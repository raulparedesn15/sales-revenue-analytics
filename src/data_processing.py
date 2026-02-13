# -*- coding: utf-8 -*-
"""
Data Processing Module
======================

Functions for loading Excel sheets and preparing the merged base.
"""

from pathlib import Path

import pandas as pd

from src.utils import (
    budget_to_common_key,
    normalize_vendor_name,
    require_columns,
)


def load_sheets(xlsx_path: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Load the 3 sheets from the Excel file: BD, Ciudad-Region and Presupuesto."""
    xls = pd.ExcelFile(xlsx_path)
    required_sheets = ["BD", "Ciudad-Region", "Presupuesto"]
    missing = [sheet for sheet in required_sheets if sheet not in xls.sheet_names]
    if missing:
        raise ValueError(
            f"Missing sheets in Excel: {missing}. Available sheets: {xls.sheet_names}"
        )
    bd = pd.read_excel(xlsx_path, sheet_name="BD")
    cdrg = pd.read_excel(xlsx_path, sheet_name="Ciudad-Region")
    pres = pd.read_excel(xlsx_path, sheet_name="Presupuesto")
    return bd, cdrg, pres


def prepare_base(
    bd: pd.DataFrame, cdrg: pd.DataFrame, pres: pd.DataFrame
) -> pd.DataFrame:
    """Clean, create keys, merge tables and return the final base (Merged_Base)."""
    require_columns(
        bd,
        ["Fecha Operación", "Vendedor", "Ingreso Operación", "No. Cliente", "Guia"],
        "BD",
    )
    require_columns(cdrg, ["NOMBRE", "CIUDAD", "REGION"], "Ciudad-Region")
    require_columns(pres, ["Vendedor", "Presupuesto"], "Presupuesto")

    bd = bd.copy()
    cdrg = cdrg.copy()
    pres = pres.copy()

    bd["Fecha Operación"] = pd.to_datetime(bd["Fecha Operación"], errors="coerce")
    if bd["Fecha Operación"].isna().any():
        raise ValueError(
            "Invalid dates found in BD (Fecha Operación). Check the column format."
        )

    # Create time columns (month/quarter)
    bd["Month_Period"] = bd["Fecha Operación"].dt.to_period("M")  # type: ignore[union-attr]
    bd["Quarter_Period"] = bd["Fecha Operación"].dt.to_period("Q")  # type: ignore[union-attr]
    bd["Month"] = bd["Month_Period"].astype(str)
    bd["Quarter"] = bd["Quarter_Period"].astype(str)

    # Create standard key (Vendedor_key) in each table
    bd["Vendedor_key"] = bd["Vendedor"].apply(normalize_vendor_name)
    cdrg["Vendedor_key"] = cdrg["NOMBRE"].apply(
        lambda x: normalize_vendor_name(x, remove_leading_digits=True)
    )
    pres["Vendedor_key"] = pres["Vendedor"].apply(budget_to_common_key)

    # Keep only minimal columns for merge
    cdrg_min = cdrg[["Vendedor_key", "CIUDAD", "REGION"]].drop_duplicates(
        subset=["Vendedor_key"]
    )
    pres_min = pres[["Vendedor_key", "Presupuesto"]].drop_duplicates(
        subset=["Vendedor_key"]
    )

    # Merge BD with Ciudad-Region and Presupuesto
    df = bd.merge(cdrg_min, on="Vendedor_key", how="left")
    df = df.merge(pres_min, on="Vendedor_key", how="left")

    # Ensure revenue and budget are numeric
    df["Ingreso Operación"] = pd.to_numeric(
        df["Ingreso Operación"], errors="coerce"
    ).fillna(0.0)
    df["Presupuesto"] = pd.to_numeric(df["Presupuesto"], errors="coerce")

    return df
