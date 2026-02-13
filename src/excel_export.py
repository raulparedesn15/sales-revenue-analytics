# -*- coding: utf-8 -*-
"""
Excel Export Module
===================

Functions for exporting data to Excel files.
"""

from pathlib import Path

import pandas as pd

from src.utils import safe_sheet_name


def export_excel(tables: dict[str, pd.DataFrame], output_path: Path) -> None:
    """Export all tables to an Excel file with properly named tabs."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        used: set[str] = set()
        for key, df in tables.items():
            sheet = safe_sheet_name(key, used)
            used.add(sheet)
            df.to_excel(writer, sheet_name=sheet, index=False)
