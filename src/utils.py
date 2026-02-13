# -*- coding: utf-8 -*-

"""
Sales & Revenue Analytics | Utility Functions
==============================================

Objective
---------
Load the Excel file (3 sheets: BD, Ciudad-Region, Presupuesto),
clean and merge the data to calculate the requested KPIs, plus an additional
free analysis focused on portfolio, concentration, and customer reactivation.
Export a final Excel with multiple tabs ready for pivot tables and dashboards.

Main output tabs
----------------
- Merged_Base: cleaned and merged base (ready for Excel pivots).
- KPI1: monthly revenue (total / region / city / seller).
- KPI2: quarterly revenue and % growth (total / region / seller).
- KPI3: simple annual projection with available data.
- KPI4: budget compliance by seller.
- Exercise_3: extra portfolio/concentration analysis and customers to reactivate."""

import re
from pathlib import Path

import numpy as np
import pandas as pd

# ===============================================================
# 1) HELPER FUNCTIONS (TEXT CLEANING AND KEY CREATION)
# ===============================================================


def norm_spaces(value: object) -> str:
    """Clean text whitespace.
    - Strip leading/trailing spaces
    - Collapse multiple consecutive spaces into one
    Used so seller/customer names don't fail due to formatting issues.
    """
    return " ".join(str(value or "").strip().split())


def normalize_vendor_name(raw: object, remove_leading_digits: bool = False) -> str:
    """Create a standard key for the seller: `Vendedor_key`.
    The goal is to make variations like:
    - "Irving   Hernandez"
    - "IRVING HERNANDEZ"
    - "  001 Irving Hernandez"  (when it has leading numbers)
    become exactly the same text for reliable merges."""
    s = norm_spaces(raw)
    if remove_leading_digits:
        # In Ciudad-Region sheet it comes as "001 FIRST LAST".
        # Remove leading digits so tables match.
        s = re.sub(r"^\d+\s*", "", s)
    # Standardize to uppercase.
    return norm_spaces(s).upper()


def budget_to_common_key(raw: object) -> str:
    """Adjust seller name from Presupuesto sheet to match common format.
    In Presupuesto the name comes as: "LASTNAME FIRSTNAME".
    In BD/Ciudad-Region it comes as: "FIRSTNAME LASTNAME".
    So here we flip the order to generate a compatible key."""
    s = normalize_vendor_name(raw)
    if not s:
        return s
    parts = s.split()
    if len(parts) >= 2:
        # E.g.: "HERNANDEZ IRVING" -> "IRVING HERNANDEZ"
        return " ".join(parts[1:] + [parts[0]])
    return s


# ============================================================
# 2) LOADING AND VALIDATION
# ============================================================


def require_columns(df: pd.DataFrame, cols: list[str], df_name: str) -> None:
    """Validate that required columns exist before proceeding.
    This avoids reaching the end only to discover something was missing.
    Fail fast with a clear message."""
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing columns in '{df_name}': {missing}. Available columns: {list(df.columns)}"
        )


def load_sheets(xlsx_path: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    # Load the 3 sheets from the Excel file: BD, Ciudad-Region and Presupuesto.
    xls = pd.ExcelFile(xlsx_path)
    required = ["BD", "Ciudad-Region", "Presupuesto"]
    missing = [s for s in required if s not in xls.sheet_names]
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
    # Clean, create keys, merge tables and return the final base (Merged_Base).
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

    # Ensure date is actual date, not text.
    bd["Fecha Operación"] = pd.to_datetime(bd["Fecha Operación"], errors="coerce")
    if bd["Fecha Operación"].isna().any():
        # If there are invalid dates, fail early to avoid affecting KPIs
        raise ValueError(
            "Invalid dates found in BD (Fecha Operación). Check the column format."
        )

    # Create time columns (month/quarter). Use Period for proper sorting, then convert to text.
    bd["Month_Period"] = bd["Fecha Operación"].dt.to_period("M")  # type: ignore[union-attr]
    bd["Quarter_Period"] = bd["Fecha Operación"].dt.to_period("Q")  # type: ignore[union-attr]
    bd["Month"] = bd["Month_Period"].astype(str)
    bd["Quarter"] = bd["Quarter_Period"].astype(str)

    # Create standard key (Vendedor_key) in each table for error-free merging.
    bd["Vendedor_key"] = bd["Vendedor"].apply(normalize_vendor_name)
    cdrg["Vendedor_key"] = cdrg["NOMBRE"].apply(
        lambda x: normalize_vendor_name(x, remove_leading_digits=True)
    )
    pres["Vendedor_key"] = pres["Vendedor"].apply(budget_to_common_key)

    # Keep only minimal columns for merge and avoid duplicates.
    # This prevents row multiplication during merge.
    cdrg_min = cdrg[["Vendedor_key", "CIUDAD", "REGION"]].drop_duplicates(
        subset=["Vendedor_key"]
    )
    pres_min = pres[["Vendedor_key", "Presupuesto"]].drop_duplicates(
        subset=["Vendedor_key"]
    )

    # Merge BD with Ciudad-Region and then with Presupuesto.
    df = bd.merge(cdrg_min, on="Vendedor_key", how="left")
    df = df.merge(pres_min, on="Vendedor_key", how="left")

    # Ensure revenue and budget are numeric.
    # Convert invalid revenues to 0 to avoid breaking sums.
    df["Ingreso Operación"] = pd.to_numeric(
        df["Ingreso Operación"], errors="coerce"
    ).fillna(0.0)
    df["Presupuesto"] = pd.to_numeric(df["Presupuesto"], errors="coerce")

    return df


# ============================================================
# 3) KPIs - Resolving the 3 requested KPIs
# ============================================================

# KPI 1: Monthly revenue with breakdown by total, city, region and seller
# to facilitate analysis.


def kpi_monthly_revenue(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    # Generate multiple tables because it's easier to build a dashboard in Excel
    # when levels are already prepared (total / region / city / seller).
    # Total Monthly
    overall = (
        df.groupby("Month_Period", as_index=False)["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Monthly_Revenue"})  # type: ignore[arg-type]
        .sort_values("Month_Period")
    )
    overall["Month"] = overall["Month_Period"].astype(str)
    overall = overall[["Month", "Monthly_Revenue"]]
    # Monthly by Region
    by_region = (
        df.groupby(["Month_Period", "REGION"], as_index=False)["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Monthly_Revenue"})  # type: ignore[arg-type]
        .sort_values(["Month_Period", "REGION"])
    )
    by_region["Month"] = by_region["Month_Period"].astype(str)
    by_region = by_region[["Month", "REGION", "Monthly_Revenue"]]
    # Monthly by city
    by_city = (
        df.groupby(["Month_Period", "REGION", "CIUDAD"], as_index=False)[
            "Ingreso Operación"
        ]
        .sum()
        .rename(columns={"Ingreso Operación": "Monthly_Revenue"})  # type: ignore[arg-type]
        .sort_values(["Month_Period", "REGION", "CIUDAD"])
    )
    by_city["Month"] = by_city["Month_Period"].astype(str)
    by_city = by_city[["Month", "REGION", "CIUDAD", "Monthly_Revenue"]]
    # Monthly by seller
    by_seller = (
        df.groupby(
            ["Month_Period", "Vendedor_key", "REGION", "CIUDAD"], as_index=False
        )["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Monthly_Revenue"})  # type: ignore[arg-type]
        .sort_values(["Month_Period", "Vendedor_key"])
    )
    by_seller["Month"] = by_seller["Month_Period"].astype(str)
    by_seller = by_seller[
        ["Month", "Vendedor_key", "REGION", "CIUDAD", "Monthly_Revenue"]
    ]

    return {
        "KPI1_Monthly_Total": overall,
        "KPI1_Monthly_Region": by_region,
        "KPI1_Monthly_City": by_city,
        "KPI1_Monthly_Seller": by_seller,
    }


# KPI 2: Quarterly revenue + % growth vs previous quarter.


def kpi_quarterly_growth(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    # Logic: sum by quarter and compare against previous quarter.
    # Similar to KPI 1, I break it down by Region and Seller for additional
    # dashboard insights.
    # Total Quarterly
    overall = (
        df.groupby("Quarter_Period", as_index=False)["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Quarterly_Revenue"})  # type: ignore[arg-type]
        .sort_values("Quarter_Period")
    )
    # Calculate percentage using pct_change, same for region and seller
    overall["% Quarterly_Growth"] = overall["Quarterly_Revenue"].pct_change() * 100
    overall["Quarter"] = overall["Quarter_Period"].astype(str)
    overall = overall[["Quarter", "Quarterly_Revenue", "% Quarterly_Growth"]]
    # Quarterly by Region
    by_region = (
        df.groupby(["Quarter_Period", "REGION"], as_index=False)["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Quarterly_Revenue"})  # type: ignore[arg-type]
        .sort_values(["REGION", "Quarter_Period"])
    )
    by_region["% Quarterly_Growth"] = (
        by_region.groupby("REGION")["Quarterly_Revenue"].pct_change() * 100
    )
    by_region["Quarter"] = by_region["Quarter_Period"].astype(str)
    by_region = by_region[
        ["Quarter", "REGION", "Quarterly_Revenue", "% Quarterly_Growth"]
    ]
    # Quarterly by Seller
    by_seller = (
        df.groupby(
            ["Quarter_Period", "Vendedor_key", "REGION", "CIUDAD"], as_index=False
        )["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Quarterly_Revenue"})  # type: ignore[arg-type]
        .sort_values(["Vendedor_key", "Quarter_Period"])
    )
    by_seller["% Quarterly_Growth"] = (
        by_seller.groupby("Vendedor_key")["Quarterly_Revenue"].pct_change() * 100
    )
    by_seller["Quarter"] = by_seller["Quarter_Period"].astype(str)
    by_seller = by_seller[
        [
            "Quarter",
            "Vendedor_key",
            "REGION",
            "CIUDAD",
            "Quarterly_Revenue",
            "% Quarterly_Growth",
        ]
    ]

    return {
        "KPI2_Quarterly_Total": overall,
        "KPI2_Quarterly_Region": by_region,
        "KPI2_Quarterly_Seller": by_seller,
    }


# KPI 3: Annual projection for year-end (monthly average * 12)
# KPI 4: Budget compliance by seller (plus projected compliance)


def kpi_projection(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    # First check how many actual months we have in the history (not all 12 months available).
    months_with_data = int(df["Month_Period"].nunique())
    if months_with_data <= 0:
        raise ValueError("No months with data to calculate projection.")

    # Calculate monthly average and project to 12 months for annual projection.
    accumulated_revenue = float(df["Ingreso Operación"].sum())
    avg_monthly_revenue = accumulated_revenue / months_with_data
    annual_projection = avg_monthly_revenue * 12

    total = pd.DataFrame(
        [
            {
                "Start_Month": str(df["Month_Period"].min()),
                "End_Month": str(df["Month_Period"].max()),
                "Months_With_Data": months_with_data,
                "Accumulated_Revenue": accumulated_revenue,
                "Avg_Monthly_Revenue": avg_monthly_revenue,
                "Annual_Projection": annual_projection,
            }
        ]
    )
    # Projection and compliance by seller. This provides additional insights
    # on each seller's performance based on their budget and revenue.
    by_seller = df.groupby(["Vendedor_key", "REGION", "CIUDAD"], as_index=False).agg(
        Accumulated_Revenue=("Ingreso Operación", "sum"),
        Months_With_Data=("Month_Period", "nunique"),
        # Budget: use "first" because it's already merged by seller.
        # Should be the same value repeated for all rows of that seller.
        Budget=("Presupuesto", "first"),
    )
    # Calculate monthly average per seller, and detect budget anomalies
    by_seller["Avg_Monthly_Revenue"] = by_seller["Accumulated_Revenue"] / by_seller[
        "Months_With_Data"
    ].replace(0, np.nan)
    by_seller["Annual_Projection"] = by_seller["Avg_Monthly_Revenue"] * 12
    by_seller["% Budget_Compliance"] = (
        by_seller["Accumulated_Revenue"] / by_seller["Budget"] * 100
    )
    by_seller["% Projected_Compliance"] = (
        by_seller["Annual_Projection"] / by_seller["Budget"] * 100
    )
    by_seller = by_seller.sort_values("% Budget_Compliance", ascending=False)

    return {
        "KPI3_Projection_Total": total,
        "KPI4_Compliance_Seller": by_seller,
    }


# ============================================================
# 4) Exercise 3 - Free Analysis
# ============================================================

# Exercise 3: Analyzing portfolio, concentration, and customer reactivation for each seller.
# The goal is to derive insights that help commercial management and decision-making,
# not just KPIs. These metrics help identify portfolio risk, customer volume, etc.


def exercise_3_free_analysis(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    # 1) First get accumulated revenue per seller
    seller_accumulated = (
        df.groupby(["Vendedor_key", "CIUDAD", "REGION"], as_index=False)[
            "Ingreso Operación"
        ]
        .sum()
        .rename(columns={"Ingreso Operación": "Accumulated_Revenue"})  # type: ignore[arg-type]
    )

    # 2) Then drill down to seller-customer level to see contribution per customer
    portfolio_seller_customer = (
        df.groupby(["Vendedor_key", "No. Cliente"], as_index=False)["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Customer_Revenue"})  # type: ignore[arg-type]
    ).merge(
        seller_accumulated[["Vendedor_key", "Accumulated_Revenue"]],
        on="Vendedor_key",
        how="left",
    )

    # 3) With seller total, calculate what % each customer represents
    portfolio_seller_customer["% of_Seller"] = portfolio_seller_customer[
        "Customer_Revenue"
    ] / portfolio_seller_customer["Accumulated_Revenue"].replace(0, np.nan)

    # 4) Using HHI concentration index to determine if a seller's revenue
    # depends on few customers or is well distributed. Makes it easier to assess
    # portfolio risk.
    # If high (above 0.18), seller depends on few customers = risky.
    # Simple example: if 1 customer = 100% -> 1^2 = 1 (max concentration).
    # If 10 equal customers (10% each) -> 10*(0.1^2) = 0.10 (more distributed).
    concentration_hhi = (
        portfolio_seller_customer.groupby("Vendedor_key", as_index=False)["% of_Seller"]
        .apply(lambda s: float((s.fillna(0) ** 2).sum()))
        .rename(columns={"% of_Seller": "HHI_Concentration"})
    )

    # 5) Customer ranking to sum Top 5, to determine if seller depends on
    # certain customers, helping assess portfolio risk.
    portfolio_seller_customer["Customer_Rank"] = portfolio_seller_customer.groupby(
        "Vendedor_key"
    )["Customer_Revenue"].rank(method="first", ascending=False)

    top5_revenue = (
        portfolio_seller_customer[portfolio_seller_customer["Customer_Rank"] <= 5]
        .groupby("Vendedor_key", as_index=False)["Customer_Revenue"]
        .sum()
        .rename(columns={"Customer_Revenue": "Top5_Customer_Revenue"})
    )

    # 6) Portfolio size: how many unique customers each seller manages
    unique_customers = (
        portfolio_seller_customer.groupby("Vendedor_key", as_index=False)["No. Cliente"]
        .nunique()
        .rename(columns={"No. Cliente": "Unique_Customers"})
    )

    # 7) Build summary table by seller
    seller_portfolio = (
        seller_accumulated.merge(unique_customers, on="Vendedor_key", how="left")
        .merge(top5_revenue, on="Vendedor_key", how="left")
        .merge(concentration_hhi, on="Vendedor_key", how="left")
    )

    seller_portfolio["Top5_Customer_Revenue"] = seller_portfolio[
        "Top5_Customer_Revenue"
    ].fillna(0.0)
    seller_portfolio["% Top5_Revenue"] = (
        seller_portfolio["Top5_Customer_Revenue"]
        / seller_portfolio["Accumulated_Revenue"].replace(0, np.nan)
        * 100
    )
    # 8) Alert flag to prioritize follow-up when 70%+ of seller's revenue
    # comes from Top 5 customers (high risk).
    seller_portfolio["Risk (Concentration)"] = np.where(
        seller_portfolio["% Top5_Revenue"] >= 70, "HIGH", "NORMAL"
    )
    seller_portfolio = seller_portfolio.sort_values("% Top5_Revenue", ascending=False)

    # 9) Reactivation: find high-value customers who haven't purchased recently,
    # enabling outreach decisions to reactivate purchases.
    max_date = df["Fecha Operación"].max()
    customer_details = df.groupby(["Vendedor_key", "No. Cliente"], as_index=False).agg(
        Total_Revenue=("Ingreso Operación", "sum"),
        Purchases=("Guia", "count"),
        Last_Purchase=("Fecha Operación", "max"),
        First_Purchase=("Fecha Operación", "min"),
        City=("CIUDAD", "first"),
        Region=("REGION", "first"),
    )
    customer_details["Days_Since_Last_Purchase"] = (
        max_date - customer_details["Last_Purchase"]
    ).dt.days

    # Define high-value customers as top 20% of revenue within each seller.
    # Done per seller since each portfolio is different.
    threshold = customer_details.groupby("Vendedor_key")["Total_Revenue"].transform(
        lambda s: s.quantile(0.80)
    )
    customer_details["High_Value"] = customer_details["Total_Revenue"] >= threshold

    # Define dormant customer as 30+ days without purchase (for practical purposes),
    # so action may be needed.
    customers_to_reactivate = (
        customer_details[
            (customer_details["High_Value"])
            & (customer_details["Days_Since_Last_Purchase"] >= 30)
        ]
        .sort_values(
            ["Days_Since_Last_Purchase", "Total_Revenue"], ascending=[False, False]
        )
        .drop(columns=["High_Value"])
    )

    return {
        "E3_Seller_Portfolio": seller_portfolio,
        "E3_Reactivate_Customers": customers_to_reactivate,
    }


# ============================================================
# 5) Export to Excel (final deliverable)
# ============================================================


def safe_sheet_name(name: str, used: set[str]) -> str:
    # Excel limits sheet names to 31 characters. Here we ensure:
    # truncate to 31 and avoid duplicate names
    base = name.strip().replace("/", "_")
    base = base[:31] if len(base) > 31 else base
    if base not in used:
        return base
    # If exists, use suffixes until unique
    for i in range(1, 1000):
        suffix = f"_{i}"
        candidate = base[: 31 - len(suffix)] + suffix
        if candidate not in used:
            return candidate
    raise ValueError("Could not generate a unique sheet name for Excel.")


def export_excel(tables: dict[str, pd.DataFrame], output_path: Path) -> None:
    # Export all tables to an Excel file with the defined tabs.
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        used: set[str] = set()
        for key, df in tables.items():
            sheet = safe_sheet_name(key, used)
            used.add(sheet)
            df.to_excel(writer, sheet_name=sheet, index=False)


def build_all_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    # Build a dictionary {tab_name: dataframe} for easy export.
    # Makes it easy to add/remove tabs without breaking anything.
    tables: dict[str, pd.DataFrame] = {}
    tables["Merged_Base"] = df.drop(
        columns=["Month_Period", "Quarter_Period"], errors="ignore"
    )
    tables.update(kpi_monthly_revenue(df))
    tables.update(kpi_quarterly_growth(df))
    tables.update(kpi_projection(df))
    tables.update(exercise_3_free_analysis(df))
    return tables


def run_selftest() -> None:
    # Quick tests to ensure basic functionality works.
    assert normalize_vendor_name("  Juan  Pérez ") == "JUAN PÉREZ"
    assert (
        normalize_vendor_name("001  Juan Pérez", remove_leading_digits=True)
        == "JUAN PÉREZ"
    )
    assert budget_to_common_key("PEREZ JUAN") == "JUAN PEREZ"
    used: set[str] = set()
    a = safe_sheet_name("A" * 40, used)
    used.add(a)
    b = safe_sheet_name("A" * 40, used)
    assert a != b
    assert len(a) <= 31 and len(b) <= 31


def run_analysis(input_path: Path, output_path: Path, quiet: bool = False) -> int:
    """
    Main entry point for the analysis.
    Receives input/output paths and executes the full workflow.
    """
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # 1) Load sheets from original Excel (BD, Ciudad-Region, Presupuesto)
    bd, cdrg, pres = load_sheets(input_path)

    # 2) Prepare final base (cleaning + merges)
    df = prepare_base(bd, cdrg, pres)

    # 3) Calculate KPIs + exercise 3
    tables = build_all_tables(df)

    # 4) Export everything to Excel (deliverable file)
    export_excel(tables, output_path)

    # If not quiet, print the path so it's easy to find the file.
    if not quiet:
        print(str(output_path))
    return 0
