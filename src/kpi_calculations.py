# -*- coding: utf-8 -*-
"""
KPI Calculations Module
=======================

Functions for calculating all KPIs and analyses.
"""

import numpy as np
import pandas as pd


def kpi_monthly_revenue(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """KPI 1: Monthly revenue with breakdown by total, city, region and seller."""
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


def kpi_quarterly_growth(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """KPI 2: Quarterly revenue + % growth vs previous quarter."""
    # Total Quarterly
    overall = (
        df.groupby("Quarter_Period", as_index=False)["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Quarterly_Revenue"})  # type: ignore[arg-type]
        .sort_values("Quarter_Period")
    )
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


def kpi_projection(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """KPI 3: Annual projection (monthly average * 12) and KPI 4: Budget compliance."""
    months_with_data = int(df["Month_Period"].nunique())
    if months_with_data <= 0:
        raise ValueError("No months with data to calculate projection.")

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

    # Projection and compliance by seller
    by_seller = df.groupby(["Vendedor_key", "REGION", "CIUDAD"], as_index=False).agg(
        Accumulated_Revenue=("Ingreso Operación", "sum"),
        Months_With_Data=("Month_Period", "nunique"),
        Budget=("Presupuesto", "first"),
    )
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


def exercise_3_free_analysis(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Exercise 3: Portfolio concentration analysis and customer reactivation."""
    # 1) Accumulated revenue per seller
    seller_accumulated = (
        df.groupby(["Vendedor_key", "CIUDAD", "REGION"], as_index=False)[
            "Ingreso Operación"
        ]
        .sum()
        .rename(columns={"Ingreso Operación": "Accumulated_Revenue"})  # type: ignore[arg-type]
    )

    # 2) Seller-customer level breakdown
    portfolio_seller_customer = (
        df.groupby(["Vendedor_key", "No. Cliente"], as_index=False)["Ingreso Operación"]
        .sum()
        .rename(columns={"Ingreso Operación": "Customer_Revenue"})  # type: ignore[arg-type]
    ).merge(
        seller_accumulated[["Vendedor_key", "Accumulated_Revenue"]],
        on="Vendedor_key",
        how="left",
    )

    # 3) Calculate what % each customer represents
    portfolio_seller_customer["% of_Seller"] = portfolio_seller_customer[
        "Customer_Revenue"
    ] / portfolio_seller_customer["Accumulated_Revenue"].replace(0, np.nan)

    # 4) HHI concentration index
    concentration_hhi = (
        portfolio_seller_customer.groupby("Vendedor_key", as_index=False)["% of_Seller"]
        .apply(lambda s: float((s.fillna(0) ** 2).sum()))
        .rename(columns={"% of_Seller": "HHI_Concentration"})
    )

    # 5) Customer ranking for Top 5
    portfolio_seller_customer["Customer_Rank"] = portfolio_seller_customer.groupby(
        "Vendedor_key"
    )["Customer_Revenue"].rank(method="first", ascending=False)

    top5_revenue = (
        portfolio_seller_customer[portfolio_seller_customer["Customer_Rank"] <= 5]
        .groupby("Vendedor_key", as_index=False)["Customer_Revenue"]
        .sum()
        .rename(columns={"Customer_Revenue": "Top5_Customer_Revenue"})
    )

    # 6) Portfolio size
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

    # 8) Risk flag for high concentration
    seller_portfolio["Risk (Concentration)"] = np.where(
        seller_portfolio["% Top5_Revenue"] >= 70, "HIGH", "NORMAL"
    )
    seller_portfolio = seller_portfolio.sort_values("% Top5_Revenue", ascending=False)

    # 9) Reactivation: high-value dormant customers
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

    threshold = customer_details.groupby("Vendedor_key")["Total_Revenue"].transform(
        lambda s: s.quantile(0.80)
    )
    customer_details["High_Value"] = customer_details["Total_Revenue"] >= threshold

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


def build_all_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Build a dictionary {tab_name: dataframe} for easy export."""
    tables: dict[str, pd.DataFrame] = {}
    tables["Merged_Base"] = df.drop(
        columns=["Month_Period", "Quarter_Period"], errors="ignore"
    )
    tables.update(kpi_monthly_revenue(df))
    tables.update(kpi_quarterly_growth(df))
    tables.update(kpi_projection(df))
    tables.update(exercise_3_free_analysis(df))
    return tables
