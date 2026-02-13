# -*- coding: utf-8 -*-
"""
Entry point for sales and revenue analytics.
Runs the analysis using the Excel file in the project root.
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

from src.data_processing import load_sheets, prepare_base
from src.excel_export import export_excel
from src.kpi_calculations import build_all_tables
from src.utils import run_selftest

# Project root path
PROJECT_ROOT = Path(__file__).parent.resolve()

# Default input/output Excel files in project root
DEFAULT_INPUT = PROJECT_ROOT / "customers_database.xlsx"
DEFAULT_OUTPUT = PROJECT_ROOT / "SalesAnalysis_Report.xlsx"


def log(message: str) -> None:
    """Print a timestamped log message."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")


def parse_args(argv: list[str]) -> argparse.Namespace:
    """
    Command line arguments.
    By default uses the Excel file in the project root.
    """
    parser = argparse.ArgumentParser(
        description="Sales and Revenue Analytics - Generates KPIs and reports in Excel"
    )
    parser.add_argument(
        "--input",
        default=str(DEFAULT_INPUT),
        help=f"Path to input Excel file (default: {DEFAULT_INPUT.name})",
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT),
        help=f"Path to output Excel file (default: {DEFAULT_OUTPUT.name})",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Don't print logs or output file path",
    )
    parser.add_argument(
        "--selftest",
        action="store_true",
        help="Run quick tests and exit",
    )
    return parser.parse_args(argv)


def run_analysis(input_path: Path, output_path: Path) -> dict:
    """
    Run the full analysis workflow.
    Returns a dict with stats for logging purposes.
    """
    # Step 1: Validate input file
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # Step 2: Load Excel sheets
    bd, cdrg, pres = load_sheets(input_path)

    # Step 3: Clean and merge data
    df = prepare_base(bd, cdrg, pres)

    # Step 4: Calculate KPIs
    tables = build_all_tables(df)

    # Step 5: Export to Excel
    export_excel(tables, output_path)

    # Return stats for logging
    return {
        "bd_rows": len(bd),
        "cdrg_rows": len(cdrg),
        "pres_rows": len(pres),
        "merged_rows": len(df),
        "unique_sellers": df["Vendedor_key"].nunique(),
        "unique_customers": df["No. Cliente"].nunique(),
        "date_min": df["Fecha Operación"].min().date(),
        "date_max": df["Fecha Operación"].max().date(),
        "sheets": list(tables.keys()),
    }


def run_with_logging(input_path: Path, output_path: Path, quiet: bool = False) -> int:
    """
    Wrapper that runs the analysis and handles logging.
    """
    if quiet:
        run_analysis(input_path, output_path)
        return 0

    log("=" * 60)
    log("SALES & REVENUE ANALYTICS - Starting Analysis")
    log("=" * 60)

    log(f"Step 1/5: Validating input file: {input_path.name}")
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    log("         Input file found")

    log("Step 2/5: Loading Excel sheets (BD, Ciudad-Region, Presupuesto)...")
    bd, cdrg, pres = load_sheets(input_path)
    log(f"         BD sheet: {len(bd)} rows")
    log(f"         Ciudad-Region sheet: {len(cdrg)} rows")
    log(f"         Presupuesto sheet: {len(pres)} rows")

    log("Step 3/5: Cleaning data and merging tables...")
    df = prepare_base(bd, cdrg, pres)
    log(f"         Merged base created: {len(df)} rows")
    log(f"         Unique sellers: {df['Vendedor_key'].nunique()}")
    log(f"         Unique customers: {df['No. Cliente'].nunique()}")
    log(
        f"         Date range: {df['Fecha Operación'].min().date()} to {df['Fecha Operación'].max().date()}"
    )

    log("Step 4/5: Calculating KPIs...")
    tables = build_all_tables(df)
    log(f"         Generated {len(tables)} Excel sheets:")
    for sheet_name in tables.keys():
        log(f"           - {sheet_name}")

    log(f"Step 5/5: Exporting to Excel: {output_path.name}")
    export_excel(tables, output_path)
    log("         Data exported successfully")

    log("Analysis complete!")
    log("=" * 60)
    log(f"OUTPUT FILE: {output_path}")
    log("=" * 60)

    return 0


def main(argv: list[str] | None = None) -> int:
    """Main entry point."""
    if argv is None:
        argv = sys.argv[1:]

    args = parse_args(argv)

    # Selftest: quick test of key functions
    if args.selftest:
        log("Running self-tests...")
        run_selftest()
        log("All tests passed!")
        return 0

    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    return run_with_logging(input_path, output_path, quiet=args.quiet)


if __name__ == "__main__":
    raise SystemExit(main())
