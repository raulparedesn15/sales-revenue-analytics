# -*- coding: utf-8 -*-
"""
Entry point for sales and revenue analytics.
Runs the analysis using the Excel file in the project root.
"""

import argparse
import sys
from pathlib import Path

from src.utils import run_analysis, run_selftest

# Project root path
PROJECT_ROOT = Path(__file__).parent.resolve()

# Default input/output Excel files in project root
DEFAULT_INPUT = PROJECT_ROOT / "customers_database.xlsx"
DEFAULT_OUTPUT = PROJECT_ROOT / "SalesAnalysis_Report.xlsx"


def parse_args(argv: list[str]) -> argparse.Namespace:
    """
    Command line arguments.
    By default uses the Excel file in the project root.
    """
    p = argparse.ArgumentParser(
        description="Sales and Revenue Analytics - Generates KPIs and reports in Excel"
    )
    p.add_argument(
        "--input",
        default=str(DEFAULT_INPUT),
        help=f"Path to input Excel file (default: {DEFAULT_INPUT.name})",
    )
    p.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT),
        help=f"Path to output Excel file (default: {DEFAULT_OUTPUT.name})",
    )
    p.add_argument(
        "--quiet",
        action="store_true",
        help="Don't print the output file path",
    )
    p.add_argument(
        "--selftest",
        action="store_true",
        help="Run quick tests and exit",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """Main entry point."""
    if argv is None:
        argv = sys.argv[1:]

    args = parse_args(argv)

    # Selftest: quick test of key functions
    if args.selftest:
        run_selftest()
        return 0

    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    return run_analysis(input_path, output_path, quiet=args.quiet)


if __name__ == "__main__":
    raise SystemExit(main())
