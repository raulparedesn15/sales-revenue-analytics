# -*- coding: utf-8 -*-
"""
Utility Functions
=================

Validation, formatting, and helper functions used across the project.
"""

import re

import pandas as pd


# ===============================================================
# TEXT CLEANING AND KEY CREATION
# ===============================================================


def norm_spaces(value: object) -> str:
    """Clean text whitespace.
    - Strip leading/trailing spaces
    - Collapse multiple consecutive spaces into one
    Used so seller/customer names don't fail due to formatting issues.
    """
    return " ".join(str(value or "").strip().split())


def normalize_vendor_name(
    raw_vendor_name: object, remove_leading_digits: bool = False
) -> str:
    """Create a standard key for the seller: `Vendedor_key`.
    The goal is to make variations like:
    - "Irving   Hernandez"
    - "IRVING HERNANDEZ"
    - "  001 Irving Hernandez"  (when it has leading numbers)
    become exactly the same text for reliable merges."""
    formatted_vendor_name = norm_spaces(raw_vendor_name)
    if remove_leading_digits:
        # In Ciudad-Region sheet it comes as "001 FIRST LAST".
        # Remove leading digits so tables match.
        formatted_vendor_name = re.sub(r"^\d+\s*", "", formatted_vendor_name)
    # Standardize to uppercase.
    formatted_vendor_name_in_upper_case = formatted_vendor_name.upper()
    return formatted_vendor_name_in_upper_case


def budget_to_common_key(
    raw_vendor_name: object, remove_leading_digits: bool = False
) -> str:
    """Adjust seller name from Presupuesto sheet to match common format.
    In Presupuesto the name comes as: "LASTNAME FIRSTNAME".
    In BD/Ciudad-Region it comes as: "FIRSTNAME LASTNAME".
    So here we flip the order to generate a compatible key."""
    formatted_vendor_name_in_upper_case = normalize_vendor_name(
        raw_vendor_name, remove_leading_digits=remove_leading_digits
    )
    if not formatted_vendor_name_in_upper_case:
        return formatted_vendor_name_in_upper_case
    parts = formatted_vendor_name_in_upper_case.split()
    if len(parts) >= 2:
        # E.g.: "HERNANDEZ IRVING" -> "IRVING HERNANDEZ"
        return " ".join(parts[1:] + [parts[0]])
    return formatted_vendor_name_in_upper_case


# ===============================================================
# VALIDATION
# ===============================================================


def require_columns(df: pd.DataFrame, cols: list[str], df_name: str) -> None:
    """Validate that required columns exist before proceeding.
    This avoids reaching the end only to discover something was missing.
    Fail fast with a clear message."""
    missing = [col for col in cols if col not in df.columns]
    if missing:
        raise ValueError(
            f"Missing columns in '{df_name}': {missing}. Available columns: {list(df.columns)}"
        )


# ===============================================================
# EXCEL HELPERS
# ===============================================================


def safe_sheet_name(name: str, used: set[str]) -> str:
    """Generate a valid Excel sheet name (max 31 chars, unique)."""
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


# ===============================================================
# SELF-TEST
# ===============================================================


def run_selftest() -> None:
    """Quick tests to ensure basic functionality works."""
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
