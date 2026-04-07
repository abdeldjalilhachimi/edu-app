"""
payroll_parser.py

Load and parse a single Excel file for payroll calculation.

Reads BRUTSS and NETPAI per employee (keyed by NUMCPT).
Reuses existing validator and cleaner modules for file loading,
column validation, and number conversion.
"""

import pandas as pd

from modules.validator import (
    load_excel_safe,
    drop_non_employee_rows,
    validate_not_empty,
    validate_required_columns,
)
from modules.cleaner import (
    convert_brutss_column,
    detect_and_clean_number_string,
    _normalize_numcpt,
)
from modules.payroll_types import PayrollEmployee

# Required columns for payroll (adds NETPAI on top of the standard set)
PAYROLL_REQUIRED_COLUMNS = {"BRUTSS", "NOM", "PRENOM", "NUMCPT", "NETPAI"}


def _float_to_cents(value: float) -> int:
    """Convert a float value to integer cents (round-safe)."""
    return round(value * 100)


def _convert_netpai_column(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    """
    Convert the NETPAI column to float64, same logic as convert_brutss_column.

    Parameters
    ----------
    df : pd.DataFrame
        Must have a 'NETPAI' column.
    filename : str
        Used in error messages.

    Returns
    -------
    pd.DataFrame
        Same DataFrame with NETPAI column as float64.
    """
    col = "NETPAI"
    errors = []
    converted = []

    for idx, raw_value in enumerate(df[col]):
        row_number = idx + 2  # 1-based + header

        if pd.isna(raw_value):
            errors.append(f"  Row {row_number} (index {idx}): empty/null value")
            converted.append(None)
            continue

        if isinstance(raw_value, (int, float)):
            converted.append(float(raw_value))
            continue

        try:
            cleaned = detect_and_clean_number_string(str(raw_value))
            converted.append(float(cleaned))
        except ValueError as e:
            errors.append(f"  Row {row_number} (index {idx}): '{raw_value}' → {e}")
            converted.append(None)

    if errors:
        error_detail = "\n".join(errors)
        raise ValueError(
            f"File '{filename}': Failed to convert NETPAI values to numbers.\n"
            f"The following rows have invalid values:\n"
            f"{error_detail}\n"
            "Please fix the source file and re-upload."
        )

    result_df = df.copy()
    result_df[col] = pd.array(converted, dtype="float64")
    return result_df


def _validate_payroll_columns(df: pd.DataFrame, filename: str) -> None:
    """
    Validate that all payroll-required columns are present.

    Required: BRUTSS, NOM, PRENOM, NUMCPT, NETPAI.
    """
    present = set(df.columns.tolist())
    missing = PAYROLL_REQUIRED_COLUMNS - present

    if missing:
        missing_list = ", ".join(f'"{c}"' for c in sorted(missing))
        found_list = ", ".join(f'"{c}"' for c in sorted(present))
        raise ValueError(
            f"File '{filename}' is missing required column(s): {missing_list}.\n"
            f"Columns found in file: {found_list}"
        )


def parse_payroll_file(file_obj, filename: str) -> list:
    """
    Load, validate, and parse a single Excel file for payroll calculation.

    Pipeline:
    1. Load with openpyxl (TEXT_DTYPE for NUMCPT)
    2. Drop non-employee rows (blank/total rows)
    3. Validate not empty + required columns (including NETPAI)
    4. Convert BRUTSS and NETPAI to float64
    5. Normalize NUMCPT for matching
    6. Group by NUMCPT, summing BRUTSS + NETPAI for duplicates
    7. Convert to integer cents

    Parameters
    ----------
    file_obj : file-like object
        Uploaded Excel file.
    filename : str
        Original filename, used in error messages.

    Returns
    -------
    list[PayrollEmployee]
        One entry per unique employee. Values in integer cents.
    """
    # Step 1: Load
    df = load_excel_safe(file_obj, filename)

    # Step 2: Drop non-employee rows
    df, _dropped = drop_non_employee_rows(df)

    # Step 3: Validate
    validate_not_empty(df, filename)
    _validate_payroll_columns(df, filename)

    # Step 4: Convert BRUTSS and NETPAI to float64
    df = convert_brutss_column(df, filename)
    df = _convert_netpai_column(df, filename)

    # Step 5: Normalize NUMCPT
    df["_NUMCPT_NORM"] = _normalize_numcpt(df["NUMCPT"])

    # Step 6 & 7: Group by NUMCPT, sum values, convert to cents
    lookup: dict = {}  # numcpt_norm → PayrollEmployee

    for _, row in df.iterrows():
        norm = row["_NUMCPT_NORM"]

        if not norm or str(norm).strip() == "":
            continue

        brutss_cents = _float_to_cents(float(row["BRUTSS"]))
        netpai_cents = _float_to_cents(float(row["NETPAI"]))

        if norm in lookup:
            existing = lookup[norm]
            lookup[norm] = PayrollEmployee(
                numcpt_raw=existing.numcpt_raw,
                numcpt_norm=existing.numcpt_norm,
                nom=existing.nom,
                prenom=existing.prenom,
                brutss_cents=existing.brutss_cents + brutss_cents,
                netpai_cents=existing.netpai_cents + netpai_cents,
            )
        else:
            lookup[norm] = PayrollEmployee(
                numcpt_raw=str(row.get("NUMCPT", "")).strip(),
                numcpt_norm=norm,
                nom=str(row.get("NOM", "")).strip(),
                prenom=str(row.get("PRENOM", "")).strip(),
                brutss_cents=brutss_cents,
                netpai_cents=netpai_cents,
            )

    # Return as sorted list
    return sorted(lookup.values(), key=lambda e: e.numcpt_norm)
