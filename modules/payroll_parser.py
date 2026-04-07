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
    col_name = "NETPAI"
    col = df[col_name]

    # ── Fast path: column is already numeric ────────────────────────────
    if pd.api.types.is_numeric_dtype(col):
        if col.isna().any():
            na_indices = col.index[col.isna()].tolist()
            errors = [
                f"  Row {idx + 2} (index {idx}): empty/null value"
                for idx in na_indices
            ]
            raise ValueError(
                f"File '{filename}': Failed to convert NETPAI values to numbers.\n"
                f"The following rows have invalid values:\n"
                + "\n".join(errors)
                + "\nPlease fix the source file and re-upload."
            )
        result_df = df.copy()
        result_df[col_name] = col.astype("float64")
        return result_df

    # ── Slow path: mixed types ──────────────────────────────────────────
    numeric_attempt = pd.to_numeric(col, errors="coerce")
    orig_na_mask = col.isna()
    needs_cleaning = numeric_attempt.isna() & ~orig_na_mask

    errors = []
    result_series = numeric_attempt.copy()

    if orig_na_mask.any():
        for idx in col.index[orig_na_mask]:
            errors.append(f"  Row {idx + 2} (index {idx}): empty/null value")

    if needs_cleaning.any():
        for idx in col.index[needs_cleaning]:
            raw_value = col.iloc[idx] if isinstance(col.index, pd.RangeIndex) else col.loc[idx]
            try:
                cleaned = detect_and_clean_number_string(str(raw_value))
                result_series.iat[col.index.get_loc(idx)] = float(cleaned)
            except ValueError as e:
                errors.append(f"  Row {idx + 2} (index {idx}): '{raw_value}' → {e}")

    if errors:
        error_detail = "\n".join(errors)
        raise ValueError(
            f"File '{filename}': Failed to convert NETPAI values to numbers.\n"
            f"The following rows have invalid values:\n"
            f"{error_detail}\n"
            "Please fix the source file and re-upload."
        )

    result_df = df.copy()
    result_df[col_name] = result_series.astype("float64")
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

    # Step 6 & 7: Filter empty keys, groupby, sum values, convert to cents

    # Filter rows with empty normalized NUMCPT (vectorized)
    valid_mask = df["_NUMCPT_NORM"].astype(str).str.strip().ne("")
    df = df[valid_mask]

    if df.empty:
        return []

    # Vectorized cleanup of identity columns
    df["NUMCPT"] = df["NUMCPT"].astype(str).str.strip()
    df["NOM"] = df["NOM"].astype(str).str.strip()
    df["PRENOM"] = df["PRENOM"].astype(str).str.strip()

    # Groupby normalized NUMCPT: sum BRUTSS + NETPAI, keep first identity
    group_key = "_NUMCPT_NORM"
    sums = df.groupby(group_key)[["BRUTSS", "NETPAI"]].sum()
    first_rows = df.groupby(group_key)[["NUMCPT", "NOM", "PRENOM"]].first()

    # Build result list from small deduplicated set
    employees = []
    for norm in sorted(sums.index):
        employees.append(PayrollEmployee(
            numcpt_raw=first_rows.at[norm, "NUMCPT"],
            numcpt_norm=norm,
            nom=first_rows.at[norm, "NOM"],
            prenom=first_rows.at[norm, "PRENOM"],
            brutss_cents=_float_to_cents(float(sums.at[norm, "BRUTSS"])),
            netpai_cents=_float_to_cents(float(sums.at[norm, "NETPAI"])),
        ))

    return employees
