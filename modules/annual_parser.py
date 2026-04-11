"""
annual_parser.py

Parse a trimestrial output file for the annual declaration.

Each quarterly file is expected to have a "Consolidation" sheet (or first sheet)
with at minimum: NUMCPT, NOM, PRENOM, BRUTSS_TOTAL.
Optional columns: NUMSS, ADM, DATNAIS, NBRTRAV, DATENT, DATSOR.

Reuses the existing validator for safe file loading and row cleanup.
"""

import pandas as pd

from modules.validator import (
    load_excel_safe,
    drop_non_employee_rows,
    validate_not_empty,
)
from modules.cleaner import _normalize_numcpt, EMPTY_BRUTSS_COLUMN
from modules.annual_types import EmployeeKey, QuarterlyEntry

# The BRUTSS_TOTAL column name as produced by the trimestrial exporter
BRUTSS_TOTAL_COLUMN = "BRUTSS_TOTAL"

# Minimum required columns
REQUIRED_COLUMNS = {"NUMCPT", "NOM", "PRENOM", "BRUTSS_TOTAL"}


def _float_to_cents(value: float) -> int:
    """Convert a float value to integer cents (round-safe)."""
    return round(value * 100)


def _safe_str(val) -> str:
    """Convert a value to a clean string, treating NaN/None as empty."""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    return "" if s in ("nan", "None") else s


def _validate_annual_columns(df: pd.DataFrame, filename: str) -> None:
    """Validate that all required columns for annual parsing are present."""
    present = set(df.columns.tolist())
    missing = REQUIRED_COLUMNS - present

    if missing:
        missing_list = ", ".join(f'"{c}"' for c in sorted(missing))
        found_list = ", ".join(f'"{c}"' for c in sorted(present))
        raise ValueError(
            f"File '{filename}' is missing required column(s): {missing_list}.\n"
            f"Columns found in file: {found_list}\n"
            "Please make sure this is a trimestrial output file containing BRUTSS_TOTAL."
        )


def parse_quarterly_file(file_obj, filename: str) -> tuple:
    """
    Load, validate, and parse a single quarterly (trimestrial output) file.

    Pipeline:
    1. Load Excel (first sheet)
    2. Drop non-employee rows
    3. Validate not empty + required columns
    4. Convert BRUTSS_TOTAL to float64
    5. Normalize NUMCPT for matching
    6. Group by NUMCPT (sum BRUTSS_TOTAL for any duplicates)

    Parameters
    ----------
    file_obj : file-like object
        Uploaded Excel file.
    filename : str
        Original filename, used in error messages.

    Returns
    -------
    tuple[dict[EmployeeKey, QuarterlyEntry], list[tuple]]
        - lookup: One entry per unique employee with BRUTSS_TOTAL in cents.
        - empty_brutss: List of (numcpt, nom, prenom) for employees with empty BRUTSS_TOTAL.
    """
    # Step 1: Load
    df = load_excel_safe(file_obj, filename)

    # Step 2: Drop non-employee rows
    df, _dropped = drop_non_employee_rows(df)

    # Step 3: Validate
    validate_not_empty(df, filename)
    _validate_annual_columns(df, filename)

    # Step 4: Convert BRUTSS_TOTAL to float64 (handle NaN → 0)
    col = df[BRUTSS_TOTAL_COLUMN]
    empty_brutss = []

    if pd.api.types.is_numeric_dtype(col):
        empty_mask = col.isna()
        if empty_mask.any():
            for _, row in df[empty_mask].iterrows():
                empty_brutss.append((
                    _safe_str(row.get("NUMCPT", "")),
                    _safe_str(row.get("NOM", "")),
                    _safe_str(row.get("PRENOM", "")),
                ))
        df[BRUTSS_TOTAL_COLUMN] = col.fillna(0.0).astype("float64")
    else:
        numeric = pd.to_numeric(col, errors="coerce")
        orig_na = col.isna()
        empty_mask = orig_na | numeric.isna()
        if empty_mask.any():
            for _, row in df[empty_mask].iterrows():
                empty_brutss.append((
                    _safe_str(row.get("NUMCPT", "")),
                    _safe_str(row.get("NOM", "")),
                    _safe_str(row.get("PRENOM", "")),
                ))
        df[BRUTSS_TOTAL_COLUMN] = numeric.fillna(0.0).astype("float64")

    # Step 5: Normalize NUMCPT
    df["_NUMCPT_NORM"] = _normalize_numcpt(df["NUMCPT"])

    # Filter empty keys
    valid_mask = df["_NUMCPT_NORM"].astype(str).str.strip().ne("")
    df = df[valid_mask]

    if df.empty:
        return {}, empty_brutss

    # Vectorized cleanup
    df["NUMCPT"] = df["NUMCPT"].astype(str).str.strip()
    df["NOM"] = df["NOM"].astype(str).str.strip()
    df["PRENOM"] = df["PRENOM"].astype(str).str.strip()

    _OPT_COLS = ["NUMSS", "ADM", "DATNAIS", "NBRTRAV", "DATENT", "DATSOR"]
    for col_name in _OPT_COLS:
        if col_name in df.columns:
            df[col_name] = (
                df[col_name].fillna("").astype(str).str.strip()
                .replace({"nan": "", "None": ""})
            )
        else:
            df[col_name] = ""

    # Step 6: Group by NUMCPT — sum BRUTSS_TOTAL, keep first identity
    group_key = "_NUMCPT_NORM"
    brutss_sums = df.groupby(group_key)[BRUTSS_TOTAL_COLUMN].sum()
    first_rows = df.groupby(group_key).first()

    # Build lookup
    lookup: dict = {}
    for norm_numcpt, row in first_rows.iterrows():
        key = EmployeeKey(numcpt=norm_numcpt)
        lookup[key] = QuarterlyEntry(
            numcpt_raw=row["NUMCPT"],
            nom=row["NOM"],
            prenom=row["PRENOM"],
            numss=row.get("NUMSS", ""),
            adm=row.get("ADM", ""),
            datnais=row.get("DATNAIS", ""),
            nbrtrav=row.get("NBRTRAV", ""),
            datent=row.get("DATENT", ""),
            datsor=row.get("DATSOR", ""),
            brutss_total_cents=_float_to_cents(float(brutss_sums[norm_numcpt])),
        )

    return lookup, empty_brutss
