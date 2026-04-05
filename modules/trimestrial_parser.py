"""
trimestrial_parser.py

Helper parsing functions for the trimestrial BRUTSS consolidation.

Loads and validates each monthly Excel file, then extracts employee
records with BRUTSS converted to integer cents for safe arithmetic.

Reuses the existing validator and cleaner modules for file loading,
column validation, and BRUTSS string-to-float conversion.
"""

import math
import pandas as pd

from modules.validator import (
    load_excel_safe,
    drop_non_employee_rows,
    validate_not_empty,
    validate_required_columns,
)
from modules.cleaner import (
    convert_brutss_column,
    _normalize_numcpt,
)
from modules.trimestrial_types import EmployeeKey, MonthlyEntry

BRUTSS_COLUMN = "BRUTSS"


def _float_to_cents(value: float) -> int:
    """
    Convert a float BRUTSS value to integer cents.

    Uses round() to handle floating-point imprecision:
    e.g. 1234.56 * 100 might be 123455.99999... → round → 123456

    Parameters
    ----------
    value : float
        BRUTSS value in euros/currency units.

    Returns
    -------
    int
        Value in integer cents.
    """
    return round(value * 100)


def _build_employee_key(normalized_numcpt: str) -> EmployeeKey:
    """
    Build a matching key from the normalized NUMCPT.

    Uses NUMCPT only — NOM/PRENOM are excluded from the key to avoid
    false negatives caused by name typos or encoding differences
    between files (e.g. "MEYMOUNA" vs "MIMOUNA").

    Parameters
    ----------
    normalized_numcpt : str
        Already-normalized NUMCPT value (leading zeros stripped).

    Returns
    -------
    EmployeeKey
    """
    return EmployeeKey(numcpt=normalized_numcpt)


def parse_monthly_file(file_obj, filename: str) -> dict:
    """
    Load, validate, and parse a single monthly Excel file.

    Pipeline:
    1. Load with openpyxl (safe loading, TEXT_DTYPE for NUMCPT)
    2. Drop non-employee rows (blank/total rows)
    3. Validate not empty + required columns
    4. Convert BRUTSS to float64
    5. Normalize NUMCPT for matching
    6. Group by EmployeeKey, summing BRUTSS for duplicates within the file
    7. Convert to integer cents

    If the same employee (NUMCPT+NOM+PRENOM) appears multiple times in the
    same monthly file, their BRUTSS values are summed first.

    Parameters
    ----------
    file_obj : file-like object
        Uploaded Excel file (from Streamlit st.file_uploader).
    filename : str
        Original filename, used in error messages.

    Returns
    -------
    dict[EmployeeKey, MonthlyEntry]
        One entry per unique employee. BRUTSS is in integer cents.

    Raises
    ------
    ValueError
        If the file cannot be read, is empty, or has invalid data.
    """
    # Step 1: Load Excel
    df = load_excel_safe(file_obj, filename)

    # Step 2: Drop non-employee rows
    df, _dropped = drop_non_employee_rows(df)

    # Step 3: Validate
    validate_not_empty(df, filename)
    validate_required_columns(df, filename)

    # Step 4: Convert BRUTSS to float64
    df = convert_brutss_column(df, filename)

    # Step 5: Normalize NUMCPT
    df["_NUMCPT_NORM"] = _normalize_numcpt(df["NUMCPT"])

    # Step 6 & 7: Group by key, sum BRUTSS, convert to cents
    lookup: dict = {}

    for _, row in df.iterrows():
        norm_numcpt = row["_NUMCPT_NORM"]

        # Skip rows with empty key
        if not norm_numcpt or str(norm_numcpt).strip() == "":
            continue

        key = _build_employee_key(norm_numcpt)
        brutss_cents = _float_to_cents(float(row[BRUTSS_COLUMN]))

        if key in lookup:
            # Same employee appears multiple times → sum BRUTSS
            existing = lookup[key]
            lookup[key] = MonthlyEntry(
                numcpt_raw=existing.numcpt_raw,
                nom=existing.nom,
                prenom=existing.prenom,
                brutss_cents=existing.brutss_cents + brutss_cents,
            )
        else:
            # First occurrence — preserve original NUMCPT (with leading zeros)
            numcpt_raw = str(row.get("NUMCPT", "")).strip()
            nom = str(row.get("NOM", "")).strip()
            prenom = str(row.get("PRENOM", "")).strip()

            lookup[key] = MonthlyEntry(
                numcpt_raw=numcpt_raw,
                nom=nom,
                prenom=prenom,
                brutss_cents=brutss_cents,
            )

    return lookup


def validate_trimestrial_files(file1, file2, file3) -> None:
    """
    Validate that all three monthly files are uploaded.

    Parameters
    ----------
    file1, file2, file3 : file-like object or None

    Raises
    ------
    ValueError
        If any file is missing.
    """
    missing = []
    if file1 is None:
        missing.append("Mois 1")
    if file2 is None:
        missing.append("Mois 2")
    if file3 is None:
        missing.append("Mois 3")

    if missing:
        raise ValueError(
            f"Fichier(s) manquant(s) : {', '.join(missing)}.\n"
            "Veuillez téléverser les trois fichiers mensuels pour continuer."
        )
