"""
trimestrial_parser.py

Helper parsing functions for the trimestrial BRUTSS consolidation.

Loads and validates each monthly Excel file, then extracts employee
records with BRUTSS converted to integer cents for safe arithmetic.

Reuses the existing validator and cleaner modules for file loading,
column validation, and BRUTSS string-to-float conversion.
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
    _normalize_numcpt,
    EMPTY_BRUTSS_COLUMN,
)
from modules.trimestrial_types import EmployeeKey, MonthlyEntry

BRUTSS_COLUMN = "BRUTSS"

# The "days worked" value is stored under a different column name depending on
# the file type. Each file carries exactly one of these; we use the first one
# present, by this priority, and map it to NBRTRAV.
#   PAIE → NBRTRAV   PRIME → JRPRIME   RAPPEL → NBRJ   RAPPEL PRIM → NBRJ/JRPRIME
NBRTRAV_ALIASES = ["NBRTRAV", "JRPRIME", "NBRJ"]


def _resolve_nbrtrav(df: pd.DataFrame) -> pd.Series:
    """
    Return the days-worked column as an integer Series, regardless of which
    alias the file uses (NBRTRAV / JRPRIME / NBRJ). Missing → all zeros.
    """
    for alias in NBRTRAV_ALIASES:
        if alias in df.columns:
            return pd.to_numeric(df[alias], errors="coerce").fillna(0).astype(int)
    return pd.Series(0, index=df.index, dtype=int)


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


def parse_monthly_file(file_obj, filename: str) -> tuple:
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
    tuple[dict[EmployeeKey, MonthlyEntry], list[tuple]]
        - lookup: One entry per unique employee. BRUTSS is in integer cents.
        - empty_brutss: List of (numcpt, nom, prenom) for employees with empty BRUTSS.

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

    # Extract empty-BRUTSS employees before further processing
    empty_brutss = []
    if EMPTY_BRUTSS_COLUMN in df.columns:
        empty_mask = df[EMPTY_BRUTSS_COLUMN] == True
        if empty_mask.any():
            for _, row in df[empty_mask].iterrows():
                empty_brutss.append((
                    str(row.get("NUMCPT", "")).strip(),
                    str(row.get("NOM", "")).strip(),
                    str(row.get("PRENOM", "")).strip(),
                ))
        df = df.drop(columns=[EMPTY_BRUTSS_COLUMN])

    # Step 5: Normalize NUMCPT
    df["_NUMCPT_NORM"] = _normalize_numcpt(df["NUMCPT"])

    # Step 6: Filter empty keys, clean optional columns, groupby, convert to cents

    # Filter rows with empty normalized NUMCPT (vectorized)
    valid_mask = df["_NUMCPT_NORM"].astype(str).str.strip().ne("")
    df = df[valid_mask]

    if df.empty:
        return {}, empty_brutss

    # Vectorized cleanup: identity columns
    df["NUMCPT"] = df["NUMCPT"].astype(str).str.strip()
    df["NOM"] = df["NOM"].astype(str).str.strip()
    df["PRENOM"] = df["PRENOM"].astype(str).str.strip()

    # Vectorized cleanup: optional string columns — fillna + strip + replace nan/None
    _OPT_STR_COLS = ["NUMSS", "ADM", "DATNAIS", "DATENT", "DATSOR"]
    for col_name in _OPT_STR_COLS:
        if col_name in df.columns:
            df[col_name] = (
                df[col_name].fillna("").astype(str).str.strip()
                .replace({"nan": "", "None": ""})
            )
        else:
            df[col_name] = ""

    # NBRTRAV: read from whichever alias the file uses (NBRTRAV/JRPRIME/NBRJ),
    # as integer days worked — will be summed per employee.
    df["NBRTRAV"] = _resolve_nbrtrav(df)

    # Group by normalized NUMCPT: sum BRUTSS + NBRTRAV, keep first row for identity
    group_key = "_NUMCPT_NORM"
    brutss_sums = df.groupby(group_key)[BRUTSS_COLUMN].sum()
    nbrtrav_sums = df.groupby(group_key)["NBRTRAV"].sum()
    first_rows = df.groupby(group_key).first()

    # Build lookup dict from grouped result (iterating over small deduplicated set)
    lookup: dict = {}
    for norm_numcpt, row in first_rows.iterrows():
        key = _build_employee_key(norm_numcpt)
        brutss_cents = _float_to_cents(float(brutss_sums[norm_numcpt]))

        lookup[key] = MonthlyEntry(
            numcpt_raw=row["NUMCPT"],
            nom=row["NOM"],
            prenom=row["PRENOM"],
            numss=row.get("NUMSS", ""),
            adm=row.get("ADM", ""),
            datnais=row.get("DATNAIS", ""),
            nbrtrav=int(nbrtrav_sums[norm_numcpt]),
            datent=row.get("DATENT", ""),
            datsor=row.get("DATSOR", ""),
            brutss_cents=brutss_cents,
        )

    return lookup, empty_brutss


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
