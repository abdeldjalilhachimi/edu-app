"""
cleaner.py

Normalize the BRUTSS column from raw string or numeric values to float64.
All conversion failures are collected and reported together (no silent errors).

Also builds the composite key column (_KEY) from NOM, PRENOM, NUMCPT.
"""

import re
import pandas as pd
from typing import Union

BRUTSS_COLUMN = "BRUTSS"
KEY_COLUMN = "_KEY"
EMPTY_BRUTSS_COLUMN = "_BRUTSS_WAS_EMPTY"


def detect_and_clean_number_string(raw: str) -> str:
    """
    Convert a raw number string to a clean string parseable by float().

    Handles the following real-world formats:
    - "1,234.56"  → Anglo-Saxon (comma=thousands, dot=decimal)   → "1234.56"
    - "1 234,56"  → French/European (space=thousands, comma=decimal) → "1234.56"
    - "1.234,56"  → European variant (dot=thousands, comma=decimal) → "1234.56"
    - "1234.56"   → Standard dot-decimal                           → "1234.56"
    - "1234,56"   → French decimal (comma=decimal)                 → "1234.56"
    - "1 234"     → French integer with thousands space            → "1234"
    - "1,234"     → Anglo-Saxon integer with thousands comma       → "1234"
    - "1234"      → Plain integer                                  → "1234"

    Detection algorithm:
    1. Strip outer whitespace; replace non-breaking spaces (\u00a0) with regular space.
    2. If both comma AND dot present: the one appearing LAST is the decimal separator.
       - dot last  → remove all commas (they were thousands) → standard decimal
       - comma last → remove all dots (they were thousands), replace comma with dot
    3. If only comma (no dot):
       - comma + exactly 1 or 2 trailing digits → decimal separator → replace with dot
       - otherwise → thousands separator → remove
    4. If only dot (no comma):
       - multiple dots → invalid (e.g. "1.2.3") → raise ValueError
       - single dot → standard decimal → leave as-is
    5. Remove remaining spaces (they are thousands separators at this point).
    6. Final float() check; raise ValueError if still unparseable.

    Parameters
    ----------
    raw : str

    Returns
    -------
    str
        A clean string that float() can parse directly.

    Raises
    ------
    ValueError
        If the string cannot be interpreted as a number after all cleaning.
    """
    s = str(raw).strip()

    # Replace non-breaking spaces with regular spaces
    s = s.replace("\u00a0", " ")

    has_comma = "," in s
    has_dot = "." in s

    if has_comma and has_dot:
        last_comma_pos = s.rfind(",")
        last_dot_pos = s.rfind(".")

        if last_dot_pos > last_comma_pos:
            # Dot is decimal: "1,234.56" or "1,234,567.89"
            s = s.replace(",", "")
        else:
            # Comma is decimal: "1.234,56" or "1.234.567,89"
            s = s.replace(".", "").replace(",", ".")

    elif has_comma and not has_dot:
        # Comma is either thousands or decimal separator
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1].strip()) in (1, 2):
            # e.g. "1234,5" or "1234,56" → decimal
            s = s.replace(",", ".")
        else:
            # e.g. "1,234" or "1,234,567" → thousands separators
            s = s.replace(",", "")

    elif has_dot and not has_comma:
        # Standard dot-decimal; multiple dots = invalid
        dot_count = s.count(".")
        if dot_count > 1:
            raise ValueError(
                f"Cannot interpret '{raw}' as a number: "
                f"multiple dots found (e.g. '1.2.3' is not valid)."
            )
        # else: single dot → leave as-is

    # Remove remaining spaces (thousands separators)
    s = s.replace(" ", "")

    # Final validation
    try:
        float(s)
    except ValueError:
        raise ValueError(
            f"Cannot interpret '{raw}' as a number after cleaning "
            f"(cleaned form: '{s}')."
        )

    return s


def convert_brutss_column(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    """
    Convert the BRUTSS column to float64 in place.

    - Already-numeric cells (int, float) are passed through unchanged.
    - String cells go through detect_and_clean_number_string().
    - NaN / None / empty values are treated as 0 (zero BRUTSS).
    - ALL conversion failures are collected before raising, so the user
      sees every problematic row at once rather than stopping at row 1.

    Parameters
    ----------
    df : pd.DataFrame
        Must already have a 'BRUTSS' column (validated upstream).
    filename : str
        Used in error messages.

    Returns
    -------
    pd.DataFrame
        Same DataFrame with BRUTSS column as float64.

    Raises
    ------
    ValueError
        If one or more values fail conversion. Lists every bad row.
    """
    col = df[BRUTSS_COLUMN]

    # ── Fast path: column is already numeric (int/float dtype) ──────────
    if pd.api.types.is_numeric_dtype(col):
        empty_mask = col.isna()
        result_df = df.copy()
        result_df[BRUTSS_COLUMN] = col.fillna(0.0).astype("float64")
        result_df[EMPTY_BRUTSS_COLUMN] = empty_mask
        return result_df

    # ── Slow path: mixed types — separate numeric from string values ────
    # Try pd.to_numeric first; values that fail become NaN
    numeric_attempt = pd.to_numeric(col, errors="coerce")

    # Identify originally-NaN values → treat as 0
    orig_na_mask = col.isna()

    # Track all empty BRUTSS positions
    empty_mask = orig_na_mask.copy()

    # Identify string values that need cleaning (not NaN originally, but failed numeric)
    needs_cleaning = numeric_attempt.isna() & ~orig_na_mask

    errors = []
    result_series = numeric_attempt.copy()

    # Empty/NaN BRUTSS → 0
    result_series[orig_na_mask] = 0.0

    # Clean string values that couldn't be converted directly
    if needs_cleaning.any():
        for idx in col.index[needs_cleaning]:
            raw_value = col.iloc[idx] if isinstance(col.index, pd.RangeIndex) else col.loc[idx]
            # Empty strings → 0
            raw_str = str(raw_value).strip()
            if raw_str == "" or raw_str.lower() == "nan":
                result_series.iat[col.index.get_loc(idx)] = 0.0
                empty_mask.iat[col.index.get_loc(idx)] = True
                continue
            try:
                cleaned = detect_and_clean_number_string(raw_str)
                result_series.iat[col.index.get_loc(idx)] = float(cleaned)
            except ValueError as e:
                errors.append(f"  Row {idx + 2} (index {idx}): '{raw_value}' → {e}")

    if errors:
        error_detail = "\n".join(errors)
        raise ValueError(
            f"File '{filename}': Failed to convert BRUTSS values to numbers.\n"
            f"The following rows have invalid values:\n"
            f"{error_detail}\n"
            "Please fix the source file and re-upload."
        )

    result_df = df.copy()
    result_df[BRUTSS_COLUMN] = result_series.fillna(0.0).astype("float64")
    result_df[EMPTY_BRUTSS_COLUMN] = empty_mask
    return result_df


def _normalize_numcpt(series: pd.Series) -> pd.Series:
    """
    Normalize NUMCPT values so they match regardless of source format.

    Problem: NUMCPT may be stored as text with leading zeros in one file
    (e.g. '007999990009457003') and as an integer in another file
    (e.g. 7999990009457003). The composite key would never match.

    Solution: convert to string, strip whitespace, strip leading zeros.
    Both '007999990009457003' and 7999990009457003 become '7999990009457003'.

    Edge case: a NUMCPT of '000' or 0 becomes '0' (not empty string).
    """
    return (
        series.astype(str)
        .str.strip()
        .str.lstrip("0")
        .replace("", "0")  # if the entire value was zeros
    )


def build_composite_key(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add a '_KEY' column based on NUMCPT alone.

    NUMCPT is 100% unique across employees, so NOM/PRENOM are not needed
    in the key. This avoids false negatives caused by name typos or
    encoding differences between files.

    NUMCPT is normalized: stripped of whitespace and leading zeros so that
    text ('007999...') and integer (7999...) representations match.

    Parameters
    ----------
    df : pd.DataFrame
        Must have NUMCPT column (validated upstream).

    Returns
    -------
    pd.DataFrame
        Copy of df with an added '_KEY' column.
    """
    return df.assign(**{KEY_COLUMN: _normalize_numcpt(df["NUMCPT"])})


def clean_all_dataframes(
    main_df: pd.DataFrame,
    additional_dfs: list,
    main_filename: str,
    additional_filenames: list,
) -> tuple:
    """
    Apply BRUTSS conversion and composite key building to all DataFrames.

    Stops at the first file that has conversion errors.

    Parameters
    ----------
    main_df : pd.DataFrame
    additional_dfs : list[pd.DataFrame]
    main_filename : str
    additional_filenames : list[str]

    Returns
    -------
    tuple[pd.DataFrame, list[pd.DataFrame]]
        All DataFrames with BRUTSS as float64 and _KEY column added.

    Raises
    ------
    ValueError
        Propagated from convert_brutss_column.
    """
    # Clean main file
    main_clean = convert_brutss_column(main_df, main_filename)
    main_clean = build_composite_key(main_clean)

    # Clean each additional file
    additional_clean = []
    for df, filename in zip(additional_dfs, additional_filenames):
        cleaned = convert_brutss_column(df, filename)
        cleaned = build_composite_key(cleaned)
        additional_clean.append(cleaned)

    return main_clean, additional_clean
