"""
splitter.py

Split a single Excel file into several files based on the distinct values of
a user-selected column.

Example: splitting by the ADM column produces one file per ADM value
(ADM=J, ADM=S, ADM=3, …). Each output file keeps all original columns and
only the rows matching that value. Rows where the column is empty go into a
dedicated "VIDE" file.

Output: a ZIP archive (bytes) containing one .xlsx per distinct value,
ready to download.
"""

import io
import os
import re
import zipfile

import pandas as pd

# Force these ID-like columns to text so leading zeros / long numbers survive
TEXT_DTYPE = {
    "NUMCPT": str, "NIN": str, "NUMSS": str, "NUMMUT": str,
    "MATRI": str, "NOMATRI": str, "CODBANK": str, "CLECPT": str,
}

# Safety guard: refuse to split into an unreasonable number of files
MAX_FILES = 500

# Token used in the filename when the split value is empty
EMPTY_TOKEN = "VIDE"


def get_columns(file_obj, filename: str) -> list:
    """
    Return the list of column names of the first sheet (header only, fast read).

    Raises
    ------
    ValueError
        If the file cannot be read.
    """
    file_obj.seek(0)
    try:
        head = pd.read_excel(file_obj, engine="openpyxl", nrows=0)
    except Exception as e:
        raise ValueError(
            f"Cannot read file '{filename}': {e}\n"
            "Please make sure it is a valid, unencrypted .xlsx file."
        ) from e
    return [str(c).strip() for c in head.columns]


def _safe_filename_token(value: str) -> str:
    """Turn a column value into a filesystem-safe filename token."""
    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none", "nat"):
        return EMPTY_TOKEN
    # Replace characters illegal in filenames on Windows/macOS
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = s.replace(" ", "_")
    return s[:80] or EMPTY_TOKEN


def split_by_column(file_obj, filename: str, column: str) -> tuple:
    """
    Split the file into one .xlsx per distinct value of `column`.

    Parameters
    ----------
    file_obj : file-like object
    filename : str
    column : str
        Name of the column to split on.

    Returns
    -------
    tuple (zip_bytes: bytes, stats: dict)
        stats = {
            "column": str,
            "total_rows": int,
            "file_count": int,
            "groups": list of {"value", "rows", "filename"},
        }

    Raises
    ------
    ValueError
        If the file can't be read, the column is missing, the file is empty,
        or the split would produce more than MAX_FILES files.
    """
    file_obj.seek(0)
    try:
        df = pd.read_excel(file_obj, engine="openpyxl", dtype=TEXT_DTYPE)
    except Exception as e:
        raise ValueError(
            f"Cannot read file '{filename}': {e}\n"
            "Please make sure it is a valid, unencrypted .xlsx file."
        ) from e

    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how="all").reset_index(drop=True)

    if column not in df.columns:
        raise ValueError(
            f"Column '{column}' not found in file '{filename}'. "
            f"Available columns: {', '.join(df.columns)}"
        )

    if df.empty:
        raise ValueError(f"File '{filename}' contains no data rows.")

    # Grouping key: trimmed string value, empty/NaN → EMPTY_TOKEN
    keys = df[column].astype(str).str.strip()
    keys = keys.mask(
        df[column].isna() | keys.str.lower().isin(["", "nan", "none", "nat"]),
        EMPTY_TOKEN,
    )

    distinct = keys.nunique()
    if distinct > MAX_FILES:
        raise ValueError(
            f"Splitting by '{column}' would create {distinct} files, which is "
            f"more than the limit of {MAX_FILES}. Please choose a column with "
            "fewer distinct values (e.g. a category/code column)."
        )

    base = os.path.splitext(os.path.basename(filename))[0]

    groups_info = []
    used_names = set()
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        # Sort values for deterministic output; EMPTY_TOKEN naturally sorts among them
        for value, sub in df.groupby(keys, sort=True):
            token = _safe_filename_token(value)
            out_name = f"{base}_{column}_{token}.xlsx"

            # Avoid collisions when two raw values sanitize to the same token
            if out_name in used_names:
                n = 2
                while f"{base}_{column}_{token}_{n}.xlsx" in used_names:
                    n += 1
                out_name = f"{base}_{column}_{token}_{n}.xlsx"
            used_names.add(out_name)

            sheet_buf = io.BytesIO()
            with pd.ExcelWriter(sheet_buf, engine="openpyxl") as writer:
                sub.to_excel(writer, index=False, sheet_name=token[:31] or "Data")
            zf.writestr(out_name, sheet_buf.getvalue())

            groups_info.append({
                "value": value,
                "rows": len(sub),
                "filename": out_name,
            })

    stats = {
        "column": column,
        "total_rows": len(df),
        "file_count": len(groups_info),
        "groups": groups_info,
    }

    return zip_buffer.getvalue(), stats
