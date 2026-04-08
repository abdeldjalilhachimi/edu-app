"""
txt_converter.py

Convert an annual declaration Excel file (.xlsx) to a pipe-delimited
text file (.txt) for administrative submission.

Reads the first sheet of the uploaded file, drops the individual
quarterly BRUTSS columns (BRUTSS_<Q1>, BRUTSS_<Q2>, ...), and keeps
only the annual total (BRUTSS_ANNUEL) alongside all other columns.

Output format: pipe-delimited (|), UTF-8 encoded, one row per line.
"""

import io
import pandas as pd

from modules.validator import load_excel_safe, drop_non_employee_rows


# Columns that must be present in the annual file
REQUIRED_COLUMNS = {"NUMCPT", "NOM", "PRENOM", "BRUTSS_ANNUEL"}

# The final column order for the txt output
# (matches the annual declaration structure minus quarterly breakdown)
FINAL_COLUMNS = [
    "NEMPLOYEUR",
    "ANREF",
    "N",
    "NUMCPT",
    "NUMSS",
    "ADM",
    "NOM",
    "PRENOM",
    "DATNAIS",
    "NBRTRAV",
    "DATENT",
    "DATSOR",
    "BRUTSS_ANNUEL",
    "UNBRTRAV",
    "OBSERV",
]


def _validate_columns(df: pd.DataFrame, filename: str) -> None:
    """Validate that the required columns are present."""
    present = set(df.columns.tolist())
    missing = REQUIRED_COLUMNS - present

    if missing:
        missing_list = ", ".join(f'"{c}"' for c in sorted(missing))
        found_list = ", ".join(f'"{c}"' for c in sorted(present))
        raise ValueError(
            f"File '{filename}' is missing required column(s): {missing_list}.\n"
            f"Columns found in file: {found_list}\n"
            "Please make sure this is an annual declaration file containing BRUTSS_ANNUEL."
        )


def _format_value(val) -> str:
    """
    Format a single cell value for txt output.

    - NaN / None → empty string
    - Float with no decimals (e.g. 1234.0) → integer string "1234"
    - Float with decimals → 2 decimal places "1234.56"
    - Everything else → stripped string
    """
    if pd.isna(val):
        return ""

    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return f"{val:.2f}"

    s = str(val).strip()
    return "" if s in ("nan", "None") else s


def convert_xlsx_to_txt(file_obj, filename: str) -> tuple:
    """
    Convert an annual declaration Excel file to pipe-delimited txt.

    Pipeline:
    1. Load Excel (first sheet)
    2. Validate required columns
    3. Drop quarterly BRUTSS columns (any column matching BRUTSS_* except BRUTSS_ANNUEL)
    4. Reorder to final column order (skip missing optional columns)
    5. Export as pipe-delimited text

    Parameters
    ----------
    file_obj : file-like object
        Uploaded annual Excel file.
    filename : str
        Original filename, used in error messages.

    Returns
    -------
    tuple[bytes, dict]
        - txt_bytes: The pipe-delimited text file content (UTF-8).
        - info: Dict with stats {total_rows, columns_kept, columns_dropped}.
    """
    # Step 1: Load
    df = load_excel_safe(file_obj, filename)

    # Step 2: Drop non-employee rows
    df, _dropped = drop_non_employee_rows(df)

    if df.empty or len(df) == 0:
        raise ValueError(
            f"File '{filename}' contains no data rows. "
            "The file appears to be empty."
        )

    # Step 3: Validate
    _validate_columns(df, filename)

    # Step 4: Identify and drop quarterly BRUTSS columns
    # Quarterly columns match pattern BRUTSS_* but are NOT BRUTSS_ANNUEL or BRUTSS_TOTAL
    all_cols = df.columns.tolist()
    quarterly_cols = [
        c for c in all_cols
        if c.startswith("BRUTSS_")
        and c not in ("BRUTSS_ANNUEL", "BRUTSS_TOTAL")
    ]

    # If file has BRUTSS_TOTAL instead of BRUTSS_ANNUEL, rename it
    if "BRUTSS_TOTAL" in df.columns and "BRUTSS_ANNUEL" not in df.columns:
        df = df.rename(columns={"BRUTSS_TOTAL": "BRUTSS_ANNUEL"})

    df_clean = df.drop(columns=quarterly_cols, errors="ignore")

    # Step 5: Reorder columns to the final order
    # Keep only columns that exist in the DataFrame
    ordered_cols = [c for c in FINAL_COLUMNS if c in df_clean.columns]

    # Also keep any extra columns not in FINAL_COLUMNS (placed at the end)
    extra_cols = [c for c in df_clean.columns if c not in FINAL_COLUMNS]
    final_col_order = ordered_cols + extra_cols

    df_out = df_clean[final_col_order]

    # Step 6: Build pipe-delimited text (data only, no header)
    lines = []

    # Data rows
    for _, row in df_out.iterrows():
        values = [_format_value(row[col]) for col in df_out.columns]
        lines.append("|".join(values))

    txt_content = "\n".join(lines) + "\n"
    txt_bytes = txt_content.encode("utf-8")

    info = {
        "total_rows": len(df_out),
        "columns_kept": len(df_out.columns),
        "columns_dropped": quarterly_cols,
        "columns_dropped_count": len(quarterly_cols),
    }

    return txt_bytes, info
