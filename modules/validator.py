"""
validator.py

Strict pre-processing validation for uploaded Excel files.
All functions raise ValueError with a descriptive, user-friendly message on failure.
The UI layer catches these and displays them cleanly.
"""

import pandas as pd

# Required columns in every uploaded file
REQUIRED_COLUMNS = {"BRUTSS", "NOM", "PRENOM", "NUMCPT"}


def load_excel_safe(file_obj, filename: str) -> pd.DataFrame:
    """
    Read an Excel file into a DataFrame, catching all parsing errors.

    Strips leading/trailing whitespace from all column names on load
    to prevent invisible-whitespace false negatives during validation.

    Parameters
    ----------
    file_obj : file-like object
        Uploaded file object (e.g. from Streamlit st.file_uploader).
    filename : str
        Original filename, used only in error messages.

    Returns
    -------
    pd.DataFrame
        Raw DataFrame with no data transformations applied.

    Raises
    ------
    ValueError
        If the file cannot be parsed (corrupted, wrong format, encrypted, etc.)
    """
    try:
        df = pd.read_excel(file_obj, sheet_name=0, engine="openpyxl")
    except Exception as e:
        raise ValueError(
            f"Cannot read file '{filename}': {e}\n"
            "Please make sure it is a valid, unencrypted .xlsx file."
        ) from e

    # Normalize column names: strip surrounding whitespace
    df.columns = df.columns.astype(str).str.strip()

    # Drop rows where every single cell is NaN (trailing blank rows from Excel)
    df = df.dropna(how="all").reset_index(drop=True)

    return df


def drop_non_employee_rows(df: pd.DataFrame) -> tuple:
    """
    Remove rows that are clearly not employee records:
    - Rows where NOM, PRENOM, and NUMCPT are all empty/NaN
      (e.g. grand-total rows appended by payroll software).

    Returns the cleaned DataFrame and a count of rows dropped.
    """
    key_cols = [c for c in ["NOM", "PRENOM", "NUMCPT"] if c in df.columns]
    if not key_cols:
        return df, 0

    mask_all_key_empty = df[key_cols].isnull().all(axis=1) | (
        df[key_cols].astype(str).apply(lambda col: col.str.strip()) == ""
    ).all(axis=1)

    dropped = mask_all_key_empty.sum()
    cleaned = df[~mask_all_key_empty].reset_index(drop=True)
    return cleaned, int(dropped)


def validate_not_empty(df: pd.DataFrame, filename: str) -> None:
    """
    Ensure the DataFrame has at least one data row.

    Parameters
    ----------
    df : pd.DataFrame
    filename : str

    Raises
    ------
    ValueError
        If the DataFrame has no rows (headers-only file).
    """
    if df.empty or len(df) == 0:
        raise ValueError(
            f"File '{filename}' contains no data rows. "
            "The file appears to be empty (headers only or completely blank)."
        )


def validate_required_columns(df: pd.DataFrame, filename: str) -> None:
    """
    Verify that all required columns are present in the DataFrame.

    Required columns: BRUTSS, NOM, PRENOM, NUMCPT.
    Comparison is case-sensitive (column names must match exactly).

    Parameters
    ----------
    df : pd.DataFrame
    filename : str

    Raises
    ------
    ValueError
        If one or more required columns are missing. The error message lists
        exactly which columns are missing and which were found.
    """
    present = set(df.columns.tolist())
    missing = REQUIRED_COLUMNS - present

    if missing:
        missing_list = ", ".join(f'"{c}"' for c in sorted(missing))
        found_list = ", ".join(f'"{c}"' for c in sorted(present))
        raise ValueError(
            f"File '{filename}' is missing required column(s): {missing_list}.\n"
            f"Columns found in file: {found_list}"
        )


def validate_all_files(
    main_file,
    main_filename: str,
    additional_files: list,
    additional_filenames: list,
) -> tuple:
    """
    Orchestrate full validation for the main file and all additional files.

    Runs load_excel_safe → validate_not_empty → validate_required_columns
    for each file in sequence. Stops and raises at the first error encountered.

    Parameters
    ----------
    main_file : file-like object
    main_filename : str
    additional_files : list of file-like objects
    additional_filenames : list of str

    Returns
    -------
    tuple[pd.DataFrame, list[pd.DataFrame]]
        (main_df, [additional_df_1, additional_df_2, ...])
        All DataFrames are raw — no cleaning applied yet.

    Raises
    ------
    ValueError
        Propagated from any sub-validator with filename context included.
    """
    if not additional_files:
        raise ValueError(
            "No additional files uploaded. "
            "Please upload at least one additional .xlsx file."
        )

    # Validate main file
    main_df = load_excel_safe(main_file, main_filename)
    main_df, dropped_main = drop_non_employee_rows(main_df)
    validate_not_empty(main_df, main_filename)
    validate_required_columns(main_df, main_filename)

    # Validate each additional file
    additional_dfs = []
    dropped_counts = {main_filename: dropped_main}
    for file_obj, filename in zip(additional_files, additional_filenames):
        df = load_excel_safe(file_obj, filename)
        df, dropped = drop_non_employee_rows(df)
        dropped_counts[filename] = dropped
        validate_not_empty(df, filename)
        validate_required_columns(df, filename)
        additional_dfs.append(df)

    return main_df, additional_dfs, dropped_counts
