"""
validator.py

Strict pre-processing validation for uploaded Excel files.
All functions raise ValueError with a descriptive, user-friendly message on failure.
The UI layer catches these and displays them cleanly.
"""

import pandas as pd
import openpyxl

# Required columns in every uploaded file
REQUIRED_COLUMNS = {"BRUTSS", "NOM", "PRENOM", "NUMCPT"}

# Force these columns to be read as text (avoid precision loss / scientific notation)
TEXT_DTYPE = {"NUMCPT": str, "NUMSS": str, "NUMMUT": str}


def load_excel_safe(file_obj, filename: str, sheet_name=0) -> pd.DataFrame:
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
        # Force ID columns to be read as string to preserve exact value:
        # - keeps leading zeros (e.g. "007999990025346277")
        # - avoids float64 precision loss on 16+ digit numbers
        df = pd.read_excel(
            file_obj, sheet_name=sheet_name, engine="openpyxl",
            dtype=TEXT_DTYPE,
        )
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


def get_sheet_names(file_obj) -> list:
    """
    Return the list of sheet names in an Excel file without reading all data.
    Resets file position after reading.
    """
    pos = file_obj.tell() if hasattr(file_obj, 'tell') else 0
    try:
        wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    finally:
        if hasattr(file_obj, 'seek'):
            file_obj.seek(pos)


def load_internal_sheets(file_obj, filename: str, main_sheet_index: int = 0) -> list:
    """
    Load all sheets EXCEPT the main sheet as additional DataFrames.

    Only returns sheets that contain the required columns (BRUTSS, NOM, PRENOM, NUMCPT).
    Sheets missing required columns are silently skipped.

    Parameters
    ----------
    file_obj : file-like object
    filename : str
    main_sheet_index : int
        Index of the main sheet (default 0 = first sheet).

    Returns
    -------
    list of (pd.DataFrame, str)
        Each tuple is (dataframe, sheet_label) for valid internal sheets.
    """
    file_obj.seek(0)
    sheet_names = get_sheet_names(file_obj)

    internal_sheets = []
    for idx, sheet_name in enumerate(sheet_names):
        if idx == main_sheet_index:
            continue  # skip main sheet

        file_obj.seek(0)
        try:
            df = load_excel_safe(file_obj, filename, sheet_name=sheet_name)
            df, _ = drop_non_employee_rows(df)

            # Check if this sheet has the required columns
            present = set(df.columns.tolist())
            if REQUIRED_COLUMNS.issubset(present) and len(df) > 0:
                label = f"{filename} → {sheet_name}"
                internal_sheets.append((df, label))
        except (ValueError, Exception):
            continue  # skip sheets that can't be read or are invalid

    return internal_sheets


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
    include_internal_sheets: bool = False,
) -> tuple:
    """
    Orchestrate full validation for the main file and all additional files.

    Runs load_excel_safe → validate_not_empty → validate_required_columns
    for each file in sequence. Stops and raises at the first error encountered.

    When include_internal_sheets=True, extra sheets in ALL uploaded files
    (main + additional) are automatically included as additional data.

    Parameters
    ----------
    main_file : file-like object
    main_filename : str
    additional_files : list of file-like objects
    additional_filenames : list of str
    include_internal_sheets : bool
        If True, extra sheets (beyond the first) in any uploaded file
        are treated as additional data sources.

    Returns
    -------
    tuple (main_df, additional_dfs, additional_names, dropped_counts)
    """
    if not additional_files and not include_internal_sheets:
        raise ValueError(
            "No additional files uploaded. "
            "Please upload at least one additional .xlsx file."
        )

    # Validate main file (first sheet)
    main_df = load_excel_safe(main_file, main_filename)
    main_df, dropped_main = drop_non_employee_rows(main_df)
    validate_not_empty(main_df, main_filename)
    validate_required_columns(main_df, main_filename)

    additional_dfs = []
    additional_names = []
    dropped_counts = {main_filename: dropped_main}

    # Load internal sheets from main file
    if include_internal_sheets:
        main_file.seek(0)
        for df, label in load_internal_sheets(main_file, main_filename):
            validate_required_columns(df, label)
            additional_dfs.append(df)
            additional_names.append(label)

    # Validate each additional file
    for file_obj, filename in zip(additional_files, additional_filenames):
        df = load_excel_safe(file_obj, filename)
        df, dropped = drop_non_employee_rows(df)
        dropped_counts[filename] = dropped
        validate_not_empty(df, filename)
        validate_required_columns(df, filename)
        additional_dfs.append(df)
        additional_names.append(filename)

        # Load internal sheets from additional files too
        if include_internal_sheets:
            file_obj.seek(0)
            for int_df, int_label in load_internal_sheets(file_obj, filename):
                validate_required_columns(int_df, int_label)
                additional_dfs.append(int_df)
                additional_names.append(int_label)

    if not additional_dfs:
        raise ValueError(
            "No additional data found. "
            "Please upload additional files or enable internal sheets."
        )

    return main_df, additional_dfs, additional_names, dropped_counts
