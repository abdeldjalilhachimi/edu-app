"""
txt_converter.py

Convert an annual declaration Excel file (.xlsx) into the official
fixed-width "Détail" record text file (ANNEXE 2).

Each output line is a fixed-width record of 194 characters: every field is
padded/truncated to its exact size and concatenated with NO separator.

Field layout (Enregistrement Détail):
     # Field                              Type            Size
     1 N° Employeur                       Alphanumérique   10
     2 Année Réf                          Alphanumérique    4
     3 N° Ligne                           Alphanumérique    6
     4 N° Immatriculation + Clé (NUMSS)   Alphanumérique   12
     5 Nom                                Alphanumérique   25
     6 Prénom                             Alphanumérique   25
     7 Date de Naissance                  JJMMAAAA          8
   8-19 Per quarter ×4: Durée(3) Unité(1) Salaire(10)
    20 Montant Annuel Des Salaires        Alphanumérique   12
    21 Date Entrée                        JJMMAAAA          8
    22 Date Sortie                        JJMMAAAA          8
    23 Observation                        Alphanumérique   20

Formatting conventions (adjust here if the administration differs):
  - Text fields:    left-justified, space-padded on the right, truncated.
  - Numeric fields (N° Ligne, Durée): right-justified, zero-padded.
  - Amount fields (Salaire, Montant): value in CENTIMES (×100, no decimal
    point), right-justified, zero-padded.
  - Dates:          JJMMAAAA (8 digits, no separators); empty → 8 spaces.
  - Unité de mesure: "J" (jours).
  - Encoding:       Windows-1252 (single-byte, so 1 char = 1 byte → widths
    stay aligned). Line terminator: CRLF.
"""

import io
import pandas as pd

from modules.validator import load_excel_safe, drop_non_employee_rows
from modules.date_utils import format_ddmmyyyy

# Required identity columns in the annual file
REQUIRED_COLUMNS = {"NUMSS", "NOM", "PRENOM"}

RECORD_WIDTH = 194
FIELD_COUNT = 23
OUTPUT_ENCODING = "cp1252"
LINE_TERMINATOR = "\r\n"
UNIT_OF_MEASURE = "J"

# Sizes
_SZ = {
    "NEMPLOYEUR": 10, "ANREF": 4, "N": 6, "NUMSS": 12,
    "NOM": 25, "PRENOM": 25, "DATNAIS": 8,
    "DUREE": 3, "UNITE": 1, "SALAIRE": 10,
    "MONTANT": 12, "DATENT": 8, "DATSOR": 8, "OBSERV": 20,
}

DEFAULT_EMPLOYER = "0840198947"


# ── Field formatters ─────────────────────────────────────────────────────────

def _alnum(value, size: int) -> str:
    """Left-justified, space-padded, truncated text field."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        s = ""
    else:
        s = str(value).strip()
        if s.lower() in ("nan", "none", "nat"):
            s = ""
    return s[:size].ljust(size)


def _num(value, size: int) -> str:
    """Right-justified, zero-padded integer field (e.g. line number, days)."""
    try:
        n = int(round(float(str(value).strip())))
    except (ValueError, TypeError):
        n = 0
    if n < 0:
        n = 0
    return str(n).rjust(size, "0")[-size:]


def _amount(value, size: int) -> str:
    """Amount in centimes (×100), right-justified, zero-padded."""
    try:
        cents = int(round(float(str(value).strip()) * 100))
    except (ValueError, TypeError):
        cents = 0
    if cents < 0:
        cents = 0
    return str(cents).rjust(size, "0")[-size:]


def _date8(value) -> str:
    """JJMMAAAA (8 digits). Empty/unparseable → 8 spaces."""
    s = format_ddmmyyyy(value)          # → DD/MM/YYYY or '' or original
    digits = s.replace("/", "")
    if len(digits) == 8 and digits.isdigit():
        return digits
    return " " * 8


# ── Column detection ─────────────────────────────────────────────────────────

def _quarter_columns(df: pd.DataFrame) -> tuple:
    """
    Return (salary_cols, day_cols), each a list of up to 4 column names in
    quarter order, padded with None to length 4.

    Salary columns:  BRUTSS_<label> (excluding BRUTSS_ANNUEL / BRUTSS_TOTAL).
    Day columns:     NBRTRAV_<label> (the plain NBRTRAV annual column is excluded).
    """
    salary_cols = [
        c for c in df.columns
        if c.startswith("BRUTSS_") and c not in ("BRUTSS_ANNUEL", "BRUTSS_TOTAL")
    ]
    day_cols = [c for c in df.columns if c.startswith("NBRTRAV_")]

    salary_cols = (salary_cols + [None, None, None, None])[:4]
    day_cols = (day_cols + [None, None, None, None])[:4]
    return salary_cols, day_cols


def _annual_amount_column(df: pd.DataFrame) -> str:
    """Return the annual total column name (BRUTSS_ANNUEL or BRUTSS_TOTAL)."""
    if "BRUTSS_ANNUEL" in df.columns:
        return "BRUTSS_ANNUEL"
    if "BRUTSS_TOTAL" in df.columns:
        return "BRUTSS_TOTAL"
    return ""


def _validate_columns(df: pd.DataFrame, filename: str) -> None:
    present = set(df.columns.tolist())
    missing = REQUIRED_COLUMNS - present
    if missing or not _annual_amount_column(df):
        need = sorted(missing) + ([] if _annual_amount_column(df) else ["BRUTSS_ANNUEL"])
        raise ValueError(
            f"File '{filename}' is missing required column(s): "
            f"{', '.join(need)}.\n"
            f"Columns found: {', '.join(sorted(present))}\n"
            "Please use the annual declaration file produced by the "
            "'Déclaration Annuelle' tab."
        )


# ── Record builder ───────────────────────────────────────────────────────────

def _build_record(row, line_no: int, salary_cols, day_cols, annual_col) -> str:
    """Build one 194-char fixed-width Détail record from a DataFrame row."""
    parts = [
        _alnum(row.get("NEMPLOYEUR", DEFAULT_EMPLOYER) or DEFAULT_EMPLOYER, _SZ["NEMPLOYEUR"]),
        _alnum(row.get("ANREF", ""), _SZ["ANREF"]),
        _num(line_no, _SZ["N"]),
        _alnum(row.get("NUMSS", ""), _SZ["NUMSS"]),
        _alnum(row.get("NOM", ""), _SZ["NOM"]),
        _alnum(row.get("PRENOM", ""), _SZ["PRENOM"]),
        _date8(row.get("DATNAIS", "")),
    ]

    # 4 quarters: Durée(3) + Unité(1) + Salaire(10)
    for q in range(4):
        day_col = day_cols[q]
        sal_col = salary_cols[q]
        days = row.get(day_col, 0) if day_col else 0
        salary = row.get(sal_col, 0) if sal_col else 0
        parts.append(_num(days, _SZ["DUREE"]))
        parts.append(_alnum(UNIT_OF_MEASURE, _SZ["UNITE"]))
        parts.append(_amount(salary, _SZ["SALAIRE"]))

    parts.append(_amount(row.get(annual_col, 0), _SZ["MONTANT"]))
    parts.append(_date8(row.get("DATENT", "")))
    parts.append(_date8(row.get("DATSOR", "")))
    parts.append(_alnum(row.get("OBSERV", ""), _SZ["OBSERV"]))

    return "".join(parts)


def convert_xlsx_to_txt(file_obj, filename: str) -> tuple:
    """
    Convert an annual declaration Excel file to the official fixed-width
    "Détail" text file.

    Returns
    -------
    tuple[bytes, dict]
        - txt_bytes: the fixed-width file (cp1252-encoded).
        - info: {total_rows, record_width, field_count, encoding,
                 quarter_salary_cols, quarter_day_cols, missing_day_columns}.
    """
    df = load_excel_safe(file_obj, filename)
    df, _ = drop_non_employee_rows(df)

    if df.empty:
        raise ValueError(
            f"File '{filename}' contains no data rows. The file appears to be empty."
        )

    _validate_columns(df, filename)

    salary_cols, day_cols = _quarter_columns(df)
    annual_col = _annual_amount_column(df)
    missing_days = all(c is None for c in day_cols)

    lines = []
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        record = _build_record(row, i, salary_cols, day_cols, annual_col)
        # Safety: guarantee exact width
        record = record[:RECORD_WIDTH].ljust(RECORD_WIDTH)
        lines.append(record)

    txt_content = LINE_TERMINATOR.join(lines) + LINE_TERMINATOR
    txt_bytes = txt_content.encode(OUTPUT_ENCODING, errors="replace")

    info = {
        "total_rows": len(df),
        "record_width": RECORD_WIDTH,
        "field_count": FIELD_COUNT,
        "encoding": OUTPUT_ENCODING,
        "quarter_salary_cols": [c for c in salary_cols if c],
        "quarter_day_cols": [c for c in day_cols if c],
        "missing_day_columns": missing_days,
    }
    return txt_bytes, info
