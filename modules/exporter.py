"""
exporter.py

Build the output Excel workbook with two sheets using pure openpyxl.

Sheet 1 ("Résultat"): Main file data with updated BRUTSS values.
Sheet 2 ("Résumé"):   Per-file totals, grand total, duplicate list.

Number formatting is applied via openpyxl cell.number_format so that
BRUTSS values remain numeric in Excel (sortable, usable in formulas)
while displaying in French format: 1 234,56.

French Excel number format code: '# ##0,00'
  Space = thousands separator
  Comma = decimal separator
  Always 2 decimal places
"""

import io
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# French number format: "1 234 567,89"
FRENCH_NUMBER_FORMAT = "# ##0,00"

# Styling constants
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
SECTION_FONT = Font(bold=True)
SECTION_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
GRAND_TOTAL_FONT = Font(bold=True)
GRAND_TOTAL_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

# Internal key column — excluded from all output
KEY_COLUMN = "_KEY"
BRUTSS_COLUMN = "BRUTSS"


def _auto_column_widths(ws) -> None:
    """Set each column width to fit the widest cell value (with a cap)."""
    for col_cells in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_length + 4, 50)


def build_result_sheet(ws, df: pd.DataFrame) -> None:
    """
    Write the updated main DataFrame to an openpyxl worksheet.

    - Excludes the _KEY column (internal use only).
    - Applies FRENCH_NUMBER_FORMAT to every BRUTSS cell.
    - Values are stored as Python float (numeric in Excel, not text).
    - Header row is bold white on blue background.

    Parameters
    ----------
    ws : openpyxl Worksheet
    df : pd.DataFrame
        Main DataFrame with updated BRUTSS (float64) and _KEY column.
    """
    # Columns to export (drop internal _KEY)
    export_cols = [c for c in df.columns if c != KEY_COLUMN]
    brutss_col_idx = export_cols.index(BRUTSS_COLUMN) + 1  # 1-based

    # Write header row
    for col_idx, col_name in enumerate(export_cols, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    # Write data rows
    for row_idx, row_data in enumerate(df[export_cols].itertuples(index=False), start=2):
        for col_idx, (col_name, value) in enumerate(
            zip(export_cols, row_data), start=1
        ):
            # Convert numpy types to native Python for openpyxl compatibility
            if hasattr(value, "item"):
                value = value.item()

            cell = ws.cell(row=row_idx, column=col_idx, value=value)

            if col_name == BRUTSS_COLUMN:
                cell.number_format = FRENCH_NUMBER_FORMAT
                cell.alignment = Alignment(horizontal="right")

    _auto_column_widths(ws)


def build_summary_sheet(
    ws,
    duplicates: list,
    stats: dict,
) -> None:
    """
    Build Sheet 2 — Summary Report.

    Layout:
      Section A: Matched rows table (NUMCPT, NOM, PRENOM, BRUTSS breakdown)
      Section B: Stats summary (total rows, matches, final BRUTSS total)

    Parameters
    ----------
    ws : openpyxl Worksheet
    duplicates : list[DuplicateMatch]
    stats : dict  {total, duplicate_count, brutss_total}
    """
    current_row = 1

    # ── Section A: Matched rows (duplicates) ────────────────────────────────

    ws.cell(row=current_row, column=1, value="CORRESPONDANCES TROUVÉES")
    ws.cell(row=current_row, column=1).font = SECTION_FONT
    ws.cell(row=current_row, column=1).fill = SECTION_FILL
    ws.merge_cells(
        start_row=current_row, start_column=1,
        end_row=current_row, end_column=6
    )
    current_row += 1

    if not duplicates:
        ws.cell(
            row=current_row, column=1,
            value="Aucune correspondance trouvée entre le fichier principal et les fichiers additionnels."
        )
        ws.cell(row=current_row, column=1).font = Font(italic=True, color="548235")
        current_row += 2
    else:
        # Sub-header count
        ws.cell(
            row=current_row, column=1,
            value=f"{len(duplicates)} correspondance(s) trouvée(s)"
        )
        ws.cell(row=current_row, column=1).font = Font(bold=True, color="2F5496")
        current_row += 1

        # Column headers
        headers = [
            "NUMCPT", "NOM", "PRENOM",
            "BRUTSS Principal", "BRUTSS Additionnels", "BRUTSS Somme"
        ]
        header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=header)
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        current_row += 1

        # Data rows
        for match in duplicates:
            ws.cell(row=current_row, column=1, value=match.numcpt)
            ws.cell(row=current_row, column=2, value=match.nom)
            ws.cell(row=current_row, column=3, value=match.prenom)

            for col_idx, val in [(4, match.brutss_main), (5, match.brutss_additional), (6, match.brutss_sum)]:
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.number_format = FRENCH_NUMBER_FORMAT
                cell.alignment = Alignment(horizontal="right")

            current_row += 1

        current_row += 1  # blank separator

    # ── Section B: Stats summary ─────────────────────────────────────────────

    ws.cell(row=current_row, column=1, value="RÉSUMÉ")
    ws.cell(row=current_row, column=1).font = SECTION_FONT
    ws.cell(row=current_row, column=1).fill = SECTION_FILL
    ws.merge_cells(
        start_row=current_row, start_column=1,
        end_row=current_row, end_column=2
    )
    current_row += 1

    stat_rows = [
        ("Lignes dans le fichier principal", stats["total"]),
        ("Correspondances trouvées", stats["duplicate_count"]),
    ]
    for label, value in stat_rows:
        ws.cell(row=current_row, column=1, value=label)
        ws.cell(row=current_row, column=2, value=value)
        current_row += 1

    # Grand total with special formatting
    label_cell = ws.cell(row=current_row, column=1, value="Total BRUTSS final")
    label_cell.font = GRAND_TOTAL_FONT
    label_cell.fill = GRAND_TOTAL_FILL
    value_cell = ws.cell(row=current_row, column=2, value=stats["brutss_total"])
    value_cell.number_format = FRENCH_NUMBER_FORMAT
    value_cell.font = GRAND_TOTAL_FONT
    value_cell.fill = GRAND_TOTAL_FILL
    value_cell.alignment = Alignment(horizontal="right")

    _auto_column_widths(ws)


def create_output_excel(
    updated_main_df: pd.DataFrame,
    duplicates: list,
    stats: dict,
) -> bytes:
    """
    Assemble the complete two-sheet output workbook and return as bytes.

    No temp files are used — serializes directly to an in-memory BytesIO buffer.

    Parameters
    ----------
    updated_main_df : pd.DataFrame
    duplicates : list[DuplicateMatch]
    stats : dict  {total, duplicate_count, brutss_total}

    Returns
    -------
    bytes
        Raw .xlsx bytes, ready for st.download_button(data=...).
    """
    wb = openpyxl.Workbook()

    # Sheet 1 — Result
    ws1 = wb.active
    ws1.title = "Résultat"
    build_result_sheet(ws1, updated_main_df)

    # Sheet 2 — Summary
    ws2 = wb.create_sheet(title="Résumé")
    build_summary_sheet(ws2, duplicates, stats)

    # Serialize to bytes
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
