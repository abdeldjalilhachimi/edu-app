"""
trimestrial_merger.py

Merge service for the trimestrial BRUTSS consolidation.

Takes three monthly lookups (parsed by trimestrial_parser) and produces
a single consolidated result with one row per unique employee.

Column names are derived from uploaded filenames (e.g. BRUTSS_paie_janvier).

All arithmetic is done in integer cents to avoid floating-point errors.
"""

from modules.trimestrial_types import (
    EmployeeKey,
    MonthlyEntry,
    MissingEmployee,
    ConsolidatedRow,
    TrimestrialStats,
    TrimestrialResult,
    filename_to_label,
)


def merge_trimestrial(
    month1: dict,
    month2: dict,
    month3: dict,
    filename1: str,
    filename2: str,
    filename3: str,
) -> TrimestrialResult:
    """
    Merge three monthly lookups into a consolidated trimestrial result.

    For each unique employee (identified by NUMCPT only):
    - BRUTSS for each month (0 if employee is missing in that month)
    - BRUTSS_TOTAL = sum of the three months

    Column labels are derived from the uploaded filenames (without extension).

    Parameters
    ----------
    month1 : dict[EmployeeKey, MonthlyEntry]
        Parsed Month 1 data.
    month2 : dict[EmployeeKey, MonthlyEntry]
        Parsed Month 2 data.
    month3 : dict[EmployeeKey, MonthlyEntry]
        Parsed Month 3 data.
    filename1 : str
        Original filename for Month 1 (e.g. "paie_janvier.xlsx").
    filename2 : str
        Original filename for Month 2.
    filename3 : str
        Original filename for Month 3.

    Returns
    -------
    TrimestrialResult
        Consolidated rows + summary statistics (including file labels).
    """
    # Derive clean labels from filenames
    label1 = filename_to_label(filename1)
    label2 = filename_to_label(filename2)
    label3 = filename_to_label(filename3)

    # Collect all unique keys across all three months
    all_keys = set(month1.keys()) | set(month2.keys()) | set(month3.keys())

    # For each key, pick the best representative identity (NOM, PRENOM, NUMCPT_RAW).
    # Priority: Month 1 > Month 2 > Month 3 (first file where the employee appears).
    rows: list = []
    total_m1 = 0
    total_m2 = 0
    total_m3 = 0

    # Track employees missing from each file
    missing_from_file1: list = []
    missing_from_file2: list = []
    missing_from_file3: list = []

    months = [month1, month2, month3]

    for key in sorted(all_keys, key=lambda k: k.numcpt):
        entry1 = month1.get(key)
        entry2 = month2.get(key)
        entry3 = month3.get(key)

        # Pick representative identity from first available month
        representative = entry1 or entry2 or entry3

        m1_cents = entry1.brutss_cents if entry1 else 0
        m2_cents = entry2.brutss_cents if entry2 else 0
        m3_cents = entry3.brutss_cents if entry3 else 0
        total_cents = m1_cents + m2_cents + m3_cents

        # Pick best value for optional fields: first non-empty across months (M1 > M2 > M3)
        numss = ""
        adm = ""
        datnais = ""
        nbrtrav = ""
        datent = ""
        datsor = ""
        for entry in (entry1, entry2, entry3):
            if entry:
                if entry.numss and not numss:
                    numss = entry.numss
                if entry.adm and not adm:
                    adm = entry.adm
                if entry.datnais and not datnais:
                    datnais = entry.datnais
                if entry.nbrtrav and not nbrtrav:
                    nbrtrav = entry.nbrtrav
                if entry.datent and not datent:
                    datent = entry.datent
                if entry.datsor and not datsor:
                    datsor = entry.datsor

        rows.append(ConsolidatedRow(
            numcpt_raw=representative.numcpt_raw,
            nom=representative.nom,
            prenom=representative.prenom,
            numss=numss,
            adm=adm,
            datnais=datnais,
            nbrtrav=nbrtrav,
            datent=datent,
            datsor=datsor,
            monthly_brutss_cents=(m1_cents, m2_cents, m3_cents),
            brutss_total=total_cents,
        ))

        total_m1 += m1_cents
        total_m2 += m2_cents
        total_m3 += m3_cents

        # Detect missing: employee exists globally but absent from this file
        missing_info = MissingEmployee(
            numcpt_raw=representative.numcpt_raw,
            nom=representative.nom,
            prenom=representative.prenom,
        )
        if entry1 is None:
            missing_from_file1.append(missing_info)
        if entry2 is None:
            missing_from_file2.append(missing_info)
        if entry3 is None:
            missing_from_file3.append(missing_info)

    grand_total = total_m1 + total_m2 + total_m3

    stats = TrimestrialStats(
        unique_count=len(rows),
        monthly_totals_cents=(total_m1, total_m2, total_m3),
        grand_total_cents=grand_total,
        file_labels=(label1, label2, label3),
    )

    return TrimestrialResult(
        rows=rows,
        stats=stats,
        missing_per_file=(missing_from_file1, missing_from_file2, missing_from_file3),
    )
