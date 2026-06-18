"""
annual_merger.py

Merge 4 quarterly files into a single annual declaration.

For each unique employee (by NUMCPT):
- BRUTSS from each quarter (0 if absent from that quarter)
- BRUTSS_ANNUEL = sum of all 4 quarters
- Identity fields (NOM, PRENOM, NUMSS, ADM, etc.) picked from first non-empty quarter

Same pattern as trimestrial_merger.py but for 4 files instead of 3.
"""

from modules.annual_types import (
    EmployeeKey,
    QuarterlyEntry,
    MissingEmployee,
    AnnualRow,
    AnnualStats,
    AnnualResult,
    filename_to_label,
)
from modules.date_utils import format_ddmmyyyy


def _pick_best(values: list) -> str:
    """Return the first non-empty string from a list of values."""
    for v in values:
        if v and str(v).strip():
            return str(v).strip()
    return ""


def _sum_days(values: list) -> int:
    """
    Sum NBRTRAV (days worked) across quarters as integers.

    Values arrive as strings (e.g. "15", "0", "15.0", ""); non-numeric or
    empty entries count as 0. Used instead of _pick_best so days accumulate
    over the year rather than taking the first quarter's value.
    """
    total = 0
    for v in values:
        s = str(v).strip()
        if not s or s.lower() in ("nan", "none"):
            continue
        try:
            total += int(round(float(s)))
        except (ValueError, TypeError):
            continue
    return total


def merge_annual(
    q1: dict, q2: dict, q3: dict, q4: dict,
    fname1: str, fname2: str, fname3: str, fname4: str,
) -> AnnualResult:
    """
    Merge 4 quarterly employee lookups into one annual result.

    Parameters
    ----------
    q1, q2, q3, q4 : dict[EmployeeKey, QuarterlyEntry]
        One lookup per quarter, from parse_quarterly_file().
    fname1, fname2, fname3, fname4 : str
        Original filenames, used for column labels and absence detection.

    Returns
    -------
    AnnualResult
        Consolidated rows, stats, and per-file missing employee lists.
    """
    quarters = [q1, q2, q3, q4]
    file_labels = tuple(filename_to_label(f) for f in (fname1, fname2, fname3, fname4))

    # Collect all unique employee keys
    all_keys = set()
    for q in quarters:
        all_keys.update(q.keys())

    # Build consolidated rows
    rows = []
    quarterly_totals = [0, 0, 0, 0]

    for key in sorted(all_keys, key=lambda k: k.numcpt):
        entries = [q.get(key) for q in quarters]

        # Identity: pick first non-empty across quarters
        non_none = [e for e in entries if e is not None]
        if not non_none:
            continue

        first = non_none[0]
        numcpt_raw = first.numcpt_raw
        nom = _pick_best([e.nom for e in non_none])
        prenom = _pick_best([e.prenom for e in non_none])
        numss = _pick_best([e.numss for e in non_none])
        adm = _pick_best([e.adm for e in non_none])
        datnais = format_ddmmyyyy(_pick_best([e.datnais for e in non_none]))
        datent = format_ddmmyyyy(_pick_best([e.datent for e in non_none]))
        datsor = format_ddmmyyyy(_pick_best([e.datsor for e in non_none]))

        # Days worked per quarter (0 if absent that quarter)
        q_days = tuple(_sum_days([e.nbrtrav]) if e is not None else 0 for e in entries)
        # Total days across all quarters
        nbrtrav = sum(q_days)

        # BRUTSS per quarter (0 if absent)
        q_cents = tuple(
            e.brutss_total_cents if e is not None else 0
            for e in entries
        )
        brutss_annual = sum(q_cents)

        # Accumulate per-quarter totals
        for i, c in enumerate(q_cents):
            quarterly_totals[i] += c

        rows.append(AnnualRow(
            numcpt_raw=numcpt_raw,
            nom=nom,
            prenom=prenom,
            numss=numss,
            adm=adm,
            datnais=datnais,
            nbrtrav=nbrtrav,
            datent=datent,
            datsor=datsor,
            quarterly_brutss_cents=q_cents,
            quarterly_nbrtrav=q_days,
            brutss_annual=brutss_annual,
        ))

    # Detect missing employees per file
    missing_per_file = []
    for i, q in enumerate(quarters):
        missing = []
        for key in sorted(all_keys, key=lambda k: k.numcpt):
            if key not in q:
                # Find identity from another quarter
                for other_q in quarters:
                    if key in other_q:
                        e = other_q[key]
                        missing.append(MissingEmployee(
                            numcpt_raw=e.numcpt_raw,
                            nom=e.nom,
                            prenom=e.prenom,
                        ))
                        break
        missing_per_file.append(missing)

    # Stats
    grand_total = sum(quarterly_totals)
    stats = AnnualStats(
        unique_count=len(rows),
        quarterly_totals_cents=tuple(quarterly_totals),
        grand_total_cents=grand_total,
        file_labels=file_labels,
    )

    return AnnualResult(
        rows=rows,
        stats=stats,
        missing_per_file=tuple(missing_per_file),
    )
