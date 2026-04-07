"""
payroll_calculator.py

Core business logic for the payroll calculation feature (Tab 3).

Classifies employees as confirmed (titulaire) or non-confirmed (non titulaire)
based on their NUMCPT, then applies the appropriate rates:

  - RETSS = BRUTSS × 9%     (both categories)
  - PARTSS = BRUTSS × 25%   (confirmed) or BRUTSS × 12.5% (non-confirmed)
  - NETPAI = read from file  (passed through unchanged)
"""

from modules.payroll_types import (
    PayrollEmployee,
    PayrollResult,
    PayrollStats,
    PayrollOutput,
    NON_CONFIRMED_NUMCPTS,
    RETSS_RATE_PERCENT,
    PARTSS_CONFIRMED_PERCENT,
    PARTSS_NON_CONFIRMED_PERMILLE,
)


def _compute_retss(brutss_cents: int) -> int:
    """RETSS = BRUTSS × 9% (both categories)."""
    return round(brutss_cents * RETSS_RATE_PERCENT / 100)


def _compute_partss_confirmed(brutss_cents: int) -> int:
    """PARTSS = BRUTSS × 25% (confirmed employees)."""
    return round(brutss_cents * PARTSS_CONFIRMED_PERCENT / 100)


def _compute_partss_non_confirmed(brutss_cents: int) -> int:
    """PARTSS = BRUTSS × 12.5% (non-confirmed employees)."""
    return round(brutss_cents * PARTSS_NON_CONFIRMED_PERMILLE / 1000)


def calculate_payroll(employees: list) -> PayrollOutput:
    """
    Classify employees and apply payroll rates.

    Parameters
    ----------
    employees : list[PayrollEmployee]
        Parsed employees from the input file.

    Returns
    -------
    PayrollOutput
        Confirmed rows, non-confirmed rows, and aggregate stats.
    """
    confirmed_rows: list = []
    non_confirmed_rows: list = []

    for emp in employees:
        is_confirmed = emp.numcpt_norm not in NON_CONFIRMED_NUMCPTS

        retss = _compute_retss(emp.brutss_cents)

        if is_confirmed:
            partss = _compute_partss_confirmed(emp.brutss_cents)
        else:
            partss = _compute_partss_non_confirmed(emp.brutss_cents)

        result = PayrollResult(
            numcpt_raw=emp.numcpt_raw,
            nom=emp.nom,
            prenom=emp.prenom,
            is_confirmed=is_confirmed,
            brutss_cents=emp.brutss_cents,
            retss_cents=retss,
            partss_cents=partss,
            netpai_cents=emp.netpai_cents,
        )

        if is_confirmed:
            confirmed_rows.append(result)
        else:
            non_confirmed_rows.append(result)

    # ── Aggregate stats ─────────────────────────────────────────────────────

    def _sum_field(rows, field):
        return sum(getattr(r, field) for r in rows)

    c_brutss = _sum_field(confirmed_rows, "brutss_cents")
    c_retss = _sum_field(confirmed_rows, "retss_cents")
    c_partss = _sum_field(confirmed_rows, "partss_cents")
    c_netpai = _sum_field(confirmed_rows, "netpai_cents")

    nc_brutss = _sum_field(non_confirmed_rows, "brutss_cents")
    nc_retss = _sum_field(non_confirmed_rows, "retss_cents")
    nc_partss = _sum_field(non_confirmed_rows, "partss_cents")
    nc_netpai = _sum_field(non_confirmed_rows, "netpai_cents")

    stats = PayrollStats(
        total_employees=len(employees),
        confirmed_count=len(confirmed_rows),
        non_confirmed_count=len(non_confirmed_rows),
        confirmed_brutss_cents=c_brutss,
        confirmed_retss_cents=c_retss,
        confirmed_partss_cents=c_partss,
        confirmed_netpai_cents=c_netpai,
        non_confirmed_brutss_cents=nc_brutss,
        non_confirmed_retss_cents=nc_retss,
        non_confirmed_partss_cents=nc_partss,
        non_confirmed_netpai_cents=nc_netpai,
        grand_brutss_cents=c_brutss + nc_brutss,
        grand_retss_cents=c_retss + nc_retss,
        grand_partss_cents=c_partss + nc_partss,
        grand_netpai_cents=c_netpai + nc_netpai,
    )

    return PayrollOutput(
        confirmed_rows=confirmed_rows,
        non_confirmed_rows=non_confirmed_rows,
        stats=stats,
    )
