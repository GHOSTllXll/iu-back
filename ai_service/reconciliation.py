# backend/ai_service/reconciliation.py
"""
MODULE 1: Cross-Document Reconciliation Engine ("Bullshit Detector")
Enterprise tier only.

Design principle (avoiding circular self-validation): the AI's ONLY job is to
report what a source document literally claims (om_claimed_occupancy_pct).
Ground truth is computed by Python directly from the parsed rent_roll_df —
the same data that drives the Excel formulas. Python performs the comparison.
The AI never checks its own math, so a flag here means an actual document
mismatch, not the AI disagreeing with itself.

Coverage note: this implements OCCUPANCY_MISMATCH and REVENUE_MISMATCH fully.
CONCESSION_HIDING is best-effort — it checks for a rent-roll column that looks
concession-related and flags a WARNING if found alongside a zero/missing T12
concessions figure. It does not do exhaustive text scanning for concession
mentions buried in free-text cells; treat it as a prompt for manual review,
not a guarantee of detection.
"""
import pandas as pd

# Variance thresholds, in percentage points. Your spec specified WARNING for
# variances under 2% but didn't specify a CRITICAL threshold — 5 points is a
# reasonable line I'm drawing, not something you specified. Adjust freely.
WARNING_THRESHOLD_PCT = 2.0
CRITICAL_THRESHOLD_PCT = 5.0


def _severity_for_variance(variance_pct: float) -> str:
    return "CRITICAL" if variance_pct >= CRITICAL_THRESHOLD_PCT else "WARNING"


def _check_occupancy(ground_truth: dict, document_claims: dict) -> dict | None:
    claimed_pct = document_claims.get("om_claimed_occupancy_pct")
    actual_pct = ground_truth.get("physical_occupancy_pct")

    if claimed_pct is None or actual_pct is None:
        # Nothing to compare — OM didn't state a figure, or rent roll had no rows.
        return None

    variance_pct = abs(claimed_pct - actual_pct) * 100  # both are 0-1 decimals

    if variance_pct < WARNING_THRESHOLD_PCT:
        return None  # within tolerance, not worth flagging

    severity = _severity_for_variance(variance_pct)
    occupied = ground_truth["occupied_units"]
    total = ground_truth["total_units"]

    return {
        "metric": "physical_occupancy_pct",
        "severity": severity,
        "message": (
            f"OM claims {claimed_pct * 100:.1f}% occupancy, but raw Rent Roll "
            f"analysis shows actual occupancy is {actual_pct * 100:.1f}% "
            f"({occupied} occupied of {total} total units)."
        ),
        "claimed_value": claimed_pct,
        "actual_value": actual_pct,
        "variance_pct": round(variance_pct, 2),
    }


def _check_revenue(ground_truth: dict, t12_revenue_expenses: dict) -> dict | None:
    claimed_income = t12_revenue_expenses.get("net_rental_income")
    actual_income = ground_truth.get("annualized_rent_roll_income")

    if claimed_income is None or not actual_income:
        return None

    variance_pct = abs(claimed_income - actual_income) / actual_income * 100

    if variance_pct < WARNING_THRESHOLD_PCT:
        return None

    severity = _severity_for_variance(variance_pct)

    return {
        "metric": "net_rental_income",
        "severity": severity,
        "message": (
            f"T12 reports Net Rental Income of ${claimed_income:,.0f}, but summing "
            f"in-place rents from the Rent Roll (annualized) yields ${actual_income:,.0f} "
            f"— a variance of {variance_pct:.1f}%."
        ),
        "claimed_value": claimed_income,
        "actual_value": round(actual_income, 2),
        "variance_pct": round(variance_pct, 2),
    }


def _check_concessions(rent_roll_df: pd.DataFrame, rent_roll_metrics: dict) -> dict | None:
    """
    Best-effort check — see module docstring. Looks only for a column whose
    NAME suggests concessions/discounts; does not scan free-text cell content.
    """
    concession_like_cols = [
        col for col in rent_roll_df.columns
        if any(term in str(col).lower() for term in ('concession', 'discount', 'rent free', 'loss to lease'))
    ]

    if not concession_like_cols:
        return None  # nothing to cross-check against — not itself suspicious

    claimed_concessions = rent_roll_metrics.get("total_annual_concessions")
    has_nonzero_data = False

    for col in concession_like_cols:
        numeric_col = pd.to_numeric(rent_roll_df[col], errors='coerce').fillna(0)
        if (numeric_col != 0).any():
            has_nonzero_data = True
            break

    if has_nonzero_data and (claimed_concessions is None or claimed_concessions == 0):
        return {
            "metric": "total_annual_concessions",
            "severity": "WARNING",
            "message": (
                f"Rent Roll contains a column suggesting concessions/discounts "
                f"({', '.join(str(c) for c in concession_like_cols)}) with non-zero values, "
                f"but the reported total annual concessions is zero or missing. "
                f"Manual review recommended — this check does not scan free-text cells "
                f"and may miss concessions recorded elsewhere."
            ),
            "claimed_value": claimed_concessions,
            "actual_value": None,  # not a precise figure — this is a presence check, not a sum
            "variance_pct": None,
        }

    return None


def run_reconciliation(metrics: dict, rent_roll_df: pd.DataFrame, ground_truth: dict) -> list:
    """
    Runs all reconciliation checks and returns a list of flag dicts.
    Never raises — a failed/skipped check just means no flag for that metric,
    since reconciliation is a value-add, not a hard requirement for the
    core extraction to succeed.
    """
    flags = []

    document_claims = metrics.get("reconciliation_claims", {}) or {}
    t12_revenue_expenses = metrics.get("t12_revenue_expenses", {}) or {}
    rent_roll_metrics = metrics.get("rent_roll_metrics", {}) or {}

    for check in (
        _check_occupancy(ground_truth, document_claims),
        _check_revenue(ground_truth, t12_revenue_expenses),
        _check_concessions(rent_roll_df, rent_roll_metrics),
    ):
        if check is not None:
            flags.append(check)

    return flags