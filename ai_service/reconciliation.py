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

Coverage note: this implements OCCUPANCY_MISMATCH, REVENUE_MISMATCH, and
VACANCY_LOSS_VS_OCCUPANCY ("Rule B") fully. CONCESSION_HIDING is best-effort —
it checks for a rent-roll column that looks concession-related and flags a
WARNING if found alongside a zero/missing T12 concessions figure. It does not
do exhaustive text scanning for concession mentions buried in free-text
cells; treat it as a prompt for manual review, not a guarantee of detection.
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


PHYSICAL_OCCUPANCY_HIGH_THRESHOLD_PCT = 95.0  # percentage points
VACANCY_LOSS_HIGH_THRESHOLD_PCT = 8.0  # percentage points, as stated on the T12


def _check_vacancy_loss_vs_occupancy(ground_truth: dict, document_claims: dict) -> dict | None:
    """
    Rule B: flags the specific anomaly pattern where a property LOOKS fully
    occupied on the Rent Roll but the T12 itself reports meaningfully higher
    vacancy/concession loss than that would suggest — a mismatch that often
    points to delinquency being miscoded as "vacancy", a stale Rent Roll, or
    concessions masking weaker real economic performance.

    Unlike the other checks in this module, this isn't a single-value
    variance comparison (there's no one "true" number both sides should
    agree on) — it's a conjunction of two independent conditions, each
    checked against its own fixed threshold. Severity is derived from how
    far the T12's claimed vacancy loss exceeds its own 8% threshold, reusing
    _severity_for_variance for a consistent WARNING/CRITICAL split with the
    rest of this module rather than inventing a parallel scale.
    """
    actual_occupancy_pct = ground_truth.get("physical_occupancy_pct")
    claimed_vacancy_loss_pct = document_claims.get("t12_claimed_vacancy_loss_pct")

    if actual_occupancy_pct is None or claimed_vacancy_loss_pct is None:
        return None  # rent roll had no rows, or the T12 didn't state a figure

    occupancy_points = actual_occupancy_pct * 100
    vacancy_loss_points = claimed_vacancy_loss_pct * 100

    if occupancy_points <= PHYSICAL_OCCUPANCY_HIGH_THRESHOLD_PCT:
        return None  # not a "looks fully occupied" property — Rule B doesn't apply
    if vacancy_loss_points <= VACANCY_LOSS_HIGH_THRESHOLD_PCT:
        return None  # vacancy loss is in line with high occupancy — no anomaly

    excess_pct = vacancy_loss_points - VACANCY_LOSS_HIGH_THRESHOLD_PCT
    severity = _severity_for_variance(excess_pct)

    return {
        "metric": "t12_claimed_vacancy_loss_pct",
        "severity": severity,
        "message": (
            f"Rent Roll shows {occupancy_points:.1f}% physical occupancy (a fully-occupied "
            f"property), but the T12 reports {vacancy_loss_points:.1f}% economic vacancy/"
            f"concession loss — above the {VACANCY_LOSS_HIGH_THRESHOLD_PCT:.0f}% threshold "
            f"expected at this occupancy level. Worth checking whether delinquency is being "
            f"coded as vacancy, the Rent Roll is stale, or concessions are masking weaker "
            f"economic performance."
        ),
        "claimed_value": claimed_vacancy_loss_pct,
        "actual_value": actual_occupancy_pct,
        "variance_pct": round(excess_pct, 2),
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


RECONCILIATION_DELTA_THRESHOLD = 0.01  # dollars — effectively "must match exactly"


def run_ppm_reconciliation(raw_sources_and_uses_items, claimed_master_acquisition_cost) -> dict:
    """
    PPM equivalent of this module's core design principle above: the AI's
    ONLY job is to extract the raw "Sources and Uses" line items and the
    document's claimed master acquisition cost figure — it never grades its
    own arithmetic or produces a color verdict. Python performs the
    summation and comparison here, same as the CRE checks above.

    Returns a tri-state result rather than a flat GREEN/RED boolean, because
    "some line items didn't parse" and "the numbers genuinely don't balance"
    are different problems that shouldn't look identical to the user:
      - GREEN: parsed cleanly AND the totals match within threshold.
      - RED: parsed cleanly, but the totals genuinely don't match — a real
        mismatch in the sponsor's own document math.
      - NEEDS_REVIEW: one or more line items (or the acquisition cost
        itself) couldn't be parsed as a number — a data-quality gap, not a
        proven mismatch. Takes priority over RED/GREEN regardless of what
        the parseable items happen to sum to, since a partial sum isn't a
        trustworthy basis for either verdict.

    Never raises — malformed AI output degrades into unparseable_items
    entries (and a NEEDS_REVIEW flag) rather than crashing the request,
    same defensive posture as every other AI-output-facing function in this
    module and in variance_analysis.py/parsers.py.
    """
    total_sources_and_uses = 0.0
    unparseable_items = []

    for item in (raw_sources_and_uses_items or []):
        if not isinstance(item, dict):
            unparseable_items.append({"item": item, "error": "Line item was not an object."})
            continue
        try:
            amount = float(item.get('amount'))
        except (TypeError, ValueError):
            unparseable_items.append({
                "item": item,
                "error": f"Non-numeric or missing 'amount': {item.get('amount')!r}",
            })
            continue
        total_sources_and_uses += amount

    acquisition_cost = None
    try:
        acquisition_cost = float(claimed_master_acquisition_cost)
    except (TypeError, ValueError):
        unparseable_items.append({
            "item": {"claimed_master_acquisition_cost": claimed_master_acquisition_cost},
            "error": "Non-numeric or missing claimed_master_acquisition_cost.",
        })

    if unparseable_items:
        flag = "NEEDS_REVIEW"
        delta = None
    else:
        delta = abs(total_sources_and_uses - acquisition_cost)
        flag = "GREEN" if delta < RECONCILIATION_DELTA_THRESHOLD else "RED"

    return {
        "flag": flag,
        "total_sources_and_uses": round(total_sources_and_uses, 2),
        "claimed_acquisition_cost": round(acquisition_cost, 2) if acquisition_cost is not None else None,
        "delta": round(delta, 2) if delta is not None else None,
        "unparseable_items": unparseable_items,
    }


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
        _check_vacancy_loss_vs_occupancy(ground_truth, document_claims),
        _check_concessions(rent_roll_df, rent_roll_metrics),
    ):
        if check is not None:
            flags.append(check)

    return flags