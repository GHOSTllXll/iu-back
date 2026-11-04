# backend/ai_service/variance_analysis.py
"""
Module 7: Multi-Family Variance Analysis Engine (Scoped & Deterministic).

Identifies month-over-month revenue drops in the T12 and attempts to
attribute recent drops (near the end of the trailing-12 period) to a cause
visible on the Rent Roll — delinquency, lease turnover, or concession
burn-off. Deliberately conservative: only flags what real data actually
supports, and says so explicitly when it can't.

DESIGN PRINCIPLES (see conversation history for full reasoning):
- All arithmetic (month-over-month deltas) is done in Python, not the AI —
  the AI has already demonstrated it cannot reliably sum/compare numbers
  extracted from raw document text.
- Categories A/B (turnover, concession burn-off) are only evaluated for
  dips near the END of the T12's own monthly sequence — a Rent Roll is a
  point-in-time snapshot, so it can only meaningfully explain what changed
  RECENTLY, not something that happened months ago and has since resolved.
- "Recent" is defined as the T12's own last 2 months, NOT a separately
  extracted "Rent Roll As-Of Date" — a T12 is inherently trailing up to the
  present, so this avoids needing another AI-extracted date field and stays
  fully deterministic.
- Category C (delinquency) is a point-in-time fact and doesn't have this
  limitation — it's evaluated whenever the Rent Roll has a usable column.
- If the T12 doesn't have real chronological monthly columns (annual-only),
  this entire module is skipped with an explicit, honest message rather
  than fabricating a trend that doesn't exist in the source data.
"""

import re
import pandas as pd


DROP_THRESHOLD_PCT = 0.10  # a month-over-month decline >= 10% counts as a "dip"
RECENT_WINDOW_MONTHS = 2   # only the last N months of the T12 are eligible for A/B

# Month name patterns we'll recognize in column headers, e.g. "Jan 2025",
# "January 2025", "Jan-25", "01/2025". Deliberately permissive — real T12
# exports vary a lot in header formatting.
_MONTH_PATTERN = re.compile(
    r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s\-/]?\d{2,4}',
    re.IGNORECASE
)


class VarianceAnalysisSkipped(Exception):
    """Raised when the source documents don't support this analysis at all."""
    def __init__(self, message):
        self.message = message
        super().__init__(message)


def detect_monthly_t12_columns(t12_df: pd.DataFrame):
    """
    Scans column headers for a run of real chronological month labels
    (e.g. "Jan 2025", "Feb 2025", ...). Returns the ordered list of matching
    column names if at least 6 are found (half a year — enough to be
    confident this is genuinely a monthly-column T12), else None.

    Deliberately does NOT require exactly 12 — some exports have 11 (partial
    year) or 13+ (includes a stub/forecast column) alongside the real ones.
    """
    matches = [col for col in t12_df.columns if _MONTH_PATTERN.search(str(col))]
    if len(matches) < 6:
        return None
    return matches


def find_revenue_row(df: pd.DataFrame, label_col):
    """
    Searches for the row representing net/total rental revenue, using a
    strict priority hierarchy AND an exclusion guardrail — deliberately
    more conservative than a simple "first substring match", because
    conflating rental-income-specific revenue with a broader property-wide
    total (or non-rental income) would silently corrupt the month-over-month
    delta calculation.

    Priority hierarchy (checked in order — ALL rows are searched for term #1
    before term #2 is ever considered, not "first row in file order"):
        1. "Net Rental Income"
        2. "Net Rent Income"
        3. "Total Rental Income"
        4. "Gross Potential Rent"
        5. "Gross Effective Rent"

    Exclusion guardrail — a row is NEVER eligible for matching, regardless
    of hierarchy term, if its label contains any of:
        "TOTAL INCOME", "TOTAL REVENUE", "OTHER INCOME",
        "EFFECTIVE GROSS INCOME (EGI)"
    (all comparisons case-insensitive, whitespace-stripped)

    Returns the matching row (pandas Series), or None if nothing in the
    hierarchy matches after exclusions are applied.
    """
    excluded_substrings = [
        "total income",
        "total revenue",
        "other income",
        "effective gross income (egi)",
    ]
    search_hierarchy = [
        "net rental income",
        "net rent income",
        "total rental income",
        "gross potential rent",
        "gross effective rent",
    ]

    # Guardrail applied FIRST, globally — a row containing an excluded
    # substring is never eligible, even if it also happens to contain one
    # of the hierarchy terms above.
    eligible_rows = []
    for _, row in df.iterrows():
        label = str(row.get(label_col, '')).strip().lower()
        if any(excluded in label for excluded in excluded_substrings):
            continue
        eligible_rows.append((label, row))

    # Priority search: exhaust term #1 across every eligible row before
    # term #2 is ever considered.
    for term in search_hierarchy:
        for label, row in eligible_rows:
            if term in label:
                return row

    return None


def compute_monthly_deltas(monthly_values: dict):
    """
    monthly_values: ordered dict of {month_column_name: numeric_value}.
    Returns a list of dicts for every month-over-month decline >= threshold:
        {"month": ..., "prior_value": ..., "value": ..., "pct_change": ...}
    Skips months with a missing/zero prior value (can't compute % change).
    """
    months = list(monthly_values.keys())
    dips = []
    for i in range(1, len(months)):
        prior_val = monthly_values[months[i - 1]]
        curr_val = monthly_values[months[i]]
        if not prior_val or prior_val == 0:
            continue
        pct_change = (curr_val - prior_val) / abs(prior_val)
        if pct_change <= -DROP_THRESHOLD_PCT:
            dips.append({
                "month": months[i],
                "prior_month": months[i - 1],
                "prior_value": round(prior_val, 2),
                "value": round(curr_val, 2),
                "pct_change": round(pct_change * 100, 1),
            })
    return dips


def detect_delinquency_column(rent_roll_df: pd.DataFrame):
    """
    Looks for a Rent Roll column representing an outstanding/past-due
    balance. Returns the column name, or None if the rent roll doesn't
    have one — this is a genuinely optional field, not every rent roll
    tracks it.
    """
    keywords = ['past due', 'delinquen', 'outstanding', 'balance due', 'ledger balance', 'current balance']
    for col in rent_roll_df.columns:
        col_lower = str(col).lower()
        if any(kw in col_lower for kw in keywords):
            return col
    return None


def flag_delinquent_units(rent_roll_df: pd.DataFrame, delinquency_col: str, unit_col: str, threshold: float = 1.0):
    """
    Returns a list of {"unit": ..., "balance": ...} for every row where the
    delinquency column shows a real positive balance above the threshold
    (default $1 — filters out $0.00/blank rows, not a meaningful cutoff
    otherwise).
    """
    flagged = []
    for _, row in rent_roll_df.iterrows():
        try:
            balance = float(row.get(delinquency_col, 0) or 0)
        except (ValueError, TypeError):
            continue
        if balance >= threshold:
            flagged.append({
                "unit": str(row.get(unit_col, 'Unknown')),
                "balance": round(balance, 2),
            })
    return flagged


def categorize_dip(dip: dict, is_recent: bool, delinquent_units: list):
    """
    Assigns a category to a single detected dip, per the module's
    hierarchical logic. `is_recent` = whether this dip's month falls within
    the T12's own last RECENT_WINDOW_MONTHS months (~60 days — see module
    docstring for why this is a monthly-granularity approximation, not an
    exact day-count cutoff).

    Returns (flag, detail):
      flag   — short, exact bracketed string for the visible Excel cell.
      detail — longer explanation for a cell comment (hover tooltip),
               matching the pattern already used for Module 1's
               reconciliation flags elsewhere in the workbook.
    """
    if not is_recent:
        flag = "[Insufficient Historical Data to Track Deep Gaps]"
        detail = (
            "This drop occurred outside the recent window near the end of "
            "the T12 period. The Rent Roll is a point-in-time snapshot, so "
            "it can't reliably explain what happened this far back — the "
            "unit/tenant situation may have already changed since."
        )
        return flag, detail

    if delinquent_units:
        total_delinquent = sum(u["balance"] for u in delinquent_units)
        flag = "[FLAG: Current Outstanding Delinquency]"
        detail = (
            f"{len(delinquent_units)} unit(s) on the Rent Roll show unpaid "
            f"balances totaling ${total_delinquent:,.2f} as of the current "
            f"snapshot — a likely contributor to this recent revenue drop."
        )
        return flag, detail

    # A/B: recent, but not explained by delinquency — genuine cross-
    # referencing of exact lease-end dates against this specific month isn't
    # reliably possible across varied real-world Rent Roll formats, so this
    # flags the pattern (recent + unexplained by non-payment) rather than
    # claiming a specific matched lease-end date.
    flag = "[Recent Lease Rollover Transition Window]"
    detail = (
        "This drop occurred within the recent window, and no outstanding "
        "delinquency was found on the Rent Roll to explain it — consistent "
        "with a vacancy/turnover or concession burn-off, though the exact "
        "unit and cause were not independently confirmed."
    )
    return flag, detail


def run_variance_analysis(t12_df: pd.DataFrame, rent_roll_df: pd.DataFrame, unit_col: str):
    """
    Main entry point. Returns a dict:
        {"skipped": True, "reason": "..."}  — if the T12 lacks monthly data
        {"skipped": False, "dips": [...], "delinquent_units": [...]}
    Never raises for "no data found" cases — that's a valid, expected
    outcome (e.g. a genuinely stable property with no dips at all).
    """
    monthly_cols = detect_monthly_t12_columns(t12_df)
    if not monthly_cols:
        return {
            "skipped": True,
            "reason": "[Reconciliation Skipped: Source Document Lacks Monthly Granularity]",
        }

    # Find the label column (first non-numeric column, typically "Account Name")
    label_col = t12_df.columns[0]

    revenue_row = find_revenue_row(t12_df, label_col)
    if revenue_row is None:
        return {
            "skipped": True,
            "reason": "[Reconciliation Skipped: Source Document Lacks Standard Rental Income Labeling]",
        }

    monthly_values = {}
    for col in monthly_cols:
        try:
            val = float(revenue_row[col])
        except (ValueError, TypeError):
            val = 0.0
        monthly_values[col] = val

    dips = compute_monthly_deltas(monthly_values)

    delinquency_col = detect_delinquency_column(rent_roll_df)
    delinquent_units = (
        flag_delinquent_units(rent_roll_df, delinquency_col, unit_col)
        if delinquency_col else []
    )

    recent_months = set(monthly_cols[-RECENT_WINDOW_MONTHS:])

    results = []
    for dip in dips:
        is_recent = dip["month"] in recent_months
        flag, detail = categorize_dip(dip, is_recent, delinquent_units if is_recent else [])
        results.append({**dip, "flag": flag, "detail": detail})

    return {
        "skipped": False,
        "dips": results,
        "delinquent_units": delinquent_units,
        "delinquency_column_found": delinquency_col is not None,
    }


# ==========================================
# Generalized line-item anomaly flagging — independent of everything above.
# run_variance_analysis()/Module 7 only ever looks at ONE row (revenue),
# only ever checks for a DROP, only ever compares against the single PRIOR
# month, and cross-references the Rent Roll to explain WHY. This instead
# flags ANY row in the raw monthly T12 data, in EITHER direction (spike or
# drop), against its own TRAILING multi-month average, and never attempts
# to explain why — it's a QA flag, not a forensic finding.
#
# Deliberately uses the source document's own row labels verbatim, NOT this
# platform's normalized 14-category T12_LINE_ITEMS list (see
# excel_generator.py) — that list is for ANNUAL totals only (the AI
# extraction schema doesn't produce a monthly breakdown mapped to those
# categories; monthly data only exists here, read directly off the raw
# source document, the same way Module 7 above already does). Mapping
# arbitrary monthly rows onto a fixed category list reliably needs real
# categorization logic, which this intentionally does not attempt.
# ==========================================
LINE_ITEM_ANOMALY_WINDOW_MONTHS = 5   # trailing months averaged for comparison
LINE_ITEM_ANOMALY_THRESHOLD_PCT = 0.30  # +/- deviation from that average that triggers a flag


def detect_line_item_anomalies(t12_df: pd.DataFrame, monthly_cols: list, label_col):
    """
    For every row (line item) in t12_df, and every month from the
    (window+1)th onward, compares that month's value against the TRAILING
    LINE_ITEM_ANOMALY_WINDOW_MONTHS-month average of the months immediately
    before it (never including the month itself, never "centered" — this
    only ever looks backward, same convention as compute_monthly_deltas
    above). Flags it when the deviation exceeds
    LINE_ITEM_ANOMALY_THRESHOLD_PCT in EITHER direction.

    Returns a list of row-level dicts, one per line item with at least one
    flagged month (a row with zero flagged months is omitted entirely —
    nothing downstream needs a "this row was checked and is fine" entry):
        {
            "label": "...",
            "values": {month_col: value_or_None, ...},   # the FULL monthly
                # series for this row, in monthly_cols order — kept even for
                # non-flagged months so the output can show real surrounding
                # context next to the flagged cell, not just that cell alone.
            "anomalies": {
                month_col: {
                    "direction": "spike" | "drop",
                    "pct_change": ...,      # signed, already *100 (percent)
                    "trailing_avg": ...,
                    "flag": "...",          # short label for the cell
                    "detail": "...",        # longer text for a cell Comment
                },
                ...
            },
        }

    Skips a (row, month) comparison entirely (no flag, no error) rather than
    guessing when: there isn't a full trailing window of real numbers behind
    it (missing/non-numeric source cell), or the trailing average is exactly
    zero (can't compute a meaningful % deviation off nothing) — same "never
    fabricate" discipline as the rest of this module.
    """
    rows = []
    if len(monthly_cols) <= LINE_ITEM_ANOMALY_WINDOW_MONTHS:
        # Not enough trailing history for even one comparison, anywhere.
        return rows

    for _, row in t12_df.iterrows():
        label = str(row.get(label_col, '')).strip()
        if not label:
            continue

        values = []
        for col in monthly_cols:
            try:
                values.append(float(row[col]))
            except (ValueError, TypeError):
                values.append(None)

        row_anomalies = {}
        for i in range(LINE_ITEM_ANOMALY_WINDOW_MONTHS, len(monthly_cols)):
            window_vals = values[i - LINE_ITEM_ANOMALY_WINDOW_MONTHS:i]
            curr_val = values[i]
            if curr_val is None or any(v is None for v in window_vals):
                continue

            trailing_avg = sum(window_vals) / len(window_vals)
            if trailing_avg == 0:
                continue

            pct_change = (curr_val - trailing_avg) / abs(trailing_avg)
            if abs(pct_change) < LINE_ITEM_ANOMALY_THRESHOLD_PCT:
                continue

            direction = "spike" if pct_change > 0 else "drop"
            month = monthly_cols[i]
            row_anomalies[month] = {
                "direction": direction,
                "pct_change": round(pct_change * 100, 1),
                "trailing_avg": round(trailing_avg, 2),
                "flag": f"{direction.capitalize()} vs. {LINE_ITEM_ANOMALY_WINDOW_MONTHS}-Mo. Avg",
                "detail": (
                    f"System Flag: {label} {'spiked' if direction == 'spike' else 'dropped'} "
                    f"{abs(pct_change) * 100:.1f}% in {month} compared to its trailing "
                    f"{LINE_ITEM_ANOMALY_WINDOW_MONTHS}-month average (${trailing_avg:,.2f})."
                ),
            }

        if row_anomalies:
            rows.append({
                "label": label,
                "values": {col: (round(v, 2) if v is not None else None) for col, v in zip(monthly_cols, values)},
                "anomalies": row_anomalies,
            })

    return rows


def run_line_item_anomaly_scan(t12_df: pd.DataFrame):
    """
    Entry point mirroring run_variance_analysis()'s shape/conventions, for
    the generalized (any-line-item) scan that feeds the output workbook's
    separate 'Monthly Variance' tab (see excel_generator.py). Fully
    independent of run_variance_analysis()/Module 7 above — never crashes
    the caller; "nothing flagged" and "couldn't run" are both valid,
    expected outcomes, same as that function.

    Returns:
        {"skipped": True, "reason": "..."} — if the T12 lacks monthly data
        {"skipped": False, "monthly_cols": [...], "rows": [...]} — see
        detect_line_item_anomalies() above for the shape of "rows".
    """
    monthly_cols = detect_monthly_t12_columns(t12_df)
    if not monthly_cols:
        return {
            "skipped": True,
            "reason": "Monthly Variance Skipped: Source Document Lacks Monthly Granularity.",
        }

    label_col = t12_df.columns[0]
    rows = detect_line_item_anomalies(t12_df, monthly_cols, label_col)

    return {
        "skipped": False,
        "monthly_cols": [str(c) for c in monthly_cols],
        "rows": rows,
    }