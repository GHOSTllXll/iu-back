# backend/ai_service/waterfall_calculator.py
"""
V2 Module 3 (Live Debt & Equity Waterfall Scenario Builder): pure Python,
zero-dependency (beyond numpy, already a project dependency) financial
math for one "what-if" financing/equity scenario on a deal. No AI calls,
no network calls, no new paid service — this is the ONE calculation
engine both the Scenario Builder's "Calculate" action and the Approve &
Lock Baseline / IC Summary PDF paths call, so the number an analyst sees
on screen and the number that ends up in front of the Investment
Committee can never silently disagree (see Phase V2-1's build-order
confirmation: "Backend Python is authoritative" was the explicit choice
over a client-side JS mirror, specifically to avoid two implementations
of promote-tier math drifting apart).

======================================================================
MODELING CONVENTION (read this before touching the waterfall math)
======================================================================
This is a SCENARIO/UNDERWRITING tool, not fund-administration software.
It deliberately uses the standard simplification real underwriting
platforms use at this stage of a deal: an ANNUAL cash flow model (not
monthly/daily) with a "BACK-ENDED" or "promote-at-exit" waterfall — every
operating year's cash flow is split between LP and GP strictly pro-rata
by their equity split (no promote during the hold), and the ENTIRE
promote calculation happens as a single true-up against the final
(sale) year's distribution. This is NOT the same thing as a full annual
"American" waterfall with interim IRR-hurdle tracking and GP catch-up
provisions paid out year-by-year — that is materially more complex
fund-accounting-grade machinery. The back-ended convention is standard
for acquisition-stage scenario modeling (what a deal's return profile
plausibly looks like) and is clearly surfaced as such in the UI/PDF
label ("Promote applied at exit"), not silently passed off as a full
waterfall engine.

ALGORITHM, step by step:
  1. Build the deal's total (LP+GP) equity cash flow stream, year 0
     through the hold period, from NOI, debt service, and net sale
     proceeds (see _build_property_cashflows).
  2. Split every year PRO RATA by equity_split_pct into an LP stream and
     a GP stream — this is "no promote yet."
  3. Solve, independently for each promote tier's hurdle_irr_pct, what
     the FINAL YEAR's distribution to LP would have had to be (holding
     every other year fixed at its pro-rata amount) to make LP's whole
     cash flow stream's IRR land exactly on that hurdle — a one-unknown
     root-find (_solve_single_cashflow_for_irr).
  4. LP's actual pro-rata final-year distribution is then sliced into
     the bands between consecutive tier hurdles; GP's promote_pct of
     each slice above the first tier is reallocated from LP to GP. The
     untouched portion below the first tier's hurdle is pure pro-rata
     (LP's preferred-return band).
  5. LP's real (post-promote) cash flow stream = pro-rata operating
     years + the tiered final year. IRR/Equity Multiple/Cash-on-Cash
     are computed from that real stream.

A cash-out event at a bridge-to-perm refinance (new loan > old balance)
is treated as an interim distribution to equity in that year, pro-rata
only (never subject to promote — promote is exit-only in this
convention, documented above); a cash-IN event (new loan < old balance,
i.e. the refinance couldn't fully retire the bridge loan) is treated as
an additional pro-rata equity call in that year.
"""
import math


# ----------------------------------------------------------------------
# Low-level financial primitives
# ----------------------------------------------------------------------
def _npv(rate, cashflows):
    """Standard NPV at `rate` (e.g. 0.08 for 8%) of a cash flow list
    where index 0 is "today" (no discounting applied to index 0)."""
    return sum(cf / ((1.0 + rate) ** i) for i, cf in enumerate(cashflows))


def solve_irr(cashflows, low=-0.9999, high=10.0, tol=1e-7, max_iter=200):
    """
    Solves for the internal rate of return of `cashflows` (index 0 =
    year 0, typically negative; later entries typically positive) via
    bisection on NPV(rate) == 0.

    Bisection (not Newton's method) deliberately: this project's cash
    flow shape — one negative outflow at year 0, all-positive
    thereafter — has NPV(rate) strictly monotonically decreasing in
    rate (by Descartes' Rule of Signs, exactly one sign change in the
    cash flow sequence means exactly one positive real root), so
    bisection on a bracket that spans it is guaranteed to converge, with
    no risk of a bad starting guess diverging the way Newton's method
    can. Returns None (not an exception, not a silently wrong number) if
    no root exists in [low, high] — e.g. total distributions never
    exceed the equity invested (a total-loss-or-worse scenario, whose
    "IRR" isn't a real number in the usual sense: nothing in [low, high]
    recovers capital).
    """
    if not cashflows or all(cf == 0 for cf in cashflows):
        return None

    npv_low = _npv(low, cashflows)
    npv_high = _npv(high, cashflows)
    if npv_low == 0:
        return low
    if npv_high == 0:
        return high
    if (npv_low > 0) == (npv_high > 0):
        # No sign change across the bracket — most commonly a deal that
        # never returns its equity even at a -99.99% discount rate (a
        # total loss), or an unrealistically high positive IRR beyond
        # 1000% (practically meaningless either way). Never guessed.
        return None

    for _ in range(max_iter):
        mid = (low + high) / 2.0
        npv_mid = _npv(mid, cashflows)
        if abs(npv_mid) < tol:
            return mid
        if (npv_mid > 0) == (npv_low > 0):
            low, npv_low = mid, npv_mid
        else:
            high = mid
    return (low + high) / 2.0


def _solve_single_cashflow_for_irr(cashflows, target_index, target_irr):
    """
    Holding every entry of `cashflows` EXCEPT index `target_index` fixed,
    solves for the value at `target_index` that makes the whole stream's
    IRR equal `target_irr`. Closed-form (not iterative) — the stream's
    NPV at `target_irr` is linear in the unknown entry, so this is a
    single algebraic rearrangement of NPV(target_irr) == 0, not a second
    root-find nested inside the IRR solve above:

        0 = sum(cf_i / (1+r)^i for i != target_index) + x / (1+r)^target_index
        x = -sum(cf_i / (1+r)^i for i != target_index) * (1+r)^target_index
    """
    r = target_irr
    other_sum_pv = sum(
        cf / ((1.0 + r) ** i) for i, cf in enumerate(cashflows) if i != target_index
    )
    return -other_sum_pv * ((1.0 + r) ** target_index)


# ----------------------------------------------------------------------
# Debt sizing / amortization
# ----------------------------------------------------------------------
def _level_payment(principal, annual_rate_pct, amortization_years):
    """Standard level-payment (fully amortizing) annual payment amount."""
    if amortization_years <= 0:
        return principal
    r = annual_rate_pct / 100.0
    if r == 0:
        return principal / amortization_years
    return principal * (r * (1 + r) ** amortization_years) / ((1 + r) ** amortization_years - 1)


def _amortization_schedule(principal, annual_rate_pct, amortization_years, num_years):
    """
    Returns a list of `num_years` dicts {interest, principal, debt_service,
    ending_balance} for a level-payment amortizing loan, truncated to
    num_years (num_years may be less than amortization_years — a 10-year
    hold on a 30-year am schedule, for instance). If num_years exceeds
    amortization_years, the loan is simply paid off (zero debt service)
    for the remaining years.
    """
    r = annual_rate_pct / 100.0
    payment = _level_payment(principal, annual_rate_pct, amortization_years)
    balance = principal
    schedule = []
    for year in range(1, num_years + 1):
        if balance <= 0.01 or year > amortization_years:
            schedule.append({'interest': 0.0, 'principal': 0.0, 'debt_service': 0.0, 'ending_balance': max(balance, 0.0)})
            continue
        interest = balance * r
        principal_paid = min(payment - interest, balance)
        balance = max(balance - principal_paid, 0.0)
        schedule.append({
            'interest': round(interest, 2), 'principal': round(principal_paid, 2),
            'debt_service': round(interest + principal_paid, 2), 'ending_balance': round(balance, 2),
        })
    return schedule


# ----------------------------------------------------------------------
# Property-level (unlevered-then-levered) cash flows
# ----------------------------------------------------------------------
class WaterfallInputError(Exception):
    """Raised for a scenario input that can't be modeled at all (not a
    finance judgment call — a genuinely missing/invalid required field).
    Caught at the view layer and returned as a 400, never silently
    defaulted."""
    pass


def _project_noi(going_in_noi, noi_growth_pct, num_years):
    """Year 1..num_years NOI, straight-line compounding growth off the
    deal's own going-in (year 1) NOI. Year num_years+1 (the "forward"
    NOI used to cap the exit sale price — see _build_property_cashflows)
    is also returned as the list's (num_years+1)th implicit value by the
    caller re-calling this with num_years+1 and taking the last entry."""
    g = noi_growth_pct / 100.0
    return [going_in_noi * ((1 + g) ** (year - 1)) for year in range(1, num_years + 1)]


def _build_property_cashflows(inputs, going_in_noi):
    """
    Returns (total_equity, annual_cashflows, sale_price, detail) where
    annual_cashflows[0] is year-0 (negative, the initial equity outflow)
    and annual_cashflows[1..hold_period] are the TOTAL (LP+GP, pre-split)
    equity cash flows for each operating/exit year.
    """
    purchase_price = inputs['purchase_price']
    hold_years = inputs['hold_period_years']
    noi_growth_pct = inputs.get('noi_growth_pct', 3.0)
    exit_cap_rate_pct = inputs['exit_cap_rate_pct']

    noi_by_year = _project_noi(going_in_noi, noi_growth_pct, hold_years + 1)  # +1 for forward NOI at exit

    debt_structure = inputs['debt_structure']  # 'fixed_rate_permanent' | 'bridge_to_perm'

    if debt_structure == 'fixed_rate_permanent':
        initial_loan = inputs['acquisition_ltv_pct'] / 100.0 * purchase_price
        schedule = _amortization_schedule(
            initial_loan, inputs['interest_rate_pct'], inputs['amortization_years'], hold_years
        )
        debt_service_by_year = [row['debt_service'] for row in schedule]
        ending_balance_at_exit = schedule[-1]['ending_balance'] if schedule else initial_loan
        interim_refi_cashflow_by_year = [0.0] * hold_years  # no refi event in this structure

    elif debt_structure == 'bridge_to_perm':
        initial_loan = inputs['acquisition_ltv_pct'] / 100.0 * purchase_price
        bridge_rate_pct = inputs['bridge_rate_pct']
        io_years = int(inputs['io_period_years'])
        refinance_ltv_pct = inputs['refinance_ltv_pct']
        refinance_dscr = inputs['refinance_dscr']

        debt_service_by_year = []
        interim_refi_cashflow_by_year = [0.0] * hold_years
        balance = initial_loan
        for year in range(1, hold_years + 1):
            if year <= io_years:
                debt_service_by_year.append(round(balance * (bridge_rate_pct / 100.0), 2))
            elif year == io_years + 1:
                # Refinance event: size the new permanent loan as the
                # LESSER of an LTV-based and a DSCR-based amount (the two
                # named constraints in the spec) — using this year's NOI
                # and the SAME bridge_rate_pct as the assumed permanent
                # rate (no separate permanent-rate/amortization input was
                # specified for this structure — see this module's own
                # docstring for why that's a deliberate, documented v1
                # simplification rather than an invented field) and
                # carrying the refinanced loan interest-only through the
                # rest of the hold, for the same reason.
                noi_at_refi = noi_by_year[year - 1]
                refi_value = noi_at_refi / (exit_cap_rate_pct / 100.0)
                ltv_capacity = refinance_ltv_pct / 100.0 * refi_value
                dscr_capacity = (noi_at_refi / refinance_dscr) / (bridge_rate_pct / 100.0) \
                    if refinance_dscr and bridge_rate_pct else 0.0
                new_loan = max(min(ltv_capacity, dscr_capacity), 0.0)
                interim_refi_cashflow_by_year[year - 1] = round(new_loan - balance, 2)
                balance = new_loan
                debt_service_by_year.append(round(balance * (bridge_rate_pct / 100.0), 2))
            else:
                debt_service_by_year.append(round(balance * (bridge_rate_pct / 100.0), 2))
        ending_balance_at_exit = balance
    else:
        raise WaterfallInputError(f"Unknown debt_structure '{debt_structure}'.")

    forward_noi = noi_by_year[hold_years]  # year hold_years+1, capped for the sale
    sale_price = forward_noi / (exit_cap_rate_pct / 100.0)
    net_sale_proceeds = sale_price - ending_balance_at_exit

    total_equity = purchase_price - initial_loan
    annual_cashflows = [-total_equity]
    for year in range(1, hold_years + 1):
        noi = noi_by_year[year - 1]
        cf = noi - debt_service_by_year[year - 1] + interim_refi_cashflow_by_year[year - 1]
        if year == hold_years:
            cf += net_sale_proceeds
        annual_cashflows.append(round(cf, 2))

    detail = {
        'purchase_price': purchase_price, 'initial_loan_amount': round(initial_loan, 2),
        'total_equity': round(total_equity, 2), 'noi_by_year': [round(n, 2) for n in noi_by_year[:hold_years]],
        'debt_service_by_year': debt_service_by_year, 'sale_price': round(sale_price, 2),
        'ending_loan_balance_at_exit': round(ending_balance_at_exit, 2),
        'net_sale_proceeds': round(net_sale_proceeds, 2),
    }
    return total_equity, annual_cashflows, sale_price, detail


# ----------------------------------------------------------------------
# LP/GP promote-tiered allocation
# ----------------------------------------------------------------------
def _allocate_promote(total_cashflows, equity_split_pct, promote_tiers):
    """
    total_cashflows: [year0 (negative), year1, ..., yearN] TOTAL (pre-
    split) equity cash flows. equity_split_pct: LP's share (e.g. 80.0
    for 80/20). promote_tiers: sorted list of
    {"hurdle_irr_pct": float or None, "gp_promote_pct": float}; a tier
    with hurdle_irr_pct=None is the top (uncapped) tier and must be last.

    Returns (lp_cashflows, gp_cashflows, tier_breakdown).
    """
    lp_share = equity_split_pct / 100.0
    pro_rata_lp = [cf * lp_share for cf in total_cashflows]
    pro_rata_gp = [cf * (1 - lp_share) for cf in total_cashflows]

    final_idx = len(total_cashflows) - 1
    lp_final_pro_rata = pro_rata_lp[final_idx]

    sorted_tiers = sorted(promote_tiers, key=lambda t: (t['hurdle_irr_pct'] is None, t.get('hurdle_irr_pct') or 0))

    # Each tier's "breakpoint" is the LP final-year distribution level
    # that exactly hits that tier's hurdle IRR, holding all other LP
    # years fixed at their pro-rata amount.
    breakpoints = []
    prior_breakpoint = 0.0
    for tier in sorted_tiers:
        hurdle = tier.get('hurdle_irr_pct')
        if hurdle is None:
            breakpoint = lp_final_pro_rata  # top tier's ceiling is LP's own actual pro-rata amount
        else:
            stream = list(pro_rata_lp)
            stream[final_idx] = 0.0  # solved value replaces this below
            breakpoint = _solve_single_cashflow_for_irr(
                [cf if i != final_idx else 0.0 for i, cf in enumerate(pro_rata_lp)],
                final_idx, hurdle / 100.0,
            )
            breakpoint = max(breakpoint, prior_breakpoint)
        breakpoints.append(breakpoint)
        prior_breakpoint = breakpoint

    lp_final_actual = 0.0
    gp_promote_total = 0.0
    tier_breakdown = []
    lower = 0.0
    for tier, breakpoint in zip(sorted_tiers, breakpoints):
        upper = min(breakpoint, lp_final_pro_rata) if tier.get('hurdle_irr_pct') is not None else lp_final_pro_rata
        upper = max(upper, lower)
        slice_amount = max(upper - lower, 0.0)
        gp_cut = slice_amount * (tier['gp_promote_pct'] / 100.0)
        lp_keeps = slice_amount - gp_cut
        lp_final_actual += lp_keeps
        gp_promote_total += gp_cut
        tier_breakdown.append({
            'hurdle_irr_pct': tier.get('hurdle_irr_pct'), 'gp_promote_pct': tier['gp_promote_pct'],
            'slice_amount': round(slice_amount, 2), 'gp_promote_amount': round(gp_cut, 2),
        })
        lower = upper

    lp_cashflows = list(pro_rata_lp)
    lp_cashflows[final_idx] = round(lp_final_actual, 2)
    gp_cashflows = list(pro_rata_gp)
    gp_cashflows[final_idx] = round(pro_rata_gp[final_idx] + gp_promote_total, 2)

    return lp_cashflows, gp_cashflows, tier_breakdown


# ----------------------------------------------------------------------
# Public entry point
# ----------------------------------------------------------------------
REQUIRED_FIELDS_COMMON = [
    'purchase_price', 'hold_period_years', 'exit_cap_rate_pct', 'acquisition_ltv_pct',
    'equity_split_pct', 'promote_tiers',
]
REQUIRED_FIELDS_FIXED = ['interest_rate_pct', 'amortization_years']
REQUIRED_FIELDS_BRIDGE = ['bridge_rate_pct', 'io_period_years', 'refinance_ltv_pct', 'refinance_dscr']


def compute_waterfall(inputs: dict, going_in_noi: float) -> dict:
    """
    The single authoritative entry point — called from GridWaterfallView
    (live "Calculate"), from the baseline-lock snapshot capture, and from
    IC Summary PDF rendering. See this module's own docstring for the
    full algorithm and its documented conventions.

    Raises WaterfallInputError for a missing/invalid required field —
    the caller is responsible for turning that into a clean 400, never
    silently defaulting a financial assumption.
    """
    if going_in_noi is None or going_in_noi <= 0:
        raise WaterfallInputError(
            "This deal has no positive Net Operating Income to build a scenario from yet."
        )

    missing = [f for f in REQUIRED_FIELDS_COMMON if inputs.get(f) in (None, '')]
    if missing:
        raise WaterfallInputError(f"Missing required field(s): {', '.join(missing)}.")

    debt_structure = inputs.get('debt_structure')
    if debt_structure == 'fixed_rate_permanent':
        missing = [f for f in REQUIRED_FIELDS_FIXED if inputs.get(f) in (None, '')]
    elif debt_structure == 'bridge_to_perm':
        missing = [f for f in REQUIRED_FIELDS_BRIDGE if inputs.get(f) in (None, '')]
    else:
        raise WaterfallInputError(
            "debt_structure must be 'fixed_rate_permanent' or 'bridge_to_perm'."
        )
    if missing:
        raise WaterfallInputError(f"Missing required field(s) for {debt_structure}: {', '.join(missing)}.")

    hold_years = int(inputs['hold_period_years'])
    if hold_years < 1 or hold_years > 40:
        raise WaterfallInputError("hold_period_years must be between 1 and 40.")

    promote_tiers = inputs['promote_tiers']
    if not isinstance(promote_tiers, list) or not promote_tiers:
        raise WaterfallInputError("promote_tiers must be a non-empty list.")
    if promote_tiers[-1].get('hurdle_irr_pct') is not None:
        raise WaterfallInputError("The last promote tier must be uncapped (hurdle_irr_pct: null).")

    total_equity, total_cashflows, sale_price, detail = _build_property_cashflows(inputs, going_in_noi)

    lp_cashflows, gp_cashflows, tier_breakdown = _allocate_promote(
        total_cashflows, inputs['equity_split_pct'], promote_tiers
    )

    lp_irr = solve_irr(lp_cashflows)
    lp_equity_invested = total_equity * (inputs['equity_split_pct'] / 100.0)
    lp_total_distributions = sum(cf for cf in lp_cashflows[1:])
    lp_equity_multiple = (lp_total_distributions / lp_equity_invested) if lp_equity_invested else None

    operating_lp_cashflows = lp_cashflows[1:-1] if hold_years > 1 else []
    lp_cash_on_cash_year1_pct = (
        (lp_cashflows[1] / lp_equity_invested * 100.0) if lp_equity_invested and len(lp_cashflows) > 1 else None
    )
    operating_years_for_avg = lp_cashflows[1:hold_years]  # excludes the sale-year entry's windfall by design
    lp_cash_on_cash_avg_pct = (
        (sum(operating_years_for_avg) / len(operating_years_for_avg) / lp_equity_invested * 100.0)
        if lp_equity_invested and operating_years_for_avg else lp_cash_on_cash_year1_pct
    )

    return {
        'lp_irr_pct': round(lp_irr * 100.0, 2) if lp_irr is not None else None,
        'lp_equity_multiple': round(lp_equity_multiple, 2) if lp_equity_multiple is not None else None,
        'lp_cash_on_cash_year1_pct': round(lp_cash_on_cash_year1_pct, 2) if lp_cash_on_cash_year1_pct is not None else None,
        'lp_cash_on_cash_avg_pct': round(lp_cash_on_cash_avg_pct, 2) if lp_cash_on_cash_avg_pct is not None else None,
        'lp_equity_invested': round(lp_equity_invested, 2),
        'gp_equity_invested': round(total_equity - lp_equity_invested, 2),
        'total_equity': round(total_equity, 2),
        'sale_price': round(sale_price, 2),
        'lp_cashflows': lp_cashflows,
        'gp_cashflows': gp_cashflows,
        'total_cashflows': total_cashflows,
        'tier_breakdown': tier_breakdown,
        'deal_detail': detail,
        'modeling_convention': (
            "Annual cash flows, promote applied as a single true-up at exit "
            "(back-ended waterfall) — see waterfall_calculator.py's own "
            "docstring."
        ),
    }
