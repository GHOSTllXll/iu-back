# backend/ai_service/market_comps.py
"""
V2 Module 2 (Tax & Rent Comps Panel): turns a cached
PropertyMarketDataSnapshot (the raw RentCast responses — expensive,
fetched at most once per deal, see rentcast_client.py) plus the deal's
CURRENT metrics (cheap, re-read live every time — so a later Excel
re-upload or grid edit is always compared against up-to-date seller-
stated figures, with no need to ever re-fetch RentCast) into the handful
of derived numbers and variance flags the "Market & Tax Comps" panel
displays.

Deliberately a SEPARATE flag namespace from reconciliation.py's Module 1
("Bullshit Detector") flags: those catch a seller's OWN documents
disagreeing with each other (OM vs Rent Roll vs T12); these catch a
seller's documents disagreeing with the EXTERNAL market (RentCast). Kept
apart so the V2 Scorecard's "Forensic Seller-Trap Flags Triggered" metric
(Module 4b) — which explicitly means the Module 1 flags per the client's
own definition — never silently double-counts these too. Same dict shape
as reconciliation.py's flags anyway ({metric, severity, message,
claimed_value, actual_value, variance_pct}), reusing its exact
WARNING/CRITICAL split, for a consistent look across both panels.
"""
WARNING_THRESHOLD_PCT = 2.0
CRITICAL_THRESHOLD_PCT = 5.0


def _severity_for_variance(variance_pct: float) -> str:
    return "CRITICAL" if variance_pct >= CRITICAL_THRESHOLD_PCT else "WARNING"


def latest_year_entry(year_keyed_dict: dict) -> dict | None:
    """
    RentCast's /properties response keys taxAssessments/propertyTaxes by
    4-digit year string, e.g. {"2022": {...}, "2023": {...}, "2024":
    {...}}. Returns the entry for the MOST RECENT year present (the most
    current assessment/tax bill), or None if the dict is empty/absent.
    Tolerant of keys already being ints, in case RentCast changes this.
    """
    if not year_keyed_dict:
        return None
    try:
        latest_key = max(year_keyed_dict.keys(), key=lambda k: int(k))
    except (ValueError, TypeError):
        return None
    return year_keyed_dict.get(latest_key)


def summarize_property_record(property_record: dict | None) -> dict:
    """
    Extracts just the handful of scalar fields the panel/variance checks
    need from RentCast's full /properties record — see
    PropertyMarketDataSnapshot's own fields in models.py for where these
    get persisted.
    """
    if not property_record:
        return {
            'market_assessed_value': None, 'market_assessed_value_year': None,
            'market_annual_property_tax': None, 'market_annual_property_tax_year': None,
        }

    tax_assessment = latest_year_entry(property_record.get('taxAssessments') or {})
    property_tax = latest_year_entry(property_record.get('propertyTaxes') or {})

    return {
        'market_assessed_value': (tax_assessment or {}).get('value'),
        'market_assessed_value_year': (tax_assessment or {}).get('year'),
        'market_annual_property_tax': (property_tax or {}).get('total'),
        'market_annual_property_tax_year': (property_tax or {}).get('year'),
    }


def summarize_rental_comps(rental_comps: list) -> dict:
    """
    Nearby active long-term rental listings -> a short display list (the
    fields the panel's comps table shows) plus a simple average asking
    rent across whatever came back, used as the market benchmark in
    check_rent_variance below. A straight mean, not weighted by
    recency/distance — RentCast already sorts by lastSeenDate
    descending and the panel shows each comp's own distance-relevant
    fields (city/state/beds/baths) so an analyst can eyeball comparability
    themselves; a fancier weighting scheme isn't something the client
    asked for.
    """
    comps = []
    rents = []
    for listing in (rental_comps or [])[:20]:
        if not isinstance(listing, dict):
            continue
        price = listing.get('price')
        comps.append({
            'formatted_address': listing.get('formattedAddress'),
            'city': listing.get('city'),
            'state': listing.get('state'),
            'property_type': listing.get('propertyType'),
            'bedrooms': listing.get('bedrooms'),
            'bathrooms': listing.get('bathrooms'),
            'square_footage': listing.get('squareFootage'),
            'monthly_rent': price,
            'days_on_market': listing.get('daysOnMarket'),
            'listed_date': listing.get('listedDate'),
        })
        if isinstance(price, (int, float)):
            rents.append(price)

    return {
        'comps': comps,
        'comp_count': len(comps),
        'market_avg_monthly_rent': round(sum(rents) / len(rents), 2) if rents else None,
    }


def summarize_avm_value(avm_value: dict | None) -> dict:
    if not avm_value:
        return {'market_value_estimate': None, 'market_value_range_low': None, 'market_value_range_high': None}
    return {
        'market_value_estimate': avm_value.get('price'),
        'market_value_range_low': avm_value.get('priceRangeLow'),
        'market_value_range_high': avm_value.get('priceRangeHigh'),
    }


def _pct_variance_flag(metric: str, claimed_value, actual_value, message_fn) -> dict | None:
    if not isinstance(claimed_value, (int, float)) or not isinstance(actual_value, (int, float)) or not actual_value:
        return None
    variance_pct = abs(claimed_value - actual_value) / actual_value * 100
    if variance_pct < WARNING_THRESHOLD_PCT:
        return None
    return {
        'metric': metric,
        'severity': _severity_for_variance(variance_pct),
        'message': message_fn(claimed_value, actual_value, variance_pct),
        'claimed_value': claimed_value,
        'actual_value': actual_value,
        'variance_pct': round(variance_pct, 2),
    }


def check_tax_variance(seller_stated_annual_tax, market_annual_tax, market_tax_year) -> dict | None:
    """
    Flags when the seller's T12-stated "Real Estate Taxes" line
    disagrees with RentCast's current assessed tax bill by >=2% — most
    commonly a sign of an impending tax REASSESSMENT at sale (many
    jurisdictions reassess at the new purchase price, so a seller's
    stale, lower tax bill understates what a buyer will actually pay
    going forward), the single most common "seller trap" in CRE tax
    lines. Symmetric either direction (same abs-variance convention as
    reconciliation.py) — a seller overstating taxes is unusual but still
    worth a flag.
    """
    def message(claimed, actual, variance_pct):
        direction = (
            "RentCast's current assessed tax bill is HIGHER than the seller's stated figure "
            "— check for a pending reassessment at sale price that would increase the buyer's "
            "actual tax burden going forward."
            if actual > claimed else
            "RentCast's current assessed tax bill is LOWER than the seller's stated figure."
        )
        year_note = f" ({market_tax_year})" if market_tax_year else ""
        return (
            f"T12 states ${claimed:,.0f} in Real Estate Taxes, but RentCast's current assessed "
            f"tax bill{year_note} is ${actual:,.0f} — a variance of {variance_pct:.1f}%. {direction}"
        )

    return _pct_variance_flag(
        'real_estate_taxes', seller_stated_annual_tax, market_annual_tax, message
    )


def check_rent_variance(seller_stated_market_rent, market_avg_monthly_rent, comp_count) -> dict | None:
    """
    Flags when the OM's claimed "market" achievable rent
    (rent_roll_metrics.avg_market_monthly_rent — the upside the seller is
    marketing, not the actual in-place rent) disagrees with the simple
    average asking rent across nearby active comps by >=2%. Requires at
    least 3 comps before flagging at all — RentCast returning only 1-2
    thin comps isn't a reliable enough benchmark to call a seller's
    number wrong.
    """
    if not comp_count or comp_count < 3:
        return None

    def message(claimed, actual, variance_pct):
        direction = (
            "the OM's claimed market rent is HIGHER than nearby comps support — a common "
            "overstated-upside seller trap."
            if claimed > actual else
            "the OM's claimed market rent is LOWER than nearby comps — the seller may be "
            "understating achievable rent."
        )
        return (
            f"OM claims ${claimed:,.0f}/mo achievable market rent, but the average asking rent "
            f"across {comp_count} nearby active comps is ${actual:,.0f}/mo — a variance of "
            f"{variance_pct:.1f}%. {direction}"
        )

    return _pct_variance_flag(
        'avg_market_monthly_rent', seller_stated_market_rent, market_avg_monthly_rent, message
    )


def check_value_variance(asking_price, market_value_estimate) -> dict | None:
    """
    Flags when the deal's asking price disagrees with RentCast's AVM
    value estimate by >=2% — purely informational context for how the
    asking price compares to an independent, automated valuation; AVMs
    are directional estimates, not appraisals, so this is deliberately
    framed as a comparison rather than an accusation either direction.
    """
    def message(claimed, actual, variance_pct):
        direction = "above" if claimed > actual else "below"
        return (
            f"Asking price of ${claimed:,.0f} is {variance_pct:.1f}% {direction} RentCast's "
            f"automated value estimate of ${actual:,.0f}. AVMs are directional estimates, not "
            f"appraisals — treat as context, not a verdict."
        )

    return _pct_variance_flag('asking_price_vs_avm', asking_price, market_value_estimate, message)


def build_market_comps_payload(snapshot, metrics: dict, asking_price) -> dict:
    """
    Assembles the full "Market & Tax Comps" panel payload for one deal
    from its cached snapshot (raw RentCast data — see
    PropertyMarketDataSnapshot in models.py) and the project's CURRENT
    metrics/asking_price. Called fresh on every GET — cheap, pure
    Python, no network call — so the panel always reflects the latest
    seller-stated figures even though the RentCast data itself is frozen
    at whatever was cached on first check.
    """
    property_summary = summarize_property_record(snapshot.raw_property_record)
    rent_summary = summarize_rental_comps(snapshot.raw_rental_comps or [])
    avm_summary = summarize_avm_value(snapshot.raw_avm_value)

    t12_revenue_expenses = metrics.get('t12_revenue_expenses') or {}
    rent_roll_metrics = metrics.get('rent_roll_metrics') or {}

    flags = []
    for flag in (
        check_tax_variance(
            t12_revenue_expenses.get('real_estate_taxes'),
            property_summary['market_annual_property_tax'],
            property_summary['market_annual_property_tax_year'],
        ),
        check_rent_variance(
            rent_roll_metrics.get('avg_market_monthly_rent'),
            rent_summary['market_avg_monthly_rent'],
            rent_summary['comp_count'],
        ),
        check_value_variance(asking_price, avm_summary['market_value_estimate']),
    ):
        if flag is not None:
            flags.append(flag)

    return {
        'address_checked': snapshot.address_checked,
        'fetched_at': snapshot.fetched_at.isoformat() if snapshot.fetched_at else None,
        'fetched_by': snapshot.fetched_by.email if snapshot.fetched_by_id else None,
        'fetch_error': snapshot.fetch_error or None,
        'tax': property_summary,
        'rent_comps': rent_summary,
        'value_estimate': avm_summary,
        'variance_flags': flags,
        # Seller-stated figures echoed back alongside the market data so
        # the frontend never has to separately re-fetch metrics just to
        # render a side-by-side comparison row.
        'seller_stated': {
            'real_estate_taxes': t12_revenue_expenses.get('real_estate_taxes'),
            'avg_market_monthly_rent': rent_roll_metrics.get('avg_market_monthly_rent'),
            'avg_in_place_monthly_rent': rent_roll_metrics.get('avg_in_place_monthly_rent'),
            'asking_price': asking_price,
        },
    }
