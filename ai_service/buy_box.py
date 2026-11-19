# backend/ai_service/buy_box.py
"""
V2 Module 4a (Buy-Box Auto-Alerts): checks a freshly-created deal's
metrics against its workspace's configured buy-box criteria (see
TeamWorkspace.buy_box_criteria / get_buy_box_criteria() in
users/models.py) and returns whether/why it's a match. Mirrors
reconciliation.py's core design principle: every criterion here is a
plain Python comparison against a number the AI already extracted into
metrics — never a second AI call grading "fit".

Called exactly once, at the moment a BRAND NEW PropertyProject is
created (see document_history._check_buy_box_match_and_notify, invoked
from record_analysis_report) — a re-analysis of an EXISTING deal never
re-triggers this, so a deal already in the pipeline doesn't repeatedly
re-alert the team every time it's reprocessed.
"""


def _asset_class_matches(criteria_classes: list, deal_asset_class) -> bool:
    if not criteria_classes:
        return True  # no filter configured -> any class passes
    if not deal_asset_class:
        return False  # filter configured, but this deal has no extracted asset class
    return str(deal_asset_class).strip().lower() in {str(c).strip().lower() for c in criteria_classes}


def _location_matches(criteria_cities: list, criteria_states: list, address: str) -> bool:
    if not criteria_cities and not criteria_states:
        return True  # no location filter configured
    if not address:
        return False
    address_lower = address.lower()
    if criteria_cities and any(city.strip().lower() in address_lower for city in criteria_cities if city):
        return True
    if criteria_states and any(
        f", {state.strip().lower()}" in address_lower for state in criteria_states if state
    ):
        return True
    return False


def check_buy_box_match(criteria: dict, metrics: dict, asking_price) -> dict | None:
    """
    criteria is always the MERGED dict from TeamWorkspace.
    get_buy_box_criteria() — every key below is always present (None/[]
    when unset), so this never needs a criteria.get(key, default).

    Returns None when the workspace has buy-box alerts disabled, OR the
    deal fails ANY configured (non-None/non-empty) criterion. Returns a
    dict describing which criteria the deal satisfied (for the
    notification's message text) when it passes every configured check
    — a workspace with alerts enabled but NO filters set at all matches
    every new deal, by design ("alert me on everything").
    """
    if not criteria.get('enabled'):
        return None

    property_metadata = metrics.get('property_metadata') or {}
    deal_valuation = metrics.get('deal_valuation') or {}

    asset_class = property_metadata.get('asset_class')
    total_units = property_metadata.get('total_unit_count')
    address = property_metadata.get('address') or ''

    entry_cap_rate_pct = None
    raw_cap_rate = deal_valuation.get('entry_cap_rate')
    if isinstance(raw_cap_rate, (int, float)):
        entry_cap_rate_pct = raw_cap_rate * 100

    matched_criteria = []

    if not _asset_class_matches(criteria.get('asset_classes'), asset_class):
        return None
    if criteria.get('asset_classes'):
        matched_criteria.append(f"Asset class: {asset_class}")

    min_units = criteria.get('min_units')
    if min_units is not None:
        if not isinstance(total_units, (int, float)) or total_units < min_units:
            return None
        matched_criteria.append(f"{total_units:.0f} units (≥ {min_units:.0f})")

    max_units = criteria.get('max_units')
    if max_units is not None:
        if not isinstance(total_units, (int, float)) or total_units > max_units:
            return None

    min_cap_rate_pct = criteria.get('min_cap_rate_pct')
    if min_cap_rate_pct is not None:
        if entry_cap_rate_pct is None or entry_cap_rate_pct < min_cap_rate_pct:
            return None
        matched_criteria.append(f"{entry_cap_rate_pct:.2f}% entry cap rate (≥ {min_cap_rate_pct:.2f}%)")

    max_cap_rate_pct = criteria.get('max_cap_rate_pct')
    if max_cap_rate_pct is not None:
        if entry_cap_rate_pct is None or entry_cap_rate_pct > max_cap_rate_pct:
            return None

    if not _location_matches(criteria.get('target_cities'), criteria.get('target_states'), address):
        return None
    if criteria.get('target_cities') or criteria.get('target_states'):
        matched_criteria.append(f"Location: {address}")

    max_price_per_unit = criteria.get('max_price_per_unit')
    if max_price_per_unit is not None:
        price = asking_price or deal_valuation.get('target_purchase_price')
        if not isinstance(price, (int, float)) or not isinstance(total_units, (int, float)) or total_units <= 0:
            return None  # can't evaluate a configured criterion -> don't claim a match
        price_per_unit = price / total_units
        if price_per_unit > max_price_per_unit:
            return None
        matched_criteria.append(f"${price_per_unit:,.0f}/unit (≤ ${max_price_per_unit:,.0f}/unit)")

    return {'matched_criteria': matched_criteria}
