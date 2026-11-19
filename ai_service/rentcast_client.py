# backend/ai_service/rentcast_client.py
"""
V2 Module 2 (Tax & Rent Comps Panel): a thin, defensive wrapper around the
RentCast API (https://developers.rentcast.io) — the one and only paid
third-party service added in Version 2.0, deliberately kept on the free
50-request/month tier during development/testing per the client's explicit
cost decision. See PropertyMarketDataSnapshot in models.py for the
per-deal CACHE this module's results are written into exactly once ("the
first time an address is checked") — this module itself never caches
anything; it is the raw HTTP layer only, called from
views.MarketCompsView._fetch_and_cache().

Confirmed against RentCast's live API documentation (developers.rentcast.io)
before writing this — endpoints, auth header, and response shapes below are
NOT guessed:
  - Base URL: https://api.rentcast.io/v1
  - Auth: header "X-Api-Key: <key>" (NOT a query param, NOT Bearer auth)
  - Rate limit: 20 req/sec/key (irrelevant to us at this volume); the
    relevant cap is the ACCOUNT's 50 free requests/month, which is exactly
    why PropertyMarketDataSnapshot caches aggressively on the deal record.

Three calls per address check (3 of the account's 50 monthly requests
spent every time this module actually runs — which is why the caller only
ever runs it once per deal, on first open):
  1. GET /properties            -> tax assessment history + annual tax bill
  2. GET /avm/value              -> current market value estimate (AVM)
  3. GET /listings/rental/long-term -> nearby ACTIVE rental listings (real
     comps, not a single AVM number — see module docstring on why: this
     app covers multiple asset classes, including multifamily buildings
     where a single "per-unit" AVM rent estimate would require per-unit
     beds/baths/sqft this system doesn't track at the property level).

lookupSubjectAttributes is left at RentCast's own default (true) and
propertyType/bedrooms/bathrooms/squareFootage are deliberately NEVER
passed — this app's own asset_class field ("Class B", "Class C", ...) does
not match RentCast's enum (Single Family/Condo/Townhouse/Manufactured/
Multi-Family/Apartment/Land), and guessing a mapping risks silently
mis-valuing a property. Letting RentCast resolve type/beds/baths/sqft
itself from public records for /properties and /avm/value is strictly
safer than a wrong guess.
"""
import requests
from django.conf import settings

RENTCAST_BASE_URL = 'https://api.rentcast.io/v1'
RENTCAST_TIMEOUT_SECONDS = 20


class RentCastError(Exception):
    """
    Raised for anything that stops a RentCast lookup from completing —
    missing API key, network failure, non-2xx response, or an
    unparseable body. Callers (views.MarketCompsView) catch this and
    store str(e) on PropertyMarketDataSnapshot.fetch_error rather than
    letting a third-party outage 500 the deal drawer.
    """
    pass


def is_configured() -> bool:
    """
    Same "disabled if unconfigured" pattern as
    POSTMARK_INBOUND_WEBHOOK_USERNAME/PASSWORD in settings.py — an empty
    key means this feature simply isn't turned on yet (e.g. a fresh dev
    environment with no RentCast account provisioned), not an error.
    """
    return bool(getattr(settings, 'RENTCAST_API_KEY', '') or '')


def _get(path: str, params: dict) -> dict | list:
    api_key = getattr(settings, 'RENTCAST_API_KEY', '') or ''
    if not api_key:
        raise RentCastError(
            'RentCast is not configured for this environment (RENTCAST_API_KEY is unset).'
        )

    try:
        response = requests.get(
            f'{RENTCAST_BASE_URL}{path}',
            params={k: v for k, v in params.items() if v is not None},
            headers={'X-Api-Key': api_key, 'Accept': 'application/json'},
            timeout=RENTCAST_TIMEOUT_SECONDS,
        )
    except requests.RequestException as e:
        raise RentCastError(f'Could not reach RentCast ({path}): {e}') from e

    if response.status_code == 429:
        raise RentCastError(
            'RentCast rate/quota limit reached for this account. Try again later, or '
            'check the account\'s monthly request usage on rentcast.io.'
        )
    if not response.ok:
        raise RentCastError(f'RentCast returned HTTP {response.status_code} for {path}: {response.text[:300]}')

    try:
        return response.json()
    except ValueError as e:
        raise RentCastError(f'RentCast returned a non-JSON response for {path}.') from e


def fetch_property_record(address: str) -> dict | None:
    """
    GET /v1/properties?address=... — tax assessment history and annual
    property tax bill (and general parcel data) for ONE specific address.

    The docs don't confirm whether this returns a bare array or an
    object (see rentcast_client's research notes) — handled defensively
    here: a list response uses its first (and expected-only, for a single
    address lookup) element; an empty list/None means RentCast has no
    record for this address, which is a valid, non-error outcome (not
    every property is in RentCast's coverage), so this returns None
    rather than raising.
    """
    data = _get('/properties', {'address': address, 'limit': 1})
    if isinstance(data, list):
        return data[0] if data else None
    if isinstance(data, dict):
        return data or None
    return None


def fetch_avm_value(address: str) -> dict | None:
    """GET /v1/avm/value?address=... — current market value estimate (AVM)."""
    data = _get('/avm/value', {'address': address})
    return data if isinstance(data, dict) and data else None


def fetch_rental_comps(address: str, radius_miles: float = 3.0, limit: int = 20) -> list:
    """
    GET /v1/listings/rental/long-term?address=...&radius=... — nearby
    ACTIVE rental listings (default RentCast sort: most recently seen
    first). Returns [] (never raises) if RentCast simply has no nearby
    rental listings for this area — a thin market is a real, displayable
    result ("no comps found nearby"), not a fetch failure.
    """
    data = _get('/listings/rental/long-term', {
        'address': address, 'radius': radius_miles, 'limit': limit, 'status': 'Active',
    })
    return data if isinstance(data, list) else []


def fetch_all_market_data(address: str) -> dict:
    """
    Runs all three calls above for one address and returns their raw
    results together — the exact shape PropertyMarketDataSnapshot's raw_*
    fields are populated from (see MarketCompsView._fetch_and_cache in
    views.py). Each of the three sub-fetches is independent: one failing
    (e.g. RentCast has no tax record for a brand-new construction, or the
    rental-comps call times out) doesn't block the other two — each
    exception is caught and recorded as that section's own None/[] plus a
    per-section error note, rather than failing the whole lookup over one
    of three calls.
    """
    result = {
        'property_record': None,
        'avm_value': None,
        'rental_comps': [],
        'section_errors': {},
    }

    try:
        result['property_record'] = fetch_property_record(address)
    except RentCastError as e:
        result['section_errors']['property_record'] = str(e)

    try:
        result['avm_value'] = fetch_avm_value(address)
    except RentCastError as e:
        result['section_errors']['avm_value'] = str(e)

    try:
        result['rental_comps'] = fetch_rental_comps(address)
    except RentCastError as e:
        result['section_errors']['rental_comps'] = str(e)

    # If EVERY section failed, treat it as a total failure (the caller
    # stores this on PropertyMarketDataSnapshot.fetch_error and shows a
    # single "lookup failed" state rather than 3 empty sub-panels).
    if (
        result['property_record'] is None and result['avm_value'] is None
        and not result['rental_comps'] and result['section_errors']
    ):
        raise RentCastError('; '.join(result['section_errors'].values()))

    return result
