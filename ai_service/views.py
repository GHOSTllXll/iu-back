# backend/ai_service/views.py
import io
import re
import json
import os
import time
import pandas as pd
from datetime import timedelta
from django.http import FileResponse
from django.utils import timezone
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from rest_framework.permissions import IsAuthenticated

from .llm_router import ai_client, AIRateLimitExhausted
from .parsers import (
    extract_document_text,
    extract_t12_text,
    extract_data_from_excel,
    read_excel_with_header_detection,
    FileValidationError,
)
from .excel_generator import generate_underwriting_excel, ExcelGenerationError
from .rent_roll_utils import (
    detect_charge_ledger_columns,
    reshape_charge_ledger,
    detect_rent_column,
    strip_rent_outlier_rows,
    detect_rent_roll_columns,
    ensure_status_column,
    compute_ground_truth,
    RentRollColumnError,
)
from .reconciliation import run_reconciliation
from .document_history import record_processed_documents, record_analysis_report
from .analysis_cache import store_analysis, get_cached_analysis


# ==========================================
# TIER CONFIGURATION
# ==========================================
TIER_BASIC = 'basic'
TIER_PROFESSIONAL = 'professional'
TIER_ENTERPRISE = 'enterprise'
TIER_TRIAL = 'trial'

# Tiers entitled to the 33rd metric (DST / Capex Budget). Basic stays at 32.
TIERS_WITH_DST_CAPEX = {TIER_PROFESSIONAL, TIER_ENTERPRISE}

# Module 1 (Cross-Document Reconciliation) is Enterprise-only.
TIERS_WITH_RECONCILIATION = {TIER_ENTERPRISE}

# Module 2 (Smart Standardization Mapping) is also Enterprise-only. Kept as a
# separate named set from TIERS_WITH_RECONCILIATION even though it's currently
# identical, in case Enterprise features ever diverge across sub-tiers later.
TIERS_WITH_STANDARDIZATION = {TIER_ENTERPRISE}

# Module 4 (Source Provenance). Scoped to ONLY the two OM-derived fields that
# genuinely need citation-based verification — see conversation for why full
# per-metric provenance was deliberately not built (redundant for
# formula-backed T12/Rent Roll numbers, and increases hallucination risk).
TIERS_WITH_PROVENANCE = {TIER_ENTERPRISE}


# ==========================================
# UPLOAD QUOTA CONFIGURATION
# ==========================================
# "Upload" = one analysis (3 source documents = 1 upload). Enforced against
# AnalysisReport rows rather than a separate counter — one row per successful
# analysis already exists (see document_history.py), so counting rows in a
# window IS counting usage, with no risk of a counter drifting out of sync.
#
# Window: rolling 30 days anchored to each Organization's own date_created —
# NOT a shared calendar month. Confirmed with the person building this system.
# Only successful analyses count against quota — a failed upload (bad file,
# parsing error) does not consume quota, also confirmed.
TIER_UPLOAD_LIMITS = {
    TIER_BASIC: 50,
    TIER_PROFESSIONAL: 150,
    TIER_ENTERPRISE: 500,
    TIER_TRIAL: 1,
}

QUOTA_WINDOW_DAYS = 30


def get_quota_window_start(organization):
    """
    Returns the start of the CURRENT 30-day billing-style cycle for this org,
    anchored to organization.date_created. E.g. if the org was created on
    Jan 1st, cycles are [Jan 1 - Jan 31), [Jan 31 - Mar 2), etc. — not simply
    "the last 30 days from right now" (that would be a plain rolling window,
    not a cycle), and not a shared calendar month across all orgs.
    """
    anchor = organization.date_created
    now = timezone.now()
    days_elapsed = (now - anchor).days
    cycles_elapsed = days_elapsed // QUOTA_WINDOW_DAYS
    return anchor + timedelta(days=cycles_elapsed * QUOTA_WINDOW_DAYS)


def get_quota_window_end(organization):
    """The end (exclusive) of the current cycle — i.e. when quota resets."""
    return get_quota_window_start(organization) + timedelta(days=QUOTA_WINDOW_DAYS)


def get_period_usage(organization) -> int:
    """
    Successful analyses recorded within the CURRENT quota cycle.
 
    DELIBERATELY does not filter is_deleted — a report a user has since
    "deleted" (soft-deleted, hidden from their own Outputs page) still
    permanently counts against the quota cycle it was created in. Filtering
    this would let a user delete a report to immediately regain an upload,
    which defeats the entire purpose of quota enforcement.
    """
    from .models import AnalysisReport
    window_start = get_quota_window_start(organization)
    return AnalysisReport.objects.filter(organization=organization, created_at__gte=window_start).count()


def check_upload_quota(organization, tier):
    """
    Returns an error Response if the org has hit its plan's upload limit for
    the current cycle, else None. organization=None (user with no org) skips
    quota enforcement entirely — matches the existing "no org, no tracking"
    pattern used elsewhere (History recording, reports).
    """
    if organization is None:
        return None

    limit = TIER_UPLOAD_LIMITS.get(tier, TIER_UPLOAD_LIMITS[TIER_BASIC])
    used = get_period_usage(organization)

    if used >= limit:
        resets_at = get_quota_window_end(organization)
        return Response({
            'error': (
                f"You've reached your plan's limit of {limit} uploads for this cycle. "
                f"Your quota resets on {resets_at.strftime('%B %d, %Y')}, or you can "
                f"upgrade your plan to continue sooner."
            ),
            'quota_exceeded': True,
            'used': used,
            'limit': limit,
            'resets_at': resets_at.isoformat(),
        }, status=status.HTTP_402_PAYMENT_REQUIRED)

    return None


# Sensible bounds — not enforcing "correctness", just catching obvious garbage
# input (negative rates, 500% LTV, etc.) before it reaches Excel formulas.
DEBT_ASSUMPTION_BOUNDS = {
    'ltv_pct': (0, 100),
    'interest_rate_pct': (0, 30),
    'amortization_years': (1, 50),
    # purchase_price intentionally has NO upper bound here — just a sanity
    # floor (must be positive). Handled separately below since it's required,
    # not defaulted.
}

# Defaults if a field is somehow omitted (frontend pre-fills these, but don't
# trust the client — validate/default server-side too).
DEBT_ASSUMPTION_DEFAULTS = {
    'ltv_pct': 75.0,
    'interest_rate_pct': 7.0,
    'amortization_years': 30,
    # No entry for purchase_price — it's required, not defaulted.
}


def _parse_debt_assumptions(request):
    """
    Reads ltv_pct, interest_rate_pct, amortization_years, and purchase_price
    from the request body. The first three fall back to sensible defaults if
    omitted; purchase_price does NOT — it's required, since there's no safe
    default for a deal-specific number. If the OM stated a price, the
    frontend pre-fills this field with the AI's extracted value; if the user
    edited it (or the OM stated none at all), whatever's in the field is
    what's used here — no merging with the AI's original figure.
    Returns (assumptions_dict, error_response). error_response is None on success.
    """
    assumptions = {}
 
    for field, (low, high) in DEBT_ASSUMPTION_BOUNDS.items():
        raw_value = request.data.get(field)
        default = DEBT_ASSUMPTION_DEFAULTS[field]
 
        if raw_value is None or raw_value == '':
            assumptions[field] = default
            continue
 
        try:
            value = float(raw_value)
        except (TypeError, ValueError):
            return None, Response(
                {'error': f"'{field}' must be a number, got '{raw_value}'."},
                status=status.HTTP_400_BAD_REQUEST
            )
 
        if not (low <= value <= high):
            return None, Response(
                {'error': f"'{field}' must be between {low} and {high}, got {value}."},
                status=status.HTTP_400_BAD_REQUEST
            )
 
        assumptions[field] = value
 
    # purchase_price: required, no default, no upper bound (just sanity floor)
    raw_price = request.data.get('purchase_price')
    if raw_price is None or raw_price == '':
        return None, Response(
            {
                'error': "Target Purchase Price is required. The source documents "
                         "didn't state one, or it wasn't confirmed — please enter "
                         "your assumed purchase price before downloading.",
                'missing_purchase_price': True,
            },
            status=status.HTTP_400_BAD_REQUEST
        )
 
    try:
        price = float(raw_price)
    except (TypeError, ValueError):
        return None, Response(
            {'error': f"'purchase_price' must be a number, got '{raw_price}'."},
            status=status.HTTP_400_BAD_REQUEST
        )
 
    if price <= 0:
        return None, Response(
            {'error': "'purchase_price' must be greater than 0."},
            status=status.HTTP_400_BAD_REQUEST
        )
 
    assumptions['purchase_price'] = price
 
    return assumptions, None


def get_user_tier(request) -> str:
    """
    Reads the requesting user's subscription tier from their Organization.
    Falls back to 'basic' if the user has no organization attached
    (CustomUser.organization is nullable) rather than erroring — adjust this
    if you'd rather that case be a hard failure instead.
    """
    user = request.user
    org = getattr(user, 'organization', None)
    if org is None:
        return TIER_BASIC
    return org.subscription_plan or TIER_BASIC


def get_user_organization(request):
    """Returns the requesting user's Organization, or None if unattached."""
    return getattr(request.user, 'organization', None)


MOCK_METRICS_BASE = {
    "property_metadata": {
        "property_name": "View High Lake (MOCK DATA)",
        "address": "10708 East 98th Terrace, Kansas City, MO",
        "year_built_renovated": "1973",
        "asset_class": "Class C",
        "total_unit_count": 308
    },
    "rent_roll_metrics": {
        "physical_occupancy_pct": 0.65,
        "economic_occupancy_pct": 0.61,
        "total_vacant_units": 108,
        "avg_in_place_monthly_rent": 1154.00,
        "avg_market_monthly_rent": 1225.00,
        "annualized_gpr": 4265000.00,
        "total_annual_concessions": 42000.00
    },
    "t12_revenue_expenses": {
        "net_rental_income": 3100000.00,
        "other_income": 210000.00,
        "total_operating_revenue": 3310000.00,
        "real_estate_taxes": 180000.00,
        "property_liability_insurance": 95000.00,
        "total_utilities": 240000.00,
        "repairs_maintenance": 310000.00,
        "contract_services": 60000.00,
        "marketing_advertising": 25000.00,
        "property_management_fee": 132000.00,
        "payroll_benefits_staffing": 280000.00,
        "general_administrative": 45000.00,
        "total_operating_expenses": 1367000.00,
        "net_operating_income": 1943000.00
    },
    "underwriting_ratios": {
        "operating_expense_ratio": 0.413,
        "total_annual_expenses_per_unit": 4438.00
    },
    "deal_valuation": {
        "target_purchase_price": 34618000.00,
        "entry_cap_rate": 0.0561,
        "acquisition_price_per_unit": 112396.00
        # dst_capex_budget deliberately omitted here — added conditionally
        # by get_mock_metrics() based on tier, same as the real AI path.
    },
    "debt_returns": {
        "cash_on_cash_return_pct": 0.082
    }
}


def get_mock_metrics(tier: str) -> dict:
    """Returns mock metrics shaped correctly for the given tier."""
    metrics = json.loads(json.dumps(MOCK_METRICS_BASE))  # cheap deep copy
    if tier in TIERS_WITH_DST_CAPEX:
        metrics["deal_valuation"]["dst_capex_budget"] = 1300000.00
    if tier in TIERS_WITH_RECONCILIATION:
        # Deliberately mismatched vs the mock rent roll's real numbers, so mock
        # mode actually exercises the reconciliation flagging logic instead of
        # silently passing every time.
        metrics["reconciliation_claims"] = {"om_claimed_occupancy_pct": 0.96}
    if tier in TIERS_WITH_STANDARDIZATION:
        # Non-empty on purpose, so mock mode exercises the standardization
        # display/Excel logic instead of always showing the empty-state path.
        metrics["uncategorized_items"] = [
            {
                "original_line_item": "Misc Admin & Maintenance Fee",
                "allocated_amount": 14000.00,
                "category": "expense"
            }
        ]
    if tier in TIERS_WITH_PROVENANCE:
        metrics["source_citations"] = {
            "dst_capex_budget": {
                "page_location": 14,
                "section_title": "Capital Improvements",
                "exact_text_anchor": "estimated deferred maintenance budget of $1.3M for roofing"
            },
            "om_claimed_occupancy_pct": {
                "page_location": 1,
                "section_title": "Loan Details",
                "exact_text_anchor": "Physical Occ. 96%"
            }
        }
    return metrics


def get_ai_metrics(system_prompt: str, user_prompt: str, tier: str) -> str:
    """
    Wraps ai_client.generate_completion with an escape hatch for local testing.
    Set AI_MOCK_MODE=true in your .env to skip the real LLM call entirely and
    return canned metrics — useful for testing file parsing + Excel generation
    without needing a working Ollama model or OpenAI credits.

    IMPORTANT: turn this back off (or remove the env var) once you're done
    testing — it's a dev convenience, not something you want on in production.
    """
    if os.environ.get("AI_MOCK_MODE", "false").lower() == "true":
        return json.dumps(get_mock_metrics(tier))

    return ai_client.generate_completion(system_prompt=system_prompt, user_prompt=user_prompt)


# Keys the AI response MUST have at the top level before we trust it downstream.
REQUIRED_METRIC_SECTIONS = [
    "property_metadata",
    "rent_roll_metrics",
    "t12_revenue_expenses",
    "underwriting_ratios",
    "deal_valuation",
    "debt_returns",
]


def validate_metrics_shape(metrics: dict):
    """
    Raises FileValidationError if the AI response is missing expected sections.
    """
    if not isinstance(metrics, dict):
        raise FileValidationError("AI response was not a JSON object.")

    missing = [key for key in REQUIRED_METRIC_SECTIONS if key not in metrics]
    if missing:
        raise FileValidationError(
            f"AI response is missing expected section(s): {', '.join(missing)}"
        )


def enforce_tier_entitlements(metrics: dict, tier: str) -> dict:
    """
    Hard enforcement of tier limits, independent of what the AI actually returned.
    Even if the model ignores the prompt and includes gated fields for a lower
    tier, this strips them before the response ever leaves the server.
    Entitlement boundaries should never depend on the AI's cooperation.
    """
    if tier not in TIERS_WITH_DST_CAPEX:
        deal_valuation = metrics.get("deal_valuation")
        if isinstance(deal_valuation, dict):
            deal_valuation.pop("dst_capex_budget", None)

    if tier not in TIERS_WITH_RECONCILIATION:
        metrics.pop("reconciliation", None)

    if tier not in TIERS_WITH_STANDARDIZATION:
        metrics.pop("standardization", None)

    if tier not in TIERS_WITH_PROVENANCE:
        metrics.pop("provenance", None)

    # reconciliation_claims, uncategorized_items, and source_citations are all
    # intermediate fields the AI populates so Python can process them — none
    # are meant to survive in the final response for ANY tier, since
    # process_underwriting_files already consumed them into "reconciliation" /
    # "standardization" / "provenance" respectively.
    metrics.pop("reconciliation_claims", None)
    metrics.pop("uncategorized_items", None)
    metrics.pop("source_citations", None)

    return metrics


def extract_json_from_ai_response(ai_response: str) -> dict:
    if not ai_response or not ai_response.strip():
        raise FileValidationError(
            "The AI returned an empty response. This can happen if the model's "
            "thinking/reasoning consumed the entire token budget, or if the "
            "request was refused. Try again, or check the AI provider's status."
        )
    """
    Strips markdown fences if present, then extracts the first {...} block.
    """
    cleaned = ai_response.replace('```json', '').replace('```', '').strip()

    match = re.search(r'\{.*\}', cleaned, re.DOTALL)
    json_str = match.group(0) if match else cleaned

    try:
        metrics = json.loads(json_str)
    except json.JSONDecodeError as e:
        raise FileValidationError(f"AI returned invalid JSON: {str(e)}")

    validate_metrics_shape(metrics)
    return metrics


# ==========================================
# SHARED PROCESSING FUNCTION (DRY Principle)
# ==========================================
def process_underwriting_files(om_file, t12_file, rent_roll_file, tier: str = TIER_BASIC,
                                 organization=None, uploaded_by=None):
    """
    Shared logic for processing files and extracting metrics.
    Returns: (metrics_dict, rent_roll_df, error_response)

    `tier` controls whether the 33rd metric (DST/Capex Budget) is requested from
    the AI and whether it's allowed to survive in the response — Basic gets 32
    metrics, Professional/Enterprise get 33.

    `organization`/`uploaded_by` are used purely for History recording — see
    document_history.py. Passing organization=None skips recording entirely
    (e.g. a user with no org attached).
    """
    # Quota check happens BEFORE any parsing/AI work — an org over their limit
    # gets rejected immediately rather than wasting processing on files that
    # were never going to be allowed through.
    quota_error = check_upload_quota(organization, tier)
    if quota_error:
        return None, None, quota_error

    pipeline_start_time = time.perf_counter()

    try:
        # 1. Parse documents
        om_text = extract_document_text(om_file)
        t12_text = extract_t12_text(t12_file)
        rent_roll_text = extract_data_from_excel(rent_roll_file)

        # 2. Parse Excel into DataFrame — uses header-row detection instead of
        #    blindly assuming row 0 (real rent rolls often have a title row
        #    above the actual column headers).
        rent_roll_df = read_excel_with_header_detection(rent_roll_file)
        rent_roll_df = rent_roll_df.dropna(how='all').dropna(axis=1, how='all')

        if rent_roll_df.empty:
            raise FileValidationError("Rent roll has no usable rows after cleaning.")

        # 2a. Detect + reshape "charge ledger" format rent rolls (multiple
        # rows per unit, one per charge type) into the standard one-row-
        # per-unit shape. Applies to all tiers — this is format support,
        # not a premium feature. If the file isn't in this format,
        # detect_charge_ledger_columns returns None and nothing changes.
        ledger_cols = detect_charge_ledger_columns(rent_roll_df)
        if ledger_cols and 'rent' not in [str(c).lower() for c in rent_roll_df.columns]:
            # Only reshape if there's no ALREADY-simple rent column — a file
            # that has both a real "Rent" column AND coincidentally a
            # "Charge"/"Amount" pair (unlikely, but possible) shouldn't be
            # reshaped unnecessarily.
            reshaped = reshape_charge_ledger(rent_roll_df, ledger_cols)
            if not reshaped.empty:
                rent_roll_df = reshaped

        # 2b. Strip likely totals/summary rows (e.g. a trailing "Totals" row
        # with a SUM-of-all-units rent value that would otherwise corrupt
        # AVERAGE()-based metrics). Applied for every tier, not just
        # Enterprise — this is a data-integrity fix, not a premium feature.
        rent_col_for_cleanup = detect_rent_column(rent_roll_df)
        rent_roll_df, excluded_rows = strip_rent_outlier_rows(rent_roll_df, rent_col_for_cleanup)
        rent_roll_cleanup_note = None
        if excluded_rows:
            rent_roll_cleanup_note = (
                f"{len(excluded_rows)} row(s) were excluded from the rent roll as likely "
                f"totals/summary rows (rent value far above the median of other rows) — "
                f"not counted as real units."
            )

        combined_context = f"""
        OFFERING MEMORANDUM (OM) EXTRACT:
        {om_text}

        T12 STATEMENT EXTRACT:
        {t12_text}

        RENT ROLL EXTRACT:
        {rent_roll_text}
        """

        system_prompt = _build_system_prompt(tier)

        # 3. Call the AI (or mock, if AI_MOCK_MODE is set)
        ai_response = get_ai_metrics(system_prompt, combined_context, tier)

        # 4. Parse + validate JSON
        metrics = extract_json_from_ai_response(ai_response)

        # 5. Module 1 (Enterprise only): Cross-Document Reconciliation.
        #    Ground truth is computed from the rent roll DataFrame itself —
        #    same data source the Excel formulas use — so a flag here reflects
        #    a real document mismatch, not the AI disagreeing with itself.
        if tier in TIERS_WITH_RECONCILIATION:
            try:
                rent_col, status_col, tenant_col = detect_rent_roll_columns(rent_roll_df)
                status_col, was_inferred = ensure_status_column(rent_roll_df, status_col, tenant_col)
                ground_truth = compute_ground_truth(rent_roll_df, rent_col, status_col)
                flags = run_reconciliation(metrics, rent_roll_df, ground_truth)
                metrics["reconciliation"] = {"flags": flags}
                if was_inferred:
                    # Occupancy wasn't explicitly stated in the rent roll — it was
                    # inferred from whether the Tenant field was populated. Real,
                    # but a genuine inference rather than a stated fact, so the
                    # user should know the ground truth carries that caveat.
                    metrics["reconciliation"]["note"] = (
                        "Occupancy was inferred from the Tenant column (no explicit "
                        "Status column found in the rent roll) — populated tenant name "
                        "= occupied, blank/vacant = vacant."
                    )
            except RentRollColumnError as e:
                # Don't fail the whole request over reconciliation specifically —
                # the core extraction still succeeded. Surface it as a soft note
                # instead. (Excel generation will raise its own clear error later
                # if the rent roll truly can't be mapped.)
                metrics["reconciliation"] = {
                    "flags": [],
                    "note": f"Reconciliation could not run: {str(e)}"
                }

        # 6. Module 2 (Enterprise only): Smart Standardization Mapping.
        #    The AI reports any T12 line item it couldn't confidently map into
        #    a standard category, rather than guessing or silently dropping it.
        #    Python computes the total by summing the list itself — not trusting
        #    any total the AI might also report — same principle as Module 1.
        if tier in TIERS_WITH_STANDARDIZATION:
            raw_items = metrics.get("uncategorized_items") or []
            cleaned_items = []
            uncategorized_total = 0.0

            for item in raw_items:
                if not isinstance(item, dict):
                    continue
                amount = item.get("allocated_amount")
                if not isinstance(amount, (int, float)):
                    continue  # skip malformed entries rather than crash on bad AI output
                cleaned_items.append({
                    "original_line_item": item.get("original_line_item", "Unknown line item"),
                    "allocated_amount": amount,
                    "category": item.get("category") if item.get("category") in ("revenue", "expense") else "expense",
                })
                uncategorized_total += amount

            metrics["standardization"] = {
                "uncategorized_items": cleaned_items,
                "uncategorized_total": round(uncategorized_total, 2),
            }

        # 7. Module 4 (Enterprise only): Source Provenance — scoped to the two
        #    OM-derived fields that genuinely need citation (dst_capex_budget,
        #    om_claimed_occupancy_pct). NOT applied to every metric — see
        #    conversation for why full coverage was deliberately skipped.
        if tier in TIERS_WITH_PROVENANCE:
            raw_citations = metrics.get("source_citations") or {}
            allowed_fields = {"dst_capex_budget", "om_claimed_occupancy_pct"}
            cleaned_citations = {}

            for field, citation in raw_citations.items():
                if field not in allowed_fields or not isinstance(citation, dict):
                    continue
                cleaned_citations[field] = {
                    "page_location": citation.get("page_location"),
                    "section_title": citation.get("section_title"),
                    "exact_text_anchor": citation.get("exact_text_anchor"),
                }

            if cleaned_citations:
                metrics["provenance"] = cleaned_citations

        # 8. Hard-enforce tier entitlements regardless of what the AI returned
        metrics = enforce_tier_entitlements(metrics, tier)

        # Surface the outlier-row cleanup note (if any) to every tier — this
        # is transparency about data quality, not a gated feature.
        if rent_roll_cleanup_note:
            metrics["data_quality_note"] = rent_roll_cleanup_note

        elapsed_seconds = time.perf_counter() - pipeline_start_time

        record_processed_documents(organization, uploaded_by, om_file, t12_file, rent_roll_file, status='completed')
        record_analysis_report(organization, uploaded_by, metrics, tier, processing_seconds=elapsed_seconds)

        return metrics, rent_roll_df, None

    except FileValidationError as e:
        record_processed_documents(organization, uploaded_by, om_file, t12_file, rent_roll_file, status='failed')
        return None, None, Response({'error': str(e)}, status=status.HTTP_400_BAD_REQUEST)
    except AIRateLimitExhausted:
        # Deliberately NOT recorded as 'failed' — this is a transient capacity
        # issue, not something wrong with the documents themselves. The user
        # is expected to retry, and a false "Failed" history entry would be
        # misleading about the actual document quality.
        return None, None, Response({
            'error': "We're experiencing high demand right now. Please try again in a minute or two.",
            'retryable': True
        }, status=status.HTTP_503_SERVICE_UNAVAILABLE)
    except Exception as e:
        record_processed_documents(organization, uploaded_by, om_file, t12_file, rent_roll_file, status='failed')
        return None, None, Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


def _build_system_prompt(tier: str = TIER_BASIC) -> str:
    include_dst_capex = tier in TIERS_WITH_DST_CAPEX
    include_reconciliation = tier in TIERS_WITH_RECONCILIATION
    include_standardization = tier in TIERS_WITH_STANDARDIZATION
    include_provenance = tier in TIERS_WITH_PROVENANCE
    metric_count = 33 if include_dst_capex else 32

    dst_capex_field = ',\n        "dst_capex_budget": number' if include_dst_capex else ''
    dst_capex_guidance = """

    Field-specific guidance:
    - "dst_capex_budget": Extract the deferred maintenance budget, planned physical
      capital repairs, and/or sponsor capital reserves as stated in the Offering
      Memorandum (often labeled "Capex Budget", "Deferred Maintenance", "Capital
      Improvements Budget", or similar). If the OM does not state this explicitly,
      set to null — do not estimate or fabricate a figure.""" if include_dst_capex else ''

    reconciliation_claims_section = """,
      "reconciliation_claims": {
        "om_claimed_occupancy_pct": number
      }""" if include_reconciliation else ''

    reconciliation_guidance = """

    CRITICAL — "reconciliation_claims" section:
    This section is NOT for your own calculations. It exists so our system can
    cross-check the OM's stated claims against the actual Rent Roll data
    independently.
    - "om_claimed_occupancy_pct": Extract EXACTLY the physical occupancy
      percentage stated in the OM itself (e.g. on its cover page, loan details
      section, or property summary) — as a decimal (e.g. 0.96 for 96%). Do NOT
      calculate this yourself from the Rent Roll. Only report what the OM
      document literally states. If the OM does not explicitly state an
      occupancy percentage, set this to null.""" if include_reconciliation else ''

    standardization_section = """,
      "uncategorized_items": [
        {
          "original_line_item": "string",
          "allocated_amount": number,
          "category": "revenue" or "expense"
        }
      ]""" if include_standardization else ''

    standardization_guidance = """

    CRITICAL — "uncategorized_items" section (Chart of Accounts mapping):
    When mapping T12 line items into the categories above, use these standard
    classifications:
    - RUBS, Utility Reimbursement, Valet Trash Income, Pet Rent, Garage Fees -> "other_income"
    - Turnover Paint, Make-Ready Supplies, HVAC Filters, Plumbing Parts -> "repairs_maintenance"
    - Lawn Care, Snow Removal, Exterminator, Fire Alarm Inspection -> "contract_services"
    - Leasing Commissions, Zillow Ads, Banners, Tenant Screening -> "marketing_advertising"

    If you encounter a T12 line item that does NOT cleanly fit any standard
    category (e.g. a cryptic or blended line item like "Misc Admin & Maintenance
    Fee"), do NOT guess and do NOT force it into the closest category. Instead:
    1. Add it to "uncategorized_items" with its exact original text, dollar
       amount, and whether it's revenue or expense.
    2. STILL include its dollar amount in the relevant total
       (total_operating_revenue or total_operating_expenses) so those totals
       remain complete and match the source document — do not exclude
       uncategorized amounts from the totals.
    If every line item maps cleanly, return an empty array for
    "uncategorized_items".""" if include_standardization else ''

    # Provenance is only requested for fields that are actually present for
    # this tier — citing a field that doesn't exist in the response makes no
    # sense. In practice both are true together for Enterprise, but this stays
    # correct even if that ever changes.
    provenance_fields = []
    if include_provenance and include_dst_capex:
        provenance_fields.append('"dst_capex_budget"')
    if include_provenance and include_reconciliation:
        provenance_fields.append('"om_claimed_occupancy_pct"')

    provenance_section = f""",
      "source_citations": {{
        {", ".join(f'{f}: {{"page_location": number, "section_title": "string", "exact_text_anchor": "string"}}' for f in provenance_fields)}
      }}""" if provenance_fields else ''

    provenance_guidance = """

    CRITICAL — "source_citations" section:
    For ONLY the specific fields listed in "source_citations" above, provide
    where in the Offering Memorandum you found that information:
    - "page_location": The page number where this was stated, as shown by the
      "--- PAGE N ---" markers in the OM extract. If the OM extract has no such
      markers (e.g. it came from a Word document with no page breaks), set to null.
    - "section_title": The heading/section name the information appeared under,
      if identifiable (e.g. "Loan Details", "Capital Improvements"). Set to null
      if not identifiable.
    - "exact_text_anchor": A SHORT excerpt (under 15 words) from the source text
      that supports this figure — enough to locate it, not a lengthy quote.
    Do NOT fabricate a page number, section, or quote if you are not confident —
    set the relevant sub-field to null rather than guess. A false citation is
    worse than no citation.""" if provenance_fields else ''

    return f"""
    You are an expert Commercial Real Estate (CRE) Underwriting AI agent.
    Your task is to process the provided OM, Rent Roll, and T12 extracts and output EXACTLY {metric_count} financial and operational metrics in STRICT JSON format.

    Rules:
    1. Output ONLY valid JSON. Do not include markdown formatting like ```json or any conversational text.
    2. If a metric cannot be found, set its value to null.
    3. Use the exact keys provided below.

    Required JSON Structure:
    {{
      "property_metadata": {{
        "property_name": "string",
        "address": "string",
        "year_built_renovated": "string",
        "asset_class": "string",
        "total_unit_count": number
      }},
      "rent_roll_metrics": {{
        "physical_occupancy_pct": number,
        "economic_occupancy_pct": number,
        "total_vacant_units": number,
        "avg_in_place_monthly_rent": number,
        "avg_market_monthly_rent": number,
        "annualized_gpr": number,
        "total_annual_concessions": number
      }},
      "t12_revenue_expenses": {{
        "net_rental_income": number,
        "other_income": number,
        "total_operating_revenue": number,
        "real_estate_taxes": number,
        "property_liability_insurance": number,
        "total_utilities": number,
        "repairs_maintenance": number,
        "contract_services": number,
        "marketing_advertising": number,
        "property_management_fee": number,
        "payroll_benefits_staffing": number,
        "general_administrative": number,
        "total_operating_expenses": number,
        "net_operating_income": number
      }},
      "underwriting_ratios": {{
        "operating_expense_ratio": number,
        "total_annual_expenses_per_unit": number
      }},
      "deal_valuation": {{
        "target_purchase_price": number,
        "entry_cap_rate": number,
        "acquisition_price_per_unit": number{dst_capex_field}
      }},
      "debt_returns": {{
        "cash_on_cash_return_pct": number
      }}{reconciliation_claims_section}{standardization_section}{provenance_section}
    }}{dst_capex_guidance}{reconciliation_guidance}{standardization_guidance}{provenance_guidance}
    """


class UnderwritePropertyView(APIView):
    """
    ENDPOINT 1: Returns JSON metrics (used by the dashboard's "Analyze Documents" button)
    URL: /api/ai/underwrite/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request):
        om_file = request.FILES.get('om_file')
        t12_file = request.FILES.get('t12_file')
        rent_roll_file = request.FILES.get('rent_roll_file')

        if not all([om_file, t12_file, rent_roll_file]):
            return Response({'error': 'All three files are required.'}, status=status.HTTP_400_BAD_REQUEST)

        tier = get_user_tier(request)
        metrics, rent_roll_df, error = process_underwriting_files(
            om_file, t12_file, rent_roll_file, tier=tier,
            organization=get_user_organization(request), uploaded_by=request.user
        )

        if error:
            return error

        # Cache this result so a follow-up "Download Excel" click (which lives
        # inside the modal this response populates) can reuse it instead of
        # re-parsing files and re-calling the AI. See analysis_cache.py.
        analysis_id = store_analysis(metrics, rent_roll_df)

        return Response({'metrics': metrics, 'tier': tier, 'analysis_id': analysis_id})


class UnderwritePropertyDownloadView(APIView):
    """
    ENDPOINT 2: Returns Excel file download
    URL: /api/ai/underwrite/download/

    Prefers reusing a cached analysis result (via analysis_id, from a prior
    call to UnderwritePropertyView) over re-parsing files and re-calling the
    AI — this is what prevents "Analyze then Download" from doing the same
    expensive work twice. Falls back to full reprocessing if no analysis_id
    is given, or if it's expired/unknown — so this endpoint still works fine
    called on its own, just without the caching benefit in that case.
    """
    permission_classes = [IsAuthenticated]

    def post(self, request):
        # Financing assumptions — user-entered on the frontend, re-entered per
        # upload (org-level saved defaults were considered and deliberately not
        # built; see conversation history). These are raw numbers, not derived
        # from the AI extraction, since no source document states a buyer's
        # intended financing terms for a new acquisition.
        debt_assumptions, debt_error = _parse_debt_assumptions(request)
        if debt_error:
            return debt_error

        analysis_id = request.data.get('analysis_id')
        cached = get_cached_analysis(analysis_id) if analysis_id else None

        if cached:
            metrics, rent_roll_df = cached
        else:
            om_file = request.FILES.get('om_file')
            t12_file = request.FILES.get('t12_file')
            rent_roll_file = request.FILES.get('rent_roll_file')

            if not all([om_file, t12_file, rent_roll_file]):
                if analysis_id:
                    # An analysis_id was given but the cache had expired/missed,
                    # AND no files were provided to fall back on — nothing we
                    # can do but ask the user to re-analyze.
                    return Response(
                        {'error': 'Your analysis has expired. Please click "Analyze Documents" again before downloading.'},
                        status=status.HTTP_400_BAD_REQUEST
                    )
                return Response(
                    {'error': 'All three files (om_file, t12_file, rent_roll_file) are required.'},
                    status=status.HTTP_400_BAD_REQUEST
                )

            tier = get_user_tier(request)
            metrics, rent_roll_df, error = process_underwriting_files(
                om_file, t12_file, rent_roll_file, tier=tier,
                organization=get_user_organization(request), uploaded_by=request.user
            )

            if error:
                return error

        try:
            excel_bytes = generate_underwriting_excel(metrics, rent_roll_df, debt_assumptions=debt_assumptions)
        except ExcelGenerationError as e:
            return Response({'error': str(e)}, status=status.HTTP_400_BAD_REQUEST)
        except Exception as e:
            return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

        response = FileResponse(
            io.BytesIO(excel_bytes),
            as_attachment=True,
            filename='Underwriting_Model.xlsx',
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        return response


# ==========================================
# Existing utility views
# ==========================================
class TestAIView(APIView):
    permission_classes = [IsAuthenticated]

    def post(self, request):
        user_prompt = request.data.get('prompt', '')

        if not user_prompt:
            return Response({'error': 'Prompt is required'}, status=status.HTTP_400_BAD_REQUEST)

        system_prompt = """
        You are an expert Commercial Real Estate (CRE) Underwriting AI agent.
        Your sole purpose is to process financial documents and extract data accurately.
        """

        try:
            response_text = ai_client.generate_completion(
                system_prompt=system_prompt,
                user_prompt=user_prompt
            )

            return Response({
                'provider': ai_client.provider,
                'response': response_text
            })

        except Exception as e:
            return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


class DocumentUploadView(APIView):
    permission_classes = [IsAuthenticated]

    def post(self, request):
        om_file = request.FILES.get('om_file')
        t12_file = request.FILES.get('t12_file')
        rent_roll_file = request.FILES.get('rent_roll_file')

        if not all([om_file, t12_file, rent_roll_file]):
            return Response(
                {'error': 'All three files (om_file, t12_file, rent_roll_file) are required.'},
                status=status.HTTP_400_BAD_REQUEST
            )

        try:
            om_text = extract_document_text(om_file)
            t12_text = extract_t12_text(t12_file)
            rent_roll_text = extract_data_from_excel(rent_roll_file)

            return Response({
                'message': 'Documents parsed successfully',
                'extracted_data': {
                    'om_preview': om_text[:500] + "...",
                    't12_preview': t12_text[:500] + "...",
                    'rent_roll_preview': rent_roll_text[:500] + "..."
                }
            })

        except FileValidationError as e:
            return Response({'error': str(e)}, status=status.HTTP_400_BAD_REQUEST)
        except Exception as e:
            return Response({'error': f'Failed to parse documents: {str(e)}'}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


class DocumentHistoryView(APIView):
    """
    Powers the Documents page: live stats + the list of processed documents
    for the requesting user's organization. Excludes soft-deleted records —
    see ProcessedDocumentDeleteView / is_deleted field notes.
    URL: GET /api/ai/documents/?search=<optional>
    """
    permission_classes = [IsAuthenticated]
 
    def get(self, request):
        from .models import ProcessedDocument
 
        organization = get_user_organization(request)
        if organization is None:
            return Response({
                'stats': {'total': 0, 'completed': 0, 'processing': 0, 'failed': 0},
                'documents': []
            })
 
        queryset = ProcessedDocument.objects.filter(organization=organization, is_deleted=False)
 
        search = request.query_params.get('search', '').strip()
        if search:
            queryset = queryset.filter(file_name__icontains=search)
 
        all_docs = ProcessedDocument.objects.filter(organization=organization, is_deleted=False)
        stats = {
            'total': all_docs.count(),
            'completed': all_docs.filter(status='completed').count(),
            'processing': all_docs.filter(status='processing').count(),
            'failed': all_docs.filter(status='failed').count(),
        }
 
        documents = [
            {
                'id': doc.id,
                'name': doc.file_name,
                'company': organization.name,
                'status': doc.get_status_display(),
                'type': doc.file_format,
                'uploaded_at': doc.uploaded_at.isoformat(),
                'size_bytes': doc.file_size_bytes,
            }
            for doc in queryset[:200]
        ]
 
        return Response({'stats': stats, 'documents': documents})


class AnalysisReportListView(APIView):
    """
    Powers the Outputs page card grid. Excludes soft-deleted reports — but
    note this is DISPLAY only. Quota counting (get_period_usage) deliberately
    does NOT use this same filtering — see is_deleted field notes.
    URL: GET /api/ai/reports/?search=<optional>
    """
    permission_classes = [IsAuthenticated]
 
    def get(self, request):
        from .models import AnalysisReport
 
        organization = get_user_organization(request)
        if organization is None:
            return Response({'reports': []})
 
        queryset = AnalysisReport.objects.filter(organization=organization, is_deleted=False)
 
        search = request.query_params.get('search', '').strip()
        if search:
            queryset = queryset.filter(property_name__icontains=search)
 
        reports = [
            {
                'id': report.id,
                'title': report.property_name or 'Untitled Property',
                'company': organization.name,
                'date': report.created_at.isoformat(),
                'status': report.get_status_display(),
            }
            for report in queryset[:200]
        ]
 
        return Response({'reports': reports})


class AnalysisReportDetailView(APIView):
    """
    GET returns the full stored metrics JSON for Preview — 404s if the report
    doesn't exist OR has been soft-deleted (deleted = fully gone from the
    user's perspective, even via direct ID).
    DELETE soft-deletes (is_deleted=True) rather than removing the row —
    CRITICAL: this row still counts toward upload quota for its cycle
    regardless of this flag. See get_period_usage() and is_deleted notes.
    URL: GET/DELETE /api/ai/reports/<id>/
    """
    permission_classes = [IsAuthenticated]
 
    def get(self, request, report_id):
        from .models import AnalysisReport
 
        organization = get_user_organization(request)
        if organization is None:
            return Response({'error': 'No organization associated with this account.'}, status=status.HTTP_404_NOT_FOUND)
 
        try:
            report = AnalysisReport.objects.get(id=report_id, organization=organization, is_deleted=False)
        except AnalysisReport.DoesNotExist:
            return Response({'error': 'Report not found.'}, status=status.HTTP_404_NOT_FOUND)
 
        return Response({
            'id': report.id,
            'title': report.property_name or 'Untitled Property',
            'metrics': report.metrics,
            'created_at': report.created_at.isoformat(),
        })
 
    def delete(self, request, report_id):
        from .models import AnalysisReport
 
        organization = get_user_organization(request)
        if organization is None:
            return Response({'error': 'No organization associated with this account.'}, status=status.HTTP_403_FORBIDDEN)
 
        try:
            report = AnalysisReport.objects.get(id=report_id, organization=organization, is_deleted=False)
        except AnalysisReport.DoesNotExist:
            return Response({'error': 'Report not found.'}, status=status.HTTP_404_NOT_FOUND)
 
        report.is_deleted = True
        report.save(update_fields=['is_deleted'])
        return Response(status=status.HTTP_204_NO_CONTENT)


class DashboardStatsView(APIView):
    """
    Powers the 4 stat cards at the top of the dashboard. Every number here is
    computed from real data — no fabricated stats. Also surfaces the org's
    current upload-quota usage, since giving people visibility into their own
    usage is part of preventing accidental overage, not just hard-blocking it
    after the fact.
    URL: GET /api/ai/dashboard-stats/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request):
        from .models import ProcessedDocument, AnalysisReport
        from django.db.models import Avg

        organization = get_user_organization(request)
        tier = get_user_tier(request)

        if organization is None:
            return Response({
                'documents_processed_total': 0,
                'processed_today': 0,
                'processed_yesterday': 0,
                'extraction_success_rate': None,
                'avg_processing_seconds': None,
                'quota': None,
            })

        now = timezone.now()
        today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        yesterday_start = today_start - timedelta(days=1)

        all_docs = ProcessedDocument.objects.filter(organization=organization, is_deleted=False)
        completed_docs = all_docs.filter(status='completed')
        failed_docs = all_docs.filter(status='failed')

        documents_processed_total = completed_docs.count()
        processed_today = completed_docs.filter(uploaded_at__gte=today_start).count()
        processed_yesterday = completed_docs.filter(
            uploaded_at__gte=yesterday_start, uploaded_at__lt=today_start
        ).count()

        total_attempts = completed_docs.count() + failed_docs.count()
        extraction_success_rate = (
            round(completed_docs.count() / total_attempts * 100, 1)
            if total_attempts > 0 else None
        )

        avg_processing_seconds = AnalysisReport.objects.filter(
            organization=organization,
            processing_seconds__isnull=False,
            is_deleted=False
        ).aggregate(avg=Avg('processing_seconds'))['avg']

        limit = TIER_UPLOAD_LIMITS.get(tier, TIER_UPLOAD_LIMITS[TIER_BASIC])
        used = get_period_usage(organization)
        resets_at = get_quota_window_end(organization)

        return Response({
            'documents_processed_total': documents_processed_total,
            'processed_today': processed_today,
            'processed_yesterday': processed_yesterday,
            'extraction_success_rate': extraction_success_rate,
            'avg_processing_seconds': round(avg_processing_seconds, 1) if avg_processing_seconds else None,
            'quota': {
                'used': used,
                'limit': limit,
                'tier': tier,
                'resets_at': resets_at.isoformat(),
            },
        })

class ProcessedDocumentDeleteView(APIView):
    """
    Soft-deletes a document-history record (sets is_deleted=True, does not
    actually remove the row). Purely cosmetic/informational — ProcessedDocument
    has no quota implication, unlike AnalysisReport — but kept consistent
    with the same pattern for predictability and so "Documents Processed"
    stats don't shrink confusingly out from under a user's own actions in a
    way that looks like real historical data was lost.
    URL: DELETE /api/ai/documents/<id>/
    """
    permission_classes = [IsAuthenticated]
 
    def delete(self, request, document_id):
        from .models import ProcessedDocument
 
        organization = get_user_organization(request)
        if organization is None:
            return Response({'error': 'No organization associated with this account.'}, status=status.HTTP_403_FORBIDDEN)
 
        try:
            doc = ProcessedDocument.objects.get(id=document_id, organization=organization, is_deleted=False)
        except ProcessedDocument.DoesNotExist:
            return Response({'error': 'Document not found.'}, status=status.HTTP_404_NOT_FOUND)
 
        doc.is_deleted = True
        doc.save(update_fields=['is_deleted'])
        return Response(status=status.HTTP_204_NO_CONTENT)