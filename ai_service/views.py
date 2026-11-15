# backend/ai_service/views.py
import io
import re
import json
import hashlib
import os
import time
import uuid
import tempfile
import datetime
import openpyxl
import pandas as pd
from datetime import timedelta
from django.http import FileResponse
from django.utils import timezone
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from rest_framework.permissions import IsAuthenticated

from celery.result import AsyncResult
from .tasks import run_cre_underwriting_task, run_ppm_underwriting_task

from .llm_router import ai_client, AIRateLimitExhausted
from .parsers import (
    extract_document_text,
    extract_t12_text,
    extract_data_from_excel,
    read_excel_with_header_detection,
    validate_uploaded_file,
    FileValidationError,
)
from .excel_generator import generate_underwriting_excel, ExcelGenerationError, CURRENCY_FORMATS, DEFAULT_CURRENCY_CODE
from .custom_template_generator import (
    generate_excel_from_template,
    TemplateValidationError as CustomTemplateValidationError,
    TemplateCapacityError,
    compute_injected_values_snapshot,
    extract_defined_names,
)
from .excel_ppm_generator import build_ppm_master_grid, PPMExcelGenerationError
from .rent_roll_utils import (
    detect_charge_ledger_columns,
    reshape_charge_ledger,
    detect_rent_column,
    strip_rent_outlier_rows,
    detect_rent_roll_columns,
    ensure_status_column,
    compute_ground_truth,
    detect_argus_fields,
    RentRollColumnError,
)
from .reconciliation import run_reconciliation, run_ppm_reconciliation
from .document_history import record_processed_documents, record_analysis_report, record_processed_ppm_document
from .analysis_cache import store_analysis, get_cached_analysis
from .ppm_splitter import extract_ppm_target_text, PPMSplitterError
from .document_classifier import (
    classify_document,
    DocumentClassificationError,
    CLASSIFIABLE_EXTENSIONS,
    DOCUMENT_TYPE_UNCLASSIFIED,
)


import logging
logger = logging.getLogger(__name__)

from django.db import connection, OperationalError, InterfaceError
from .variance_analysis import run_variance_analysis, run_line_item_anomaly_scan, VarianceAnalysisSkipped

def _db_write_with_retry(fn):
    """
    Runs a DB write once; on a dead/stale MySQL connection (common after a
    long AI call on shared hosting), closes the connection and retries once
    so Django opens a fresh one. Non-connection errors are re-raised as-is.
    """
    try:
        return fn()
    except (OperationalError, InterfaceError) as e:
        logger.warning(f"DB write failed on stale connection — retrying once: {e}")
        connection.close()
        return fn()


def _safe_record_failure(organization, uploaded_by, om_file, t12_file, rent_roll_file):
    """
    Best-effort history write for failed analyses. Must NEVER raise — if the
    original failure was itself a dead DB connection, recording 'failed' on
    that same connection would mask the real error and crash the request.
    """
    try:
        _db_write_with_retry(lambda: record_processed_documents(
            organization, uploaded_by, om_file, t12_file, rent_roll_file, status='failed'
        ))
    except Exception as e:
        logger.warning(f"Could not record failed document history: {e}")


def _safe_record_ppm_failure(organization, uploaded_by, ppm_file):
    """
    PPM equivalent of _safe_record_failure — single-file shape, same
    never-raise guarantee (see above).
    """
    try:
        _db_write_with_retry(lambda: record_processed_ppm_document(
            organization, uploaded_by, ppm_file, status='failed'
        ))
    except Exception as e:
        logger.warning(f"Could not record failed PPM document history: {e}")


# ==========================================
# TIER CONFIGURATION
# ==========================================
TIER_BASIC = 'basic'
TIER_PROFESSIONAL = 'professional'
TIER_ENTERPRISE = 'enterprise'
TIER_TRIAL = 'trial'

# Tiers entitled to the 33rd metric (DST / Capex Budget). Basic stays at 32.
TIERS_WITH_DST_CAPEX = {TIER_PROFESSIONAL, TIER_ENTERPRISE, TIER_TRIAL}

# Module 1 (Cross-Document Reconciliation) is Enterprise-only.
TIERS_WITH_RECONCILIATION = {TIER_ENTERPRISE, TIER_TRIAL}

# Module 2 (Smart Standardization Mapping) is also Enterprise-only. Kept as a
# separate named set from TIERS_WITH_RECONCILIATION even though it's currently
# identical, in case Enterprise features ever diverge across sub-tiers later.
TIERS_WITH_STANDARDIZATION = {TIER_ENTERPRISE, TIER_TRIAL}

# Module 4 (Source Provenance). Scoped to ONLY the two OM-derived fields that
# genuinely need citation-based verification — see conversation for why full
# per-metric provenance was deliberately not built (redundant for
# formula-backed T12/Rent Roll numbers, and increases hallucination risk).
TIERS_WITH_PROVENANCE = {TIER_ENTERPRISE, TIER_TRIAL}
# ==========================================
# Module 7 (Multi-Family Variance Analysis Engine) — same Enterprise/Trial
# gating as the other advanced modules. See variance_analysis.py for the
# full deterministic logic and design rationale.
TIERS_WITH_VARIANCE_ANALYSIS = {TIER_ENTERPRISE, TIER_TRIAL}
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

    organization.quota_override (nullable) takes priority over the standard
    tier-based limit when set — lets a specific org get a bespoke rolling
    30-day ceiling (e.g. a negotiated arrangement) without needing a new
    tier or a hardcoded per-org special case in this function.
    """
    if organization is None:
        return None

    if organization.quota_override is not None:
        limit = organization.quota_override
    else:
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


# ==========================================
# INSTITUTIONAL PPM CONFIGURATION
# ==========================================
# PPMs run heavier (up to ~300 pages) than the standard OM/T12 documents, so
# they get a larger cap than parsers.MAX_FILE_SIZE (25MB) — passed explicitly
# to validate_uploaded_file rather than changing the shared default, so the
# existing CRE pipeline's limit is untouched.
PPM_MAX_FILE_SIZE = 100 * 1024 * 1024  # 100 MB


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


def _dataframe_to_json_safe(df):
    """
    Feature 3: converts a pandas DataFrame into a plain JSON-serializable
    structure (list of dicts) safe to store in a JSONField — via pandas' own
    to_json (NaN -> null, numpy scalar types -> plain Python types, dates ->
    ISO 8601 strings) rather than hand-rolling DataFrame -> dict conversion,
    which silently mishandles exactly those cases (a raw .to_dict() call
    leaves numpy.int64/float64 and NaN in place, and JSONField's underlying
    json.dumps() either rejects or mis-renders every one of those).

    Returns [] for None/empty input rather than raising — callers store the
    result directly on AnalysisReport.rent_roll_data/t12_monthly_data, where
    "no rows" and "not available" are both legitimately representable as an
    empty/absent value, not an error.
    """
    if df is None or df.empty:
        return []
    return json.loads(df.to_json(orient='records', date_format='iso'))


def _parse_currency_code(request):
    """
    Shared currency_code parsing/validation — used by both
    UnderwritePropertyDownloadView and PropertyProjectDownloadView (Feature
    3), pulled out so the two never drift. An omitted field silently
    defaults to USD (a frontend dropdown that always sends a value, or an
    old cached client, shouldn't break); an explicitly-sent but unrecognized
    code IS rejected loudly, since that can only mean a frontend/backend
    mismatch worth surfacing. Returns (currency_code, error_response) —
    error_response is None on success.
    """
    currency_code = (request.data.get('currency_code') or DEFAULT_CURRENCY_CODE).strip().upper()
    if currency_code not in CURRENCY_FORMATS:
        return None, Response(
            {'error': f"'{currency_code}' is not a supported currency. Supported: {', '.join(sorted(CURRENCY_FORMATS.keys()))}."},
            status=status.HTTP_400_BAD_REQUEST
        )
    return currency_code, None


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
    },
    # This base mock property (View High Lake) is multifamily, so every
    # industrial field is null here — exactly the shape the real AI path
    # returns for a non-industrial property (see _build_system_prompt's
    # industrial_metrics_guidance). This exercises the "industrial tab is
    # absent for a standard multifamily export" path in AI_MOCK_MODE.
    "industrial_metrics": {
        "rentable_square_footage": None,
        "annual_base_rent": None,
        "annual_nnn_reimbursements": None
    },
    # Same reasoning as industrial_metrics above, for the Office asset class.
    "office_metrics": {
        "tenant_rsf": None,
        "building_total_rsf": None,
        "annual_base_rent_per_sf": None,
        "expense_stop_value": None,
        "ti_allowance_total": None,
        "lease_term_months": None
    },
    # Same reasoning again, for the Retail asset class.
    "retail_metrics": {
        "tenant_gla": None,
        "center_total_gla": None,
        "annual_base_rent_per_sf": None,
        "percentage_rent_rate": None,
        "tenant_breakpoint_threshold": None,
        "tenant_gross_annual_sales": None
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
        # silently passing every time. t12_claimed_vacancy_loss_pct is a
        # plausible non-triggering value included only for schema shape —
        # the mock rent roll's actual physical occupancy is 65%, so Rule B
        # (which requires >95%) never fires in mock mode regardless of this
        # figure; exercising Rule B specifically needs a real/manual test.
        metrics["reconciliation_claims"] = {
            "om_claimed_occupancy_pct": 0.96,
            "t12_claimed_vacancy_loss_pct": 0.03,
        }
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


# Deliberately mismatched: 29.25M + 15.75M (Senior Loan + Sponsor Equity) =
# 45M, not the 45.5M claimed_master_acquisition_cost below — so AI_MOCK_MODE
# exercises the RED reconciliation path instead of silently passing every
# time, same principle as MOCK_METRICS_BASE's deliberately-mismatched
# reconciliation_claims for the CRE pipeline above.
MOCK_PPM_METRICS = {
    "target_asset_name": "Sunbelt Logistics Portfolio (MOCK DATA)",
    "sponsor_platform_name": "Meridian Capital Partners",
    "offering_volume": 45000000.00,
    "loan_to_cost_ratio": 0.65,
    "weighted_exit_cap": 0.058,
    "total_sponsor_fees": 1350000.00,
    "year_1_yield": 0.071,
    "raw_sources_and_uses_items": [
        {"line_item": "Senior Loan Proceeds", "amount": 29250000.00},
        {"line_item": "Sponsor Equity", "amount": 15750000.00},
        {"line_item": "Acquisition Fee", "amount": 450000.00},
    ],
    "claimed_master_acquisition_cost": 45500000.00,
    "provenance_citations": {
        "target_asset_name": {
            "page_location": 12,
            "section_title": "Sources and Uses of Funds",
            "exact_text_anchor": "Sunbelt Logistics Portfolio"
        },
    },
}


def get_ppm_ai_metrics(system_prompt: str, user_prompt: str) -> str:
    """
    PPM equivalent of get_ai_metrics — same AI_MOCK_MODE escape hatch for
    local testing without a working AI provider (see get_ai_metrics for the
    "turn this back off before production" caveat, which applies here too).
    Kept as a separate function rather than adding a branch to get_ai_metrics
    because the two pipelines' mock data shapes are completely different and
    get_ai_metrics' tier parameter has no PPM equivalent.
    """
    if os.environ.get("AI_MOCK_MODE", "false").lower() == "true":
        return json.dumps(MOCK_PPM_METRICS)

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

    if tier not in TIERS_WITH_VARIANCE_ANALYSIS:
        metrics.pop("variance_analysis", None)
        # Same tier gate as Module 7 above — line_item_anomalies is the
        # generalized counterpart computed alongside it (see 7b.), not a
        # separately-gated feature.
        metrics.pop("line_item_anomalies", None)

    # reconciliation_claims, uncategorized_items, and source_citations are all
    # intermediate fields the AI populates so Python can process them — none
    # are meant to survive in the final response for ANY tier, since
    # process_underwriting_files already consumed them into "reconciliation" /
    # "standardization" / "provenance" respectively.
    metrics.pop("reconciliation_claims", None)
    metrics.pop("uncategorized_items", None)
    metrics.pop("source_citations", None)

    return metrics


# Keys the PPM AI response MUST have at the top level before we trust it
# downstream. raw_sources_and_uses_items + claimed_master_acquisition_cost
# are the two fields run_ppm_reconciliation needs — the AI extracts them,
# Python does the actual math (see reconciliation.py). provenance_citations
# is deliberately NOT required here, same as the CRE path's source_citations
# — a citation-less response is still usable, just less verifiable.
PPM_REQUIRED_FIELDS = [
    "target_asset_name",
    "sponsor_platform_name",
    "offering_volume",
    "loan_to_cost_ratio",
    "weighted_exit_cap",
    "total_sponsor_fees",
    "year_1_yield",
    "raw_sources_and_uses_items",
    "claimed_master_acquisition_cost",
]


def validate_ppm_metrics_shape(metrics: dict):
    """
    PPM equivalent of validate_metrics_shape — raises FileValidationError if
    the AI response is missing expected top-level fields.
    """
    if not isinstance(metrics, dict):
        raise FileValidationError("AI response was not a JSON object.")

    missing = [key for key in PPM_REQUIRED_FIELDS if key not in metrics]
    if missing:
        raise FileValidationError(
            f"AI response is missing expected PPM field(s): {', '.join(missing)}"
        )


def _parse_ai_json_response(ai_response: str) -> dict:
    """
    Shared by both the CRE and PPM pipelines: strips markdown fences if
    present, extracts the first {...} block, and parses it as JSON. Does NOT
    validate shape — callers apply their own validator (validate_metrics_shape
    or validate_ppm_metrics_shape) since the two pipelines expect completely
    different top-level keys.
    """
    if not ai_response or not ai_response.strip():
        raise FileValidationError(
            "The AI returned an empty response. This can happen if the model's "
            "thinking/reasoning consumed the entire token budget, or if the "
            "request was refused. Try again, or check the AI provider's status."
        )

    cleaned = ai_response.replace('```json', '').replace('```', '').strip()

    match = re.search(r'\{.*\}', cleaned, re.DOTALL)
    json_str = match.group(0) if match else cleaned

    try:
        return json.loads(json_str)
    except json.JSONDecodeError as e:
        raise FileValidationError(f"AI returned invalid JSON: {str(e)}")


def extract_json_from_ai_response(ai_response: str) -> dict:
    """
    CRE pipeline entry point — unchanged behavior from before this function
    was split into _parse_ai_json_response + per-pipeline validators.
    """
    metrics = _parse_ai_json_response(ai_response)
    validate_metrics_shape(metrics)
    return metrics


def extract_ppm_json_from_ai_response(ai_response: str) -> dict:
    """PPM pipeline entry point — same parsing, PPM-shaped validation."""
    metrics = _parse_ai_json_response(ai_response)
    validate_ppm_metrics_shape(metrics)
    return metrics

# ==========================================
# SHARED PROCESSING FUNCTION (DRY Principle)
# ==========================================
def process_underwriting_files(om_file, t12_file, rent_roll_file, tier: str = TIER_BASIC,
                                 organization=None, uploaded_by=None, project_id=None,
                                 asking_price=None):
    """
    Shared logic for processing files and extracting metrics.
    Returns: (metrics_dict, rent_roll_df, error_response, deal_scorecard)

    `tier` controls whether the 33rd metric (DST/Capex Budget) is requested from
    the AI and whether it's allowed to survive in the response — Basic gets 32
    metrics, Professional/Enterprise get 33.

    `organization`/`uploaded_by` are used purely for History recording — see
    document_history.py. Passing organization=None skips recording entirely
    (e.g. a user with no org attached).

    `project_id` (Feature 3): forwarded as-is to record_analysis_report — see
    its docstring for exactly how it's resolved. None (the default) means
    "start a brand-new deal", which is every call site's behavior before
    this parameter existed.

    Sprint "Workflow Upgrades" Feature 1 (10-Second "Kill or Proceed"
    Scorecard & Missing Doc Check):

    - om_file/t12_file/rent_roll_file may now each individually be None —
      any NON-EMPTY subset of the three is accepted (at least one document),
      not just all-three-or-nothing. This lets an analyst run an initial
      screen on a partial submission (e.g. a broker sends only a Rent Roll
      + T12, no OM yet) instead of being blocked until every document
      arrives. Every step below that depends on a specific missing file is
      skipped cleanly (not an error), and the AI is told exactly which
      document(s) are absent via a "[NOT PROVIDED]" marker in its prompt —
      it already has a standing instruction (see _build_system_prompt,
      "If a metric cannot be found, set its value to null") to null out
      whatever that marker makes unknowable, rather than guessing.
    - asking_price: the analyst's own entered/confirmed price for this deal
      (NOT the AI-extracted deal_valuation.target_purchase_price) —
      threaded through to the Scorecard's cap-rate/price-per-unit math and
      persisted onto the PropertyProject record. None when the analyst
      didn't enter one for this run.
    - The 4th return value, deal_scorecard, is the 10-second "kill or
      proceed" read-out itself — see _compute_deal_scorecard's own
      docstring. Callers that don't need it (the Excel-download path, which
      always has all 3 files anyway) simply discard it.

    NOTE: this loosened gate applies ONLY to this function (the Analyze/
    Scorecard path). UnderwritePropertyDownloadView's own direct-
    reprocessing gate deliberately still requires all 3 files, unchanged —
    full Excel generation is structurally dependent on a real rent roll
    DataFrame and isn't part of this feature's scope.
    """
    # Quota check happens BEFORE any parsing/AI work — an org over their limit
    # gets rejected immediately rather than wasting processing on files that
    # were never going to be allowed through.
    quota_error = check_upload_quota(organization, tier)
    if quota_error:
        return None, None, quota_error, None

    if not any([om_file, t12_file, rent_roll_file]):
        return None, None, Response(
            {'error': 'At least one document (OM, T12, or Rent Roll) is required.'},
            status=status.HTTP_400_BAD_REQUEST
        ), None

    documents_provided = {
        'om': om_file is not None,
        't12': t12_file is not None,
        'rent_roll': rent_roll_file is not None,
    }

    pipeline_start_time = time.perf_counter()

    try:
        # 1. Parse documents — extract_document_text/extract_t12_text
        #    (parsers.py) are already None-tolerant and return "" immediately
        #    for a missing file. extract_data_from_excel is NOT (it's also
        #    used standalone elsewhere where a rent roll is always required),
        #    so it's guarded here explicitly instead.
        om_text = extract_document_text(om_file)
        t_after_om = time.perf_counter()
        t12_text = extract_t12_text(t12_file)
        t_after_t12 = time.perf_counter()
        rent_roll_text = extract_data_from_excel(rent_roll_file) if rent_roll_file is not None else ""
        t_after_rr_text = time.perf_counter()

        # 2. Parse Excel into DataFrame — uses header-row detection instead of
        #    blindly assuming row 0 (real rent rolls often have a title row
        #    above the actual column headers). Skipped entirely when no rent
        #    roll was submitted: rent_roll_df stays an EMPTY DataFrame, which
        #    every downstream consumer below is made to tolerate explicitly.
        rent_roll_cleanup_note = None
        if rent_roll_file is not None:
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
            if excluded_rows:
                rent_roll_cleanup_note = (
                    f"{len(excluded_rows)} row(s) were excluded from the rent roll as likely "
                    f"totals/summary rows (rent value far above the median of other rows) — "
                    f"not counted as real units."
                )
        else:
            rent_roll_df = pd.DataFrame()
        t_after_rr_processing = time.perf_counter()

        _NOT_PROVIDED_MARKER = (
            "[NOT PROVIDED — this document was not submitted for this run. "
            "Per your instructions, set every field that depends on it to "
            "null rather than estimating or inferring it from the other "
            "documents.]"
        )

        combined_context = f"""
        OFFERING MEMORANDUM (OM) EXTRACT:
        {om_text if om_file is not None else _NOT_PROVIDED_MARKER}

        T12 STATEMENT EXTRACT:
        {t12_text if t12_file is not None else _NOT_PROVIDED_MARKER}

        RENT ROLL EXTRACT:
        {rent_roll_text if rent_roll_file is not None else _NOT_PROVIDED_MARKER}
        """

        system_prompt = _build_system_prompt(tier)

        # 3. Call the AI (or mock, if AI_MOCK_MODE is set)
        t_before_ai = time.perf_counter()
        ai_response = get_ai_metrics(system_prompt, combined_context, tier)
        t_after_ai = time.perf_counter()

        # The AI call can take 1-2+ minutes. Any DB connection opened earlier
        # in this request (e.g. the quota check at the top of this function)
        # may have gone stale during that wait — same root cause as the
        # earlier ActiveUserMiddleware fix, just a different DB write this
        # time. Unconditionally closing here (rather than the conditional
        # close_if_unusable_or_obsolete(), which never fires because Django
        # has no way to know the idle connection went stale) forces a fresh
        # connection to open automatically on the next query — protects
        # every database write for the rest of the function, not just one
        # specific spot.
        connection.close()

        # 4. Parse + validate JSON
        metrics = extract_json_from_ai_response(ai_response)

        # 5. Module 1 (Enterprise only): Cross-Document Reconciliation.
        #    Ground truth is computed from the rent roll DataFrame itself —
        #    same data source the Excel formulas use — so a flag here reflects
        #    a real document mismatch, not the AI disagreeing with itself.
        if tier in TIERS_WITH_RECONCILIATION:
            if rent_roll_file is not None:
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
            else:
                # Sprint "Workflow Upgrades" Feature 1: no rent roll submitted —
                # there's no ground truth to reconcile against at all.
                metrics["reconciliation"] = {
                    "flags": [],
                    "note": "Reconciliation skipped: no rent roll was submitted.",
                }

        t_after_module1 = time.perf_counter()

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

        t_after_module2 = time.perf_counter()

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

        t_after_module4 = time.perf_counter()

        # 7b. Module 7 (Enterprise/Trial only): Multi-Family Variance
        #     Analysis Engine. Fully deterministic — see variance_analysis.py
        #     for the full design rationale. Only runs for Excel-format T12s
        #     with real chronological monthly columns; cleanly skipped
        #     otherwise (PDF T12s, annual-only statements).
        # Feature 3: captured here (module scope of this try block) rather
        # than only inside the branch below, so it survives to the
        # record_analysis_report call further down regardless of which
        # branch ran — None whenever this tier/format combination never
        # produces a parsed monthly grid in the first place (PDF T12s,
        # annual-only statements, or tiers without variance analysis at all).
        t12_df_for_persistence = None

        if tier in TIERS_WITH_VARIANCE_ANALYSIS:
            if t12_file is None or rent_roll_file is None:
                # Sprint "Workflow Upgrades" Feature 1: this module pairs a
                # parsed T12 monthly grid against the rent roll — needs BOTH,
                # so a partial submission missing either one skips cleanly
                # rather than crashing on a None file.
                _missing_for_variance = [
                    label for file_obj, label in ((t12_file, "T12"), (rent_roll_file, "Rent Roll"))
                    if file_obj is None
                ]
                _variance_skip_reason = (
                    f"Skipped: {' and '.join(_missing_for_variance)} not submitted for this run."
                )
                metrics["variance_analysis"] = {"skipped": True, "reason": _variance_skip_reason}
                metrics["line_item_anomalies"] = {"skipped": True, "reason": _variance_skip_reason}
            else:
                try:
                    t12_ext = os.path.splitext(t12_file.name)[1].lower()
                    if t12_ext in ('.xlsx', '.xls', '.csv'):
                        t12_file.seek(0)  # already consumed once by extract_t12_text() above
                        t12_df = read_excel_with_header_detection(t12_file)
                        t12_df_for_persistence = t12_df
                        rent_col_va, status_col_va, tenant_col_va = detect_rent_roll_columns(rent_roll_df)
                        variance_result = run_variance_analysis(t12_df, rent_roll_df, unit_col=tenant_col_va)
                        metrics["variance_analysis"] = variance_result

                        # Generalized (any-line-item) anomaly scan feeding the
                        # workbook's separate 'Monthly Variance' tab — reuses
                        # this exact same already-parsed t12_df, no extra file
                        # I/O, and is fully independent of Module 7 above (see
                        # variance_analysis.run_line_item_anomaly_scan).
                        metrics["line_item_anomalies"] = run_line_item_anomaly_scan(t12_df)
                    else:
                        metrics["variance_analysis"] = {
                            "skipped": True,
                            "reason": "Reconciliation Skipped: Source Document Lacks Monthly Granularity (PDF T12 not supported for this analysis).",
                        }
                        metrics["line_item_anomalies"] = {
                            "skipped": True,
                            "reason": "Monthly Variance Skipped: Source Document Lacks Monthly Granularity (PDF T12 not supported for this analysis).",
                        }
                except Exception as e:
                    # Never let this optional module crash the whole analysis —
                    # same principle as Module 1's RentRollColumnError handling.
                    metrics["variance_analysis"] = {
                        "skipped": True,
                        "reason": f"Variance analysis could not run: {str(e)}",
                    }
                    metrics["line_item_anomalies"] = {
                        "skipped": True,
                        "reason": f"Monthly variance scan could not run: {str(e)}",
                    }

        t_after_module7 = time.perf_counter()

        # 8. Hard-enforce tier entitlements regardless of what the AI returned
        metrics = enforce_tier_entitlements(metrics, tier)

        # Surface the outlier-row cleanup note (if any) to every tier — this
        # is transparency about data quality, not a gated feature.
        if rent_roll_cleanup_note:
            metrics["data_quality_note"] = rent_roll_cleanup_note

        # Sprint "Workflow Upgrades" Feature 1: the 10-second Kill-or-Proceed
        # Scorecard — missing-document flags, rent-roll data-quality flags,
        # and the three quick ratios (in-place cap rate / price-per-unit-or-
        # SF / opex ratio) — computed from the SAME metrics/rent_roll_df this
        # run just produced. Re-derived read-only from persisted data later,
        # whenever this deal's pipeline drawer is reopened (see
        # PropertyProjectDetailView.get() + _compute_deal_scorecard itself).
        deal_scorecard = _compute_deal_scorecard(metrics, rent_roll_df, documents_provided, asking_price)

        # elapsed_seconds is measured BEFORE the DB writes below, matching
        # the original behavior exactly — this is what powers the "Avg
        # Processing" dashboard stat, so it must reflect actual pipeline
        # processing time, not include the DB write itself.
        elapsed_seconds = time.perf_counter() - pipeline_start_time

        t_before_db = time.perf_counter()

        # Retry once on stale MySQL connections — shared hosts often kill idle
        # connections during the long AI call even after connection.close()
        # above (e.g. race with wait_timeout, or a reconnect that dies mid-write).
        _db_write_with_retry(lambda: record_processed_documents(
            organization, uploaded_by, om_file, t12_file, rent_roll_file, status='completed'
        ))
        _db_write_with_retry(lambda: record_analysis_report(
            organization, uploaded_by, metrics, tier, processing_seconds=elapsed_seconds,
            project_id=project_id,
            # Sprint "Workflow Upgrades" Feature 1: explicitly None (NOT an
            # empty-rows dict) when no rent roll was submitted — can_download
            # (PropertyProjectDetailView / _serialize_project_summary) relies
            # on this being falsy for "no rent roll on this run", not merely
            # an empty one.
            rent_roll_data={"rows": _dataframe_to_json_safe(rent_roll_df)} if rent_roll_file is not None else None,
            t12_monthly_data={"rows": _dataframe_to_json_safe(t12_df_for_persistence)} if t12_df_for_persistence is not None else None,
            asking_price=asking_price,
            documents_provided=documents_provided,
        ))

        t_after_db = time.perf_counter()

       # TIMING BREAKDOWN — logged via the standard logging module so it
        # reliably lands in passenger.log regardless of DEBUG setting.
        # Temporary instrumentation to find where the "AI call takes ~100s
        # but total request takes 5+ minutes" gap is actually coming from.
        logger.info(
            f"TIMING BREAKDOWN — "
            f"OM parse: {t_after_om - pipeline_start_time:.1f}s | "
            f"T12 parse: {t_after_t12 - t_after_om:.1f}s | "
            f"RentRoll text: {t_after_rr_text - t_after_t12:.1f}s | "
            f"RentRoll DF processing (header/ledger/outlier): {t_after_rr_processing - t_after_rr_text:.1f}s | "
            f"AI call: {t_after_ai - t_before_ai:.1f}s | "
            f"Module 1 (Reconciliation): {t_after_module1 - t_after_ai:.1f}s | "
            f"Module 2 (Standardization): {t_after_module2 - t_after_module1:.1f}s | "
            f"Module 4 (Provenance): {t_after_module4 - t_after_module2:.1f}s | "
            f"Module 7 (Variance Analysis): {t_after_module7 - t_after_module4:.1f}s | "
            f"DB writes: {t_after_db - t_before_db:.1f}s | "
            f"TOTAL (excl. DB writes): {elapsed_seconds:.1f}s"
        )

        return metrics, rent_roll_df, None, deal_scorecard

    except FileValidationError as e:
        _safe_record_failure(organization, uploaded_by, om_file, t12_file, rent_roll_file)
        return None, None, Response({'error': str(e)}, status=status.HTTP_400_BAD_REQUEST), None
    except AIRateLimitExhausted:
        # Deliberately NOT recorded as 'failed' — this is a transient capacity
        # issue, not something wrong with the documents themselves. The user
        # is expected to retry, and a false "Failed" history entry would be
        # misleading about the actual document quality.
        return None, None, Response({
            'error': "We're experiencing high demand right now. Please try again in a minute or two.",
            'retryable': True
        }, status=status.HTTP_503_SERVICE_UNAVAILABLE), None
    except Exception as e:
        _safe_record_failure(organization, uploaded_by, om_file, t12_file, rent_roll_file)
        return None, None, Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR), None
def _build_system_prompt(tier: str = TIER_BASIC) -> str:
    include_dst_capex = tier in TIERS_WITH_DST_CAPEX
    include_reconciliation = tier in TIERS_WITH_RECONCILIATION
    include_standardization = tier in TIERS_WITH_STANDARDIZATION
    include_provenance = tier in TIERS_WITH_PROVENANCE
    metric_count = 48 if include_dst_capex else 47

    # "retail_metrics" — same always-included, always-nullable pattern as
    # industrial_metrics/office_metrics. Unlike office_metrics's dropped
    # Expense Overage formula, every field referenced by the Retail Analysis
    # Excel formulas (Sales Overage, Percentage Rent Owed, Total Gross Rent
    # Revenue) IS present in this schema, so nothing here is a stub for a
    # missing figure. "percentage_rent_rate" is a DECIMAL FRACTION (0.05 for
    # 5%), matching every other percentage field already in this schema
    # (e.g. physical_occupancy_pct) — not a whole-number percent.
    retail_metrics_guidance = """

    Field-specific guidance:
    - "retail_metrics": This section applies ONLY to retail properties
      (strip centers, shopping centers, or regional malls) with a
      Percentage Rent / sales-breakpoint lease structure. If the subject
      property is NOT a retail asset, or the source documents don't state
      these figures, set ALL SIX fields to null — do not estimate, and do
      not repurpose figures from another asset class to fill these in.
      - "tenant_gla": The Gross Leasable Area (GLA), in square feet, of the
        specific retail shop/unit being underwritten.
      - "center_total_gla": The total Gross Leasable Area of the entire
        shopping center/mall, as stated in the OM.
      - "annual_base_rent_per_sf": The flat ANNUAL base rent rate charged per
        square foot, before any percentage rent.
      - "percentage_rent_rate": The commission rate the tenant owes on sales
        above their breakpoint, as a DECIMAL FRACTION (e.g. 0.05 for a
        5% percentage rent clause), not a whole number.
      - "tenant_breakpoint_threshold": The exact gross sales dollar volume
        (natural or artificial breakpoint, as stated) the tenant must exceed
        before percentage rent applies.
      - "tenant_gross_annual_sales": The tenant's actual historical annual
        gross sales dollar volume, as reported in a sales log or estoppel, if
        stated. If the source documents don't report actual tenant sales,
        set to null — do not estimate from the breakpoint or rent alone."""

    # "office_metrics" — same always-included, always-nullable pattern as
    # industrial_metrics below. Note "expense_stop_value" is extracted as a
    # raw informational figure ONLY — there is deliberately no "Expense
    # Overage" formula built on it yet, because that would need the
    # building's ACTUAL operating expense per SF, which is not part of this
    # schema. Wiring an overage formula against data we never extract would
    # produce a formula pointing at nothing; that's a real follow-up ask, not
    # something to fake tonight. "lease_term_months" is explicitly in MONTHS
    # (not "months or years") so the Excel formula below has one unambiguous
    # unit to work with.
    office_metrics_guidance = """

    Field-specific guidance:
    - "office_metrics": This section applies ONLY to office properties (a
      single tenant suite or a multi-tenant office building), typically
      leased with a Base Year, Expense Stop, and/or Tenant Improvement (TI)
      allowance structure. If the subject property is NOT an office asset, or
      the source documents don't state these figures, set ALL SIX fields to
      null — do not estimate, and do not repurpose residential or industrial
      figures to fill these in.
      - "tenant_rsf": The square footage footprint of the specific tenant
        suite being underwritten (often labeled "Usable Square Feet"/"USF" or
        "Rentable Square Feet"/"RSF" on the lease abstract or rent roll).
      - "building_total_rsf": The total rentable square footage of the entire
        office building/tower, as stated in the OM.
      - "annual_base_rent_per_sf": The starting ANNUAL base rent rate charged
        per square foot (before expense reimbursements or TI amortization).
      - "expense_stop_value": The maximum annual operating expense (per SF or
        total, as stated) the landlord absorbs before the tenant must
        contribute — often labeled "Expense Stop" or tied to a "Base Year".
        Extracted for reference only; no downstream formula depends on it yet.
      - "ti_allowance_total": The total (lump-sum, not per-SF) cash
        improvement/build-out budget given to the tenant, as stated in the
        lease or OM.
      - "lease_term_months": The total lease duration, converted to MONTHS
        even if the source document states it in years (e.g. "5-year lease"
        -> 60)."""

    # "industrial_metrics" — ALWAYS included in the schema (not tier-gated:
    # asset class is a property of the deal, not the subscription), but every
    # field is explicitly nullable and the guidance below tells the AI to
    # null the whole section out for non-industrial (e.g. multifamily)
    # properties. This mirrors the existing "if a metric cannot be found, set
    # to null" rule already stated below, so a standard multifamily
    # underwrite behaves exactly as it did before this section existed — it
    # just gets back three more nulls that downstream code already discards
    # (see excel_generator.py: the new Industrial Analysis tab only appears
    # when at least one of these three is non-null).
    industrial_metrics_guidance = """

    Field-specific guidance:
    - "industrial_metrics": This section applies ONLY to industrial,
      warehouse, distribution, or logistics properties leased on a triple-net
      (NNN) or similar expense-reimbursement basis. If the subject property is
      NOT an industrial asset (e.g. it's multifamily), or the source documents
      don't state these figures, set ALL THREE fields to null — do not
      estimate, and do not repurpose residential rent-roll figures to fill
      these in.
      - "rentable_square_footage": The property's total rentable square
        footage (RSF), as stated in the OM or Rent Roll (often labeled "RSF",
        "Total SF", "Building SF", or "GLA").
      - "annual_base_rent": The total ANNUAL base rent (before any NNN/expense
        reimbursements), summed across all tenant leases if the property is
        multi-tenant.
      - "annual_nnn_reimbursements": The total ANNUAL triple-net (NNN) or CAM
        expense reimbursement income billed to tenants, as stated in the T12
        or Rent Roll (often labeled "NNN Reimbursement", "CAM Recovery", or
        "Expense Reimbursement Income")."""

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
        "om_claimed_occupancy_pct": number,
        "t12_claimed_vacancy_loss_pct": number
      }""" if include_reconciliation else ''

    reconciliation_guidance = """

    CRITICAL — "reconciliation_claims" section:
    This section is NOT for your own calculations. It exists so our system can
    cross-check the OM's/T12's stated claims against independently-computed
    ground truth from the Rent Roll.
    - "om_claimed_occupancy_pct": Extract EXACTLY the physical occupancy
      percentage stated in the OM itself (e.g. on its cover page, loan details
      section, or property summary) — as a decimal (e.g. 0.96 for 96%). Do NOT
      calculate this yourself from the Rent Roll. Only report what the OM
      document literally states. If the OM does not explicitly state an
      occupancy percentage, set this to null.
    - "t12_claimed_vacancy_loss_pct": Extract EXACTLY the economic vacancy
      loss / vacancy & concession loss percentage AS STATED on the T12 itself
      — this is commonly a line item near the top of the statement (e.g.
      "Vacancy Loss", "Economic Vacancy", "Vacancy & Concessions") shown as a
      percentage of Gross Potential Rent, as a decimal (e.g. 0.08 for 8%). Do
      NOT derive or calculate this from other figures — only report a
      percentage that is literally printed on the T12. If the T12 only shows
      this as a dollar amount with no percentage stated, or doesn't show it at
      all, set this to null.""" if include_reconciliation else ''

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
      }},
      "industrial_metrics": {{
        "rentable_square_footage": number,
        "annual_base_rent": number,
        "annual_nnn_reimbursements": number
      }},
      "office_metrics": {{
        "tenant_rsf": number,
        "building_total_rsf": number,
        "annual_base_rent_per_sf": number,
        "expense_stop_value": number,
        "ti_allowance_total": number,
        "lease_term_months": number
      }},
      "retail_metrics": {{
        "tenant_gla": number,
        "center_total_gla": number,
        "annual_base_rent_per_sf": number,
        "percentage_rent_rate": number,
        "tenant_breakpoint_threshold": number,
        "tenant_gross_annual_sales": number
      }}{reconciliation_claims_section}{standardization_section}{provenance_section}
    }}{industrial_metrics_guidance}{office_metrics_guidance}{retail_metrics_guidance}{dst_capex_guidance}{reconciliation_guidance}{standardization_guidance}{provenance_guidance}
    """


def _build_ppm_system_prompt() -> str:
    """
    Institutional PPM pipeline's prompt builder. Lives here next to
    _build_system_prompt (its CRE equivalent) and PPMUnderwriteView, NOT in
    llm_router.py — that module is purely the provider-routing/retry
    transport layer and contains no prompt content today; keeping prompt
    text out of it preserves that separation of concerns.

    CRITICAL design point (see reconciliation.run_ppm_reconciliation): the
    AI is explicitly instructed to extract raw Sources & Uses line items and
    the claimed acquisition cost as separate raw figures, and told NOT to
    compute or judge whether they reconcile — that comparison is done in
    Python afterward, so a reconciliation flag reflects a real document
    mismatch, never the AI grading its own math.
    """
    return """
    You are an expert Institutional Real Estate Private Placement analyst.
    You will be given extracted text from SPECIFIC PAGES of a Private Placement
    Memorandum (PPM) — the financial summary pages only (Sources and Uses, Fee
    Structure, and Historical Operations / Pro Forma sections), not the full
    legal document. Extract the following fields and output STRICT JSON only.

    Rules:
    1. Output ONLY valid JSON. Do not include markdown formatting like ```json or any conversational text.
    2. If a field cannot be found in the provided text, set its value to null.
    3. Use the exact keys provided below.
    4. Do NOT calculate whether the Sources and Uses total matches the acquisition
       cost, and do NOT state an opinion on whether the numbers reconcile — simply
       report each raw line item and each raw figure exactly as the document
       states them. A separate system performs that comparison independently.

    Required JSON structure:
    {
      "target_asset_name": "string",
      "sponsor_platform_name": "string",
      "offering_volume": number,
      "loan_to_cost_ratio": number,
      "weighted_exit_cap": number,
      "total_sponsor_fees": number,
      "year_1_yield": number,
      "raw_sources_and_uses_items": [
        {
          "line_item": "string",
          "amount": number
        }
      ],
      "claimed_master_acquisition_cost": number,
      "provenance_citations": {
        "target_asset_name": {"page_location": number, "section_title": "string", "exact_text_anchor": "string"},
        "offering_volume": {"page_location": number, "section_title": "string", "exact_text_anchor": "string"},
        "loan_to_cost_ratio": {"page_location": number, "section_title": "string", "exact_text_anchor": "string"},
        "weighted_exit_cap": {"page_location": number, "section_title": "string", "exact_text_anchor": "string"},
        "total_sponsor_fees": {"page_location": number, "section_title": "string", "exact_text_anchor": "string"},
        "year_1_yield": {"page_location": number, "section_title": "string", "exact_text_anchor": "string"}
      }
    }

    Field-specific guidance:
    - "loan_to_cost_ratio", "weighted_exit_cap", "year_1_yield": express as decimals
      (e.g. 0.65 for 65%, 0.055 for 5.5%), not whole-number percentages.
    - "raw_sources_and_uses_items": list EVERY line item from the Sources and Uses
      table exactly as labeled in the document, with its dollar amount. Include
      both "Sources" and "Uses" line items — do not net them against each other or
      pre-sum them into a single figure. If a line item's amount is unclear or
      unstated, omit that line item rather than guessing a number.
    - "claimed_master_acquisition_cost": the total property/asset acquisition cost
      as explicitly stated in the document (often labeled "Purchase Price",
      "Acquisition Cost", or appearing as a "Total Uses" line) — extract exactly
      what is stated, do not derive it by summing other figures yourself.

    CRITICAL — "provenance_citations" section:
    For each field listed above, provide where in the provided text you found
    that information:
    - "page_location": the page number shown by the "--- PAGE N ---" markers in
      the extract. Set to null if you cannot identify it.
    - "section_title": the heading/section name the information appeared under,
      if identifiable. Set to null if not identifiable.
    - "exact_text_anchor": a SHORT excerpt (under 15 words) from the source text
      that supports this figure.
    Do NOT fabricate a page number, section, or quote if you are not confident —
    set the relevant sub-field to null rather than guess. A false citation is
    worse than no citation. If you cannot cite a field at all, omit it from
    "provenance_citations" entirely rather than including it with all-null values.
    """


# Temp-file staging directory for uploads about to be handed off to a
# Celery task - see _save_upload_to_temp's docstring for why this exists.
# Each task deletes its own temp file(s) in a finally block on both the
# success and failure path; a cron job also sweeps anything older than an
# hour as a backstop against a worker crashing before reaching that finally
# block (see deploy notes).
PENDING_UPLOAD_DIR = os.path.join(tempfile.gettempdir(), 'ppm_underwriting_pending_uploads')


def _save_upload_to_temp(uploaded_file):
    """
    Saves an in-memory/temp-file Django UploadedFile to a plain file on disk
    so a Celery task - running in a completely separate OS process, after
    this request has already finished - can read it. Django's UploadedFile
    objects are bound to the request and can't be serialized across the
    Celery broker; a disk path (a plain string) can.

    Returns (temp_path, original_filename). The caller enqueues both into
    the Celery task, which re-wraps the file in django.core.files.File
    (restoring the .name/.seek() interface the existing parsers expect) and
    deletes the temp file once it's done with it.

    Sprint "Workflow Upgrades" Feature 1 (Missing Doc Check): uploaded_file
    may now legitimately be None — a partial submission (1 or 2 of the 3
    core documents). Returns (None, None) immediately in that case rather
    than crashing on uploaded_file.name; the Celery task (run_cre_
    underwriting_task) is responsible for not opening a None path.
    """
    if uploaded_file is None:
        return None, None

    os.makedirs(PENDING_UPLOAD_DIR, exist_ok=True)
    ext = os.path.splitext(uploaded_file.name)[1]
    temp_path = os.path.join(PENDING_UPLOAD_DIR, f"{uuid.uuid4().hex}{ext}")
    with open(temp_path, 'wb') as out:
        for chunk in uploaded_file.chunks():
            out.write(chunk)
    return temp_path, uploaded_file.name


class ClassifyDocumentView(APIView):
    """
    ENDPOINT 0: Intelligent Document Classification Layer — the front of the
    ingestion pipeline, ahead of UnderwritePropertyView. Given ONE uploaded
    file, runs a fast local keyword scan (see document_classifier.py) to
    guess whether it's a Rent Roll, a Trailing-12, or an Offering
    Memorandum — no AI call, no file parsing beyond a small sample.

    This exists so the dashboard's upload area can be a single mixed-file
    drop zone instead of 3 separate boxes: the frontend calls this once per
    dropped file and uses document_type to auto-place it into the om_file /
    t12_file / rent_roll_file slot, entirely client-side. It does NOT change
    how UnderwritePropertyView itself works — that endpoint still expects
    the same 3 named fields it always has, sent together in one request, and
    runs the exact same combined OM+T12+RentRoll AI analysis it did before
    this classification layer existed. "Routing" a classified file therefore
    means nothing more than placing it in the right named slot; the vacancy
    calculation (rent_roll_utils.compute_ground_truth), the T12 revenue/NOI
    extraction, and the reconciliation math (reconciliation.run_reconciliation)
    are all pre-existing logic inside that combined pipeline, not duplicated
    here.

    DESIGN PRINCIPLE — fail loud, not silent (see document_classifier.py's
    module docstring): a low-confidence or tied result comes back as
    UNCLASSIFIED rather than a guess, so the frontend can ask the user to
    place that one file manually instead of silently feeding a Rent Roll
    into the T12 slot (or vice versa).

    Deliberately NOT gated by upload quota or tier — this never calls the
    AI, so it costs nothing to run, and gating it would defeat the point of
    an instant per-file sort.

    URL: POST /api/ai/classify-document/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request):
        uploaded_file = request.FILES.get('file')
        if uploaded_file is None:
            return Response({'error': 'file is required.'}, status=status.HTTP_400_BAD_REQUEST)

        if uploaded_file.size == 0:
            return Response({
                'error': f"'{uploaded_file.name}' is empty (0 bytes).",
                'document_type': DOCUMENT_TYPE_UNCLASSIFIED,
            }, status=status.HTTP_400_BAD_REQUEST)

        ext = os.path.splitext(uploaded_file.name)[1].lower()
        if ext not in CLASSIFIABLE_EXTENSIONS:
            return Response({
                'error': (
                    f"'{uploaded_file.name}' has an unsupported extension '{ext}'. "
                    f"Allowed: {', '.join(sorted(CLASSIFIABLE_EXTENSIONS))}"
                ),
                'document_type': DOCUMENT_TYPE_UNCLASSIFIED,
            }, status=status.HTTP_400_BAD_REQUEST)

        try:
            result = classify_document(uploaded_file)
        except DocumentClassificationError as e:
            # Unreadable file (corrupt, scanned-image-only, etc.) still comes
            # back as a 200 with UNCLASSIFIED rather than a hard error — the
            # user can still place it in a slot manually, and the real
            # read/validation error will surface properly if they try to
            # Analyze with it regardless. This mirrors how a genuinely
            # unreadable file is handled once it reaches the real pipeline.
            return Response({
                'document_type': DOCUMENT_TYPE_UNCLASSIFIED,
                'scores': {},
                'note': str(e),
            })

        return Response(result)


class UnderwritePropertyView(APIView):
    """
    ENDPOINT 1: Enqueues the CRE/Multifamily extraction pipeline as a
    background Celery task and returns a task_id immediately, instead of
    blocking the request on the AI call (which can take 1-2+ minutes - see
    the TIMING BREAKDOWN log line inside process_underwriting_files). The
    frontend polls TaskStatusView with this task_id until it resolves to
    the exact same {metrics, tier, analysis_id} shape this endpoint used to
    return directly, so no downstream response-handling code had to change.

    URL: /api/ai/underwrite/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request):
        om_file = request.FILES.get('om_file')
        t12_file = request.FILES.get('t12_file')
        rent_roll_file = request.FILES.get('rent_roll_file')

        # Sprint "Workflow Upgrades" Feature 1 (Missing Doc Check): loosened
        # from "all three required" to "at least one" — an analyst can run
        # an initial screen on a partial submission (e.g. a broker sends
        # only a Rent Roll + T12, no OM yet). process_underwriting_files
        # (via the Celery task below) handles whichever subset arrives,
        # nulling out whatever a missing document makes unknowable rather
        # than guessing, and the Scorecard it returns flags exactly what's
        # still missing.
        if not any([om_file, t12_file, rent_roll_file]):
            return Response(
                {'error': 'At least one document (OM, T12, or Rent Roll) is required.'},
                status=status.HTTP_400_BAD_REQUEST
            )

        tier = get_user_tier(request)

        # Feature 3: present only when this upload came from the "Re-upload &
        # Re-analyze Files" button inside a specific PropertyProject view —
        # absent (the default) from every other upload path, which always
        # starts a brand-new deal. Not validated here beyond basic parsing;
        # record_analysis_report re-checks it belongs to the caller's own
        # organization before ever using it, so a bad/foreign id just falls
        # back to "start a new deal" rather than needing to be rejected here.
        raw_project_id = request.data.get('project_id')
        try:
            project_id = int(raw_project_id) if raw_project_id not in (None, '') else None
        except (TypeError, ValueError):
            project_id = None

        # Sprint "Workflow Upgrades" Feature 1: the analyst's own entered/
        # confirmed asking price for this deal (the upload screen's
        # "Target Purchase Price" field) — optional, and distinct from
        # whatever price the OM itself states (see process_underwriting_
        # files' own docstring, or PropertyProject.asking_price in
        # models.py, for exactly why these are kept separate).
        raw_asking_price = request.data.get('asking_price')
        try:
            asking_price = float(raw_asking_price) if raw_asking_price not in (None, '') else None
        except (TypeError, ValueError):
            asking_price = None

        om_path, om_name = _save_upload_to_temp(om_file)
        t12_path, t12_name = _save_upload_to_temp(t12_file)
        rent_roll_path, rent_roll_name = _save_upload_to_temp(rent_roll_file)

        try:
            task = run_cre_underwriting_task.delay(
                om_path, om_name, t12_path, t12_name, rent_roll_path, rent_roll_name,
                tier, request.user.pk, project_id, asking_price,
            )
        except Exception:
            # Enqueueing itself failed (e.g. Redis unreachable) - the temp
            # files we just wrote will never be picked up by a task to clean
            # them, so remove them ourselves rather than leaking them until
            # the hourly stale-upload cron sweep catches them. Any of the
            # three may be None now (Feature 1: partial submissions) —
            # os.remove(None) would raise TypeError, not OSError, so skip
            # a None path entirely rather than attempting it.
            for p in (om_path, t12_path, rent_roll_path):
                if not p:
                    continue
                try:
                    os.remove(p)
                except OSError:
                    pass
            raise
        return Response({'task_id': task.id, 'status': 'queued'}, status=status.HTTP_202_ACCEPTED)

def _build_underwriting_excel_response(request, organization, metrics, rent_roll_df, debt_assumptions, currency_code, analysis=None):
    """
    Given a fully-resolved metrics/rent_roll_df pair, builds and returns the
    Excel FileResponse (or an error Response). This is the custom-template-
    vs-standard branching shared by UnderwritePropertyDownloadView (a fresh
    or just-cached analysis) and PropertyProjectDownloadView (Feature 3:
    regenerating from a PropertyProject's persisted history) — pulled out
    into one place specifically so the two call sites can never drift from
    each other on this logic.

    `analysis` (Feature 4, Cell Override Audit Log): the AnalysisReport this
    download is FOR, if the caller has one in hand — only
    PropertyProjectDownloadView does today. When given, the exact values
    just written into the generated file (scalars + per-tenant rows) are
    captured onto analysis.injected_values_snapshot, which a later
    "Re-upload Edited Underwriting File" diffs against. Deliberately
    captured HERE (at generation time), not back when the analysis was
    first created — 3 of the 18 scalar fields are debt assumptions
    (ltv/rate/amortization) typed in at download time, which this function
    already has fresh from THIS request, and can differ between repeat
    downloads of the same analysis. UnderwritePropertyDownloadView doesn't
    pass one — its direct-download path has no AnalysisReport handle
    without a larger refactor of process_underwriting_files' return value,
    and the pipeline download path above is the one this feature's brief
    actually describes ("weeks later, from the pipeline"). A download made
    only through that older path simply has no snapshot yet — override
    tracking becomes available for a deal once it's been downloaded at
    least once from its PropertyProject view.
    """
    # Custom-template branch: if this org has its own proprietary Excel
    # template on file (see ClientExcelTemplate / custom_template_generator.py),
    # skip our own standard workbook entirely and fill in THEIRS instead —
    # same metrics/rent_roll_df, completely different output path. Falls
    # back to the standard generator below for every org without one.
    #
    # template_choice lets the frontend's upload-flow dropdown override
    # this. Deliberately defaults to 'auto' (preserving the exact
    # behavior that existed before this parameter did — prefer the org's
    # custom template whenever one is active) rather than defaulting to
    # 'standard', so an org already relying on automatic substitution
    # (e.g. Colliers — see CUSTOM_TEMPLATE_NAMING_GUIDE.md) keeps getting
    # their own workbook on every existing integration that doesn't send
    # this field at all, not just ones updated to the new dropdown.
    template_choice = (request.data.get('template_choice') or 'auto').strip().lower()
    if template_choice not in ('auto', 'standard', 'custom'):
        return Response(
            {'error': f"'{template_choice}' is not a valid template_choice. Use 'auto', 'standard', or 'custom'."},
            status=status.HTTP_400_BAD_REQUEST
        )

    from .models import ClientExcelTemplate
    custom_template = (
        ClientExcelTemplate.objects.filter(organization=organization, is_active=True).first()
        if organization else None
    )

    if template_choice == 'custom' and not custom_template:
        return Response(
            {'error': "Your organization has no active custom Excel template to use. Upload one in Settings, or choose 'System Standard'."},
            status=status.HTTP_400_BAD_REQUEST
        )

    # 'standard' is an explicit opt-out even when a custom template IS
    # active — the one case the pre-existing automatic behavior could
    # never do at all (there was no way to get the generic workbook for
    # a template-enabled org). 'auto'/'custom' both use it when present.
    use_custom_template = custom_template is not None and template_choice != 'standard'

    def _save_injected_values_snapshot():
        # Feature 4: only when the caller actually has an AnalysisReport to
        # attach this to (see this function's own docstring) — and only
        # called AFTER generation has already succeeded below, so a failed
        # download never overwrites a perfectly good earlier snapshot with
        # nothing useful.
        if analysis is None:
            return
        # Feature 2 (Immutable Snapshots): once THIS analysis's snapshot
        # has been locked as a deal's permanent baseline (see
        # ExcelBaselineLock's own docstring), never overwrite it again —
        # re-downloading a locked analysis's Excel (e.g. to send another
        # copy to IC) must keep reproducing the EXACT frozen baseline, not
        # silently drift if a re-entered debt assumption differs between
        # downloads. A deal can still be re-analyzed after locking — that
        # creates a BRAND NEW AnalysisReport, which is free to have its own
        # snapshot computed normally; only this specific, already-locked
        # row is frozen.
        if analysis.baseline_locks.exists():
            return
        analysis.injected_values_snapshot = compute_injected_values_snapshot(
            metrics, debt_assumptions, rent_roll_df
        )
        analysis.save(update_fields=['injected_values_snapshot'])

    if use_custom_template:
        try:
            excel_bytes, template_report = generate_excel_from_template(
                custom_template.file, metrics, rent_roll_df, debt_assumptions=debt_assumptions
            )
        except (CustomTemplateValidationError, TemplateCapacityError) as e:
            return Response(
                {'error': f"Your organization's custom Excel template couldn't be used for this analysis: {e}"},
                status=status.HTTP_400_BAD_REQUEST
            )
        except Exception as e:
            return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

        logger.info("Custom template export for org %s: %s", organization.name, template_report)
        _save_injected_values_snapshot()

        is_macro_enabled = custom_template.original_filename.lower().endswith('.xlsm')
        return FileResponse(
            io.BytesIO(excel_bytes),
            as_attachment=True,
            filename=custom_template.original_filename or 'Underwriting_Model.xlsm',
            content_type=(
                'application/vnd.ms-excel.sheet.macroEnabled.12' if is_macro_enabled
                else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        )

    try:
        excel_bytes = generate_underwriting_excel(metrics, rent_roll_df, debt_assumptions=debt_assumptions, currency_code=currency_code)
    except ExcelGenerationError as e:
        return Response({'error': str(e)}, status=status.HTTP_400_BAD_REQUEST)
    except Exception as e:
        return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

    _save_injected_values_snapshot()

    return FileResponse(
        io.BytesIO(excel_bytes),
        as_attachment=True,
        filename='Underwriting_Model.xlsx',
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


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

        currency_code, currency_error = _parse_currency_code(request)
        if currency_error:
            return currency_error

        analysis_id = request.data.get('analysis_id')
        cached = get_cached_analysis(analysis_id) if analysis_id else None

        organization = get_user_organization(request)

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
            # Sprint "Workflow Upgrades" Feature 1: process_underwriting_files
            # now returns a 4th value (deal_scorecard) — discarded here, since
            # this gate above still requires all 3 files unchanged, and Excel
            # download has no use for the Scorecard anyway.
            metrics, rent_roll_df, error, _deal_scorecard = process_underwriting_files(
                om_file, t12_file, rent_roll_file, tier=tier,
                organization=organization, uploaded_by=request.user
            )

            if error:
                return error

        return _build_underwriting_excel_response(
            request, organization, metrics, rent_roll_df, debt_assumptions, currency_code
        )

def process_ppm_file(ppm_file, manual_pages, tier, organization=None, uploaded_by=None):
    """
    Shared logic for the institutional PPM pipeline - extracted out of
    PPMUnderwriteView.post so both the Celery task (run_ppm_underwriting_task)
    and this view can share one implementation, mirroring
    process_underwriting_files' role for the CRE pipeline above.
    Organization-gating (has_ppm_access) and request-level parsing (the
    ppm_file-required check, manual_pages string parsing) stay in the view
    itself, since those need the request object directly and are cheap,
    fast checks worth rejecting on before ever touching a Celery worker.

    Returns: (metrics_dict, error_response_or_None) - same two-value shape
    process_underwriting_files uses for the CRE pipeline (minus the
    DataFrame, which the PPM pipeline has no equivalent of).
    """
    quota_error = check_upload_quota(organization, tier)
    if quota_error:
        return None, quota_error

    try:
        validate_uploaded_file(ppm_file, {'.pdf'}, label="PPM Document", max_size_bytes=PPM_MAX_FILE_SIZE)
    except FileValidationError as e:
        return None, Response({'error': str(e)}, status=status.HTTP_400_BAD_REQUEST)

    pipeline_start_time = time.perf_counter()

    try:
        ppm_text, located_sections = extract_ppm_target_text(ppm_file, manual_pages=manual_pages)
    except PPMSplitterError as e:
        return None, Response({
            'error': str(e),
            'needs_manual_page_range': True,
        }, status=status.HTTP_400_BAD_REQUEST)
    except FileValidationError as e:
        return None, Response({'error': str(e)}, status=status.HTTP_400_BAD_REQUEST)

    system_prompt = _build_ppm_system_prompt()

    try:
        ai_response = get_ppm_ai_metrics(system_prompt, ppm_text)
        connection.close()
        extracted_data = extract_ppm_json_from_ai_response(ai_response)
    except FileValidationError as e:
        _safe_record_ppm_failure(organization, uploaded_by, ppm_file)
        return None, Response({'error': str(e)}, status=status.HTTP_400_BAD_REQUEST)
    except AIRateLimitExhausted:
        return None, Response({
            'error': "We're experiencing high demand right now. Please try again in a minute or two.",
            'retryable': True
        }, status=status.HTTP_503_SERVICE_UNAVAILABLE)
    except Exception as e:
        _safe_record_ppm_failure(organization, uploaded_by, ppm_file)
        return None, Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

    reconciliation_result = run_ppm_reconciliation(
        extracted_data.get('raw_sources_and_uses_items'),
        extracted_data.get('claimed_master_acquisition_cost'),
    )

    metrics = {
        'target_asset_name': extracted_data.get('target_asset_name'),
        'sponsor_platform_name': extracted_data.get('sponsor_platform_name'),
        'offering_volume': extracted_data.get('offering_volume'),
        'loan_to_cost_ratio': extracted_data.get('loan_to_cost_ratio'),
        'weighted_exit_cap': extracted_data.get('weighted_exit_cap'),
        'total_sponsor_fees': extracted_data.get('total_sponsor_fees'),
        'year_1_yield': extracted_data.get('year_1_yield'),
        'provenance': extracted_data.get('provenance_citations') or {},
        'reconciliation': reconciliation_result,
        'located_sections': located_sections,
    }

    elapsed_seconds = time.perf_counter() - pipeline_start_time

    _db_write_with_retry(lambda: record_processed_ppm_document(
        organization, uploaded_by, ppm_file, status='completed'
    ))
    _db_write_with_retry(lambda: record_analysis_report(
        organization, uploaded_by, metrics, tier,
        processing_seconds=elapsed_seconds, document_type='PPM'
    ))

    return metrics, None


class PPMUnderwriteView(APIView):
    """
    ENDPOINT: Institutional PPM analysis. Enqueues process_ppm_file as a
    background Celery task and returns a task_id immediately - same
    migration as UnderwritePropertyView above, for the same reason (the AI
    call here routinely takes 1-2+ minutes). Still processes exactly ONE PPM
    document per request; the frontend's "Institutional PPM Batch Console"
    still loops this call once per uploaded file (up to 5 per batch)
    client-side, polling each task to completion before firing the next.

    Gated on organization.has_ppm_access, which is INDEPENDENT of
    subscription_plan - a feature flag rather than a new tier, so PPM access
    can be granted per-org without restructuring TIER_UPLOAD_LIMITS. Counts
    against the same rolling 30-day upload quota as the CRE pipeline (see
    document_history.record_analysis_report / get_period_usage - usage
    counting doesn't filter by document_type, by design).

    URL: POST /api/ai/ppm/underwrite/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request):
        organization = get_user_organization(request)

        if organization is None or not organization.has_ppm_access:
            return Response(
                {'error': 'Institutional PPM analysis is not enabled for your organization.'},
                status=status.HTTP_403_FORBIDDEN
            )

        ppm_file = request.FILES.get('ppm_file')
        if ppm_file is None:
            return Response({'error': 'ppm_file is required.'}, status=status.HTTP_400_BAD_REQUEST)

        manual_pages_raw = (request.data.get('manual_pages') or '').strip()
        manual_pages = None
        if manual_pages_raw:
            try:
                manual_pages = sorted({int(p.strip()) for p in manual_pages_raw.split(',') if p.strip()})
            except ValueError:
                return Response(
                    {'error': "'manual_pages' must be a comma-separated list of page numbers, e.g. '4,5,12,13'."},
                    status=status.HTTP_400_BAD_REQUEST
                )
            if not manual_pages:
                manual_pages = None

        tier = get_user_tier(request)

        temp_path, temp_name = _save_upload_to_temp(ppm_file)
        try:
            task = run_ppm_underwriting_task.delay(
                temp_path, temp_name, manual_pages, tier, request.user.pk,
            )
        except Exception:
            try:
                os.remove(temp_path)
            except OSError:
                pass
            raise
        return Response({'task_id': task.id, 'status': 'queued'}, status=status.HTTP_202_ACCEPTED)


class TaskStatusView(APIView):
    """
    Polling endpoint shared by both the CRE and PPM Celery tasks (see
    ai_service/tasks.py and UnderwritePropertyView / PPMUnderwriteView
    above). On completion, returns EXACTLY the same body the old fully
    synchronous endpoints used to return directly - {metrics, tier,
    analysis_id} on success, or the identical error shape
    ({error, retryable, needs_manual_page_range, ...} with the original
    status code) on a handled failure - so the frontend's existing
    response-handling and error-handling code needed no changes beyond
    wrapping the initial request in a poll loop.

    URL: GET /api/ai/task-status/<task_id>/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, task_id):
        result = AsyncResult(task_id)

        if not result.ready():
            return Response({'status': 'processing'})

        if result.failed():
            logger.error(f"Celery task {task_id} failed: {result.result!r}")
            return Response(
                {'error': 'Something went wrong while processing your document. Please try again.'},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR
            )

        task_result = result.result

        if isinstance(task_result, dict) and task_result.get('error'):
            return Response(
                task_result.get('data', {'error': 'Request failed.'}),
                status=task_result.get('status_code', status.HTTP_500_INTERNAL_SERVER_ERROR)
            )

        return Response(task_result)


class ExportPPMMasterGridView(APIView):
    """
    ENDPOINT: Institutional cross-analysis export — a single Master Grid
    spreadsheet listing every PPM analysis for the requesting user's
    organization, one row per deal, for side-by-side portfolio comparison.

    Distinct from UnderwritePropertyDownloadView's per-analysis download:
    this is a many-rows-at-once export across an org's PPM history, not a
    single deal's underwriting workbook.

    SECURITY: organization is derived ONLY from get_user_organization(request)
    — the authenticated user's own organization — never from a client-
    supplied org_id or similar request parameter. Every other org-scoped
    view in this file (DocumentHistoryView, AnalysisReportListView,
    AnalysisReportDetailView, DashboardStatsView) follows this same rule; do
    not add an org_id parameter here, as that would let one organization
    pull another's PPM deal data.

    URL: GET /api/ai/ppm/export-master-grid/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request):
        from .models import AnalysisReport

        organization = get_user_organization(request)

        if organization is None:
            # Matches the read-view convention used elsewhere in this file
            # (DocumentHistoryView, AnalysisReportListView, DashboardStatsView)
            # — no organization means nothing to export, not an error.
            return Response({'message': 'No organization associated with this account.'}, status=status.HTTP_200_OK)

        if not organization.has_ppm_access:
            return Response(
                {'error': 'Institutional PPM analysis is not enabled for your organization.'},
                status=status.HTTP_403_FORBIDDEN
            )

        reports = AnalysisReport.objects.filter(
            organization=organization,
            document_type='PPM',
            is_deleted=False,
        ).order_by('-created_at')

        try:
            excel_bytes = build_ppm_master_grid(reports)
        except PPMExcelGenerationError as e:
            # Not really an error from the user's perspective — they simply
            # haven't analyzed any PPMs yet. 200 + a plain message, same
            # spirit as the empty-state responses elsewhere in this file.
            return Response({'message': str(e)}, status=status.HTTP_200_OK)
        except Exception as e:
            return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

        response = FileResponse(
            io.BytesIO(excel_bytes),
            as_attachment=True,
            filename='PPM_Master_Grid.xlsx',
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

        # Feature 2 (Immutable Snapshots): once this specific analysis has
        # been locked as a deal's permanent baseline (ExcelBaselineLock),
        # it can never be deleted — soft or otherwise — through this
        # endpoint. This is an explicit application-layer check, not just
        # the model's on_delete=PROTECT (which only guards a real DB-level
        # .delete(), and this view only ever sets is_deleted=True).
        if report.baseline_locks.exists():
            return Response(
                {'error': "This analysis is a locked deal baseline and can't be deleted."},
                status=status.HTTP_400_BAD_REQUEST
            )

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


# ---------------------------------------------------------------------------
# Feature 3 (Enterprise Pipeline & Workflow Dashboard): the Deal Pipeline.
# Every view below delegates "who can see/manage this deal" entirely to
# permissions.py (visible_projects_for / can_manage_project) — see that
# module's docstring for the exact platform-superuser / org-admin /
# workspace-member / own-projects-only rule. A project that isn't visible
# 404s rather than 403s — revealing "it exists but you can't see it" is
# itself information this multi-tenant platform shouldn't leak.
# ---------------------------------------------------------------------------

def _serialize_project_summary(project, user=None):
    from .permissions import can_lock_project

    latest = project.analyses.order_by('-created_at').first()
    # Feature 2 (Immutable Snapshots): OneToOneField reverse accessor —
    # Django raises RelatedObjectDoesNotExist (an AttributeError subclass,
    # specifically so getattr's default works) rather than returning None
    # when this project has never been locked.
    lock = getattr(project, 'baseline_lock', None)
    return {
        'id': project.id,
        'property_name': project.property_name or 'Untitled',
        'deal_status': project.deal_status,
        'deal_status_display': project.get_deal_status_display(),
        'organization': project.organization.name,
        'workspace': project.workspace.name if project.workspace_id else None,
        'uploaded_by': (
            latest.uploaded_by.email if latest and latest.uploaded_by
            else (project.created_by.email if project.created_by else None)
        ),
        'date_processed': latest.created_at.isoformat() if latest else project.created_at.isoformat(),
        'analysis_count': project.analyses.count(),
        'can_download': bool(latest and latest.rent_roll_data),
        'updated_at': project.updated_at.isoformat(),
        # Sprint "Workflow Upgrades" Feature 1: the analyst's own entered/
        # confirmed price for this deal — see PropertyProject.asking_price's
        # own docstring in models.py for how this differs from the AI-
        # extracted deal_valuation.target_purchase_price inside metrics.
        'asking_price': project.asking_price,
        # Feature 2 (Immutable Snapshots): `user` is optional (None) rather
        # than required, so every pre-existing call site that didn't yet
        # know about this feature still works unchanged — it just gets
        # can_lock=False, which is the safe default for an unknown caller.
        'is_locked': lock is not None,
        'locked_at': lock.locked_at.isoformat() if lock else None,
        'locked_by': lock.locked_by.email if lock and lock.locked_by else None,
        'can_lock': lock is None and user is not None and can_lock_project(user, project),
    }


class PropertyProjectListView(APIView):
    """
    The Deal Pipeline dashboard's table: property, uploader, date
    processed, status, and whether a historical Download is available.
    Optional ?deal_status=reviewing|approved|archived filter for the
    dashboard's status tabs/filter dropdown.
    URL: GET /api/ai/pipeline/projects/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request):
        from .permissions import visible_projects_for

        projects = visible_projects_for(request.user).select_related(
            'organization', 'workspace', 'created_by', 'baseline_lock', 'baseline_lock__locked_by'
        )

        deal_status_filter = request.query_params.get('deal_status')
        if deal_status_filter:
            projects = projects.filter(deal_status=deal_status_filter)

        return Response([_serialize_project_summary(p, user=request.user) for p in projects])


class PropertyProjectDetailView(APIView):
    """
    A single deal's full detail for the pipeline side-drawer — latest
    metrics, every past analysis run (each one individually downloadable
    when it has persisted rent_roll_data), and its Team Notes. PATCH
    updates deal_status only — notes have their own endpoint below.
    URL: GET/PATCH /api/ai/pipeline/projects/<int:project_id>/
    """
    permission_classes = [IsAuthenticated]

    def _get_visible_project_or_404(self, request, project_id):
        from .permissions import visible_projects_for
        project = visible_projects_for(request.user).select_related(
            'organization', 'workspace', 'created_by'
        ).filter(pk=project_id).first()
        if project is None:
            return None, Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)
        return project, None

    def get(self, request, project_id):
        project, error = self._get_visible_project_or_404(request, project_id)
        if error:
            return error

        analyses = list(project.analyses.select_related('uploaded_by').order_by('-created_at'))
        latest = analyses[0] if analyses else None
        notes = project.notes.select_related('author').order_by('-created_at')
        # Feature 2 (Immutable Snapshots) reverse accessor — see
        # _serialize_project_summary's own comment on why getattr(default)
        # is required here instead of a plain attribute access.
        lock = getattr(project, 'baseline_lock', None)

        return Response({
            **_serialize_project_summary(project, user=request.user),
            'latest_metrics': latest.metrics if latest else None,
            # Sprint "Workflow Upgrades" Feature 4 (Broker Offer Email):
            # the LOCKED analysis's metrics specifically — deliberately
            # NOT latest.metrics, which can move on past the lock via a
            # later re-analysis. Null whenever this deal isn't locked.
            'locked_metrics': lock.analysis.metrics if lock else None,
            # Sprint "Workflow Upgrades" Feature 2: "Seller vs. Underwritten"
            # table — see _compute_seller_vs_underwritten()'s own docstring.
            'seller_vs_underwritten': _compute_seller_vs_underwritten(project, latest),
            # Sprint "Workflow Upgrades" Feature 5: DSCR-constrained (and,
            # now that Feature 1 persists an asking_price, LTV-constrained
            # too) max loan sizing — see _compute_debt_sizing()'s docstring.
            'debt_sizing': _compute_debt_sizing(
                project.organization, latest.metrics if latest else None, asking_price=project.asking_price
            ),
            # Sprint "Workflow Upgrades" Feature 1: the 10-second Kill-or-
            # Proceed Scorecard, re-derived from the latest analysis's
            # PERSISTED data (not re-run against the original files) every
            # time this deal's pipeline drawer is reopened — see
            # _compute_deal_scorecard()'s own docstring. rent_roll_df is
            # rebuilt from AnalysisReport.rent_roll_data the same way the
            # Cell-Override flow above does (see _compute_seller_vs_underwritten's
            # neighboring helpers); None-safe when there's no analysis yet.
            'deal_scorecard': _compute_deal_scorecard(
                latest.metrics if latest else None,
                pd.DataFrame((latest.rent_roll_data or {}).get('rows', [])) if latest else None,
                latest.documents_provided if latest else None,
                project.asking_price,
            ),
            'analyses': [
                {
                    'id': a.id,
                    'created_at': a.created_at.isoformat(),
                    'uploaded_by': a.uploaded_by.email if a.uploaded_by else None,
                    'tier': a.tier,
                    'can_download': bool(a.rent_roll_data),
                }
                for a in analyses
            ],
            'notes': [
                {
                    'id': n.id,
                    'body': n.body,
                    'author': n.author.email if n.author else None,
                    'created_at': n.created_at.isoformat(),
                }
                for n in notes
            ],
        })

    def patch(self, request, project_id):
        from .permissions import can_manage_project
        from .models import PropertyProject

        project, error = self._get_visible_project_or_404(request, project_id)
        if error:
            return error
        if not can_manage_project(request.user, project):
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        new_status = request.data.get('deal_status')
        valid_statuses = dict(PropertyProject.STATUS_CHOICES)
        if new_status not in valid_statuses:
            return Response(
                {'detail': f"'deal_status' must be one of: {', '.join(valid_statuses)}."},
                status=status.HTTP_400_BAD_REQUEST
            )

        project.deal_status = new_status
        project.save(update_fields=['deal_status', 'updated_at'])

        return Response(_serialize_project_summary(project, user=request.user))


class ApproveLockBaselineView(APIView):
    """
    Feature 2 (Immutable Snapshots): "Approve & Lock Baseline" — freezes
    this deal's CURRENT analysis snapshot as its permanent underwriting
    baseline. See ExcelBaselineLock's own docstring for exactly what
    locking does, and deliberately does NOT do (deal_status and Team
    Notes stay fully editable either way).

    Gated to can_lock_project() — role='admin' (Org Admin) or
    workspace_role='head' (Workspace Head) only; see that function's
    docstring for why this is deliberately stricter than every other
    pipeline endpoint's can_manage_project. A plain analyst gets a 403
    even on a deal they created themselves.

    Requires the project's latest analysis to already have a captured
    injected_values_snapshot — i.e. its Excel has been downloaded at
    least once — there's nothing to freeze otherwise. Locking an
    already-locked project 400s rather than silently no-op'ing or
    re-locking to a newer analysis — this is a one-way action with no
    "re-lock" story in v1 (see the model's own docstring on why).
    URL: POST /api/ai/pipeline/projects/<int:project_id>/lock/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request, project_id):
        from django.db import transaction
        from .permissions import visible_projects_for, can_lock_project
        from .models import ExcelBaselineLock

        project = visible_projects_for(request.user).select_related('workspace').filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        if not can_lock_project(request.user, project):
            return Response(
                {'detail': (
                    "Only an Org Admin or this deal's Workspace Head can approve "
                    "and lock a baseline."
                )},
                status=status.HTTP_403_FORBIDDEN
            )

        existing_lock = getattr(project, 'baseline_lock', None)
        if existing_lock is not None:
            return Response(
                {'detail': f"This deal's baseline was already locked on {existing_lock.locked_at:%Y-%m-%d}."},
                status=status.HTTP_400_BAD_REQUEST
            )

        analysis = project.latest_analysis
        if analysis is None or not analysis.injected_values_snapshot:
            return Response(
                {'detail': (
                    "This deal has no downloaded Excel yet to lock as a baseline — "
                    "download it at least once, then lock it."
                )},
                status=status.HTTP_400_BAD_REQUEST
            )

        # Per the client's explicit instruction: hash ONLY the JSON
        # snapshot (never the .xlsx bytes, never the full metrics blob —
        # see ExcelBaselineLock's own docstring). sort_keys=True so the
        # hash is stable regardless of dict insertion/JSONField round-trip
        # ordering — the same logical snapshot must always hash identically.
        snapshot_json = json.dumps(analysis.injected_values_snapshot, sort_keys=True)
        snapshot_sha256 = hashlib.sha256(snapshot_json.encode('utf-8')).hexdigest()

        with transaction.atomic():
            lock = ExcelBaselineLock.objects.create(
                project=project,
                analysis=analysis,
                locked_by=request.user,
                snapshot_sha256=snapshot_sha256,
            )

        return Response({
            'id': lock.id,
            'locked_at': lock.locked_at.isoformat(),
            'locked_by': request.user.email,
            'analysis_id': analysis.id,
            'snapshot_sha256': lock.snapshot_sha256,
        }, status=status.HTTP_201_CREATED)


class TeamNoteListCreateView(APIView):
    """
    Team Notes on a deal — a flat, unthreaded running log (see TeamNote's
    own docstring for why; v1 scope, same reasoning as IssueReport). GET/POST
    both require the same visibility as the project itself — anyone who can
    see a deal can read and add to its notes.
    URL: GET/POST /api/ai/pipeline/projects/<int:project_id>/notes/
    """
    permission_classes = [IsAuthenticated]

    def _get_visible_project_or_404(self, request, project_id):
        from .permissions import visible_projects_for
        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return None, Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)
        return project, None

    def get(self, request, project_id):
        project, error = self._get_visible_project_or_404(request, project_id)
        if error:
            return error

        notes = project.notes.select_related('author').order_by('-created_at')
        return Response([
            {
                'id': n.id,
                'body': n.body,
                'author': n.author.email if n.author else None,
                'created_at': n.created_at.isoformat(),
            }
            for n in notes
        ])

    def post(self, request, project_id):
        from .models import TeamNote

        project, error = self._get_visible_project_or_404(request, project_id)
        if error:
            return error

        body = (request.data.get('body') or '').strip()
        if not body:
            return Response({'detail': 'Note text is required.'}, status=status.HTTP_400_BAD_REQUEST)

        note = TeamNote.objects.create(project=project, author=request.user, body=body)

        return Response({
            'id': note.id,
            'body': note.body,
            'author': note.author.email if note.author else None,
            'created_at': note.created_at.isoformat(),
        }, status=status.HTTP_201_CREATED)


class PropertyProjectDownloadView(APIView):
    """
    Regenerates an Excel workbook for a PropertyProject's LATEST analysis,
    from its persisted rent_roll_data/metrics — without needing the
    original source files again (which are never stored — see
    ProcessedDocument's docstring). Reuses the exact same template-choice/
    currency/debt-assumption branching as UnderwritePropertyDownloadView
    via _build_underwriting_excel_response, so the two paths can never
    drift from each other.

    Debt assumptions are re-prompted for here (parsed fresh from this
    request), never persisted on the analysis itself — see
    AnalysisReport.rent_roll_data's own docstring for why: they're scenario
    inputs the user may want to vary between downloads, not a fixed
    property of the analysis.
    URL: POST /api/ai/pipeline/projects/<int:project_id>/download/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request, project_id):
        from .permissions import visible_projects_for

        project = visible_projects_for(request.user).select_related('organization').filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        analysis = project.analyses.order_by('-created_at').first()
        if analysis is None:
            return Response({'detail': 'This deal has no analysis yet.'}, status=status.HTTP_404_NOT_FOUND)

        if not analysis.rent_roll_data:
            return Response(
                {'detail': 'Detailed download unavailable for legacy analyses processed prior to system upgrade.'},
                status=status.HTTP_400_BAD_REQUEST
            )

        debt_assumptions, debt_error = _parse_debt_assumptions(request)
        if debt_error:
            return debt_error

        currency_code, currency_error = _parse_currency_code(request)
        if currency_error:
            return currency_error

        rent_roll_df = pd.DataFrame((analysis.rent_roll_data or {}).get('rows', []))

        return _build_underwriting_excel_response(
            request, project.organization, analysis.metrics, rent_roll_df, debt_assumptions, currency_code,
            analysis=analysis,
        )


# ---------------------------------------------------------------------------
# Feature 4 (Cell Override Audit Log): "Re-upload Edited Underwriting File".
# Catches an analyst editing a figure locally in Excel AFTER downloading our
# generated workbook, before it goes any further (a lender, an IC memo).
# ---------------------------------------------------------------------------

def _normalize_value_for_compare(value):
    """
    Normalizes ONE value — either a value already stored in
    AnalysisReport.injected_values_snapshot (a plain JSON scalar) or a
    value just read back from a re-uploaded workbook's cell (which
    openpyxl may hand back as a datetime for a date-formatted cell, or a
    raw float/int/str) — into a single consistent string, for BOTH the
    change-detection comparison and the display value stored on
    ExcelOverrideDelta. Mirrors clean_cell_value's own date formatting
    (excel_generator.py) so a date stored as an ISO string and the SAME
    date read back from a live cell compare equal rather than false-positive
    as "changed" over formatting alone.
    """
    if value is None or value == "":
        return ""
    if isinstance(value, (pd.Timestamp, datetime.datetime)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        # Collapses float noise (5.0 vs 5.000000000001) without losing a
        # genuine change — 6dp is far finer than any figure here is ever
        # meaningfully precise to.
        return f"{round(value, 6):g}"
    return str(value).strip()


# Sprint "Workflow Upgrades" Feature 2 ("Seller vs. Underwritten" table):
# the 5 rows shown on the Deal View, in display order. "Vacancy /
# Concessions" uses IU_TotalAnnualConcessions as its tracked value — the
# only vacancy-loss-adjacent dollar figure Cell Override already tracks
# (there's no separately-tracked vacancy-loss-$ field, only the
# IU_PhysicalOccupancyPct/IU_EconomicOccupancyPct percentages, which aren't
# dollar figures this table can line up against GPR/EGI/OpEx/NOI).
SELLER_VS_UNDERWRITTEN_FIELDS = {
    'IU_AnnualizedGPR': 'GPR',
    'IU_TotalAnnualConcessions': 'Vacancy / Concessions',
    'IU_TotalOperatingRevenue': 'EGI',
    'IU_TotalOperatingExpenses': 'OpEx',
    'IU_NetOperatingIncome': 'NOI',
}


def _parse_numeric_display_value(display_value):
    """
    ExcelOverrideDelta.new_value is intentionally just a stringified
    display value (see that model's own docstring) — never a second copy
    of the raw number. For the 5 fields above, which are always currency
    figures, parse it back to a float for variance math; returns None for
    anything that isn't cleanly numeric (blank, a date, text) rather than
    guessing.
    """
    if display_value is None or display_value == "":
        return None
    try:
        return float(display_value)
    except (TypeError, ValueError):
        return None


def _compute_seller_vs_underwritten(project, analysis):
    """
    Sprint "Workflow Upgrades" Feature 2: Column A (Baseline Extracted) vs
    Column B (Latest Re-Uploaded/Post-Lock) vs Column C ($/% variance) for
    the 5 fields in SELLER_VS_UNDERWRITTEN_FIELDS.

    Baseline comes straight from analysis.injected_values_snapshot['scalars']
    — the values this system itself wrote the last time an Excel was
    generated, the same source of truth ExcelOverrideCheckView diffs
    against. "Latest" is the most recent ExcelOverrideDelta for that
    specific field across EVERY override upload ever submitted for this
    project/analysis — not just whichever upload is newest overall, since
    different fields can have last changed on different re-uploads —
    falling back to the baseline value itself for a field that was never
    overridden at all. Unlike _variance_pct() below (deliberately
    unsigned, since it only feeds a ">threshold" alert check), the $/%
    variance here is signed, since this table's whole point is showing
    which direction a number moved.
    """
    from .models import ExcelOverrideDelta

    if analysis is None or not analysis.injected_values_snapshot:
        return []

    baseline_scalars = analysis.injected_values_snapshot.get('scalars') or {}

    latest_deltas = {}
    deltas_qs = (
        ExcelOverrideDelta.objects
        .filter(
            upload__project=project,
            upload__analysis=analysis,
            row_index__isnull=True,
            field_name__in=SELLER_VS_UNDERWRITTEN_FIELDS,
        )
        .order_by('-upload__uploaded_at', '-id')
    )
    for delta in deltas_qs:
        latest_deltas.setdefault(delta.field_name, delta)

    rows = []
    for field_name, label in SELLER_VS_UNDERWRITTEN_FIELDS.items():
        baseline_value = baseline_scalars.get(field_name)
        latest_delta = latest_deltas.get(field_name)
        if latest_delta is not None:
            current_value = _parse_numeric_display_value(latest_delta.new_value)
        else:
            current_value = baseline_value

        dollar_variance = None
        pct_variance = None
        if (
            isinstance(baseline_value, (int, float)) and not isinstance(baseline_value, bool)
            and isinstance(current_value, (int, float)) and not isinstance(current_value, bool)
        ):
            dollar_variance = round(current_value - baseline_value, 2)
            if baseline_value != 0:
                pct_variance = round((current_value - baseline_value) / abs(baseline_value) * 100, 2)

        rows.append({
            'field_name': field_name,
            'label': label,
            'baseline_value': baseline_value,
            'current_value': current_value,
            'dollar_variance': dollar_variance,
            'pct_variance': pct_variance,
        })
    return rows


def _compute_deal_scorecard(metrics, rent_roll_df, documents_provided, asking_price):
    """
    Sprint "Workflow Upgrades" Feature 1 (10-Second "Kill or Proceed"
    Scorecard & Missing Doc Check): the quick read-out an analyst sees
    immediately after an Analyze run (or whenever a deal's pipeline drawer
    is reopened later) — which core documents/data points are missing, plus
    the handful of ratios that can usually be computed even from a partial
    submission.

    Called from two places with slightly different inputs:
    - Right after extraction (process_underwriting_files), with the
      freshly-built rent_roll_df and this run's own documents_provided.
    - Later, read-only, from PropertyProjectDetailView.get(), rebuilding
      rent_roll_df from the latest AnalysisReport's persisted rent_roll_data
      JSON (same pd.DataFrame(... .get('rows', [])) pattern used elsewhere
      in this file) and its persisted documents_provided/asking_price.

    Every figure here is computed in PYTHON from the underlying extracted
    numbers — never from an AI-self-reported ratio — same "don't trust a
    total the AI might also report" principle Module 1/Module 2 above
    already follow. Anything that can't be computed (a missing input) is
    None, never estimated.

    metrics/rent_roll_df/documents_provided may each be None/empty (e.g. a
    brand-new project with no analysis yet) — every branch below tolerates
    that and simply produces more Nones/flags rather than raising.
    """
    metrics = metrics or {}
    documents_provided = documents_provided or {}
    rent_roll_df = rent_roll_df if rent_roll_df is not None else pd.DataFrame()

    DOC_LABELS = (('om', 'Offering Memorandum (OM)'), ('t12', 'T12 (Trailing 12-Month Statement)'), ('rent_roll', 'Rent Roll'))
    # .get(key, True): an ABSENT key means "was provided" — correct for
    # every row that predates partial uploads (see AnalysisReport.
    # documents_provided's own docstring in models.py), not just a
    # coincidental default.
    missing_documents = [label for key, label in DOC_LABELS if not documents_provided.get(key, True)]

    data_quality_flags = []
    rent_roll_was_submitted = documents_provided.get('rent_roll', True)

    if rent_roll_was_submitted and not rent_roll_df.empty:
        try:
            detect_rent_column(rent_roll_df)
        except RentRollColumnError:
            data_quality_flags.append({
                'code': 'missing_rent_column',
                'message': "Rent roll has no identifiable in-place contract-rent column.",
                'email_bullet': "The in-place contract-rent column on the rent roll",
            })

        try:
            argus_fields = detect_argus_fields(rent_roll_df)
        except Exception:
            argus_fields = {}
        lease_end_col = argus_fields.get('lease_end_col')
        total_rows = len(rent_roll_df)
        if lease_end_col is None:
            data_quality_flags.append({
                'code': 'no_lease_end_column',
                'message': "Rent roll has no identifiable lease-expiration-date column.",
                'email_bullet': "Lease expiration dates on the rent roll",
            })
        elif total_rows > 0:
            missing_lease_end = int(rent_roll_df[lease_end_col].isna().sum())
            if missing_lease_end / total_rows > 0.20:
                pct = round(missing_lease_end / total_rows * 100)
                data_quality_flags.append({
                    'code': 'missing_lease_dates',
                    'message': (
                        f"{missing_lease_end} of {total_rows} rent roll rows ({pct}%) are "
                        f"missing a lease-expiration date."
                    ),
                    'email_bullet': "Lease expiration dates for the units missing them on the rent roll",
                })

    t12_section = metrics.get('t12_revenue_expenses') or {}
    noi = t12_section.get('net_operating_income')
    opex = t12_section.get('total_operating_expenses')
    egi = t12_section.get('total_operating_revenue')

    def _is_number(v):
        return isinstance(v, (int, float)) and not isinstance(v, bool)

    opex_ratio = round(opex / egi * 100, 2) if _is_number(opex) and _is_number(egi) and egi else None

    has_asking_price = _is_number(asking_price) and asking_price > 0
    in_place_cap_rate = round(noi / asking_price * 100, 2) if _is_number(noi) and has_asking_price else None

    property_metadata = metrics.get('property_metadata') or {}
    total_units = property_metadata.get('total_unit_count')
    total_sf = (
        (metrics.get('industrial_metrics') or {}).get('rentable_square_footage')
        or (metrics.get('office_metrics') or {}).get('building_total_rsf')
        or (metrics.get('retail_metrics') or {}).get('center_total_gla')
    )

    price_per_unit_or_sf = None
    price_per_unit_or_sf_basis = None
    if has_asking_price and _is_number(total_units) and total_units > 0:
        price_per_unit_or_sf = round(asking_price / total_units, 2)
        price_per_unit_or_sf_basis = 'unit'
    elif has_asking_price and _is_number(total_sf) and total_sf > 0:
        price_per_unit_or_sf = round(asking_price / total_sf, 2)
        price_per_unit_or_sf_basis = 'sf'

    return {
        'documents_provided': {
            'om': documents_provided.get('om', True),
            't12': documents_provided.get('t12', True),
            'rent_roll': documents_provided.get('rent_roll', True),
        },
        'missing_documents': missing_documents,
        'data_quality_flags': data_quality_flags,
        'asking_price': asking_price,
        'noi': noi,
        'in_place_cap_rate': in_place_cap_rate,
        'price_per_unit_or_sf': price_per_unit_or_sf,
        'price_per_unit_or_sf_basis': price_per_unit_or_sf_basis,
        'opex_ratio': opex_ratio,
        # Convenience for the frontend's "Copy Broker Request Email" button —
        # every missing-document label plus every data-quality flag's own
        # bullet text, already in one flat list. Empty means nothing's
        # missing, i.e. this deal is ready to move past the initial screen.
        'missing_items_for_email': missing_documents + [f['email_bullet'] for f in data_quality_flags],
    }


def _compute_debt_sizing(organization, metrics, asking_price=None):
    """
    Sprint "Workflow Upgrades" Feature 5 (Automatic Loan Sizing & Debt
    Service Check): standard mortgage-constant DSCR-based max loan sizing,
    using this org's debt_assumptions (Organization.get_debt_assumptions())
    and the deal's own extracted T12 NOI.

    max_loan_by_ltv / constrained_max_loan / estimated_dscr (the LTV-
    constrained half of the client's spec) need a persisted, analyst-
    entered asking_price on the deal record — Sprint "Workflow Upgrades"
    Feature 1 added PropertyProject.asking_price for exactly this. All
    three stay None (never guessed) whenever asking_price isn't available
    yet for this deal — e.g. an analysis run before Feature 1 shipped, or
    simply a deal nobody has entered a price for.

    Returns None when there's no usable NOI to size a loan against at all
    (no analysis yet, or NOI missing/zero/negative).
    """
    noi = (metrics or {}).get('t12_revenue_expenses', {}).get('net_operating_income')
    if not isinstance(noi, (int, float)) or isinstance(noi, bool) or noi <= 0:
        return None

    assumptions = organization.get_debt_assumptions()
    interest_rate_pct = assumptions['interest_rate']
    amortization_years = assumptions['amortization_years']
    min_dscr = assumptions['min_dscr']
    target_ltv_pct = assumptions['target_ltv']

    monthly_rate = (interest_rate_pct / 100.0) / 12.0
    n_payments = amortization_years * 12
    if monthly_rate == 0:
        # Degenerate but valid input (0% rate) — straight-line principal
        # only, no interest component to the mortgage constant.
        annual_debt_service_per_dollar = 1.0 / amortization_years
    else:
        monthly_constant = (
            (monthly_rate * (1 + monthly_rate) ** n_payments)
            / (((1 + monthly_rate) ** n_payments) - 1)
        )
        annual_debt_service_per_dollar = monthly_constant * 12

    max_loan_by_dscr = (noi / min_dscr) / annual_debt_service_per_dollar

    result = {
        'noi': noi,
        'min_dscr': min_dscr,
        'target_ltv': target_ltv_pct,
        'interest_rate': interest_rate_pct,
        'amortization_years': amortization_years,
        'annual_debt_service_per_dollar': round(annual_debt_service_per_dollar, 6),
        'max_loan_by_dscr': round(max_loan_by_dscr, 2),
        'max_loan_by_ltv': None,
        'constrained_max_loan': None,
        'estimated_dscr': None,
    }

    has_asking_price = (
        isinstance(asking_price, (int, float)) and not isinstance(asking_price, bool) and asking_price > 0
    )
    if has_asking_price:
        max_loan_by_ltv = asking_price * (target_ltv_pct / 100.0)
        constrained_max_loan = min(max_loan_by_dscr, max_loan_by_ltv)
        estimated_dscr = (
            round(noi / (constrained_max_loan * annual_debt_service_per_dollar), 2)
            if constrained_max_loan > 0 else None
        )
        result['max_loan_by_ltv'] = round(max_loan_by_ltv, 2)
        result['constrained_max_loan'] = round(constrained_max_loan, 2)
        result['estimated_dscr'] = estimated_dscr

    return result


def _variance_pct(original_value, current_value):
    """
    Feature 1 (Variance & Change Ledger) extension: percentage change
    between a tracked field's snapshotted value and its re-uploaded
    current value — fed the RAW values (before _normalize_value_for_compare
    stringifies them), so a real int/float is still recognizable as one.

    Returns None — never a fabricated number — for anything that isn't a
    genuine before/after percentage comparison: either side missing or not
    a real number, either side a bool (bool is technically an int subclass
    in Python, but "True -> False" is not a percentage change), or a zero
    original (a %-change FROM zero is undefined, not infinite-and-worth-
    flagging). This covers every text/date tracked field (IU_PropertyName,
    IU_Row_LeaseStart, etc.) automatically, with no field-name allowlist to
    maintain — they just never produce a variance_pct.
    """
    if isinstance(original_value, bool) or isinstance(current_value, bool):
        return None
    if not isinstance(original_value, (int, float)) or not isinstance(current_value, (int, float)):
        return None
    if original_value == 0:
        return None
    return round(abs(current_value - original_value) / abs(original_value) * 100, 2)


def _serialize_override_upload(upload):
    return {
        'id': upload.id,
        'uploaded_at': upload.uploaded_at.isoformat(),
        'uploaded_by': upload.uploaded_by.email if upload.uploaded_by else None,
        'original_filename': upload.original_filename,
        'is_post_lock_revision': upload.is_post_lock_revision,
        # Sprint "Workflow Upgrades" Feature 2: blank unless this upload's
        # deltas crossed the >2.0% GPR/OpEx/NOI threshold at the time it
        # was submitted — see ExcelOverrideCheckView.post()'s two-step flow.
        'variance_note': upload.variance_note,
        'deltas': [
            {
                'field_name': d.field_name,
                'row_index': d.row_index,
                'original_value': d.original_value,
                'new_value': d.new_value,
                'variance_pct': d.variance_pct,
            }
            for d in upload.deltas.all()
        ],
    }


class ExcelOverrideCheckView(APIView):
    """
    Parses the re-uploaded file's named ranges with the EXACT SAME
    extract_defined_names() the custom-template path already uses (it has
    no idea — and doesn't need to know — whether this file came from the
    custom-template path or the standard generator's own "Source Values
    (IU)" sheet/Argus tab, see excel_generator.py), reads each tracked
    field's CURRENT value, and diffs it against
    analysis.injected_values_snapshot — the values this system itself
    wrote the last time an Excel was actually generated for this deal.
    Every detected change is persisted as one ExcelOverrideUpload (the
    event) with one ExcelOverrideDelta per CHANGED field only — see that
    model's own docstring for why an unchanged field produces no row.

    Row fields are matched by POSITION ONLY, against exactly as many
    tenant rows as existed when the snapshot was captured — confirmed
    acceptable for v1 (position-based matching is standard practice in
    multifamily underwriting anyway; see ExcelOverrideDelta's docstring).
    A field whose named range can no longer be found at all in the
    re-upload (e.g. this platform's reference sheet/named ranges were
    deleted) is silently skipped, not reported as a change — there's
    nothing to compare it against.

    Requires only the same visibility as the project itself (like Team
    Notes below) — not can_manage_project — since re-uploading for a check
    is an audit action on the deal, not a change to deal_status itself.
    URL: POST /api/ai/pipeline/projects/<int:project_id>/override-check/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request, project_id):
        from django.db import transaction
        from .permissions import visible_projects_for
        from .models import ExcelOverrideUpload, ExcelOverrideDelta

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        analysis = project.analyses.order_by('-created_at').first()
        if analysis is None:
            return Response({'detail': 'This deal has no analysis yet.'}, status=status.HTTP_404_NOT_FOUND)

        snapshot = analysis.injected_values_snapshot
        if not snapshot:
            return Response(
                {'detail': (
                    "Override tracking isn't available yet for this deal — "
                    "download the Excel from this project at least once, "
                    "then re-uploading an edited copy will work."
                )},
                status=status.HTTP_400_BAD_REQUEST
            )

        uploaded_file = request.FILES.get('file')
        if not uploaded_file:
            return Response({'detail': 'Please upload the edited Excel file.'}, status=status.HTTP_400_BAD_REQUEST)

        try:
            validate_uploaded_file(uploaded_file, {'.xlsx', '.xlsm'}, label="Edited Underwriting File")
        except FileValidationError as e:
            return Response({'detail': str(e)}, status=status.HTTP_400_BAD_REQUEST)

        try:
            workbook = openpyxl.load_workbook(io.BytesIO(uploaded_file.read()), data_only=False)
        except Exception as e:
            return Response(
                {'detail': f"Couldn't open this file as an Excel workbook: {e}"},
                status=status.HTTP_400_BAD_REQUEST
            )

        resolved_names, _skipped = extract_defined_names(workbook)
        if not resolved_names:
            return Response(
                {'detail': (
                    "No tracked fields were found in this file — this "
                    "doesn't look like a workbook generated by this platform."
                )},
                status=status.HTTP_400_BAD_REQUEST
            )

        deltas = []

        # Scalars — one value each, direct compare.
        for name, original_value in (snapshot.get('scalars') or {}).items():
            if name not in resolved_names:
                continue  # named range no longer present — nothing to compare
            sheet_name, row, col = resolved_names[name]
            current_value = workbook[sheet_name].cell(row=row, column=col).value
            original_str = _normalize_value_for_compare(original_value)
            current_str = _normalize_value_for_compare(current_value)
            if original_str != current_str:
                deltas.append({
                    'field_name': name, 'row_index': None,
                    'original_value': original_str, 'new_value': current_str,
                    'variance_pct': _variance_pct(original_value, current_value),
                })

        # Rows — same anchor cell as a scalar, one row further down per
        # tenant position (the same "first data cell, write downward"
        # convention used everywhere these are written).
        snapshot_rows = snapshot.get('rows') or []
        if snapshot_rows:
            row_field_names = [n for n in snapshot_rows[0].keys() if n in resolved_names]
            for i, original_row in enumerate(snapshot_rows):
                for name in row_field_names:
                    sheet_name, anchor_row, col = resolved_names[name]
                    current_value = workbook[sheet_name].cell(row=anchor_row + i, column=col).value
                    original_raw = original_row.get(name)
                    original_str = _normalize_value_for_compare(original_raw)
                    current_str = _normalize_value_for_compare(current_value)
                    if original_str != current_str:
                        deltas.append({
                            'field_name': name, 'row_index': i,
                            'original_value': original_str, 'new_value': current_str,
                            'variance_pct': _variance_pct(original_raw, current_value),
                        })

        # Sprint "Workflow Upgrades" Feature 2 (Variance Alert Context
        # Notes): this re-upload requires a leadership-facing 1-sentence
        # reason BEFORE it commits if any detected delta touches one of the
        # same three fields/threshold already used by the firm-wide alert
        # feed below (OVERRIDE_ALERT_TRACKED_FIELDS /
        # OVERRIDE_ALERT_THRESHOLD_PCT) — intentionally the exact same
        # fields/threshold, not a separate configuration, so "what triggers
        # a leadership alert" and "what requires an explanation" never
        # drift apart. Nothing has been written to the database yet at
        # this point (deltas is still a plain list), so returning early
        # here needs no rollback — the frontend is expected to resubmit
        # the SAME file with variance_note set once the user supplies one.
        variance_note = (request.data.get('variance_note') or '').strip()
        note_required = any(
            d['field_name'] in OVERRIDE_ALERT_TRACKED_FIELDS
            and d['variance_pct'] is not None
            and d['variance_pct'] > OVERRIDE_ALERT_THRESHOLD_PCT
            for d in deltas
        )
        if note_required and not variance_note:
            return Response(
                {
                    'detail': (
                        "Key numbers shifted by more than 2%. Please add a "
                        "quick 1-sentence reason for leadership before this "
                        "upload is recorded."
                    ),
                    'requires_variance_note': True,
                },
                status=status.HTTP_400_BAD_REQUEST
            )

        with transaction.atomic():
            upload = ExcelOverrideUpload.objects.create(
                project=project,
                analysis=analysis,
                uploaded_by=request.user,
                original_filename=getattr(uploaded_file, 'name', '') or '',
                # Feature 2 (Immutable Snapshots) extension: labels this
                # upload as a post-lock "child version branch" when the
                # deal already has a baseline lock at the moment of this
                # re-upload — see ExcelOverrideUpload.is_post_lock_revision
                # and ExcelBaselineLock's own docstring. Checked via
                # hasattr rather than a query — baseline_lock is a
                # OneToOneField reverse accessor that raises
                # RelatedObjectDoesNotExist (not returning None) when unset.
                is_post_lock_revision=hasattr(project, 'baseline_lock'),
                variance_note=variance_note,
            )
            ExcelOverrideDelta.objects.bulk_create([
                ExcelOverrideDelta(
                    upload=upload,
                    field_name=d['field_name'],
                    row_index=d['row_index'],
                    original_value=d['original_value'],
                    new_value=d['new_value'],
                    variance_pct=d['variance_pct'],
                )
                for d in deltas
            ])

        return Response({'upload': _serialize_override_upload(upload)}, status=status.HTTP_201_CREATED)


class ExcelOverrideHistoryView(APIView):
    """
    Every past "Re-upload Edited Underwriting File" event on this deal,
    newest first — powers the pipeline drawer's "Override Summary" panel.
    URL: GET /api/ai/pipeline/projects/<int:project_id>/overrides/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, project_id):
        from .permissions import visible_projects_for
        from .models import ExcelOverrideUpload

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        uploads = ExcelOverrideUpload.objects.filter(project=project).select_related(
            'uploaded_by'
        ).prefetch_related('deltas').order_by('-uploaded_at')

        return Response({'uploads': [_serialize_override_upload(u) for u in uploads]})


# Feature 1 (Variance & Change Ledger) extension: per the client's explicit
# decision, the firm-wide PROACTIVE alert feed below is deliberately
# narrower than the full 18-scalar/13-row Cell Override Audit Log (every
# tracked field keeps being diffed and logged per-deal regardless — see
# ExcelOverrideCheckView above) — only these three fields, labeled for
# display, ever surface here.
OVERRIDE_ALERT_TRACKED_FIELDS = {
    'IU_AnnualizedGPR': 'Annualized GPR',
    'IU_TotalOperatingExpenses': 'Total Operating Expenses',
    'IU_NetOperatingIncome': 'Net Operating Income',
}
OVERRIDE_ALERT_THRESHOLD_PCT = 2.0


class OverrideAlertFeedView(APIView):
    """
    Feature 1 (Variance & Change Ledger) extension: a firm-wide, proactive
    feed of Cell Override changes that (a) touch GPR, OpEx, or NOI and (b)
    exceed OVERRIDE_ALERT_THRESHOLD_PCT — surfaced here so a Master/Head
    dashboard can show "something worth a second look happened somewhere
    in the pipeline" WITHOUT opening individual deals, exactly as scoped.

    Reuses visible_projects_for() verbatim rather than inventing a new
    permission function or a parallel isolation system — that's the
    client's own explicit decision on the related Feature 4 ("our existing
    Org -> Team Workspace -> Analyst/Head hierarchy already gives us the
    exact 'need-to-know' isolation we want"), and it applies just as
    cleanly here: for an org admin or workspace head, "firm-wide" already
    means every deal in their org/workspace through that same function;
    for a plain analyst it means only their own deals. No special case.

    Capped at the 200 most recent qualifying changes — a dashboard feed,
    not a full export (ExcelOverrideHistoryView above already gives the
    complete, uncapped per-deal history for anyone who opens a deal).
    URL: GET /api/ai/pipeline/override-alerts/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request):
        from .permissions import visible_projects_for
        from .models import ExcelOverrideDelta

        projects = visible_projects_for(request.user)

        deltas = (
            ExcelOverrideDelta.objects
            .filter(
                field_name__in=OVERRIDE_ALERT_TRACKED_FIELDS,
                variance_pct__isnull=False,
                variance_pct__gt=OVERRIDE_ALERT_THRESHOLD_PCT,
                upload__project__in=projects,
            )
            .select_related('upload', 'upload__uploaded_by', 'upload__project', 'upload__project__workspace')
            .order_by('-upload__uploaded_at')[:200]
        )

        alerts = [
            {
                'delta_id': d.id,
                'upload_id': d.upload_id,
                'uploaded_at': d.upload.uploaded_at.isoformat(),
                'uploaded_by': d.upload.uploaded_by.email if d.upload.uploaded_by else None,
                'project_id': d.upload.project_id,
                'property_name': d.upload.project.property_name,
                'workspace_name': d.upload.project.workspace.name if d.upload.project.workspace_id else None,
                'field_name': d.field_name,
                'field_label': OVERRIDE_ALERT_TRACKED_FIELDS[d.field_name],
                'original_value': d.original_value,
                'new_value': d.new_value,
                'variance_pct': d.variance_pct,
                # Sprint "Workflow Upgrades" Feature 2: the analyst's
                # required 1-sentence explanation for this upload, shown
                # directly under the alert row in the firm-wide feed.
                'variance_note': d.upload.variance_note,
            }
            for d in deltas
        ]

        return Response({'alerts': alerts, 'threshold_pct': OVERRIDE_ALERT_THRESHOLD_PCT})