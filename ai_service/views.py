# backend/ai_service/views.py
import io
import re
import json
import hashlib
import os
import time
import uuid
import base64
import zipfile
import tempfile
import datetime
import openpyxl
import pandas as pd
from datetime import timedelta
from email.utils import parseaddr
from django.http import FileResponse
from django.conf import settings as django_settings
from django.utils import timezone
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from rest_framework.permissions import IsAuthenticated, AllowAny

from celery.result import AsyncResult
from .tasks import run_cre_underwriting_task, run_ppm_underwriting_task, run_inbound_email_underwriting_task

from .llm_router import ai_client, AIRateLimitExhausted
from .parsers import (
    extract_document_text,
    extract_t12_text,
    extract_data_from_excel,
    read_excel_with_header_detection,
    validate_uploaded_file,
    FileValidationError,
)
from .excel_generator import generate_underwriting_excel, ExcelGenerationError, CURRENCY_FORMATS, DEFAULT_CURRENCY_CODE, safe_get
from .waterfall_calculator import compute_waterfall, WaterfallInputError
from .custom_template_generator import (
    generate_excel_from_template,
    TemplateValidationError as CustomTemplateValidationError,
    TemplateCapacityError,
    compute_injected_values_snapshot,
    extract_defined_names,
    SCALAR_FIELD_CATALOG,
    ROW_FIELD_CATALOG,
    apply_scalar_edits_to_metrics,
    apply_row_edits_to_rent_roll_data,
)
from openpyxl.workbook.defined_name import DefinedName
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
    DOCUMENT_TYPE_RENT_ROLL,
    DOCUMENT_TYPE_TRAILING_12,
    DOCUMENT_TYPE_OFFERING_MEMORANDUM,
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
    from .permissions import can_lock_project, can_manage_project

    latest = project.analyses.order_by('-created_at').first()
    # Feature 2 (Immutable Snapshots): OneToOneField reverse accessor —
    # Django raises RelatedObjectDoesNotExist (an AttributeError subclass,
    # specifically so getattr's default works) rather than returning None
    # when this project has never been locked.
    lock = getattr(project, 'baseline_lock', None)
    return {
        'id': project.id,
        'property_name': project.property_name or 'Untitled',
        'pipeline_stage': project.pipeline_stage,
        'pipeline_stage_display': project.get_pipeline_stage_display(),
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
        # Pipeline Modules Phase 1: who's actively working this deal, and
        # the team's own target date for closing/offer — see
        # PropertyProject.assigned_analyst/target_closing_or_offer_date's
        # docstrings in models.py. Both null until someone sets them.
        'assigned_analyst': project.assigned_analyst.email if project.assigned_analyst_id else None,
        'target_closing_or_offer_date': (
            project.target_closing_or_offer_date.isoformat()
            if project.target_closing_or_offer_date else None
        ),
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
        # Pipeline Modules Phase 2 (VP Sign-Off Gate): deliberately NOT
        # gated on `lock is None` like can_lock above — a VP/Org Admin can
        # still sign off and advance an already-locked 'vp_review' deal
        # (see VPSignOffView's docstring), so the frontend needs this
        # permission flag available independent of lock state.
        'can_sign_off': user is not None and can_lock_project(user, project),
        # Pipeline Modules Phase 6: gates the drawer's assigned_analyst/
        # target_closing_or_offer_date edit controls and the "Generate
        # Lender Packet" action — same can_manage_project check the PATCH
        # endpoint and the lender-link endpoints enforce server-side, just
        # exposed here so the frontend can show/hide the controls instead
        # of letting a user try an edit that's just going to 403.
        'can_manage': user is not None and can_manage_project(user, project),
    }


class PropertyProjectListView(APIView):
    """
    The Deal Pipeline dashboard's table: property, uploader, date
    processed, stage, and whether a historical Download is available.
    Optional ?pipeline_stage=new_incoming|underwriting|vp_review|
    offer_submitted|dead_passed filter for the dashboard's stage
    tabs/filter dropdown.
    URL: GET /api/ai/pipeline/projects/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request):
        from .permissions import visible_projects_for

        projects = visible_projects_for(request.user).select_related(
            'organization', 'workspace', 'created_by', 'assigned_analyst',
            'baseline_lock', 'baseline_lock__locked_by'
        )

        pipeline_stage_filter = request.query_params.get('pipeline_stage')
        if pipeline_stage_filter:
            projects = projects.filter(pipeline_stage=pipeline_stage_filter)

        return Response([_serialize_project_summary(p, user=request.user) for p in projects])


def _get_deal_activity_feed(project):
    """
    Pipeline Modules Phase 4 (Unified Activity Feed): ONE chronological
    feed for the pipeline drawer, synthesized server-side from THREE
    already-existing sources rather than a new table of its own —
    consistent with this module's "don't duplicate audit infrastructure"
    discipline (see TeamNote/DealActivityLog/ExcelOverrideUpload's own
    docstrings for earlier examples of the same principle). Per the
    client's explicit decision, Team Notes are no longer shown in their
    own separate drawer section — they're interleaved here as "comment"
    events alongside stage changes and checklist resolutions, newest
    first, matching every other list in this module.

      - DealActivityLog: every pipeline_stage transition (Phase 1),
        including the from_stage=None "deal created" row.
      - TeamNote: every comment, each carrying its resolved @mentions
        (Phase 4) — see _parse_mentions above.
      - DealChecklistItem: only RESOLUTION events, keyed off resolved_at
        — NOT creation events. An item being auto-seeded as missing
        isn't itself something that happened ON the deal worth a feed
        entry (every deal would otherwise get noisy "checklist item
        created" spam the moment it's analyzed); a human actually
        resolving one is a real, feed-worthy event.
      - SyncedEmailMessage (V2 Module 1, Phase V2-3): every message the
        5-minute poll (or the one-time backfill right after a thread is
        tagged) has pulled from any thread EVER linked to this deal —
        deliberately including messages synced through a thread that's
        since been soft-unlinked (LinkedEmailThread.is_active=False
        only stops FUTURE polling; it was never meant to erase history
        that's already in the feed — see that model's own docstring).

    Each entry carries a 'type' discriminator the frontend switches on
    to render it, and an 'at' ISO timestamp (already sorted on before
    being stringified, so the feed array itself is the final order).
    """
    from .models import PropertyProject

    stage_labels = dict(PropertyProject.PIPELINE_STAGE_CHOICES)
    events = []

    for log in project.activity_log.select_related('user'):
        events.append({
            'type': 'stage_change',
            'at': log.created_at,
            'id': f'stage-{log.id}',
            'user': log.user.email if log.user else None,
            'from_stage': log.from_stage,
            'from_stage_display': stage_labels.get(log.from_stage) if log.from_stage else None,
            'to_stage': log.to_stage,
            'to_stage_display': stage_labels.get(log.to_stage, log.to_stage),
        })

    for note in project.notes.select_related('author').prefetch_related('mentioned_users'):
        events.append({
            'type': 'comment',
            'at': note.created_at,
            'id': f'note-{note.id}',
            'user': note.author.email if note.author else None,
            'body': note.body,
            'mentioned_users': [u.email for u in note.mentioned_users.all()],
        })

    for item in project.checklist_items.filter(
        is_resolved=True, resolved_at__isnull=False
    ).select_related('resolved_by'):
        events.append({
            'type': 'checklist_resolved',
            'at': item.resolved_at,
            'id': f'checklist-{item.id}',
            'user': item.resolved_by.email if item.resolved_by else None,
            'label': item.label,
        })

    # Pipeline Modules Phase 6 (Lender Packet & View Tracking): ONE feed
    # entry per LenderShareLink that's been opened at least once —
    # LenderShareLink.view_count/last_viewed_at is a running counter, not
    # a per-visit log, so this is "viewed N times, most recently at..."
    # rather than N separate events. Deliberately NOT sourced from
    # DealActivityLog (that model is stage-transitions-only — see its own
    # docstring); the view count/timestamp live on LenderShareLink itself.
    for link in project.lender_share_links.filter(view_count__gt=0, last_viewed_at__isnull=False):
        events.append({
            'type': 'lender_link_viewed',
            'at': link.last_viewed_at,
            'id': f'lender-link-{link.id}',
            'user': None,
            'recipient_label': link.recipient_label or 'Lender',
            'view_count': link.view_count,
        })

    from .models import SyncedEmailMessage

    for msg in SyncedEmailMessage.objects.filter(thread_link__project=project).select_related('thread_link', 'thread_link__connection'):
        thread_link = msg.thread_link
        events.append({
            'type': 'email',
            'at': msg.sent_at,
            'id': f'email-{msg.id}',
            'user': None,
            'direction': msg.direction,
            'sender_email': msg.sender_email,
            'sender_name': msg.sender_name,
            'thread_label': thread_link.label or thread_link.subject or '(no subject)',
            'thread_link_id': thread_link.id,
            'provider': thread_link.connection.provider,
            'body_text': msg.body_text,
            'attachment_summary': msg.attachment_summary,
        })

    events.sort(key=lambda e: e['at'], reverse=True)
    for e in events:
        e['at'] = e['at'].isoformat()
    return events


class PropertyProjectDetailView(APIView):
    """
    A single deal's full detail for the pipeline side-drawer — latest
    metrics, every past analysis run (each one individually downloadable
    when it has persisted rent_roll_data), and its Team Notes. PATCH
    updates any of pipeline_stage / assigned_analyst /
    target_closing_or_offer_date, independently — notes have their own
    endpoint below. See the patch() method's own comments for exactly
    what each does.
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
            # Pipeline Modules Phase 4 (Unified Activity Feed): supersedes
            # the old standalone 'notes' list in this payload — Team
            # Notes are now interleaved here as "comment" events
            # alongside stage changes and checklist resolutions. See
            # _get_deal_activity_feed()'s own docstring. Posting a note
            # still goes through TeamNoteListCreateView's own endpoint
            # unchanged; only this read-side shape changed.
            'activity_feed': _get_deal_activity_feed(project),
            # Pipeline Modules Phase 4 (Persistent Missing-Doc Checklist):
            # inlined here (not just its own /checklist/ endpoint) so the
            # drawer has it the instant it opens, same reasoning as
            # 'analyses' above. Resolving/adding an item still goes
            # through DealChecklistListCreateView/
            # DealChecklistItemResolveView's own endpoints.
            'checklist_items': [
                _serialize_checklist_item(i)
                for i in project.checklist_items.select_related('created_by', 'resolved_by')
            ],
        })

    def patch(self, request, project_id):
        from .permissions import can_manage_project
        from .models import PropertyProject, DealActivityLog

        project, error = self._get_visible_project_or_404(request, project_id)
        if error:
            return error
        if not can_manage_project(request.user, project):
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        update_fields = []

        # Pipeline Modules Phase 1: a stage change is the one field here
        # that also writes a DealActivityLog row — see that model's
        # docstring. Key is optional on this PATCH; omit it entirely to
        # update only assigned_analyst/target_closing_or_offer_date below
        # without touching the deal's stage.
        if 'pipeline_stage' in request.data:
            new_stage = request.data.get('pipeline_stage')
            valid_stages = dict(PropertyProject.PIPELINE_STAGE_CHOICES)
            if new_stage not in valid_stages:
                return Response(
                    {'detail': f"'pipeline_stage' must be one of: {', '.join(valid_stages)}."},
                    status=status.HTTP_400_BAD_REQUEST
                )
            if new_stage != project.pipeline_stage:
                # Pipeline Modules Phase 2 (VP Sign-Off Gate): moving a deal
                # INTO 'offer_submitted' requires its baseline to already be
                # locked — see VPSignOffView's own docstring for why
                # "locked" doubles as "a VP has signed off" (the client's
                # explicit combined-action decision). This only blocks this
                # generic dropdown path; VPSignOffView itself performs the
                # lock (if needed) and the stage change together in one
                # action. Once a deal IS locked, pipeline_stage goes back to
                # being as freely editable as every other stage — see
                # ExcelBaselineLock's own docstring ("pipeline_stage ...
                # left fully editable").
                if new_stage == 'offer_submitted' and getattr(project, 'baseline_lock', None) is None:
                    return Response(
                        {'detail': (
                            "This deal needs VP Sign-Off before it can move to "
                            "'Offer Submitted' — use the Sign Off action once "
                            "it's in VP Review."
                        )},
                        status=status.HTTP_400_BAD_REQUEST
                    )
                DealActivityLog.objects.create(
                    project=project,
                    user=request.user,
                    from_stage=project.pipeline_stage,
                    to_stage=new_stage,
                )
                project.pipeline_stage = new_stage
                update_fields.append('pipeline_stage')

        # Pipeline Modules Phase 1: who's actively working this deal day
        # to day. An empty string clears the assignment; otherwise the
        # email must belong to THIS project's own organization — never
        # trust client input to already be scoped right, same rule
        # applied everywhere else in this module.
        if 'assigned_analyst' in request.data:
            raw_email = (request.data.get('assigned_analyst') or '').strip()
            if not raw_email:
                project.assigned_analyst = None
            else:
                from users.models import CustomUser
                analyst = CustomUser.objects.filter(
                    email=raw_email, organization_id=project.organization_id
                ).first()
                if analyst is None:
                    return Response(
                        {'detail': "'assigned_analyst' must be a member of this deal's organization."},
                        status=status.HTTP_400_BAD_REQUEST
                    )
                project.assigned_analyst = analyst
            update_fields.append('assigned_analyst')

        # Pipeline Modules Phase 1: the team's own target date — a
        # planning input, never validated against anything else on the
        # deal (e.g. not required to be in the future; teams backdate
        # these sometimes when entering history).
        if 'target_closing_or_offer_date' in request.data:
            raw_date = request.data.get('target_closing_or_offer_date')
            if not raw_date:
                project.target_closing_or_offer_date = None
            else:
                from datetime import date as _date
                try:
                    project.target_closing_or_offer_date = _date.fromisoformat(raw_date)
                except (TypeError, ValueError):
                    return Response(
                        {'detail': "'target_closing_or_offer_date' must be an ISO date (YYYY-MM-DD)."},
                        status=status.HTTP_400_BAD_REQUEST
                    )
            update_fields.append('target_closing_or_offer_date')

        if not update_fields:
            return Response(
                {'detail': (
                    'Provide at least one of: pipeline_stage, assigned_analyst, '
                    'target_closing_or_offer_date.'
                )},
                status=status.HTTP_400_BAD_REQUEST
            )

        project.save(update_fields=update_fields + ['updated_at'])

        return Response(_serialize_project_summary(project, user=request.user))


class DealAssignableAnalystsView(APIView):
    """
    Pipeline Modules Phase 6 add-on: the roster the deal drawer's
    assigned_analyst picker selects from — the same org-membership data
    MyOrganizationMembersView already serves (users/views.py), but gated
    to can_manage_project(user, THIS project) instead of role=='admin',
    so a plain analyst managing their own deal isn't 403'd just because
    they aren't an Org Admin (can_manage_project already allows exactly
    this case via its own-created-deal fallback — see permissions.py).

    Deliberately NOT a general-purpose member directory: scoped to one
    project's organization, and requires the caller already be allowed
    to manage THAT project — MyOrganizationMembersView's admin-only
    "manage my team's roster" scope is left completely untouched.
    URL: GET /api/ai/pipeline/projects/<int:project_id>/assignable-analysts/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, project_id):
        from .permissions import visible_projects_for, can_manage_project
        from users.models import CustomUser

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)
        if not can_manage_project(request.user, project):
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        members = CustomUser.objects.filter(
            organization_id=project.organization_id, is_active=True
        ).order_by('first_name', 'last_name')

        return Response([
            {
                'email': m.email,
                'first_name': m.first_name,
                'last_name': m.last_name,
            }
            for m in members
        ])


def _vault_locked_baseline_excel(project, lock, user):
    """
    Pipeline Modules Phase 5 (Excel Versioning & Supporting Document
    Vault): auto-captures a DealVaultFile the instant a baseline is
    locked — see DealVaultFile's own docstring in models.py (file_type=
    'locked_baseline_excel') for exactly what this does and its
    documented limitation (regenerated from this org's CURRENT debt
    assumptions, since purchase price/LTV/rate are never persisted on an
    analysis — see AnalysisReport.rent_roll_data's own docstring — so
    this vaulted copy isn't guaranteed byte-identical to any one
    historical download).

    The ONE caller (_create_baseline_lock below) wraps this entire call
    in try/except — a failure here must NEVER block the lock itself,
    which is the actual point of the action the user asked for. This
    function itself raises normally on failure; it's the caller's job to
    swallow it.

    Deliberately skips organizations using a ClientExcelTemplate (custom
    Excel template) in this v1 — see DealVaultFile's own docstring for
    why (avoids custom_template_generator.py's more complex branching).
    """
    from .models import ClientExcelTemplate, DealVaultFile
    from django.core.files.base import ContentFile

    if ClientExcelTemplate.objects.filter(organization=project.organization_id, is_active=True).exists():
        logger.info(
            "Skipping auto-vault of locked baseline Excel for project %s — "
            "org %s uses a custom Excel template (v1 scope excludes "
            "custom-template orgs from auto-vault-on-lock).",
            project.id, project.organization_id,
        )
        return

    analysis = lock.analysis
    rent_roll_df = pd.DataFrame((analysis.rent_roll_data or {}).get('rows', []))
    assumptions = project.organization.get_debt_assumptions()
    debt_assumptions = {
        'ltv_pct': assumptions['target_ltv'],
        'interest_rate_pct': assumptions['interest_rate'],
        'amortization_years': assumptions['amortization_years'],
        # project.asking_price is the best available stand-in — debt
        # inputs are re-prompted at each download, never persisted (see
        # above). generate_underwriting_excel tolerates a missing
        # purchase_price (debt_assumptions.get('purchase_price'), no
        # default — see excel_generator.py) by simply leaving that one
        # cell blank, so None here is safe, not a crash.
        'purchase_price': project.asking_price,
    }

    excel_bytes = generate_underwriting_excel(
        analysis.metrics, rent_roll_df, debt_assumptions=debt_assumptions, currency_code=DEFAULT_CURRENCY_CODE,
    )

    DealVaultFile.objects.create(
        project=project,
        file=ContentFile(excel_bytes, name=f"locked_baseline_{project.id}.xlsx"),
        file_type='locked_baseline_excel',
        original_filename=f"{(project.property_name or 'Deal').replace(' ', '_')}_Locked_Baseline.xlsx",
        version_label=f"Locked {lock.locked_at:%Y-%m-%d}",
        uploaded_by=user,
    )


def _create_baseline_lock(project, user):
    """
    Pipeline Modules Phase 2: the actual lock-creation logic shared by
    ApproveLockBaselineView (the standalone "Approve & Lock Baseline"
    action) and VPSignOffView (which locks, if not already locked, as
    HALF of its combined sign-off action — see that view's docstring).
    Factored out so there is exactly one place that computes the
    snapshot hash and creates the row, rather than two copies that could
    silently drift apart.

    Callers are responsible for confirming can_lock_project(user, project)
    and that project.baseline_lock doesn't already exist — this helper
    only handles the one remaining precondition (a downloaded snapshot to
    freeze) and the creation itself. Returns (lock, error_response) —
    exactly one of the two is ever non-None.
    """
    from .models import ExcelBaselineLock

    analysis = project.latest_analysis
    if analysis is None or not analysis.injected_values_snapshot:
        return None, Response(
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

    lock = ExcelBaselineLock.objects.create(
        project=project,
        analysis=analysis,
        locked_by=user,
        snapshot_sha256=snapshot_sha256,
    )

    # Pipeline Modules Phase 5: best-effort, never blocks the lock itself
    # — see _vault_locked_baseline_excel's own docstring.
    try:
        _vault_locked_baseline_excel(project, lock, user)
    except Exception:
        logger.exception(
            "Failed to auto-vault locked baseline Excel for project %s — "
            "the lock itself still succeeded.", project.id,
        )

    # V2 Module 3: freeze whatever Scenario Builder results exist at
    # lock time onto the lock itself — same best-effort, never-blocks
    # posture as vaulting above. Left as None (not recomputed) if the
    # analyst never ran "Calculate" before locking — never fabricated.
    try:
        saved_scenario = analysis.waterfall_scenario
        if saved_scenario and saved_scenario.get('results'):
            lock.waterfall_summary = saved_scenario['results']
            lock.save(update_fields=['waterfall_summary'])
    except Exception:
        logger.exception(
            "Failed to freeze waterfall scenario summary for project %s — "
            "the lock itself still succeeded.", project.id,
        )

    return lock, None


class ApproveLockBaselineView(APIView):
    """
    Feature 2 (Immutable Snapshots): "Approve & Lock Baseline" — freezes
    this deal's CURRENT analysis snapshot as its permanent underwriting
    baseline. See ExcelBaselineLock's own docstring for exactly what
    locking does, and deliberately does NOT do (pipeline_stage and Team
    Notes stay fully editable either way).

    Gated to can_lock_project() — role='admin' (Org Admin) or
    workspace_role='head' (Workspace Head) only; see that function's
    docstring for why this is deliberately stricter than every other
    pipeline endpoint's can_manage_project. A plain analyst gets a 403
    even on a deal they created themselves.

    Requires the project's latest analysis to already have a captured
    injected_values_snapshot — i.e. its Excel has been downloaded at
    least once — there's nothing to freeze otherwise (see
    _create_baseline_lock, shared with VPSignOffView below). Locking an
    already-locked project 400s rather than silently no-op'ing or
    re-locking to a newer analysis — this is a one-way action with no
    "re-lock" story in v1 (see the model's own docstring on why).
    URL: POST /api/ai/pipeline/projects/<int:project_id>/lock/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request, project_id):
        from django.db import transaction
        from .permissions import visible_projects_for, can_lock_project

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

        with transaction.atomic():
            lock, error = _create_baseline_lock(project, request.user)
            if error:
                return error

        return Response({
            'id': lock.id,
            'locked_at': lock.locked_at.isoformat(),
            'locked_by': request.user.email,
            'analysis_id': lock.analysis_id,
            'snapshot_sha256': lock.snapshot_sha256,
        }, status=status.HTTP_201_CREATED)


class VPSignOffView(APIView):
    """
    Pipeline Modules Phase 2 (VP Sign-Off Gate & Leadership Digest):
    "Sign Off & Advance to Offer Submitted" — the one combined action
    that both satisfies the VP Sign-Off Gate (see
    PropertyProjectDetailView.patch's own gate check on the
    'offer_submitted' transition) and freezes the deal's financial
    baseline, per the client's explicit decision that these are a single
    action rather than two separate steps: by the time an offer is going
    out the door, the numbers backing it should already be frozen.

    Deliberately reuses can_lock_project (NOT can_manage_project) — the
    same Org Admin / Workspace Head-only gate as the existing "Approve &
    Lock Baseline" action, since this view performs that exact same
    freeze (via the shared _create_baseline_lock() helper above) as part
    of signing off. A plain analyst gets a 403 even on a deal they
    created and are assigned to.

    Requires the deal to currently be in the 'vp_review' stage — signing
    off a deal that hasn't reached VP Review yet (or has already moved
    past it) is rejected with a 400 rather than silently fast-forwarding
    it through stages it was never actually reviewed at.

    If the deal's baseline is ALREADY locked (e.g. an earlier, unrelated
    lock from before this deal ever reached VP Review), this does NOT
    error or attempt to re-lock — it treats that existing lock as
    already satisfying the freeze half of sign-off and proceeds straight
    to advancing the stage. Either way, the whole operation (lock-if-
    needed + stage change + DealActivityLog row) is one atomic
    transaction, so a deal can never end up "signed off" with only half
    of this applied.
    URL: POST /api/ai/pipeline/projects/<int:project_id>/vp-sign-off/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request, project_id):
        from django.db import transaction
        from .permissions import visible_projects_for, can_lock_project
        from .models import DealActivityLog

        project = visible_projects_for(request.user).select_related('workspace').filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        if not can_lock_project(request.user, project):
            return Response(
                {'detail': "Only an Org Admin or this deal's Workspace Head can sign off on a deal."},
                status=status.HTTP_403_FORBIDDEN
            )

        if project.pipeline_stage != 'vp_review':
            return Response(
                {'detail': (
                    f"This deal is in '{project.get_pipeline_stage_display()}', not "
                    "'VP Review' — move it to VP Review before it can be signed off."
                )},
                status=status.HTTP_400_BAD_REQUEST
            )

        existing_lock = getattr(project, 'baseline_lock', None)

        with transaction.atomic():
            if existing_lock is None:
                lock, error = _create_baseline_lock(project, request.user)
                if error:
                    return error

            DealActivityLog.objects.create(
                project=project,
                user=request.user,
                from_stage=project.pipeline_stage,
                to_stage='offer_submitted',
            )
            project.pipeline_stage = 'offer_submitted'
            project.save(update_fields=['pipeline_stage', 'updated_at'])

        # Belt-and-suspenders: refresh_from_db() clears any cached relations
        # on this instance (including the baseline_lock reverse one-to-one),
        # so _serialize_project_summary below is guaranteed a fresh read
        # rather than depending on exactly how/whether a prior getattr()
        # cached its result.
        project.refresh_from_db()
        return Response(_serialize_project_summary(project, user=request.user))


class ICSummaryPDFView(APIView):
    """
    Pipeline Modules Phase 2 follow-up: "Download IC Summary" in the deal
    drawer — a per-deal, one-page reportlab PDF (see ic_summary_pdf.py)
    handing a VP/analyst something easy to pass to an Investment
    Committee right after signing off on a deal.

    Built exclusively from the LOCKED analysis (ExcelBaselineLock.analysis)
    — never the project's current/latest one — same "frozen, not a moving
    target" principle as Draft Broker Email's use of locked_metrics and
    ApproveLockBaselineView itself. Requires the deal's baseline to
    already be locked; there's nothing frozen to summarize otherwise, so
    this 400s with a message pointing at Approve & Lock Baseline / VP
    Sign-Off rather than silently falling back to an unlocked snapshot.

    Deliberately just visible_projects_for() — the same visibility as the
    rest of the deal drawer — rather than can_lock_project. This is a
    read/export action on a deal you can already see, not a workflow
    action like VPSignOffView above, so it doesn't need the stricter gate.
    URL: GET /api/ai/pipeline/projects/<int:project_id>/ic-summary/pdf/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, project_id):
        from .permissions import visible_projects_for
        from .ic_summary_pdf import render_ic_summary_pdf

        project = visible_projects_for(request.user).select_related(
            'organization', 'workspace', 'assigned_analyst',
            'baseline_lock', 'baseline_lock__locked_by', 'baseline_lock__analysis',
        ).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        lock = getattr(project, 'baseline_lock', None)
        if lock is None:
            return Response(
                {'detail': (
                    "This deal's baseline must be locked before an IC Summary can "
                    "be generated — use Approve & Lock Baseline, or Sign Off if "
                    "it's in VP Review."
                )},
                status=status.HTTP_400_BAD_REQUEST
            )

        analysis = lock.analysis
        rent_roll_df = (
            pd.DataFrame((analysis.rent_roll_data or {}).get('rows', []))
            if analysis.rent_roll_data else None
        )
        deal_scorecard = _compute_deal_scorecard(
            analysis.metrics, rent_roll_df, analysis.documents_provided, project.asking_price
        )
        debt_sizing = _compute_debt_sizing(
            project.organization, analysis.metrics, asking_price=project.asking_price
        )

        pdf_bytes = render_ic_summary_pdf(
            project, lock, analysis.metrics, deal_scorecard, debt_sizing, lock.waterfall_summary,
        )

        safe_name = (project.property_name or 'Deal').replace('/', '-').replace('\\', '-')
        filename = f"IC_Summary_{safe_name}.pdf"
        return FileResponse(
            io.BytesIO(pdf_bytes),
            as_attachment=True,
            filename=filename,
            content_type='application/pdf',
        )


class PresentDealDataView(APIView):
    """
    V2 Module 4c (Boardroom IC Presentation Mode): the JSON data behind
    the deal drawer's "Present Deal" button — a full-screen, multi-slide
    view the frontend renders (Property Overview, Seller vs.
    Underwritten, Forensic Audit Flags, Debt & Waterfall Summary, Local
    Comps). Property Overview/Debt Sizing/Waterfall are computed
    IDENTICALLY to ICSummaryPDFView just above (same
    _compute_deal_scorecard/_compute_debt_sizing calls, same
    lock.waterfall_summary) — this just returns that data as JSON
    instead of rendering it into a PDF.

    Same "frozen, not a moving target" requirement as ICSummaryPDFView:
    requires the deal's baseline to already be locked, and reads
    EXCLUSIVELY from ExcelBaselineLock.analysis, never the project's
    current/latest one — a boardroom presentation shows the numbers
    the VP actually signed off on, not whatever's been edited since.

    Seller vs. Underwritten Table slide: reuses
    _compute_seller_vs_underwritten(project, analysis) VERBATIM — the
    exact same "Sprint Workflow Upgrades Feature 2" table
    PropertyProjectDetailView.get() already shows in the deal drawer
    (baseline-extracted value vs. latest re-uploaded/edited value, for
    GPR/Vacancy-Concessions/EGI/OpEx/NOI) — called against the LOCKED
    analysis specifically (not the project's current/latest one), so
    this slide reflects the baseline that was actually signed off on
    plus whatever was re-uploaded against THAT locked analysis, never a
    newer re-analysis.

    Forensic Audit Flags slide: reads metrics['reconciliation']['flags']
    (Module 1 — see reconciliation.py) directly — a DIFFERENT, separate
    concept from the table above (internal OM/T12/Rent Roll document
    consistency, not baseline-vs-edit drift). An empty list is a real,
    meaningful result (every check was within the 2% tolerance, or
    reconciliation never ran for this deal — e.g. no rent roll was
    submitted) and the frontend should say so explicitly rather than
    treating the slide as broken.

    Local Comps slide: reads whatever V2 Module 2 cached
    PropertyMarketDataSnapshot already exists for this deal — this
    view NEVER triggers a fresh RentCast fetch (see MarketCompsView
    above); opening a presentation must never silently spend the
    account's capped monthly RentCast quota. A deal whose Market & Tax
    Comps were never checked simply gets market_comps: null, and the
    frontend shows a prompt to check it from the main panel instead.

    Deliberately just visible_projects_for() — same reasoning as
    ICSummaryPDFView: a read/export action on a deal you can already
    see, not a workflow action like VPSignOffView/ApproveLockBaselineView.
    URL: GET /api/ai/pipeline/projects/<int:project_id>/present-deal/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, project_id):
        from .permissions import visible_projects_for
        from .market_comps import build_market_comps_payload

        project = visible_projects_for(request.user).select_related(
            'organization', 'workspace', 'assigned_analyst',
            'baseline_lock', 'baseline_lock__locked_by', 'baseline_lock__analysis',
            'market_data_snapshot',
        ).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        lock = getattr(project, 'baseline_lock', None)
        if lock is None:
            return Response(
                {'detail': (
                    "This deal's baseline must be locked before Boardroom Presentation "
                    "Mode is available — use Approve & Lock Baseline, or Sign Off if "
                    "it's in VP Review."
                )},
                status=status.HTTP_400_BAD_REQUEST,
            )

        analysis = lock.analysis
        metrics = analysis.metrics or {}
        rent_roll_df = (
            pd.DataFrame((analysis.rent_roll_data or {}).get('rows', []))
            if analysis.rent_roll_data else None
        )
        deal_scorecard = _compute_deal_scorecard(
            metrics, rent_roll_df, analysis.documents_provided, project.asking_price
        )
        debt_sizing = _compute_debt_sizing(
            project.organization, metrics, asking_price=project.asking_price
        )
        property_metadata = metrics.get('property_metadata') or {}
        reconciliation = metrics.get('reconciliation') or {}

        market_comps_payload = None
        snapshot = getattr(project, 'market_data_snapshot', None)
        if snapshot is not None:
            market_comps_payload = build_market_comps_payload(snapshot, metrics, project.asking_price)

        return Response({
            'property_overview': {
                'property_name': project.property_name or 'Untitled',
                'address': property_metadata.get('address'),
                'asset_class': property_metadata.get('asset_class'),
                'total_unit_count': property_metadata.get('total_unit_count'),
                'workspace_name': project.workspace.name if project.workspace_id else None,
                'assigned_analyst': project.assigned_analyst.email if project.assigned_analyst_id else None,
                'pipeline_stage': project.pipeline_stage,
                'asking_price': project.asking_price,
                'target_closing_or_offer_date': (
                    project.target_closing_or_offer_date.isoformat()
                    if project.target_closing_or_offer_date else None
                ),
                'noi': deal_scorecard.get('noi'),
                'in_place_cap_rate': deal_scorecard.get('in_place_cap_rate'),
                'opex_ratio': deal_scorecard.get('opex_ratio'),
                'price_per_unit_or_sf': deal_scorecard.get('price_per_unit_or_sf'),
                'price_per_unit_or_sf_basis': deal_scorecard.get('price_per_unit_or_sf_basis'),
            },
            'seller_vs_underwritten': _compute_seller_vs_underwritten(project, analysis),
            'forensic_audit_flags': {
                'flags': reconciliation.get('flags') or [],
                'note': reconciliation.get('note'),
            },
            'debt_sizing': debt_sizing,
            'waterfall_summary': lock.waterfall_summary,
            'market_comps': market_comps_payload,
            'locked_at': lock.locked_at.isoformat(),
            'locked_by': lock.locked_by.email if lock.locked_by else None,
        })


# Pipeline Modules Phase 4 (@Mentions, in-app only): matches "@" followed
# immediately by a full email address inside a Team Note's body text —
# e.g. "cc @alex@acme.com on this". Typing the full email is the entire
# mechanism for v1: this codebase has no org-member-directory/autocomplete
# endpoint to build a nicer @-picker against yet (confirmed by inspection —
# a real one is a reasonable later addition, not built here to keep this
# phase scoped), and no email notification infrastructure exists either —
# a mention only ever surfaces as an in-app highlight/"mentions you" filter
# on the unified Activity Feed below, never a sent email.
_MENTION_RE = re.compile(r'@([A-Za-z0-9_.+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9.-]+)')


def _parse_mentions(body, organization_id):
    """
    Resolves every @-mentioned email in `body` against a member of
    `organization_id` — the exact same org-scoped
    CustomUser.objects.filter(email=..., organization_id=...) pattern
    already used for PropertyProject.assigned_analyst in
    PropertyProjectDetailView.patch(), so "who counts as a valid
    mention" never drifts from "who counts as a valid assignee".
    Case-insensitive (email addresses aren't case-sensitive in practice
    here — CustomUser.email is normalized lowercase at creation). A
    mention of an email with no matching org member (a typo, or someone
    outside this org) simply resolves to nobody — silently dropped, not
    an error, since a comment's body is otherwise free-form text.
    Returns a CustomUser queryset (possibly empty), never a list, so the
    caller can hand it straight to a ManyToManyField's .set().
    """
    from users.models import CustomUser

    emails = {m.group(1).lower() for m in _MENTION_RE.finditer(body)}
    if not emails:
        return CustomUser.objects.none()
    return CustomUser.objects.filter(email__in=emails, organization_id=organization_id)


class TeamNoteListCreateView(APIView):
    """
    Team Notes on a deal — a flat, unthreaded running log (see TeamNote's
    own docstring for why; v1 scope, same reasoning as IssueReport). GET/POST
    both require the same visibility as the project itself — anyone who can
    see a deal can read and add to its notes.

    Pipeline Modules Phase 4: a note's body can now carry @mentions (see
    _parse_mentions above) and is interleaved into the pipeline drawer's
    unified Activity Feed as a "comment" event (see
    _get_deal_activity_feed) rather than shown in a separate section —
    this endpoint itself is unchanged either way; it's still the one
    place a Team Note is read or written.
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

        notes = project.notes.select_related('author').prefetch_related('mentioned_users').order_by('-created_at')
        return Response([
            {
                'id': n.id,
                'body': n.body,
                'author': n.author.email if n.author else None,
                'created_at': n.created_at.isoformat(),
                'mentioned_users': [u.email for u in n.mentioned_users.all()],
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

        # Pipeline Modules Phase 4 (@Mentions): resolved once, at creation
        # — see TeamNote.mentioned_users' own docstring in models.py for
        # why there's no "re-parse on edit" path.
        mentioned = list(_parse_mentions(body, project.organization_id))
        if mentioned:
            note.mentioned_users.set(mentioned)

        return Response({
            'id': note.id,
            'body': note.body,
            'author': note.author.email if note.author else None,
            'created_at': note.created_at.isoformat(),
            'mentioned_users': [u.email for u in mentioned],
        }, status=status.HTTP_201_CREATED)


def _serialize_checklist_item(item):
    return {
        'id': item.id,
        'doc_key': item.doc_key,
        'label': item.label,
        'is_resolved': item.is_resolved,
        'created_by': item.created_by.email if item.created_by else None,
        'resolved_by': item.resolved_by.email if item.resolved_by else None,
        'created_at': item.created_at.isoformat(),
        'resolved_at': item.resolved_at.isoformat() if item.resolved_at else None,
    }


class DealChecklistListCreateView(APIView):
    """
    Pipeline Modules Phase 4 (Persistent Missing-Doc Checklist): GET lists
    every checklist item on a deal — the 3 core documents, auto-seeded
    whenever an analysis run reports one missing (see
    _seed_missing_doc_checklist_items() in document_history.py), plus
    whatever the team has added by hand. POST adds a CUSTOM item only —
    there is no client-settable doc_key here; the 3 doc-backed items only
    ever come from the auto-seed path, never from this endpoint.

    Same visibility gate as Team Notes (GET/POST both require only that
    the caller can see the deal, not can_manage_project) — a checklist
    item is a collaborative, low-stakes artifact like a note, not a
    pipeline-stage change.
    URL: GET/POST /api/ai/pipeline/projects/<int:project_id>/checklist/
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

        items = project.checklist_items.select_related('created_by', 'resolved_by')
        return Response([_serialize_checklist_item(i) for i in items])

    def post(self, request, project_id):
        from .models import DealChecklistItem

        project, error = self._get_visible_project_or_404(request, project_id)
        if error:
            return error

        label = (request.data.get('label') or '').strip()
        if not label:
            return Response({'detail': 'A label is required for a custom checklist item.'}, status=status.HTTP_400_BAD_REQUEST)

        # doc_key is deliberately never accepted from the client here —
        # see this view's own docstring. Every item created through this
        # endpoint is a custom, hand-added one.
        item = DealChecklistItem.objects.create(
            project=project, label=label, created_by=request.user,
        )

        return Response(_serialize_checklist_item(item), status=status.HTTP_201_CREATED)


class DealChecklistItemResolveView(APIView):
    """
    Pipeline Modules Phase 4: toggle (or explicitly set) one checklist
    item's resolved state. Body: {"is_resolved": true|false} — explicit,
    not a bare toggle-with-no-body, so a double-click/retry can't
    accidentally flip it twice. Setting resolved_by/resolved_at only on
    the transition INTO resolved — clearing both back to null when
    un-resolving, rather than leaving a stale "resolved by/at" pair on an
    item that's currently showing as outstanding again.
    URL: PATCH /api/ai/pipeline/projects/<int:project_id>/checklist/<int:item_id>/resolve/
    """
    permission_classes = [IsAuthenticated]

    def patch(self, request, project_id, item_id):
        from .permissions import visible_projects_for
        from .models import DealChecklistItem

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        item = DealChecklistItem.objects.filter(pk=item_id, project=project).first()
        if item is None:
            return Response({'detail': 'Checklist item not found.'}, status=status.HTTP_404_NOT_FOUND)

        if 'is_resolved' not in request.data:
            return Response({'detail': "'is_resolved' (true/false) is required."}, status=status.HTTP_400_BAD_REQUEST)

        is_resolved = bool(request.data.get('is_resolved'))
        item.is_resolved = is_resolved
        if is_resolved:
            item.resolved_by = request.user
            item.resolved_at = timezone.now()
        else:
            item.resolved_by = None
            item.resolved_at = None
        item.save(update_fields=['is_resolved', 'resolved_by', 'resolved_at'])

        return Response(_serialize_checklist_item(item))


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


# Shared between ExcelOverrideCheckView (a real re-uploaded file) and
# Pipeline Modules Phase 8's Live Grid Editor (GridFieldValuesView /
# GridEditSaveView below) — both roads lead to the same "nothing to diff
# against yet" condition, so they use the exact same wording rather than
# two messages that could drift apart.
OVERRIDE_SNAPSHOT_UNAVAILABLE_DETAIL = (
    "Override tracking isn't available yet for this deal — download the "
    "Excel from this project at least once (the Live Grid Editor needs "
    "that same snapshot), then editing will work."
)


def _apply_excel_override(*, project, analysis, snapshot, workbook, file_bytes,
                           original_filename, user, variance_note, vault_file_type,
                           vault_label):
    """
    The core diff-and-persist logic shared by ExcelOverrideCheckView's
    "Re-upload Edited Underwriting File" (a real re-uploaded .xlsx/.xlsm)
    and Pipeline Modules Phase 8's Live Grid Editor "Save" action (a
    minimal workbook synthesized in-memory purely to carry the grid's
    edited values through this SAME named-range-diff machinery — see
    _build_grid_edit_workbook below). Neither caller's own identity
    matters here: both already resolved `workbook` down to "a workbook
    whose defined names we can read," and this function doesn't care
    which path produced it.

    Returns a DRF Response either way — a 400 for "no tracked fields
    found" or "a >2% field changed without a variance_note yet", or a 201
    with the newly-created ExcelOverrideUpload serialized. Never raises
    for an ordinary validation outcome; the caller is only responsible for
    getting as far as a parsed `workbook` and a non-empty `snapshot`.
    """
    from django.db import transaction
    from .models import ExcelOverrideUpload, ExcelOverrideDelta

    resolved_names, _skipped = extract_defined_names(workbook)
    if not resolved_names:
        return Response(
            {'detail': (
                "No tracked fields were found — this doesn't look like a "
                "workbook generated by this platform."
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

    # Sprint "Workflow Upgrades" Feature 2 (Variance Alert Context Notes):
    # requires a leadership-facing 1-sentence reason BEFORE this commits if
    # any detected delta touches one of the same three fields/threshold
    # already used by the firm-wide alert feed below
    # (OVERRIDE_ALERT_TRACKED_FIELDS / OVERRIDE_ALERT_THRESHOLD_PCT) —
    # intentionally the exact same fields/threshold, not a separate
    # configuration, so "what triggers a leadership alert" and "what
    # requires an explanation" never drift apart. Nothing has been written
    # to the database yet at this point (deltas is still a plain list), so
    # returning early here needs no rollback — the caller is expected to
    # resubmit with variance_note set once the user supplies one.
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
                    "is recorded."
                ),
                'requires_variance_note': True,
            },
            status=status.HTTP_400_BAD_REQUEST
        )

    with transaction.atomic():
        upload = ExcelOverrideUpload.objects.create(
            project=project,
            analysis=analysis,
            uploaded_by=user,
            original_filename=original_filename,
            # Feature 2 (Immutable Snapshots) extension: labels this
            # upload as a post-lock "child version branch" when the deal
            # already has a baseline lock at the moment of this
            # re-upload/save — see ExcelOverrideUpload.is_post_lock_revision
            # and ExcelBaselineLock's own docstring. Checked via hasattr
            # rather than a query — baseline_lock is a OneToOneField
            # reverse accessor that raises RelatedObjectDoesNotExist (not
            # returning None) when unset.
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

        # Pipeline Modules Phase 5 (Excel Versioning): vault the exact
        # bytes this came from — byte-identical, no regeneration needed,
        # unlike the locked-baseline-on-lock path (see
        # _vault_locked_baseline_excel's own docstring for why THAT path
        # can't do the same). try/except INSIDE this atomic block, not
        # around it — a vaulting failure must never roll back (or block)
        # the override record this action exists to create; it's only
        # ever best-effort.
        try:
            from .models import DealVaultFile
            from django.core.files.base import ContentFile

            DealVaultFile.objects.create(
                project=project,
                file=ContentFile(file_bytes, name=original_filename or 'override.xlsx'),
                file_type=vault_file_type,
                original_filename=original_filename,
                version_label=f"{vault_label} — {upload.uploaded_at:%Y-%m-%d %H:%M}",
                uploaded_by=user,
            )
        except Exception:
            logger.exception(
                "Failed to vault %s for project %s (override upload %s) — "
                "the override record itself still succeeded.",
                vault_file_type, project.id, upload.id,
            )

    return Response({'upload': _serialize_override_upload(upload)}, status=status.HTTP_201_CREATED)


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
    model's own docstring for why an unchanged field produces no row. The
    actual diff/persist/vault logic lives in the shared _apply_excel_
    override above — Pipeline Modules Phase 8's Live Grid Editor
    (GridEditSaveView below) uses that exact same function with a
    synthesized workbook instead of a real re-upload.

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
    is an audit action on the deal, not a change to pipeline_stage itself.
    URL: POST /api/ai/pipeline/projects/<int:project_id>/override-check/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request, project_id):
        from .permissions import visible_projects_for

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        analysis = project.analyses.order_by('-created_at').first()
        if analysis is None:
            return Response({'detail': 'This deal has no analysis yet.'}, status=status.HTTP_404_NOT_FOUND)

        snapshot = analysis.injected_values_snapshot
        if not snapshot:
            return Response({'detail': OVERRIDE_SNAPSHOT_UNAVAILABLE_DETAIL}, status=status.HTTP_400_BAD_REQUEST)

        uploaded_file = request.FILES.get('file')
        if not uploaded_file:
            return Response({'detail': 'Please upload the edited Excel file.'}, status=status.HTTP_400_BAD_REQUEST)

        try:
            validate_uploaded_file(uploaded_file, {'.xlsx', '.xlsm'}, label="Edited Underwriting File")
        except FileValidationError as e:
            return Response({'detail': str(e)}, status=status.HTTP_400_BAD_REQUEST)

        # Captured once, up front — Pipeline Modules Phase 5 vaults these
        # EXACT bytes (byte-identical, no regeneration needed) once this
        # upload is actually recorded below. uploaded_file.read() can only
        # be called once against the underlying stream, so every later
        # use in this method reads from this variable, never the file
        # object again.
        uploaded_file_bytes = uploaded_file.read()
        try:
            workbook = openpyxl.load_workbook(io.BytesIO(uploaded_file_bytes), data_only=False)
        except Exception as e:
            return Response(
                {'detail': f"Couldn't open this file as an Excel workbook: {e}"},
                status=status.HTTP_400_BAD_REQUEST
            )

        variance_note = (request.data.get('variance_note') or '').strip()

        return _apply_excel_override(
            project=project, analysis=analysis, snapshot=snapshot, workbook=workbook,
            file_bytes=uploaded_file_bytes,
            original_filename=getattr(uploaded_file, 'name', '') or '',
            user=request.user, variance_note=variance_note,
            vault_file_type='re_upload', vault_label='Re-upload',
        )


# ---------------------------------------------------------------------------
# Pipeline Modules Phase 8: Live In-Browser Spreadsheet Editor
# ---------------------------------------------------------------------------
# Explicit client decision from design review: a CONSTRAINED grid bound
# only to the already-tracked IU_* fields (SCALAR_FIELD_CATALOG /
# ROW_FIELD_CATALOG — the exact same catalog "Re-upload Edited
# Underwriting File" and the custom-template path already use), not a
# full formula-driven workbook clone. "Save" never introduces a new
# live-edit data model of its own — it reuses ExcelOverrideCheckView's
# diff/variance-alert/vault pipeline (_apply_excel_override above)
# end-to-end, via a MINIMAL synthetic workbook (_build_grid_edit_workbook)
# built purely to carry the grid's edited values through that existing
# named-range-diff machinery.
#
# v2 (Live Grid Editor for Custom Excel Templates): this grid is bound to
# OUR field catalog either way, so a ClientExcelTemplate org edits through
# EXACTLY the same snapshot/diff/variance/vault pipeline as a standard
# org below — nothing in _build_grid_edit_workbook or _apply_excel_override
# cares which path produced the snapshot it's diffing against, since both
# already resolve values through the IU_* catalog, just via different
# physical cells. The only custom-template-specific behavior lives in
# GridFieldValuesView.get() below: it additionally resolves each known
# IU_* name against the ORG'S OWN uploaded template (read-only — see
# _resolve_custom_template_field_locations) so the grid can show which
# tab/cell each field actually maps to, same information an admin can
# already see on the Excel Template settings panel. "Save" writes no
# differently for a custom-template org than for a standard one — it was
# never touching any real workbook in the first place (see
# _build_grid_edit_workbook's own docstring), so there was nothing
# custom-template-specific to adapt there at all. The deal's ACTUAL custom
# .xlsx is still only ever generated fresh at Download/Lock time from
# analysis.metrics/rent_roll_data (see generate_excel_from_template's
# caller above) — this grid never loads, mutates, or re-saves that file.


def _build_grid_edit_workbook(scalars: dict, rows: list):
    """
    Builds a small, purely-mechanical in-memory workbook whose ONLY job is
    to carry the Live Grid Editor's submitted values through
    extract_defined_names() + _apply_excel_override() exactly the way a
    real re-uploaded "Source Values (IU)" sheet/Argus tab would (see
    excel_generator.py's own Named Ranges section for that convention).
    Deliberately ugly/undecorated — this is never shown to a human as a
    deliverable, only vaulted for audit (file_type='live_grid_edit', a
    distinct label from a real re-upload) and immediately diffed.

    Silently drops any key that isn't a known IU_* name (defensive against
    stale/tampered client input — see GridEditSaveView.post()) and any row
    dict that is not actually a dict. Scalars get one row each (col B);
    row fields share one block below them, one column per field, written
    downward one row per tenant — anchored on a fresh defined name per
    field, same "first data cell, write downward" convention used
    everywhere else in this codebase.
    """
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = 'Grid Edit'

    r = 1
    for name, value in (scalars or {}).items():
        if name not in SCALAR_FIELD_CATALOG:
            continue
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=value if value not in (None, '') else None)
        wb.defined_names[name] = DefinedName(name, attr_text=f"'{ws.title}'!$B${r}")
        r += 1

    clean_rows = [row for row in (rows or []) if isinstance(row, dict)]
    row_field_names = sorted({
        k for row in clean_rows for k in row if k in ROW_FIELD_CATALOG
    })
    if row_field_names:
        anchor_row = r + 1  # one blank separator row between scalars and the row table
        for col_offset, name in enumerate(row_field_names):
            col = 2 + col_offset  # leave column A free, same as the scalar block above
            ws.cell(row=anchor_row - 1, column=col, value=name)  # cosmetic header only
            wb.defined_names[name] = DefinedName(
                name, attr_text=f"'{ws.title}'!${get_column_letter(col)}${anchor_row}"
            )
        for i, row_data in enumerate(clean_rows):
            for col_offset, name in enumerate(row_field_names):
                value = row_data.get(name)
                ws.cell(row=anchor_row + i, column=2 + col_offset, value=value if value not in (None, '') else None)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _resolve_custom_template_field_locations(custom_template):
    """
    Live Grid Editor for Custom Excel Templates: opens THIS org's own
    uploaded template (read-only — never re-saved, same discipline as
    validate_template() in custom_template_generator.py) and resolves
    every known IU_* named range to a human-readable "Tab!Cell" location,
    so the grid can show an analyst exactly which tab/cell each field maps
    to in their own workbook, same as an admin can already see (in raw
    field-name form) on the Excel Template settings panel.

    Deliberately re-resolves from the live file on every call rather than
    trusting ClientExcelTemplate.detected_named_ranges — that column is
    only ever a flat list of field NAMES (captured at upload time for the
    admin's benefit), never the (sheet, row, col) a display needs, and
    re-opening a workbook purely to read its defined names is cheap (the
    same thing validate_template() already does on every upload).

    Best-effort: if the stored file can't be opened for any reason (a
    corrupt upload, a storage hiccup), returns None rather than failing
    the whole grid-values request — the grid still works with plain field
    labels, it just loses the Tab/Cell column for this one request.
    """
    from openpyxl.utils import get_column_letter

    try:
        workbook = openpyxl.load_workbook(custom_template.file, data_only=False, read_only=True)
    except Exception:
        logger.exception(
            "Live Grid Editor: couldn't open the custom template for org %s "
            "to resolve field locations — grid will still work without the "
            "Tab/Cell column.",
            custom_template.organization_id,
        )
        return None

    try:
        resolved, _skipped = extract_defined_names(workbook)
    finally:
        workbook.close()

    field_locations = {
        name: {'sheet': sheet_name, 'cell': f"{get_column_letter(col)}{row}"}
        for name, (sheet_name, row, col) in resolved.items()
        if name in SCALAR_FIELD_CATALOG or name in ROW_FIELD_CATALOG
    }
    return {
        'original_filename': custom_template.original_filename,
        'field_locations': field_locations,
    }


class GridFieldValuesView(APIView):
    """
    GET: the current editable values for the Live Grid Editor — exactly
    analysis.injected_values_snapshot (the same {scalars, rows} shape
    "Re-upload Edited Underwriting File" diffs against), so the grid's
    starting values and the backend's diff baseline can never silently
    drift apart from each other. No separate re-resolution of SCALAR_
    FIELD_CATALOG from metrics here, by design — see that field's own
    docstring for why it's only ever captured at actual generation time.

    Same "no snapshot yet" gate as the Save action below, so a disabled
    grid and a rejected save are never surprises of each other.

    For an organization on an active ClientExcelTemplate, the response
    also includes a `custom_template` object ({original_filename,
    field_locations}) resolving each known field to where it actually
    lives in THIS org's own workbook — see
    _resolve_custom_template_field_locations above. `custom_template` is
    null for a standard-template org, and best-effort null (logged, not
    raised) if this org's template file couldn't be opened — either way
    the grid itself still renders from `scalars`/`rows` as normal.
    URL: GET /api/ai/pipeline/projects/<int:project_id>/grid-values/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, project_id):
        from .permissions import visible_projects_for
        from .models import ClientExcelTemplate

        project = visible_projects_for(request.user).select_related('organization').filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        analysis = project.analyses.order_by('-created_at').first()
        if analysis is None:
            return Response({'detail': 'This deal has no analysis yet.'}, status=status.HTTP_404_NOT_FOUND)

        snapshot = analysis.injected_values_snapshot
        if not snapshot:
            return Response({'detail': OVERRIDE_SNAPSHOT_UNAVAILABLE_DETAIL}, status=status.HTTP_400_BAD_REQUEST)

        custom_template_info = None
        custom_template = ClientExcelTemplate.objects.filter(
            organization=project.organization_id, is_active=True
        ).first()
        if custom_template is not None:
            custom_template_info = _resolve_custom_template_field_locations(custom_template)

        return Response({
            'scalars': snapshot.get('scalars') or {},
            'rows': snapshot.get('rows') or [],
            'is_post_lock_revision': hasattr(project, 'baseline_lock'),
            'custom_template': custom_template_info,
        })


def _sync_grid_edits_to_working_state(analysis, scalars, rows):
    """
    Follow-up fix to Pipeline Modules Phase 8: makes a Live Grid Editor
    "Save" actually update the deal's WORKING state — analysis.metrics
    and analysis.rent_roll_data, the exact two fields
    PropertyProjectDownloadView and _vault_locked_baseline_excel both
    regenerate an Excel from — not just the ExcelOverrideUpload/Delta
    audit trail _apply_excel_override already writes. Before this,
    "Download .xlsx" and "Approve & Lock Baseline" would silently keep
    reflecting the ORIGINAL extraction forever, no matter how many grid
    edits were saved — only the audit log and the vaulted synthetic
    workbook ever saw them.

    Also recomputes analysis.injected_values_snapshot from the new
    metrics/rent_roll_data, so reopening the grid shows what was just
    saved rather than appearing to silently revert it, and so the NEXT
    save's diff/variance-alert check is against this save's values, not
    the original extraction forever. Deliberately does NOT touch
    ExcelOverrideCheckView's re-upload path at all — a plain re-upload's
    "working state" semantics are intentionally left exactly as they
    were; this sync only ever runs from GridEditSaveView below.

    One explicit, permanent limitation: IU_DebtLTVPct/
    IU_DebtInterestRatePct/IU_DebtAmortizationYears have no persisted
    home to write into (see SCALAR_FIELD_WRITE_PATHS's own comment in
    custom_template_generator.py — debt assumptions are re-prompted at
    every download/lock, never stored on an analysis). Whatever the
    analyst saved for those 3 fields is kept in the snapshot (so the
    grid itself keeps showing it) and in the audit trail, but can never
    flow through to Download/Lock. Likewise IU_Row_AnnualRent/
    IU_Row_MLAProfile are always recomputed at generation time (never
    read from a real rent-roll column), so an edit to either is tracked
    in the audit trail but never changes what a regeneration produces.

    Best-effort, like DealVaultFile's own vaulting step inside
    _apply_excel_override: by the time this runs, the override/audit
    record it's called after has ALREADY been committed successfully, so
    a failure here is logged and swallowed rather than surfaced as a
    failed save — worst case, Download/Lock keeps showing the pre-edit
    numbers until the next successful save, never silently corrupted
    data.
    """
    try:
        metrics = analysis.metrics or {}
        metrics, _written = apply_scalar_edits_to_metrics(metrics, scalars)

        rent_roll_data = apply_row_edits_to_rent_roll_data(analysis.rent_roll_data, rows)
        rent_roll_df = pd.DataFrame((rent_roll_data or {}).get('rows') or [])

        new_snapshot = compute_injected_values_snapshot(metrics, None, rent_roll_df)

        old_scalars = (analysis.injected_values_snapshot or {}).get('scalars') or {}
        for debt_field in ('IU_DebtLTVPct', 'IU_DebtInterestRatePct', 'IU_DebtAmortizationYears'):
            if debt_field in (scalars or {}):
                new_snapshot['scalars'][debt_field] = scalars[debt_field]
            elif debt_field in old_scalars:
                new_snapshot['scalars'][debt_field] = old_scalars[debt_field]

        analysis.metrics = metrics
        analysis.rent_roll_data = rent_roll_data
        analysis.injected_values_snapshot = new_snapshot
        analysis.save(update_fields=['metrics', 'rent_roll_data', 'injected_values_snapshot'])
    except Exception:
        logger.exception(
            "Live Grid Editor: saved override/audit record for analysis "
            "%s, but failed to sync edits into metrics/rent_roll_data — "
            "Download/Lock will keep showing pre-edit values until the "
            "next successful save.",
            analysis.id,
        )


class GridEditSaveView(APIView):
    """
    POST: "Save" on the Live Grid Editor. Body: {"scalars": {IU_name:
    value, ...}, "rows": [{IU_Row_name: value, ...}, ...], "variance_note":
    "..." (optional, required only if a tracked field crossed the >2%
    alert threshold — see _apply_excel_override)}.

    Builds a minimal synthetic workbook from the posted values
    (_build_grid_edit_workbook) and hands it to the EXACT SAME
    _apply_excel_override() "Re-upload Edited Underwriting File" itself
    calls — this endpoint exists only to translate "JSON from a grid"
    into "a workbook extract_defined_names() can read," not to duplicate
    any of the diff/variance-alert/vault logic itself. See this module's
    Phase 8 section comment above for the full design.

    On a successful save (201), ALSO calls _sync_grid_edits_to_working_state
    so Download/Lock pick up the edits too — see that function's own
    docstring for exactly what does and doesn't flow through, and why
    it's deliberately never invoked from ExcelOverrideCheckView's
    re-upload path.

    Only known IU_* names are ever written into that synthetic workbook —
    anything else in the posted body is silently ignored, the same
    defensive posture every other endpoint in this codebase takes toward
    client-supplied field names.

    Works identically for an organization on an active ClientExcelTemplate
    — this view never touches any real workbook, custom or standard (see
    _build_grid_edit_workbook's own docstring), so there's nothing
    custom-template-specific to branch on here. The org's own template
    file is only ever read (never written) by GridFieldValuesView above,
    purely to show where each field lives.
    URL: POST /api/ai/pipeline/projects/<int:project_id>/grid-save/
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

        snapshot = analysis.injected_values_snapshot
        if not snapshot:
            return Response({'detail': OVERRIDE_SNAPSHOT_UNAVAILABLE_DETAIL}, status=status.HTTP_400_BAD_REQUEST)

        scalars = request.data.get('scalars')
        rows = request.data.get('rows')
        if not isinstance(scalars, dict) or not isinstance(rows, list):
            return Response(
                {'detail': '"scalars" (object) and "rows" (array) are both required.'},
                status=status.HTTP_400_BAD_REQUEST
            )

        variance_note = (request.data.get('variance_note') or '').strip()

        file_bytes = _build_grid_edit_workbook(scalars, rows)
        workbook = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)

        response = _apply_excel_override(
            project=project, analysis=analysis, snapshot=snapshot, workbook=workbook,
            file_bytes=file_bytes,
            # Not a real filename — there is no uploaded file on this path
            # (see this view's own docstring) — but original_filename is
            # exactly what the Override Summary / Vault UIs display, so a
            # readable label here is more honest than a fabricated-looking
            # "grid-edit-42.xlsx".
            original_filename="Live Grid Editor",
            user=request.user, variance_note=variance_note,
            vault_file_type='live_grid_edit', vault_label='Live Grid Edit',
        )
        if response.status_code == status.HTTP_201_CREATED:
            # The "Download .xlsx"/"Approve & Lock Baseline" sync fix —
            # see _sync_grid_edits_to_working_state's own docstring.
            _sync_grid_edits_to_working_state(analysis, scalars, rows)
        return response


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


# ---------------------------------------------------------------------------
# V2 Module 3 (Live Debt & Equity Waterfall Scenario Builder) — Phase V2-1.
# Pure Python, $0/month infrastructure, no AI API calls: see
# waterfall_calculator.py's own module docstring for the full algorithm
# and its documented modeling conventions. Backend Python is the
# AUTHORITATIVE calculation (an explicit build-order decision — the
# number an analyst sees here is the exact same code path that feeds the
# locked baseline snapshot and the IC Summary PDF, never a second JS
# implementation that could silently drift from it).
# ---------------------------------------------------------------------------

DEFAULT_PROMOTE_TIERS = [
    {'hurdle_irr_pct': 8.0, 'gp_promote_pct': 0.0},
    {'hurdle_irr_pct': None, 'gp_promote_pct': 20.0},
]


def _resolve_going_in_noi(analysis):
    """Same T12 NOI metrics already resolve to everywhere else in this
    codebase (SCALAR_FIELD_CATALOG['IU_NetOperatingIncome']) — the
    waterfall's going-in NOI is never a separate, independently-typed
    input; it's always this deal's own underwritten NOI."""
    return safe_get(analysis.metrics or {}, 't12_revenue_expenses', 'net_operating_income')


class WaterfallScenarioView(APIView):
    """
    GET: the deal's saved Scenario Builder state — {"inputs": {...},
    "results": {...}, "computed_at": ...} from
    AnalysisReport.waterfall_scenario, or a seeded-but-uncalculated
    default (DEFAULT_PROMOTE_TIERS, a blank debt structure, results:
    null) if the analyst has never run "Calculate" on this deal before.
    Also returns going_in_noi so the UI can show WHY a scenario can't be
    built yet (no positive NOI) without a wasted round-trip to Calculate.
    URL: GET /api/ai/pipeline/projects/<int:project_id>/waterfall-scenario/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, project_id):
        from .permissions import visible_projects_for

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        analysis = project.latest_analysis
        if analysis is None:
            return Response({'detail': 'This deal has no analysis yet.'}, status=status.HTTP_404_NOT_FOUND)

        going_in_noi = _resolve_going_in_noi(analysis)
        saved = analysis.waterfall_scenario or {}

        return Response({
            'inputs': saved.get('inputs') or {
                'debt_structure': 'fixed_rate_permanent',
                'purchase_price': project.asking_price,
                'promote_tiers': DEFAULT_PROMOTE_TIERS,
            },
            'results': saved.get('results'),
            'computed_at': saved.get('computed_at'),
            'going_in_noi': going_in_noi,
            'is_post_lock_revision': hasattr(project, 'baseline_lock'),
        })


class WaterfallScenarioCalculateView(APIView):
    """
    POST: "Calculate" on the Scenario Builder. Body is the raw `inputs`
    dict compute_waterfall() expects (debt_structure, purchase_price,
    hold_period_years, exit_cap_rate_pct, acquisition_ltv_pct,
    equity_split_pct, promote_tiers, plus the structure-specific fields —
    see waterfall_calculator.REQUIRED_FIELDS_*). Runs the authoritative
    Python calculation, persists {"inputs", "results", "computed_at"} to
    AnalysisReport.waterfall_scenario (overwriting any prior scenario —
    there is deliberately only ever ONE "current" scenario per deal for
    v1, not a saved list of named scenarios), and returns the results.

    A WaterfallInputError (missing/invalid field, or no positive NOI to
    build from) is a clean 400 with the exact reason — never a silently
    guessed default for a financial assumption.
    URL: POST /api/ai/pipeline/projects/<int:project_id>/waterfall-calculate/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request, project_id):
        from .permissions import visible_projects_for

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        analysis = project.latest_analysis
        if analysis is None:
            return Response({'detail': 'This deal has no analysis yet.'}, status=status.HTTP_404_NOT_FOUND)

        inputs = request.data if isinstance(request.data, dict) else None
        if not inputs:
            return Response({'detail': 'Scenario inputs are required.'}, status=status.HTTP_400_BAD_REQUEST)

        # purchase_price falls back to the deal's own asking price if the
        # analyst hasn't overridden it in the scenario itself — same
        # "AI-extracted/asking price is the default, always editable"
        # convention Module A's Target:/Price: tag override established.
        inputs = dict(inputs)
        if inputs.get('purchase_price') in (None, ''):
            inputs['purchase_price'] = project.asking_price

        going_in_noi = _resolve_going_in_noi(analysis)

        try:
            results = compute_waterfall(inputs, going_in_noi)
        except WaterfallInputError as e:
            return Response({'detail': str(e)}, status=status.HTTP_400_BAD_REQUEST)

        computed_at = timezone.now().isoformat()
        analysis.waterfall_scenario = {'inputs': inputs, 'results': results, 'computed_at': computed_at}
        analysis.save(update_fields=['waterfall_scenario'])

        return Response({'inputs': inputs, 'results': results, 'computed_at': computed_at})


# ---------------------------------------------------------------------------
# Pipeline Modules Phase 5 (Excel Versioning & Supporting Document Vault).
# See DealVaultFile's own docstring in models.py for exactly which 3 moments
# auto-capture a version (lock, re-upload) vs. the one manual upload path
# below — there is no "save version now" button anywhere in the UI.
# ---------------------------------------------------------------------------

ALLOWED_VAULT_SUPPORTING_DOC_EXTENSIONS = {
    '.pdf', '.doc', '.docx', '.xls', '.xlsx', '.xlsm', '.csv', '.png', '.jpg', '.jpeg',
}


def _serialize_vault_file(f):
    return {
        'id': f.id,
        'file_type': f.file_type,
        'file_type_display': f.get_file_type_display(),
        'original_filename': f.original_filename,
        'version_label': f.version_label,
        'uploaded_by': f.uploaded_by.email if f.uploaded_by else None,
        'uploaded_at': f.uploaded_at.isoformat(),
    }


class DealVaultListView(APIView):
    """
    Every versioned file kept against a deal — auto-vaulted Locked
    Baseline/Re-upload Excel snapshots plus any manually-added supporting
    documents — newest first. Read-only; see DealVaultFileUploadView
    below for adding a supporting document and DealVaultFileDownloadView
    for reading one back. Same visibility gate as every other read-only
    pipeline list (Team Notes, checklist) — anyone who can see the deal
    can see its vault.
    URL: GET /api/ai/pipeline/projects/<int:project_id>/vault/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, project_id):
        from .permissions import visible_projects_for

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        files = project.vault_files.select_related('uploaded_by')
        return Response([_serialize_vault_file(f) for f in files])


class DealVaultFileUploadView(APIView):
    """
    The ONE manual path into the vault: attaching an arbitrary supporting
    document (an appraisal, an insurance quote, a signed LOI) to a deal —
    see DealVaultFile's own docstring ('supporting_document') for why
    this is a deliberate, named exception to this platform's usual zero-
    retention policy. Same visibility gate as Team Notes/checklist — not
    can_manage_project — a supporting document is a collaborative
    artifact like a note, not a pipeline-stage change.
    URL: POST /api/ai/pipeline/projects/<int:project_id>/vault/upload/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request, project_id):
        from .permissions import visible_projects_for
        from .models import DealVaultFile

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        uploaded_file = request.FILES.get('file')
        try:
            validate_uploaded_file(
                uploaded_file, ALLOWED_VAULT_SUPPORTING_DOC_EXTENSIONS, label="Supporting Document"
            )
        except FileValidationError as e:
            return Response({'detail': str(e)}, status=status.HTTP_400_BAD_REQUEST)

        version_label = (request.data.get('version_label') or '').strip()

        vault_file = DealVaultFile.objects.create(
            project=project,
            file=uploaded_file,
            file_type='supporting_document',
            original_filename=uploaded_file.name,
            version_label=version_label,
            uploaded_by=request.user,
        )

        return Response(_serialize_vault_file(vault_file), status=status.HTTP_201_CREATED)


class DealVaultFileDownloadView(APIView):
    """
    Reads one vault file back — the only way any vault file is ever
    reachable, per PrivateFileSystemStorage's own docstring (no public
    URL exists). Visibility-checked against the file's OWN project
    first, then re-confirmed the file actually belongs to the
    project_id in the URL — a vault file id alone is never trusted to
    imply the right deal.
    URL: GET /api/ai/pipeline/projects/<int:project_id>/vault/<int:file_id>/download/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, project_id, file_id):
        import mimetypes
        from .permissions import visible_projects_for
        from .models import DealVaultFile

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        vault_file = DealVaultFile.objects.filter(pk=file_id, project=project).first()
        if vault_file is None:
            return Response({'detail': 'File not found.'}, status=status.HTTP_404_NOT_FOUND)

        filename = vault_file.original_filename or vault_file.file.name.rsplit('/', 1)[-1]
        content_type, _ = mimetypes.guess_type(filename)

        return FileResponse(
            vault_file.file.open('rb'),
            as_attachment=True,
            filename=filename,
            content_type=content_type or 'application/octet-stream',
        )


def _serialize_lender_share_link(link):
    return {
        'id': link.id,
        'access_token': link.access_token,
        'recipient_label': link.recipient_label,
        'is_active': link.is_active,
        'view_count': link.view_count,
        'last_viewed_at': link.last_viewed_at.isoformat() if link.last_viewed_at else None,
        'expires_at': link.expires_at.isoformat() if link.expires_at else None,
        'created_at': link.created_at.isoformat(),
        'is_expired': link.is_expired,
        'is_usable': link.is_usable,
    }


class LenderShareLinkListCreateView(APIView):
    """
    Pipeline Modules Phase 6: "Generate Lender Packet" — every tokenized
    public share link ever generated for a deal (GET), and generating a
    new one (POST). Gated to can_manage_project — the same bar as editing
    assigned_analyst/pipeline_stage — since a generated link is a
    standing, EXTERNALLY reachable grant on this deal's locked data, not
    a purely collaborative artifact like a Team Note or vault upload
    (those only require visible_projects_for).

    POST requires the deal to already be locked — see LenderShareLink's
    own docstring ("nothing to show before lock").

    expires_in_days (optional int in the POST body): omitted defaults to
    30 days (the client's explicit default); 0/null/'' means no expiry;
    a positive integer sets expires_at that many days from now.
    URL: GET/POST /api/ai/pipeline/projects/<int:project_id>/lender-links/
    """
    permission_classes = [IsAuthenticated]

    def _get_project(self, request, project_id):
        from .permissions import visible_projects_for
        return visible_projects_for(request.user).select_related(
            'workspace', 'baseline_lock'
        ).filter(pk=project_id).first()

    def get(self, request, project_id):
        project = self._get_project(request, project_id)
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        links = project.lender_share_links.all()
        return Response([_serialize_lender_share_link(l) for l in links])

    def post(self, request, project_id):
        from .permissions import can_manage_project
        from .models import LenderShareLink

        project = self._get_project(request, project_id)
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)
        if not can_manage_project(request.user, project):
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        if getattr(project, 'baseline_lock', None) is None:
            return Response(
                {'detail': (
                    "This deal's baseline must be locked before a Lender Packet can "
                    "be generated — use Approve & Lock Baseline, or Sign Off if it's "
                    "in VP Review."
                )},
                status=status.HTTP_400_BAD_REQUEST
            )

        recipient_label = (request.data.get('recipient_label') or '').strip()

        if 'expires_in_days' in request.data:
            raw_days = request.data.get('expires_in_days')
            expires_at = None
            if raw_days:
                try:
                    days = int(raw_days)
                except (TypeError, ValueError):
                    return Response(
                        {'detail': "'expires_in_days' must be an integer."},
                        status=status.HTTP_400_BAD_REQUEST
                    )
                if days > 0:
                    expires_at = timezone.now() + timedelta(days=days)
            # falsy (0 / null / '') -> no expiry, expires_at stays None.
        else:
            # Omitted entirely -> the client's explicit default.
            expires_at = timezone.now() + timedelta(days=30)

        link = LenderShareLink.objects.create(
            project=project,
            workspace=project.workspace,
            recipient_label=recipient_label,
            expires_at=expires_at,
            created_by=request.user,
        )
        return Response(_serialize_lender_share_link(link), status=status.HTTP_201_CREATED)


class LenderShareLinkRevokeView(APIView):
    """
    Toggles is_active on one of this deal's LenderShareLink rows — the
    manual revoke/reactivate switch (see LenderShareLink's own docstring
    for how this interacts with expires_at). Same can_manage_project gate
    as generating the link. The link id is re-confirmed against THIS
    project_id, same double-check pattern as DealVaultFileDownloadView —
    a link id alone is never trusted to imply the right deal.
    URL: PATCH /api/ai/pipeline/projects/<int:project_id>/lender-links/<int:link_id>/
    """
    permission_classes = [IsAuthenticated]

    def patch(self, request, project_id, link_id):
        from .permissions import visible_projects_for, can_manage_project
        from .models import LenderShareLink

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)
        if not can_manage_project(request.user, project):
            return Response({'detail': 'Forbidden'}, status=status.HTTP_403_FORBIDDEN)

        link = LenderShareLink.objects.filter(pk=link_id, project=project).first()
        if link is None:
            return Response({'detail': 'Lender link not found.'}, status=status.HTTP_404_NOT_FOUND)

        if 'is_active' not in request.data:
            return Response({'detail': "Provide 'is_active' (true/false)."}, status=status.HTTP_400_BAD_REQUEST)

        link.is_active = bool(request.data.get('is_active'))
        link.save(update_fields=['is_active'])
        return Response(_serialize_lender_share_link(link))


def _get_usable_lender_link_or_none(access_token):
    """
    Shared lookup for every public lender-portal view below. Deliberately
    collapses "no such token" / "revoked" / "expired" into the same None
    result — an unauthenticated caller is never told WHICH reason a token
    doesn't work (same minimal-disclosure principle as a login form never
    confirming whether an email exists), and every caller below turns
    None into a 404, NEVER a 401 — see LenderSharePortalView's own
    docstring for exactly why that distinction matters for this one
    group of endpoints.
    """
    from .models import LenderShareLink

    link = LenderShareLink.objects.select_related(
        'project', 'project__organization',
        'project__baseline_lock', 'project__baseline_lock__analysis',
    ).filter(access_token=access_token).first()
    if link is None or not link.is_usable:
        return None
    return link


class LenderSharePortalView(APIView):
    """
    Pipeline Modules Phase 6: the public, UNAUTHENTICATED read-only lender
    portal — the frontend's /share/deal/:token page calls this to render
    everything a lender sees. permission_classes=[AllowAny] (the same
    no-gate pattern as users.views.LogoutView, for the same reason: this
    request may have no valid session at all, by design — an external
    lender was never asked to create an account on this platform).

    IMPORTANT — this view (and every other public lender-portal view
    below) must NEVER return HTTP 401, under any failure path. Why: the
    frontend's global $fetch override (frontend/app/plugins/
    auth_interceptor.client.ts) force-logs-out and redirects to /login on
    ANY 401 response anywhere in the app. An external lender has no
    account and no /login to go to — a 401 here would silently strand
    them on a broken redirect instead of a plain "this link doesn't
    work" message. Every failure path below returns 404 instead.
    (Belt-and-suspenders note: CookieJWTAuthentication.authenticate()
    already returns None rather than raising for a missing OR invalid
    access_token cookie — see users/authentication.py — so an ordinary
    unauthenticated request was never going to 401 at the auth layer in
    the first place. authentication_classes=[] below is an extra,
    explicit guarantee on top of that, specifically for this one group of
    public-facing views where getting it wrong has a real user-facing
    consequence.)

    Bundles exactly what the spec calls for: the 1-Page Deal Scorecard &
    Executive Summary (the same _compute_deal_scorecard() computation
    ICSummaryPDFView uses, built off the LOCKED analysis only — never a
    moving "latest" one), whether a Locked Underwritten Excel Model is
    available (see LenderShareExcelDownloadView — absent for
    ClientExcelTemplate orgs, which skip auto-vault-on-lock in v1), and
    any vault 'supporting_document' files — deliberately EXCLUDING
    're_upload' vault entries (post-lock internal cell-override audit
    history, not meant for an external audience) and never exposing raw
    OM/T12/Rent Roll inputs, which this platform's zero-retention policy
    never persists anywhere to begin with.

    Increments view_count/last_viewed_at on every successful load — this
    feature's entire "view tracking" mechanism; see LenderShareLink's own
    docstring and _get_deal_activity_feed's 'lender_link_viewed' event for
    how this surfaces back to the team.
    URL: GET /api/ai/share/deal/<str:access_token>/
    """
    permission_classes = [AllowAny]
    authentication_classes = []

    def get(self, request, access_token):
        link = _get_usable_lender_link_or_none(access_token)
        if link is None:
            return Response({'detail': 'This link is no longer available.'}, status=status.HTTP_404_NOT_FOUND)

        project = link.project
        lock = getattr(project, 'baseline_lock', None)
        if lock is None:
            # Shouldn't happen in practice — generating a link requires a
            # lock, and nothing in this codebase ever unlocks a deal — but
            # this view must never crash the public portal if it somehow
            # does (e.g. data fixed up by hand).
            return Response({'detail': 'This link is no longer available.'}, status=status.HTTP_404_NOT_FOUND)

        analysis = lock.analysis
        rent_roll_df = (
            pd.DataFrame((analysis.rent_roll_data or {}).get('rows', []))
            if analysis.rent_roll_data else None
        )
        deal_scorecard = _compute_deal_scorecard(
            analysis.metrics, rent_roll_df, analysis.documents_provided, project.asking_price
        )

        has_locked_excel = project.vault_files.filter(file_type='locked_baseline_excel').exists()
        supporting_docs = project.vault_files.filter(file_type='supporting_document')

        link.view_count += 1
        link.last_viewed_at = timezone.now()
        link.save(update_fields=['view_count', 'last_viewed_at'])

        return Response({
            'recipient_label': link.recipient_label,
            'property_name': project.property_name or 'Untitled',
            'locked_at': lock.locked_at.isoformat(),
            'deal_scorecard': deal_scorecard,
            'has_locked_excel': has_locked_excel,
            'supporting_documents': [
                {
                    'id': f.id,
                    'original_filename': f.original_filename,
                    'version_label': f.version_label,
                    'uploaded_at': f.uploaded_at.isoformat(),
                }
                for f in supporting_docs
            ],
        })


class LenderShareExcelDownloadView(APIView):
    """
    Serves the Locked Underwritten Excel Model's bytes for one usable
    LenderShareLink — see LenderSharePortalView's docstring for the
    401-avoidance discipline every view in this group follows. Reads the
    'locked_baseline_excel' DealVaultFile vaulted at lock time (see
    _vault_locked_baseline_excel above), exactly like
    DealVaultFileDownloadView does for an authenticated team member, just
    gated by the share token instead of visible_projects_for.
    404s (never 401) for a bad/expired/inactive token OR an org with no
    vaulted locked-baseline Excel (e.g. a ClientExcelTemplate org — see
    DealVaultFile's own docstring on why those are skipped in v1).
    URL: GET /api/ai/share/deal/<str:access_token>/excel/
    """
    permission_classes = [AllowAny]
    authentication_classes = []

    def get(self, request, access_token):
        import mimetypes
        from .models import DealVaultFile

        link = _get_usable_lender_link_or_none(access_token)
        if link is None:
            return Response({'detail': 'This link is no longer available.'}, status=status.HTTP_404_NOT_FOUND)

        vault_file = DealVaultFile.objects.filter(
            project=link.project, file_type='locked_baseline_excel'
        ).order_by('-uploaded_at').first()
        if vault_file is None:
            return Response(
                {'detail': 'No locked Excel model is available for this deal.'},
                status=status.HTTP_404_NOT_FOUND
            )

        filename = vault_file.original_filename or vault_file.file.name.rsplit('/', 1)[-1]
        content_type, _ = mimetypes.guess_type(filename)

        return FileResponse(
            vault_file.file.open('rb'),
            as_attachment=True,
            filename=filename,
            content_type=content_type or 'application/octet-stream',
        )


class LenderShareVaultFileDownloadView(APIView):
    """
    Serves one 'supporting_document' vault file for a usable
    LenderShareLink — same 404-never-401 discipline as the other public
    lender-portal views above. Restricted to file_type=
    'supporting_document': a lender can never reach a 're_upload'
    internal-audit vault entry through this endpoint even by guessing its
    id, since that type is already excluded from LenderSharePortalView's
    own listing.
    URL: GET /api/ai/share/deal/<str:access_token>/vault/<int:file_id>/
    """
    permission_classes = [AllowAny]
    authentication_classes = []

    def get(self, request, access_token, file_id):
        import mimetypes
        from .models import DealVaultFile

        link = _get_usable_lender_link_or_none(access_token)
        if link is None:
            return Response({'detail': 'This link is no longer available.'}, status=status.HTTP_404_NOT_FOUND)

        vault_file = DealVaultFile.objects.filter(
            pk=file_id, project=link.project, file_type='supporting_document'
        ).first()
        if vault_file is None:
            return Response({'detail': 'File not found.'}, status=status.HTTP_404_NOT_FOUND)

        filename = vault_file.original_filename or vault_file.file.name.rsplit('/', 1)[-1]
        content_type, _ = mimetypes.guess_type(filename)

        return FileResponse(
            vault_file.file.open('rb'),
            as_attachment=True,
            filename=filename,
            content_type=content_type or 'application/octet-stream',
        )


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


def _get_override_alerts(user, limit=200):
    """
    The actual query + serialization behind OverrideAlertFeedView, factored
    out so the Pipeline Modules Phase 2 Leadership Digest (below) can pull
    the same firm-wide variance feed into its own payload — capped smaller,
    as a preview — without duplicating this logic or risking it drifting
    out of sync with the dedicated full-feed endpoint.
    """
    from .permissions import visible_projects_for
    from .models import ExcelOverrideDelta

    projects = visible_projects_for(user)

    deltas = (
        ExcelOverrideDelta.objects
        .filter(
            field_name__in=OVERRIDE_ALERT_TRACKED_FIELDS,
            variance_pct__isnull=False,
            variance_pct__gt=OVERRIDE_ALERT_THRESHOLD_PCT,
            upload__project__in=projects,
        )
        .select_related('upload', 'upload__uploaded_by', 'upload__project', 'upload__project__workspace')
        .order_by('-upload__uploaded_at')[:limit]
    )

    return [
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
    The actual query now lives in _get_override_alerts() above, shared
    with the Leadership Digest (Pipeline Modules Phase 2).
    URL: GET /api/ai/pipeline/override-alerts/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request):
        alerts = _get_override_alerts(request.user, limit=200)
        return Response({'alerts': alerts, 'threshold_pct': OVERRIDE_ALERT_THRESHOLD_PCT})


class TeamScorecardView(APIView):
    """
    V2 Module 4b (Analyst Speed & Accuracy Scorecard): per-analyst
    performance metrics for the Leadership Morning Digest's "Team
    Performance" tab — see team_scorecard.compute_analyst_scorecards
    for the 4 metrics (Deals Underwritten, Average Turnaround Time,
    >2.0% Post-Upload Variance Rate reusing OVERRIDE_ALERT_
    TRACKED_FIELDS/_THRESHOLD_PCT above verbatim, and Forensic
    Seller-Trap Flags Triggered).

    Gated to the same leadership tier as every other org-wide/
    workspace-wide view in this codebase (Org Admin or Workspace Head,
    or a platform superuser) — a plain analyst, or a member with no
    workspace, has no "team" to see a scorecard FOR, same reasoning as
    can_lock_project's own gate. Scoped through visible_projects_for
    exactly like OverrideAlertFeedView/the Leadership Digest above: an
    Org Admin's scorecard is firm-wide, a Workspace Head's is their own
    workspace only. No special case.
    URL: GET /api/ai/pipeline/team-scorecard/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request):
        from .permissions import visible_projects_for
        from .team_scorecard import compute_analyst_scorecards

        user = request.user
        is_leader = user.is_admin or user.role == 'admin' or (
            user.workspace_id is not None and user.workspace_role == 'head'
        )
        if not is_leader:
            return Response(
                {'detail': 'Only an Org Admin or Workspace Head can view the Team Scorecard.'},
                status=status.HTTP_403_FORBIDDEN,
            )

        projects = visible_projects_for(user)
        rows = compute_analyst_scorecards(projects, OVERRIDE_ALERT_TRACKED_FIELDS, OVERRIDE_ALERT_THRESHOLD_PCT)
        return Response({'analysts': rows, 'variance_threshold_pct': OVERRIDE_ALERT_THRESHOLD_PCT})


# ---------------------------------------------------------------------------
# Pipeline Modules Phase 2: VP Sign-Off Gate & Leadership Digest. The gate
# itself and the combined sign-off/lock action live above (see
# PropertyProjectDetailView.patch's gate check and VPSignOffView). What
# follows is the second half: a "Leadership Morning Digest" leadership (or,
# per the "no special case" philosophy above, anyone — it just narrows to
# their own deals via visible_projects_for like every other pipeline view)
# can open any time, plus a reportlab-rendered PDF of the same data. Per the
# client's explicit decision, this phase ships BOTH: the in-app page now,
# and a PDF download button on it — no automatic scheduled email yet.
# ---------------------------------------------------------------------------

# A deal with no DealActivityLog entry in this many days (and not already
# Dead/Passed) shows up in the digest's "stale deals" section. Deliberately
# a plain module constant, not an org setting — v1 scope, same reasoning as
# OVERRIDE_ALERT_THRESHOLD_PCT above.
STALE_DEAL_DAYS = 7


def _get_leadership_digest_data(user):
    """
    Gathers every section of the Leadership Morning Digest in one place so
    LeadershipDigestView (JSON) and LeadershipDigestPDFView (PDF) render
    the exact same underlying data — never two independently-computed
    copies that could disagree with each other.

    Sections:
      - stage_counts: how many of this user's visible deals sit in each of
        the 5 pipeline stages right now.
      - awaiting_sign_off: visible deals sitting in 'vp_review' whose
        baseline isn't locked yet — i.e. still waiting on the VP Sign-Off
        Gate (see VPSignOffView's docstring).
      - stale_deals: visible deals (any stage except 'dead_passed') with no
        DealActivityLog entry in at least STALE_DEAL_DAYS days — "nobody's
        touched this deal in a while."
      - variance_alerts: a 10-row preview of the same firm-wide variance
        feed OverrideAlertFeedView exposes in full (see
        _get_override_alerts() above).

    Every section is scoped through visible_projects_for(user) — the exact
    same Org -> Workspace -> Analyst/Head hierarchy as every other pipeline
    view, so a plain analyst's "Leadership Digest" is really just their own
    morning digest, and an Org Admin's is genuinely firm-wide. No special
    case, same philosophy as OverrideAlertFeedView.
    """
    from django.db.models import Max
    from collections import Counter
    from .permissions import visible_projects_for
    from .models import PropertyProject, DealActivityLog

    now = timezone.now()
    projects = list(
        visible_projects_for(user).select_related('workspace', 'assigned_analyst', 'baseline_lock')
    )

    stage_counts_raw = Counter(p.pipeline_stage for p in projects)
    stage_counts = {
        stage: stage_counts_raw.get(stage, 0) for stage, _ in PropertyProject.PIPELINE_STAGE_CHOICES
    }

    last_activity_by_project = dict(
        DealActivityLog.objects
        .filter(project__in=projects)
        .values('project_id')
        .annotate(last_at=Max('created_at'))
        .values_list('project_id', 'last_at')
    )

    awaiting_sign_off = []
    stale_deals = []

    for p in projects:
        # Every project gets a "created" DealActivityLog row going
        # forward (see document_history.record_analysis_report), but a
        # deal created before Pipeline Modules Phase 1 shipped may have
        # none at all yet — fall back to created_at for those.
        last_activity = last_activity_by_project.get(p.id) or p.created_at
        days_since_activity = (now - last_activity).days

        lock = getattr(p, 'baseline_lock', None)
        if p.pipeline_stage == 'vp_review' and lock is None:
            awaiting_sign_off.append({
                'id': p.id,
                'property_name': p.property_name or 'Untitled',
                'workspace': p.workspace.name if p.workspace_id else None,
                'assigned_analyst': p.assigned_analyst.email if p.assigned_analyst_id else None,
                'days_in_vp_review': days_since_activity,
                'target_closing_or_offer_date': (
                    p.target_closing_or_offer_date.isoformat() if p.target_closing_or_offer_date else None
                ),
            })

        if p.pipeline_stage != 'dead_passed' and days_since_activity >= STALE_DEAL_DAYS:
            stale_deals.append({
                'id': p.id,
                'property_name': p.property_name or 'Untitled',
                'workspace': p.workspace.name if p.workspace_id else None,
                'pipeline_stage': p.pipeline_stage,
                'pipeline_stage_display': p.get_pipeline_stage_display(),
                'days_since_activity': days_since_activity,
            })

    awaiting_sign_off.sort(key=lambda d: d['days_in_vp_review'], reverse=True)
    stale_deals.sort(key=lambda d: d['days_since_activity'], reverse=True)

    return {
        'generated_at': now.isoformat(),
        'stage_counts': stage_counts,
        'awaiting_sign_off': awaiting_sign_off,
        # Capped — a dashboard/PDF digest, not a full export (same
        # "preview, not export" reasoning as _get_override_alerts' own
        # 200-row cap, just tighter since this renders in full on the page).
        'stale_deals': stale_deals[:50],
        'variance_alerts': _get_override_alerts(user, limit=10),
        'stale_deal_threshold_days': STALE_DEAL_DAYS,
    }


class LeadershipDigestView(APIView):
    """
    Pipeline Modules Phase 2: the Leadership Morning Digest's in-app page —
    see _get_leadership_digest_data()'s own docstring for exactly what it
    returns and how visibility is scoped.
    URL: GET /api/ai/pipeline/leadership-digest/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request):
        return Response(_get_leadership_digest_data(request.user))


class LeadershipDigestPDFView(APIView):
    """
    Pipeline Modules Phase 2: "Download as PDF" on the Leadership Digest
    page — renders the EXACT same data _get_leadership_digest_data()
    returns to the JSON endpoint above, via reportlab (see that library's
    own choice reasoning in requirements.txt: this server's cPanel hosting
    has no Cairo/Pango for WeasyPrint). Never re-queries independently, so
    the PDF can never show different numbers than the page it was
    downloaded from.
    URL: GET /api/ai/pipeline/leadership-digest/pdf/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request):
        from .leadership_digest_pdf import render_leadership_digest_pdf

        digest_data = _get_leadership_digest_data(request.user)
        pdf_bytes = render_leadership_digest_pdf(digest_data, generated_for_email=request.user.email)

        filename = f"Leadership_Digest_{timezone.now():%Y-%m-%d}.pdf"
        return FileResponse(
            io.BytesIO(pdf_bytes),
            as_attachment=True,
            filename=filename,
            content_type='application/pdf',
        )


# ---------------------------------------------------------------------------
# Pipeline Modules Phase 7: Email-to-Deal Intake (Postmark Inbound)
# ---------------------------------------------------------------------------
# An analyst (or a broker, directly) sends/forwards an email with the OM/T12/
# Rent Roll attached to a workspace's own deals-<slug>@... address (see
# TeamWorkspace.inbound_email_address in users/models.py). Postmark's Inbound
# product receives it and POSTs the parsed message to
# InboundEmailWebhookView below, which turns it into a brand-new
# PropertyProject and enqueues the same background extraction pipeline a
# manual dashboard upload uses. See that view's own docstring for the full
# design (authentication, idempotency, rejection taxonomy).

# Postmark's own inbound limits are 10 attachments / 10MB total per message;
# these are defensive backstops that mirror those numbers rather than new
# arbitrary limits of our own.
MAX_INBOUND_ATTACHMENTS = 10
MAX_INBOUND_ATTACHMENT_BYTES = 25 * 1024 * 1024

# Matches an optional "Target: $X" / "Price: $X" tag anywhere in the
# subject or body - e.g. "Target: $5.2M", "Price: 4,500,000", "target:$950k".
# Case-insensitive; the K/M suffix is a convenience this codebase's own
# analysts asked for informally, not something the original spec wording
# required, but it costs nothing extra to support.
_ASKING_PRICE_TAG_RE = re.compile(
    r'(?:target|price)\s*:\s*\$?\s*([\d,]+(?:\.\d+)?)\s*([km])?\b', re.IGNORECASE
)


def _parse_asking_price_tag(text):
    """
    Returns the first Target:/Price: tag value found in `text` as a float,
    or None if no tag is present (or it fails to parse) - None leaves
    asking_price entirely to run_inbound_email_underwriting_task's own
    AI-extracted-OM-price fallback (see that task's docstring in tasks.py).
    """
    if not text:
        return None
    match = _ASKING_PRICE_TAG_RE.search(text)
    if not match:
        return None
    try:
        value = float(match.group(1).replace(',', ''))
    except (TypeError, ValueError):
        return None
    suffix = (match.group(2) or '').lower()
    if suffix == 'k':
        value *= 1_000
    elif suffix == 'm':
        value *= 1_000_000
    return value if value > 0 else None


def _check_postmark_basic_auth(request) -> bool:
    """
    The ONLY authenticity check this webhook has - Postmark Inbound doesn't
    sign each payload cryptographically; its own recommended verification
    is HTTP Basic Auth baked into the webhook URL registered with Postmark
    itself (see POSTMARK_INBOUND_WEBHOOK_USERNAME/PASSWORD's own docstrings
    in settings.py). Both unset (the dev-friendly default) SKIPS this check
    entirely - returns True unconditionally - which is only safe before
    this URL is ever registered live with Postmark.
    """
    expected_user = django_settings.POSTMARK_INBOUND_WEBHOOK_USERNAME
    expected_pass = django_settings.POSTMARK_INBOUND_WEBHOOK_PASSWORD
    if not expected_user and not expected_pass:
        return True

    auth_header = request.META.get('HTTP_AUTHORIZATION', '')
    if not auth_header.startswith('Basic '):
        return False
    try:
        decoded = base64.b64decode(auth_header[len('Basic '):]).decode('utf-8')
        sent_user, _, sent_pass = decoded.partition(':')
    except Exception:
        return False
    return sent_user == expected_user and sent_pass == expected_pass


def _iter_postmark_attachment_files(attachments):
    """
    Decodes Postmark's base64 Attachments list into (filename, raw_bytes)
    pairs. Expands any .zip attachment's contents inline as additional
    candidates (one level deep only) since brokers routinely zip an
    OM+T12+Rent Roll bundle together rather than attaching 3 separate
    files. Caps total candidates at MAX_INBOUND_ATTACHMENTS as a defensive
    bound against a hostile/oversized zip; anything past the cap (or any
    single file over MAX_INBOUND_ATTACHMENT_BYTES) is silently dropped
    rather than erroring the whole email out over it.
    """
    candidates = []
    for att in (attachments or []):
        if len(candidates) >= MAX_INBOUND_ATTACHMENTS:
            break
        name = att.get('Name') or 'attachment'
        content_b64 = att.get('Content')
        if not content_b64:
            continue
        try:
            raw = base64.b64decode(content_b64)
        except Exception:
            continue
        if not raw or len(raw) > MAX_INBOUND_ATTACHMENT_BYTES:
            continue

        if name.lower().endswith('.zip'):
            try:
                with zipfile.ZipFile(io.BytesIO(raw)) as zf:
                    for zip_name in zf.namelist():
                        if len(candidates) >= MAX_INBOUND_ATTACHMENTS:
                            break
                        if zip_name.endswith('/'):
                            continue
                        try:
                            zip_bytes = zf.read(zip_name)
                        except Exception:
                            continue
                        if not zip_bytes or len(zip_bytes) > MAX_INBOUND_ATTACHMENT_BYTES:
                            continue
                        candidates.append((os.path.basename(zip_name), zip_bytes))
            except zipfile.BadZipFile:
                # Not actually a valid zip despite the extension - fall
                # through and try it as a regular file instead.
                candidates.append((name, raw))
        else:
            candidates.append((name, raw))
    return candidates


def _classify_inbound_attachments(attachments):
    """
    Runs every decoded attachment candidate through the same
    classify_document() keyword scan ClassifyDocumentView uses, routing
    each into the om/t12/rent_roll slot it scores highest for - FIRST
    MATCH WINS per slot (there's no analyst standing by at inbox time to
    pick between two Rent Rolls the way the dashboard's drop zone lets one
    do interactively). Returns (om, t12, rent_roll) as (ContentFile or
    None) x3. UNCLASSIFIED and unreadable attachments are silently
    skipped - this is a best-effort automatic sort, not a hard validation
    gate (the "no usable attachments at all" check happens one level up,
    in the view itself).
    """
    from django.core.files.base import ContentFile

    slots = {'om': None, 't12': None, 'rent_roll': None}
    type_to_slot = {
        DOCUMENT_TYPE_OFFERING_MEMORANDUM: 'om',
        DOCUMENT_TYPE_TRAILING_12: 't12',
        DOCUMENT_TYPE_RENT_ROLL: 'rent_roll',
    }

    for filename, raw_bytes in _iter_postmark_attachment_files(attachments):
        ext = os.path.splitext(filename)[1].lower()
        if ext not in CLASSIFIABLE_EXTENSIONS:
            continue

        wrapped = ContentFile(raw_bytes, name=filename)
        try:
            result = classify_document(wrapped)
        except DocumentClassificationError:
            continue

        slot = type_to_slot.get(result.get('document_type'))
        if slot and slots[slot] is None:
            wrapped.seek(0)
            slots[slot] = wrapped

    return slots['om'], slots['t12'], slots['rent_roll']


def _resolve_inbound_workspace(to_full):
    """
    Resolves the recipient address (Postmark's ToFull[0] - the actual
    envelope recipient this inbound stream is configured against, not
    whatever else may be cc'd) against TeamWorkspace.inbound_email_slug.
    Every address this feature hands out has the literal local-part
    'deals-<slug>' (see TeamWorkspace.inbound_email_address in
    users/models.py) - anything that doesn't start with that exact prefix
    can't be one of ours. Returns the TeamWorkspace, or None.
    """
    from users.models import TeamWorkspace

    if not to_full:
        return None
    recipient = (to_full[0].get('Email') or '').strip().lower()
    local_part = recipient.split('@', 1)[0]
    if not local_part.startswith('deals-'):
        return None
    slug = local_part[len('deals-'):]
    if not slug:
        return None
    return TeamWorkspace.objects.filter(inbound_email_slug=slug).first()


def _resolve_inbound_sender(from_email, workspace):
    """
    Verifies the sender is an ACTIVE CustomUser who is a MEMBER OF THIS
    EXACT WORKSPACE (CustomUser.workspace_id == workspace.id) - a
    deliberate, literal reading of the spec, not merely "same
    organization". Worth flagging as a real consequence: an org admin who
    oversees several workspaces but isn't personally a member of THIS one
    can't submit deals through its address unless they're also added as a
    member - the same membership model CustomUser.workspace already
    enforces everywhere else in this codebase. Returns the CustomUser, or
    None.
    """
    from django.contrib.auth import get_user_model

    if not from_email or workspace is None:
        return None
    User = get_user_model()
    return User.objects.filter(
        email__iexact=from_email, workspace_id=workspace.id, is_active=True,
    ).first()


class InboundEmailWebhookView(APIView):
    """
    Pipeline Modules Phase 7 (Multi-Tenant Email-to-Deal Intake): the
    server-to-server endpoint Postmark's Inbound stream POSTs to every
    time an email arrives at one of this app's deals-<slug>@... addresses
    (see TeamWorkspace.inbound_email_address in users/models.py). An
    analyst forwards (or a broker sends directly to) that address with the
    OM/T12/Rent Roll attached; this view resolves which workspace it
    belongs to, verifies the sender is actually a member of it, sorts the
    attachments into the right slots, creates a brand-new PropertyProject
    immediately (titled from the email's subject line, so it shows up in
    the pipeline before AI extraction even starts), and enqueues the exact
    same background extraction pipeline a manual dashboard upload uses
    (run_inbound_email_underwriting_task - see tasks.py for the two
    deliberate, scoped differences from a manual upload: the asking_price
    pre-fill/override and the placeholder-subject-to-AI-extracted-name
    upgrade).

    AUTHENTICATION: see _check_postmark_basic_auth above - the ONLY path
    in this view that ever returns 401. Every other rejection below
    (unknown workspace slug, unauthorized sender, no usable attachments,
    malformed payload) is a PERMANENT rejection Postmark should stop
    retrying, so each of those returns 200 instead - exactly what
    Postmark's own docs recommend for "don't retry this." A row is logged
    to InboundEmailEvent for every one of these outcomes, accepted or
    rejected (never for a 401, which never reaches any application logic
    at all).

    IDEMPOTENCY: Postmark retries a webhook delivery on anything other
    than a fast 200, so the very first thing this view does (after the
    Basic Auth gate, and after confirming the payload even has a
    MessageID) is check whether this exact MessageID has already been
    recorded in InboundEmailEvent.provider_message_id - a retried
    delivery of an already-accepted (or already-rejected) email returns
    200 immediately without doing anything a second time.

    URL: POST /api/webhooks/inbound-email/ (see config/urls.py; deliberately
    NOT under /api/ai/, since this is never called by this app's own
    frontend - only by Postmark, server to server).
    """
    permission_classes = [AllowAny]
    authentication_classes = []

    def post(self, request):
        from .models import InboundEmailEvent, PropertyProject, DealActivityLog, TeamNote

        if not _check_postmark_basic_auth(request):
            return Response(status=status.HTTP_401_UNAUTHORIZED)

        payload = request.data if isinstance(request.data, dict) else {}

        message_id = (payload.get('MessageID') or '').strip()
        subject = (payload.get('Subject') or '').strip() or '(no subject)'
        from_email = parseaddr(payload.get('From') or '')[1].strip().lower()
        text_body = payload.get('TextBody') or ''
        html_body = payload.get('HtmlBody') or ''
        to_full = payload.get('ToFull') or []
        attachments = payload.get('Attachments') or []

        if not message_id or not from_email:
            # Malformed beyond anything meaningful to dedupe against - a
            # missing MessageID means a retry of THIS rejection can't be
            # recognized as a retry, but that's fine: a payload missing
            # even these two fields is garbage hitting this URL, not a
            # real Postmark delivery that will be retried sensibly.
            InboundEmailEvent.objects.create(
                provider_message_id=message_id or f'malformed-{uuid.uuid4().hex}',
                sender_email=from_email,
                subject=subject,
                status='rejected_malformed',
            )
            return Response(status=status.HTTP_200_OK)

        if InboundEmailEvent.objects.filter(provider_message_id=message_id).exists():
            return Response(status=status.HTTP_200_OK)

        workspace = _resolve_inbound_workspace(to_full)
        if workspace is None:
            InboundEmailEvent.objects.create(
                provider_message_id=message_id,
                sender_email=from_email,
                subject=subject,
                status='rejected_unknown_slug',
            )
            return Response(status=status.HTTP_200_OK)

        sender = _resolve_inbound_sender(from_email, workspace)
        if sender is None:
            InboundEmailEvent.objects.create(
                provider_message_id=message_id,
                workspace=workspace,
                sender_email=from_email,
                subject=subject,
                status='rejected_unauthorized_sender',
            )
            return Response(status=status.HTTP_200_OK)

        organization = sender.organization
        if organization is None:
            # Can't happen in practice - a CustomUser with a workspace_id
            # always has an organization too - but fail gracefully rather
            # than crash on a None if that invariant is ever violated.
            InboundEmailEvent.objects.create(
                provider_message_id=message_id,
                workspace=workspace,
                sender_email=from_email,
                subject=subject,
                status='rejected_malformed',
            )
            return Response(status=status.HTTP_200_OK)

        om_file, t12_file, rent_roll_file = _classify_inbound_attachments(attachments)
        if not any([om_file, t12_file, rent_roll_file]):
            InboundEmailEvent.objects.create(
                provider_message_id=message_id,
                workspace=workspace,
                sender_email=from_email,
                subject=subject,
                status='rejected_no_usable_attachments',
            )
            return Response(status=status.HTTP_200_OK)

        # Target:/Price: tag override - subject first, then the plain-text
        # body, then (only if neither carried one) a crude tag-stripped
        # version of the HTML body. None (no tag found anywhere) leaves
        # asking_price entirely to run_inbound_email_underwriting_task's
        # own AI-extracted-OM-price fallback - see that task's docstring.
        explicit_asking_price = (
            _parse_asking_price_tag(subject)
            or _parse_asking_price_tag(text_body)
            or _parse_asking_price_tag(re.sub(r'<[^>]+>', ' ', html_body))
        )

        tier = organization.subscription_plan or TIER_BASIC

        project = PropertyProject.objects.create(
            organization=organization,
            workspace=workspace,
            property_name=subject,
            created_by=sender,
            asking_price=explicit_asking_price,
        )
        # Pipeline Modules Phase 1: seed the Activity feed with a "deal
        # created" row exactly like record_analysis_report does for every
        # other new-project path (see DealActivityLog's own docstring) -
        # this view creates the PropertyProject itself, up front, rather
        # than going through record_analysis_report, so that seeding has
        # to happen here instead.
        DealActivityLog.objects.create(
            project=project, user=sender, from_stage=None, to_stage=project.pipeline_stage,
        )
        # "Received via email" provenance surfaces in the Unified Activity
        # Feed as a normal TeamNote authored by the sender - reusing the
        # existing 'comment' event infrastructure (_get_deal_activity_feed)
        # rather than inventing a new event type for this one case.
        TeamNote.objects.create(
            project=project,
            author=sender,
            body=f'Deal received via email from {from_email} (subject: "{subject}").',
        )

        om_path, om_name = _save_upload_to_temp(om_file)
        t12_path, t12_name = _save_upload_to_temp(t12_file)
        rent_roll_path, rent_roll_name = _save_upload_to_temp(rent_roll_file)

        try:
            run_inbound_email_underwriting_task.delay(
                om_path, om_name, t12_path, t12_name, rent_roll_path, rent_roll_name,
                tier, sender.pk, project.pk, explicit_asking_price,
            )
        except Exception:
            # Enqueueing itself failed (e.g. Redis unreachable) - clean up
            # the temp files we just wrote, same reasoning as
            # UnderwritePropertyView's own enqueue failure path above. The
            # deal itself still exists (with its placeholder title/
            # attachments unprocessed) - an analyst can retry extraction
            # manually from inside it - so this is logged but NOT treated
            # as a rejection of the InboundEmailEvent below.
            logger.exception(
                "Inbound email intake: failed to enqueue extraction for "
                "project %s - deal was created but extraction never ran.",
                project.pk,
            )
            for p in (om_path, t12_path, rent_roll_path):
                if not p:
                    continue
                try:
                    os.remove(p)
                except OSError:
                    pass

        InboundEmailEvent.objects.create(
            provider_message_id=message_id,
            workspace=workspace,
            sender_email=from_email,
            subject=subject,
            status='accepted',
            project=project,
        )
        return Response(status=status.HTTP_200_OK)


# ---------------------------------------------------------------------------
# V2 Module 2 (Tax & Rent Comps Panel via RentCast): a "Market & Tax Comps"
# side panel the frontend renders next to the Live In-Browser Spreadsheet
# Editor (Pipeline Modules Phase 8) in the deal drawer. See
# rentcast_client.py for the raw RentCast HTTP layer, market_comps.py for
# the derived-numbers/variance-flag math, and PropertyMarketDataSnapshot in
# models.py for the per-deal cache this view reads from and writes to —
# that model's own docstring explains why caching here is mandatory (the
# $0-testing RentCast free tier is capped at 50 requests/month).
# ---------------------------------------------------------------------------

MARKET_COMPS_RENTCAST_UNCONFIGURED_DETAIL = (
    "Market & Tax Comps isn't turned on for this environment yet (no RentCast "
    "API key configured)."
)


class MarketCompsView(APIView):
    """
    GET: returns this deal's Market & Tax Comps data — tax assessment vs.
    the seller's stated T12 taxes, nearby active rent comps vs. the OM's
    claimed market rent, and an AVM value estimate vs. asking price (see
    market_comps.build_market_comps_payload for the exact shape).

    Fetches AND caches from RentCast on the FIRST call for a deal (no
    PropertyMarketDataSnapshot exists yet); every later call reads the
    cached snapshot only — no RentCast call at all — per the client's
    explicit "cache on the deal record the first time an address is
    checked" instruction. The variance flags themselves are still
    recomputed live on every call from the project's CURRENT metrics
    against that frozen snapshot, so a later Excel re-upload or grid
    edit is always compared against up-to-date seller-stated figures.

    `?refresh=true` forces a fresh RentCast lookup even when a snapshot
    already exists, overwriting it in place — gated to Org Admin/
    Workspace Head (can_lock_project's exact permission check) rather
    than open to every analyst, since every refresh spends real, capped
    RentCast quota and an address correction is exactly the kind of
    deliberate, occasional action a lead should gate, not something a
    plain analyst should be able to trigger by accident (or repeatedly,
    by mashing a button).

    URL: GET /api/ai/pipeline/projects/<int:project_id>/market-comps/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, project_id):
        from .permissions import visible_projects_for, can_lock_project
        from .models import PropertyMarketDataSnapshot
        from .market_comps import build_market_comps_payload
        from . import rentcast_client

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        analysis = project.latest_analysis
        if analysis is None:
            return Response({'detail': 'This deal has no analysis yet.'}, status=status.HTTP_404_NOT_FOUND)

        metrics = analysis.metrics or {}
        address = (metrics.get('property_metadata') or {}).get('address') or ''

        refresh_requested = str(request.query_params.get('refresh', '')).lower() in ('1', 'true', 'yes')
        snapshot = getattr(project, 'market_data_snapshot', None)

        if refresh_requested and not can_lock_project(request.user, project):
            return Response(
                {'detail': 'Only an Org Admin or Workspace Head can refresh Market & Tax Comps.'},
                status=status.HTTP_403_FORBIDDEN,
            )

        if snapshot is None or refresh_requested:
            if not address.strip():
                return Response({
                    'detail': (
                        'This deal has no extracted property address yet — Market & Tax '
                        'Comps needs one to look up.'
                    ),
                }, status=status.HTTP_400_BAD_REQUEST)

            if not rentcast_client.is_configured():
                return Response(
                    {'detail': MARKET_COMPS_RENTCAST_UNCONFIGURED_DETAIL}, status=status.HTTP_400_BAD_REQUEST
                )

            fetch_error = ''
            fetched = {'property_record': None, 'avm_value': None, 'rental_comps': []}
            try:
                fetched = rentcast_client.fetch_all_market_data(address)
            except rentcast_client.RentCastError as e:
                fetch_error = str(e)

            if snapshot is None:
                snapshot = PropertyMarketDataSnapshot(project=project)

            snapshot.address_checked = address
            snapshot.fetched_by = request.user
            snapshot.fetched_at = timezone.now()
            snapshot.raw_property_record = fetched.get('property_record')
            snapshot.raw_avm_value = fetched.get('avm_value')
            snapshot.raw_rental_comps = fetched.get('rental_comps')
            snapshot.fetch_error = fetch_error
            snapshot.save()

        payload = build_market_comps_payload(snapshot, metrics, project.asking_price)
        # Same "server computes the permission, frontend just reads a
        # boolean" pattern as PropertyProjectDetailView's own 'can_lock'
        # field — lets the Market & Tax Comps panel show/hide its
        # Refresh control without re-deriving can_lock_project's rule
        # client-side.
        payload['can_refresh'] = can_lock_project(request.user, project)
        return Response(payload)


# ---------------------------------------------------------------------------
# V2 Module 4a (Buy-Box Auto-Alerts): per-workspace match criteria settings
# (WorkspaceBuyBoxView — see TeamWorkspace.buy_box_criteria/
# get_buy_box_criteria in users/models.py and buy_box.py for the actual
# matching logic, run from document_history._check_buy_box_match_and_notify
# whenever a brand-new deal is created) plus the generic in-app
# notification feed the resulting alerts land in (Notification*View — see
# Notification in models.py for why this table's shape is deliberately
# generic, not buy-box-specific).
# ---------------------------------------------------------------------------

class WorkspaceBuyBoxView(APIView):
    """
    GET: this workspace's buy-box criteria (merged over defaults — see
    get_buy_box_criteria). Open to ANY authenticated member of the same
    organization, mirroring MyOrganizationDebtAssumptionsView's exact
    "anyone in the org can see the assumptions deals are measured
    against" reasoning in users/views.py — an analyst should be able to
    see why a deal they're working did or didn't get flagged as a
    buy-box match.

    PATCH: gated to this workspace's Org Admin OR Workspace Head only
    (mirrors can_lock_project's exact role check — a plain analyst, even
    one who created deals in this workspace, can never edit its buy-box
    criteria), per the client's explicit "editable by ORG_ADMIN /
    WORKSPACE_HEAD" scope decision.

    URL: GET/PATCH /api/ai/pipeline/workspaces/<int:workspace_id>/buy-box/
    """
    permission_classes = [IsAuthenticated]

    FIELD_SPECS = {
        'enabled': ('bool', None),
        'asset_classes': ('str_list', None),
        'target_cities': ('str_list', None),
        'target_states': ('str_list', None),
        'min_units': ('nonneg_float_or_null', None),
        'max_units': ('nonneg_float_or_null', None),
        'min_cap_rate_pct': ('nonneg_float_or_null', (0.0, 100.0)),
        'max_cap_rate_pct': ('nonneg_float_or_null', (0.0, 100.0)),
        'max_price_per_unit': ('nonneg_float_or_null', None),
    }

    def _get_workspace_or_error(self, request, workspace_id):
        from users.models import TeamWorkspace

        workspace = TeamWorkspace.objects.filter(pk=workspace_id).select_related('organization').first()
        if workspace is None:
            return None, Response({'detail': 'Workspace not found.'}, status=status.HTTP_404_NOT_FOUND)

        user = request.user
        same_org = user.is_admin or (user.organization_id == workspace.organization_id)
        if not same_org:
            return None, Response({'detail': 'Workspace not found.'}, status=status.HTTP_404_NOT_FOUND)

        return workspace, None

    def _can_edit(self, user, workspace) -> bool:
        if user.is_admin:
            return True
        if user.organization_id != workspace.organization_id:
            return False
        if user.role == 'admin':
            return True
        return user.workspace_id == workspace.id and user.workspace_role == 'head'

    def get(self, request, workspace_id):
        workspace, error = self._get_workspace_or_error(request, workspace_id)
        if error:
            return error
        data = workspace.get_buy_box_criteria()
        # Same "server computes the permission, frontend just reads a
        # boolean" pattern as PropertyProjectDetailView's 'can_lock' —
        # lets Workspace Settings show/hide the edit form per workspace
        # without re-deriving the Org Admin/Workspace Head rule itself.
        data['can_edit'] = self._can_edit(request.user, workspace)
        return Response(data)

    def patch(self, request, workspace_id):
        workspace, error = self._get_workspace_or_error(request, workspace_id)
        if error:
            return error

        if not self._can_edit(request.user, workspace):
            return Response(
                {'detail': 'Only an Org Admin or this Workspace\'s Head can edit its buy-box criteria.'},
                status=status.HTTP_403_FORBIDDEN,
            )

        current = workspace.get_buy_box_criteria()
        updated = dict(current)
        errors = {}

        for key, (kind, bounds) in self.FIELD_SPECS.items():
            if key not in request.data:
                continue
            raw = request.data.get(key)

            if kind == 'bool':
                if not isinstance(raw, bool):
                    errors[key] = 'Must be true or false.'
                    continue
                updated[key] = raw

            elif kind == 'str_list':
                if not isinstance(raw, list) or not all(isinstance(v, str) for v in raw):
                    errors[key] = 'Must be a list of strings.'
                    continue
                updated[key] = [v.strip() for v in raw if v.strip()]

            elif kind == 'nonneg_float_or_null':
                if raw is None:
                    updated[key] = None
                    continue
                try:
                    value = float(raw)
                except (TypeError, ValueError):
                    errors[key] = 'Must be a number, or null to clear it.'
                    continue
                if value < 0:
                    errors[key] = 'Must be zero or positive.'
                    continue
                if bounds and not (bounds[0] <= value <= bounds[1]):
                    errors[key] = f'Must be between {bounds[0]:.0f} and {bounds[1]:.0f}.'
                    continue
                updated[key] = value

        if errors:
            return Response({'detail': 'Invalid buy-box criteria.', 'errors': errors}, status=status.HTTP_400_BAD_REQUEST)

        workspace.buy_box_criteria = updated
        workspace.save(update_fields=['buy_box_criteria'])
        return Response(workspace.get_buy_box_criteria())


class NotificationListView(APIView):
    """
    GET: the current user's most recent notifications (buy-box matches
    today — see Notification's own docstring in models.py for why the
    shape is generic) plus their unread count, for the dashboard
    header's notification badge.
    URL: GET /api/ai/notifications/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request):
        from .models import Notification

        notifications = (
            Notification.objects.filter(recipient=request.user)
            .select_related('project')
            .order_by('-created_at')[:50]
        )
        unread_count = Notification.objects.filter(recipient=request.user, is_read=False).count()

        return Response({
            'notifications': [
                {
                    'id': n.id,
                    'notification_type': n.notification_type,
                    'project_id': n.project_id,
                    'property_name': n.project.property_name if n.project_id else None,
                    'message': n.message,
                    'is_read': n.is_read,
                    'created_at': n.created_at.isoformat(),
                }
                for n in notifications
            ],
            'unread_count': unread_count,
        })


class NotificationMarkReadView(APIView):
    """
    POST: marks one notification read (body: {"id": N}), or every
    unread notification for this user when no id is given — the header
    badge's "mark all read" action. Always scoped to
    recipient=request.user, so there's no way to mark (or probe the
    existence of) another user's notification by id.
    URL: POST /api/ai/notifications/mark-read/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request):
        from .models import Notification

        notif_id = request.data.get('id')
        qs = Notification.objects.filter(recipient=request.user, is_read=False)
        if notif_id is not None:
            qs = qs.filter(pk=notif_id)
        updated_count = qs.update(is_read=True)
        return Response({'updated': updated_count})


# ---------------------------------------------------------------------------
# V2 Module 1 (Two-Way Broker & Lender Email Thread Sync, Phase V2-3)
# ---------------------------------------------------------------------------
# Two layers, same split as the models (see models.py's own section-header
# comment above EmailOAuthConnection):
#   CONNECTION-level (one CustomUser's own mailbox — never project- or
#   org-scoped, per the client's own "at the CustomUser level" wording):
#     EmailConnectionListView, EmailConnectStartView, EmailOAuthCallbackView,
#     EmailConnectionDisconnectView.
#   THREAD-level (an explicit, per-deal "tag this thread" action — always
#   gated through visible_projects_for, the same visibility check every
#   other pipeline drawer panel uses):
#     EmailThreadSearchView, ProjectEmailThreadsView, ProjectEmailThreadUnlinkView.
# _get_deal_activity_feed() above already got its new 'email' event type —
# see that function's own docstring.

EMAIL_OAUTH_STATE_SALT = 'ai_service.email_oauth_state'


def _serialize_email_connection(conn):
    return {
        'id': conn.id,
        'provider': conn.provider,
        'provider_display': conn.get_provider_display(),
        'email_address': conn.email_address,
        'status': conn.status,
        'last_error': conn.last_error if conn.status == 'error' else '',
        'connected_at': conn.connected_at.isoformat(),
    }


class EmailConnectionListView(APIView):
    """
    GET: the CURRENT user's own connected mailboxes (at most one Gmail
    + one Microsoft 365 — see EmailOAuthConnection's own docstring) for
    the Settings page's "Connected Email Accounts" section. Deliberately
    request.user-scoped only, never org-wide — this is a per-analyst
    OAuth grant to THEIR OWN mailbox, not a shared workspace resource
    like an Excel Template.
    URL: GET /api/ai/email-sync/connections/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request):
        from . import email_oauth
        from .models import EmailOAuthConnection

        connections = EmailOAuthConnection.objects.filter(user=request.user)
        return Response({
            'connections': [_serialize_email_connection(c) for c in connections],
            'google_configured': email_oauth.is_google_configured(),
            'microsoft_configured': email_oauth.is_microsoft_configured(),
        })


class EmailConnectStartView(APIView):
    """
    POST (body: {"provider": "google"|"microsoft"}): returns
    {"authorization_url": "..."} for the frontend to do a full-page
    redirect to (not a popup — simplest, no popup-blocker edge cases,
    standard for a consent screen like this). The `state` param baked
    into that URL is a signed, TTL'd token (django.core.signing — no new
    dependency) carrying {user_id, provider}, NOT a Django session or
    this request's own auth cookie: the callback that follows is a
    plain top-level browser redirect FROM the provider back to this
    backend, and signed state is the standard, provider-agnostic way an
    OAuth client both prevents CSRF on that redirect AND recovers which
    of its OWN users started the flow — see EmailOAuthCallbackView
    below, which trusts this signature as its ONLY source of identity.
    URL: POST /api/ai/email-sync/connect/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request):
        from django.core import signing
        from . import email_oauth

        provider = (request.data.get('provider') or '').strip().lower()
        if provider not in email_oauth.PROVIDER_CHOICES:
            return Response({'detail': "provider must be 'google' or 'microsoft'."}, status=status.HTTP_400_BAD_REQUEST)
        if not email_oauth.is_configured(provider):
            return Response(
                {'detail': f"{provider.title()} email sync isn't configured on this server yet."},
                status=status.HTTP_400_BAD_REQUEST,
            )

        # request.user.pk, NOT .id - CustomUser's primary key field is
        # `email` (see users/models.py; USER_ID_FIELD in SIMPLE_JWT is
        # 'email' for the same reason), so there is no .id attribute at
        # all. .pk resolves to whatever the actual primary key field is,
        # which is the only attribute access on a user instance that's
        # ever safe to assume here.
        state = signing.dumps({'user_id': request.user.pk, 'provider': provider}, salt=EMAIL_OAUTH_STATE_SALT)
        if provider == email_oauth.GOOGLE:
            authorization_url = email_oauth.build_google_authorization_url(state)
        else:
            authorization_url = email_oauth.build_microsoft_authorization_url(state)

        return Response({'authorization_url': authorization_url})


class EmailOAuthCallbackView(APIView):
    """
    GET /api/ai/email-sync/oauth/callback/<provider>/ — exactly the path
    email_oauth._redirect_uri() builds, and what must be registered as
    this app's "Authorized redirect URI" with each provider. AllowAny:
    this is a plain browser redirect FROM Google/Microsoft, not an
    authenticated call from the frontend's JS — see EmailConnectStartView's
    own docstring on why the signed `state` param, not request.user, is
    this view's actual source of identity. Always ends in a 302 back to
    the frontend's Settings page (never a JSON error — there's no
    frontend JS request here to receive one), with a query param the
    Settings page reads once on load to show a toast.
    URL: GET /api/ai/email-sync/oauth/callback/google/
    URL: GET /api/ai/email-sync/oauth/callback/microsoft/
    """
    permission_classes = [AllowAny]

    def get(self, request, provider):
        from django.core import signing
        from django.shortcuts import redirect
        from django.contrib.auth import get_user_model
        from . import email_oauth
        from .models import EmailOAuthConnection

        settings_url = f"{django_settings.FRONTEND_BASE_URL.rstrip('/')}/dashboard/settings"

        if provider not in email_oauth.PROVIDER_CHOICES:
            return redirect(f"{settings_url}?email_sync_error=unknown_provider")

        provider_error = request.GET.get('error')
        if provider_error:
            logger.info("Email OAuth (%s): provider returned error=%s", provider, provider_error)
            return redirect(f"{settings_url}?email_sync_error=provider_denied")

        code = request.GET.get('code')
        state = request.GET.get('state')
        if not code or not state:
            return redirect(f"{settings_url}?email_sync_error=missing_code")

        try:
            # max_age matches the realistic time a user takes to click
            # through a consent screen — never the whole session.
            payload = signing.loads(state, salt=EMAIL_OAUTH_STATE_SALT, max_age=600)
        except signing.SignatureExpired:
            return redirect(f"{settings_url}?email_sync_error=expired_state")
        except signing.BadSignature:
            return redirect(f"{settings_url}?email_sync_error=bad_state")

        if payload.get('provider') != provider:
            return redirect(f"{settings_url}?email_sync_error=provider_mismatch")

        CustomUser = get_user_model()
        user = CustomUser.objects.filter(pk=payload.get('user_id')).first()
        if user is None:
            return redirect(f"{settings_url}?email_sync_error=unknown_user")

        try:
            if provider == email_oauth.GOOGLE:
                result = email_oauth.exchange_google_code(code)
            else:
                result = email_oauth.exchange_microsoft_code(code)
        except email_oauth.EmailOAuthError as e:
            logger.warning("Email OAuth (%s) code exchange failed for user %s: %s", provider, user.id, e)
            return redirect(f"{settings_url}?email_sync_error=exchange_failed")

        # unique_together(user, provider) — reconnecting the SAME
        # provider replaces the existing row rather than creating a
        # second one (see EmailOAuthConnection's own docstring).
        connection, _created = EmailOAuthConnection.objects.update_or_create(
            user=user, provider=provider,
            defaults={
                'email_address': result.get('email_address') or '',
                'scopes': result.get('scope', ''),
                'status': 'active',
                'last_error': '',
            },
        )
        connection.set_access_token(result['access_token'], result['expires_at'])
        if result.get('refresh_token'):
            connection.set_refresh_token(result['refresh_token'])
        elif not connection.encrypted_refresh_token:
            # Google omits refresh_token on a returning-consent exchange
            # ONLY if it never issued one before — access_type=offline +
            # prompt=consent on the authorization request (see
            # build_google_authorization_url) is specifically what
            # guarantees this doesn't happen on a fresh connect, but a
            # connection with no refresh token at all can never be
            # polled — surface that now rather than let the very first
            # poll attempt silently discover it.
            connection.status = 'error'
            connection.last_error = 'No refresh token was returned. Please reconnect and accept the consent screen.'
        connection.save()

        return redirect(f"{settings_url}?email_connected={provider}")


class EmailConnectionDisconnectView(APIView):
    """
    POST /api/ai/email-sync/connections/<provider>/disconnect/ — revokes
    (best-effort) and clears this user's OWN connection for one
    provider. Deliberately does NOT delete the EmailOAuthConnection row:
    LinkedEmailThread.connection is ON_DELETE=CASCADE, so hard-deleting
    it would wipe every linked thread AND its already-synced Activity
    Feed history for every deal it ever touched. Instead: tokens are
    blanked, status becomes 'disconnected', and every thread linked
    through this connection is soft-unlinked (is_active=False) exactly
    like a manual per-thread unlink — there's no access token left to
    poll them with, and their history stays in the Activity Feed either
    way. Reconnecting the same provider later (EmailOAuthCallbackView)
    reuses this same row rather than creating a new one.
    URL: POST /api/ai/email-sync/connections/<provider>/disconnect/
    """
    permission_classes = [IsAuthenticated]

    def post(self, request, provider):
        from . import email_oauth
        from .models import EmailOAuthConnection

        connection = EmailOAuthConnection.objects.filter(user=request.user, provider=provider).first()
        if connection is None:
            return Response({'detail': 'No connection to disconnect.'}, status=status.HTTP_404_NOT_FOUND)

        access_token = ''
        try:
            access_token = connection.get_access_token()
        except Exception:
            logger.warning("Could not decrypt access token for connection %s during disconnect — revoking skipped, local disconnect proceeds anyway.", connection.id)

        if access_token:
            try:
                if connection.provider == email_oauth.GOOGLE:
                    email_oauth.revoke_google(access_token)
                else:
                    email_oauth.revoke_microsoft(access_token)
            except Exception:
                logger.exception("Best-effort revoke failed for connection %s — disconnecting locally anyway.", connection.id)

        connection.encrypted_access_token = ''
        connection.encrypted_refresh_token = ''
        connection.access_token_expires_at = None
        connection.status = 'disconnected'
        connection.last_error = ''
        connection.save()

        connection.linked_threads.filter(is_active=True).update(is_active=False)

        return Response({'detail': 'Disconnected.'})


class EmailThreadSearchView(APIView):
    """
    GET /api/ai/email-sync/projects/<project_id>/thread-search/?q=...
    Searches every ACTIVE connection the CURRENT user owns (never other
    users' connections — this is "search MY mailbox for a thread to tag
    onto a deal I can see," never a shared org-wide inbox) and merges
    results across providers, newest-message-first, so the "Link Email
    Thread" picker shows one list regardless of how many mailboxes are
    connected. `project_id` is used ONLY for the visibility check (an
    analyst must be able to see the deal to search for a thread to tag
    onto it) — the search itself isn't project-scoped in any other way.
    A connection whose refresh token has died (ReauthRequired) is
    skipped from the results, flipped to status='error', and the user
    is notified — exactly like every other call site in this module —
    rather than failing the whole search over one broken connection.
    URL: GET /api/ai/email-sync/projects/<project_id>/thread-search/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, project_id):
        from .permissions import visible_projects_for
        from .models import EmailOAuthConnection
        from . import email_sync

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        query = (request.GET.get('q') or '').strip()
        connections = EmailOAuthConnection.objects.filter(user=request.user, status='active')

        results = []
        for connection in connections:
            try:
                threads = email_sync.search_candidate_threads(connection, query)
            except email_sync.ReauthRequired as e:
                email_sync.mark_connection_needs_reauth(connection, str(e))
                continue
            except email_sync.EmailSyncError as e:
                logger.warning("Email thread search failed for connection %s: %s", connection.id, e)
                continue
            for t in threads:
                t['connection_id'] = connection.id
                t['provider'] = connection.provider
                results.append(t)

        results.sort(key=lambda t: t.get('last_message_at') or '', reverse=True)
        return Response({'results': results})


def _serialize_linked_thread(thread_link):
    return {
        'id': thread_link.id,
        'provider': thread_link.connection.provider,
        'provider_display': thread_link.connection.get_provider_display(),
        'connection_email': thread_link.connection.email_address,
        'provider_thread_id': thread_link.provider_thread_id,
        'subject': thread_link.subject,
        'label': thread_link.label,
        'linked_by': thread_link.linked_by.email if thread_link.linked_by else None,
        'linked_at': thread_link.linked_at.isoformat(),
        'is_active': thread_link.is_active,
        'last_synced_at': thread_link.last_synced_at.isoformat() if thread_link.last_synced_at else None,
        'backfill_status': thread_link.backfill_status,
        'backfill_error': thread_link.backfill_error,
        'message_count': thread_link.messages.count(),
    }


class ProjectEmailThreadsView(APIView):
    """
    GET: every thread ever linked to this deal, active AND soft-
    unlinked — the "Linked Email Threads" panel shows unlinked ones too
    (dimmed), since their history is still in the Activity Feed.
    POST (body: {"connection_id", "provider_thread_id", "subject",
    "label"}): the ONE explicit "tag this thread to this deal" action —
    see LinkedEmailThread's own docstring on why this is never auto-
    matched. connection_id must be one of the CURRENT user's own
    connections (the thread was found by searching THEIR mailbox — see
    EmailThreadSearchView above); linking through someone else's
    connection_id 404s exactly like a project the user can't see.
    Re-linking a thread that's currently soft-unlinked on this SAME
    project+connection REACTIVATES that existing row instead of
    creating a second one (a second row would double-backfill and
    duplicate every message already in the Activity Feed); linking one
    that's already active is a 400, not a silent no-op, so the user
    actually sees why nothing happened.
    URL: GET/POST /api/ai/email-sync/projects/<project_id>/threads/
    """
    permission_classes = [IsAuthenticated]

    def get(self, request, project_id):
        from .permissions import visible_projects_for

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        threads = project.linked_email_threads.select_related('connection', 'linked_by').prefetch_related('messages')
        return Response([_serialize_linked_thread(t) for t in threads])

    def post(self, request, project_id):
        from .permissions import visible_projects_for
        from .models import EmailOAuthConnection, LinkedEmailThread
        from .tasks import run_email_thread_backfill_task

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        connection_id = request.data.get('connection_id')
        provider_thread_id = (request.data.get('provider_thread_id') or '').strip()
        subject = (request.data.get('subject') or '').strip()
        label = (request.data.get('label') or '').strip()

        if not provider_thread_id:
            return Response({'detail': 'provider_thread_id is required.'}, status=status.HTTP_400_BAD_REQUEST)

        connection = EmailOAuthConnection.objects.filter(pk=connection_id, user=request.user).first()
        if connection is None:
            return Response({'detail': 'That connection was not found for your account.'}, status=status.HTTP_404_NOT_FOUND)

        existing = LinkedEmailThread.objects.filter(
            project=project, connection=connection, provider_thread_id=provider_thread_id,
        ).first()

        if existing is not None and existing.is_active:
            return Response({'detail': 'This thread is already linked to this deal.'}, status=status.HTTP_400_BAD_REQUEST)

        if existing is not None:
            existing.is_active = True
            existing.subject = subject or existing.subject
            existing.label = label or existing.label
            existing.linked_by = request.user
            existing.save(update_fields=['is_active', 'subject', 'label', 'linked_by'])
            thread_link = existing
        else:
            thread_link = LinkedEmailThread.objects.create(
                project=project, connection=connection, provider_thread_id=provider_thread_id,
                subject=subject, label=label, linked_by=request.user,
            )

        run_email_thread_backfill_task.delay(thread_link.id)

        return Response(_serialize_linked_thread(thread_link), status=status.HTTP_201_CREATED)


class ProjectEmailThreadUnlinkView(APIView):
    """
    DELETE: soft-unlinks ONE thread (is_active=False) — see
    LinkedEmailThread's own docstring on why there's no hard-delete
    endpoint. Every message already synced through it stays exactly
    where it is in the Activity Feed; only the 5-minute poll stops
    touching this thread going forward.
    URL: DELETE /api/ai/email-sync/projects/<project_id>/threads/<thread_link_id>/
    """
    permission_classes = [IsAuthenticated]

    def delete(self, request, project_id, thread_link_id):
        from .permissions import visible_projects_for
        from .models import LinkedEmailThread

        project = visible_projects_for(request.user).filter(pk=project_id).first()
        if project is None:
            return Response({'detail': 'Project not found.'}, status=status.HTTP_404_NOT_FOUND)

        thread_link = LinkedEmailThread.objects.filter(pk=thread_link_id, project=project).first()
        if thread_link is None:
            return Response({'detail': 'Linked thread not found.'}, status=status.HTTP_404_NOT_FOUND)

        thread_link.is_active = False
        thread_link.save(update_fields=['is_active'])
        return Response({'detail': 'Unlinked.'})