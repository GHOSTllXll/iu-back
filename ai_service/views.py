# backend/ai_service/views.py
import io
import re
import json
import os
import pandas as pd
from django.http import FileResponse
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from rest_framework.permissions import IsAuthenticated

from .llm_router import ai_client, AIRateLimitExhausted
from .parsers import (
    extract_document_text,
    extract_t12_text,
    extract_data_from_excel,
    FileValidationError,
)
from .excel_generator import generate_underwriting_excel, ExcelGenerationError
from .rent_roll_utils import detect_rent_roll_columns, compute_ground_truth, RentRollColumnError
from .reconciliation import run_reconciliation
from .document_history import record_processed_documents


# ==========================================
# TIER CONFIGURATION
# ==========================================
TIER_BASIC = 'basic'
TIER_PROFESSIONAL = 'professional'
TIER_ENTERPRISE = 'enterprise'

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


# Sensible bounds — not enforcing "correctness", just catching obvious garbage
# input (negative rates, 500% LTV, etc.) before it reaches Excel formulas.
DEBT_ASSUMPTION_BOUNDS = {
    'ltv_pct': (0, 100),
    'interest_rate_pct': (0, 30),
    'amortization_years': (1, 50),
}

# Defaults if a field is somehow omitted (frontend pre-fills these, but don't
# trust the client — validate/default server-side too).
DEBT_ASSUMPTION_DEFAULTS = {
    'ltv_pct': 75.0,
    'interest_rate_pct': 7.0,
    'amortization_years': 30,
}


def _parse_debt_assumptions(request):
    """
    Reads ltv_pct, interest_rate_pct, amortization_years from the request body,
    validates they're sane numbers within reasonable bounds, and falls back to
    defaults for anything missing or invalid.
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
    try:
        # 1. Parse documents
        om_text = extract_document_text(om_file)
        t12_text = extract_t12_text(t12_file)
        rent_roll_text = extract_data_from_excel(rent_roll_file)

        # 2. Parse Excel into DataFrame
        rent_roll_df = pd.read_excel(rent_roll_file, sheet_name=0)
        rent_roll_df = rent_roll_df.dropna(how='all').dropna(axis=1, how='all')

        if rent_roll_df.empty:
            raise FileValidationError("Rent roll has no usable rows after cleaning.")

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
                rent_col, status_col = detect_rent_roll_columns(rent_roll_df)
                ground_truth = compute_ground_truth(rent_roll_df, rent_col, status_col)
                flags = run_reconciliation(metrics, rent_roll_df, ground_truth)
                metrics["reconciliation"] = {"flags": flags}
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

        record_processed_documents(organization, uploaded_by, om_file, t12_file, rent_roll_file, status='completed')

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

        return Response({'metrics': metrics, 'tier': tier})


class UnderwritePropertyDownloadView(APIView):
    """
    ENDPOINT 2: Returns Excel file download
    URL: /api/ai/underwrite/download/
    """
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

        # Financing assumptions — user-entered on the frontend, re-entered per
        # upload (org-level saved defaults were considered and deliberately not
        # built; see conversation history). These are raw numbers, not derived
        # from the AI extraction, since no source document states a buyer's
        # intended financing terms for a new acquisition.
        debt_assumptions, debt_error = _parse_debt_assumptions(request)
        if debt_error:
            return debt_error

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
    for the requesting user's organization. Files themselves are never stored
    — this is metadata only (name, type, status, size, timestamp).
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

        queryset = ProcessedDocument.objects.filter(organization=organization)

        search = request.query_params.get('search', '').strip()
        if search:
            queryset = queryset.filter(file_name__icontains=search)

        # Stats are computed from the FULL org history (not filtered by search),
        # so the stat cards reflect the org's overall state regardless of what
        # the user is currently searching for.
        all_docs = ProcessedDocument.objects.filter(organization=organization)
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
            for doc in queryset[:200]  # reasonable cap — pagination can be added later if needed
        ]

        return Response({'stats': stats, 'documents': documents})