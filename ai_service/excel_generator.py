# backend/ai_service/excel_generator.py
import io
import datetime
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

from .rent_roll_utils import detect_rent_roll_columns, ensure_status_column, RentRollColumnError


class ExcelGenerationError(Exception):
    """Raised when the rent roll doesn't have columns we can confidently map."""
    pass


# ==========================================
# MULTI-CURRENCY SUPPORT
# Maps a currency code to (a) the literal symbol used in a couple of plain-text
# summary strings, and (b) the openpyxl number_format code used on every
# currency cell in the workbook.
#
# NOTE on the number_format codes: these are intentionally simple
# quoted-literal formats (e.g. '"€"#,##0.00'), NOT the '[$€-407]'-style
# locale-code formats. A quoted literal renders identically in every Excel
# build/locale/OS with zero risk of a malformed format code — Excel treats
# anything inside "..." as plain text, full stop. The '[$SYMBOL-LCID]' syntax
# exists to make the *number* (decimal separator, digit grouping) follow a
# specific locale rather than the workbook's own locale, which isn't something
# this feature needs — we just want the right symbol next to a normally
# grouped number. Going this route also avoids depending on any specific LCID
# codes being exactly right, which is easy to get subtly wrong and, if wrong,
# is exactly the kind of thing that produces "unreadable formatting" or a
# repair prompt when a file is opened in certain Excel builds.
# ==========================================
CURRENCY_FORMATS = {
    'USD': {'symbol': '$', 'number_format': '"$"#,##0.00'},
    'EUR': {'symbol': '€', 'number_format': '"€"#,##0.00'},
    'GBP': {'symbol': '£', 'number_format': '"£"#,##0.00'},
    'CHF': {'symbol': 'CHF', 'number_format': '"CHF" #,##0.00'},
    'CAD': {'symbol': 'C$', 'number_format': '"C$"#,##0.00'},
    'AUD': {'symbol': 'A$', 'number_format': '"A$"#,##0.00'},
    'ZAR': {'symbol': 'R', 'number_format': '"R"#,##0.00'},
    # NOTE: uses standard thousands grouping (1,23,456.00 style lakh/crore
    # grouping is NOT applied here), for consistency with every other currency
    # in this map and to avoid a locale-specific format code. Flag if Indian
    # analysts specifically need lakh/crore grouping — that's a real, separate
    # ask worth doing deliberately, not something to slip in silently here.
    'INR': {'symbol': '₹', 'number_format': '"₹"#,##0.00'},
    'SGD': {'symbol': 'S$', 'number_format': '"S$"#,##0.00'},
}
DEFAULT_CURRENCY_CODE = 'USD'


def format_citation(citation: dict) -> str:
    """
    Module 4 (Source Provenance) — builds a human-readable citation string from
    a provenance dict, gracefully handling any field being null (the AI is
    instructed to use null rather than fabricate a citation it isn't sure of).
    Returns "" if there's nothing usable to cite.
    """
    if not citation:
        return ""

    parts = []
    if citation.get("page_location"):
        parts.append(f"Page {citation['page_location']}")
    if citation.get("section_title"):
        parts.append(f"'{citation['section_title']}'")

    location = ", ".join(parts)
    anchor = citation.get("exact_text_anchor")

    if location and anchor:
        return f"[Source: Offering Memorandum, {location} — \"{anchor}\"]"
    elif location:
        return f"[Source: Offering Memorandum, {location}]"
    elif anchor:
        return f"[Source: Offering Memorandum — \"{anchor}\"]"
    return ""


def clean_cell_value(val):
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (ValueError, TypeError):
        pass
    if isinstance(val, np.integer):
        return int(val)
    if isinstance(val, np.floating):
        return float(val)
    if isinstance(val, (pd.Timestamp, datetime.datetime)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, str):
        if val in ['#N/A', '#VALUE!', '#REF!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!']:
            return ""
        return val.strip()
    return val


def safe_get(d: dict, *keys, default=None):
    """Nested dict.get() that never raises, even if a section is missing."""
    for key in keys[:-1]:
        d = d.get(key, {}) if isinstance(d, dict) else {}
    return d.get(keys[-1], default) if isinstance(d, dict) else default


# ==========================================
# T12 LINE ITEM DEFINITIONS
# Maps metrics JSON keys -> display label -> row on the T12 tab.
# NOTE: This writes ANNUAL totals only, one column, because the AI extraction
# schema only produces annual figures per line item (see conversation: monthly
# breakdown was scoped out pending a decision on expanding the AI schema).
# ==========================================
T12_LINE_ITEMS = [
    ("net_rental_income", "Net Rental Income"),
    ("other_income", "Other Income"),
    ("total_operating_revenue", "TOTAL OPERATING REVENUE", True),   # subtotal row
    ("real_estate_taxes", "Real Estate Taxes"),
    ("property_liability_insurance", "Property & Liability Insurance"),
    ("total_utilities", "Total Utilities"),
    ("repairs_maintenance", "Repairs & Maintenance"),
    ("contract_services", "Contract Services"),
    ("marketing_advertising", "Marketing & Advertising"),
    ("property_management_fee", "Property Management Fee"),
    ("payroll_benefits_staffing", "Payroll, Benefits & Staffing"),
    ("general_administrative", "General & Administrative"),
    ("total_operating_expenses", "TOTAL OPERATING EXPENSES", True),  # subtotal row
    ("net_operating_income", "NET OPERATING INCOME (NOI)", True),    # subtotal row
]
def generate_underwriting_excel(metrics: dict, rent_roll_df: pd.DataFrame, debt_assumptions: dict = None, currency_code: str = None) -> bytes:
    wb = Workbook()

    # Multi-currency: purely a display/formatting concern — every numeric value
    # written to the sheet below is untouched (no conversion, no rate lookup).
    # This only changes which symbol + number_format string gets stamped onto
    # currency cells. Invalid/unknown codes fall back to USD rather than
    # raising, since this is a formatting nicety, not something that should be
    # able to break a download.
    currency_info = CURRENCY_FORMATS.get((currency_code or DEFAULT_CURRENCY_CODE).upper(), CURRENCY_FORMATS[DEFAULT_CURRENCY_CODE])
    currency_format = currency_info['number_format']
    currency_symbol = currency_info['symbol']

    # debt_assumptions: {'ltv_pct': 75.0, 'interest_rate_pct': 7.0, 'amortization_years': 30}
    # User-entered per upload — see conversation history for why these can't be
    # AI-extracted (they're the buyer's own financing assumptions, not facts
    # stated in the source documents). Falls back to safe defaults if omitted
    # (e.g. if this function is ever called from a path that doesn't collect them).
    if debt_assumptions is None:
        debt_assumptions = {'ltv_pct': 75.0, 'interest_rate_pct': 7.0, 'amortization_years': 30}

    # ==========================================
    # PRE-PROCESSING: Rent Roll column detection
    # (Shared with reconciliation.py — see rent_roll_utils.py. Both the Excel
    # file and the reconciliation flags must agree on which columns are which.)
    # ==========================================
    try:
        rent_col, status_col, tenant_col = detect_rent_roll_columns(rent_roll_df)
        status_col, status_was_inferred = ensure_status_column(rent_roll_df, status_col, tenant_col)
    except RentRollColumnError as e:
        raise ExcelGenerationError(str(e))

    rent_roll_df[rent_col] = pd.to_numeric(rent_roll_df[rent_col], errors='coerce').fillna(0)
    rent_roll_df[status_col] = rent_roll_df[status_col].astype(str).str.strip()
    rent_roll_df = rent_roll_df.replace({pd.NaT: None, pd.NA: None, np.nan: None})

    # ==========================================
    # TAB 3: RENT ROLL ANALYSIS
    # Dynamic row range (NOT hardcoded to row 150 — a property with more units
    # than that would silently undercount with a fixed range).
    # Summary formulas live BELOW the last real data row, not overlapping it.
    # ==========================================
    ws_rr = wb.active
    ws_rr.title = "Rent Roll Analysis"

    for col_idx, col_name in enumerate(rent_roll_df.columns, 1):
        cell = ws_rr.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)
        if status_was_inferred and col_name == status_col:
            cell.comment = Comment(
                "This rent roll had no explicit Status column. Occupancy was "
                "INFERRED from the Tenant column: a populated tenant name = "
                "Occupied, blank/vacant = Vacant. Verify against the source "
                "rent roll before relying on this.",
                "Underwriting AI"
            )

    rent_col_idx = rent_roll_df.columns.get_loc(rent_col) + 1

    for row_idx, row in enumerate(rent_roll_df.itertuples(index=False), 2):
        for col_idx, val in enumerate(row, 1):
            cleaned_val = clean_cell_value(val)
            cell = ws_rr.cell(row=row_idx, column=col_idx, value=cleaned_val)
            if col_idx == rent_col_idx:
                cell.number_format = currency_format
                if isinstance(cleaned_val, str) and cleaned_val.replace('.', '').replace('-', '').isdigit():
                    cell.value = float(cleaned_val)
    last_row = len(rent_roll_df) + 1
    status_letter = get_column_letter(rent_roll_df.columns.get_loc(status_col) + 1)
    rent_letter = get_column_letter(rent_col_idx)

    unit_count_row = last_row + 2
    vacant_units_row = last_row + 3
    occupancy_row = last_row + 4
    avg_rent_row = last_row + 5

    ws_rr.cell(row=unit_count_row, column=1, value="TOTAL UNIT COUNT:").font = Font(bold=True)
    ws_rr.cell(row=unit_count_row, column=2, value=f"=COUNTA(A2:A{last_row})")

    ws_rr.cell(row=vacant_units_row, column=1, value="TOTAL VACANT UNITS:").font = Font(bold=True)
    ws_rr.cell(row=vacant_units_row, column=2, value=f'=COUNTIF({status_letter}2:{status_letter}{last_row}, "Vacant")')

    ws_rr.cell(row=occupancy_row, column=1, value="PHYSICAL OCCUPANCY %:").font = Font(bold=True)
    ws_rr.cell(row=occupancy_row, column=2,
               value=f'=COUNTIF({status_letter}2:{status_letter}{last_row}, "Occupied") / COUNTA(A2:A{last_row})')
    ws_rr.cell(row=occupancy_row, column=2).number_format = '0.00%'

    ws_rr.cell(row=avg_rent_row, column=1, value="AVG IN-PLACE RENT:").font = Font(bold=True)
    ws_rr.cell(row=avg_rent_row, column=2, value=f"=AVERAGE({rent_letter}2:{rent_letter}{last_row})")
    ws_rr.cell(row=avg_rent_row, column=2).number_format = currency_format

    # ==========================================
    # TAB 2: CLEANED T12
    # Real extracted annual figures — NOT hardcoded placeholder data.
    # (Previous version wrote a fixed $15,000/month regardless of actual metrics —
    # fixed here.)
    # ==========================================
    ws_t12 = wb.create_sheet("Cleaned T12")
    ws_t12.cell(row=1, column=1, value="LINE ITEM")
    ws_t12.cell(row=1, column=2, value="ANNUAL TOTAL")
    for cell in ws_t12[1]:
        cell.font = Font(bold=True, color="FFFFFFFF")
        cell.fill = PatternFill(start_color="FF333333", end_color="FF333333", fill_type="solid")

    t12_data = metrics.get("t12_revenue_expenses", {})
    line_item_rows = {}  # key -> row number, so we can reference rows for subtotals

    row_num = 2
    for item in T12_LINE_ITEMS:
        key, label = item[0], item[1]
        is_subtotal = len(item) > 2 and item[2]

        value = t12_data.get(key)
        cell_label = ws_t12.cell(row=row_num, column=1, value=label)
        cell_value = ws_t12.cell(row=row_num, column=2)

        if is_subtotal:
            cell_label.font = Font(bold=True)
            cell_value.font = Font(bold=True)
            # Subtotals are written as real values from the AI extraction, not
            # re-derived via SUM() here, since the AI's total may legitimately
            # differ slightly from a naive sum if it applied its own judgment
            # (e.g. rounding, or excluded a line item as non-recurring).
            # If you want the sheet to self-recalculate subtotals from the line
            # items above instead of trusting the AI's stated total, that's a
            # one-line change — say so and I'll switch it to a SUM() formula.

        cell_value.value = value if value is not None else ""
        if isinstance(value, (int, float)):
            cell_value.number_format = currency_format

        line_item_rows[key] = row_num
        row_num += 1
    # ==========================================
    # MODULE 2 (Enterprise only): Uncategorized Expenses
    # Per spec: a dedicated, visible row for any T12 line item the AI couldn't
    # confidently map into a standard category, so the model's total still
    # balances against the raw source document instead of silently dropping
    # or misfiling anything.
    # ==========================================
    standardization = metrics.get("standardization")
    if standardization and standardization.get("uncategorized_items"):
        uncategorized_header_row = row_num + 1
        ws_t12.cell(row=uncategorized_header_row, column=1, value="UNCATEGORIZED DISCREPANCIES").font = Font(bold=True, size=12, color="FFB8860B")

        item_row = uncategorized_header_row + 1
        for item in standardization["uncategorized_items"]:
            label = f"{item['original_line_item']} ({item['category']})"
            label_cell = ws_t12.cell(row=item_row, column=1, value=label)
            label_cell.comment = Comment(
                "AI could not confidently map this line item to a standard "
                "category. Included in totals as-is — please manually reallocate "
                "if a more specific category applies.",
                "Standardization Engine"
            )
            amount_cell = ws_t12.cell(row=item_row, column=2, value=item["allocated_amount"])
            amount_cell.number_format = currency_format
            item_row += 1

        total_uncategorized_row = item_row
        ws_t12.cell(row=total_uncategorized_row, column=1, value="TOTAL UNCATEGORIZED").font = Font(bold=True)
        total_cell = ws_t12.cell(row=total_uncategorized_row, column=2, value=standardization["uncategorized_total"])
        total_cell.font = Font(bold=True)
        total_cell.number_format = currency_format

        row_num = total_uncategorized_row + 1
    # ==========================================
    # TAB 2 continued: Capital Structure
    # Only builds what current data actually supports: Purchase Price + DST/Capex.
    # Loan sizing / debt service / DSCR are DEFERRED — see conversation: those
    # need LTV%, Interest Rate%, and Amortization Term inputs that don't exist
    # anywhere in this system yet (not extracted, not user-entered).
    # ==========================================
    capital_header_row = row_num + 1
    ws_t12.cell(row=capital_header_row, column=1, value="CAPITAL STRUCTURE").font = Font(bold=True, size=12)

    deal_valuation = metrics.get("deal_valuation", {})
    purchase_price = debt_assumptions.get('purchase_price') 
    dst_capex = deal_valuation.get("dst_capex_budget")  # None for Basic tier — already stripped upstream

    purchase_price_row = capital_header_row + 1
    ws_t12.cell(row=purchase_price_row, column=1, value="Target Purchase Price")
    pp_cell = ws_t12.cell(row=purchase_price_row, column=2, value=purchase_price)
    pp_cell.number_format = currency_format
    pp_cell.comment = Comment(
        "Confirmed purchase price used for this analysis — either extracted "
        "from the Offering Memorandum, or manually entered/adjusted before "
        "download if the source documents didn't state one.",
        "Underwriting AI"
    )

    total_acq_cost_row = purchase_price_row + 1

    if dst_capex is not None:
        dst_row = purchase_price_row + 1
        ws_t12.cell(row=dst_row, column=1, value="DST / Capex Budget")
        dst_cell = ws_t12.cell(row=dst_row, column=2, value=dst_capex)
        dst_cell.number_format = currency_format

        comment_text = (
            "Deferred maintenance, physical capital repairs, and/or sponsor reserves "
            "as stated in the Offering Memorandum. Extracted by AI — verify against "
            "source document before relying on this figure."
        )
        # Module 4: append the actual source citation if the AI provided one
        citation_text = format_citation(metrics.get("provenance", {}).get("dst_capex_budget", {}))
        if citation_text:
            comment_text += f"\n\n{citation_text}"

        dst_cell.comment = Comment(comment_text, "Underwriting AI")
        total_acq_cost_row = dst_row + 1
        ws_t12.cell(row=total_acq_cost_row, column=1, value="TOTAL ACQUISITION COST").font = Font(bold=True)
        acq_cell = ws_t12.cell(
            row=total_acq_cost_row, column=2,
            value=f"=B{purchase_price_row}+B{dst_row}"
        )
        acq_cell.font = Font(bold=True)
        acq_cell.number_format = currency_format
    else:
        # No DST/Capex (Basic tier, or AI found none) — Total Acquisition Cost is
        # just the purchase price, still expressed as a formula for consistency.
        ws_t12.cell(row=total_acq_cost_row, column=1, value="TOTAL ACQUISITION COST").font = Font(bold=True)
        acq_cell = ws_t12.cell(row=total_acq_cost_row, column=2, value=f"=B{purchase_price_row}")
        acq_cell.font = Font(bold=True)
        acq_cell.number_format = currency_format
    # ==========================================
    # TAB 2 continued: Debt Sizing & Returns
    # LTV / Interest Rate / Amortization are raw user-entered numbers (not
    # formulas) — the underwriter's own assumptions, editable directly in the
    # spreadsheet afterward. Everything downstream (loan amount, debt service,
    # DSCR, cash-on-cash) is a live formula referencing these input cells.
    # ==========================================
    ltv_pct = debt_assumptions.get('ltv_pct', 75.0)
    interest_rate_pct = debt_assumptions.get('interest_rate_pct', 7.0)
    amortization_years = debt_assumptions.get('amortization_years', 30)

    debt_header_row = debt_note_row = total_acq_cost_row + 2
    ws_t12.cell(row=debt_header_row, column=1, value="DEBT ASSUMPTIONS & SIZING").font = Font(bold=True, size=12)

    ltv_row = debt_header_row + 1
    ws_t12.cell(row=ltv_row, column=1, value="Loan-to-Value (LTV) %")
    ltv_cell = ws_t12.cell(row=ltv_row, column=2, value=ltv_pct / 100)  # stored as a decimal for % formatting
    ltv_cell.number_format = '0.00%'
    ltv_cell.comment = Comment(
        "User-entered assumption — not extracted from source documents. "
        "Edit this cell to see the model recalculate.",
        "Underwriting AI"
    )

    interest_row = ltv_row + 1
    ws_t12.cell(row=interest_row, column=1, value="Interest Rate %")
    interest_cell = ws_t12.cell(row=interest_row, column=2, value=interest_rate_pct / 100)
    interest_cell.number_format = '0.000%'
    interest_cell.comment = Comment(
        "User-entered assumption — not extracted from source documents.",
        "Underwriting AI"
    )

    amort_row = interest_row + 1
    ws_t12.cell(row=amort_row, column=1, value="Amortization Term (years)")
    amort_cell = ws_t12.cell(row=amort_row, column=2, value=amortization_years)
    amort_cell.comment = Comment(
        "User-entered assumption — not extracted from source documents.",
        "Underwriting AI"
    )

    loan_amount_row = amort_row + 1
    ws_t12.cell(row=loan_amount_row, column=1, value="Sized Loan Amount").font = Font(bold=True)
    loan_cell = ws_t12.cell(
        row=loan_amount_row, column=2,
        value=f"=B{total_acq_cost_row}*B{ltv_row}"
    )
    loan_cell.font = Font(bold=True)
    loan_cell.number_format = currency_format

    equity_row = loan_amount_row + 1
    ws_t12.cell(row=equity_row, column=1, value="Total Initial Cash Equity Required").font = Font(bold=True)
    equity_cell = ws_t12.cell(
        row=equity_row, column=2,
        value=f"=B{total_acq_cost_row}-B{loan_amount_row}"
    )
    equity_cell.font = Font(bold=True)
    equity_cell.number_format = currency_format
    debt_service_row = equity_row + 1
    ws_t12.cell(row=debt_service_row, column=1, value="Annual Debt Service")
    # PMT(rate/12, term*12, -loan) * 12 — standard amortizing loan payment,
    # negated because PMT returns a negative value for an outflow.
    debt_service_cell = ws_t12.cell(
        row=debt_service_row, column=2,
        value=f"=PMT(B{interest_row}/12, B{amort_row}*12, -B{loan_amount_row})*12"
    )
    debt_service_cell.number_format = currency_format

    dscr_row = debt_service_row + 1
    ws_t12.cell(row=dscr_row, column=1, value="Debt Service Coverage Ratio (DSCR)").font = Font(bold=True)
    noi_row_ref = line_item_rows.get("net_operating_income")
    if noi_row_ref:
        dscr_cell = ws_t12.cell(
            row=dscr_row, column=2,
            value=f"=B{noi_row_ref}/B{debt_service_row}"
        )
        dscr_cell.font = Font(bold=True)
        dscr_cell.number_format = '0.00"x"'

    cash_on_cash_row = dscr_row + 1
    ws_t12.cell(row=cash_on_cash_row, column=1, value="Cash-on-Cash Return %").font = Font(bold=True)
    if noi_row_ref:
        coc_cell = ws_t12.cell(
            row=cash_on_cash_row, column=2,
            value=f"=(B{noi_row_ref}-B{debt_service_row})/B{equity_row}"
        )
        coc_cell.font = Font(bold=True)
        coc_cell.number_format = '0.00%'

    # ==========================================
    # TAB 1: DASHBOARD SUMMARY
    # Every numeric cell is a cross-tab reference — no hardcoded values.
    # ==========================================
    ws_dash = wb.create_sheet("Dashboard Summary", 0)
    ws_dash.cell(row=1, column=1, value="METRIC")
    ws_dash.cell(row=1, column=2, value="VALUE")
    for cell in ws_dash[1]:
        cell.font = Font(bold=True, size=12, color="FFD4AF37")

    property_metadata = metrics.get("property_metadata", {})
    underwriting_ratios = metrics.get("underwriting_ratios", {})
    debt_returns = metrics.get("debt_returns", {})

    noi_row = line_item_rows.get("net_operating_income")
    revenue_row = line_item_rows.get("total_operating_revenue")
    opex_row = line_item_rows.get("total_operating_expenses")
    dash_rows = [
        ("Property Name", property_metadata.get("property_name"), None, False, None, None),
        ("Address", property_metadata.get("address"), None, False, None, None),
        ("Asset Class", property_metadata.get("asset_class"), None, False, None, None),
        ("Total Unit Count", None, f"='Rent Roll Analysis'!B{unit_count_row}", False, None, None),
        ("Total Vacant Units", None, f"='Rent Roll Analysis'!B{vacant_units_row}", False, None, None),
        ("Physical Occupancy %", None, f"='Rent Roll Analysis'!B{occupancy_row}", True, None, "physical_occupancy_pct"),
        ("Avg In-Place Monthly Rent", None, f"='Rent Roll Analysis'!B{avg_rent_row}", False, '$', None),
        ("Total Annual Concessions", metrics.get("rent_roll_metrics", {}).get("total_annual_concessions"), None, False, '$', "total_annual_concessions"),
        ("Total Operating Revenue", None, f"='Cleaned T12'!B{revenue_row}" if revenue_row else None, False, '$', None),
        ("Total Operating Expenses", None, f"='Cleaned T12'!B{opex_row}" if opex_row else None, False, '$', None),
        ("Net Operating Income (NOI)", None, f"='Cleaned T12'!B{noi_row}" if noi_row else None, False, '$', None),
        ("Net Rental Income", None, f"='Cleaned T12'!B{line_item_rows.get('net_rental_income')}" if line_item_rows.get('net_rental_income') else None, False, '$', "net_rental_income"),
        ("Operating Expense Ratio", underwriting_ratios.get("operating_expense_ratio"), None, True, None, None),
        ("Total Annual Expenses / Unit", underwriting_ratios.get("total_annual_expenses_per_unit"), None, False, '$', None),
        ("Target Purchase Price", None, f"='Cleaned T12'!B{purchase_price_row}", False, '$', None),
        ("Entry Cap Rate %", deal_valuation.get("entry_cap_rate"), None, True, None, None),
        ("Acquisition Price / Unit", deal_valuation.get("acquisition_price_per_unit"), None, False, '$', None),
        ("Cash-on-Cash Return %", debt_returns.get("cash_on_cash_return_pct"), None, True, None, None),
    ]

    if dst_capex is not None:
        dash_rows.append(("DST / Capex Budget", None, f"='Cleaned T12'!B{total_acq_cost_row - 1}", False, '$', None))

    dash_rows.append(("Sized Loan Amount", None, f"='Cleaned T12'!B{loan_amount_row}", False, '$', None))
    dash_rows.append(("Total Equity Required", None, f"='Cleaned T12'!B{equity_row}", False, '$', None))
    dash_rows.append(("Annual Debt Service", None, f"='Cleaned T12'!B{debt_service_row}", False, '$', None))
    if noi_row_ref:
        dash_rows.append(("DSCR", None, f"='Cleaned T12'!B{dscr_row}", False, None, None))
        dash_rows.append(("Cash-on-Cash Return % (Modeled)", None, f"='Cleaned T12'!B{cash_on_cash_row}", True, None, None))

    row_idx = 2
    dash_row_by_metric = {}  # metric key -> row number, for reconciliation highlighting below

    for entry in dash_rows:
        label, static_val, formula, is_pct, currency, metric_key = entry

        ws_dash.cell(row=row_idx, column=1, value=label)
        value_cell = ws_dash.cell(row=row_idx, column=2, value=formula if formula else static_val)

        if is_pct:
            value_cell.number_format = '0.00%'
        elif currency == '$':
            value_cell.number_format = currency_format

        if metric_key:
            dash_row_by_metric[metric_key] = row_idx

        row_idx += 1
    # ==========================================
    # MODULE 1 (Enterprise only): Reconciliation flag highlighting
    # If flags were computed (see reconciliation.py — Python compares OM claims
    # against ground truth from the actual rent roll data), highlight the
    # corresponding Dashboard cell in light red with a hoverable comment
    # explaining the discrepancy, per the module spec.
    # ==========================================
    reconciliation = metrics.get("reconciliation")
    if reconciliation and reconciliation.get("flags"):
        flag_fill = PatternFill(start_color="FFF8D7DA", end_color="FFF8D7DA", fill_type="solid")

        for flag in reconciliation["flags"]:
            target_row = dash_row_by_metric.get(flag["metric"])
            if target_row is None:
                continue  # flag references a metric not shown on this tab — skip highlighting, still listed below

            label_cell = ws_dash.cell(row=target_row, column=1)
            value_cell = ws_dash.cell(row=target_row, column=2)
            label_cell.fill = flag_fill
            value_cell.fill = flag_fill

            comment_text = f"[{flag['severity']}] {flag['message']}"
            # Module 4: the occupancy flag is specifically about the OM's claim —
            # attach its citation here if the AI provided one.
            if flag["metric"] == "physical_occupancy_pct":
                citation_text = format_citation(metrics.get("provenance", {}).get("om_claimed_occupancy_pct", {}))
                if citation_text:
                    comment_text += f"\n\n{citation_text}"

            value_cell.comment = Comment(comment_text, "Reconciliation Engine")

        # Full flags list below the main table too, so nothing is hidden behind
        # a hover-only comment — visible at a glance even without inspecting cells.
        flags_header_row = row_idx + 2
        ws_dash.cell(row=flags_header_row, column=1, value="RECONCILIATION FLAGS").font = Font(bold=True, size=12, color="FFD4AF37")

        flag_row = flags_header_row + 1
        for flag in reconciliation["flags"]:
            severity_cell = ws_dash.cell(row=flag_row, column=1, value=f"[{flag['severity']}]")
            severity_cell.font = Font(bold=True, color="FFCC0000" if flag['severity'] == "CRITICAL" else "FFB8860B")
            ws_dash.cell(row=flag_row, column=2, value=flag["message"])
            flag_row += 1
        next_free_row = flag_row
    elif reconciliation and reconciliation.get("note"):
        # Reconciliation couldn't run (e.g. rent roll columns unclear) — surface why.
        note_row = row_idx + 2
        ws_dash.cell(row=note_row, column=1, value="RECONCILIATION:").font = Font(bold=True, italic=True)
        ws_dash.cell(row=note_row, column=2, value=reconciliation["note"]).font = Font(italic=True, color="FF888888")
        next_free_row = note_row + 1
    else:
        next_free_row = row_idx
    # ==========================================
    # MODULE 2 (Enterprise only): Standardization summary
    # Quick visibility on the Dashboard tab so uncategorized items aren't only
    # discoverable by scrolling into Tab 2's detail section.
    # ==========================================
    if standardization and standardization.get("uncategorized_items"):
        std_header_row = next_free_row + 2
        ws_dash.cell(row=std_header_row, column=1, value="UNCATEGORIZED ITEMS DETECTED").font = Font(bold=True, size=12, color="FFB8860B")

        count = len(standardization["uncategorized_items"])
        total = standardization["uncategorized_total"]
        summary_row = std_header_row + 1
        ws_dash.cell(
            row=summary_row, column=1,
            value=f"{count} line item(s) totaling {currency_symbol}{total:,.2f} could not be confidently "
                  f"categorized. See 'Cleaned T12' tab for details — manual review recommended."
        ).font = Font(italic=True, color="FF888888")

        # FIX: previously this section never advanced next_free_row after
        # writing itself — anything added below it (e.g. Module 7 below)
        # would have silently overlapped these rows.
        next_free_row = summary_row + 1

    # ==========================================
    # MODULE 7 (Enterprise/Trial only): Forensic Variance Analysis
    # Per spec: a dedicated, visible 3-column grid (Month/Period | Calculated
    # Variance % | Forensic System Flag), with a light-red highlight on
    # ACTIVE flags only — "Insufficient Historical Data" rows are an explicit
    # non-finding, not something requiring the executive's urgent attention,
    # so they're shown but not highlighted. Full explanation for each flagged
    # row is available as a hoverable cell comment, same pattern as Module 1.
    #
    # Placed as a section on the Dashboard Summary tab (matching how Modules
    # 1 and 2 are laid out above), rather than a separate tab — keeps every
    # advanced-tier finding in one place for a fast top-to-bottom review.
    # ==========================================
    variance_analysis = metrics.get("variance_analysis")
    if variance_analysis is not None:
        va_header_row = next_free_row + 2
        ws_dash.cell(row=va_header_row, column=1, value="FORENSIC VARIANCE ANALYSIS").font = Font(bold=True, size=12, color="FFD4AF37")

        if variance_analysis.get("skipped"):
            skip_row = va_header_row + 1
            ws_dash.cell(row=skip_row, column=1, value=variance_analysis["reason"]).font = Font(italic=True, color="FF888888")
            next_free_row = skip_row + 1
        else:
            dips = variance_analysis.get("dips", [])
            if not dips:
                no_dips_row = va_header_row + 1
                ws_dash.cell(
                    row=no_dips_row, column=1,
                    value="No month-over-month revenue drops of 10% or more detected."
                ).font = Font(italic=True, color="FF888888")
                next_free_row = no_dips_row + 1
            else:
                table_header_row = va_header_row + 1
                ws_dash.cell(row=table_header_row, column=1, value="Month / Period").font = Font(bold=True)
                ws_dash.cell(row=table_header_row, column=2, value="Calculated Variance %").font = Font(bold=True)
                ws_dash.cell(row=table_header_row, column=3, value="Forensic System Flag").font = Font(bold=True)

                active_flag_fill = PatternFill(start_color="FFF8D7DA", end_color="FFF8D7DA", fill_type="solid")

                dip_row = table_header_row + 1
                for dip in dips:
                    month_cell = ws_dash.cell(row=dip_row, column=1, value=dip["month"])
                    pct_cell = ws_dash.cell(row=dip_row, column=2, value=dip["pct_change"] / 100)
                    pct_cell.number_format = '0.0%'
                    flag_cell = ws_dash.cell(row=dip_row, column=3, value=dip["flag"])

                    if dip.get("detail"):
                        flag_cell.comment = Comment(dip["detail"], "Variance Analysis Engine")

                    # Only highlight actionable flags — explicitly NOT the
                    # "Insufficient Historical Data" case, which is a
                    # deliberate non-finding rather than something urgent.
                    if "Insufficient Historical Data" not in dip["flag"]:
                        month_cell.fill = active_flag_fill
                        pct_cell.fill = active_flag_fill
                        flag_cell.fill = active_flag_fill

                    dip_row += 1

                next_free_row = dip_row

    # ==========================================
    # TAB: INDUSTRIAL ANALYSIS (Industrial & Logistics assets only)
    # Purely additive: only created when the AI extraction actually returned
    # industrial-specific data (rentable_square_footage / annual_base_rent /
    # annual_nnn_reimbursements not all null). A standard multifamily export
    # gets no new tab at all and is otherwise byte-for-byte unaffected by
    # this block's existence.
    #
    # Per-SF figures are LIVE FORMULAS referencing the raw extracted totals
    # written just above them on this same tab — never AI-computed or
    # pre-divided values — same "live formula" principle already used for
    # DSCR / Cash-on-Cash / Sized Loan Amount on the Cleaned T12 tab, so a
    # lender opening this file can trace exactly how each per-SF number was
    # derived instead of trusting a black box.
    # ==========================================
    industrial_metrics = metrics.get("industrial_metrics")
    if industrial_metrics and any(
        industrial_metrics.get(k) is not None
        for k in ("rentable_square_footage", "annual_base_rent", "annual_nnn_reimbursements")
    ):
        ws_ind = wb.create_sheet("Industrial Analysis")
        ws_ind.cell(row=1, column=1, value="INDUSTRIAL & LOGISTICS METRIC").font = Font(bold=True, size=12, color="FFD4AF37")
        ws_ind.cell(row=1, column=2, value="VALUE").font = Font(bold=True, size=12, color="FFD4AF37")

        rsf = safe_get(industrial_metrics, "rentable_square_footage")
        base_rent = safe_get(industrial_metrics, "annual_base_rent")
        nnn_reimb = safe_get(industrial_metrics, "annual_nnn_reimbursements")

        rsf_row = 2
        ws_ind.cell(row=rsf_row, column=1, value="Rentable Square Footage (RSF)").font = Font(bold=True)
        ws_ind.cell(row=rsf_row, column=2, value=rsf).number_format = '#,##0'

        base_rent_row = 3
        ws_ind.cell(row=base_rent_row, column=1, value="Annual Base Rent").font = Font(bold=True)
        ws_ind.cell(row=base_rent_row, column=2, value=base_rent).number_format = currency_format

        nnn_row = 4
        ws_ind.cell(row=nnn_row, column=1, value="Annual NNN Reimbursements").font = Font(bold=True)
        ws_ind.cell(row=nnn_row, column=2, value=nnn_reimb).number_format = currency_format

        total_rent_row = 5
        total_rent_cell = ws_ind.cell(row=total_rent_row, column=1, value="Total Annual Rent (Base + NNN)")
        total_rent_cell.font = Font(bold=True)
        ws_ind.cell(row=total_rent_row, column=2, value=f"=B{base_rent_row}+B{nnn_row}").number_format = currency_format

        # Per-SF formulas only make sense with a nonzero RSF — if RSF is
        # missing/zero, write an explanatory note rather than a formula that
        # would surface as #DIV/0! in the opened workbook.
        na_font = Font(italic=True, color="FF888888")

        base_rent_psf_row = 7
        ws_ind.cell(row=base_rent_psf_row, column=1, value="Base Rent per SF").font = Font(bold=True)
        if rsf:
            ws_ind.cell(row=base_rent_psf_row, column=2, value=f"=B{base_rent_row}/B{rsf_row}").number_format = currency_format
        else:
            ws_ind.cell(row=base_rent_psf_row, column=2, value="N/A (RSF not available)").font = na_font

        nnn_psf_row = 8
        ws_ind.cell(row=nnn_psf_row, column=1, value="NNN Reimbursement per SF").font = Font(bold=True)
        if rsf:
            ws_ind.cell(row=nnn_psf_row, column=2, value=f"=B{nnn_row}/B{rsf_row}").number_format = currency_format
        else:
            ws_ind.cell(row=nnn_psf_row, column=2, value="N/A (RSF not available)").font = na_font

        total_psf_row = 9
        ws_ind.cell(row=total_psf_row, column=1, value="Total Rent per SF (Base + NNN)").font = Font(bold=True)
        if rsf:
            ws_ind.cell(row=total_psf_row, column=2, value=f"=B{total_rent_row}/B{rsf_row}").number_format = currency_format
        else:
            ws_ind.cell(row=total_psf_row, column=2, value="N/A (RSF not available)").font = na_font

        ws_ind.column_dimensions['A'].width = 34
        ws_ind.column_dimensions['B'].width = 22

    # ==========================================
    # TAB: OFFICE ANALYSIS (Office assets only)
    # Same additive/gated pattern as Industrial Analysis above: only created
    # when the AI extraction returned office-specific data, otherwise the
    # export is completely unaffected.
    #
    # Pro-Rata Share and Net Effective Rent/SF are LIVE FORMULAS referencing
    # the raw cells above them, both wrapped in IFERROR(...,"N/A") per spec
    # so a missing/zero denominator degrades to a clean text label instead of
    # a #DIV/0!/#VALUE! error in the opened workbook.
    #
    # NOTE on Net Effective Rent/SF: the originally-specified formula
    # subtracted "TI Allowance Total" (a lump-sum $ figure) directly from
    # "Base Rent/SF x Lease Term" (a $/SF figure) -- a unit mismatch that
    # would silently produce a meaningless number. Fixed here by amortizing
    # TI Allowance Total over Tenant RSF first (ti_allowance_total /
    # tenant_rsf) so every term in the formula is in $/SF before combining.
    # Lease Term is stored in MONTHS (see views.py's office_metrics_guidance)
    # and converted to years (/12) since Base Rent/SF is an annual rate.
    #
    # "Expense Stop Value" is written as a raw reference figure only -- no
    # Expense Overage formula is built from it, since that would need the
    # building's actual operating expense per SF, which isn't part of this
    # schema (deliberately deferred, see office_metrics_guidance).
    # ==========================================
    office_metrics = metrics.get("office_metrics")
    if office_metrics and any(
        office_metrics.get(k) is not None
        for k in ("tenant_rsf", "building_total_rsf", "annual_base_rent_per_sf",
                   "expense_stop_value", "ti_allowance_total", "lease_term_months")
    ):
        ws_off = wb.create_sheet("Office Analysis")
        ws_off.cell(row=1, column=1, value="OFFICE METRIC").font = Font(bold=True, size=12, color="FFD4AF37")
        ws_off.cell(row=1, column=2, value="VALUE").font = Font(bold=True, size=12, color="FFD4AF37")

        tenant_rsf = safe_get(office_metrics, "tenant_rsf")
        building_rsf = safe_get(office_metrics, "building_total_rsf")
        base_rent_psf = safe_get(office_metrics, "annual_base_rent_per_sf")
        expense_stop = safe_get(office_metrics, "expense_stop_value")
        ti_allowance = safe_get(office_metrics, "ti_allowance_total")
        lease_term_months = safe_get(office_metrics, "lease_term_months")

        tenant_rsf_row = 2
        ws_off.cell(row=tenant_rsf_row, column=1, value="Tenant RSF").font = Font(bold=True)
        ws_off.cell(row=tenant_rsf_row, column=2, value=tenant_rsf).number_format = '#,##0'

        building_rsf_row = 3
        ws_off.cell(row=building_rsf_row, column=1, value="Building Total RSF").font = Font(bold=True)
        ws_off.cell(row=building_rsf_row, column=2, value=building_rsf).number_format = '#,##0'

        base_rent_psf_row = 4
        ws_off.cell(row=base_rent_psf_row, column=1, value="Annual Base Rent per SF").font = Font(bold=True)
        ws_off.cell(row=base_rent_psf_row, column=2, value=base_rent_psf).number_format = currency_format

        expense_stop_row = 5
        expense_stop_label_cell = ws_off.cell(row=expense_stop_row, column=1, value="Expense Stop Value")
        expense_stop_label_cell.font = Font(bold=True)
        expense_stop_label_cell.comment = Comment(
            "Reference figure only. No Expense Overage formula is calculated "
            "against this yet -- that requires the building's actual operating "
            "expense per SF, which this extraction does not currently capture.",
            "Underwriting AI"
        )
        ws_off.cell(row=expense_stop_row, column=2, value=expense_stop).number_format = currency_format

        ti_allowance_row = 6
        ws_off.cell(row=ti_allowance_row, column=1, value="TI Allowance Total").font = Font(bold=True)
        ws_off.cell(row=ti_allowance_row, column=2, value=ti_allowance).number_format = currency_format

        lease_term_row = 7
        ws_off.cell(row=lease_term_row, column=1, value="Lease Term (Months)").font = Font(bold=True)
        ws_off.cell(row=lease_term_row, column=2, value=lease_term_months).number_format = '#,##0'

        pro_rata_row = 9
        ws_off.cell(row=pro_rata_row, column=1, value="Pro-Rata Share").font = Font(bold=True)
        pro_rata_cell = ws_off.cell(
            row=pro_rata_row, column=2,
            value=f'=IFERROR(B{tenant_rsf_row}/B{building_rsf_row},"N/A")'
        )
        pro_rata_cell.number_format = '0.00%'

        ner_row = 10
        ws_off.cell(row=ner_row, column=1, value="Net Effective Rent per SF").font = Font(bold=True)
        ner_cell = ws_off.cell(
            row=ner_row, column=2,
            value=(
                f'=IFERROR((B{base_rent_psf_row}*(B{lease_term_row}/12)'
                f'-B{ti_allowance_row}/B{tenant_rsf_row})/(B{lease_term_row}/12),"N/A")'
            )
        )
        ner_cell.number_format = currency_format

        ws_off.column_dimensions['A'].width = 34
        ws_off.column_dimensions['B'].width = 22

    # ==========================================
    # TAB: RETAIL ANALYSIS (Retail assets only)
    # Same additive/gated pattern as Industrial/Office Analysis above.
    #
    # All 4 derived rows (Pro-Rata CAM Share, Sales Overage, Percentage Rent
    # Owed, Total Gross Rent Revenue) are LIVE FORMULAS chained off the raw
    # cells above them and off each other, each independently wrapped in
    # IFERROR(...,"N/A") per spec -- including the ones that reference an
    # upstream formula cell, so if an upstream cell degrades to the text
    # "N/A" (missing/zero denominator), the downstream formula's own IFERROR
    # catches the resulting #VALUE! and degrades to "N/A" too rather than
    # propagating a broken error chain.
    #
    # Unlike Office's Expense Stop Value, every retail field here IS used by
    # a formula -- this schema has no informational-only leftover field.
    # ==========================================
    retail_metrics = metrics.get("retail_metrics")
    if retail_metrics and any(
        retail_metrics.get(k) is not None
        for k in ("tenant_gla", "center_total_gla", "annual_base_rent_per_sf",
                   "percentage_rent_rate", "tenant_breakpoint_threshold", "tenant_gross_annual_sales")
    ):
        ws_ret = wb.create_sheet("Retail Analysis")
        ws_ret.cell(row=1, column=1, value="RETAIL METRIC").font = Font(bold=True, size=12, color="FFD4AF37")
        ws_ret.cell(row=1, column=2, value="VALUE").font = Font(bold=True, size=12, color="FFD4AF37")

        tenant_gla = safe_get(retail_metrics, "tenant_gla")
        center_gla = safe_get(retail_metrics, "center_total_gla")
        base_rent_psf = safe_get(retail_metrics, "annual_base_rent_per_sf")
        pct_rent_rate = safe_get(retail_metrics, "percentage_rent_rate")
        breakpoint_threshold = safe_get(retail_metrics, "tenant_breakpoint_threshold")
        gross_sales = safe_get(retail_metrics, "tenant_gross_annual_sales")

        tenant_gla_row = 2
        ws_ret.cell(row=tenant_gla_row, column=1, value="Tenant GLA").font = Font(bold=True)
        ws_ret.cell(row=tenant_gla_row, column=2, value=tenant_gla).number_format = '#,##0'

        center_gla_row = 3
        ws_ret.cell(row=center_gla_row, column=1, value="Center Total GLA").font = Font(bold=True)
        ws_ret.cell(row=center_gla_row, column=2, value=center_gla).number_format = '#,##0'

        base_rent_psf_row = 4
        ws_ret.cell(row=base_rent_psf_row, column=1, value="Annual Base Rent per SF").font = Font(bold=True)
        ws_ret.cell(row=base_rent_psf_row, column=2, value=base_rent_psf).number_format = currency_format

        pct_rent_rate_row = 5
        ws_ret.cell(row=pct_rent_rate_row, column=1, value="Percentage Rent Rate").font = Font(bold=True)
        ws_ret.cell(row=pct_rent_rate_row, column=2, value=pct_rent_rate).number_format = '0.00%'

        breakpoint_row = 6
        ws_ret.cell(row=breakpoint_row, column=1, value="Tenant Breakpoint Threshold").font = Font(bold=True)
        ws_ret.cell(row=breakpoint_row, column=2, value=breakpoint_threshold).number_format = currency_format

        gross_sales_row = 7
        ws_ret.cell(row=gross_sales_row, column=1, value="Tenant Gross Annual Sales").font = Font(bold=True)
        ws_ret.cell(row=gross_sales_row, column=2, value=gross_sales).number_format = currency_format

        pro_rata_row = 9
        ws_ret.cell(row=pro_rata_row, column=1, value="Pro-Rata CAM Share").font = Font(bold=True)
        pro_rata_cell = ws_ret.cell(
            row=pro_rata_row, column=2,
            value=f'=IFERROR(B{tenant_gla_row}/B{center_gla_row},"N/A")'
        )
        pro_rata_cell.number_format = '0.00%'

        overage_row = 10
        ws_ret.cell(row=overage_row, column=1, value="Retail Sales Overage").font = Font(bold=True)
        overage_cell = ws_ret.cell(
            row=overage_row, column=2,
            value=(
                f'=IFERROR(IF(B{gross_sales_row}>B{breakpoint_row},'
                f'B{gross_sales_row}-B{breakpoint_row},0),"N/A")'
            )
        )
        overage_cell.number_format = currency_format

        pct_rent_owed_row = 11
        ws_ret.cell(row=pct_rent_owed_row, column=1, value="Total Percentage Rent Owed").font = Font(bold=True)
        pct_rent_owed_cell = ws_ret.cell(
            row=pct_rent_owed_row, column=2,
            value=f'=IFERROR(B{overage_row}*B{pct_rent_rate_row},"N/A")'
        )
        pct_rent_owed_cell.number_format = currency_format

        total_revenue_row = 12
        total_revenue_cell = ws_ret.cell(row=total_revenue_row, column=1, value="Total Gross Rent Revenue (Base + Percentage)")
        total_revenue_cell.font = Font(bold=True)
        total_gross_revenue_cell = ws_ret.cell(
            row=total_revenue_row, column=2,
            value=(
                f'=IFERROR((B{tenant_gla_row}*B{base_rent_psf_row})+B{pct_rent_owed_row},"N/A")'
            )
        )
        total_gross_revenue_cell.font = Font(bold=True)
        total_gross_revenue_cell.number_format = currency_format

        ws_ret.column_dimensions['A'].width = 34
        ws_ret.column_dimensions['B'].width = 22

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()