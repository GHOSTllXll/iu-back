# backend/ai_service/custom_template_generator.py
"""
The "parallel path" for an institutional client's own proprietary Excel
template (see ai_service.models.ClientExcelTemplate) — generates a filled-in
copy of a CLIENT-PROVIDED workbook instead of building one of our own from
scratch (that's generate_underwriting_excel() in excel_generator.py).

CORE DIFFERENCE FROM excel_generator.py: that module OWNS the workbook it
produces — every cell, every format, every formula is ours, so it's free to
decide how to represent "this field is missing" (blank + an explanatory
Comment) or "this field is a default" (DEFAULTED_ZERO_COMMENT and friends).
This module is a GUEST in someone else's workbook. Their formulas, number
formats, macros (if .xlsm), logos, and layout are the whole point of this
feature and must survive completely untouched — we only ever write raw
VALUES into the specific cells the client themselves named, and never touch
formatting, never add new formulas, never write to any other cell.

MAPPING MECHANISM: named ranges (Excel's own Name Manager), not addresses or
column positions. The client opens their own template once, names each input
cell with one of the IU_* names in SCALAR_FIELD_CATALOG / ROW_FIELD_CATALOG
below, and uploads it. We resolve those names back to real (sheet, row, col)
coordinates via openpyxl — never by guessing a layout. A name this client's
template doesn't define is simply never written to; a field our own
extraction doesn't have a value for is simply never written either (see
"never fabricate" note on _write_scalars below). Nothing is ever placed
somewhere we weren't explicitly told to place it.

REPEATING ROW TABLE: a rent roll has a variable number of tenants, but a
named range is a single cell. The convention (documented for clients in
CUSTOM_TEMPLATE_NAMING_GUIDE.md): each IU_Row_* name marks the FIRST DATA
CELL of its column — e.g. if tenant names start at B10, name that cell
IU_Row_TenantName. We write downward from there, one row per tenant. Before
writing anything, we pre-flight-scan every cell we're about to touch; if any
of them already holds a formula, we refuse to write ANY of the row data and
raise TemplateCapacityError instead of silently overwriting part of the
client's own model. This is the same "never silently clobber/guess" standard
used throughout this codebase (see excel_generator.py's Argus tab), applied
to a workbook we don't own.
"""
import datetime
import io

import pandas as pd
from openpyxl import load_workbook
from openpyxl.comments import Comment

from .excel_generator import safe_get, clean_cell_value
from .rent_roll_utils import (
    detect_rent_roll_columns,
    ensure_status_column,
    detect_argus_fields,
    RentRollColumnError,
)


class TemplateValidationError(Exception):
    """
    Raised when a client's template can't be used at all — e.g. it defines
    none of our known IU_* names, or the IU_Row_* names it does define
    don't agree on a single anchor row. Used both at UPLOAD time (so the
    admin finds out immediately, not on the client's first real analysis)
    and defensively at generation time.
    """
    pass


class TemplateCapacityError(Exception):
    """
    Raised when writing this analysis's rent roll into the client's
    template would overwrite a cell that already holds a formula — i.e.
    the client has more tenants than their template's row table (as built)
    can safely hold without touching something of theirs. Deliberately
    stops the ENTIRE row-write rather than writing some rows and silently
    truncating the rest.
    """
    pass


# ==========================================
# SCALAR FIELD CATALOG — one value per analysis, each resolved from `metrics`
# (or `debt_assumptions`) exactly the way excel_generator.py already does
# for our own workbook. A field with no value for this particular deal is
# simply never written — see _write_scalars — not forced to a blank/zero,
# since we have no (and shouldn't invent) a way to annotate "why" inside
# someone else's model the way our own Comment()-based discipline does.
# ==========================================
def _total_building_area(metrics):
    return (
        safe_get(metrics, 'industrial_metrics', 'rentable_square_footage')
        or safe_get(metrics, 'office_metrics', 'building_total_rsf')
        or safe_get(metrics, 'retail_metrics', 'center_total_gla')
    )


SCALAR_FIELD_CATALOG = {
    'IU_PropertyName': lambda m, d: safe_get(m, 'property_metadata', 'property_name'),
    'IU_PropertyAddress': lambda m, d: safe_get(m, 'property_metadata', 'address'),
    'IU_AssetClass': lambda m, d: safe_get(m, 'property_metadata', 'asset_class'),
    'IU_TotalUnitCount': lambda m, d: safe_get(m, 'property_metadata', 'total_unit_count'),
    'IU_TotalBuildingAreaSF': lambda m, d: _total_building_area(m),
    'IU_PhysicalOccupancyPct': lambda m, d: safe_get(m, 'rent_roll_metrics', 'physical_occupancy_pct'),
    'IU_EconomicOccupancyPct': lambda m, d: safe_get(m, 'rent_roll_metrics', 'economic_occupancy_pct'),
    'IU_TotalVacantUnits': lambda m, d: safe_get(m, 'rent_roll_metrics', 'total_vacant_units'),
    'IU_AvgInPlaceRent': lambda m, d: safe_get(m, 'rent_roll_metrics', 'avg_in_place_monthly_rent'),
    'IU_AvgMarketRent': lambda m, d: safe_get(m, 'rent_roll_metrics', 'avg_market_monthly_rent'),
    'IU_AnnualizedGPR': lambda m, d: safe_get(m, 'rent_roll_metrics', 'annualized_gpr'),
    'IU_TotalAnnualConcessions': lambda m, d: safe_get(m, 'rent_roll_metrics', 'total_annual_concessions'),
    'IU_NetRentalIncome': lambda m, d: safe_get(m, 't12_revenue_expenses', 'net_rental_income'),
    # Buyer-entered financing assumptions (see _parse_debt_assumptions in
    # views.py) — not derived from source documents, same caveat that
    # applies to these three everywhere else they're used.
    'IU_DebtLTVPct': lambda m, d: (d or {}).get('ltv_pct'),
    'IU_DebtInterestRatePct': lambda m, d: (d or {}).get('interest_rate_pct'),
    'IU_DebtAmortizationYears': lambda m, d: (d or {}).get('amortization_years'),
}

# ==========================================
# ROW FIELD CATALOG — one per tenant, repeating down from each name's
# anchor cell. Sourced via the SAME column-detection helpers used for the
# Argus Rent Roll Export tab (detect_rent_roll_columns / detect_argus_fields)
# — not a parallel/duplicate detection path, so a fix to column detection
# benefits both tabs' worth of consumers at once.
#
# Each entry: (lookup_key, is_date). lookup_key names which column-detection
# result to read from the per-row context dict built in _build_row_context.
# ==========================================
ROW_FIELD_CATALOG = {
    'IU_Row_TenantName': ('tenant', False),
    'IU_Row_SuiteID': ('suite', False),
    'IU_Row_SpaceType': ('space_type', False),
    'IU_Row_SF': ('sf', False),
    'IU_Row_LeaseStart': ('lease_start', True),
    'IU_Row_LeaseEnd': ('lease_end', True),
    'IU_Row_BaseRent': ('base_rent', False),
    'IU_Row_AnnualRent': ('annual_rent', False),
    'IU_Row_Concessions': ('concessions', False),
    'IU_Row_Escalation': ('escalation', False),
    'IU_Row_RecoveryStructure': ('recovery', False),
    'IU_Row_BaseYear': ('base_year', False),
    'IU_Row_OccupancyStatus': ('occupancy_status', False),
    # Same placeholder-not-fabricated convention as the Argus tab's MLA
    # Profile column — see MLA_DEFAULT_COMMENT in excel_generator.py.
    'IU_Row_MLAProfile': ('mla_profile', False),
}

ALL_KNOWN_FIELD_NAMES = sorted(set(SCALAR_FIELD_CATALOG) | set(ROW_FIELD_CATALOG))

NOT_IN_SOURCE_COMMENT = (
    "Not present in the uploaded Rent Roll for this row -- left blank "
    "rather than estimated or guessed."
)


# ==========================================
# NAMED RANGE RESOLUTION
# ==========================================
def extract_defined_names(workbook):
    """
    Returns {name: (sheet_name, row, col)} for every defined name in the
    workbook that resolves to exactly one cell on exactly one sheet. A name
    that resolves to a multi-cell range, an external reference, or nothing
    at all is skipped (and reported separately by validate_template) --
    we only ever support a single-cell anchor per name, by design.
    """
    resolved = {}
    skipped = []
    for name, defined_name in workbook.defined_names.items():
        destinations = list(defined_name.destinations)
        if len(destinations) != 1:
            skipped.append(name)
            continue
        sheet_name, cell_ref = destinations[0]
        # cell_ref looks like "$B$10" for a single cell, "$B$10:$B$20" for a
        # range -- we only support the single-cell case.
        if ':' in cell_ref:
            skipped.append(name)
            continue
        if sheet_name not in workbook.sheetnames:
            skipped.append(name)
            continue
        sheet = workbook[sheet_name]
        try:
            cell = sheet[cell_ref.replace('$', '')]
        except (KeyError, ValueError):
            skipped.append(name)
            continue
        resolved[name] = (sheet_name, cell.row, cell.column)
    return resolved, skipped


def validate_template(workbook):
    """
    Checks a client's template against our known field catalog WITHOUT
    writing anything. Used at upload time (so a bad template is caught
    immediately) and defensively at generation time. Returns a report dict;
    raises TemplateValidationError only for a genuinely unusable template
    (zero known fields found, or IU_Row_* names that disagree on a row).
    """
    resolved, skipped_names = extract_defined_names(workbook)

    found_scalar = {n: resolved[n] for n in SCALAR_FIELD_CATALOG if n in resolved}
    found_row = {n: resolved[n] for n in ROW_FIELD_CATALOG if n in resolved}
    missing_scalar = sorted(set(SCALAR_FIELD_CATALOG) - set(found_scalar))
    missing_row = sorted(set(ROW_FIELD_CATALOG) - set(found_row))

    if not found_scalar and not found_row:
        raise TemplateValidationError(
            "This template defines none of our known field names. Nothing "
            f"would ever be written to it. Supported names: {', '.join(ALL_KNOWN_FIELD_NAMES)}"
        )

    anchor_row = None
    anchor_conflicts = []
    for name, (sheet_name, row, _col) in found_row.items():
        if anchor_row is None:
            anchor_row = (sheet_name, row)
        elif (sheet_name, row) != anchor_row:
            anchor_conflicts.append(
                f"{name} is on '{sheet_name}' row {row}, but other IU_Row_* "
                f"names are anchored on '{anchor_row[0]}' row {anchor_row[1]}"
            )

    if anchor_conflicts:
        raise TemplateValidationError(
            "The IU_Row_* names in this template don't all point to the "
            "same row. Every IU_Row_* name must be the FIRST DATA CELL of "
            "its own column, all on the same row (see "
            "CUSTOM_TEMPLATE_NAMING_GUIDE.md). Conflicts found: "
            + "; ".join(anchor_conflicts)
        )

    return {
        'found_fields': sorted(set(found_scalar) | set(found_row)),
        'missing_scalar_fields': missing_scalar,
        'missing_row_fields': missing_row,
        'row_anchor': {'sheet': anchor_row[0], 'row': anchor_row[1]} if anchor_row else None,
        'unresolvable_defined_names': skipped_names,
    }


# ==========================================
# ROW DATA SOURCING — mirrors the Argus tab's per-row logic in
# excel_generator.py exactly (same columns detected the same way), just
# collected into a plain dict per row instead of written to our own fixed
# column layout.
# ==========================================
def _build_row_contexts(rent_roll_df):
    try:
        rent_col, status_col, tenant_col = detect_rent_roll_columns(rent_roll_df)
        status_col, _ = ensure_status_column(rent_roll_df, status_col, tenant_col)
    except RentRollColumnError as e:
        raise TemplateValidationError(f"Could not read this rent roll: {e}")

    rent_roll_df = rent_roll_df.copy()
    rent_roll_df[rent_col] = pd.to_numeric(rent_roll_df[rent_col], errors='coerce').fillna(0)
    rent_roll_df[status_col] = rent_roll_df[status_col].astype(str).str.strip()

    fields = detect_argus_fields(rent_roll_df)
    suite_col = fields['suite_col']
    sf_col = fields['sf_col']
    lease_start_col = fields['lease_start_col']
    lease_end_col = fields['lease_end_col']
    concession_col = fields['concession_col']
    escalation_col = fields['escalation_col']
    space_type_col = fields['space_type_col']
    recovery_col = fields['recovery_col']
    base_year_col = fields['base_year_col']

    lease_start_dates = pd.to_datetime(rent_roll_df[lease_start_col], errors='coerce') if lease_start_col else None
    lease_end_dates = pd.to_datetime(rent_roll_df[lease_end_col], errors='coerce') if lease_end_col else None

    contexts = []
    for i in range(len(rent_roll_df)):
        src = rent_roll_df.iloc[i]
        status_val = clean_cell_value(src[status_col])
        is_vacant = isinstance(status_val, str) and status_val.strip().lower() == 'vacant'

        base_rent_raw = clean_cell_value(src[rent_col])
        base_rent_val = base_rent_raw if isinstance(base_rent_raw, (int, float)) else 0

        if is_vacant:
            lease_start_val = None
            lease_end_val = None
        else:
            lease_start_val = lease_start_dates.iloc[i].to_pydatetime() if lease_start_dates is not None and pd.notna(lease_start_dates.iloc[i]) else None
            lease_end_val = lease_end_dates.iloc[i].to_pydatetime() if lease_end_dates is not None and pd.notna(lease_end_dates.iloc[i]) else None

        escalation_raw = clean_cell_value(src[escalation_col]) if escalation_col else ""
        escalation_val = escalation_raw if (isinstance(escalation_raw, str) and escalation_raw.strip()) else "0"

        contexts.append({
            'tenant': clean_cell_value(src[tenant_col]) if tenant_col else "",
            'suite': clean_cell_value(src[suite_col]) if suite_col else "",
            'space_type': clean_cell_value(src[space_type_col]) if space_type_col else "",
            'sf': clean_cell_value(src[sf_col]) if sf_col else "",
            'lease_start': lease_start_val,
            'lease_end': lease_end_val,
            'base_rent': base_rent_val,
            'annual_rent': round(base_rent_val * 12, 2),
            'concessions': clean_cell_value(src[concession_col]) if concession_col else 0.0,
            'escalation': escalation_val,
            'recovery': clean_cell_value(src[recovery_col]) if recovery_col else "",
            'base_year': clean_cell_value(src[base_year_col]) if base_year_col else "",
            'occupancy_status': status_val,
            'mla_profile': "General_Market",
        })
    return contexts


# ==========================================
# WRITING
# ==========================================
def _write_scalars(workbook, resolved_names, metrics, debt_assumptions):
    written, skipped_empty = [], []
    for name, resolver in SCALAR_FIELD_CATALOG.items():
        if name not in resolved_names:
            continue
        value = resolver(metrics, debt_assumptions)
        if value is None or value == "":
            # Never force a blank/zero into the client's own cell when we
            # genuinely have nothing — we have no safe way to annotate
            # "why" inside their model, so leaving whatever was already
            # there (normally nothing) is the honest choice here.
            skipped_empty.append(name)
            continue
        sheet_name, row, col = resolved_names[name]
        workbook[sheet_name].cell(row=row, column=col, value=value)
        written.append(name)
    return written, skipped_empty


def _write_rows(workbook, resolved_names, row_contexts):
    row_names = {n: resolved_names[n] for n in ROW_FIELD_CATALOG if n in resolved_names}
    if not row_names or not row_contexts:
        return [], 0

    any_name = next(iter(row_names.values()))
    sheet_name, anchor_row, _ = any_name
    sheet = workbook[sheet_name]
    num_rows = len(row_contexts)

    # PRE-FLIGHT: refuse to write ANYTHING if any target cell already holds
    # a formula — the client's template has fewer pre-built rows than this
    # analysis has tenants, and row `anchor_row + num_rows - 1` has run into
    # their own formula (a summary row, usually). Better to fail loudly here
    # than silently overwrite part of their model or silently truncate the
    # rent roll with no indication anything was cut off.
    conflicts = []
    for name, (_sheet, _row, col) in row_names.items():
        for offset in range(num_rows):
            target_row = anchor_row + offset
            cell = sheet.cell(row=target_row, column=col)
            if isinstance(cell.value, str) and cell.value.startswith('='):
                conflicts.append(f"{name}: {sheet_name}!{cell.coordinate} already contains a formula ({cell.value!r})")
    if conflicts:
        raise TemplateCapacityError(
            f"This analysis has {num_rows} rent-roll rows, starting at "
            f"{sheet_name} row {anchor_row}, which runs into the client's "
            "own formula(s) before reaching the end. Nothing was written. "
            "Ask the client to extend their template's row table past row "
            f"{anchor_row + num_rows - 1}, or reduce to fewer tenants. "
            + "; ".join(conflicts)
        )

    for offset, ctx in enumerate(row_contexts):
        target_row = anchor_row + offset
        for name, (_sheet, _row, col) in row_names.items():
            lookup_key, is_date = ROW_FIELD_CATALOG[name]
            value = ctx[lookup_key]
            cell = sheet.cell(row=target_row, column=col)
            if value is None or value == "":
                cell.value = ""
                cell.comment = Comment(NOT_IN_SOURCE_COMMENT, "Underwriting AI")
            elif is_date:
                cell.value = value  # already a real datetime, see _build_row_contexts
            else:
                cell.value = value

    return sorted(row_names), num_rows


def generate_excel_from_template(template_path, metrics, rent_roll_df, debt_assumptions=None):
    """
    Loads a CLIENT-PROVIDED workbook (template_path: a filesystem path or
    file-like object) and writes this analysis's data into whichever IU_*
    named cells it defines. Returns (xlsx_bytes, report) where report is a
    dict describing exactly what was written/skipped — surfaced to the
    admin/analyst, never silently swallowed.

    keep_vba is driven by the file's own extension: an .xlsm's macros are
    preserved byte-for-byte on save (openpyxl never executes them, at load
    or at save), an .xlsx stays macro-free.
    """
    name = getattr(template_path, 'name', None) or str(template_path)
    is_macro_enabled = name.lower().endswith('.xlsm')

    workbook = load_workbook(template_path, keep_vba=is_macro_enabled)

    validation = validate_template(workbook)  # raises TemplateValidationError if unusable
    resolved_names, _ = extract_defined_names(workbook)

    written_scalars, skipped_scalars = _write_scalars(workbook, resolved_names, metrics, debt_assumptions)

    row_contexts = _build_row_contexts(rent_roll_df)
    written_row_fields, num_rows_written = _write_rows(workbook, resolved_names, row_contexts)

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    report = {
        'written_scalar_fields': written_scalars,
        'skipped_scalar_fields_no_value': skipped_scalars,
        'missing_scalar_fields_not_in_template': validation['missing_scalar_fields'],
        'written_row_fields': written_row_fields,
        'missing_row_fields_not_in_template': validation['missing_row_fields'],
        'rows_written': num_rows_written,
    }
    return buffer.getvalue(), report
