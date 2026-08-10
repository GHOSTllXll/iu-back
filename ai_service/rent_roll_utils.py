# backend/ai_service/rent_roll_utils.py
"""
Shared rent roll column detection, ground-truth computation, and format
normalization (including charge-ledger reshaping — see reshape_charge_ledger).

Extracted into its own module so excel_generator.py and reconciliation.py
(Module 1) both use the EXACT same column-detection logic and the EXACT same
ground-truth math.
"""
import pandas as pd


class RentRollColumnError(Exception):
    """Raised when the rent roll doesn't have columns we can confidently map."""
    pass


def detect_charge_ledger_columns(rent_roll_df: pd.DataFrame):
    """
    Detects the "charge ledger" rent roll format: multiple rows per unit, one
    per charge type (rent, utilities, deposits, etc.), rather than one row
    per unit with a single rent figure. Common in property-management-system
    exports (e.g. ResAnalytics-style exports).

    Returns a dict of column names if detected, else None. Detection requires
    finding columns for: charge type, charge amount, and unit identifier —
    without those three, there's no ledger structure to reshape.
    """
    charge_col = None
    amount_col = None
    unit_col = None
    resident_col = None
    name_col = None

    for col in rent_roll_df.columns:
        col_str = str(col).lower().strip()
        if 'charge' in col_str:
            charge_col = col
        if col_str == 'amount' or col_str.startswith('amount'):
            amount_col = col
        if col_str == 'unit':
            unit_col = col
        if col_str == 'resident':
            resident_col = col
        if col_str == 'name':
            name_col = col

    if charge_col is None or amount_col is None or unit_col is None:
        return None

    return {
        'charge_col': charge_col,
        'amount_col': amount_col,
        'unit_col': unit_col,
        'resident_col': resident_col,
        'name_col': name_col,
    }


def reshape_charge_ledger(rent_roll_df: pd.DataFrame, ledger_cols: dict) -> pd.DataFrame:
    """
    Collapses a charge-ledger rent roll (multiple rows per unit, one per
    charge type, in contiguous row blocks) into the standard one-row-per-unit
    shape — Unit / Tenant / Rent / Status — that everything else in this
    system (Excel formulas, reconciliation ground truth) already expects.

    Assumptions about the source format (verified against a real export, but
    a heuristic nonetheless — see conversation history):
    - Each unit occupies a contiguous block of rows. The first row of a
      block has the Unit/resident info filled in; continuation rows (more
      charges for the same unit) have those fields blank.
    - The base rent charge is identified by a Charge Code value of
      literally "rent" (case-insensitive) — confirmed against real data,
      but a property using a different literal code (e.g. "base rent",
      "rnt") would not be matched by this and would show $0 rent instead.
    - "DOWN" or a blank resident name means vacant — also confirmed against
      real data; other systems may use different vacancy markers.
    - "Total" rows and blank separator rows are not real charge lines and
      are dropped before grouping.
    """
    unit_col = ledger_cols['unit_col']
    resident_col = ledger_cols['resident_col']
    name_col = ledger_cols['name_col']
    charge_col = ledger_cols['charge_col']
    amount_col = ledger_cols['amount_col']

    display_name_col = name_col or resident_col

    df = rent_roll_df.reset_index(drop=True).copy()

    # Forward-fill unit-identifying columns down through each block's
    # continuation rows (blank Unit/resident on rows after the first).
    id_cols = [c for c in [unit_col, resident_col, name_col] if c is not None]
    df[id_cols] = df[id_cols].ffill()

    # Drop "Total" summary rows and fully-blank charge rows — neither is a
    # real charge line to inspect.
    charge_clean = df[charge_col].astype(str).str.strip().str.lower()
    df = df[~charge_clean.isin(['total', 'nan', ''])]
    df = df[df[unit_col].notna()]

    records = []
    for unit_id, group in df.groupby(unit_col, sort=False):
        charges = group[charge_col].astype(str).str.strip().str.lower()
        rent_rows = group[charges == 'rent']
        rent_amount = (
            pd.to_numeric(rent_rows[amount_col], errors='coerce').sum()
            if not rent_rows.empty else 0
        )

        resident_val = group[display_name_col].iloc[0] if display_name_col else None
        resident_str = str(resident_val).strip().lower() if pd.notna(resident_val) else ''
        is_vacant = resident_str in ('', 'nan', 'down') or 'vacant' in resident_str

        records.append({
            'Unit': unit_id,
            'Tenant': resident_val if pd.notna(resident_val) else '',
            'Rent': rent_amount,
            'Status': 'Vacant' if is_vacant else 'Occupied',
        })

    result = pd.DataFrame(records)

    if result.empty:
        return result

    # Filter out stray section-header rows that leak into the "unit" grouping
    # (e.g. a literal row of text like "Current/Notice/Vacant Residents"
    # sitting in the source sheet as a section divider, not a real unit) —
    # same statistical-outlier approach already used for rent totals-row
    # stripping: a real unit ID is short, a stray label is much longer.
    unit_id_lengths = result['Unit'].astype(str).str.len()
    median_len = unit_id_lengths.median()
    result = result[unit_id_lengths <= max(median_len * 4, 10)].reset_index(drop=True)

    return result


def detect_rent_column(rent_roll_df: pd.DataFrame):
    """
    Finds just the rent amount column, without requiring status/tenant info.
    Used early in the pipeline (for outlier-row cleanup) before we know
    whether the full occupancy detection will even be needed for this tier.
    """
    for col in rent_roll_df.columns:
        col_str = str(col).lower()
        if 'rent' in col_str and 'market' not in col_str:
            return col

    raise RentRollColumnError(
        "Could not confidently map the rent roll: a rent amount column "
        "(expected a column name containing 'rent', excluding 'market'). "
        f"Found columns: {list(rent_roll_df.columns)}"
    )


def strip_rent_outlier_rows(rent_roll_df: pd.DataFrame, rent_col, outlier_multiplier: float = 5.0):
    """
    Some rent rolls include a trailing "Totals" or summary row that isn't a
    real unit — blank tenant/unit-type/square-footage, but a rent value
    that's actually the SUM across every unit, not one unit's rent. Left in,
    this silently corrupts AVERAGE()-based metrics like Avg In-Place Rent.

    Heuristic: flag any row whose rent value exceeds `outlier_multiplier`
    times the MEDIAN of all nonzero rent values in the column, and exclude
    it.

    Returns (cleaned_df, excluded_rows: list of dicts) so callers can surface
    what was excluded rather than silently dropping data.
    """
    numeric_rent = pd.to_numeric(rent_roll_df[rent_col], errors='coerce')
    nonzero = numeric_rent[numeric_rent > 0]

    if len(nonzero) < 3:
        return rent_roll_df, []  # not enough data points to judge outliers meaningfully

    median_rent = nonzero.median()
    threshold = median_rent * outlier_multiplier

    is_outlier = numeric_rent > threshold
    excluded_rows = []

    if is_outlier.any():
        excluded_rows = rent_roll_df.loc[is_outlier].to_dict('records')
        rent_roll_df = rent_roll_df.loc[~is_outlier].reset_index(drop=True)

    return rent_roll_df, excluded_rows


def detect_rent_roll_columns(rent_roll_df: pd.DataFrame):
    """
    Returns (rent_col, status_col, tenant_col). status_col and tenant_col may
    individually be None (but not both), depending on what the rent roll
    actually has. Callers should pass the result to ensure_status_column()
    before using it for ground truth / Excel formulas.

    Raises RentRollColumnError if:
    - no rent amount column can be found, or
    - NEITHER a status column NOR a tenant name column can be found.
    """
    rent_col = detect_rent_column(rent_roll_df)
    status_col = None
    tenant_col = None

    for col in rent_roll_df.columns:
        col_str = str(col).lower()
        if 'status' in col_str:
            status_col = col
        if tenant_col is None and 'tenant' in col_str:
            tenant_col = col

    if status_col is None and tenant_col is None:
        raise RentRollColumnError(
            "Could not confidently map the rent roll: no status column (containing "
            "'status') and no tenant name column (containing 'tenant') found — "
            "unable to determine occupancy by any method. "
            f"Found columns: {list(rent_roll_df.columns)}"
        )

    return rent_col, status_col, tenant_col


def ensure_status_column(rent_roll_df: pd.DataFrame, status_col, tenant_col):
    """
    Guarantees the DataFrame has a real column containing literal 'Occupied' /
    'Vacant' text. Two cases:
    1. An explicit status column already exists — normalize its values.
    2. No status column, but a Tenant column does — infer occupancy from
       blank/vacant-containing tenant names.

    Returns (status_col_name, was_inferred: bool).
    """
    if status_col is not None:
        normalized = rent_roll_df[status_col].astype(str).str.strip().str.lower()
        rent_roll_df[status_col] = normalized.map({
            'occupied': 'Occupied',
            'vacant': 'Vacant',
        }).fillna(rent_roll_df[status_col])
        return status_col, False

    tenant_clean = rent_roll_df[tenant_col].astype(str).str.strip().str.lower()
    is_vacant = tenant_clean.isin(['', 'nan', 'none']) | tenant_clean.str.contains('vacant', na=False)
    rent_roll_df['Status'] = is_vacant.map({True: 'Vacant', False: 'Occupied'})
    return 'Status', True


def compute_ground_truth(rent_roll_df: pd.DataFrame, rent_col, status_col) -> dict:
    """
    Computes real occupancy and revenue figures directly from the rent roll
    DataFrame. Expects status_col to already contain literal 'Occupied'/
    'Vacant' text — i.e. this should be called AFTER ensure_status_column().
    """
    numeric_rent = pd.to_numeric(rent_roll_df[rent_col], errors='coerce').fillna(0)
    status_clean = rent_roll_df[status_col].astype(str).str.strip()

    total_units = len(rent_roll_df)
    occupied_units = int((status_clean.str.lower() == 'occupied').sum())
    vacant_units = int((status_clean.str.lower() == 'vacant').sum())

    physical_occupancy_pct = (occupied_units / total_units) if total_units > 0 else None
    annualized_rent_roll_income = float(numeric_rent.sum() * 12)

    return {
        "total_units": total_units,
        "occupied_units": occupied_units,
        "vacant_units": vacant_units,
        "physical_occupancy_pct": physical_occupancy_pct,
        "annualized_rent_roll_income": annualized_rent_roll_income,
    }