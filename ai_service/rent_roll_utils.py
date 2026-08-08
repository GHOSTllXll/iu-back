# backend/ai_service/rent_roll_utils.py
"""
Shared rent roll column detection and ground-truth computation.

Extracted into its own module so excel_generator.py and reconciliation.py
(Module 1) both use the EXACT same column-detection logic and the EXACT same
ground-truth math. If these ever drifted apart, the Excel file and the
reconciliation flags could disagree with each other about what the rent roll
actually says — which would be a real problem for a feature whose entire
purpose is catching inconsistencies.
"""
import pandas as pd


class RentRollColumnError(Exception):
    """Raised when the rent roll doesn't have columns we can confidently map."""
    pass


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
    real unit — blank tenant/unit-type/square-footage, but a rent value that's
    actually the SUM across every unit, not one unit's rent. Left in, this
    silently corrupts AVERAGE()-based metrics like Avg In-Place Rent (one
    absurdly large "rent" value drags the whole average way up).

    Heuristic: flag any row whose rent value exceeds `outlier_multiplier`
    times the MEDIAN of all nonzero rent values in the column, and exclude
    it. Deliberately simple and format-agnostic — doesn't need to guess
    column names for unit type / square footage, which vary a lot between
    rent rolls. This is a heuristic, not a certainty: a real property with
    one dramatically higher-rent unit (e.g. a penthouse) could in theory be
    caught by this too, though 5x the median is a generous threshold for
    that to be likely in practice.

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
    before using it for ground truth / Excel formulas — that function
    guarantees a real Occupied/Vacant column exists one way or another.

    Raises RentRollColumnError if:
    - no rent amount column can be found, or
    - NEITHER a status column NOR a tenant name column can be found (no
      reasonable way to determine occupancy by any method).
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
    'Vacant' text — the form every downstream consumer (Excel COUNTIF formulas,
    ground-truth math, reconciliation) actually needs. Two cases:

    1. An explicit status column already exists — normalize its values to
       exact 'Occupied'/'Vacant' casing (source data is often inconsistently
       cased) and use it as-is.
    2. No status column, but a Tenant column does — INFER occupancy: a blank/
       empty tenant name, or one literally containing "vacant", means vacant;
       anything else means occupied. This is a genuine inference, not a fact
       stated by the source document, so callers displaying this to the user
       (e.g. Module 1 reconciliation messages) should note it was inferred
       rather than presenting it with the same certainty as an explicit
       status column.

    Either way, this MATERIALIZES real values onto the DataFrame (mutates it
    in place, matching the pattern already used elsewhere for rent/status
    cleanup) — not just an in-memory calculation — so the exact same values
    end up in both the Excel export and the reconciliation engine's ground
    truth, with no risk of the two disagreeing.

    Returns (status_col_name, was_inferred: bool).
    """
    if status_col is not None:
        normalized = rent_roll_df[status_col].astype(str).str.strip().str.lower()
        rent_roll_df[status_col] = normalized.map({
            'occupied': 'Occupied',
            'vacant': 'Vacant',
        }).fillna(rent_roll_df[status_col])  # leave unrecognized values as-is rather than blanking them
        return status_col, False

    # Infer from tenant column
    tenant_clean = rent_roll_df[tenant_col].astype(str).str.strip().str.lower()
    is_vacant = tenant_clean.isin(['', 'nan', 'none']) | tenant_clean.str.contains('vacant', na=False)
    rent_roll_df['Status'] = is_vacant.map({True: 'Vacant', False: 'Occupied'})
    return 'Status', True


def compute_ground_truth(rent_roll_df: pd.DataFrame, rent_col, status_col) -> dict:
    """
    Computes real occupancy and revenue figures directly from the rent roll
    DataFrame — the same "ground truth" your Excel COUNTIF/AVERAGE formulas
    compute, just done in Python so it can also feed the reconciliation engine
    (Module 1) before the Excel file even exists.

    Expects status_col to already contain literal 'Occupied'/'Vacant' text —
    i.e. this should be called AFTER ensure_status_column(), not on a raw
    tenant-name column.
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