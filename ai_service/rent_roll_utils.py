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


def detect_rent_roll_columns(rent_roll_df: pd.DataFrame):
    """
    Returns (rent_col, status_col) — the actual column names/labels in the
    DataFrame. Raises RentRollColumnError if either can't be confidently found.
    Column names are matched by content ('rent' / 'status' substrings), not
    hardcoded letters — a rent roll's column order varies file to file.
    """
    rent_col = None
    status_col = None

    for col in rent_roll_df.columns:
        col_str = str(col).lower()
        if 'rent' in col_str and 'market' not in col_str:
            rent_col = col
        if 'status' in col_str:
            status_col = col

    missing = []
    if rent_col is None:
        missing.append("a rent amount column (expected a column name containing 'rent', excluding 'market')")
    if status_col is None:
        missing.append("a status column (expected a column name containing 'status')")

    if missing:
        raise RentRollColumnError(
            "Could not confidently map the rent roll: " + "; ".join(missing) +
            f". Found columns: {list(rent_roll_df.columns)}"
        )

    return rent_col, status_col


def compute_ground_truth(rent_roll_df: pd.DataFrame, rent_col, status_col) -> dict:
    """
    Computes real occupancy and revenue figures directly from the rent roll
    DataFrame — the same "ground truth" your Excel COUNTIF/AVERAGE formulas
    compute, just done in Python so it can also feed the reconciliation engine
    (Module 1) before the Excel file even exists.
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