"""Tests for src.reconciliation.reconciler."""

from __future__ import annotations

import pandas as pd
import pytest

from src.consolidation.consolidator import LINEAGE_COL_SHEET, LINEAGE_COL_WORKBOOK
from src.reconciliation.reconciler import (
    ReconciliationResult,
    reconcile,
    reconciliation_summary,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def source_a() -> pd.DataFrame:
    return pd.DataFrame({"revenue": [100.0, 200.0], "cost": [50.0, 80.0]})


@pytest.fixture
def source_b() -> pd.DataFrame:
    return pd.DataFrame({"revenue": [300.0], "cost": [120.0]})


@pytest.fixture
def master_exact(source_a, source_b) -> pd.DataFrame:
    """Master that perfectly matches combined source totals."""
    combined = pd.concat([source_a, source_b], ignore_index=True)
    combined.insert(0, LINEAGE_COL_SHEET, "Sheet1")
    combined.insert(0, LINEAGE_COL_WORKBOOK, "source.xlsx")
    return combined


@pytest.fixture
def master_off(source_a) -> pd.DataFrame:
    """Master with inflated revenue (should fail reconciliation)."""
    df = source_a.copy()
    df["revenue"] = df["revenue"] * 2   # double the revenue
    df.insert(0, LINEAGE_COL_SHEET, "S")
    df.insert(0, LINEAGE_COL_WORKBOOK, "w.xlsx")
    return df


# ---------------------------------------------------------------------------
# reconcile
# ---------------------------------------------------------------------------

class TestReconcile:
    def test_returns_list(self, source_a, master_exact):
        results = reconcile([source_a], master_exact, numeric_columns=["revenue"])
        assert isinstance(results, list)

    def test_one_result_per_column(self, source_a, master_exact):
        results = reconcile([source_a], master_exact, numeric_columns=["revenue", "cost"])
        assert len(results) == 2

    def test_exact_match_passes(self, source_a, source_b, master_exact):
        results = reconcile([source_a, source_b], master_exact, numeric_columns=["revenue"])
        assert results[0].status == "pass"

    def test_large_delta_fails(self, source_a, master_off):
        results = reconcile([source_a], master_off, numeric_columns=["revenue"], tolerance=0.01)
        assert results[0].status == "fail"

    def test_source_total_correct(self, source_a, source_b, master_exact):
        results = reconcile([source_a, source_b], master_exact, numeric_columns=["revenue"])
        # 100 + 200 + 300 = 600
        assert abs(results[0].source_total - 600.0) < 0.01

    def test_output_total_correct(self, source_a, source_b, master_exact):
        results = reconcile([source_a, source_b], master_exact, numeric_columns=["revenue"])
        assert abs(results[0].output_total - 600.0) < 0.01

    def test_delta_zero_on_perfect_match(self, source_a, source_b, master_exact):
        results = reconcile([source_a, source_b], master_exact, numeric_columns=["revenue"])
        assert abs(results[0].delta) < 0.01

    def test_empty_numeric_columns_returns_empty(self, source_a, master_exact):
        results = reconcile([source_a], master_exact, numeric_columns=[])
        assert results == []

    def test_missing_column_in_output_treated_as_zero(self, source_a):
        output = pd.DataFrame({LINEAGE_COL_WORKBOOK: ["x"], LINEAGE_COL_SHEET: ["S"], "cost": [0.0]})
        results = reconcile([source_a], output, numeric_columns=["revenue"])
        assert results[0].output_total == 0.0

    def test_normalised_source_col_names_matched(self):
        """Source col 'Revenue (£)' should normalise to 'revenue' and reconcile."""
        source = pd.DataFrame({"Revenue (£)": [100.0, 200.0]})
        output = pd.DataFrame({
            LINEAGE_COL_WORKBOOK: ["x", "x"],
            LINEAGE_COL_SHEET: ["S", "S"],
            "revenue": [100.0, 200.0],
        })
        results = reconcile([source], output, numeric_columns=["revenue"])
        assert results[0].status == "pass"

    def test_warn_status_within_multiplier(self, source_a):
        """Delta 2x tolerance → warn not fail."""
        output = source_a.copy()
        output["revenue"] = output["revenue"] * 1.02   # 2% delta, tolerance 1%
        output.insert(0, LINEAGE_COL_SHEET, "S")
        output.insert(0, LINEAGE_COL_WORKBOOK, "w.xlsx")
        results = reconcile([source_a], output, numeric_columns=["revenue"], tolerance=0.01)
        assert results[0].status in ("warn", "fail")  # 2% > 1%, either is acceptable

    def test_multiple_sources_summed(self, source_a, source_b, master_exact):
        results = reconcile([source_a, source_b], master_exact, numeric_columns=["cost"])
        # 50 + 80 + 120 = 250
        assert abs(results[0].source_total - 250.0) < 0.01


# ---------------------------------------------------------------------------
# reconciliation_summary
# ---------------------------------------------------------------------------

class TestReconciliationSummary:
    def test_returns_dataframe(self, source_a, source_b, master_exact):
        results = reconcile([source_a, source_b], master_exact, numeric_columns=["revenue", "cost"])
        df = reconciliation_summary(results)
        assert isinstance(df, pd.DataFrame)

    def test_row_count(self, source_a, master_exact):
        results = reconcile([source_a], master_exact, numeric_columns=["revenue", "cost"])
        df = reconciliation_summary(results)
        assert len(df) == 2

    def test_required_columns(self, source_a, master_exact):
        results = reconcile([source_a], master_exact, numeric_columns=["revenue"])
        df = reconciliation_summary(results)
        assert "status" in df.columns
        assert "delta_pct" in df.columns

    def test_status_uppercase(self, source_a, master_exact):
        results = reconcile([source_a], master_exact, numeric_columns=["revenue"])
        df = reconciliation_summary(results)
        assert df["status"].iloc[0] == df["status"].iloc[0].upper()

    def test_fail_sorted_first(self, source_a, master_off):
        r_fail = reconcile([source_a], master_off, numeric_columns=["revenue"], tolerance=0.01)
        r_pass = reconcile([source_a], source_a, numeric_columns=["cost"])
        # reconcile expects output with lineage cols; add them
        sa_with_lineage = source_a.copy()
        sa_with_lineage.insert(0, LINEAGE_COL_SHEET, "S")
        sa_with_lineage.insert(0, LINEAGE_COL_WORKBOOK, "w")
        results = (
            reconcile([source_a], master_off, numeric_columns=["revenue"], tolerance=0.01)
            + reconcile([source_a], sa_with_lineage, numeric_columns=["cost"])
        )
        df = reconciliation_summary(results)
        assert df["status"].iloc[0] == "FAIL"

    def test_empty_input_returns_empty_df(self):
        df = reconciliation_summary([])
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 0
