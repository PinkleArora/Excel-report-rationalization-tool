"""Tests for src.rationalization.rationalizer."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from src.ingestion.loader import WorkbookBundle
from src.rationalization.rationalizer import (
    ReportRecommendation,
    ReportStatus,
    assess_redundancy,
    build_inventory,
)
from src.schema_matching.matcher import ColumnMatch


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def bundle_a(tmp_path: Path) -> WorkbookBundle:
    path = tmp_path / "report_a.xlsx"
    df = pd.DataFrame({"revenue": [1], "cost": [2], "headcount": [3]})
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="S1", index=False)
    from src.ingestion.loader import load_workbook
    return load_workbook(path)


@pytest.fixture
def bundle_b(tmp_path: Path) -> WorkbookBundle:
    path = tmp_path / "report_b.xlsx"
    df = pd.DataFrame({"revenue": [10], "cost": [20], "headcount": [30]})
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="S1", index=False)
    from src.ingestion.loader import load_workbook
    return load_workbook(path)


@pytest.fixture
def bundle_unique(tmp_path: Path) -> WorkbookBundle:
    path = tmp_path / "unique.xlsx"
    df = pd.DataFrame({"fx_rate": [1.2], "basket_code": ["A"], "ccy": ["GBP"]})
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Rates", index=False)
    from src.ingestion.loader import load_workbook
    return load_workbook(path)


def _make_full_matches(bundle_a, bundle_b) -> list[ColumnMatch]:
    """Create exact matches for all 3 columns between a and b."""
    cols = ["revenue", "cost", "headcount"]
    return [
        ColumnMatch(
            source_workbook="report_a.xlsx", source_sheet="S1", source_col=c,
            target_workbook="report_b.xlsx", target_sheet="S1", target_col=c,
            score=100.0, match_type="exact",
        )
        for c in cols
    ]


# ---------------------------------------------------------------------------
# assess_redundancy
# ---------------------------------------------------------------------------

class TestAssessRedundancy:
    def test_returns_list(self, bundle_a, bundle_b):
        recs = assess_redundancy([bundle_a, bundle_b], [])
        assert isinstance(recs, list)

    def test_one_rec_per_bundle(self, bundle_a, bundle_b):
        recs = assess_redundancy([bundle_a, bundle_b], [])
        assert len(recs) == 2

    def test_single_bundle_no_matches_status_keep(self, bundle_a):
        recs = assess_redundancy([bundle_a], [])
        assert recs[0].status == ReportStatus.KEEP

    def test_full_overlap_recommends_retire(self, bundle_a, bundle_b):
        matches = _make_full_matches(bundle_a, bundle_b)
        recs = assess_redundancy([bundle_a, bundle_b], matches)
        statuses = {r.report_name: r.status for r in recs}
        # Both have 100% overlap — both should be retire or merge
        assert all(s in (ReportStatus.RETIRE, ReportStatus.MERGE) for s in statuses.values())

    def test_no_overlap_recommends_keep(self, bundle_a, bundle_unique):
        recs = assess_redundancy([bundle_a, bundle_unique], [])
        statuses = {r.report_name: r.status for r in recs}
        assert all(s == ReportStatus.KEEP for s in statuses.values())

    def test_overlap_pct_range(self, bundle_a, bundle_b):
        matches = _make_full_matches(bundle_a, bundle_b)
        recs = assess_redundancy([bundle_a, bundle_b], matches)
        for r in recs:
            assert 0 <= r.overlap_pct <= 100

    def test_rec_has_rationale(self, bundle_a, bundle_b):
        recs = assess_redundancy([bundle_a, bundle_b], [])
        assert all(len(r.rationale) > 0 for r in recs)

    def test_total_columns_correct(self, bundle_a):
        recs = assess_redundancy([bundle_a], [])
        assert recs[0].total_columns == 3  # revenue, cost, headcount

    def test_sorted_by_overlap_descending(self, bundle_a, bundle_b, bundle_unique):
        matches = _make_full_matches(bundle_a, bundle_b)
        recs = assess_redundancy([bundle_a, bundle_b, bundle_unique], matches)
        overlaps = [r.overlap_pct for r in recs]
        assert overlaps == sorted(overlaps, reverse=True)

    def test_merge_target_set_when_overlap(self, bundle_a, bundle_b):
        matches = _make_full_matches(bundle_a, bundle_b)
        recs = assess_redundancy([bundle_a, bundle_b], matches)
        rec_a = next(r for r in recs if r.report_name == "report_a.xlsx")
        assert rec_a.merge_target is not None

    def test_empty_bundles_returns_empty(self):
        recs = assess_redundancy([], [])
        assert recs == []


# ---------------------------------------------------------------------------
# build_inventory
# ---------------------------------------------------------------------------

class TestBuildInventory:
    def test_returns_dataframe(self, bundle_a, bundle_b):
        recs = assess_redundancy([bundle_a, bundle_b], [])
        df = build_inventory(recs)
        assert isinstance(df, pd.DataFrame)

    def test_one_row_per_rec(self, bundle_a, bundle_b):
        recs = assess_redundancy([bundle_a, bundle_b], [])
        df = build_inventory(recs)
        assert len(df) == 2

    def test_required_columns(self, bundle_a):
        recs = assess_redundancy([bundle_a], [])
        df = build_inventory(recs)
        required = {"report_name", "recommendation", "overlap_pct", "rationale"}
        assert required.issubset(set(df.columns))

    def test_recommendation_is_uppercase(self, bundle_a):
        recs = assess_redundancy([bundle_a], [])
        df = build_inventory(recs)
        assert df["recommendation"].iloc[0] == df["recommendation"].iloc[0].upper()

    def test_empty_input_returns_empty_df(self):
        df = build_inventory([])
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 0
