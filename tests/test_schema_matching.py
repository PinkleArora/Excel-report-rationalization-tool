"""Tests for src.schema_matching — normalizer, matcher, and utilities."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from src.ingestion.loader import WorkbookBundle
from src.schema_matching.matcher import (
    ColumnMatch,
    build_match_matrix,
    match_all_bundles,
    match_columns,
    match_workbooks,
)
from src.schema_matching.normalizer import canonical_for_group, normalize_column_name


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def bundle_a(tmp_path: Path) -> WorkbookBundle:
    path = tmp_path / "a.xlsx"
    df = pd.DataFrame({"Revenue": [1, 2], "Cost": [3, 4], "Headcount": [5, 6]})
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Sheet1", index=False)
    from src.ingestion.loader import load_workbook
    return load_workbook(path)


@pytest.fixture
def bundle_b(tmp_path: Path) -> WorkbookBundle:
    path = tmp_path / "b.xlsx"
    df = pd.DataFrame({"revenue": [10, 20], "cost": [30, 40], "headcount": [50, 60]})
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Sheet1", index=False)
    from src.ingestion.loader import load_workbook
    return load_workbook(path)


@pytest.fixture
def bundle_unique(tmp_path: Path) -> WorkbookBundle:
    path = tmp_path / "unique.xlsx"
    df = pd.DataFrame({"FX_Rate": [1.2, 1.3], "Basket_Code": ["A", "B"]})
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Rates", index=False)
    from src.ingestion.loader import load_workbook
    return load_workbook(path)


# ---------------------------------------------------------------------------
# normalize_column_name
# ---------------------------------------------------------------------------

class TestNormalizeColumnName:
    def test_lowercase(self):
        assert normalize_column_name("Revenue") == "revenue"

    def test_strips_whitespace(self):
        assert normalize_column_name("  cost  ") == "cost"

    def test_replaces_special_chars(self):
        assert normalize_column_name("Revenue (£)") == "revenue"

    def test_collapses_spaces(self):
        assert normalize_column_name("TOTAL  COST") == "total_cost"

    def test_numeric_preserved(self):
        assert normalize_column_name("Q1_2024") == "q1_2024"

    def test_empty_string(self):
        assert normalize_column_name("") == "unnamed"

    def test_only_special_chars(self):
        assert normalize_column_name("###") == "unnamed"

    def test_leading_trailing_underscores_stripped(self):
        result = normalize_column_name("__hello__")
        assert not result.startswith("_")
        assert not result.endswith("_")


# ---------------------------------------------------------------------------
# canonical_for_group
# ---------------------------------------------------------------------------

class TestCanonicalForGroup:
    def test_picks_shortest(self):
        assert canonical_for_group(["revenue_total", "rev", "revenue"]) == "rev"

    def test_ties_broken_alphabetically(self):
        result = canonical_for_group(["cost", "beta"])
        assert result == "beta"

    def test_single_element(self):
        assert canonical_for_group(["revenue"]) == "revenue"


# ---------------------------------------------------------------------------
# ColumnMatch dataclass
# ---------------------------------------------------------------------------

class TestColumnMatch:
    def test_fields(self):
        m = ColumnMatch(
            source_workbook="a.xlsx", source_sheet="S1", source_col="Revenue",
            target_workbook="b.xlsx", target_sheet="S1", target_col="revenue",
            score=100.0, match_type="exact",
        )
        assert m.score == 100.0
        assert m.match_type == "exact"


# ---------------------------------------------------------------------------
# match_columns
# ---------------------------------------------------------------------------

class TestMatchColumns:
    @pytest.fixture
    def src_sheets(self):
        return {"Sheet1": pd.DataFrame({"Revenue": [1], "Cost": [2]})}

    @pytest.fixture
    def tgt_sheets_exact(self):
        return {"Sheet1": pd.DataFrame({"revenue": [10], "cost": [20]})}

    @pytest.fixture
    def tgt_sheets_fuzzy(self):
        return {"Sheet1": pd.DataFrame({"Rev": [10], "Expenditure": [20]})}

    def test_exact_match_found(self, src_sheets, tgt_sheets_exact):
        matches = match_columns(src_sheets, tgt_sheets_exact, threshold=80.0)
        col_names = {m.source_col for m in matches}
        assert "Revenue" in col_names

    def test_exact_match_score_100(self, src_sheets, tgt_sheets_exact):
        matches = match_columns(src_sheets, tgt_sheets_exact, threshold=80.0)
        exact = [m for m in matches if m.match_type == "exact"]
        assert all(m.score == 100.0 for m in exact)

    def test_fuzzy_match_above_threshold(self, src_sheets, tgt_sheets_fuzzy):
        # "revenue" vs "Rev" — should score high
        matches = match_columns(src_sheets, tgt_sheets_fuzzy, threshold=70.0)
        assert len(matches) >= 1

    def test_below_threshold_excluded(self):
        src = {"S": pd.DataFrame({"alpha": [1]})}
        tgt = {"S": pd.DataFrame({"zzzzzzzz": [2]})}
        matches = match_columns(src, tgt, threshold=90.0)
        assert matches == []

    def test_empty_target_returns_empty(self, src_sheets):
        matches = match_columns(src_sheets, {}, threshold=80.0)
        assert matches == []

    def test_empty_source_returns_empty(self, tgt_sheets_exact):
        matches = match_columns({}, tgt_sheets_exact, threshold=80.0)
        assert matches == []

    def test_sorted_by_score_descending(self, src_sheets, tgt_sheets_exact):
        matches = match_columns(src_sheets, tgt_sheets_exact, threshold=50.0)
        scores = [m.score for m in matches]
        assert scores == sorted(scores, reverse=True)

    def test_workbook_labels_set(self, src_sheets, tgt_sheets_exact):
        matches = match_columns(
            src_sheets, tgt_sheets_exact,
            source_workbook="a.xlsx", target_workbook="b.xlsx",
        )
        assert all(m.source_workbook == "a.xlsx" for m in matches)
        assert all(m.target_workbook == "b.xlsx" for m in matches)

    def test_each_source_col_appears_at_most_once(self, src_sheets, tgt_sheets_exact):
        matches = match_columns(src_sheets, tgt_sheets_exact)
        source_keys = [(m.source_sheet, m.source_col) for m in matches]
        assert len(source_keys) == len(set(source_keys))


# ---------------------------------------------------------------------------
# match_workbooks
# ---------------------------------------------------------------------------

class TestMatchWorkbooks:
    def test_returns_list(self, bundle_a, bundle_b):
        matches = match_workbooks(bundle_a, bundle_b)
        assert isinstance(matches, list)

    def test_workbook_names_set(self, bundle_a, bundle_b):
        matches = match_workbooks(bundle_a, bundle_b)
        assert all(m.source_workbook == "a.xlsx" for m in matches)
        assert all(m.target_workbook == "b.xlsx" for m in matches)

    def test_exact_matches_for_same_normalised_cols(self, bundle_a, bundle_b):
        matches = match_workbooks(bundle_a, bundle_b)
        exact = [m for m in matches if m.match_type == "exact"]
        assert len(exact) >= 3  # Revenue/revenue, Cost/cost, Headcount/headcount


# ---------------------------------------------------------------------------
# match_all_bundles
# ---------------------------------------------------------------------------

class TestMatchAllBundles:
    def test_empty_returns_empty(self):
        assert match_all_bundles([]) == []

    def test_single_bundle_returns_empty(self, bundle_a):
        assert match_all_bundles([bundle_a]) == []

    def test_two_bundles_returns_matches(self, bundle_a, bundle_b):
        matches = match_all_bundles([bundle_a, bundle_b])
        assert len(matches) >= 1

    def test_no_overlap_with_unique_bundle(self, bundle_a, bundle_unique):
        matches = match_all_bundles([bundle_a, bundle_unique], threshold=90.0)
        # Revenue/Cost/Headcount vs FX_Rate/Basket_Code — no matches expected at 90
        assert all(m.score < 90 or m.match_type != "exact" for m in matches)


# ---------------------------------------------------------------------------
# build_match_matrix
# ---------------------------------------------------------------------------

class TestBuildMatchMatrix:
    def test_square_matrix(self, bundle_a, bundle_b):
        matrix = build_match_matrix([bundle_a, bundle_b])
        assert matrix.shape == (2, 2)

    def test_diagonal_is_zero(self, bundle_a, bundle_b):
        matrix = build_match_matrix([bundle_a, bundle_b])
        assert matrix.loc["a.xlsx", "a.xlsx"] == 0
        assert matrix.loc["b.xlsx", "b.xlsx"] == 0

    def test_symmetric(self, bundle_a, bundle_b):
        matrix = build_match_matrix([bundle_a, bundle_b])
        assert matrix.loc["a.xlsx", "b.xlsx"] == matrix.loc["b.xlsx", "a.xlsx"]

    def test_single_bundle_returns_1x1(self, bundle_a):
        matrix = build_match_matrix([bundle_a])
        assert matrix.shape == (1, 1)
        assert matrix.iloc[0, 0] == 0
