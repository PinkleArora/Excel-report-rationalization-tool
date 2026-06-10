"""Tests for src.schema_matching — normalizer, matcher, and utilities."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from src.ingestion.loader import WorkbookBundle
from src.schema_matching.matcher import (
    ColumnMatch,
    MatchConfidence,
    build_match_matrix,
    detect_many_to_one_mappings,
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


# ---------------------------------------------------------------------------
# MatchConfidence classification
# ---------------------------------------------------------------------------

class TestMatchConfidence:
    def test_exact_match_is_safe_to_merge(self, bundle_a, bundle_b):
        matches = match_all_bundles([bundle_a, bundle_b], threshold=80.0)
        exact = [m for m in matches if m.match_type == "exact"]
        assert all(m.safe_to_merge for m in exact)
        assert all(m.confidence == MatchConfidence.EXACT for m in exact)

    def test_fuzzy_match_not_safe_to_merge_by_default(self):
        """Fuzzy matches must NOT be marked safe_to_merge unless merge_high_confidence=True."""
        from unittest.mock import MagicMock
        import pandas as pd
        src = MagicMock()
        src.file_name = "a.xlsx"
        src.sheets = {"S1": pd.DataFrame({"policy_no": [1]})}
        tgt = MagicMock()
        tgt.file_name = "b.xlsx"
        tgt.sheets = {"S1": pd.DataFrame({"policy_number": [1]})}
        matches = match_workbooks(src, tgt, threshold=70.0, merge_high_confidence=False)
        fuzzy = [m for m in matches if m.match_type == "fuzzy"]
        assert all(not m.safe_to_merge for m in fuzzy)

    def test_high_confidence_fuzzy_safe_when_opted_in(self):
        """Score ≥ 95 + merge_high_confidence=True → safe_to_merge=True."""
        from unittest.mock import MagicMock
        import pandas as pd
        # Use columns with near-identical names to force a high fuzzy score
        src = MagicMock()
        src.file_name = "a.xlsx"
        src.sheets = {"S": pd.DataFrame({"revenue_total": [1]})}
        tgt = MagicMock()
        tgt.file_name = "b.xlsx"
        tgt.sheets = {"S": pd.DataFrame({"revenue_totals": [1]})}
        matches = match_workbooks(src, tgt, threshold=80.0, merge_high_confidence=True)
        if matches:
            high = [m for m in matches if m.score >= 95.0]
            assert all(m.safe_to_merge for m in high)

    def test_review_confidence_level(self):
        """Score 80–94 → MatchConfidence.REVIEW."""
        from unittest.mock import MagicMock
        import pandas as pd
        src = MagicMock()
        src.file_name = "a.xlsx"
        src.sheets = {"S": pd.DataFrame({"gaap_reserve": [1]})}
        tgt = MagicMock()
        tgt.file_name = "b.xlsx"
        tgt.sheets = {"S": pd.DataFrame({"non_tai_gaap_reserve": [1]})}
        matches = match_workbooks(src, tgt, threshold=70.0)
        review = [m for m in matches if m.confidence == MatchConfidence.REVIEW]
        for m in review:
            assert 70.0 <= m.score < 95.0
            assert not m.safe_to_merge


# ---------------------------------------------------------------------------
# Many-to-one detection
# ---------------------------------------------------------------------------

class TestDetectManyToOne:
    def _make_safe_match(self, src_col, tgt_col, score=100.0, match_type="exact"):
        m = ColumnMatch(
            source_workbook="wb_a.xlsx", source_sheet="S1", source_col=src_col,
            target_workbook="wb_b.xlsx", target_sheet="S1", target_col=tgt_col,
            score=score, match_type=match_type,
        )
        if match_type == "exact":
            object.__setattr__(m, "safe_to_merge", True)
            object.__setattr__(m, "confidence", MatchConfidence.EXACT)
        return m

    def test_no_many_to_one_when_distinct_targets(self):
        """Different target columns → no many-to-one issue."""
        matches = [
            self._make_safe_match("Revenue", "Revenue"),
            self._make_safe_match("Cost", "Cost"),
        ]
        issues = detect_many_to_one_mappings(matches)
        assert issues == []

    def test_same_col_name_different_workbooks_not_flagged(self):
        """Same column name from two workbooks mapping to same canonical is fine."""
        m1 = self._make_safe_match("Revenue", "Revenue")
        m2 = ColumnMatch(
            source_workbook="wb_c.xlsx", source_sheet="S1", source_col="Revenue",
            target_workbook="wb_b.xlsx", target_sheet="S1", target_col="Revenue",
            score=100.0, match_type="exact",
        )
        object.__setattr__(m2, "safe_to_merge", True)
        object.__setattr__(m2, "confidence", MatchConfidence.EXACT)
        issues = detect_many_to_one_mappings([m1, m2])
        # "revenue" from two sources — same normalized name, not many-to-one
        assert issues == []

    def test_many_to_one_detected_and_suppressed(self):
        """Two different source cols mapping to same canonical → issue + safe_to_merge=False."""
        m1 = self._make_safe_match("GAAP Reserve", "GAAP Reserve Total")   # gaap_reserve → gaap_reserve_total
        m2 = self._make_safe_match("Non-TAI Reserve", "GAAP Reserve Total") # non_tai_reserve → gaap_reserve_total
        issues = detect_many_to_one_mappings([m1, m2])
        assert len(issues) >= 1
        # Both matches should be downgraded
        assert not m1.safe_to_merge
        assert not m2.safe_to_merge

    def test_many_to_one_issue_has_required_fields(self):
        m1 = self._make_safe_match("GAAP Reserve", "GAAP Reserve Total")
        m2 = self._make_safe_match("Tax Reserve",  "GAAP Reserve Total")
        issues = detect_many_to_one_mappings([m1, m2])
        assert len(issues) >= 1
        issue = issues[0]
        assert "target_canonical" in issue
        assert "description" in issue
        assert "source_columns" in issue


# ---------------------------------------------------------------------------
# KPI rationalization gating
# ---------------------------------------------------------------------------

class TestKpiRationalizationGating:
    def _make_analysis(self, wb_name, label, agg_func, ref_cols):
        from src.ingestion.workbook_analyzer import KPIDefinition, WorkbookAnalysis
        kpi = KPIDefinition(wb_name, "Summary", "B2", label,
                            f"={agg_func}(Data!A1:A10)", agg_func,
                            ["Data"], ref_cols)
        return WorkbookAnalysis(wb_name, [], [kpi])

    def test_same_label_same_func_shared_col_is_candidate(self):
        from src.ingestion.workbook_analyzer import build_kpi_inventory_df
        a = self._make_analysis("wb_a.xlsx", "Total Revenue", "SUM", ["revenue"])
        b = self._make_analysis("wb_b.xlsx", "Total Revenue", "SUM", ["revenue"])
        df = build_kpi_inventory_df([a, b])
        assert df[df["workbook_name"] == "wb_a.xlsx"]["rationalization_candidate"].iloc[0]

    def test_same_label_different_func_not_candidate(self):
        from src.ingestion.workbook_analyzer import build_kpi_inventory_df
        a = self._make_analysis("wb_a.xlsx", "Revenue KPI", "SUM",   ["revenue"])
        b = self._make_analysis("wb_b.xlsx", "Revenue KPI", "COUNT", ["revenue"])
        df = build_kpi_inventory_df([a, b])
        assert not df[df["workbook_name"] == "wb_a.xlsx"]["rationalization_candidate"].iloc[0]

    def test_same_label_same_func_no_shared_cols_not_candidate(self):
        from src.ingestion.workbook_analyzer import build_kpi_inventory_df
        # "Total Reserve" in both workbooks but they reference completely different fields
        a = self._make_analysis("wb_a.xlsx", "Total Reserve", "SUM", ["gaap_reserve"])
        b = self._make_analysis("wb_b.xlsx", "Total Reserve", "SUM", ["captive_reserve"])
        df = build_kpi_inventory_df([a, b])
        assert not df[df["workbook_name"] == "wb_a.xlsx"]["rationalization_candidate"].iloc[0]

    def test_unique_kpi_not_candidate(self):
        from src.ingestion.workbook_analyzer import build_kpi_inventory_df
        a = self._make_analysis("wb_a.xlsx", "LOB-Specific KPI", "SUM", ["net_premium"])
        df = build_kpi_inventory_df([a])
        assert not df["rationalization_candidate"].iloc[0]
        assert not df["is_common"].iloc[0]

    def test_rationalization_note_populated(self):
        from src.ingestion.workbook_analyzer import build_kpi_inventory_df
        a = self._make_analysis("wb_a.xlsx", "Total Revenue", "SUM", ["revenue"])
        b = self._make_analysis("wb_b.xlsx", "Total Revenue", "SUM", ["revenue"])
        df = build_kpi_inventory_df([a, b])
        assert df["rationalization_note"].notna().all()
        assert (df["rationalization_note"].str.len() > 0).all()
