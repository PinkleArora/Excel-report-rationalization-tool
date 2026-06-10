"""Tests for src.ingestion.workbook_analyzer — signal-driven classifier."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from src.ingestion.workbook_analyzer import (
    ClassificationConfig,
    KPIDefinition,
    SignalVector,
    TabAnalysis,
    TabType,
    WorkbookAnalysis,
    _col_letter_to_index,
    _compute_avg_unique_ratio,
    _compute_dtype_homogeneity,
    _detect_aggregate_function,
    _has_aggregate_function,
    _parse_formula_refs,
    analyze_many,
    analyze_workbook,
    build_dependency_report_df,
    build_kpi_inventory_df,
)
from src.ingestion.loader import load_workbook


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _signal(**kwargs) -> SignalVector:
    defaults = dict(
        row_count=100, col_count=5,
        formula_cell_count=0, non_empty_cell_count=500,
        formula_density=0.0,
        inter_sheet_ref_count=0, referenced_sheet_count=0,
        referenced_sheets=[],
        agg_func_count=0,
        avg_unique_ratio=0.5,
        dtype_homogeneity=0.8,
    )
    defaults.update(kwargs)
    return SignalVector(**defaults)


# ---------------------------------------------------------------------------
# _col_letter_to_index
# ---------------------------------------------------------------------------

class TestColLetterToIndex:
    def test_a(self):     assert _col_letter_to_index("A") == 0
    def test_b(self):     assert _col_letter_to_index("B") == 1
    def test_z(self):     assert _col_letter_to_index("Z") == 25
    def test_aa(self):    assert _col_letter_to_index("AA") == 26
    def test_ab(self):    assert _col_letter_to_index("AB") == 27
    def test_lower(self): assert _col_letter_to_index("a") == 0


# ---------------------------------------------------------------------------
# _has_aggregate_function / _detect_aggregate_function
# ---------------------------------------------------------------------------

class TestAggregateDetection:
    def test_sum(self):            assert _has_aggregate_function("=SUM(A1:A10)")
    def test_countif(self):        assert _has_aggregate_function("=COUNTIF(A1:A10,1)")
    def test_if_not_agg(self):     assert not _has_aggregate_function("=IF(A1>0,1,0)")
    def test_detect_sum(self):     assert _detect_aggregate_function("=SUM(A1:A10)") == "SUM"
    def test_detect_avg(self):     assert _detect_aggregate_function("=AVERAGE(A1:A10)") == "AVERAGE"
    def test_detect_other(self):   assert _detect_aggregate_function("=IF(A1,1,0)") == "OTHER"
    def test_case_insensitive(self): assert _detect_aggregate_function("=sum(a1:a5)") == "SUM"


# ---------------------------------------------------------------------------
# _compute_avg_unique_ratio
# ---------------------------------------------------------------------------

class TestAvgUniqueRatio:
    def test_all_unique(self):
        df = pd.DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
        assert _compute_avg_unique_ratio(df) == pytest.approx(1.0)

    def test_all_same(self):
        df = pd.DataFrame({"a": [1, 1, 1]})
        assert _compute_avg_unique_ratio(df) == pytest.approx(1 / 3)

    def test_empty_df(self):
        assert _compute_avg_unique_ratio(pd.DataFrame()) == 0.0


# ---------------------------------------------------------------------------
# _compute_dtype_homogeneity
# ---------------------------------------------------------------------------

class TestDtypeHomogeneity:
    def test_all_same_dtype(self):
        df = pd.DataFrame({"a": [1, 2], "b": [3, 4], "c": [5, 6]})
        assert _compute_dtype_homogeneity(df) == pytest.approx(1.0)

    def test_mixed_dtypes(self):
        df = pd.DataFrame({"a": [1, 2], "b": ["x", "y"]})
        assert _compute_dtype_homogeneity(df) == pytest.approx(0.5)

    def test_empty_df(self):
        assert _compute_dtype_homogeneity(pd.DataFrame()) == 0.0


# ---------------------------------------------------------------------------
# _parse_formula_refs
# ---------------------------------------------------------------------------

class TestParseFormulaRefs:
    def test_simple_ref(self):
        headers = {"DataSheet": ["id", "name", "revenue"]}
        sheets, cols = _parse_formula_refs("=SUM(DataSheet!C2:C10)", headers)
        assert "DataSheet" in sheets
        assert "revenue" in cols

    def test_no_cross_sheet(self):
        sheets, cols = _parse_formula_refs("=SUM(A1:A10)", {})
        assert sheets == []
        assert cols == []

    def test_multiple_refs(self):
        headers = {"S1": ["a", "b"], "S2": ["x", "y"]}
        sheets, cols = _parse_formula_refs("=S1!A1+S2!B1", headers)
        assert set(sheets) == {"S1", "S2"}

    def test_quoted_sheet_name(self):
        headers = {"My Sheet": ["val"]}
        sheets, _ = _parse_formula_refs("='My Sheet'!A1", headers)
        assert "My Sheet" in sheets


# ---------------------------------------------------------------------------
# Multi-signal classifier
# ---------------------------------------------------------------------------

from src.ingestion.workbook_analyzer import _classify_tab


class TestClassifier:
    cfg = ClassificationConfig()

    def test_high_rows_low_formula_is_source(self):
        sv = _signal(row_count=500, formula_density=0.01, dtype_homogeneity=0.9,
                     inter_sheet_ref_count=0)
        tab = _classify_tab(sv, "Sheet1", "wb.xlsx", self.cfg)
        assert tab.tab_type == TabType.SOURCE_DATA

    def test_high_formula_density_with_agg_is_kpi(self):
        sv = _signal(row_count=5, formula_density=0.60, agg_func_count=8,
                     inter_sheet_ref_count=5, referenced_sheet_count=1)
        tab = _classify_tab(sv, "Sheet1", "wb.xlsx", self.cfg)
        assert tab.tab_type == TabType.KPI_SUMMARY

    def test_many_sheet_refs_is_dashboard_or_output(self):
        sv = _signal(row_count=10, formula_density=0.50, agg_func_count=3,
                     inter_sheet_ref_count=20, referenced_sheet_count=5)
        tab = _classify_tab(sv, "Sheet1", "wb.xlsx", self.cfg)
        assert tab.tab_type in (TabType.DASHBOARD, TabType.OUTPUT)

    def test_small_high_unique_is_reference_or_mapping(self):
        # Very few rows + near-total uniqueness → classic lookup/code table
        sv = _signal(row_count=5, col_count=2, formula_density=0.0,
                     avg_unique_ratio=0.99, inter_sheet_ref_count=0)
        tab = _classify_tab(sv, "Sheet1", "wb.xlsx", self.cfg)
        assert tab.tab_type in (TabType.REFERENCE_DATA, TabType.MAPPING_TABLE, TabType.SOURCE_DATA)

    def test_classification_reason_non_empty(self):
        sv = _signal()
        tab = _classify_tab(sv, "Sheet1", "wb.xlsx", self.cfg)
        assert len(tab.classification_reason) > 0

    def test_signal_scores_populated(self):
        sv = _signal(row_count=200, formula_density=0.02)
        tab = _classify_tab(sv, "Sheet1", "wb.xlsx", self.cfg)
        assert isinstance(tab.signal_scores, dict)
        assert len(tab.signal_scores) > 0

    def test_custom_config_changes_score(self):
        """Raising min_source_rows should reduce the source_data score for small sheets."""
        sv = _signal(row_count=5, formula_density=0.0, avg_unique_ratio=0.3)
        strict_cfg = ClassificationConfig(min_source_rows=50)
        relaxed_cfg = ClassificationConfig(min_source_rows=3)
        tab_strict  = _classify_tab(sv, "S", "wb.xlsx", strict_cfg)
        tab_relaxed = _classify_tab(sv, "S", "wb.xlsx", relaxed_cfg)
        assert (tab_relaxed.signal_scores.get("source_data", 0)
                >= tab_strict.signal_scores.get("source_data", 0))

    def test_sheet_name_does_not_affect_classification(self):
        """Identical signals must produce the same type regardless of sheet name."""
        sv = _signal(row_count=500, formula_density=0.01)
        names = ["SQL_Data", "Summary", "Dashboard", "KPI", "Sheet1", "Données"]
        types = {_classify_tab(sv, n, "wb.xlsx", self.cfg).tab_type for n in names}
        assert len(types) == 1, "Classification should not vary by sheet name"


# ---------------------------------------------------------------------------
# WorkbookAnalysis properties
# ---------------------------------------------------------------------------

class TestWorkbookAnalysisProperties:
    def _make_analysis(self) -> WorkbookAnalysis:
        tabs = [
            TabAnalysis("wb", "S1", TabType.SOURCE_DATA,   200, 5,  0, 1000, 0.00, 0),
            TabAnalysis("wb", "S2", TabType.REFERENCE_DATA, 20, 3,  0,   60, 0.00, 0),
            TabAnalysis("wb", "S3", TabType.KPI_SUMMARY,     5, 4, 10,   20, 0.50, 3),
            TabAnalysis("wb", "S4", TabType.DASHBOARD,        8, 6, 15,   48, 0.31, 5),
            TabAnalysis("wb", "S5", TabType.CALCULATION,     50, 8, 40,  400, 0.10, 2),
        ]
        return WorkbookAnalysis(workbook_name="wb", tab_analyses=tabs)

    def test_source_tabs(self):
        assert self._make_analysis().source_tabs == ["S1"]

    def test_reference_tabs(self):
        assert self._make_analysis().reference_tabs == ["S2"]

    def test_summary_tabs(self):
        assert set(self._make_analysis().summary_tabs) == {"S3", "S4"}

    def test_calculation_tabs(self):
        assert self._make_analysis().calculation_tabs == ["S5"]

    def test_data_bearing_default(self):
        result = self._make_analysis().data_bearing_tabs()
        assert set(result) == {"S1", "S2"}

    def test_data_bearing_custom_types(self):
        result = self._make_analysis().data_bearing_tabs(
            types={TabType.SOURCE_DATA, TabType.KPI_SUMMARY}
        )
        assert set(result) == {"S1", "S3"}

    def test_tab_type_lookup(self):
        assert self._make_analysis().tab_type("S3") == TabType.KPI_SUMMARY

    def test_tab_type_missing(self):
        assert self._make_analysis().tab_type("nonexistent") == TabType.UNKNOWN


# ---------------------------------------------------------------------------
# analyze_workbook (integration — no openpyxl formula metadata)
# ---------------------------------------------------------------------------

class TestAnalyzeWorkbook:
    def test_source_tab_classified(self, tmp_path: Path):
        path = tmp_path / "report.xlsx"
        df = pd.DataFrame({
            "id":    range(50),
            "name":  [f"item_{i}" for i in range(50)],
            "value": [float(i) for i in range(50)],
        })
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="Records", index=False)
        bundle = load_workbook(path)
        analysis = analyze_workbook(bundle)
        assert analysis.workbook_name == "report.xlsx"
        assert len(analysis.tab_analyses) == 1
        assert analysis.tab_analyses[0].tab_type == TabType.SOURCE_DATA

    def test_custom_config_accepted(self, tmp_path: Path):
        path = tmp_path / "r.xlsx"
        pd.DataFrame({"a": range(20)}).to_excel(path, index=False)
        bundle = load_workbook(path)
        cfg = ClassificationConfig(min_source_rows=5)
        analysis = analyze_workbook(bundle, config=cfg)
        assert len(analysis.tab_analyses) == 1

    def test_analyze_many_returns_one_per_bundle(self, tmp_path: Path):
        bundles = []
        for i in range(3):
            p = tmp_path / f"wb{i}.xlsx"
            pd.DataFrame({"a": range(20)}).to_excel(p, index=False)
            bundles.append(load_workbook(p))
        assert len(analyze_many(bundles)) == 3

    def test_analyze_many_skips_failures(self):
        from unittest.mock import MagicMock
        bad = MagicMock()
        bad.file_name = "bad.xlsx"
        bad.sheets = None
        assert analyze_many([bad]) == []


# ---------------------------------------------------------------------------
# build_kpi_inventory_df
# ---------------------------------------------------------------------------

class TestBuildKpiInventoryDf:
    def test_empty_returns_correct_columns(self):
        df = build_kpi_inventory_df([])
        assert "is_common" in df.columns
        assert len(df) == 0

    def test_is_common_true_when_label_in_two_workbooks(self):
        kpi_a = KPIDefinition("wb_a.xlsx", "S1", "B2", "Total Revenue",
                              "=SUM(D!A1:A10)", "SUM", [], [])
        kpi_b = KPIDefinition("wb_b.xlsx", "S1", "B2", "Total Revenue",
                              "=SUM(E!A1:A10)", "SUM", [], [])
        kpi_c = KPIDefinition("wb_a.xlsx", "S1", "B3", "Unique KPI",
                              "=COUNT(D!A1:A10)", "COUNT", [], [])
        a = WorkbookAnalysis("wb_a.xlsx", [], [kpi_a, kpi_c])
        b = WorkbookAnalysis("wb_b.xlsx", [], [kpi_b])
        df = build_kpi_inventory_df([a, b])
        assert df[df["kpi_label"] == "Total Revenue"]["is_common"].all()
        assert not df[df["kpi_label"] == "Unique KPI"]["is_common"].any()

    def test_is_common_false_when_single_workbook(self):
        kpi = KPIDefinition("wb.xlsx", "S1", "B2", "Revenue", "=SUM(A1:A10)", "SUM", [], [])
        df = build_kpi_inventory_df([WorkbookAnalysis("wb.xlsx", [], [kpi])])
        assert not df["is_common"].any()


# ---------------------------------------------------------------------------
# build_dependency_report_df
# ---------------------------------------------------------------------------

class TestBuildDependencyReportDf:
    def test_returns_dataframe(self):
        tab = TabAnalysis("wb", "S1", TabType.SOURCE_DATA, 100, 5, 0, 500, 0.0, 0)
        df = build_dependency_report_df([WorkbookAnalysis("wb", [tab])])
        assert isinstance(df, pd.DataFrame)
        assert "tab_type" in df.columns
        assert "formula_density_pct" in df.columns

    def test_one_row_per_tab(self):
        tabs = [
            TabAnalysis("wb", "S1", TabType.SOURCE_DATA, 100, 5, 0, 500, 0.0, 0),
            TabAnalysis("wb", "S2", TabType.KPI_SUMMARY,   5, 3, 10,  15, 0.67, 2),
        ]
        df = build_dependency_report_df([WorkbookAnalysis("wb", tabs)])
        assert len(df) == 2

    def test_empty_analyses(self):
        df = build_dependency_report_df([])
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 0
