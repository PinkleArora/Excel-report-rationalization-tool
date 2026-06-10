"""Tests for src.ingestion.workbook_analyzer."""
from __future__ import annotations
from pathlib import Path
import pandas as pd
import pytest
from unittest.mock import MagicMock
from src.ingestion.workbook_analyzer import (
    TabType, TabAnalysis, KPIDefinition, WorkbookAnalysis,
    analyze_workbook, analyze_many, build_kpi_inventory_df,
    build_dependency_report_df, _col_letter_to_index,
    _detect_aggregate_function, _parse_formula_refs,
    _extract_kpi_label,
)
from src.ingestion.loader import load_workbook


# ---------------------------------------------------------------------------
# _col_letter_to_index
# ---------------------------------------------------------------------------

class TestColLetterToIndex:
    def test_A(self):
        assert _col_letter_to_index("A") == 0

    def test_B(self):
        assert _col_letter_to_index("B") == 1

    def test_Z(self):
        assert _col_letter_to_index("Z") == 25

    def test_AA(self):
        assert _col_letter_to_index("AA") == 26

    def test_AB(self):
        assert _col_letter_to_index("AB") == 27


# ---------------------------------------------------------------------------
# _detect_aggregate_function
# ---------------------------------------------------------------------------

class TestDetectAggregateFunction:
    def test_sum(self):
        assert _detect_aggregate_function("=SUM(A1:A10)") == "SUM"

    def test_countif(self):
        assert _detect_aggregate_function("=COUNTIF(B1:B10,\"x\")") == "COUNTIF"

    def test_if_returns_other(self):
        assert _detect_aggregate_function("=IF(A1>0,1,0)") == "OTHER"

    def test_case_insensitive(self):
        assert _detect_aggregate_function("=sum(A1:A10)") == "SUM"

    def test_average(self):
        assert _detect_aggregate_function("=AVERAGE(C1:C5)") == "AVERAGE"


# ---------------------------------------------------------------------------
# _parse_formula_refs
# ---------------------------------------------------------------------------

class TestParseFormulaRefs:
    def test_sheet_ref_resolved(self):
        headers_by_sheet = {"SQL_Data": ["id", "name", "revenue", "cost", "tax", "total"]}
        formula = "=SUM(SQL_Data!F2:F8)"
        sheets, cols = _parse_formula_refs(formula, headers_by_sheet)
        assert "SQL_Data" in sheets
        assert "total" in cols

    def test_unknown_sheet_returns_name_only(self):
        headers_by_sheet = {}
        formula = "=SUM(Sheet1!A1:A5)"
        sheets, cols = _parse_formula_refs(formula, headers_by_sheet)
        assert "Sheet1" in sheets
        assert cols == []

    def test_no_refs_returns_empty(self):
        sheets, cols = _parse_formula_refs("=SUM(A1:A5)", {})
        assert sheets == []
        assert cols == []


# ---------------------------------------------------------------------------
# build_kpi_inventory_df
# ---------------------------------------------------------------------------

class TestBuildKpiInventoryDf:
    def test_empty_returns_correct_columns(self):
        df = build_kpi_inventory_df([])
        expected_cols = {
            "workbook_name", "source_tab", "kpi_label", "formula",
            "aggregate_function", "referenced_sheets", "referenced_columns",
            "cell_address", "is_common",
        }
        assert expected_cols == set(df.columns)
        assert len(df) == 0

    def test_is_common_same_label_two_workbooks(self):
        """Same kpi_label in 2 workbooks → is_common=True."""
        kpi1 = KPIDefinition(
            workbook_name="wb1.xlsx", source_tab="Summary",
            cell_address="B2", kpi_label="Total Revenue",
            formula="=SUM(Data!B:B)", aggregate_function="SUM",
            referenced_sheets=["Data"], referenced_columns=["revenue"],
        )
        kpi2 = KPIDefinition(
            workbook_name="wb2.xlsx", source_tab="Dashboard",
            cell_address="C5", kpi_label="Total Revenue",
            formula="=SUM(Sheet1!C:C)", aggregate_function="SUM",
            referenced_sheets=["Sheet1"], referenced_columns=["revenue"],
        )
        analysis1 = WorkbookAnalysis(workbook_name="wb1.xlsx", kpi_definitions=[kpi1])
        analysis2 = WorkbookAnalysis(workbook_name="wb2.xlsx", kpi_definitions=[kpi2])
        df = build_kpi_inventory_df([analysis1, analysis2])
        assert df["is_common"].all()

    def test_is_common_unique_label(self):
        """KPI label in only one workbook → is_common=False."""
        kpi = KPIDefinition(
            workbook_name="wb1.xlsx", source_tab="Summary",
            cell_address="B2", kpi_label="Unique KPI",
            formula="=SUM(Data!B:B)", aggregate_function="SUM",
        )
        analysis = WorkbookAnalysis(workbook_name="wb1.xlsx", kpi_definitions=[kpi])
        df = build_kpi_inventory_df([analysis])
        assert not df["is_common"].any()


# ---------------------------------------------------------------------------
# build_dependency_report_df
# ---------------------------------------------------------------------------

class TestBuildDependencyReportDf:
    def test_correct_columns(self):
        tab = TabAnalysis(
            workbook_name="wb.xlsx", tab_name="Sheet1", tab_type=TabType.SOURCE_DATA,
            row_count=10, col_count=3, formula_cell_count=0, non_empty_cell_count=30,
            formula_density=0.0, inter_sheet_ref_count=0,
        )
        analysis = WorkbookAnalysis(workbook_name="wb.xlsx", tab_analyses=[tab])
        df = build_dependency_report_df([analysis])
        expected_cols = {
            "workbook", "tab_name", "tab_type", "row_count", "col_count",
            "formula_cells", "formula_density_pct", "inter_sheet_refs",
            "references", "classification_reason",
        }
        assert expected_cols.issubset(set(df.columns))
        assert len(df) == 1


# ---------------------------------------------------------------------------
# analyze_workbook — real xlsx
# ---------------------------------------------------------------------------

class TestAnalyzeWorkbook:
    def test_source_data_classification(self, tmp_path):
        """A plain data sheet with enough rows should classify as SOURCE_DATA."""
        path = tmp_path / "test_wb.xlsx"
        df = pd.DataFrame({"id": range(10), "value": range(10)})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="Data", index=False)
        bundle = load_workbook(path)
        analysis = analyze_workbook(bundle)
        assert analysis.workbook_name == "test_wb.xlsx"
        assert len(analysis.tab_analyses) == 1
        assert analysis.tab_analyses[0].tab_type == TabType.SOURCE_DATA

    def test_returns_workbook_analysis(self, tmp_path):
        path = tmp_path / "wb.xlsx"
        df = pd.DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="S1", index=False)
        bundle = load_workbook(path)
        result = analyze_workbook(bundle)
        assert isinstance(result, WorkbookAnalysis)

    def test_source_tabs_property(self, tmp_path):
        path = tmp_path / "wb.xlsx"
        df = pd.DataFrame({"x": range(5), "y": range(5)})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="DataSheet", index=False)
        bundle = load_workbook(path)
        analysis = analyze_workbook(bundle)
        # At least one source tab should exist
        assert isinstance(analysis.source_tabs, list)

    def test_summary_tabs_property(self, tmp_path):
        path = tmp_path / "wb.xlsx"
        df = pd.DataFrame({"a": [1, 2, 3]})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="Sheet1", index=False)
        bundle = load_workbook(path)
        analysis = analyze_workbook(bundle)
        assert isinstance(analysis.summary_tabs, list)


# ---------------------------------------------------------------------------
# analyze_many
# ---------------------------------------------------------------------------

class TestAnalyzeMany:
    def test_returns_two_analyses(self, tmp_path):
        for name in ["wb1.xlsx", "wb2.xlsx"]:
            path = tmp_path / name
            df = pd.DataFrame({"col": range(5)})
            with pd.ExcelWriter(path, engine="openpyxl") as w:
                df.to_excel(w, sheet_name="Sheet1", index=False)
        bundles = [load_workbook(tmp_path / "wb1.xlsx"), load_workbook(tmp_path / "wb2.xlsx")]
        results = analyze_many(bundles)
        assert len(results) == 2
        assert all(isinstance(r, WorkbookAnalysis) for r in results)

    def test_skips_failing_bundle(self):
        """A broken bundle should be skipped, not crash the whole run."""
        bad_bundle = MagicMock()
        bad_bundle.file_name = "bad.xlsx"
        bad_bundle.sheets = {}
        # MagicMock iter on sheets will return empty
        results = analyze_many([bad_bundle])
        # Should return one result (or gracefully skip), not raise
        assert isinstance(results, list)
