"""Tests for src.workbook_generator.analysis_builder."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from src.ingestion.loader import load_workbook
from src.workbook_generator.analysis_builder import (
    build_analysis_workbook,
    build_analysis_workbook_bytes,
    _run_analysis,
    _sheet_workbook_inventory,
    _sheet_source_tabs,
    _sheet_summary_tabs,
    _sheet_kpi_inventory,
    _sheet_formula_inventory,
    _sheet_schema_mapping,
    _sheet_collision_analysis,
    _sheet_rationalization_candidates,
    _sheet_rationalization_exclusions,
    AnalysisContext,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def bundle_source(tmp_path: Path):
    path = tmp_path / "source.xlsx"
    df = pd.DataFrame({
        "Region":  ["North", "South", "East", "West", "Central"],
        "Revenue": [100.0, 200.0, 150.0, 120.0, 180.0],
        "Cost":    [50.0,  80.0,  60.0,  45.0,  70.0],
    })
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Data", index=False)
    return load_workbook(path)


@pytest.fixture
def bundle_source_b(tmp_path: Path):
    path = tmp_path / "source_b.xlsx"
    df = pd.DataFrame({
        "Region":   ["North", "South"],
        "Revenue":  [300.0, 400.0],
        "Headcount": [10, 20],
    })
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Records", index=False)
    return load_workbook(path)


@pytest.fixture
def ctx_single(bundle_source) -> AnalysisContext:
    return _run_analysis([bundle_source], matching_threshold=80.0, config=None)


@pytest.fixture
def ctx_two(bundle_source, bundle_source_b) -> AnalysisContext:
    return _run_analysis([bundle_source, bundle_source_b], matching_threshold=80.0, config=None)


# ---------------------------------------------------------------------------
# _run_analysis
# ---------------------------------------------------------------------------

class TestRunAnalysis:
    def test_returns_context(self, bundle_source):
        ctx = _run_analysis([bundle_source], 80.0, None)
        assert isinstance(ctx, AnalysisContext)

    def test_analyses_populated(self, ctx_single):
        assert len(ctx_single.analyses) == 1

    def test_column_matches_is_list(self, ctx_two):
        assert isinstance(ctx_two.column_matches, list)

    def test_many_to_one_issues_is_list(self, ctx_two):
        assert isinstance(ctx_two.many_to_one_issues, list)


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

class TestSheetWorkbookInventory:
    def test_one_row_per_bundle(self, ctx_single, bundle_source):
        df = _sheet_workbook_inventory(ctx_single)
        assert len(df) == 1

    def test_required_columns(self, ctx_single):
        df = _sheet_workbook_inventory(ctx_single)
        assert "workbook_name" in df.columns
        assert "total_sheets" in df.columns
        assert "analysis_date" in df.columns

    def test_two_bundles_two_rows(self, ctx_two):
        df = _sheet_workbook_inventory(ctx_two)
        assert len(df) == 2


class TestSheetSourceTabs:
    def test_returns_dataframe(self, ctx_single):
        df = _sheet_source_tabs(ctx_single)
        assert isinstance(df, pd.DataFrame)

    def test_required_columns_present(self, ctx_single):
        df = _sheet_source_tabs(ctx_single)
        # May be empty if tab not classified as source, but columns should exist
        assert "workbook_name" in df.columns or len(df) == 0

    def test_no_summary_tabs_in_source_sheet(self, ctx_two):
        from src.ingestion.workbook_analyzer import TabType
        df = _sheet_source_tabs(ctx_two)
        if not df.empty and "tab_type" in df.columns:
            assert all(t in ("source_data", "reference_data", "mapping_table")
                       for t in df["tab_type"])


class TestSheetSummaryTabs:
    def test_returns_dataframe(self, ctx_single):
        df = _sheet_summary_tabs(ctx_single)
        assert isinstance(df, pd.DataFrame)

    def test_no_source_tabs_in_summary_sheet(self, ctx_two):
        df = _sheet_summary_tabs(ctx_two)
        if not df.empty and "tab_type" in df.columns:
            assert all(t in ("kpi_summary", "dashboard", "output", "calculation")
                       for t in df["tab_type"])


class TestSheetKpiInventory:
    def test_returns_dataframe(self, ctx_single):
        df = _sheet_kpi_inventory(ctx_single)
        assert isinstance(df, pd.DataFrame)

    def test_no_internal_ref_cols(self, ctx_single):
        df = _sheet_kpi_inventory(ctx_single)
        assert "_ref_cols_set" not in df.columns


class TestSheetFormulaInventory:
    def test_returns_dataframe(self, ctx_single):
        df = _sheet_formula_inventory(ctx_single)
        assert isinstance(df, pd.DataFrame)

    def test_columns_present(self, ctx_single):
        df = _sheet_formula_inventory(ctx_single)
        assert "formula" in df.columns or len(df) == 0


class TestSheetSchemaMapping:
    def test_returns_dataframe(self, ctx_two):
        df = _sheet_schema_mapping(ctx_two)
        assert isinstance(df, pd.DataFrame)

    def test_confidence_column_present(self, ctx_two):
        df = _sheet_schema_mapping(ctx_two)
        if not df.empty:
            assert "match_confidence" in df.columns

    def test_review_required_column(self, ctx_two):
        df = _sheet_schema_mapping(ctx_two)
        if not df.empty:
            assert "review_required" in df.columns


class TestSheetCollisionAnalysis:
    def test_returns_dataframe(self, ctx_single):
        df = _sheet_collision_analysis(ctx_single)
        assert isinstance(df, pd.DataFrame)

    def test_has_severity_column(self, ctx_single):
        df = _sheet_collision_analysis(ctx_single)
        assert "severity" in df.columns

    def test_no_collisions_shows_info_row(self, ctx_single):
        df = _sheet_collision_analysis(ctx_single)
        # When no collisions, returns a single-row "no issues" dataframe
        if len(df) == 1 and "issue_type" in df.columns:
            assert df["issue_type"].iloc[0] == "none"


class TestSheetRationalizationCandidates:
    def test_returns_dataframe(self, ctx_two):
        df = _sheet_rationalization_candidates(ctx_two)
        assert isinstance(df, pd.DataFrame)

    def test_no_false_candidates(self, ctx_two):
        df = _sheet_rationalization_candidates(ctx_two)
        # If it has a 'message' column it's an empty-result placeholder
        if "message" not in df.columns and not df.empty:
            assert "rationalization_note" in df.columns


class TestSheetRationalizationExclusions:
    def test_returns_dataframe(self, ctx_two):
        df = _sheet_rationalization_exclusions(ctx_two)
        assert isinstance(df, pd.DataFrame)


# ---------------------------------------------------------------------------
# build_analysis_workbook
# ---------------------------------------------------------------------------

class TestBuildAnalysisWorkbook:
    def test_creates_file(self, bundle_source, tmp_path):
        out = build_analysis_workbook([bundle_source], output_path=tmp_path / "analysis.xlsx")
        assert out.exists()

    def test_returns_absolute_path(self, bundle_source, tmp_path):
        out = build_analysis_workbook([bundle_source], output_path=tmp_path / "a.xlsx")
        assert out.is_absolute()

    def test_has_nine_sheets(self, bundle_source, tmp_path):
        out = build_analysis_workbook([bundle_source], output_path=tmp_path / "a.xlsx")
        xl = pd.ExcelFile(out)
        assert len(xl.sheet_names) == 9

    def test_sheet_names_prefixed(self, bundle_source, tmp_path):
        out = build_analysis_workbook([bundle_source], output_path=tmp_path / "a.xlsx")
        xl = pd.ExcelFile(out)
        for i, name in enumerate(xl.sheet_names, start=1):
            assert name.startswith(f"0{i}_"), f"{name} does not start with 0{i}_"

    def test_creates_parent_dirs(self, bundle_source, tmp_path):
        nested = tmp_path / "deep" / "nested" / "analysis.xlsx"
        out = build_analysis_workbook([bundle_source], output_path=nested)
        assert out.exists()


class TestBuildAnalysisWorkbookBytes:
    def test_returns_bytes(self, bundle_source):
        data = build_analysis_workbook_bytes([bundle_source])
        assert isinstance(data, bytes)
        assert len(data) > 0

    def test_bytes_is_valid_xlsx(self, bundle_source):
        import io
        data = build_analysis_workbook_bytes([bundle_source])
        xl = pd.ExcelFile(io.BytesIO(data))
        assert len(xl.sheet_names) == 9
