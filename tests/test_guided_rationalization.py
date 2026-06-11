"""Tests for the guided_rationalization module."""

from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import pytest
import openpyxl

from src.ingestion.loader import WorkbookBundle, load_workbook_from_bytes
from src.guided_rationalization.config import RationalizationConfig, WorkbookTabConfig
from src.guided_rationalization.source_analyzer import analyze_source_data, ColumnProfile
from src.guided_rationalization.kpi_analyzer import analyze_kpi_dependencies, KpiDependency
from src.guided_rationalization.workbook_builder import (
    build_master_source_df,
    build_rationalized_workbook_bytes,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _make_bundle(file_name: str, sheets: dict[str, pd.DataFrame]) -> WorkbookBundle:
    """Create a minimal WorkbookBundle without needing a real file."""
    fake_path = Path(file_name)
    return WorkbookBundle(source_path=fake_path, sheets=sheets, _raw_wb=None)


@pytest.fixture
def bundle_a():
    df = pd.DataFrame({
        "Region":  ["North", "South"],
        "Revenue": [100.0, 200.0],
        "Cost":    [50.0, 80.0],
    })
    return _make_bundle("workbook_a.xlsx", {"Data": df})


@pytest.fixture
def bundle_b():
    df = pd.DataFrame({
        "Region":    ["East", "West"],
        "Revenue":   [150.0, 120.0],
        "Headcount": [10, 20],
    })
    return _make_bundle("workbook_b.xlsx", {"Records": df})


@pytest.fixture
def config_two(bundle_a, bundle_b) -> RationalizationConfig:
    return RationalizationConfig(
        workbook_configs=[
            WorkbookTabConfig(
                workbook_name="workbook_a.xlsx",
                source_tab="Data",
                kpi_tabs=[],
                lob_identifier="LOB-A",
            ),
            WorkbookTabConfig(
                workbook_name="workbook_b.xlsx",
                source_tab="Records",
                kpi_tabs=[],
                lob_identifier="LOB-B",
            ),
        ],
        matching_threshold=80.0,
        remove_unused_columns=False,
    )


# ---------------------------------------------------------------------------
# RationalizationConfig
# ---------------------------------------------------------------------------

class TestRationalizationConfig:
    def test_config_for_returns_matching(self, config_two):
        cfg = config_two.config_for("workbook_a.xlsx")
        assert cfg is not None
        assert cfg.source_tab == "Data"

    def test_config_for_returns_none_for_unknown(self, config_two):
        assert config_two.config_for("unknown.xlsx") is None

    def test_kpi_tab_name_auto(self, config_two):
        name = config_two.kpi_tab_name_for("workbook_a.xlsx", "Summary")
        assert "workbook_a" in name
        assert "Summary" in name

    def test_kpi_tab_name_explicit(self):
        cfg = RationalizationConfig(
            workbook_configs=[],
            future_kpi_tab_names={"wb.xlsx": "MyKPIs"},
        )
        assert cfg.kpi_tab_name_for("wb.xlsx", "KPI") == "MyKPIs"

    def test_kpi_tab_name_truncated_to_31(self):
        cfg = RationalizationConfig(workbook_configs=[])
        name = cfg.kpi_tab_name_for("very_long_workbook_name_file.xlsx", "LongTabName")
        assert len(name) <= 31


# ---------------------------------------------------------------------------
# analyze_source_data
# ---------------------------------------------------------------------------

class TestAnalyzeSourceData:
    def test_returns_source_analysis_result(self, bundle_a, bundle_b, config_two):
        result = analyze_source_data([bundle_a, bundle_b], config_two)
        assert result is not None

    def test_common_columns_detected(self, bundle_a, bundle_b, config_two):
        result = analyze_source_data([bundle_a, bundle_b], config_two)
        common_names = {p.canonical_name for p in result.common_columns}
        assert "region" in common_names
        assert "revenue" in common_names

    def test_unique_column_detected(self, bundle_a, bundle_b, config_two):
        result = analyze_source_data([bundle_a, bundle_b], config_two)
        unique_names = {p.canonical_name for p in result.unique_columns}
        assert "cost" in unique_names or "headcount" in unique_names

    def test_column_mapping_populated(self, bundle_a, bundle_b, config_two):
        result = analyze_source_data([bundle_a, bundle_b], config_two)
        assert len(result.column_mapping) > 0

    def test_kpi_referenced_marks_columns(self, bundle_a, bundle_b, config_two):
        result = analyze_source_data(
            [bundle_a, bundle_b], config_two,
            kpi_referenced_canonicals={"revenue"},
        )
        revenue_profile = next(
            (p for p in result.column_profiles if p.canonical_name == "revenue"), None
        )
        assert revenue_profile is not None
        assert revenue_profile.is_kpi_referenced is True

    def test_excluded_columns_when_remove_unused(self, bundle_a, bundle_b):
        config = RationalizationConfig(
            workbook_configs=[
                WorkbookTabConfig("workbook_a.xlsx", "Data", [], ""),
                WorkbookTabConfig("workbook_b.xlsx", "Records", [], ""),
            ],
            remove_unused_columns=True,
        )
        result = analyze_source_data(
            [bundle_a, bundle_b], config,
            kpi_referenced_canonicals={"revenue"},
        )
        # Columns not in {"revenue"} should be excluded
        excluded_canonicals = {
            result.column_mapping.get((wb, tab, col), "")
            for wb, tab, col in result.excluded_columns
        }
        assert "revenue" not in excluded_canonicals

    def test_to_dataframe_returns_dataframe(self, bundle_a, bundle_b, config_two):
        result = analyze_source_data([bundle_a, bundle_b], config_two)
        df = result.to_dataframe()
        assert isinstance(df, pd.DataFrame)
        assert "canonical_column" in df.columns

    def test_to_mapping_dataframe_returns_dataframe(self, bundle_a, bundle_b, config_two):
        result = analyze_source_data([bundle_a, bundle_b], config_two)
        df = result.to_mapping_dataframe()
        assert isinstance(df, pd.DataFrame)
        assert "canonical_column" in df.columns


# ---------------------------------------------------------------------------
# analyze_kpi_dependencies (without real formulas — covers no-formula path)
# ---------------------------------------------------------------------------

class TestAnalyzeKpiDependencies:
    def test_empty_result_when_no_kpi_tabs(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        kpi_result = analyze_kpi_dependencies(
            [bundle_a, bundle_b], config_two, source_result.column_mapping
        )
        assert kpi_result.dependencies == []
        assert len(kpi_result.referenced_canonicals) == 0

    def test_unreferenced_includes_all_source_cols(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        kpi_result = analyze_kpi_dependencies(
            [bundle_a, bundle_b], config_two, source_result.column_mapping
        )
        # All source canonicals should be unreferenced when no KPI tabs exist
        all_canonicals = set(source_result.column_mapping.values())
        assert kpi_result.unreferenced_canonicals == all_canonicals

    def test_to_dependency_dataframe(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        kpi_result = analyze_kpi_dependencies(
            [bundle_a, bundle_b], config_two, source_result.column_mapping
        )
        df = kpi_result.to_dependency_dataframe()
        assert isinstance(df, pd.DataFrame)

    def test_to_coverage_dataframe(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        kpi_result = analyze_kpi_dependencies(
            [bundle_a, bundle_b], config_two, source_result.column_mapping
        )
        all_canonicals = set(source_result.column_mapping.values())
        df = kpi_result.to_coverage_dataframe(all_canonicals)
        assert isinstance(df, pd.DataFrame)
        assert "canonical_column" in df.columns
        assert "recommendation" in df.columns


# ---------------------------------------------------------------------------
# build_master_source_df
# ---------------------------------------------------------------------------

class TestBuildMasterSourceDf:
    def test_combines_rows(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        master = build_master_source_df([bundle_a, bundle_b], config_two, source_result)
        assert len(master) == 4  # 2 rows from each workbook

    def test_lineage_columns_present(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        master = build_master_source_df([bundle_a, bundle_b], config_two, source_result)
        assert "Source_Workbook" in master.columns
        assert "Source_Sheet" in master.columns
        assert "LOB_Identifier" in master.columns

    def test_lob_identifier_populated(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        master = build_master_source_df([bundle_a, bundle_b], config_two, source_result)
        assert "LOB-A" in master["LOB_Identifier"].values
        assert "LOB-B" in master["LOB_Identifier"].values

    def test_empty_bundles_returns_df(self, config_two):
        from src.guided_rationalization.source_analyzer import SourceAnalysisResult
        empty_result = SourceAnalysisResult(
            column_profiles=[], column_mapping={}, excluded_columns=[]
        )
        master = build_master_source_df([], config_two, empty_result)
        assert isinstance(master, pd.DataFrame)


# ---------------------------------------------------------------------------
# build_rationalized_workbook_bytes
# ---------------------------------------------------------------------------

class TestBuildRationalizedWorkbookBytes:
    def test_returns_bytes(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        kpi_result = analyze_kpi_dependencies(
            [bundle_a, bundle_b], config_two, source_result.column_mapping
        )
        data = build_rationalized_workbook_bytes(
            [bundle_a, bundle_b], config_two, source_result, kpi_result
        )
        assert isinstance(data, bytes)
        assert len(data) > 0

    def test_valid_xlsx(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        kpi_result = analyze_kpi_dependencies(
            [bundle_a, bundle_b], config_two, source_result.column_mapping
        )
        data = build_rationalized_workbook_bytes(
            [bundle_a, bundle_b], config_two, source_result, kpi_result
        )
        xl = pd.ExcelFile(io.BytesIO(data))
        assert len(xl.sheet_names) >= 1

    def test_master_source_sheet_present(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        kpi_result = analyze_kpi_dependencies(
            [bundle_a, bundle_b], config_two, source_result.column_mapping
        )
        data = build_rationalized_workbook_bytes(
            [bundle_a, bundle_b], config_two, source_result, kpi_result
        )
        xl = pd.ExcelFile(io.BytesIO(data))
        assert config_two.future_source_tab_name in xl.sheet_names

    def test_fixed_sheets_present(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        kpi_result = analyze_kpi_dependencies(
            [bundle_a, bundle_b], config_two, source_result.column_mapping
        )
        data = build_rationalized_workbook_bytes(
            [bundle_a, bundle_b], config_two, source_result, kpi_result
        )
        xl = pd.ExcelFile(io.BytesIO(data))
        for expected in ("Column_Mapping", "KPI_Dependencies", "Column_Coverage", "Documentation"):
            assert expected in xl.sheet_names, f"Missing sheet: {expected}"


# ---------------------------------------------------------------------------
# load_workbook_from_bytes
# ---------------------------------------------------------------------------

class TestLoadWorkbookFromBytes:
    def test_loads_xlsx_bytes(self, tmp_path):
        path = tmp_path / "test.xlsx"
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="Sheet1", index=False)
        data = path.read_bytes()
        bundle = load_workbook_from_bytes(data, "test.xlsx")
        assert bundle.file_name == "test.xlsx"
        assert "Sheet1" in bundle.sheets

    def test_dataframe_content_correct(self, tmp_path):
        path = tmp_path / "t.xlsx"
        df = pd.DataFrame({"X": [10, 20]})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="Data", index=False)
        bundle = load_workbook_from_bytes(path.read_bytes(), "t.xlsx")
        assert list(bundle.sheets["Data"]["X"]) == [10, 20]
