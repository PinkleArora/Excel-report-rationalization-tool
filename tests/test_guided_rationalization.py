"""Tests for the guided_rationalization module."""

from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import pytest

from src.ingestion.loader import WorkbookBundle, load_workbook_from_bytes
from src.guided_rationalization.config import RationalizationConfig, WorkbookTabConfig
from src.guided_rationalization.source_analyzer import analyze_source_data
from src.guided_rationalization.kpi_analyzer import analyze_kpi_dependencies
from src.guided_rationalization.workbook_builder import (
    DuplicateColumnError,
    DuplicateFinding,
    SourceFrameProfile,
    build_duplicate_column_analysis_df,
    build_issues_log_df,
    build_master_source_df,
    build_rationalized_workbook_bytes,
    build_workbook_source_analysis_df,
    inspect_frames_for_duplicates,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _make_bundle(file_name: str, sheets: dict[str, pd.DataFrame]) -> WorkbookBundle:
    return WorkbookBundle(source_path=Path(file_name), sheets=sheets, _raw_wb=None)


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
def bundle_with_duplicates():
    """Source tab has two columns that normalise to the same canonical name."""
    df = pd.DataFrame({
        "Policy Number": ["P001", "P002"],
        "Policy_Number": ["P003", "P004"],   # will collide with 'Policy Number' → policy_number
        "Revenue":       [100.0, 200.0],
    })
    return _make_bundle("dup_workbook.xlsx", {"Data": df})


@pytest.fixture
def config_two(bundle_a, bundle_b) -> RationalizationConfig:
    return RationalizationConfig(
        workbook_configs=[
            WorkbookTabConfig("workbook_a.xlsx", "Data", [], "LOB-A"),
            WorkbookTabConfig("workbook_b.xlsx", "Records", [], "LOB-B"),
        ],
        matching_threshold=80.0,
        remove_unused_columns=False,
    )


@pytest.fixture
def config_dup(bundle_with_duplicates) -> RationalizationConfig:
    return RationalizationConfig(
        workbook_configs=[
            WorkbookTabConfig("dup_workbook.xlsx", "Data", [], "LOB-DUP"),
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
# analyze_kpi_dependencies
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
# inspect_frames_for_duplicates
# ---------------------------------------------------------------------------

class TestInspectFramesForDuplicates:
    def test_clean_frames_produce_no_findings(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        # Build frames manually
        frames_with_meta = []
        for bundle in [bundle_a, bundle_b]:
            wb_cfg = config_two.config_for(bundle.file_name)
            df = bundle.sheets[wb_cfg.source_tab].copy()
            frames_with_meta.append((df, bundle.file_name, wb_cfg.source_tab, {}))
        profiles = inspect_frames_for_duplicates(frames_with_meta)
        all_findings = [f for p in profiles for f in p.duplicate_findings]
        assert all_findings == []

    def test_duplicate_frame_detected(self):
        df = pd.DataFrame({"A": [1], "B": [2]})
        df.columns = pd.Index(["col_x", "col_x"])  # force duplicate
        frames_with_meta = [(df, "wb.xlsx", "Sheet1", {"A": "col_x", "B": "col_x"})]
        profiles = inspect_frames_for_duplicates(frames_with_meta)
        assert profiles[0].duplicate_column_count > 0
        assert any(f.column_name == "col_x" for f in profiles[0].duplicate_findings)

    def test_profile_shape_recorded(self, bundle_a, config_two):
        wb_cfg = config_two.config_for("workbook_a.xlsx")
        df = bundle_a.sheets[wb_cfg.source_tab]
        profiles = inspect_frames_for_duplicates(
            [(df, "workbook_a.xlsx", wb_cfg.source_tab, {})]
        )
        assert profiles[0].shape == df.shape


# ---------------------------------------------------------------------------
# DuplicateColumnError
# ---------------------------------------------------------------------------

class TestDuplicateColumnError:
    def test_str_contains_workbook_info(self):
        finding = DuplicateFinding(
            workbook="wb.xlsx",
            source_tab="Data",
            column_name="policy_number",
            duplicate_count=2,
            column_positions=[0, 3],
            original_names=["Policy Number", "Policy_Number"],
            likely_cause="normalisation_collision",
        )
        exc = DuplicateColumnError([finding])
        assert "wb.xlsx" in str(exc)
        assert "policy_number" in str(exc)

    def test_diagnostic_bytes_default_none(self):
        exc = DuplicateColumnError([])
        assert exc.diagnostic_bytes is None

    def test_diagnostic_bytes_attached(self):
        exc = DuplicateColumnError([], diagnostic_bytes=b"xlsx")
        assert exc.diagnostic_bytes == b"xlsx"


# ---------------------------------------------------------------------------
# build_duplicate_column_analysis_df / build_workbook_source_analysis_df
# ---------------------------------------------------------------------------

class TestDiagnosticDataFrames:
    def _finding(self):
        return DuplicateFinding(
            workbook="wb.xlsx",
            source_tab="Data",
            column_name="policy_number",
            duplicate_count=2,
            column_positions=[0, 3],
            original_names=["Policy Number", "Policy_Number"],
            likely_cause="normalisation_collision — test",
        )

    def _profile_with_finding(self):
        f = self._finding()
        return SourceFrameProfile(
            workbook="wb.xlsx",
            source_tab="Data",
            shape=(10, 5),
            all_columns=["policy_number", "revenue", "cost", "policy_number", "region"],
            duplicate_findings=[f],
        )

    def test_dup_analysis_df_has_finding_row(self):
        profile = self._profile_with_finding()
        df = build_duplicate_column_analysis_df([profile])
        assert "policy_number" in df["Duplicate Column Name"].values

    def test_dup_analysis_df_no_findings_returns_placeholder(self):
        clean_profile = SourceFrameProfile(
            workbook="wb.xlsx", source_tab="Data",
            shape=(5, 3), all_columns=["a", "b", "c"],
        )
        df = build_duplicate_column_analysis_df([clean_profile])
        assert "(none)" in df["Workbook"].values

    def test_workbook_source_analysis_columns(self):
        profile = self._profile_with_finding()
        df = build_workbook_source_analysis_df([profile])
        assert "Total Columns" in df.columns
        assert "Duplicate Columns Count" in df.columns
        assert df["Has Duplicates"].iloc[0] == "YES"

    def test_issues_log_high_severity_for_duplicate(self):
        profile = self._profile_with_finding()
        df = build_issues_log_df([profile])
        assert "HIGH" in df["Severity"].values


# ---------------------------------------------------------------------------
# build_master_source_df — returns (df, profiles)
# ---------------------------------------------------------------------------

class TestBuildMasterSourceDf:
    def test_combines_rows(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        master, profiles = build_master_source_df([bundle_a, bundle_b], config_two, source_result)
        assert len(master) == 4

    def test_lineage_columns_present(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        master, _ = build_master_source_df([bundle_a, bundle_b], config_two, source_result)
        assert "Source_Workbook" in master.columns
        assert "Source_Sheet" in master.columns
        assert "LOB_Identifier" in master.columns

    def test_lob_identifier_populated(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        master, _ = build_master_source_df([bundle_a, bundle_b], config_two, source_result)
        assert "LOB-A" in master["LOB_Identifier"].values
        assert "LOB-B" in master["LOB_Identifier"].values

    def test_profiles_returned(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        _, profiles = build_master_source_df([bundle_a, bundle_b], config_two, source_result)
        assert isinstance(profiles, list)
        assert len(profiles) == 2

    def test_empty_bundles_returns_df(self, config_two):
        from src.guided_rationalization.source_analyzer import SourceAnalysisResult
        empty_result = SourceAnalysisResult(
            column_profiles=[], column_mapping={}, excluded_columns=[]
        )
        master, profiles = build_master_source_df([], config_two, empty_result)
        assert isinstance(master, pd.DataFrame)
        assert profiles == []

    def test_raises_duplicate_column_error(self, bundle_with_duplicates, config_dup):
        source_result = analyze_source_data([bundle_with_duplicates], config_dup)
        with pytest.raises(DuplicateColumnError) as exc_info:
            build_master_source_df([bundle_with_duplicates], config_dup, source_result)
        assert len(exc_info.value.findings) > 0

    def test_duplicate_error_identifies_correct_column(self, bundle_with_duplicates, config_dup):
        source_result = analyze_source_data([bundle_with_duplicates], config_dup)
        with pytest.raises(DuplicateColumnError) as exc_info:
            build_master_source_df([bundle_with_duplicates], config_dup, source_result)
        col_names = {f.column_name for f in exc_info.value.findings}
        # "Policy Number" and "Policy_Number" both normalise to "policy_number"
        assert "policy_number" in col_names


# ---------------------------------------------------------------------------
# build_rationalized_workbook_bytes — clean path
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
        # 8 fixed sheets minimum (no KPI tabs configured in config_two)
        assert len(xl.sheet_names) >= 8

    def test_numbered_sheet_names(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        kpi_result = analyze_kpi_dependencies(
            [bundle_a, bundle_b], config_two, source_result.column_mapping
        )
        data = build_rationalized_workbook_bytes(
            [bundle_a, bundle_b], config_two, source_result, kpi_result
        )
        xl = pd.ExcelFile(io.BytesIO(data))
        numbered = [s for s in xl.sheet_names if s[:2].isdigit()]
        assert len(numbered) >= 6  # at least the fixed sheets are numbered

    def test_all_fixed_sheets_present(self, bundle_a, bundle_b, config_two):
        source_result = analyze_source_data([bundle_a, bundle_b], config_two)
        kpi_result = analyze_kpi_dependencies(
            [bundle_a, bundle_b], config_two, source_result.column_mapping
        )
        data = build_rationalized_workbook_bytes(
            [bundle_a, bundle_b], config_two, source_result, kpi_result
        )
        xl = pd.ExcelFile(io.BytesIO(data))
        names = xl.sheet_names
        for expected in (
            "01_Master_Source_Data",
            "03_Source_Mapping",
            "04_Data_Dictionary",
            "05_Reconciliation",
            "06_Issues_Log",
            "07_Duplicate_Column_Analysis",
            "08_Workbook_Source_Analysis",
            "09_Documentation",
        ):
            assert expected in names, f"Missing sheet: {expected}"

    def test_duplicate_raises_with_diagnostic_bytes(
        self, bundle_with_duplicates, config_dup
    ):
        source_result = analyze_source_data([bundle_with_duplicates], config_dup)
        kpi_result = analyze_kpi_dependencies(
            [bundle_with_duplicates], config_dup, source_result.column_mapping
        )
        with pytest.raises(DuplicateColumnError) as exc_info:
            build_rationalized_workbook_bytes(
                [bundle_with_duplicates], config_dup, source_result, kpi_result
            )
        assert exc_info.value.diagnostic_bytes is not None
        assert isinstance(exc_info.value.diagnostic_bytes, bytes)
        # Diagnostic workbook must be readable
        xl = pd.ExcelFile(io.BytesIO(exc_info.value.diagnostic_bytes))
        assert "07_Duplicate_Column_Analysis" in xl.sheet_names


# ---------------------------------------------------------------------------
# load_workbook_from_bytes
# ---------------------------------------------------------------------------

class TestLoadWorkbookFromBytes:
    def test_loads_xlsx_bytes(self, tmp_path):
        path = tmp_path / "test.xlsx"
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="Sheet1", index=False)
        bundle = load_workbook_from_bytes(path.read_bytes(), "test.xlsx")
        assert bundle.file_name == "test.xlsx"
        assert "Sheet1" in bundle.sheets

    def test_dataframe_content_correct(self, tmp_path):
        path = tmp_path / "t.xlsx"
        df = pd.DataFrame({"X": [10, 20]})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="Data", index=False)
        bundle = load_workbook_from_bytes(path.read_bytes(), "t.xlsx")
        assert list(bundle.sheets["Data"]["X"]) == [10, 20]
