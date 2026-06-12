"""Tests for the guided_rationalization module."""

from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import pytest

from src.ingestion.loader import WorkbookBundle, load_workbook_from_bytes
from src.guided_rationalization.config import RationalizationConfig, WorkbookTabConfig
from src.guided_rationalization.source_analyzer import analyze_source_data
from src.guided_rationalization.kpi_analyzer import (
    KpiAnalysisResult, KpiDependency, analyze_kpi_dependencies,
)
from src.guided_rationalization.workbook_builder import (
    DuplicateColumnError,
    DuplicateFinding,
    DuplicateResolution,
    SourceFrameProfile,
    _detect_raw_duplicates,
    _index_to_col_letter,
    _kpis_using_position,
    build_duplicate_column_analysis_df,
    build_issues_log_df,
    build_master_source_df,
    build_rationalized_workbook_bytes,
    build_workbook_source_analysis_df,
    resolve_duplicate_group,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_bundle(file_name: str, sheets: dict[str, pd.DataFrame]) -> WorkbookBundle:
    return WorkbookBundle(source_path=Path(file_name), sheets=sheets, _raw_wb=None)


def _empty_kpi_result() -> KpiAnalysisResult:
    return KpiAnalysisResult(dependencies=[], referenced_canonicals=set(), unreferenced_canonicals=set())


def _kpi_result_with_refs(refs: list[tuple[str, str, str, str]]) -> KpiAnalysisResult:
    """refs: list of (workbook, kpi_tab, cell_address, (source_tab, col_letter))"""
    deps = []
    for wb, kpi_tab, cell, (src_tab, col_letter) in refs:
        deps.append(KpiDependency(
            workbook_name=wb,
            kpi_tab=kpi_tab,
            cell_address=cell,
            kpi_label=cell,
            formula=f"=SUM({src_tab}!{col_letter}:{col_letter})",
            aggregate_function="SUM",
            raw_refs=[(src_tab, col_letter)],
            canonical_source_columns=[],
            refs_source_tab_only=True,
        ))
    return KpiAnalysisResult(
        dependencies=deps,
        referenced_canonicals=set(),
        unreferenced_canonicals=set(),
    )


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

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
            WorkbookTabConfig("workbook_a.xlsx", "Data", []),
            WorkbookTabConfig("workbook_b.xlsx", "Records", []),
        ],
        matching_threshold=80.0,
        remove_unused_columns=False,
    )


# A workbook where two columns normalise to the same canonical ("policy_number")
@pytest.fixture
def bundle_dup_normcol():
    df = pd.DataFrame({
        "Policy Number": ["P001", "P002"],  # pos 0 → col A
        "Policy_Number": ["P003", "P004"],  # pos 1 → col B
        "Revenue":       [100.0, 200.0],
    })
    return _make_bundle("dup_wb.xlsx", {"Data": df})


@pytest.fixture
def config_dup(bundle_dup_normcol) -> RationalizationConfig:
    return RationalizationConfig(
        workbook_configs=[
            WorkbookTabConfig("dup_wb.xlsx", "Data", ["KPIs"]),
        ],
        matching_threshold=80.0,
        remove_unused_columns=False,
    )


# ---------------------------------------------------------------------------
# _index_to_col_letter
# ---------------------------------------------------------------------------

class TestIndexToColLetter:
    def test_first_column(self):
        assert _index_to_col_letter(0) == "A"

    def test_26th_column(self):
        assert _index_to_col_letter(25) == "Z"

    def test_27th_column(self):
        assert _index_to_col_letter(26) == "AA"

    def test_round_trip(self):
        for i in range(50):
            letter = _index_to_col_letter(i)
            assert len(letter) >= 1


# ---------------------------------------------------------------------------
# _detect_raw_duplicates
# ---------------------------------------------------------------------------

class TestDetectRawDuplicates:
    def test_no_duplicates_returns_empty(self, bundle_a, config_two):
        df = bundle_a.sheets["Data"]
        rename_map = {"Region": "region", "Revenue": "revenue", "Cost": "cost"}
        findings = _detect_raw_duplicates(df, rename_map, "workbook_a.xlsx", "Data")
        assert findings == []

    def test_normcol_duplicate_detected(self, bundle_dup_normcol):
        df = bundle_dup_normcol.sheets["Data"]
        rename_map = {
            "Policy Number": "policy_number",
            "Policy_Number": "policy_number",
            "Revenue": "revenue",
        }
        findings = _detect_raw_duplicates(df, rename_map, "dup_wb.xlsx", "Data")
        assert len(findings) == 1
        assert findings[0].canonical_name == "policy_number"
        assert sorted(findings[0].original_positions) == [0, 1]

    def test_finding_records_original_names(self, bundle_dup_normcol):
        df = bundle_dup_normcol.sheets["Data"]
        rename_map = {
            "Policy Number": "policy_number",
            "Policy_Number": "policy_number",
            "Revenue": "revenue",
        }
        findings = _detect_raw_duplicates(df, rename_map, "dup_wb.xlsx", "Data")
        assert set(findings[0].original_col_names) == {"Policy Number", "Policy_Number"}


# ---------------------------------------------------------------------------
# _kpis_using_position
# ---------------------------------------------------------------------------

class TestKpisUsingPosition:
    def test_position_referenced(self):
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),  # col A = position 0
        ])
        using = _kpis_using_position("dup_wb.xlsx", "Data", 0, kpi_result)
        assert "B2" in using

    def test_position_not_referenced(self):
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),
        ])
        using = _kpis_using_position("dup_wb.xlsx", "Data", 1, kpi_result)  # col B not referenced
        assert using == []

    def test_different_workbook_not_matched(self):
        kpi_result = _kpi_result_with_refs([
            ("other_wb.xlsx", "KPIs", "B2", ("Data", "A")),
        ])
        using = _kpis_using_position("dup_wb.xlsx", "Data", 0, kpi_result)
        assert using == []


# ---------------------------------------------------------------------------
# resolve_duplicate_group — Scenario A
# ---------------------------------------------------------------------------

class TestScenarioA:
    """One position used by KPIs, other not → keep used, remove unused."""

    def _finding(self):
        return DuplicateFinding(
            workbook="dup_wb.xlsx", source_tab="Data",
            canonical_name="policy_number", duplicate_count=2,
            original_positions=[0, 1],
            original_col_names=["Policy Number", "Policy_Number"],
            likely_cause="normalisation_collision",
        )

    def test_used_position_is_kept(self):
        # Position 0 (col A) used by a KPI
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),
        ])
        df = pd.DataFrame({"Policy Number": ["P1"], "Policy_Number": ["P2"], "Rev": [1]})
        resolutions = resolve_duplicate_group(self._finding(), kpi_result, df)
        kept = [r for r in resolutions if r.action == "KEEP"]
        removed = [r for r in resolutions if r.action == "REMOVE"]
        assert len(kept) == 1
        assert kept[0].original_position == 0
        assert len(removed) == 1
        assert removed[0].original_position == 1

    def test_all_resolutions_scenario_a(self):
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),
        ])
        df = pd.DataFrame({"Policy Number": ["P1"], "Policy_Number": ["P2"], "Rev": [1]})
        resolutions = resolve_duplicate_group(self._finding(), kpi_result, df)
        assert all(r.scenario == "A" for r in resolutions)


# ---------------------------------------------------------------------------
# resolve_duplicate_group — Scenario B
# ---------------------------------------------------------------------------

class TestScenarioB:
    """Neither position used by KPIs → remove all."""

    def _finding(self):
        return DuplicateFinding(
            workbook="dup_wb.xlsx", source_tab="Data",
            canonical_name="policy_number", duplicate_count=2,
            original_positions=[0, 1],
            original_col_names=["Policy Number", "Policy_Number"],
            likely_cause="normalisation_collision",
        )

    def test_all_removed(self):
        kpi_result = _empty_kpi_result()
        df = pd.DataFrame({"Policy Number": ["P1"], "Policy_Number": ["P2"], "Rev": [1]})
        resolutions = resolve_duplicate_group(self._finding(), kpi_result, df)
        assert all(r.action == "REMOVE" for r in resolutions)
        assert all(r.scenario == "B" for r in resolutions)


# ---------------------------------------------------------------------------
# resolve_duplicate_group — Scenario C
# ---------------------------------------------------------------------------

class TestScenarioC:
    """Both positions used by different KPIs → MANUAL_REVIEW."""

    def _finding(self):
        return DuplicateFinding(
            workbook="dup_wb.xlsx", source_tab="Data",
            canonical_name="policy_number", duplicate_count=2,
            original_positions=[0, 1],
            original_col_names=["Policy Number", "Policy_Number"],
            likely_cause="normalisation_collision",
        )

    def test_all_manual_review(self):
        # Both col A and col B referenced by different KPIs
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),
            ("dup_wb.xlsx", "KPIs", "C3", ("Data", "B")),
        ])
        df = pd.DataFrame({"Policy Number": ["P1"], "Policy_Number": ["P2"], "Rev": [1]})
        resolutions = resolve_duplicate_group(self._finding(), kpi_result, df)
        assert all(r.scenario == "C" for r in resolutions)
        assert all(r.action == "MANUAL_REVIEW" for r in resolutions)


# ---------------------------------------------------------------------------
# resolve_duplicate_group — Scenario D
# ---------------------------------------------------------------------------

class TestScenarioD:
    """Both positions used by same KPIs with identical values → keep first, remove rest."""

    def _finding(self):
        return DuplicateFinding(
            workbook="dup_wb.xlsx", source_tab="Data",
            canonical_name="policy_number", duplicate_count=2,
            original_positions=[0, 1],
            original_col_names=["Policy Number", "Policy_Number"],
            likely_cause="normalisation_collision",
        )

    def test_identical_values_same_kpis(self):
        # Both col A and col B referenced by the SAME KPI, with identical values
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "B")),  # same cell B2 references both
        ])
        df = pd.DataFrame({
            "Policy Number": ["P1", "P2"],
            "Policy_Number": ["P1", "P2"],  # identical values
            "Rev": [1, 2],
        })
        resolutions = resolve_duplicate_group(self._finding(), kpi_result, df)
        kept = [r for r in resolutions if r.action == "KEEP"]
        removed = [r for r in resolutions if r.action == "REMOVE"]
        assert len(kept) == 1
        assert kept[0].original_position == 0
        assert all(r.scenario == "D" for r in resolutions)


# ---------------------------------------------------------------------------
# RationalizationConfig
# ---------------------------------------------------------------------------

class TestRationalizationConfig:
    def test_config_for_returns_matching(self, config_two):
        assert config_two.config_for("workbook_a.xlsx").source_tab == "Data"

    def test_config_for_returns_none_for_unknown(self, config_two):
        assert config_two.config_for("unknown.xlsx") is None

    def test_kpi_tab_name_truncated_to_31(self):
        cfg = RationalizationConfig(workbook_configs=[])
        assert len(cfg.kpi_tab_name_for("very_long_workbook_name_file.xlsx", "LongTabName")) <= 31


# ---------------------------------------------------------------------------
# analyze_source_data
# ---------------------------------------------------------------------------

class TestAnalyzeSourceData:
    def test_common_columns_detected(self, bundle_a, bundle_b, config_two):
        result = analyze_source_data([bundle_a, bundle_b], config_two)
        common_names = {p.canonical_name for p in result.common_columns}
        assert "region" in common_names and "revenue" in common_names

    def test_column_mapping_populated(self, bundle_a, bundle_b, config_two):
        result = analyze_source_data([bundle_a, bundle_b], config_two)
        assert len(result.column_mapping) > 0

    def test_kpi_referenced_marks_columns(self, bundle_a, bundle_b, config_two):
        result = analyze_source_data(
            [bundle_a, bundle_b], config_two, kpi_referenced_canonicals={"revenue"}
        )
        rev = next(p for p in result.column_profiles if p.canonical_name == "revenue")
        assert rev.is_kpi_referenced is True

    def test_to_dataframe_has_canonical_column(self, bundle_a, bundle_b, config_two):
        df = analyze_source_data([bundle_a, bundle_b], config_two).to_dataframe()
        assert "canonical_column" in df.columns


# ---------------------------------------------------------------------------
# analyze_kpi_dependencies
# ---------------------------------------------------------------------------

class TestAnalyzeKpiDependencies:
    def test_empty_when_no_kpi_tabs(self, bundle_a, bundle_b, config_two):
        sr = analyze_source_data([bundle_a, bundle_b], config_two)
        kr = analyze_kpi_dependencies([bundle_a, bundle_b], config_two, sr.column_mapping)
        assert kr.dependencies == []

    def test_to_coverage_dataframe_has_recommendation(self, bundle_a, bundle_b, config_two):
        sr = analyze_source_data([bundle_a, bundle_b], config_two)
        kr = analyze_kpi_dependencies([bundle_a, bundle_b], config_two, sr.column_mapping)
        df = kr.to_coverage_dataframe(set(sr.column_mapping.values()))
        assert "recommendation" in df.columns


# ---------------------------------------------------------------------------
# build_master_source_df — clean path
# ---------------------------------------------------------------------------

class TestBuildMasterSourceDfClean:
    def test_combines_rows(self, bundle_a, bundle_b, config_two):
        sr = analyze_source_data([bundle_a, bundle_b], config_two)
        master, _ = build_master_source_df([bundle_a, bundle_b], config_two, sr, _empty_kpi_result())
        assert len(master) == 4

    def test_lineage_columns_present(self, bundle_a, bundle_b, config_two):
        sr = analyze_source_data([bundle_a, bundle_b], config_two)
        master, _ = build_master_source_df([bundle_a, bundle_b], config_two, sr, _empty_kpi_result())
        for col in ("Source_Workbook", "Source_Sheet"):
            assert col in master.columns
        assert "LOB_Identifier" not in master.columns

    def test_empty_bundles_returns_df(self, config_two):
        from src.guided_rationalization.source_analyzer import SourceAnalysisResult
        empty_sr = SourceAnalysisResult(column_profiles=[], column_mapping={}, excluded_columns=[])
        master, profiles = build_master_source_df([], config_two, empty_sr, _empty_kpi_result())
        assert isinstance(master, pd.DataFrame)
        assert profiles == []


# ---------------------------------------------------------------------------
# build_master_source_df — scenario A: one position used, one unused
# ---------------------------------------------------------------------------

class TestBuildMasterSourceScenarioA:
    def test_unused_duplicate_removed_automatically(self, bundle_dup_normcol, config_dup):
        sr = analyze_source_data([bundle_dup_normcol], config_dup)
        # Only col A (position 0, "Policy Number") is referenced by a KPI
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),
        ])
        master, profiles = build_master_source_df(
            [bundle_dup_normcol], config_dup, sr, kpi_result
        )
        # Should succeed without raising
        assert isinstance(master, pd.DataFrame)
        # Only one policy_number column in output
        assert list(master.columns).count("policy_number") == 1

    def test_resolution_scenario_a_in_profile(self, bundle_dup_normcol, config_dup):
        sr = analyze_source_data([bundle_dup_normcol], config_dup)
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),
        ])
        _, profiles = build_master_source_df(
            [bundle_dup_normcol], config_dup, sr, kpi_result
        )
        assert any(r.scenario == "A" for p in profiles for r in p.resolutions)


# ---------------------------------------------------------------------------
# build_master_source_df — scenario B: neither used
# ---------------------------------------------------------------------------

class TestBuildMasterSourceScenarioB:
    def test_both_removed_automatically(self, bundle_dup_normcol, config_dup):
        sr = analyze_source_data([bundle_dup_normcol], config_dup)
        # No KPI references at all
        master, profiles = build_master_source_df(
            [bundle_dup_normcol], config_dup, sr, _empty_kpi_result()
        )
        assert isinstance(master, pd.DataFrame)
        assert "policy_number" not in master.columns

    def test_resolution_scenario_b_in_profile(self, bundle_dup_normcol, config_dup):
        sr = analyze_source_data([bundle_dup_normcol], config_dup)
        _, profiles = build_master_source_df(
            [bundle_dup_normcol], config_dup, sr, _empty_kpi_result()
        )
        assert any(r.scenario == "B" for p in profiles for r in p.resolutions)


# ---------------------------------------------------------------------------
# build_master_source_df — scenario C: both used by different KPIs → raises
# ---------------------------------------------------------------------------

class TestBuildMasterSourceScenarioC:
    def test_raises_duplicate_column_error(self, bundle_dup_normcol, config_dup):
        sr = analyze_source_data([bundle_dup_normcol], config_dup)
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),
            ("dup_wb.xlsx", "KPIs", "C3", ("Data", "B")),
        ])
        with pytest.raises(DuplicateColumnError) as exc_info:
            build_master_source_df([bundle_dup_normcol], config_dup, sr, kpi_result)
        assert any(r.scenario == "C" for r in exc_info.value.resolutions)

    def test_error_message_names_workbook(self, bundle_dup_normcol, config_dup):
        sr = analyze_source_data([bundle_dup_normcol], config_dup)
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),
            ("dup_wb.xlsx", "KPIs", "C3", ("Data", "B")),
        ])
        with pytest.raises(DuplicateColumnError) as exc_info:
            build_master_source_df([bundle_dup_normcol], config_dup, sr, kpi_result)
        assert "dup_wb.xlsx" in str(exc_info.value)


# ---------------------------------------------------------------------------
# Diagnostic DataFrames
# ---------------------------------------------------------------------------

class TestDiagnosticDataFrames:
    def _profile_with_scenario_a(self):
        finding = DuplicateFinding(
            workbook="wb.xlsx", source_tab="Data",
            canonical_name="policy_number", duplicate_count=2,
            original_positions=[0, 1],
            original_col_names=["Policy Number", "Policy_Number"],
            likely_cause="normalisation_collision",
        )
        resolutions = [
            DuplicateResolution(
                workbook="wb.xlsx", source_tab="Data", canonical_name="policy_number",
                original_position=0, original_col_name="Policy Number",
                scenario="A", action="KEEP", kpis_using=["B2"],
                reason="Referenced by KPI B2.",
            ),
            DuplicateResolution(
                workbook="wb.xlsx", source_tab="Data", canonical_name="policy_number",
                original_position=1, original_col_name="Policy_Number",
                scenario="A", action="REMOVE", kpis_using=[],
                reason="Not referenced by any KPI.",
            ),
        ]
        return SourceFrameProfile(
            workbook="wb.xlsx", source_tab="Data",
            shape=(5, 3), all_columns=["policy_number", "revenue"],
            duplicate_findings=[finding], resolutions=resolutions,
        )

    def test_dup_analysis_has_kpi_usage_column(self):
        df = build_duplicate_column_analysis_df([self._profile_with_scenario_a()])
        assert "KPI(s) Using Column" in df.columns

    def test_dup_analysis_has_scenario_column(self):
        df = build_duplicate_column_analysis_df([self._profile_with_scenario_a()])
        assert "Scenario" in df.columns

    def test_dup_analysis_has_action_column(self):
        df = build_duplicate_column_analysis_df([self._profile_with_scenario_a()])
        assert "Recommended Action" in df.columns

    def test_dup_analysis_no_findings_returns_placeholder(self):
        clean = SourceFrameProfile(
            workbook="wb.xlsx", source_tab="Data",
            shape=(5, 3), all_columns=["a", "b", "c"],
        )
        df = build_duplicate_column_analysis_df([clean])
        assert "(none)" in df["Workbook"].values

    def test_workbook_source_analysis_has_status(self):
        df = build_workbook_source_analysis_df([self._profile_with_scenario_a()])
        assert "Status" in df.columns
        assert "AUTO-RESOLVED" in df["Status"].values

    def test_issues_log_info_for_auto_removed(self):
        df = build_issues_log_df([self._profile_with_scenario_a()])
        # REMOVE entries become INFO, not HIGH
        severities = set(df["Severity"].tolist())
        assert "HIGH" not in severities

    def test_issues_log_high_for_scenario_c(self):
        finding = DuplicateFinding(
            workbook="wb.xlsx", source_tab="Data",
            canonical_name="policy_number", duplicate_count=2,
            original_positions=[0, 1],
            original_col_names=["A", "B"],
            likely_cause="test",
        )
        res_c = [
            DuplicateResolution(
                workbook="wb.xlsx", source_tab="Data", canonical_name="policy_number",
                original_position=p, original_col_name=n,
                scenario="C", action="MANUAL_REVIEW", kpis_using=["B2"],
                reason="Both used by different KPIs.",
            )
            for p, n in [(0, "A"), (1, "B")]
        ]
        profile = SourceFrameProfile(
            workbook="wb.xlsx", source_tab="Data",
            shape=(5, 3), all_columns=["policy_number", "policy_number"],
            duplicate_findings=[finding], resolutions=res_c,
        )
        df = build_issues_log_df([profile])
        assert "HIGH" in df["Severity"].values


# ---------------------------------------------------------------------------
# build_rationalized_workbook_bytes — end-to-end
# ---------------------------------------------------------------------------

class TestBuildRationalizedWorkbookBytes:
    def test_clean_input_returns_bytes(self, bundle_a, bundle_b, config_two):
        sr = analyze_source_data([bundle_a, bundle_b], config_two)
        kr = analyze_kpi_dependencies([bundle_a, bundle_b], config_two, sr.column_mapping)
        data = build_rationalized_workbook_bytes([bundle_a, bundle_b], config_two, sr, kr)
        assert isinstance(data, bytes) and len(data) > 0

    def test_all_fixed_sheets_present(self, bundle_a, bundle_b, config_two):
        sr = analyze_source_data([bundle_a, bundle_b], config_two)
        kr = analyze_kpi_dependencies([bundle_a, bundle_b], config_two, sr.column_mapping)
        data = build_rationalized_workbook_bytes([bundle_a, bundle_b], config_two, sr, kr)
        names = pd.ExcelFile(io.BytesIO(data)).sheet_names
        # config_two has no KPI tabs, so base = 2 (01_Master_Source_Data + no summary sheets)
        for expected in (
            "Master_Source_Data",   # default future_source_tab_name
            "02_Source_Mapping",
            "03_Reconciliation",
            "04_Issues_Log",
            "05_Documentation",
        ):
            assert expected in names, f"Missing: {expected}"

    def test_scenario_a_auto_resolved_no_raise(self, bundle_dup_normcol, config_dup):
        sr = analyze_source_data([bundle_dup_normcol], config_dup)
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),
        ])
        # Should NOT raise — Scenario A resolves automatically
        data = build_rationalized_workbook_bytes(
            [bundle_dup_normcol], config_dup, sr, kpi_result
        )
        assert isinstance(data, bytes)

    def test_scenario_b_auto_resolved_no_raise(self, bundle_dup_normcol, config_dup):
        sr = analyze_source_data([bundle_dup_normcol], config_dup)
        # Should NOT raise — Scenario B (neither used) resolves automatically
        data = build_rationalized_workbook_bytes(
            [bundle_dup_normcol], config_dup, sr, _empty_kpi_result()
        )
        assert isinstance(data, bytes)

    def test_scenario_c_raises_with_diagnostic_bytes(self, bundle_dup_normcol, config_dup):
        sr = analyze_source_data([bundle_dup_normcol], config_dup)
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),
            ("dup_wb.xlsx", "KPIs", "C3", ("Data", "B")),
        ])
        with pytest.raises(DuplicateColumnError) as exc_info:
            build_rationalized_workbook_bytes(
                [bundle_dup_normcol], config_dup, sr, kpi_result
            )
        assert exc_info.value.diagnostic_bytes is not None
        xl = pd.ExcelFile(io.BytesIO(exc_info.value.diagnostic_bytes))
        assert any("Duplicate_Column_Analysis" in s for s in xl.sheet_names)

    def test_scenario_c_error_carries_resolutions(self, bundle_dup_normcol, config_dup):
        sr = analyze_source_data([bundle_dup_normcol], config_dup)
        kpi_result = _kpi_result_with_refs([
            ("dup_wb.xlsx", "KPIs", "B2", ("Data", "A")),
            ("dup_wb.xlsx", "KPIs", "C3", ("Data", "B")),
        ])
        with pytest.raises(DuplicateColumnError) as exc_info:
            build_rationalized_workbook_bytes(
                [bundle_dup_normcol], config_dup, sr, kpi_result
            )
        blocking = [r for r in exc_info.value.resolutions if r.scenario == "C"]
        assert len(blocking) > 0


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
