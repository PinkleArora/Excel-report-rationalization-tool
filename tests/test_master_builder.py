"""Tests for src.workbook_generator.master_builder — end-to-end pipeline."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from src.ingestion.loader import WorkbookBundle, load_workbook
from src.workbook_generator.master_builder import (
    ALL_SHEET_NAMES,
    MasterWorkbookContext,
    SHEET_DATA_DICTIONARY,
    SHEET_ISSUES_LOG,
    SHEET_KPI_SUMMARY,
    SHEET_MASTER_DATA,
    SHEET_RATIONALIZATION,
    SHEET_RECONCILIATION,
    SHEET_SOP,
    SHEET_SOURCE_MAPPING,
    _build_issues_log,
    _build_kpi_summary,
    _build_sop,
    _stage_consolidate,
    _stage_data_dictionary,
    _stage_profile,
    _stage_rationalization,
    _stage_reconciliation,
    _stage_schema_matching,
    _stage_source_mapping,
    build_master_workbook,
    build_master_workbook_bytes,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def bundle_sales(tmp_path: Path) -> WorkbookBundle:
    path = tmp_path / "sales_report.xlsx"
    df = pd.DataFrame({
        "Region": ["North", "South", "East"],
        "Revenue": [100.0, 200.0, 150.0],
        "Cost":    [50.0,  80.0,  60.0],
    })
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Q1", index=False)
    return load_workbook(path)


@pytest.fixture
def bundle_hr(tmp_path: Path) -> WorkbookBundle:
    path = tmp_path / "hr_report.xlsx"
    df = pd.DataFrame({
        "Department": ["Eng", "Finance"],
        "Headcount":  [42,    15],
        "Cost":       [420000.0, 150000.0],
    })
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Headcount", index=False)
    return load_workbook(path)


@pytest.fixture
def ctx_two_bundles(bundle_sales, bundle_hr) -> MasterWorkbookContext:
    ctx = MasterWorkbookContext(bundles=[bundle_sales, bundle_hr])
    _stage_profile(ctx)
    _stage_schema_matching(ctx, 80.0)
    _stage_consolidate(ctx)
    _stage_source_mapping(ctx)
    _stage_data_dictionary(ctx)
    _stage_reconciliation(ctx, 0.01)
    _stage_rationalization(ctx)
    return ctx


# ---------------------------------------------------------------------------
# ALL_SHEET_NAMES constant
# ---------------------------------------------------------------------------

class TestSheetNames:
    def test_count(self):
        assert len(ALL_SHEET_NAMES) == 8

    def test_ordered_prefixes(self):
        for i, name in enumerate(ALL_SHEET_NAMES, start=1):
            assert name.startswith(f"0{i}_"), f"{name} does not start with 0{i}_"


# ---------------------------------------------------------------------------
# Pipeline stages (unit-level)
# ---------------------------------------------------------------------------

class TestPipelineStages:
    def test_stage_profile_populates_metas(self, bundle_sales, bundle_hr):
        ctx = MasterWorkbookContext(bundles=[bundle_sales, bundle_hr])
        _stage_profile(ctx)
        assert len(ctx.workbook_metas) == 2

    def test_stage_schema_matching_populates_matches(self, bundle_sales, bundle_hr):
        ctx = MasterWorkbookContext(bundles=[bundle_sales, bundle_hr])
        _stage_profile(ctx)
        _stage_schema_matching(ctx, 80.0)
        assert isinstance(ctx.column_matches, list)

    def test_stage_consolidate_produces_master_df(self, bundle_sales, bundle_hr):
        ctx = MasterWorkbookContext(bundles=[bundle_sales, bundle_hr])
        _stage_profile(ctx)
        _stage_schema_matching(ctx, 80.0)
        _stage_consolidate(ctx)
        assert not ctx.master_df.empty

    def test_stage_reconciliation_produces_results(self, ctx_two_bundles):
        assert isinstance(ctx_two_bundles.reconciliation_results, list)

    def test_stage_rationalization_one_rec_per_bundle(self, ctx_two_bundles):
        assert len(ctx_two_bundles.rationalization_recs) == 2


# ---------------------------------------------------------------------------
# build_master_workbook
# ---------------------------------------------------------------------------

class TestBuildMasterWorkbook:
    def test_creates_file(self, bundle_sales, bundle_hr, tmp_path):
        out = build_master_workbook(
            [bundle_sales, bundle_hr],
            output_path=tmp_path / "master.xlsx",
        )
        assert out.exists()

    def test_returns_absolute_path(self, bundle_sales, tmp_path):
        out = build_master_workbook([bundle_sales], output_path=tmp_path / "m.xlsx")
        assert out.is_absolute()

    def test_has_all_eight_sheets(self, bundle_sales, bundle_hr, tmp_path):
        out = build_master_workbook(
            [bundle_sales, bundle_hr],
            output_path=tmp_path / "master.xlsx",
        )
        xl = pd.ExcelFile(out)
        assert set(xl.sheet_names) == set(ALL_SHEET_NAMES)

    def test_master_data_has_lineage_columns(self, bundle_sales, bundle_hr, tmp_path):
        out = build_master_workbook(
            [bundle_sales, bundle_hr],
            output_path=tmp_path / "master.xlsx",
        )
        df = pd.read_excel(out, sheet_name=SHEET_MASTER_DATA)
        assert "_source_workbook" in df.columns
        assert "_source_sheet" in df.columns

    def test_master_data_row_count(self, bundle_sales, bundle_hr, tmp_path):
        out = build_master_workbook(
            [bundle_sales, bundle_hr],
            output_path=tmp_path / "master.xlsx",
        )
        df = pd.read_excel(out, sheet_name=SHEET_MASTER_DATA)
        # 3 rows from sales + 2 rows from hr = 5
        assert len(df) == 5

    def test_source_mapping_not_empty(self, bundle_sales, bundle_hr, tmp_path):
        out = build_master_workbook(
            [bundle_sales, bundle_hr],
            output_path=tmp_path / "master.xlsx",
        )
        df = pd.read_excel(out, sheet_name=SHEET_SOURCE_MAPPING)
        assert len(df) > 0

    def test_data_dictionary_not_empty(self, bundle_sales, bundle_hr, tmp_path):
        out = build_master_workbook(
            [bundle_sales, bundle_hr],
            output_path=tmp_path / "master.xlsx",
        )
        df = pd.read_excel(out, sheet_name=SHEET_DATA_DICTIONARY)
        assert len(df) > 0

    def test_kpi_summary_not_empty(self, bundle_sales, bundle_hr, tmp_path):
        out = build_master_workbook(
            [bundle_sales, bundle_hr],
            output_path=tmp_path / "master.xlsx",
        )
        df = pd.read_excel(out, sheet_name=SHEET_KPI_SUMMARY)
        assert len(df) > 0

    def test_reconciliation_sheet_present(self, bundle_sales, bundle_hr, tmp_path):
        out = build_master_workbook(
            [bundle_sales, bundle_hr],
            output_path=tmp_path / "master.xlsx",
        )
        df = pd.read_excel(out, sheet_name=SHEET_RECONCILIATION)
        assert isinstance(df, pd.DataFrame)

    def test_sop_sheet_has_rows(self, bundle_sales, tmp_path):
        out = build_master_workbook([bundle_sales], output_path=tmp_path / "m.xlsx")
        df = pd.read_excel(out, sheet_name=SHEET_SOP)
        assert len(df) >= 8

    def test_creates_parent_dirs(self, bundle_sales, tmp_path):
        nested = tmp_path / "a" / "b" / "master.xlsx"
        out = build_master_workbook([bundle_sales], output_path=nested)
        assert out.exists()

    def test_raises_on_empty_bundles(self, tmp_path):
        with pytest.raises(ValueError, match="at least one"):
            build_master_workbook([], output_path=tmp_path / "m.xlsx")

    def test_single_bundle_works(self, bundle_sales, tmp_path):
        out = build_master_workbook([bundle_sales], output_path=tmp_path / "m.xlsx")
        assert out.exists()

    def test_default_output_path_created(self, bundle_sales, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        out = build_master_workbook([bundle_sales])
        assert out.name == "master_workbook.xlsx"
        assert out.exists()


# ---------------------------------------------------------------------------
# build_master_workbook_bytes
# ---------------------------------------------------------------------------

class TestBuildMasterWorkbookBytes:
    def test_returns_bytes(self, bundle_sales, bundle_hr):
        data = build_master_workbook_bytes([bundle_sales, bundle_hr])
        assert isinstance(data, bytes)
        assert len(data) > 0

    def test_bytes_is_valid_xlsx(self, bundle_sales, bundle_hr):
        import io
        data = build_master_workbook_bytes([bundle_sales, bundle_hr])
        xl = pd.ExcelFile(io.BytesIO(data))
        assert set(xl.sheet_names) == set(ALL_SHEET_NAMES)


# ---------------------------------------------------------------------------
# Sheet-specific helpers
# ---------------------------------------------------------------------------

class TestBuildKpiSummary:
    def test_contains_all_sources_row(self, ctx_two_bundles):
        df = _build_kpi_summary(ctx_two_bundles)
        assert "ALL SOURCES" in df["source_workbook"].values

    def test_metrics_present(self, ctx_two_bundles):
        df = _build_kpi_summary(ctx_two_bundles)
        assert set(df["metric"]).issuperset({"sum", "count", "mean"})

    def test_empty_master_returns_empty(self):
        ctx = MasterWorkbookContext(bundles=[])
        df = _build_kpi_summary(ctx)
        assert len(df) == 0


class TestBuildIssuesLog:
    def test_returns_dataframe(self, ctx_two_bundles):
        df = _build_issues_log(ctx_two_bundles)
        assert isinstance(df, pd.DataFrame)

    def test_required_columns(self, ctx_two_bundles):
        df = _build_issues_log(ctx_two_bundles)
        assert "issue_id" in df.columns
        assert "issue_description" in df.columns
        assert "status" in df.columns


class TestBuildSop:
    def test_returns_dataframe(self):
        df = _build_sop()
        assert isinstance(df, pd.DataFrame)

    def test_has_eight_or_more_steps(self):
        df = _build_sop()
        assert len(df) >= 8

    def test_step_column_present(self):
        df = _build_sop()
        assert "step" in df.columns
