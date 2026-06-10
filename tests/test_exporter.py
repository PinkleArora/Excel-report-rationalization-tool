"""Tests for src.profiling.exporter — Excel export."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from src.ingestion.loader import load_workbook
from src.profiling.exporter import export_profiling_results
from src.profiling.profiler import profile_workbook


@pytest.fixture
def workbook_meta(tmp_path):
    path = tmp_path / "sample.xlsx"
    df = pd.DataFrame({"product": ["A", "B", "C"], "revenue": [100, 200, 300]})
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sales", index=False)
    bundle = load_workbook(path)
    return profile_workbook(bundle)


class TestExportProfilingResults:
    def test_creates_file(self, workbook_meta, tmp_path):
        out = export_profiling_results([workbook_meta], output_path=tmp_path / "out.xlsx")
        assert out.exists()

    def test_default_output_dir(self, workbook_meta, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        out = export_profiling_results([workbook_meta])
        assert out.name == "profiling_summary.xlsx"
        assert out.exists()

    def test_output_has_three_sheets(self, workbook_meta, tmp_path):
        out = export_profiling_results([workbook_meta], output_path=tmp_path / "out.xlsx")
        xl = pd.ExcelFile(out)
        assert set(xl.sheet_names) == {"Workbook_Summary", "Sheet_Summary", "Column_Summary"}

    def test_workbook_summary_not_empty(self, workbook_meta, tmp_path):
        out = export_profiling_results([workbook_meta], output_path=tmp_path / "out.xlsx")
        df = pd.read_excel(out, sheet_name="Workbook_Summary")
        assert len(df) == 1

    def test_column_summary_not_empty(self, workbook_meta, tmp_path):
        out = export_profiling_results([workbook_meta], output_path=tmp_path / "out.xlsx")
        df = pd.read_excel(out, sheet_name="Column_Summary")
        assert len(df) >= 1

    def test_raises_on_empty_input(self, tmp_path):
        with pytest.raises(ValueError, match="No workbook metadata"):
            export_profiling_results([], output_path=tmp_path / "out.xlsx")

    def test_creates_parent_dirs(self, workbook_meta, tmp_path):
        nested = tmp_path / "a" / "b" / "c" / "out.xlsx"
        out = export_profiling_results([workbook_meta], output_path=nested)
        assert out.exists()

    def test_returned_path_is_absolute(self, workbook_meta, tmp_path):
        out = export_profiling_results([workbook_meta], output_path=tmp_path / "out.xlsx")
        assert out.is_absolute()
