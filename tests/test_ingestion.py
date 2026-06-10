"""Tests for src.ingestion.loader — placeholder suite."""

import pytest

from src.ingestion.loader import WorkbookBundle, load_workbook


def test_load_workbook_missing_file(tmp_path):
    with pytest.raises((FileNotFoundError, NotImplementedError)):
        load_workbook(tmp_path / "nonexistent.xlsx")


def test_workbook_bundle_sheet_names():
    import pandas as pd

    bundle = WorkbookBundle(source_path="dummy.xlsx", sheets={"Sheet1": pd.DataFrame()})
    assert bundle.sheet_names == ["Sheet1"]
