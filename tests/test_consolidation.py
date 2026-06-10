"""Tests for src.consolidation.consolidator."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pandas as pd
import pytest

from src.consolidation.consolidator import (
    LINEAGE_COL_SHEET,
    LINEAGE_COL_WORKBOOK,
    _build_canonical_map,
    build_source_mapping_df,
    consolidate,
    resolve_conflicts,
)
from src.ingestion.loader import WorkbookBundle
from src.schema_matching.matcher import ColumnMatch


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def bundle_sales(tmp_path: Path) -> WorkbookBundle:
    path = tmp_path / "sales.xlsx"
    df = pd.DataFrame({"Region": ["North", "South"], "Revenue": [100, 200], "Cost": [50, 80]})
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Sales", index=False)
    from src.ingestion.loader import load_workbook
    return load_workbook(path)


@pytest.fixture
def bundle_hr(tmp_path: Path) -> WorkbookBundle:
    path = tmp_path / "hr.xlsx"
    df = pd.DataFrame({"Department": ["Eng", "HR"], "Headcount": [10, 5], "Cost": [500, 200]})
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="HR", index=False)
    from src.ingestion.loader import load_workbook
    return load_workbook(path)


@pytest.fixture
def exact_matches(bundle_sales, bundle_hr) -> list[ColumnMatch]:
    return [
        ColumnMatch(
            source_workbook="sales.xlsx", source_sheet="Sales", source_col="Cost",
            target_workbook="hr.xlsx",    target_sheet="HR",    target_col="Cost",
            score=100.0, match_type="exact",
        )
    ]


# ---------------------------------------------------------------------------
# consolidate
# ---------------------------------------------------------------------------

class TestConsolidate:
    def test_returns_dataframe(self, bundle_sales, bundle_hr):
        df = consolidate([bundle_sales, bundle_hr], [])
        assert isinstance(df, pd.DataFrame)

    def test_has_lineage_columns(self, bundle_sales, bundle_hr):
        df = consolidate([bundle_sales, bundle_hr], [])
        assert LINEAGE_COL_WORKBOOK in df.columns
        assert LINEAGE_COL_SHEET in df.columns

    def test_lineage_workbook_values(self, bundle_sales, bundle_hr):
        df = consolidate([bundle_sales, bundle_hr], [])
        assert set(df[LINEAGE_COL_WORKBOOK]) == {"sales.xlsx", "hr.xlsx"}

    def test_row_count_equals_sum_of_source_rows(self, bundle_sales, bundle_hr):
        df = consolidate([bundle_sales, bundle_hr], [])
        expected = sum(len(b.sheets[s]) for b in [bundle_sales, bundle_hr] for s in b.sheets)
        assert len(df) == expected

    def test_union_strategy_keeps_all_columns(self, bundle_sales, bundle_hr):
        df = consolidate([bundle_sales, bundle_hr], [])
        # Should have columns from both sources (normalised)
        assert "region" in df.columns or "Region" in df.columns or "region" in df.columns
        assert "headcount" in df.columns or "Headcount" in df.columns

    def test_intersection_strategy_keeps_only_matched(self, bundle_sales, bundle_hr, exact_matches):
        df = consolidate([bundle_sales, bundle_hr], exact_matches, strategy="intersection")
        data_cols = [c for c in df.columns if c not in (LINEAGE_COL_WORKBOOK, LINEAGE_COL_SHEET)]
        assert len(data_cols) <= 1  # only "cost" (the matched column)

    def test_matched_column_renamed_to_canonical(self, bundle_sales, bundle_hr, exact_matches):
        df = consolidate([bundle_sales, bundle_hr], exact_matches)
        # "Cost" from both sources should be unified as "cost"
        assert "cost" in df.columns

    def test_empty_bundles_returns_empty(self):
        df = consolidate([], [])
        assert df.empty

    def test_single_bundle(self, bundle_sales):
        df = consolidate([bundle_sales], [])
        assert len(df) == 2
        assert LINEAGE_COL_WORKBOOK in df.columns

    def test_lineage_sheet_values(self, bundle_sales):
        df = consolidate([bundle_sales], [])
        assert set(df[LINEAGE_COL_SHEET]) == {"Sales"}


# ---------------------------------------------------------------------------
# resolve_conflicts
# ---------------------------------------------------------------------------

class TestResolveConflicts:
    @pytest.fixture
    def master_with_dupes(self, bundle_sales, bundle_hr):
        # Duplicate by concatenating sales with itself
        df = consolidate([bundle_sales, bundle_sales], [])
        return df

    def test_last_wins_removes_duplicates(self, master_with_dupes):
        result = resolve_conflicts(master_with_dupes, strategy="last_wins")
        data_cols = [c for c in result.columns if c not in (LINEAGE_COL_WORKBOOK, LINEAGE_COL_SHEET)]
        assert not result.duplicated(subset=data_cols).any()

    def test_first_wins_removes_duplicates(self, master_with_dupes):
        result = resolve_conflicts(master_with_dupes, strategy="first_wins")
        data_cols = [c for c in result.columns if c not in (LINEAGE_COL_WORKBOOK, LINEAGE_COL_SHEET)]
        assert not result.duplicated(subset=data_cols).any()

    def test_raise_strategy_on_duplicates(self, master_with_dupes):
        with pytest.raises(ValueError, match="duplicate"):
            resolve_conflicts(master_with_dupes, strategy="raise")

    def test_no_duplicates_returns_unchanged(self, bundle_sales, bundle_hr):
        df = consolidate([bundle_sales, bundle_hr], [])
        result = resolve_conflicts(df, strategy="raise")
        assert len(result) == len(df)


# ---------------------------------------------------------------------------
# build_source_mapping_df
# ---------------------------------------------------------------------------

class TestBuildSourceMappingDf:
    def test_returns_dataframe(self, bundle_sales, exact_matches):
        df = build_source_mapping_df([bundle_sales], exact_matches)
        assert isinstance(df, pd.DataFrame)

    def test_required_columns_present(self, bundle_sales, exact_matches):
        df = build_source_mapping_df([bundle_sales], exact_matches)
        required = {"source_workbook", "source_sheet", "source_column", "canonical_column", "is_matched"}
        assert required.issubset(set(df.columns))

    def test_row_count_equals_total_columns(self, bundle_sales, bundle_hr):
        df = build_source_mapping_df([bundle_sales, bundle_hr], [])
        expected = sum(len(b.sheets[s].columns) for b in [bundle_sales, bundle_hr] for s in b.sheets)
        assert len(df) == expected

    def test_matched_column_flagged(self, bundle_sales, bundle_hr, exact_matches):
        df = build_source_mapping_df([bundle_sales, bundle_hr], exact_matches)
        matched = df[df["is_matched"]]
        assert len(matched) >= 1

    def test_unmatched_columns_have_none_score(self, bundle_sales):
        df = build_source_mapping_df([bundle_sales], [])
        assert df["match_score"].isna().all()

    def test_workbook_name_in_output(self, bundle_sales):
        df = build_source_mapping_df([bundle_sales], [])
        assert "sales.xlsx" in df["source_workbook"].values


# ---------------------------------------------------------------------------
# _build_canonical_map
# ---------------------------------------------------------------------------

class TestBuildCanonicalMap:
    def test_all_keys_present(self, bundle_sales):
        canon = _build_canonical_map([bundle_sales], [])
        assert ("sales.xlsx", "Sales", "Revenue") in canon

    def test_matched_cols_share_canonical(self, bundle_sales, bundle_hr, exact_matches):
        canon = _build_canonical_map([bundle_sales, bundle_hr], exact_matches)
        src = canon.get(("sales.xlsx", "Sales", "Cost"))
        tgt = canon.get(("hr.xlsx", "HR", "Cost"))
        assert src == tgt == "cost"


class TestCollisionSafePropagation:
    """Tests for intra-frame collision detection in _build_canonical_map."""

    def test_no_collision_match_applied(self):
        """When no collision, match propagates normally."""
        from src.consolidation.consolidator import _build_canonical_map
        from src.schema_matching.matcher import ColumnMatch
        import pandas as pd

        bundle = MagicMock()
        bundle.file_name = "wb.xlsx"
        bundle.sheets = {"Sheet1": pd.DataFrame({"revenue": [1], "cost": [2]})}

        bundle2 = MagicMock()
        bundle2.file_name = "wb2.xlsx"
        bundle2.sheets = {"Sheet1": pd.DataFrame({"Revenue": [10]})}

        match = ColumnMatch(
            source_workbook="wb.xlsx", source_sheet="Sheet1", source_col="revenue",
            target_workbook="wb2.xlsx", target_sheet="Sheet1", target_col="Revenue",
            score=100.0, match_type="exact",
        )
        result = _build_canonical_map([bundle, bundle2], [match])
        # Both should map to "revenue" (shorter/earlier)
        assert result[("wb.xlsx", "Sheet1", "revenue")] == "revenue"
        assert result[("wb2.xlsx", "Sheet1", "Revenue")] == "revenue"

    def test_intra_frame_collision_skipped(self):
        """Match that would create duplicate canonical in same sheet is skipped."""
        from src.consolidation.consolidator import _build_canonical_map
        from src.schema_matching.matcher import ColumnMatch
        import pandas as pd

        # wb.xlsx has BOTH "revenue_total" and "rev_total" in the same sheet
        bundle = MagicMock()
        bundle.file_name = "wb.xlsx"
        bundle.sheets = {"S": pd.DataFrame({"revenue_total": [1], "rev_total": [2]})}

        bundle2 = MagicMock()
        bundle2.file_name = "wb2.xlsx"
        bundle2.sheets = {"S": pd.DataFrame({"rev_total": [10]})}

        # This match would try to unify "revenue_total" and "rev_total" → both → "rev_total"
        # but wb.xlsx/S already has both → collision
        match = ColumnMatch(
            source_workbook="wb.xlsx", source_sheet="S", source_col="revenue_total",
            target_workbook="wb2.xlsx", target_sheet="S", target_col="rev_total",
            score=82.0, match_type="fuzzy",
        )
        collision_log = []
        result = _build_canonical_map([bundle, bundle2], [match], collision_log=collision_log)
        # wb.xlsx/S columns should NOT be unified (would cause collision)
        assert result[("wb.xlsx", "S", "revenue_total")] != result[("wb.xlsx", "S", "rev_total")]
        assert len(collision_log) == 1

    def test_collision_log_populated(self):
        """collision_log gets an entry for each skipped match."""
        from src.consolidation.consolidator import _build_canonical_map
        from src.schema_matching.matcher import ColumnMatch
        import pandas as pd

        bundle = MagicMock()
        bundle.file_name = "wb.xlsx"
        bundle.sheets = {"S": pd.DataFrame({"revenue_total": [1], "rev_total": [2]})}
        bundle2 = MagicMock()
        bundle2.file_name = "wb2.xlsx"
        bundle2.sheets = {"S": pd.DataFrame({"rev_total": [10]})}
        match = ColumnMatch("wb.xlsx", "S", "revenue_total", "wb2.xlsx", "S", "rev_total", 82.0, "fuzzy")

        log = []
        _build_canonical_map([bundle, bundle2], [match], collision_log=log)
        assert log[0]["source_col"] == "revenue_total"
        assert log[0]["target_col"] == "rev_total"

    def test_disambiguate_columns_renames_dupes(self):
        """_disambiguate_columns adds _dup2, _dup3 suffixes for repeated col names."""
        from src.consolidation.consolidator import _disambiguate_columns
        import pandas as pd
        df = pd.DataFrame([[1, 2, 3]], columns=["a", "a", "a"])
        out = _disambiguate_columns(df)
        assert list(out.columns) == ["a", "a_dup2", "a_dup3"]

    def test_consolidate_no_crash_on_collision_scenario(self):
        """consolidate() should not raise InvalidIndexError even with fuzzy collision risk."""
        from src.consolidation.consolidator import consolidate
        from src.schema_matching.matcher import ColumnMatch
        import pandas as pd

        bundle = MagicMock()
        bundle.file_name = "wb.xlsx"
        bundle.sheets = {"S": pd.DataFrame({"revenue_total": [1, 2], "rev_total": [3, 4]})}
        bundle2 = MagicMock()
        bundle2.file_name = "wb2.xlsx"
        bundle2.sheets = {"S": pd.DataFrame({"rev_total": [10, 20]})}
        match = ColumnMatch("wb.xlsx", "S", "revenue_total", "wb2.xlsx", "S", "rev_total", 82.0, "fuzzy")

        # Should not raise
        result = consolidate([bundle, bundle2], [match])
        assert not result.empty
