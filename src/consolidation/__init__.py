"""Consolidation module — merge aligned reports into a master dataset."""

from src.consolidation.consolidator import (
    LINEAGE_COL_SHEET,
    LINEAGE_COL_WORKBOOK,
    build_source_mapping_df,
    consolidate,
    resolve_conflicts,
)

__all__ = [
    "consolidate",
    "resolve_conflicts",
    "build_source_mapping_df",
    "LINEAGE_COL_WORKBOOK",
    "LINEAGE_COL_SHEET",
]
