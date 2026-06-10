"""Auto-generate data dictionaries and lineage metadata from profiling results."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd

from src.profiling.metadata import WorkbookMetadata
from src.schema_matching.matcher import ColumnMatch
from src.schema_matching.normalizer import normalize_column_name

logger = logging.getLogger(__name__)


@dataclass
class DataDictionaryEntry:
    """One entry in the data dictionary — corresponds to one canonical column."""

    canonical_column: str
    display_name: str
    data_types: list[str] = field(default_factory=list)       # unique dtypes seen
    semantic_types: list[str] = field(default_factory=list)   # unique semantic types
    source_workbooks: list[str] = field(default_factory=list)
    source_columns: list[str] = field(default_factory=list)   # original names before canonicalisation
    avg_null_pct: float = 0.0
    sample_values: list[Any] = field(default_factory=list)
    notes: str = ""

    # Back-compat alias kept for tests that used the old field name
    @property
    def column_name(self) -> str:
        return self.canonical_column

    @property
    def description(self) -> str:
        return self.notes

    @property
    def source_reports(self) -> list[str]:
        return self.source_workbooks


def build_data_dictionary(
    workbook_metas: list[WorkbookMetadata],
    matches: list[ColumnMatch],
) -> list[DataDictionaryEntry]:
    """Produce a data dictionary from profiling and match results.

    Each :class:`DataDictionaryEntry` corresponds to one canonical column in the
    master dataset.  Columns that match across workbooks are merged into a single
    entry; unique columns get their own entry.

    Args:
        workbook_metas: Profiling output from :func:`~src.profiling.profiler.profile_workbook`.
        matches: Column matches from :func:`~src.schema_matching.matcher.match_all_bundles`.

    Returns:
        List of :class:`DataDictionaryEntry`, one per canonical column, sorted alphabetically.
    """
    # Build canonical map: (wb_name, sheet_name, original_col) → canonical_col
    canonical_map: dict[tuple[str, str, str], str] = {}
    for wm in workbook_metas:
        for sm in wm.sheets:
            for cm in sm.columns:
                key = (wm.workbook_name, sm.sheet_name, cm.column_name)
                canonical_map[key] = normalize_column_name(cm.column_name)

    # Apply match unifications
    for m in matches:
        src_key = (m.source_workbook, m.source_sheet, m.source_col)
        tgt_key = (m.target_workbook, m.target_sheet, m.target_col)
        src_c = canonical_map.get(src_key, normalize_column_name(m.source_col))
        tgt_c = canonical_map.get(tgt_key, normalize_column_name(m.target_col))
        canonical = min(src_c, tgt_c, key=lambda s: (len(s), s))
        for key, val in canonical_map.items():
            if val in (src_c, tgt_c):
                canonical_map[key] = canonical

    # Aggregate per canonical column
    entries: dict[str, DataDictionaryEntry] = {}

    for wm in workbook_metas:
        for sm in wm.sheets:
            for cm in sm.columns:
                key = (wm.workbook_name, sm.sheet_name, cm.column_name)
                canonical = canonical_map.get(key, normalize_column_name(cm.column_name))

                if canonical not in entries:
                    entries[canonical] = DataDictionaryEntry(
                        canonical_column=canonical,
                        display_name=cm.column_name,
                    )
                entry = entries[canonical]

                if cm.inferred_dtype not in entry.data_types:
                    entry.data_types.append(cm.inferred_dtype)
                if cm.semantic_type not in entry.semantic_types:
                    entry.semantic_types.append(cm.semantic_type)
                if wm.workbook_name not in entry.source_workbooks:
                    entry.source_workbooks.append(wm.workbook_name)
                if cm.column_name not in entry.source_columns:
                    entry.source_columns.append(cm.column_name)

                # Merge sample values (up to 5 unique)
                for v in cm.sample_values:
                    if v not in entry.sample_values and len(entry.sample_values) < 5:
                        entry.sample_values.append(v)

    # Recompute avg_null_pct
    null_pcts: dict[str, list[float]] = {c: [] for c in entries}
    for wm in workbook_metas:
        for sm in wm.sheets:
            for cm in sm.columns:
                key = (wm.workbook_name, sm.sheet_name, cm.column_name)
                canonical = canonical_map.get(key, normalize_column_name(cm.column_name))
                null_pcts[canonical].append(cm.null_pct)

    for canonical, pcts in null_pcts.items():
        if pcts and canonical in entries:
            entries[canonical].avg_null_pct = round(sum(pcts) / len(pcts), 2)

    result = sorted(entries.values(), key=lambda e: e.canonical_column)
    logger.info("Built data dictionary: %d canonical column(s)", len(result))
    return result


def data_dictionary_to_df(entries: list[DataDictionaryEntry]) -> pd.DataFrame:
    """Convert data dictionary entries to a tidy DataFrame for export."""
    rows = []
    for e in entries:
        rows.append({
            "canonical_column": e.canonical_column,
            "display_name": e.display_name,
            "data_types": ", ".join(e.data_types),
            "semantic_types": ", ".join(e.semantic_types),
            "source_workbooks": ", ".join(e.source_workbooks),
            "source_columns": ", ".join(e.source_columns),
            "avg_null_pct": e.avg_null_pct,
            "sample_values": " | ".join(str(v) for v in e.sample_values),
            "notes": e.notes,
        })
    return pd.DataFrame(rows)


def build_lineage_map(
    workbook_metas: list[WorkbookMetadata],
    matches: list[ColumnMatch],
) -> dict[str, list[tuple[str, str, str]]]:
    """Return column-level lineage: canonical_col → [(workbook, sheet, original_col), …]."""
    entries = build_data_dictionary(workbook_metas, matches)
    lineage: dict[str, list[tuple[str, str, str]]] = {}

    for e in entries:
        # Re-derive source tuples from workbook_metas
        sources: list[tuple[str, str, str]] = []
        for wm in workbook_metas:
            for sm in wm.sheets:
                for cm in sm.columns:
                    if cm.column_name in e.source_columns and wm.workbook_name in e.source_workbooks:
                        sources.append((wm.workbook_name, sm.sheet_name, cm.column_name))
        lineage[e.canonical_column] = sources

    return lineage


def export_to_excel(entries: list[DataDictionaryEntry], output_path: Path) -> None:
    """Write the data dictionary to a single-sheet Excel file at *output_path*."""
    df = data_dictionary_to_df(entries)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False, engine="openpyxl")
    logger.info("Data dictionary exported to %s", output_path)
