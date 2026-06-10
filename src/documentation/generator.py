"""Auto-generate data dictionaries and lineage metadata."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


@dataclass
class DataDictionaryEntry:
    column_name: str
    data_type: str
    description: str
    source_reports: list[str] = field(default_factory=list)
    sample_values: list = field(default_factory=list)
    notes: str = ""


def build_data_dictionary(
    profiles: dict,
    matches: list,
) -> list[DataDictionaryEntry]:
    """Produce a data dictionary from profiling and match results.

    Args:
        profiles: Output from :mod:`src.profiling.profiler`.
        matches: Output from :mod:`src.schema_matching.matcher`.

    Returns:
        List of :class:`DataDictionaryEntry` objects.
    """
    raise NotImplementedError


def export_to_excel(entries: list[DataDictionaryEntry], output_path: Path) -> None:
    """Write the data dictionary to an Excel file at *output_path*."""
    raise NotImplementedError


def build_lineage_map(bundles: list, matches: list) -> dict:
    """Return a column-level lineage map: output column → list of source (report, column) pairs."""
    raise NotImplementedError
