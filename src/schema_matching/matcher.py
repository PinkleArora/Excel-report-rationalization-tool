"""Detect semantically equivalent columns across workbooks using fuzzy matching."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass
class ColumnMatch:
    source_sheet: str
    source_col: str
    target_sheet: str
    target_col: str
    score: float  # 0–100 (RapidFuzz WRatio)
    match_type: str  # "exact" | "fuzzy" | "ai"


def match_columns(
    source: dict[str, pd.DataFrame],
    target: dict[str, pd.DataFrame],
    threshold: float = 80.0,
) -> list[ColumnMatch]:
    """Find column-level matches between *source* and *target* workbooks.

    Args:
        source: Sheet-name → DataFrame mapping for the source workbook.
        target: Sheet-name → DataFrame mapping for the target workbook.
        threshold: Minimum RapidFuzz score (0–100) to accept a match.

    Returns:
        Sorted list of :class:`ColumnMatch` objects, highest score first.
    """
    raise NotImplementedError


def build_match_matrix(
    bundles: list[dict[str, pd.DataFrame]],
    threshold: float = 80.0,
) -> pd.DataFrame:
    """Build a pairwise match matrix across all provided workbook sheet dicts."""
    raise NotImplementedError
