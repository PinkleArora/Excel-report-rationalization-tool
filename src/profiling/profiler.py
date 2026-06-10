"""Compute column-level and sheet-level quality profiles."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pandas as pd


@dataclass
class ColumnProfile:
    name: str
    dtype: str
    non_null_count: int
    null_count: int
    null_pct: float
    unique_count: int
    sample_values: list[Any]


@dataclass
class SheetProfile:
    sheet_name: str
    row_count: int
    col_count: int
    columns: list[ColumnProfile]


def profile_sheet(df: pd.DataFrame, sheet_name: str = "") -> SheetProfile:
    """Return a :class:`SheetProfile` for *df*.

    Args:
        df: The sheet data to profile.
        sheet_name: Optional human-readable label.

    Returns:
        A fully populated :class:`SheetProfile`.
    """
    raise NotImplementedError


def profile_workbook(sheets: dict[str, pd.DataFrame]) -> dict[str, SheetProfile]:
    """Profile every sheet in *sheets* and return a mapping of sheet name → profile."""
    raise NotImplementedError
