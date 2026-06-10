"""Validate consolidated output totals against original source reports."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import pandas as pd


@dataclass
class ReconciliationResult:
    column: str
    source_total: float
    output_total: float
    delta: float
    delta_pct: float
    status: Literal["pass", "fail", "warn"]
    tolerance: float


def reconcile(
    source_frames: list[pd.DataFrame],
    output_frame: pd.DataFrame,
    numeric_columns: list[str],
    tolerance: float = 0.01,
) -> list[ReconciliationResult]:
    """Check that numeric column sums in *output_frame* match the sum across *source_frames*.

    Args:
        source_frames: Original per-report DataFrames.
        output_frame: The consolidated output to validate.
        numeric_columns: Columns to reconcile.
        tolerance: Acceptable relative difference (default 1 %).

    Returns:
        One :class:`ReconciliationResult` per column.
    """
    raise NotImplementedError


def reconciliation_summary(results: list[ReconciliationResult]) -> pd.DataFrame:
    """Convert reconciliation results to a tidy summary DataFrame."""
    raise NotImplementedError
