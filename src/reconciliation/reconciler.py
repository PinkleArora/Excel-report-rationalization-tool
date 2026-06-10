"""Validate consolidated output totals against original source reports."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Literal

import pandas as pd

from src.consolidation.consolidator import LINEAGE_COL_SHEET, LINEAGE_COL_WORKBOOK
from src.schema_matching.normalizer import normalize_column_name

logger = logging.getLogger(__name__)

# Warn (rather than fail) when abs delta_pct is between tolerance and _WARN_MULTIPLIER*tolerance
_WARN_MULTIPLIER = 3.0


@dataclass
class ReconciliationResult:
    """Totals-check result for one canonical column."""

    column: str
    source_total: float
    output_total: float
    delta: float
    delta_pct: float
    status: Literal["pass", "warn", "fail"]
    tolerance: float


def reconcile(
    source_frames: list[pd.DataFrame],
    output_frame: pd.DataFrame,
    numeric_columns: list[str],
    tolerance: float = 0.01,
) -> list[ReconciliationResult]:
    """Check that numeric column sums in *output_frame* match those in *source_frames*.

    Column names in *source_frames* are normalised before comparison so that
    original names (e.g. ``"Revenue (£)"``) align with their canonical master
    names (e.g. ``"revenue"``)

    Args:
        source_frames: Original per-report DataFrames (no lineage columns).
        output_frame: The consolidated master DataFrame (may contain lineage columns).
        numeric_columns: Canonical column names to reconcile.
        tolerance: Acceptable relative difference as a fraction (default 0.01 = 1 %).

    Returns:
        One :class:`ReconciliationResult` per column in *numeric_columns*.
    """
    if not numeric_columns:
        return []

    # Strip lineage columns from output before summing
    output_data = output_frame.drop(
        columns=[c for c in (LINEAGE_COL_WORKBOOK, LINEAGE_COL_SHEET) if c in output_frame.columns],
        errors="ignore",
    )

    results: list[ReconciliationResult] = []

    for col in numeric_columns:
        # Sum across all source frames (normalise source column names on the fly)
        source_total = 0.0
        for sf in source_frames:
            norm_map = {c: normalize_column_name(str(c)) for c in sf.columns}
            renamed = sf.rename(columns=norm_map)
            if col in renamed.columns:
                numeric_series = pd.to_numeric(renamed[col], errors="coerce")
                source_total += float(numeric_series.sum(skipna=True))

        # Sum from consolidated output
        output_total = 0.0
        if col in output_data.columns:
            numeric_series = pd.to_numeric(output_data[col], errors="coerce")
            output_total = float(numeric_series.sum(skipna=True))

        delta = output_total - source_total
        if source_total != 0:
            delta_pct = abs(delta / source_total)
        else:
            delta_pct = 0.0 if output_total == 0 else 1.0

        if delta_pct <= tolerance:
            status: Literal["pass", "warn", "fail"] = "pass"
        elif delta_pct <= tolerance * _WARN_MULTIPLIER:
            status = "warn"
        else:
            status = "fail"

        results.append(ReconciliationResult(
            column=col,
            source_total=round(source_total, 6),
            output_total=round(output_total, 6),
            delta=round(delta, 6),
            delta_pct=round(delta_pct * 100, 4),  # store as percentage
            status=status,
            tolerance=round(tolerance * 100, 4),
        ))
        logger.debug(
            "Reconcile '%s': source=%.2f, output=%.2f, delta_pct=%.4f%%, status=%s",
            col, source_total, output_total, delta_pct * 100, status,
        )

    return results


def reconciliation_summary(results: list[ReconciliationResult]) -> pd.DataFrame:
    """Convert reconciliation results to a tidy summary DataFrame."""
    rows = []
    for r in results:
        rows.append({
            "column": r.column,
            "source_total": r.source_total,
            "master_total": r.output_total,
            "delta": r.delta,
            "delta_pct": r.delta_pct,
            "tolerance_pct": r.tolerance,
            "status": r.status.upper(),
        })
    df = pd.DataFrame(rows)
    # Sort: fail first, then warn, then pass
    order = {"FAIL": 0, "WARN": 1, "PASS": 2}
    if not df.empty:
        df["_sort"] = df["status"].map(order)
        df = df.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)
    return df
