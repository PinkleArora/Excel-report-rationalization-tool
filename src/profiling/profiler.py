"""Compute column-level and sheet-level quality profiles from WorkbookBundles."""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

import pandas as pd

from src.ingestion.loader import WorkbookBundle
from src.profiling.metadata import ColumnMetadata, SheetMetadata, WorkbookMetadata

logger = logging.getLogger(__name__)

_SAMPLE_SIZE = 5  # number of non-null sample values to retain


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def profile_workbook(bundle: WorkbookBundle) -> WorkbookMetadata:
    """Build a full :class:`WorkbookMetadata` from a loaded workbook bundle.

    Args:
        bundle: A :class:`~src.ingestion.loader.WorkbookBundle` returned by
            :func:`~src.ingestion.loader.load_workbook`.

    Returns:
        Fully populated :class:`WorkbookMetadata`.
    """
    wb_name = bundle.file_name
    logger.info("Profiling workbook: %s", wb_name)

    sheet_metas: list[SheetMetadata] = []
    for sheet_name, df in bundle.sheets.items():
        raw_ws = _get_raw_ws(bundle, sheet_name)
        sm = profile_sheet(df, sheet_name=sheet_name, workbook_name=wb_name, raw_ws=raw_ws)
        sheet_metas.append(sm)
        logger.debug("  Profiled sheet '%s': %d rows, %d cols", sheet_name, sm.row_count, sm.col_count)

    wm = WorkbookMetadata(
        workbook_name=wb_name,
        source_path=bundle.source_path,
        file_extension=bundle.source_path.suffix.lower(),
        sheet_count=len(sheet_metas),
        total_rows=sum(s.row_count for s in sheet_metas),
        total_cols=sum(s.col_count for s in sheet_metas),
        total_formula_cells=sum(s.formula_cell_count for s in sheet_metas),
        total_empty_cols=sum(s.empty_col_count for s in sheet_metas),
        sheets=sheet_metas,
    )
    logger.info(
        "Profiled '%s': %d sheet(s), %d total rows, %d formula cells",
        wb_name, wm.sheet_count, wm.total_rows, wm.total_formula_cells,
    )
    return wm


def profile_sheet(
    df: pd.DataFrame,
    sheet_name: str = "",
    workbook_name: str = "",
    raw_ws: Any = None,
) -> SheetMetadata:
    """Profile a single sheet DataFrame.

    Args:
        df: Sheet data (header already parsed as column names).
        sheet_name: Human-readable sheet label.
        workbook_name: Parent workbook label.
        raw_ws: openpyxl Worksheet object for formula scanning (may be None).

    Returns:
        Populated :class:`SheetMetadata`.
    """
    col_metas: list[ColumnMetadata] = []
    for idx, col in enumerate(df.columns):
        cm = _profile_column(df[col], col_name=str(col), col_idx=idx,
                             sheet_name=sheet_name, workbook_name=workbook_name)
        col_metas.append(cm)

    formula_count = _count_formula_cells(df, raw_ws)
    empty_count = sum(1 for cm in col_metas if cm.is_empty)

    return SheetMetadata(
        workbook_name=workbook_name,
        sheet_name=sheet_name,
        row_count=len(df),
        col_count=len(df.columns),
        empty_col_count=empty_count,
        formula_cell_count=formula_count,
        column_names=[str(c) for c in df.columns],
        columns=col_metas,
    )


def profile_many(bundles: list[WorkbookBundle]) -> list[WorkbookMetadata]:
    """Profile every bundle in *bundles*, skipping failures with a warning."""
    results: list[WorkbookMetadata] = []
    for bundle in bundles:
        try:
            results.append(profile_workbook(bundle))
        except Exception as exc:
            logger.warning("Could not profile '%s': %s", bundle.file_name, exc)
    return results


# ---------------------------------------------------------------------------
# Summary DataFrames
# ---------------------------------------------------------------------------

def build_summary_dataframes(
    workbook_metas: list[WorkbookMetadata],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Convert metadata objects into three tidy summary DataFrames.

    Returns:
        ``(workbook_summary, sheet_summary, column_summary)``
    """
    wb_rows: list[dict] = []
    sheet_rows: list[dict] = []
    col_rows: list[dict] = []

    for wm in workbook_metas:
        wb_rows.append({
            "workbook_name": wm.workbook_name,
            "source_path": str(wm.source_path),
            "file_extension": wm.file_extension,
            "sheet_count": wm.sheet_count,
            "total_rows": wm.total_rows,
            "total_cols": wm.total_cols,
            "total_formula_cells": wm.total_formula_cells,
            "total_empty_cols": wm.total_empty_cols,
        })

        for sm in wm.sheets:
            sheet_rows.append({
                "workbook_name": sm.workbook_name,
                "sheet_name": sm.sheet_name,
                "row_count": sm.row_count,
                "col_count": sm.col_count,
                "empty_col_count": sm.empty_col_count,
                "formula_cell_count": sm.formula_cell_count,
                "column_names": ", ".join(sm.column_names),
            })

            for cm in sm.columns:
                col_rows.append({
                    "workbook_name": cm.workbook_name,
                    "sheet_name": cm.sheet_name,
                    "column_index": cm.column_index,
                    "column_name": cm.column_name,
                    "inferred_dtype": cm.inferred_dtype,
                    "semantic_type": cm.semantic_type,
                    "non_null_count": cm.non_null_count,
                    "null_count": cm.null_count,
                    "null_pct": round(cm.null_pct, 2),
                    "unique_count": cm.unique_count,
                    "is_empty": cm.is_empty,
                    "sample_values": " | ".join(str(v) for v in cm.sample_values),
                    "min_value": cm.min_value,
                    "max_value": cm.max_value,
                    "mean_value": cm.mean_value,
                })

    return (
        pd.DataFrame(wb_rows),
        pd.DataFrame(sheet_rows),
        pd.DataFrame(col_rows),
    )


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _profile_column(
    series: pd.Series,
    col_name: str,
    col_idx: int,
    sheet_name: str,
    workbook_name: str,
) -> ColumnMetadata:
    total = len(series)
    null_count = int(series.isna().sum())
    non_null_count = total - null_count
    null_pct = (null_count / total * 100) if total > 0 else 100.0
    is_empty = non_null_count == 0

    unique_count = int(series.nunique(dropna=True))
    sample_values = _extract_samples(series)

    dtype_str = str(series.dtype)
    semantic = _infer_semantic_type(series)

    min_val: Any = None
    max_val: Any = None
    mean_val: float | None = None

    if not is_empty and pd.api.types.is_numeric_dtype(series):
        try:
            min_val = series.min(skipna=True)
            max_val = series.max(skipna=True)
            mean_val = round(float(series.mean(skipna=True)), 6)
        except Exception:
            pass
    elif not is_empty and pd.api.types.is_datetime64_any_dtype(series):
        try:
            min_val = str(series.min())
            max_val = str(series.max())
        except Exception:
            pass

    return ColumnMetadata(
        workbook_name=workbook_name,
        sheet_name=sheet_name,
        column_index=col_idx,
        column_name=col_name,
        inferred_dtype=dtype_str,
        semantic_type=semantic,
        non_null_count=non_null_count,
        null_count=null_count,
        null_pct=null_pct,
        unique_count=unique_count,
        is_empty=is_empty,
        sample_values=sample_values,
        min_value=min_val,
        max_value=max_val,
        mean_value=mean_val,
    )


def _infer_semantic_type(series: pd.Series) -> str:
    if series.isna().all():
        return "empty"
    if pd.api.types.is_bool_dtype(series):
        return "boolean"
    if pd.api.types.is_integer_dtype(series):
        return "integer"
    if pd.api.types.is_float_dtype(series):
        return "float"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "datetime"
    # Attempt datetime parse on object columns
    if series.dtype == object:
        sample = series.dropna().head(20)
        parsed = pd.to_datetime(sample, errors="coerce", infer_datetime_format=True)
        if parsed.notna().sum() / max(len(sample), 1) >= 0.8:
            return "datetime"
    return "text"


def _extract_samples(series: pd.Series) -> list[Any]:
    non_null = series.dropna()
    samples = non_null.head(_SAMPLE_SIZE).tolist()
    return [_safe_scalar(v) for v in samples]


def _safe_scalar(v: Any) -> Any:
    """Convert numpy scalars to native Python types for serialisation."""
    try:
        import numpy as np
        if isinstance(v, (np.integer,)):
            return int(v)
        if isinstance(v, (np.floating,)):
            return float(v)
        if isinstance(v, (np.bool_,)):
            return bool(v)
    except ImportError:
        pass
    return v


def _count_formula_cells(df: pd.DataFrame, raw_ws: Any) -> int:
    """Count cells whose value starts with '=' using the openpyxl worksheet if available."""
    if raw_ws is not None:
        # openpyxl worksheet (data_only=False preserves formula strings)
        count = 0
        for row in raw_ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    count += 1
        return count

    # Fallback: scan string values in the DataFrame itself
    count = 0
    for col in df.columns:
        if df[col].dtype == object:
            count += int(df[col].astype(str).str.startswith("=").sum())
    return count


def _get_raw_ws(bundle: WorkbookBundle, sheet_name: str) -> Any:
    """Retrieve the raw openpyxl Worksheet from the bundle, if available."""
    raw_wb = getattr(bundle, "_raw_wb", None)
    if raw_wb is None:
        return None
    try:
        return raw_wb[sheet_name]
    except (KeyError, TypeError):
        return None
