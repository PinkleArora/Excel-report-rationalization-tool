"""Analyse workbook structure: classify tabs and extract formula-based KPI dependencies."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Any

import pandas as pd

logger = logging.getLogger(__name__)

# Minimum rows a sheet must have to be considered a source-data candidate
_MIN_SOURCE_ROWS = 3
# Formula density (formula_cells / non_empty_cells) above which a tab is classed as SUMMARY
_SUMMARY_FORMULA_DENSITY = 0.20
# Regex to match inter-sheet cell references: SheetName!ColLetters[RowNum]
_SHEET_REF_RE = re.compile(
    r"'?([A-Za-z0-9_][\w\s]*?)'?\s*!\s*\$?([A-Z]+)\$?(\d*)"
)
# Common KPI aggregate function names
_AGGREGATE_FUNCS = {"SUM", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "AVERAGE",
                    "AVERAGEIF", "MAX", "MIN", "SUMIF", "SUMIFS", "VLOOKUP", "INDEX"}


class TabType(str, Enum):
    SOURCE_DATA = "source_data"
    SUMMARY = "summary"
    KPI = "kpi"
    MIXED = "mixed"
    UNKNOWN = "unknown"


@dataclass
class TabAnalysis:
    """Analysis result for a single worksheet."""

    workbook_name: str
    tab_name: str
    tab_type: TabType
    row_count: int
    col_count: int
    formula_cell_count: int
    non_empty_cell_count: int
    formula_density: float      # formula_cells / non_empty_cells, 0-1
    inter_sheet_ref_count: int  # cells that reference other sheets
    referenced_sheets: list[str] = field(default_factory=list)
    classification_reason: str = ""


@dataclass
class KPIDefinition:
    """A single KPI extracted from a summary/formula tab."""

    workbook_name: str
    source_tab: str
    cell_address: str    # e.g. "B2"
    kpi_label: str       # label from adjacent cell
    formula: str         # raw formula string, e.g. "=SUM(SQL_Data!F2:F8)"
    aggregate_function: str   # SUM, COUNT, etc.
    referenced_sheets: list[str] = field(default_factory=list)
    referenced_columns: list[str] = field(default_factory=list)  # resolved header names


@dataclass
class WorkbookAnalysis:
    """Complete structural analysis of one workbook."""

    workbook_name: str
    tab_analyses: list[TabAnalysis] = field(default_factory=list)
    kpi_definitions: list[KPIDefinition] = field(default_factory=list)

    @property
    def source_tabs(self) -> list[str]:
        return [t.tab_name for t in self.tab_analyses if t.tab_type == TabType.SOURCE_DATA]

    @property
    def summary_tabs(self) -> list[str]:
        return [t.tab_name for t in self.tab_analyses
                if t.tab_type in (TabType.SUMMARY, TabType.KPI)]

    def tab_type(self, name: str) -> TabType:
        for t in self.tab_analyses:
            if t.tab_name == name:
                return t.tab_type
        return TabType.UNKNOWN


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def analyze_workbook(bundle) -> WorkbookAnalysis:
    """Classify every tab and extract KPI definitions from formula tabs.

    Args:
        bundle: A :class:`~src.ingestion.loader.WorkbookBundle`.

    Returns:
        :class:`WorkbookAnalysis` with classified tabs and extracted KPIs.
    """
    raw_wb = getattr(bundle, "_raw_wb", None)
    wb_name = bundle.file_name

    # Build a mapping sheet_name → list of header strings for formula resolution
    headers_by_sheet: dict[str, list[str]] = {
        sn: [str(c) for c in df.columns]
        for sn, df in bundle.sheets.items()
    }

    tab_analyses: list[TabAnalysis] = []
    kpi_definitions: list[KPIDefinition] = []

    for sheet_name, df in bundle.sheets.items():
        raw_ws = _get_raw_ws(raw_wb, sheet_name)
        tab = _analyze_tab(raw_ws, df, sheet_name, wb_name)
        tab_analyses.append(tab)
        logger.debug("Tab '%s/%s' classified as %s (formula_density=%.2f, inter_sheet=%d)",
                     wb_name, sheet_name, tab.tab_type.value,
                     tab.formula_density, tab.inter_sheet_ref_count)

        if tab.tab_type in (TabType.SUMMARY, TabType.KPI, TabType.MIXED) and raw_ws is not None:
            kpis = _extract_kpis(raw_ws, sheet_name, wb_name, headers_by_sheet)
            kpi_definitions.extend(kpis)
            if kpis:
                logger.info("Extracted %d KPI definition(s) from '%s/%s'",
                            len(kpis), wb_name, sheet_name)

    analysis = WorkbookAnalysis(
        workbook_name=wb_name,
        tab_analyses=tab_analyses,
        kpi_definitions=kpi_definitions,
    )
    logger.info(
        "Analysis '%s': %d source tab(s), %d summary tab(s), %d KPI definition(s)",
        wb_name, len(analysis.source_tabs), len(analysis.summary_tabs), len(kpi_definitions),
    )
    return analysis


def analyze_many(bundles: list) -> list[WorkbookAnalysis]:
    """Analyse multiple bundles, skipping failures with a warning."""
    results = []
    for bundle in bundles:
        try:
            results.append(analyze_workbook(bundle))
        except Exception as exc:
            logger.warning("Could not analyse '%s': %s", getattr(bundle, 'file_name', '?'), exc)
    return results


def build_kpi_inventory_df(analyses: list[WorkbookAnalysis]) -> pd.DataFrame:
    """Flatten all KPI definitions into a tidy DataFrame.

    Adds an ``is_common`` flag: True when the same label appears in ≥2 workbooks.
    """
    rows = []
    for analysis in analyses:
        for kpi in analysis.kpi_definitions:
            rows.append({
                "workbook_name": kpi.workbook_name,
                "source_tab": kpi.source_tab,
                "kpi_label": kpi.kpi_label,
                "formula": kpi.formula,
                "aggregate_function": kpi.aggregate_function,
                "referenced_sheets": ", ".join(kpi.referenced_sheets),
                "referenced_columns": ", ".join(kpi.referenced_columns),
                "cell_address": kpi.cell_address,
            })

    if not rows:
        return pd.DataFrame(columns=[
            "workbook_name", "source_tab", "kpi_label", "formula",
            "aggregate_function", "referenced_sheets", "referenced_columns",
            "cell_address", "is_common",
        ])

    df = pd.DataFrame(rows)
    # Mark KPIs whose label appears in more than one workbook
    label_counts = df.groupby("kpi_label")["workbook_name"].nunique()
    df["is_common"] = df["kpi_label"].map(lambda lbl: label_counts.get(lbl, 0) >= 2)
    return df


def build_dependency_report_df(analyses: list[WorkbookAnalysis]) -> pd.DataFrame:
    """Build a row-per-KPI dependency report suitable for the workbook output."""
    rows = []
    for analysis in analyses:
        for tab in analysis.tab_analyses:
            # Source tabs: no formula info needed beyond classification
            rows.append({
                "workbook": analysis.workbook_name,
                "tab_name": tab.tab_name,
                "tab_type": tab.tab_type.value,
                "row_count": tab.row_count,
                "col_count": tab.col_count,
                "formula_cells": tab.formula_cell_count,
                "formula_density_pct": round(tab.formula_density * 100, 1),
                "inter_sheet_refs": tab.inter_sheet_ref_count,
                "references": ", ".join(tab.referenced_sheets),
                "classification_reason": tab.classification_reason,
            })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Tab classification
# ---------------------------------------------------------------------------

def _analyze_tab(raw_ws: Any, df: pd.DataFrame, sheet_name: str, wb_name: str) -> TabAnalysis:
    row_count = len(df)
    col_count = len(df.columns)

    if raw_ws is None:
        # No formula info available — use shape heuristics
        tab_type = TabType.SOURCE_DATA if row_count >= _MIN_SOURCE_ROWS else TabType.UNKNOWN
        return TabAnalysis(
            workbook_name=wb_name, tab_name=sheet_name, tab_type=tab_type,
            row_count=row_count, col_count=col_count,
            formula_cell_count=0, non_empty_cell_count=row_count * col_count,
            formula_density=0.0, inter_sheet_ref_count=0,
            classification_reason="no formula metadata — classified by shape",
        )

    formula_cells = 0
    inter_sheet_refs = 0
    non_empty = 0
    referenced_sheets: set[str] = set()

    for row in raw_ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                non_empty += 1
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formula_cells += 1
                    if "!" in cell.value:
                        inter_sheet_refs += 1
                        for m in _SHEET_REF_RE.finditer(cell.value):
                            referenced_sheets.add(m.group(1).strip())

    formula_density = formula_cells / non_empty if non_empty > 0 else 0.0

    # Classification logic
    if inter_sheet_refs > 0 and formula_density >= _SUMMARY_FORMULA_DENSITY:
        tab_type = TabType.SUMMARY
        reason = (f"formula_density={formula_density:.2f}, "
                  f"inter_sheet_refs={inter_sheet_refs} → classified as SUMMARY")
    elif formula_cells > 0 and formula_density >= _SUMMARY_FORMULA_DENSITY and row_count < 20:
        tab_type = TabType.SUMMARY
        reason = (f"formula_density={formula_density:.2f}, "
                  f"few rows → classified as SUMMARY")
    elif row_count >= _MIN_SOURCE_ROWS and formula_density < 0.05:
        tab_type = TabType.SOURCE_DATA
        reason = (f"row_count={row_count}, "
                  f"formula_density={formula_density:.2f} → classified as SOURCE_DATA")
    elif formula_cells > 0 and row_count >= _MIN_SOURCE_ROWS:
        tab_type = TabType.MIXED
        reason = (f"formula_density={formula_density:.2f}, "
                  f"row_count={row_count} → classified as MIXED")
    else:
        tab_type = TabType.UNKNOWN
        reason = f"row_count={row_count}, formula_density={formula_density:.2f} → UNKNOWN"

    return TabAnalysis(
        workbook_name=wb_name, tab_name=sheet_name, tab_type=tab_type,
        row_count=row_count, col_count=col_count,
        formula_cell_count=formula_cells, non_empty_cell_count=non_empty,
        formula_density=round(formula_density, 4),
        inter_sheet_ref_count=inter_sheet_refs,
        referenced_sheets=sorted(referenced_sheets),
        classification_reason=reason,
    )


# ---------------------------------------------------------------------------
# KPI extraction
# ---------------------------------------------------------------------------

def _extract_kpis(
    raw_ws: Any,
    sheet_name: str,
    wb_name: str,
    headers_by_sheet: dict[str, list[str]],
) -> list[KPIDefinition]:
    kpis: list[KPIDefinition] = []

    for row in raw_ws.iter_rows():
        for cell in row:
            if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                continue
            formula = cell.value
            col_letter = cell.column_letter
            row_idx = cell.row
            col_idx = cell.column

            label = _extract_kpi_label(raw_ws, row_idx, col_idx)
            agg_func = _detect_aggregate_function(formula)
            sheet_refs, col_refs = _parse_formula_refs(formula, headers_by_sheet)

            kpis.append(KPIDefinition(
                workbook_name=wb_name,
                source_tab=sheet_name,
                cell_address=f"{col_letter}{row_idx}",
                kpi_label=label,
                formula=formula,
                aggregate_function=agg_func,
                referenced_sheets=sheet_refs,
                referenced_columns=col_refs,
            ))

    return kpis


def _extract_kpi_label(raw_ws: Any, row_idx: int, col_idx: int) -> str:
    """Find the human-readable label adjacent to a formula cell.

    Priority: left cell → above cell → cell coordinate fallback.
    """
    # Check cell to the left
    if col_idx > 1:
        left = raw_ws.cell(row=row_idx, column=col_idx - 1)
        if left.value and not (isinstance(left.value, str) and left.value.startswith("=")):
            return str(left.value).strip()
    # Check cell above
    if row_idx > 1:
        above = raw_ws.cell(row=row_idx - 1, column=col_idx)
        if above.value and not (isinstance(above.value, str) and above.value.startswith("=")):
            return str(above.value).strip()
    return f"KPI_R{row_idx}C{col_idx}"


def _detect_aggregate_function(formula: str) -> str:
    """Extract the outermost aggregate function name from a formula string."""
    m = re.match(r"=\s*([A-Z]+)\s*\(", formula.upper())
    if m and m.group(1) in _AGGREGATE_FUNCS:
        return m.group(1)
    return "OTHER"


def _parse_formula_refs(
    formula: str,
    headers_by_sheet: dict[str, list[str]],
) -> tuple[list[str], list[str]]:
    """Extract (referenced_sheet_names, resolved_column_names) from a formula."""
    seen_sheets: list[str] = []
    seen_cols: list[str] = []

    for m in _SHEET_REF_RE.finditer(formula):
        sheet_name = m.group(1).strip()
        col_letter = m.group(2).upper()

        if sheet_name not in seen_sheets:
            seen_sheets.append(sheet_name)

        if sheet_name in headers_by_sheet:
            col_idx = _col_letter_to_index(col_letter)
            headers = headers_by_sheet[sheet_name]
            if 0 <= col_idx < len(headers):
                col_name = headers[col_idx]
                if col_name not in seen_cols:
                    seen_cols.append(col_name)

    return seen_sheets, seen_cols


def _col_letter_to_index(letters: str) -> int:
    """Convert a column letter (A, B, …, AA, AB, …) to a 0-based index."""
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------

def _get_raw_ws(raw_wb: Any, sheet_name: str) -> Any:
    if raw_wb is None:
        return None
    try:
        return raw_wb[sheet_name]
    except (KeyError, TypeError):
        return None
