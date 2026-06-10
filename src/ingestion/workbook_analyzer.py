"""Workbook Intelligence Engine.

Classifies worksheet tabs using structural signals — never sheet names.
Extracts formula-based KPI dependency chains.

Classification is driven by a weighted multi-signal scorer.  Every measurable
characteristic of a worksheet contributes to a score for each candidate tab
type; the type with the highest score wins.  All thresholds and weights are
held in :class:`ClassificationConfig` so callers can tune them without
touching source code.

Tab types produced
------------------
SOURCE_DATA        Raw, detailed records; low formula density, many rows.
REFERENCE_DATA     Small lookup / code tables; high unique-value ratio, few rows.
MAPPING_TABLE      Two-to-four column translation tables; very high unique ratio.
KPI_SUMMARY        High formula density + aggregation functions + few rows.
DASHBOARD          High formula density + inter-sheet refs + mixed data types.
OUTPUT             Dense formulas referencing many other sheets; no raw data.
CALCULATION        Formula-heavy helper sheet; moderate inter-sheet refs.
MIXED              Meaningful mix of raw data and formulas.
UNKNOWN            Cannot be determined from available signals.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Any

import pandas as pd

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Regex: inter-sheet cell reference  SheetName!ColLetters[RowNum]
# Deliberately ignores sheet name content — purely structural detection.
# ---------------------------------------------------------------------------
_SHEET_REF_RE = re.compile(
    r"'?([A-Za-z0-9_][\w\s]*?)'?\s*!\s*\$?([A-Z]+)\$?(\d*)"
)

# Excel aggregate function names used to detect KPI / summary intent
_AGGREGATE_FUNCS = frozenset({
    "SUM", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS",
    "AVERAGE", "AVERAGEIF", "AVERAGEIFS",
    "MAX", "MIN", "MEDIAN",
    "SUMIF", "SUMIFS",
    "VLOOKUP", "HLOOKUP", "INDEX", "MATCH",
    "PIVOT", "GETPIVOTDATA",
})


# ---------------------------------------------------------------------------
# Public enumerations and dataclasses
# ---------------------------------------------------------------------------

class TabType(str, Enum):
    SOURCE_DATA    = "source_data"
    REFERENCE_DATA = "reference_data"
    MAPPING_TABLE  = "mapping_table"
    KPI_SUMMARY    = "kpi_summary"
    DASHBOARD      = "dashboard"
    OUTPUT         = "output"
    CALCULATION    = "calculation"
    MIXED          = "mixed"
    UNKNOWN        = "unknown"

# Backward-compatible aliases (module-level, not enum members — avoids enum alias pitfalls)
TabType.SUMMARY = TabType.KPI_SUMMARY  # type: ignore[attr-defined]
TabType.KPI     = TabType.KPI_SUMMARY  # type: ignore[attr-defined]


@dataclass
class ClassificationConfig:
    """Tunable thresholds for the multi-signal classifier.

    All values have tested defaults; override only what you need.
    """
    # --- row/column count signals ---
    min_source_rows: int = 10          # fewer rows → unlikely to be raw data
    max_reference_rows: int = 200      # reference/mapping tables are small
    max_mapping_cols: int = 6          # mapping tables have few columns

    # --- formula density signals (formula_cells / non_empty_cells) ---
    kpi_formula_density_min: float = 0.25   # ≥ this → KPI/summary candidate
    source_formula_density_max: float = 0.05  # ≤ this → raw data candidate
    calculation_formula_density_min: float = 0.15  # helper sheets are formula-heavy

    # --- unique-value ratio signals (unique_values / total_values per col, avg) ---
    reference_unique_ratio_min: float = 0.80  # lookup cols are mostly unique
    mapping_unique_ratio_min: float = 0.70

    # --- inter-sheet reference signals ---
    dashboard_ref_sheet_min: int = 2   # dashboards pull from ≥ N sheets
    output_ref_sheet_min: int = 3      # output tabs pull from ≥ N sheets

    # --- aggregation function signal ---
    kpi_agg_func_min: int = 1          # ≥ 1 agg function detected → KPI candidate

    # --- data-type homogeneity signal ---
    # fraction of columns sharing the majority dtype → high = structured source data
    homogeneity_source_min: float = 0.70

    # --- scoring weights (relative; need not sum to 1) ---
    w_formula_density: float = 3.0
    w_inter_sheet_refs: float = 2.5
    w_agg_functions: float = 2.0
    w_row_count: float = 1.5
    w_unique_ratio: float = 1.5
    w_col_homogeneity: float = 1.0
    w_col_count: float = 0.5


@dataclass
class SignalVector:
    """Raw measurable signals extracted from a single worksheet."""

    row_count: int
    col_count: int
    formula_cell_count: int
    non_empty_cell_count: int
    formula_density: float          # formula_cells / non_empty_cells
    inter_sheet_ref_count: int      # formula cells referencing other sheets
    referenced_sheet_count: int     # distinct sheets referenced
    referenced_sheets: list[str]
    agg_func_count: int             # cells using an aggregate function
    avg_unique_ratio: float         # mean(unique_values/total_values) across cols
    dtype_homogeneity: float        # fraction of cols sharing the majority dtype


@dataclass
class TabAnalysis:
    """Classification result for a single worksheet."""

    workbook_name: str
    tab_name: str
    tab_type: TabType
    row_count: int
    col_count: int
    formula_cell_count: int
    non_empty_cell_count: int
    formula_density: float
    inter_sheet_ref_count: int
    referenced_sheets: list[str] = field(default_factory=list)
    classification_reason: str = ""
    signal_scores: dict[str, float] = field(default_factory=dict)


@dataclass
class KPIDefinition:
    """A single KPI extracted from a formula-bearing tab."""

    workbook_name: str
    source_tab: str
    cell_address: str
    kpi_label: str
    formula: str
    aggregate_function: str
    referenced_sheets: list[str] = field(default_factory=list)
    referenced_columns: list[str] = field(default_factory=list)


@dataclass
class WorkbookAnalysis:
    """Complete structural analysis of one workbook."""

    workbook_name: str
    tab_analyses: list[TabAnalysis] = field(default_factory=list)
    kpi_definitions: list[KPIDefinition] = field(default_factory=list)

    @property
    def source_tabs(self) -> list[str]:
        return [t.tab_name for t in self.tab_analyses
                if t.tab_type == TabType.SOURCE_DATA]

    @property
    def reference_tabs(self) -> list[str]:
        return [t.tab_name for t in self.tab_analyses
                if t.tab_type in (TabType.REFERENCE_DATA, TabType.MAPPING_TABLE)]

    @property
    def summary_tabs(self) -> list[str]:
        return [t.tab_name for t in self.tab_analyses
                if t.tab_type in (TabType.KPI_SUMMARY, TabType.DASHBOARD, TabType.OUTPUT)]

    @property
    def calculation_tabs(self) -> list[str]:
        return [t.tab_name for t in self.tab_analyses
                if t.tab_type == TabType.CALCULATION]

    def tab_type(self, name: str) -> TabType:
        for t in self.tab_analyses:
            if t.tab_name == name:
                return t.tab_type
        return TabType.UNKNOWN

    def data_bearing_tabs(
        self,
        types: set[TabType] | None = None,
    ) -> list[str]:
        """Tabs suitable for consolidation.  Defaults to SOURCE_DATA + REFERENCE_DATA."""
        if types is None:
            types = {TabType.SOURCE_DATA, TabType.REFERENCE_DATA}
        return [t.tab_name for t in self.tab_analyses if t.tab_type in types]


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def analyze_workbook(
    bundle,
    config: ClassificationConfig | None = None,
) -> WorkbookAnalysis:
    """Classify every tab and extract KPI definitions from formula tabs.

    Classification is entirely signal-driven; no sheet names are inspected.

    Args:
        bundle: A :class:`~src.ingestion.loader.WorkbookBundle`.
        config: Optional :class:`ClassificationConfig`.  Defaults are used if
            omitted.

    Returns:
        :class:`WorkbookAnalysis` with classified tabs and extracted KPIs.
    """
    cfg = config or ClassificationConfig()
    raw_wb = getattr(bundle, "_raw_wb", None)
    wb_name = bundle.file_name

    headers_by_sheet: dict[str, list[str]] = {
        sn: [str(c) for c in df.columns]
        for sn, df in bundle.sheets.items()
    }

    tab_analyses: list[TabAnalysis] = []
    kpi_definitions: list[KPIDefinition] = []

    for sheet_name, df in bundle.sheets.items():
        raw_ws = _get_raw_ws(raw_wb, sheet_name)
        signals = _extract_signals(raw_ws, df)
        tab = _classify_tab(signals, sheet_name, wb_name, cfg)
        tab_analyses.append(tab)
        logger.debug(
            "Tab '%s/%s' → %s  (formula_density=%.2f, inter_sheet=%d, "
            "agg_funcs=%d, avg_unique=%.2f, homogeneity=%.2f)",
            wb_name, sheet_name, tab.tab_type.value,
            signals.formula_density, signals.inter_sheet_ref_count,
            signals.agg_func_count, signals.avg_unique_ratio,
            signals.dtype_homogeneity,
        )

        if tab.tab_type in (
            TabType.KPI_SUMMARY, TabType.DASHBOARD, TabType.OUTPUT,
            TabType.CALCULATION, TabType.MIXED,
        ) and raw_ws is not None:
            kpis = _extract_kpis(raw_ws, sheet_name, wb_name, headers_by_sheet)
            kpi_definitions.extend(kpis)
            if kpis:
                logger.info(
                    "Extracted %d KPI definition(s) from '%s/%s'",
                    len(kpis), wb_name, sheet_name,
                )

    analysis = WorkbookAnalysis(
        workbook_name=wb_name,
        tab_analyses=tab_analyses,
        kpi_definitions=kpi_definitions,
    )
    logger.info(
        "Analysis '%s': %d source, %d reference, %d summary, %d KPI def(s)",
        wb_name, len(analysis.source_tabs), len(analysis.reference_tabs),
        len(analysis.summary_tabs), len(kpi_definitions),
    )
    return analysis


def analyze_many(
    bundles: list,
    config: ClassificationConfig | None = None,
) -> list[WorkbookAnalysis]:
    """Analyse multiple bundles, skipping failures with a warning."""
    results = []
    for bundle in bundles:
        try:
            results.append(analyze_workbook(bundle, config=config))
        except Exception as exc:
            logger.warning(
                "Could not analyse '%s': %s",
                getattr(bundle, "file_name", "?"), exc,
            )
    return results


def build_kpi_inventory_df(analyses: list[WorkbookAnalysis]) -> pd.DataFrame:
    """Flatten all KPI definitions into a tidy DataFrame.

    Adds three rationalization columns:

    ``is_common``
        True when the same KPI label appears in ≥ 2 workbooks (name similarity
        only — may still represent different metrics).

    ``rationalization_candidate``
        True when a KPI pair shares the same label **and** the same aggregate
        function **and** at least one overlapping referenced source column.
        Only these pairs are truly candidates for consolidation.

    ``rationalization_note``
        Human-readable explanation of why the KPI is or is not a candidate.
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
                "_ref_cols_set": set(kpi.referenced_columns),
            })

    if not rows:
        return pd.DataFrame(columns=[
            "workbook_name", "source_tab", "kpi_label", "formula",
            "aggregate_function", "referenced_sheets", "referenced_columns",
            "cell_address", "is_common", "rationalization_candidate",
            "rationalization_note",
        ])

    df = pd.DataFrame(rows)
    label_counts = df.groupby("kpi_label")["workbook_name"].nunique()
    df["is_common"] = df["kpi_label"].map(lambda lbl: label_counts.get(lbl, 0) >= 2)

    candidates: list[bool] = []
    notes: list[str] = []

    for _, row in df.iterrows():
        if not row["is_common"]:
            candidates.append(False)
            notes.append("Unique to one workbook — no rationalization needed.")
            continue

        # Find peer KPIs with same label from other workbooks
        peers = df[
            (df["kpi_label"] == row["kpi_label"]) &
            (df["workbook_name"] != row["workbook_name"])
        ]
        same_func = peers[peers["aggregate_function"] == row["aggregate_function"]]
        if same_func.empty:
            candidates.append(False)
            notes.append(
                f"Same label in multiple workbooks but different aggregate functions "
                f"({row['aggregate_function']} vs "
                f"{', '.join(peers['aggregate_function'].unique())}). "
                f"Review before consolidating."
            )
            continue

        # Check for overlapping referenced columns
        overlapping = same_func[
            same_func["_ref_cols_set"].apply(
                lambda peer_cols: bool(peer_cols & row["_ref_cols_set"])
                if isinstance(peer_cols, set) and isinstance(row["_ref_cols_set"], set)
                else False
            )
        ]
        if not overlapping.empty:
            candidates.append(True)
            shared = row["_ref_cols_set"] & overlapping.iloc[0]["_ref_cols_set"]
            notes.append(
                f"Rationalization candidate: same label, same function "
                f"({row['aggregate_function']}), shared source columns: "
                f"{', '.join(sorted(shared))}."
            )
        else:
            candidates.append(False)
            notes.append(
                f"Same label and function ({row['aggregate_function']}) but no "
                f"overlapping referenced source columns. These KPIs likely measure "
                f"different data — keep separate."
            )

    df["rationalization_candidate"] = candidates
    df["rationalization_note"] = notes
    df.drop(columns=["_ref_cols_set"], inplace=True)
    return df


def build_dependency_report_df(analyses: list[WorkbookAnalysis]) -> pd.DataFrame:
    """Build a row-per-tab structural dependency report."""
    rows = []
    for analysis in analyses:
        for tab in analysis.tab_analyses:
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
# Signal extraction
# ---------------------------------------------------------------------------

def _extract_signals(raw_ws: Any, df: pd.DataFrame) -> SignalVector:
    """Compute all measurable signals from a worksheet."""
    row_count = len(df)
    col_count = len(df.columns)

    # --- unique-value ratio (from pandas data) ---
    avg_unique_ratio = _compute_avg_unique_ratio(df)

    # --- dtype homogeneity ---
    dtype_homogeneity = _compute_dtype_homogeneity(df)

    if raw_ws is None:
        return SignalVector(
            row_count=row_count, col_count=col_count,
            formula_cell_count=0,
            non_empty_cell_count=max(row_count * col_count, 1),
            formula_density=0.0,
            inter_sheet_ref_count=0, referenced_sheet_count=0,
            referenced_sheets=[],
            agg_func_count=0,
            avg_unique_ratio=avg_unique_ratio,
            dtype_homogeneity=dtype_homogeneity,
        )

    formula_cells = 0
    inter_sheet_refs = 0
    non_empty = 0
    agg_funcs = 0
    referenced_sheets: set[str] = set()

    for row in raw_ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            non_empty += 1
            if isinstance(cell.value, str) and cell.value.startswith("="):
                formula = cell.value
                formula_cells += 1
                if "!" in formula:
                    inter_sheet_refs += 1
                    for m in _SHEET_REF_RE.finditer(formula):
                        referenced_sheets.add(m.group(1).strip())
                if _has_aggregate_function(formula):
                    agg_funcs += 1

    formula_density = formula_cells / non_empty if non_empty > 0 else 0.0

    return SignalVector(
        row_count=row_count,
        col_count=col_count,
        formula_cell_count=formula_cells,
        non_empty_cell_count=non_empty,
        formula_density=round(formula_density, 4),
        inter_sheet_ref_count=inter_sheet_refs,
        referenced_sheet_count=len(referenced_sheets),
        referenced_sheets=sorted(referenced_sheets),
        agg_func_count=agg_funcs,
        avg_unique_ratio=round(avg_unique_ratio, 4),
        dtype_homogeneity=round(dtype_homogeneity, 4),
    )


def _compute_avg_unique_ratio(df: pd.DataFrame) -> float:
    """Mean of (unique_values / total_values) across all non-empty columns."""
    if df.empty or len(df.columns) == 0:
        return 0.0
    ratios = []
    for col in df.columns:
        series = df[col].dropna()
        if len(series) > 0:
            ratios.append(series.nunique() / len(series))
    return sum(ratios) / len(ratios) if ratios else 0.0


def _compute_dtype_homogeneity(df: pd.DataFrame) -> float:
    """Fraction of columns sharing the single most common inferred dtype."""
    if df.empty or len(df.columns) == 0:
        return 0.0
    dtype_counts: dict[str, int] = {}
    for col in df.columns:
        dtype_counts[str(df[col].dtype)] = dtype_counts.get(str(df[col].dtype), 0) + 1
    majority = max(dtype_counts.values())
    return majority / len(df.columns)


def _has_aggregate_function(formula: str) -> bool:
    upper = formula.upper()
    return any(f"={fn}(" in upper or f",{fn}(" in upper or f"({fn}(" in upper
               for fn in _AGGREGATE_FUNCS)


# ---------------------------------------------------------------------------
# Multi-signal classifier
# ---------------------------------------------------------------------------

def _classify_tab(
    sv: SignalVector,
    sheet_name: str,
    wb_name: str,
    cfg: ClassificationConfig,
) -> TabAnalysis:
    """Score each candidate TabType against the signal vector; return the winner."""

    scores: dict[TabType, float] = {t: 0.0 for t in TabType}

    # -----------------------------------------------------------------------
    # SOURCE_DATA signals
    # -----------------------------------------------------------------------
    if sv.row_count >= cfg.min_source_rows:
        scores[TabType.SOURCE_DATA] += cfg.w_row_count * min(sv.row_count / cfg.min_source_rows, 3.0)
    if sv.formula_density <= cfg.source_formula_density_max:
        scores[TabType.SOURCE_DATA] += cfg.w_formula_density * (1.0 - sv.formula_density / max(cfg.source_formula_density_max, 0.001))
    if sv.dtype_homogeneity >= cfg.homogeneity_source_min:
        scores[TabType.SOURCE_DATA] += cfg.w_col_homogeneity * sv.dtype_homogeneity
    if sv.inter_sheet_ref_count == 0:
        scores[TabType.SOURCE_DATA] += cfg.w_inter_sheet_refs * 0.5

    # -----------------------------------------------------------------------
    # REFERENCE_DATA signals
    # -----------------------------------------------------------------------
    if 0 < sv.row_count <= cfg.max_reference_rows:
        scores[TabType.REFERENCE_DATA] += cfg.w_row_count * (1.0 - sv.row_count / cfg.max_reference_rows)
    if sv.avg_unique_ratio >= cfg.reference_unique_ratio_min:
        scores[TabType.REFERENCE_DATA] += cfg.w_unique_ratio * sv.avg_unique_ratio
    if sv.formula_density <= cfg.source_formula_density_max:
        scores[TabType.REFERENCE_DATA] += cfg.w_formula_density * 0.5

    # -----------------------------------------------------------------------
    # MAPPING_TABLE signals (very small, high-uniqueness, few columns)
    # -----------------------------------------------------------------------
    if 0 < sv.row_count <= cfg.max_reference_rows and sv.col_count <= cfg.max_mapping_cols:
        scores[TabType.MAPPING_TABLE] += cfg.w_col_count * (1.0 - sv.col_count / cfg.max_mapping_cols)
    if sv.avg_unique_ratio >= cfg.mapping_unique_ratio_min:
        scores[TabType.MAPPING_TABLE] += cfg.w_unique_ratio * sv.avg_unique_ratio * 1.2
    if sv.formula_density <= cfg.source_formula_density_max:
        scores[TabType.MAPPING_TABLE] += cfg.w_formula_density * 0.5

    # -----------------------------------------------------------------------
    # KPI_SUMMARY signals
    # -----------------------------------------------------------------------
    if sv.formula_density >= cfg.kpi_formula_density_min:
        scores[TabType.KPI_SUMMARY] += cfg.w_formula_density * (sv.formula_density / cfg.kpi_formula_density_min)
    if sv.agg_func_count >= cfg.kpi_agg_func_min:
        scores[TabType.KPI_SUMMARY] += cfg.w_agg_functions * min(sv.agg_func_count, 10)
    if sv.inter_sheet_ref_count > 0:
        scores[TabType.KPI_SUMMARY] += cfg.w_inter_sheet_refs * 1.0
    if sv.row_count < cfg.min_source_rows:
        scores[TabType.KPI_SUMMARY] += cfg.w_row_count * 0.5

    # -----------------------------------------------------------------------
    # DASHBOARD signals (multi-source, visually mixed, inter-sheet heavy)
    # -----------------------------------------------------------------------
    if sv.referenced_sheet_count >= cfg.dashboard_ref_sheet_min:
        scores[TabType.DASHBOARD] += cfg.w_inter_sheet_refs * sv.referenced_sheet_count
    if sv.formula_density >= cfg.kpi_formula_density_min:
        scores[TabType.DASHBOARD] += cfg.w_formula_density * 0.8
    if sv.agg_func_count >= cfg.kpi_agg_func_min:
        scores[TabType.DASHBOARD] += cfg.w_agg_functions * 0.5

    # -----------------------------------------------------------------------
    # OUTPUT signals (very dense formulas, referencing many sheets)
    # -----------------------------------------------------------------------
    if sv.referenced_sheet_count >= cfg.output_ref_sheet_min:
        scores[TabType.OUTPUT] += cfg.w_inter_sheet_refs * sv.referenced_sheet_count * 1.2
    if sv.formula_density >= cfg.kpi_formula_density_min:
        scores[TabType.OUTPUT] += cfg.w_formula_density * sv.formula_density * 2.0

    # -----------------------------------------------------------------------
    # CALCULATION signals (helper/intermediate formula sheet)
    # -----------------------------------------------------------------------
    if sv.formula_density >= cfg.calculation_formula_density_min:
        scores[TabType.CALCULATION] += cfg.w_formula_density * sv.formula_density
    if 0 < sv.referenced_sheet_count < cfg.dashboard_ref_sheet_min:
        scores[TabType.CALCULATION] += cfg.w_inter_sheet_refs * 0.8
    if sv.agg_func_count == 0 and sv.formula_cell_count > 0:
        scores[TabType.CALCULATION] += cfg.w_agg_functions * 0.5

    # -----------------------------------------------------------------------
    # MIXED signals (meaningful raw data AND formulas)
    # -----------------------------------------------------------------------
    if (sv.row_count >= cfg.min_source_rows
            and cfg.source_formula_density_max < sv.formula_density < cfg.kpi_formula_density_min):
        scores[TabType.MIXED] += cfg.w_formula_density + cfg.w_row_count

    # -----------------------------------------------------------------------
    # Remove UNKNOWN from competition (fallback only) and pick winner
    # -----------------------------------------------------------------------
    del scores[TabType.UNKNOWN]

    best_type = max(scores, key=lambda t: scores[t])
    best_score = scores[best_type]

    # If the best score is near zero the sheet is truly ambiguous
    if best_score < 0.1:
        best_type = TabType.UNKNOWN

    reason = (
        f"top_type={best_type.value} score={best_score:.2f} | "
        f"formula_density={sv.formula_density:.3f} "
        f"inter_sheet_refs={sv.inter_sheet_ref_count} "
        f"agg_funcs={sv.agg_func_count} "
        f"rows={sv.row_count} cols={sv.col_count} "
        f"avg_unique={sv.avg_unique_ratio:.2f} "
        f"homogeneity={sv.dtype_homogeneity:.2f}"
    )

    return TabAnalysis(
        workbook_name=wb_name,
        tab_name=sheet_name,
        tab_type=best_type,
        row_count=sv.row_count,
        col_count=sv.col_count,
        formula_cell_count=sv.formula_cell_count,
        non_empty_cell_count=sv.non_empty_cell_count,
        formula_density=sv.formula_density,
        inter_sheet_ref_count=sv.inter_sheet_ref_count,
        referenced_sheets=sv.referenced_sheets,
        classification_reason=reason,
        signal_scores={t.value: round(s, 3) for t, s in scores.items()},
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

    Priority: left cell → above cell → cell-coordinate fallback.
    Sheet names are never consulted.
    """
    if col_idx > 1:
        left = raw_ws.cell(row=row_idx, column=col_idx - 1)
        if left.value and not (isinstance(left.value, str) and left.value.startswith("=")):
            return str(left.value).strip()
    if row_idx > 1:
        above = raw_ws.cell(row=row_idx - 1, column=col_idx)
        if above.value and not (isinstance(above.value, str) and above.value.startswith("=")):
            return str(above.value).strip()
    return f"KPI_R{row_idx}C{col_idx}"


def _detect_aggregate_function(formula: str) -> str:
    """Return the outermost aggregate function name, or 'OTHER'."""
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
