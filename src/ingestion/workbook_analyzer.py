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
from concurrent.futures import ThreadPoolExecutor, as_completed
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
# Pre-compiled regex — much faster than 18 string constructions per formula cell
_AGG_FUNC_RE = re.compile(
    r"(?:^|[=,(])(" + "|".join(_AGGREGATE_FUNCS) + r")\s*\(",
    re.IGNORECASE,
)


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
class KpiBlock:
    """A self-contained KPI section within a summary/KPI tab.

    Many actuarial reserve worksheets contain multiple named sections — each
    with a title row, a header row defining dimensions and KPI measure columns,
    and several data rows.  For example, a "Summary" tab might contain separate
    blocks for "STAT Reserve", "Tax Reserves", "GAAP Ben Reserves", etc.
    """

    block_name: str           # section title, e.g. "STAT Reserve"
    dimensions: list[str]     # row-label columns (text), e.g. ["Product Subtype"]
    kpi_measures: list[str]   # KPI/measure column names from the header row
    header_row: int           # 1-based worksheet row index of the block header
    data_start_row: int       # 1-based row index of the first data row
    data_end_row: int         # 1-based row index of the last data row


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
    # Pivot metadata — populated when sheet contains or feeds a pivot table
    is_pivot_sheet: bool = False             # True when this sheet has pivot tables
    pivot_source_tabs: list[str] = field(default_factory=list)   # sheets feeding this pivot
    pivot_value_fields: list[str] = field(default_factory=list)  # measure/KPI field names
    # Multi-block KPI structure — populated for KPI_SUMMARY / DASHBOARD / OUTPUT tabs
    kpi_blocks: list[KpiBlock] = field(default_factory=list)


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

    @property
    def pivot_kpi_tabs(self) -> list[str]:
        """Sheets with pivot tables classified as KPI summary tabs."""
        return [t.tab_name for t in self.tab_analyses if t.is_pivot_sheet]

    @property
    def all_pivot_value_fields(self) -> list[str]:
        """All pivot value (measure) fields across all KPI tabs."""
        fields: list[str] = []
        for t in self.tab_analyses:
            for f in t.pivot_value_fields:
                if f not in fields:
                    fields.append(f)
        return fields

    def detection_type(self) -> str:
        """'Formula-Based' | 'Pivot-Based' | 'Mixed'."""
        has_formula_kpi = any(
            t.tab_type in (TabType.KPI_SUMMARY, TabType.DASHBOARD, TabType.OUTPUT)
            and not t.is_pivot_sheet
            for t in self.tab_analyses
        )
        has_pivot_kpi = any(t.is_pivot_sheet for t in self.tab_analyses)
        if has_formula_kpi and has_pivot_kpi:
            return "Mixed"
        if has_pivot_kpi:
            return "Pivot-Based"
        return "Formula-Based"


# ---------------------------------------------------------------------------
# Pivot pre-scanner
# ---------------------------------------------------------------------------

_BUSINESS_MEASURE_WORDS = frozenset({
    "reserve", "reserves", "count", "counts", "balance", "balances",
    "premium", "premiums", "claim", "claims", "amount", "amounts",
    "total", "sum", "net", "gross", "earned", "incurred", "paid",
    "outstanding", "ibnr", "ulr", "loss", "expense", "exposure",
    "policy", "policies", "rate", "fee", "revenue", "cost",
    "quota", "share", "stat", "statutory", "gaap",
})


def _looks_like_measure(name: str) -> bool:
    words = set(re.split(r"[\s_\-]+", name.lower()))
    return bool(words & _BUSINESS_MEASURE_WORDS)


@dataclass
class _PivotInfo:
    """Pivot table metadata extracted from a single sheet."""
    sheet_name: str         # the sheet that contains the pivot table
    source_sheets: list[str]
    value_fields: list[str]
    row_fields: list[str]
    filter_fields: list[str]


def _find_source_by_cache_fields(
    raw_wb: Any,
    cache_fields: list[str],
    pivot_sheet: str,
    min_overlap: int = 2,
) -> str:
    """Heuristic fallback: find the sheet whose column headers best match
    the pivot cache field names.  Used when worksheetSource.sheet is absent.

    Returns the sheet name with the highest overlap, or "" if no candidate
    reaches ``min_overlap`` matching fields.
    """
    if raw_wb is None or not cache_fields:
        return ""
    cache_lower = {f.lower() for f in cache_fields if f}
    best_sheet = ""
    best_overlap = 0
    for sn in raw_wb.sheetnames:
        if sn == pivot_sheet:
            continue
        try:
            ws = raw_wb[sn]
            first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            if not first_row:
                continue
            col_names = {str(v).lower() for v in first_row if v is not None}
            overlap = len(cache_lower & col_names)
            if overlap > best_overlap:
                best_overlap = overlap
                best_sheet = sn
        except Exception:
            continue
    return best_sheet if best_overlap >= min_overlap else ""


def _scan_pivots(raw_wb: Any) -> list[_PivotInfo]:
    """Read all pivot tables in a workbook via openpyxl metadata.

    Returns one _PivotInfo per sheet that contains at least one pivot table.
    Fields with no explicit dataField list are heuristically filtered for
    business-measure keywords.

    Source sheet detection uses three strategies in order:
    1. ``cache.worksheetSource.sheet`` — the explicit sheet name (most reliable).
    2. ``cache.worksheetSource.name`` → workbook defined-names lookup.
    3. Cache-field name matching against every other sheet's column headers.
    """
    if raw_wb is None:
        return []

    infos: list[_PivotInfo] = []
    for sheet_name in raw_wb.sheetnames:
        try:
            raw_ws = raw_wb[sheet_name]
        except Exception:
            continue
        pivots = getattr(raw_ws, "_pivots", [])
        if not pivots:
            continue

        source_sheets: list[str] = []
        value_fields: list[str] = []
        row_fields_raw: list[str] = []
        filter_fields_raw: list[str] = []
        all_cache_fields: list[str] = []   # accumulated across all pivots on this sheet

        for pv in pivots:
            cache = getattr(pv, "cache", None)
            if cache is not None:
                ws_src = getattr(cache, "worksheetSource", None)
                if ws_src is not None:
                    # Strategy 1: explicit sheet attribute
                    src_sheet = getattr(ws_src, "sheet", None) or ""
                    if src_sheet and src_sheet not in source_sheets:
                        source_sheets.append(src_sheet)
                    else:
                        # Strategy 2: named range — resolve via workbook defined names
                        named = getattr(ws_src, "name", None) or ""
                        if named:
                            for dn_key in getattr(raw_wb, "defined_names", {}):
                                if dn_key == named:
                                    dn_val = raw_wb.defined_names[dn_key]
                                    for dest_sheet, _ in getattr(dn_val, "destinations", []):
                                        if dest_sheet and dest_sheet not in source_sheets:
                                            source_sheets.append(dest_sheet)
                                    break

                # All field names in the cache
                cache_fields: list[str] = [
                    getattr(f, "name", None)
                    for f in getattr(cache, "cacheFields", [])
                    if getattr(f, "name", None)
                ]
                for cf in cache_fields:
                    if cf not in all_cache_fields:
                        all_cache_fields.append(cf)
            else:
                cache_fields = []

            # Value (data) fields
            data_fields_obj = getattr(pv, "dataFields", None)
            pv_value_fields: list[str] = []
            if data_fields_obj is not None:
                for df in getattr(data_fields_obj, "dataField", []):
                    idx = getattr(df, "field", None)
                    name = getattr(df, "name", None) or (
                        cache_fields[idx] if idx is not None and 0 <= idx < len(cache_fields) else None
                    )
                    if name:
                        pv_value_fields.append(name)

            # Fallback: use business-measure heuristic on all cache fields
            if not pv_value_fields:
                pv_value_fields = [f for f in cache_fields if _looks_like_measure(f)]

            for f in pv_value_fields:
                if f not in value_fields:
                    value_fields.append(f)

            # Row fields
            row_fields_obj = getattr(pv, "rowFields", None)
            if row_fields_obj is not None:
                for rf in getattr(row_fields_obj, "field", []):
                    idx = getattr(rf, "x", None)
                    if idx is not None and 0 <= idx < len(cache_fields):
                        name = cache_fields[idx]
                        if name and name not in row_fields_raw:
                            row_fields_raw.append(name)

            # Filter fields
            pf_obj = getattr(pv, "pageFields", None)
            if pf_obj is not None:
                for pf in getattr(pf_obj, "pageField", []):
                    idx = getattr(pf, "field", None)
                    if idx is not None and 0 <= idx < len(cache_fields):
                        name = cache_fields[idx]
                        if name and name not in filter_fields_raw:
                            filter_fields_raw.append(name)

        # Strategy 3: if no source sheet found yet, match cache fields against
        # all other sheets' column headers to find the best candidate.
        if not source_sheets and all_cache_fields:
            fallback = _find_source_by_cache_fields(raw_wb, all_cache_fields, sheet_name)
            if fallback:
                source_sheets.append(fallback)
                logger.info(
                    "Pivot source fallback: '%s' matched to '%s' via cache-field overlap",
                    sheet_name, fallback,
                )
            else:
                logger.warning(
                    "Pivot on '%s': could not identify source sheet "
                    "(worksheetSource.sheet absent and no cache-field match found)",
                    sheet_name,
                )

        infos.append(_PivotInfo(
            sheet_name=sheet_name,
            source_sheets=source_sheets,
            value_fields=value_fields,
            row_fields=row_fields_raw,
            filter_fields=filter_fields_raw,
        ))

    return infos


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

    # ── Pivot pre-scan ───────────────────────────────────────────────────────
    # Read all pivot tables BEFORE signal classification so we can override
    # tab types with certainty: pivot sheets → KPI_SUMMARY, their source
    # sheets → SOURCE_DATA (regardless of formula density or row count).
    pivot_infos = _scan_pivots(raw_wb)
    # pivot_by_sheet: sheet_name → _PivotInfo (for sheets that contain pivots)
    pivot_by_sheet: dict[str, _PivotInfo] = {p.sheet_name: p for p in pivot_infos}
    # sheets that are referenced as the data source by ≥1 pivot
    pivot_source_sheet_names: set[str] = {
        src for p in pivot_infos for src in p.source_sheets
    }

    tab_analyses: list[TabAnalysis] = []
    kpi_definitions: list[KPIDefinition] = []

    for sheet_name, df in bundle.sheets.items():
        raw_ws = _get_raw_ws(raw_wb, sheet_name)
        signals = _extract_signals(raw_ws, df)
        tab = _classify_tab(signals, sheet_name, wb_name, cfg)

        # ── Pivot override ───────────────────────────────────────────────────
        if sheet_name in pivot_by_sheet:
            pinfo = pivot_by_sheet[sheet_name]
            # A pivot sheet is a KPI summary tab — never source data
            tab.tab_type = TabType.KPI_SUMMARY
            tab.is_pivot_sheet = True
            tab.pivot_source_tabs = pinfo.source_sheets
            tab.pivot_value_fields = pinfo.value_fields
            n_kpi = len(pinfo.value_fields)
            src_str = ", ".join(pinfo.source_sheets) if pinfo.source_sheets else "unknown"
            tab.classification_reason = (
                f"PIVOT OVERRIDE: sheet contains pivot table(s) with {n_kpi} value "
                f"field(s) — classified as KPI_SUMMARY. "
                f"Pivot data source: [{src_str}]. "
                f"KPI fields: {', '.join(pinfo.value_fields[:5])}"
                + ("…" if n_kpi > 5 else "") + "."
            )
            logger.info(
                "Pivot KPI tab detected: '%s/%s' — %d value field(s), source: %s",
                wb_name, sheet_name, n_kpi, src_str,
            )

        elif sheet_name in pivot_source_sheet_names:
            # A sheet that feeds a pivot is source data — even with formula cols
            if tab.tab_type not in (TabType.SOURCE_DATA, TabType.REFERENCE_DATA):
                tab.tab_type = TabType.SOURCE_DATA
                consumers = [
                    p.sheet_name for p in pivot_infos if sheet_name in p.source_sheets
                ]
                tab.classification_reason = (
                    f"PIVOT SOURCE OVERRIDE: this sheet is the data source for pivot "
                    f"table(s) on [{', '.join(consumers)}]. Classified as SOURCE_DATA. "
                    f"Row count: {signals.row_count:,}, cols: {signals.col_count}. "
                    + tab.classification_reason
                )
                logger.info(
                    "Source tab promoted via pivot trace: '%s/%s' feeds pivot(s) on %s",
                    wb_name, sheet_name, consumers,
                )

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
            # KPI block detection is intentionally deferred — it is computed
            # lazily (and cached) by the UI when a KPI tab is selected in the
            # Column Explorer.  Running it here during analysis would scan up
            # to 2 000 rows per KPI tab on every discovery run.

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
    """Analyse multiple bundles in parallel, skipping failures with a warning.

    Uses a ThreadPoolExecutor so multiple workbooks are analysed concurrently.
    Results are returned in the same order as the input bundles.
    """
    if not bundles:
        return []

    results: list[WorkbookAnalysis | None] = [None] * len(bundles)

    def _analyse(idx: int, bundle) -> tuple[int, WorkbookAnalysis]:
        return idx, analyze_workbook(bundle, config=config)

    with ThreadPoolExecutor(max_workers=min(len(bundles), 8)) as pool:
        futures = {pool.submit(_analyse, i, b): i for i, b in enumerate(bundles)}
        for future in as_completed(futures):
            try:
                idx, analysis = future.result()
                results[idx] = analysis
            except Exception as exc:
                bundle = bundles[futures[future]]
                logger.warning(
                    "Could not analyse '%s': %s",
                    getattr(bundle, "file_name", "?"), exc,
                )

    return [r for r in results if r is not None]


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

def _extract_signals(raw_ws: Any, df: pd.DataFrame, max_scan_rows: int = 500) -> SignalVector:
    """Compute all measurable signals from a worksheet.

    For large worksheets the cell scan is capped at *max_scan_rows* to keep
    discovery fast.  Classification signals (formula density, inter-sheet refs)
    are estimated from the sample, which is accurate enough for tab typing.
    """
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
    rows_scanned = 0

    for row in raw_ws.iter_rows():
        if rows_scanned >= max_scan_rows:
            break
        rows_scanned += 1
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

    # If we sampled fewer rows than the full sheet, scale formula_cells up
    # proportionally so formula_density reflects the whole tab accurately.
    if rows_scanned > 0 and row_count > rows_scanned:
        scale = row_count / rows_scanned
        formula_cells_est = int(formula_cells * scale)
        non_empty_est     = int(non_empty * scale)
    else:
        formula_cells_est = formula_cells
        non_empty_est     = non_empty

    formula_density = formula_cells_est / non_empty_est if non_empty_est > 0 else 0.0

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


_LARGE_SHEET_ROW_THRESHOLD = 500  # sheets with more rows use sampled scan


_PANDAS_SAMPLE_ROWS = 5_000  # cap for nunique / dtype scans


def _compute_avg_unique_ratio(df: pd.DataFrame) -> float:
    """Mean of (unique_values / total_values) across all non-empty columns.

    Sampled to at most _PANDAS_SAMPLE_ROWS rows so large source tabs don't
    dominate discovery time; uniqueness ratios are stable with small samples.
    """
    if df.empty or len(df.columns) == 0:
        return 0.0
    sample = df if len(df) <= _PANDAS_SAMPLE_ROWS else df.iloc[:_PANDAS_SAMPLE_ROWS]
    ratios = []
    for col in sample.columns:
        series = sample[col].dropna()
        if len(series) > 0:
            ratios.append(series.nunique() / len(series))
    return sum(ratios) / len(ratios) if ratios else 0.0


def _compute_dtype_homogeneity(df: pd.DataFrame) -> float:
    """Fraction of columns sharing the single most common inferred dtype."""
    if df.empty or len(df.columns) == 0:
        return 0.0
    # dtypes are column-level metadata — no row scan needed
    dtype_counts: dict[str, int] = {}
    for dtype in df.dtypes:
        key = str(dtype)
        dtype_counts[key] = dtype_counts.get(key, 0) + 1
    majority = max(dtype_counts.values())
    return majority / len(df.columns)


def _has_aggregate_function(formula: str) -> bool:
    return bool(_AGG_FUNC_RE.search(formula))


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
    # Large row count is a strong signal for source data even with formula columns
    if sv.row_count >= 500:
        scores[TabType.SOURCE_DATA] += cfg.w_row_count * 2.0
    if sv.formula_density <= cfg.source_formula_density_max:
        scores[TabType.SOURCE_DATA] += cfg.w_formula_density * (1.0 - sv.formula_density / max(cfg.source_formula_density_max, 0.001))
    elif sv.row_count >= 500:
        # Large tabs with formula columns still get partial source credit
        scores[TabType.SOURCE_DATA] += cfg.w_formula_density * 0.5
    if sv.dtype_homogeneity >= cfg.homogeneity_source_min:
        scores[TabType.SOURCE_DATA] += cfg.w_col_homogeneity * sv.dtype_homogeneity
    if sv.inter_sheet_ref_count == 0:
        scores[TabType.SOURCE_DATA] += cfg.w_inter_sheet_refs * 0.5
    elif sv.row_count >= 500:
        # Referenced from other sheets is fine for large source tabs
        scores[TabType.SOURCE_DATA] += cfg.w_inter_sheet_refs * 0.25

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
    max_scan_rows: int = 500,
) -> list[KPIDefinition]:
    kpis: list[KPIDefinition] = []

    for row_idx_0, row in enumerate(raw_ws.iter_rows()):
        if row_idx_0 >= max_scan_rows:
            break
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


# ---------------------------------------------------------------------------
# KPI block scanner
# ---------------------------------------------------------------------------

_KPI_BLOCK_MAX_BLANK_GAP = 3  # consecutive blank rows that terminate a block


def scan_kpi_blocks(raw_ws: Any, max_rows: int = 2_000) -> list[KpiBlock]:
    """Detect multi-section (block) structure in a KPI/summary worksheet.

    Many actuarial reserve summary tabs contain multiple self-contained sections,
    each with a title row, a header row, and data rows.  This function identifies
    those boundaries by detecting the pattern::

        [title row: 1-3 text cells, no numbers]
        [optional blank rows]
        [header row: ≥2 text cells, no numbers]
        [data rows: dimension text + KPI formulas/numbers]
        [blank separator rows]
        [next title ...]

    Returns an empty list when no multi-block structure is found (i.e. the sheet
    looks like a single flat table — normal DataFrame reading is appropriate).
    """
    if raw_ws is None:
        return []

    try:
        raw_rows = []
        for i, row in enumerate(raw_ws.iter_rows(values_only=True)):
            if i >= max_rows:
                break
            raw_rows.append(row)
    except Exception:
        return []

    if not raw_rows:
        return []

    n_cols = max(len(r) for r in raw_rows)
    if n_cols == 0:
        return []

    # Pad rows to uniform width
    rows: list[list] = [list(r) + [None] * (n_cols - len(r)) for r in raw_rows]
    n_rows = len(rows)

    def _blank(v: Any) -> bool:
        return v is None or (isinstance(v, str) and v.strip() == "")

    def _text(v: Any) -> bool:
        return isinstance(v, str) and not v.startswith("=") and v.strip() != ""

    def _numeric_or_formula(v: Any) -> bool:
        if isinstance(v, (int, float)):
            return True
        return isinstance(v, str) and v.startswith("=")

    def _row_stats(row: list) -> tuple[int, int, int]:
        """Return (non_empty_count, text_count, numeric_or_formula_count)."""
        ne = sum(1 for v in row if not _blank(v))
        nt = sum(1 for v in row if _text(v))
        nn = sum(1 for v in row if _numeric_or_formula(v))
        return ne, nt, nn

    rstats: list[tuple[int, int, int]] = [_row_stats(r) for r in rows]

    def _next_nonempty(start: int) -> int:
        for i in range(start, n_rows):
            if rstats[i][0] > 0:
                return i
        return n_rows

    # ── Find header rows ─────────────────────────────────────────────────────
    # Header: ≥2 non-empty cells, ≥50% are text, zero numeric/formula cells,
    # AND the immediately following non-blank row has ≥1 numeric/formula cell.
    candidate_headers: list[int] = []
    for i in range(n_rows):
        ne, nt, nn = rstats[i]
        if ne >= 2 and nn == 0 and nt >= max(1, ne * 0.5):
            j = _next_nonempty(i + 1)
            if j < n_rows and rstats[j][2] > 0:
                candidate_headers.append(i)

    if not candidate_headers:
        return []

    # Deduplicate adjacent candidates (keep the later one)
    deduped: list[int] = []
    for h in candidate_headers:
        if deduped and h - deduped[-1] <= 1:
            deduped[-1] = h
        else:
            deduped.append(h)
    candidate_set: set[int] = set(deduped)

    # ── Build one KpiBlock per header ────────────────────────────────────────
    blocks: list[KpiBlock] = []

    for h_idx in deduped:
        # ── Title: look back up to 5 rows for a 1-3 cell pure-text row ──────
        title = ""
        for back in range(1, 6):
            t = h_idx - back
            if t < 0:
                break
            ne, nt, nn = rstats[t]
            if ne == 0:
                continue  # blank row — keep looking back
            # Title: 1–3 non-empty cells, all text, no numeric values
            if 1 <= ne <= 3 and nt == ne and nn == 0:
                text_vals = [v for v in rows[t] if _text(v)]
                if text_vals and len(str(text_vals[0]).strip()) > 2:
                    title = str(text_vals[0]).strip()
            break  # stop at the first non-blank row (title or not)

        # ── Header column names ───────────────────────────────────────────────
        col_names: list[str | None] = [
            str(v).strip() if not _blank(v) else None
            for v in rows[h_idx]
        ]

        # ── Collect data rows ─────────────────────────────────────────────────
        k = h_idx + 1
        data_start = k
        data_end = h_idx
        consecutive_blanks = 0
        sample_txt: list[int] = [0] * n_cols
        sample_num: list[int] = [0] * n_cols
        sample_count = 0
        data_row_indices: list[int] = []

        while k < n_rows:
            ne, nt, nn = rstats[k]
            if ne == 0:
                consecutive_blanks += 1
                if consecutive_blanks >= _KPI_BLOCK_MAX_BLANK_GAP:
                    break
                k += 1
                continue
            consecutive_blanks = 0

            # Stop at the next detected header
            if k in candidate_set and k != h_idx:
                break

            data_end = k
            data_row_indices.append(k)
            if sample_count < 5:
                for ci, v in enumerate(rows[k][:n_cols]):
                    if _text(v):
                        sample_txt[ci] += 1
                    elif _numeric_or_formula(v):
                        sample_num[ci] += 1
                sample_count += 1
            k += 1

        if not data_row_indices:
            continue

        # ── Classify columns as dimension or KPI measure ──────────────────────
        # Dimension: column has only text values in data rows (row labels).
        # KPI measure: column has numeric or formula values.
        dimensions: list[str] = []
        kpi_measures: list[str] = []
        for ci, col_name in enumerate(col_names):
            if col_name is None:
                continue
            if sample_txt[ci] > 0 and sample_num[ci] == 0:
                dimensions.append(col_name)
            else:
                kpi_measures.append(col_name)

        if not kpi_measures:
            continue

        blocks.append(KpiBlock(
            block_name=title or f"Block {len(blocks) + 1}",
            dimensions=dimensions,
            kpi_measures=kpi_measures,
            header_row=h_idx + 1,
            data_start_row=data_start + 1,
            data_end_row=data_end + 1,
        ))

    # Only surface blocks when there are multiple, or a single named block.
    # A single "Block 1" (no real title found) is just a normal table.
    named = [b for b in blocks if not b.block_name.startswith("Block ")]
    if len(blocks) >= 2 or named:
        return blocks
    return []


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
