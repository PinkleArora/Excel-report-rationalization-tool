"""Build the Workbook Analysis report.

Produces ``output/workbook_analysis.xlsx`` — a nine-sheet diagnostic workbook
that shows exactly how the application interpreted every uploaded file,
without requiring access to the source files themselves.

Sheets produced
---------------
01_Workbook_Inventory       One row per uploaded workbook: metadata summary.
02_Source_Tab_Detection     Tabs classified as data-bearing, with signal scores.
03_Summary_Tab_Detection    Tabs classified as formula/KPI/dashboard, with scores.
04_KPI_Inventory            All extracted KPIs with rationalization verdict.
05_Formula_Inventory        Detailed formula-level listing from formula tabs.
06_Schema_Mapping           Every source column with match confidence and verdict.
07_Collision_Analysis       Intra-frame collisions and many-to-one mapping flags.
08_Rationalization_Candidates   KPIs confirmed as consolidation candidates.
09_Rationalization_Exclusions   KPIs excluded from rationalization, with reasons.
"""

from __future__ import annotations

import io
import logging
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pandas as pd

from src.ingestion.workbook_analyzer import (
    ClassificationConfig,
    TabType,
    WorkbookAnalysis,
    analyze_many,
    build_dependency_report_df,
    build_kpi_inventory_df,
)
from src.profiling.profiler import profile_many
from src.schema_matching.matcher import (
    MatchConfidence,
    detect_many_to_one_mappings,
    match_all_bundles,
)
from src.consolidation.consolidator import build_source_mapping_df
from src.workbook_generator.generator import SheetSpec, generate_workbook

logger = logging.getLogger(__name__)

DEFAULT_OUTPUT_PATH = Path("output") / "workbook_analysis.xlsx"

# Tab types treated as data-bearing for schema matching
_DATA_BEARING = {TabType.SOURCE_DATA, TabType.REFERENCE_DATA}

# Tab types treated as formula/summary for KPI extraction
_SUMMARY_TYPES = {
    TabType.KPI_SUMMARY, TabType.DASHBOARD,
    TabType.OUTPUT, TabType.CALCULATION,
}


@dataclass
class AnalysisContext:
    """All intermediate artefacts produced during analysis."""

    bundles: list
    analyses: list[WorkbookAnalysis] = field(default_factory=list)
    collision_log: list = field(default_factory=list)
    many_to_one_issues: list = field(default_factory=list)
    column_matches: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_analysis_workbook(
    bundles: list,
    output_path: str | Path | None = None,
    matching_threshold: float = 80.0,
    classification_config: ClassificationConfig | None = None,
) -> Path:
    """Run the full analysis pipeline and write ``workbook_analysis.xlsx``.

    Args:
        bundles: Loaded :class:`~src.ingestion.loader.WorkbookBundle` objects.
        output_path: Destination path.  Defaults to ``output/workbook_analysis.xlsx``.
        matching_threshold: Minimum fuzzy-match score for schema matching.
        classification_config: Optional :class:`ClassificationConfig` to tune
            tab classification thresholds and weights.

    Returns:
        Absolute :class:`pathlib.Path` of the written workbook.
    """
    resolved = Path(output_path) if output_path else DEFAULT_OUTPUT_PATH
    ctx = _run_analysis(bundles, matching_threshold, classification_config)

    sheets = [
        SheetSpec("01_Workbook_Inventory",       _sheet_workbook_inventory(ctx)),
        SheetSpec("02_Source_Tab_Detection",      _sheet_source_tabs(ctx)),
        SheetSpec("03_Summary_Tab_Detection",     _sheet_summary_tabs(ctx)),
        SheetSpec("04_KPI_Inventory",             _sheet_kpi_inventory(ctx)),
        SheetSpec("05_Formula_Inventory",         _sheet_formula_inventory(ctx)),
        SheetSpec("06_Schema_Mapping",            _sheet_schema_mapping(ctx)),
        SheetSpec("07_Collision_Analysis",        _sheet_collision_analysis(ctx)),
        SheetSpec("08_Rationalization_Candidates", _sheet_rationalization_candidates(ctx)),
        SheetSpec("09_Rationalization_Exclusions", _sheet_rationalization_exclusions(ctx)),
    ]

    path = generate_workbook(sheets, resolved)
    logger.info("Analysis workbook written: %s", path)
    return path


def build_analysis_workbook_bytes(
    bundles: list,
    matching_threshold: float = 80.0,
    classification_config: ClassificationConfig | None = None,
) -> bytes:
    """Build the analysis workbook in-memory and return raw bytes."""
    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp) / "workbook_analysis.xlsx"
        build_analysis_workbook(
            bundles,
            output_path=tmp_path,
            matching_threshold=matching_threshold,
            classification_config=classification_config,
        )
        return tmp_path.read_bytes()


# ---------------------------------------------------------------------------
# Analysis pipeline
# ---------------------------------------------------------------------------

def _run_analysis(
    bundles: list,
    matching_threshold: float,
    config: ClassificationConfig | None,
) -> AnalysisContext:
    ctx = AnalysisContext(bundles=bundles)

    logger.info("Analysis: classifying %d bundle(s)…", len(bundles))
    ctx.analyses = analyze_many(bundles, config=config)

    # Schema matching across data-bearing tabs only
    import copy
    data_bundles = []
    for bundle in bundles:
        analysis = next((a for a in ctx.analyses if a.workbook_name == bundle.file_name), None)
        if analysis is None:
            data_bundles.append(bundle)
            continue
        data_sheets = set(analysis.data_bearing_tabs(types=_DATA_BEARING))
        if not data_sheets:
            data_bundles.append(bundle)
            continue
        nb = copy.copy(bundle)
        nb.sheets = {k: v for k, v in bundle.sheets.items() if k in data_sheets}
        if nb.sheets:
            data_bundles.append(nb)

    logger.info("Analysis: schema matching across %d data bundle(s)…", len(data_bundles))
    ctx.column_matches = match_all_bundles(data_bundles, threshold=matching_threshold)
    ctx.many_to_one_issues = detect_many_to_one_mappings(ctx.column_matches)

    # Build source mapping (also populates collision_log)
    build_source_mapping_df(data_bundles, ctx.column_matches, collision_log=ctx.collision_log)

    return ctx


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

def _sheet_workbook_inventory(ctx: AnalysisContext) -> pd.DataFrame:
    """One row per uploaded workbook."""
    wm_list = profile_many(ctx.bundles)
    rows = []
    for wm in wm_list:
        analysis = next(
            (a for a in ctx.analyses if a.workbook_name == wm.workbook_name), None
        )
        source_tabs  = analysis.source_tabs  if analysis else []
        ref_tabs     = analysis.reference_tabs  if analysis else []
        summary_tabs = analysis.summary_tabs if analysis else []
        calc_tabs    = analysis.calculation_tabs if analysis else []
        unknown_tabs = [
            t.tab_name for t in (analysis.tab_analyses if analysis else [])
            if t.tab_type in (TabType.MIXED, TabType.UNKNOWN)
        ]
        rows.append({
            "workbook_name":        wm.workbook_name,
            "file_extension":       wm.file_extension,
            "total_sheets":         wm.sheet_count,
            "total_rows":           wm.total_rows,
            "total_columns":        wm.total_cols,
            "formula_cells":        wm.total_formula_cells,
            "source_data_tabs":     ", ".join(source_tabs),
            "reference_tabs":       ", ".join(ref_tabs),
            "summary_kpi_tabs":     ", ".join(summary_tabs),
            "calculation_tabs":     ", ".join(calc_tabs),
            "mixed_unknown_tabs":   ", ".join(unknown_tabs),
            "kpi_definitions_found": sum(
                1 for a in ctx.analyses
                if a.workbook_name == wm.workbook_name
                for _ in a.kpi_definitions
            ),
            "analysis_date":        str(date.today()),
        })
    return pd.DataFrame(rows)


def _sheet_source_tabs(ctx: AnalysisContext) -> pd.DataFrame:
    """Tabs classified as data-bearing (SOURCE_DATA, REFERENCE_DATA, MAPPING_TABLE)."""
    data_types = {TabType.SOURCE_DATA, TabType.REFERENCE_DATA, TabType.MAPPING_TABLE}
    rows = []
    for analysis in ctx.analyses:
        for tab in analysis.tab_analyses:
            if tab.tab_type not in data_types:
                continue
            rows.append(_tab_row(analysis.workbook_name, tab))
    if not rows:
        return pd.DataFrame(columns=_tab_columns())
    return pd.DataFrame(rows)


def _sheet_summary_tabs(ctx: AnalysisContext) -> pd.DataFrame:
    """Tabs classified as formula-driven (KPI_SUMMARY, DASHBOARD, OUTPUT, CALCULATION)."""
    formula_types = {
        TabType.KPI_SUMMARY, TabType.DASHBOARD,
        TabType.OUTPUT, TabType.CALCULATION,
    }
    rows = []
    for analysis in ctx.analyses:
        for tab in analysis.tab_analyses:
            if tab.tab_type not in formula_types:
                continue
            rows.append(_tab_row(analysis.workbook_name, tab))
    if not rows:
        return pd.DataFrame(columns=_tab_columns())
    return pd.DataFrame(rows)


def _sheet_kpi_inventory(ctx: AnalysisContext) -> pd.DataFrame:
    """Full KPI inventory with rationalization verdicts."""
    df = build_kpi_inventory_df(ctx.analyses)
    if df.empty:
        return df
    # Reorder columns for readability
    preferred_order = [
        "workbook_name", "source_tab", "kpi_label", "aggregate_function",
        "formula", "referenced_sheets", "referenced_columns", "cell_address",
        "is_common", "rationalization_candidate", "rationalization_note",
    ]
    cols = [c for c in preferred_order if c in df.columns]
    return df[cols]


def _sheet_formula_inventory(ctx: AnalysisContext) -> pd.DataFrame:
    """Detailed formula-level listing extracted from formula tabs."""
    rows = []
    for analysis in ctx.analyses:
        for kpi in analysis.kpi_definitions:
            rows.append({
                "workbook_name":      kpi.workbook_name,
                "tab_name":           kpi.source_tab,
                "cell_address":       kpi.cell_address,
                "kpi_label":          kpi.kpi_label,
                "formula":            kpi.formula,
                "aggregate_function": kpi.aggregate_function,
                "referenced_sheets":  ", ".join(kpi.referenced_sheets),
                "referenced_columns": ", ".join(kpi.referenced_columns),
                "cross_sheet":        len(kpi.referenced_sheets) > 0,
            })
    if not rows:
        return pd.DataFrame(columns=[
            "workbook_name", "tab_name", "cell_address", "kpi_label", "formula",
            "aggregate_function", "referenced_sheets", "referenced_columns", "cross_sheet",
        ])
    return pd.DataFrame(rows)


def _sheet_schema_mapping(ctx: AnalysisContext) -> pd.DataFrame:
    """Full source mapping with confidence levels and review flags."""
    import copy
    data_bundles = []
    for bundle in ctx.bundles:
        analysis = next((a for a in ctx.analyses if a.workbook_name == bundle.file_name), None)
        if analysis is None:
            data_bundles.append(bundle)
            continue
        data_sheets = set(analysis.data_bearing_tabs(types=_DATA_BEARING))
        if not data_sheets:
            data_bundles.append(bundle)
            continue
        nb = copy.copy(bundle)
        nb.sheets = {k: v for k, v in bundle.sheets.items() if k in data_sheets}
        if nb.sheets:
            data_bundles.append(nb)

    df = build_source_mapping_df(data_bundles, ctx.column_matches)
    if df.empty:
        return df
    preferred_order = [
        "source_workbook", "source_sheet", "source_column", "canonical_column",
        "match_type", "match_score", "match_confidence", "safe_to_merge",
        "review_required", "matched_to_workbook", "matched_to_sheet",
        "matched_to_column", "is_matched",
    ]
    cols = [c for c in preferred_order if c in df.columns]
    return df[cols]


def _sheet_collision_analysis(ctx: AnalysisContext) -> pd.DataFrame:
    """Combined collision and many-to-one issue report."""
    rows = []

    for c in ctx.collision_log:
        rows.append({
            "issue_type":       "intra_frame_collision",
            "severity":         "MEDIUM",
            "workbook":         c.get("workbook", ""),
            "sheet":            c.get("sheet", ""),
            "source_column":    c.get("source_col", ""),
            "target_column":    c.get("target_col", ""),
            "canonical_attempted": c.get("canonical_attempted", ""),
            "colliding_columns": ", ".join(c.get("colliding_columns", [])),
            "description": (
                f"Matching '{c.get('source_col','')}' → "
                f"'{c.get('target_col','')}' would create duplicate "
                f"canonical name '{c.get('canonical_attempted','')}' within "
                f"sheet '{c.get('sheet','')}'. Match suppressed."
            ),
            "action_required":  "Review whether these are truly the same column.",
        })

    for m in ctx.many_to_one_issues:
        rows.append({
            "issue_type":          "many_to_one_mapping",
            "severity":            "HIGH",
            "workbook":            "multiple",
            "sheet":               "multiple",
            "source_column":       " | ".join(m.get("source_columns", [])),
            "target_column":       m.get("target_canonical", ""),
            "canonical_attempted": m.get("target_canonical", ""),
            "colliding_columns":   " | ".join(m.get("source_columns", [])),
            "description":         m.get("description", ""),
            "action_required": (
                "Review 06_Schema_Mapping. Confirm which source columns are "
                "genuinely equivalent before enabling merge."
            ),
        })

    if not rows:
        return pd.DataFrame({
            "issue_type": ["none"], "severity": ["INFO"],
            "description": ["No collisions or many-to-one mappings detected."],
            "action_required": [""],
            "workbook": [""], "sheet": [""], "source_column": [""],
            "target_column": [""], "canonical_attempted": [""],
            "colliding_columns": [""],
        })
    return pd.DataFrame(rows)


def _sheet_rationalization_candidates(ctx: AnalysisContext) -> pd.DataFrame:
    """KPIs confirmed as rationalization candidates."""
    df = build_kpi_inventory_df(ctx.analyses)
    if df.empty:
        return pd.DataFrame(columns=[
            "workbook_name", "source_tab", "kpi_label", "aggregate_function",
            "formula", "referenced_columns", "rationalization_note",
        ])
    candidates = df[df["rationalization_candidate"] == True].copy()  # noqa: E712
    if candidates.empty:
        return pd.DataFrame({
            "message": ["No rationalization candidates detected. "
                        "No two workbooks share equivalent KPIs with matching "
                        "aggregate functions and overlapping source fields."]
        })
    preferred = [
        "workbook_name", "source_tab", "kpi_label", "aggregate_function",
        "formula", "referenced_sheets", "referenced_columns",
        "cell_address", "is_common", "rationalization_note",
    ]
    cols = [c for c in preferred if c in candidates.columns]
    return candidates[cols].reset_index(drop=True)


def _sheet_rationalization_exclusions(ctx: AnalysisContext) -> pd.DataFrame:
    """KPIs seen in multiple workbooks but excluded from rationalization, with reasons."""
    df = build_kpi_inventory_df(ctx.analyses)
    if df.empty:
        return pd.DataFrame(columns=[
            "workbook_name", "source_tab", "kpi_label", "aggregate_function",
            "formula", "referenced_columns", "rationalization_note", "exclusion_reason",
        ])
    # Exclusions = appears in ≥2 workbooks (is_common) but NOT a candidate
    exclusions = df[
        (df["is_common"] == True) &  # noqa: E712
        (df["rationalization_candidate"] == False)  # noqa: E712
    ].copy()
    if exclusions.empty:
        return pd.DataFrame({
            "message": ["No exclusions. All common KPIs are either rationalization "
                        "candidates or unique to a single workbook."]
        })
    exclusions = exclusions.rename(columns={"rationalization_note": "exclusion_reason"})
    preferred = [
        "workbook_name", "source_tab", "kpi_label", "aggregate_function",
        "formula", "referenced_sheets", "referenced_columns",
        "cell_address", "is_common", "exclusion_reason",
    ]
    cols = [c for c in preferred if c in exclusions.columns]
    return exclusions[cols].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Tab row helpers
# ---------------------------------------------------------------------------

_DATA_TYPE_LABELS = {
    TabType.SOURCE_DATA:    "Source Data",
    TabType.REFERENCE_DATA: "Reference / Lookup",
    TabType.MAPPING_TABLE:  "Mapping Table",
    TabType.KPI_SUMMARY:    "KPI / Summary",
    TabType.DASHBOARD:      "Dashboard",
    TabType.OUTPUT:         "Output",
    TabType.CALCULATION:    "Calculation Helper",
    TabType.MIXED:          "Mixed",
    TabType.UNKNOWN:        "Unknown",
}


def _tab_columns() -> list[str]:
    return [
        "workbook_name", "tab_name", "tab_type", "tab_type_label",
        "row_count", "col_count", "formula_cells", "formula_density_pct",
        "inter_sheet_refs", "referenced_sheets",
        "top_signal_scores", "classification_reason",
    ]


def _tab_row(workbook_name: str, tab) -> dict:
    top_scores = sorted(
        tab.signal_scores.items(), key=lambda kv: kv[1], reverse=True
    )[:3] if tab.signal_scores else []
    score_str = " | ".join(f"{k}:{v:.2f}" for k, v in top_scores)
    return {
        "workbook_name":       workbook_name,
        "tab_name":            tab.tab_name,
        "tab_type":            tab.tab_type.value,
        "tab_type_label":      _DATA_TYPE_LABELS.get(tab.tab_type, tab.tab_type.value),
        "row_count":           tab.row_count,
        "col_count":           tab.col_count,
        "formula_cells":       tab.formula_cell_count,
        "formula_density_pct": round(tab.formula_density * 100, 1),
        "inter_sheet_refs":    tab.inter_sheet_ref_count,
        "referenced_sheets":   ", ".join(tab.referenced_sheets),
        "top_signal_scores":   score_str,
        "classification_reason": tab.classification_reason,
    }
