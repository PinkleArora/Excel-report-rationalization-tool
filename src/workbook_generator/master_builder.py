"""Orchestrate the full pipeline to produce the BAU master workbook.

Entry point::

    from src.workbook_generator.master_builder import build_master_workbook
    path = build_master_workbook(bundles, output_path="output/master_workbook.xlsx")
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pandas as pd

from src.consolidation.consolidator import (
    LINEAGE_COL_SHEET,
    LINEAGE_COL_WORKBOOK,
    build_source_mapping_df,
    consolidate,
)
from src.documentation.generator import (
    DataDictionaryEntry,
    build_data_dictionary,
    data_dictionary_to_df,
)
from src.ingestion.loader import WorkbookBundle
from src.profiling.metadata import WorkbookMetadata
from src.profiling.profiler import profile_many, profile_workbook
from src.rationalization.rationalizer import (
    ReportRecommendation,
    assess_redundancy,
    build_inventory,
)
from src.reconciliation.reconciler import (
    ReconciliationResult,
    reconcile,
    reconciliation_summary,
)
from src.schema_matching.matcher import ColumnMatch, match_all_bundles
from src.workbook_generator.generator import SheetSpec, generate_workbook

logger = logging.getLogger(__name__)

DEFAULT_OUTPUT_PATH = Path("output") / "master_workbook.xlsx"

# Sheet names — prefixed for natural sort order in Excel
SHEET_MASTER_DATA = "01_Master_Data"
SHEET_KPI_SUMMARY = "02_KPI_Summary"
SHEET_SOURCE_MAPPING = "03_Source_Mapping"
SHEET_DATA_DICTIONARY = "04_Data_Dictionary"
SHEET_RECONCILIATION = "05_Reconciliation"
SHEET_ISSUES_LOG = "06_Issues_Log"
SHEET_RATIONALIZATION = "07_Rationalization_Report"
SHEET_SOP = "08_Update_SOP"

ALL_SHEET_NAMES = [
    SHEET_MASTER_DATA,
    SHEET_KPI_SUMMARY,
    SHEET_SOURCE_MAPPING,
    SHEET_DATA_DICTIONARY,
    SHEET_RECONCILIATION,
    SHEET_ISSUES_LOG,
    SHEET_RATIONALIZATION,
    SHEET_SOP,
]


@dataclass
class MasterWorkbookContext:
    """Carries all intermediate artefacts produced during the master build."""

    bundles: list[WorkbookBundle]
    workbook_metas: list[WorkbookMetadata] = field(default_factory=list)
    column_matches: list[ColumnMatch] = field(default_factory=list)
    master_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    source_mapping_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    data_dict_entries: list[DataDictionaryEntry] = field(default_factory=list)
    reconciliation_results: list[ReconciliationResult] = field(default_factory=list)
    rationalization_recs: list[ReportRecommendation] = field(default_factory=list)
    output_path: Path = field(default=DEFAULT_OUTPUT_PATH)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_master_workbook(
    bundles: list[WorkbookBundle],
    output_path: str | Path | None = None,
    matching_threshold: float = 80.0,
    reconciliation_tolerance: float = 0.01,
) -> Path:
    """Run the full pipeline and write ``master_workbook.xlsx``.

    Pipeline stages:
    1. Profile all bundles
    2. Schema-match columns across all workbook pairs
    3. Consolidate into a master DataFrame (with lineage columns)
    4. Build source mapping table
    5. Build data dictionary
    6. Run reconciliation checks on numeric columns
    7. Assess rationalization recommendations
    8. Assemble 8 sheets and export

    Args:
        bundles: Loaded :class:`~src.ingestion.loader.WorkbookBundle` objects.
        output_path: Destination path.  Defaults to ``output/master_workbook.xlsx``.
        matching_threshold: Minimum fuzzy-match score (0–100).
        reconciliation_tolerance: Acceptable relative delta fraction (default 0.01 = 1 %).

    Returns:
        Absolute :class:`pathlib.Path` of the written workbook.

    Raises:
        ValueError: If *bundles* is empty.
    """
    if not bundles:
        raise ValueError("build_master_workbook: at least one WorkbookBundle is required.")

    resolved_path = Path(output_path) if output_path else DEFAULT_OUTPUT_PATH
    ctx = MasterWorkbookContext(bundles=bundles, output_path=resolved_path)

    logger.info("=== Master Workbook Build started ===")
    logger.info("Input bundles: %s", [b.file_name for b in bundles])

    _stage_profile(ctx)
    _stage_schema_matching(ctx, matching_threshold)
    _stage_consolidate(ctx)
    _stage_source_mapping(ctx)
    _stage_data_dictionary(ctx)
    _stage_reconciliation(ctx, reconciliation_tolerance)
    _stage_rationalization(ctx)

    path = _stage_export(ctx)
    logger.info("=== Master Workbook Build complete → %s ===", path)
    return path


def build_master_workbook_bytes(
    bundles: list[WorkbookBundle],
    matching_threshold: float = 80.0,
    reconciliation_tolerance: float = 0.01,
) -> bytes:
    """Build the master workbook in-memory and return raw bytes.

    Useful for Streamlit download buttons without touching the filesystem.
    """
    import io
    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp) / "master_workbook.xlsx"
        build_master_workbook(
            bundles,
            output_path=tmp_path,
            matching_threshold=matching_threshold,
            reconciliation_tolerance=reconciliation_tolerance,
        )
        return tmp_path.read_bytes()


# ---------------------------------------------------------------------------
# Pipeline stages
# ---------------------------------------------------------------------------

def _stage_profile(ctx: MasterWorkbookContext) -> None:
    logger.info("Stage 1/7: Profiling %d bundle(s)…", len(ctx.bundles))
    ctx.workbook_metas = profile_many(ctx.bundles)
    logger.info("  → %d workbook profile(s) produced", len(ctx.workbook_metas))


def _stage_schema_matching(ctx: MasterWorkbookContext, threshold: float) -> None:
    logger.info("Stage 2/7: Schema matching (threshold=%.0f)…", threshold)
    ctx.column_matches = match_all_bundles(ctx.bundles, threshold=threshold)
    logger.info("  → %d column match(es) found", len(ctx.column_matches))


def _stage_consolidate(ctx: MasterWorkbookContext) -> None:
    logger.info("Stage 3/7: Consolidating into master DataFrame…")
    ctx.master_df = consolidate(ctx.bundles, ctx.column_matches, strategy="union")
    logger.info(
        "  → Master data: %d rows × %d cols",
        len(ctx.master_df), len(ctx.master_df.columns),
    )


def _stage_source_mapping(ctx: MasterWorkbookContext) -> None:
    logger.info("Stage 4/7: Building source mapping…")
    ctx.source_mapping_df = build_source_mapping_df(ctx.bundles, ctx.column_matches)
    logger.info("  → %d source column mapping(s)", len(ctx.source_mapping_df))


def _stage_data_dictionary(ctx: MasterWorkbookContext) -> None:
    logger.info("Stage 5/7: Building data dictionary…")
    ctx.data_dict_entries = build_data_dictionary(ctx.workbook_metas, ctx.column_matches)
    logger.info("  → %d dictionary entry(ies)", len(ctx.data_dict_entries))


def _stage_reconciliation(ctx: MasterWorkbookContext, tolerance: float) -> None:
    logger.info("Stage 6/7: Running reconciliation checks…")
    numeric_cols = _identify_numeric_columns(ctx.master_df)
    if not numeric_cols:
        logger.info("  → No numeric columns found; skipping reconciliation")
        ctx.reconciliation_results = []
        return

    source_frames = [
        df
        for bundle in ctx.bundles
        for df in bundle.sheets.values()
    ]
    ctx.reconciliation_results = reconcile(
        source_frames=source_frames,
        output_frame=ctx.master_df,
        numeric_columns=numeric_cols,
        tolerance=tolerance,
    )
    n_fail = sum(1 for r in ctx.reconciliation_results if r.status == "fail")
    n_warn = sum(1 for r in ctx.reconciliation_results if r.status == "warn")
    logger.info(
        "  → %d check(s): %d fail, %d warn, %d pass",
        len(ctx.reconciliation_results), n_fail, n_warn,
        len(ctx.reconciliation_results) - n_fail - n_warn,
    )


def _stage_rationalization(ctx: MasterWorkbookContext) -> None:
    logger.info("Stage 7/7: Assessing rationalization…")
    ctx.rationalization_recs = assess_redundancy(ctx.bundles, ctx.column_matches)
    for r in ctx.rationalization_recs:
        logger.info("  %s → %s (overlap %.0f%%)", r.report_name, r.status.value, r.overlap_pct)


def _stage_export(ctx: MasterWorkbookContext) -> Path:
    logger.info("Assembling %d sheet(s)…", len(ALL_SHEET_NAMES))
    specs = _build_sheet_specs(ctx)
    return generate_workbook(specs, output_path=ctx.output_path, include_cover_sheet=False)


# ---------------------------------------------------------------------------
# Sheet builders — one function per sheet
# ---------------------------------------------------------------------------

def _build_sheet_specs(ctx: MasterWorkbookContext) -> list[SheetSpec]:
    return [
        SheetSpec(name=SHEET_MASTER_DATA,     data=_build_master_data(ctx)),
        SheetSpec(name=SHEET_KPI_SUMMARY,     data=_build_kpi_summary(ctx)),
        SheetSpec(name=SHEET_SOURCE_MAPPING,  data=ctx.source_mapping_df),
        SheetSpec(name=SHEET_DATA_DICTIONARY, data=data_dictionary_to_df(ctx.data_dict_entries)),
        SheetSpec(name=SHEET_RECONCILIATION,  data=reconciliation_summary(ctx.reconciliation_results),
                  status_col=_reconciliation_status_col(ctx.reconciliation_results)),
        SheetSpec(name=SHEET_ISSUES_LOG,      data=_build_issues_log(ctx)),
        SheetSpec(name=SHEET_RATIONALIZATION, data=build_inventory(ctx.rationalization_recs)),
        SheetSpec(name=SHEET_SOP,             data=_build_sop()),
    ]


def _build_master_data(ctx: MasterWorkbookContext) -> pd.DataFrame:
    """Return the master DataFrame, ensuring lineage columns come first."""
    df = ctx.master_df
    if df.empty:
        return df
    other_cols = [c for c in df.columns if c not in (LINEAGE_COL_WORKBOOK, LINEAGE_COL_SHEET)]
    ordered_cols = [LINEAGE_COL_WORKBOOK, LINEAGE_COL_SHEET] + other_cols
    return df[[c for c in ordered_cols if c in df.columns]]


def _build_kpi_summary(ctx: MasterWorkbookContext) -> pd.DataFrame:
    """Compute aggregate statistics on each numeric column, broken down by source."""
    df = ctx.master_df
    if df.empty:
        return pd.DataFrame(columns=["source_workbook", "metric", "column", "value"])

    numeric_cols = _identify_numeric_columns(df)
    if not numeric_cols:
        return pd.DataFrame(columns=["source_workbook", "metric", "column", "value"])

    rows: list[dict] = []

    def _add_stats(label: str, frame: pd.DataFrame) -> None:
        for col in numeric_cols:
            if col not in frame.columns:
                continue
            series = pd.to_numeric(frame[col], errors="coerce").dropna()
            if series.empty:
                continue
            rows.append({"source_workbook": label, "column": col, "metric": "count",   "value": int(series.count())})
            rows.append({"source_workbook": label, "column": col, "metric": "sum",     "value": round(float(series.sum()), 4)})
            rows.append({"source_workbook": label, "column": col, "metric": "mean",    "value": round(float(series.mean()), 4)})
            rows.append({"source_workbook": label, "column": col, "metric": "min",     "value": round(float(series.min()), 4)})
            rows.append({"source_workbook": label, "column": col, "metric": "max",     "value": round(float(series.max()), 4)})

    _add_stats("ALL SOURCES", df)

    if LINEAGE_COL_WORKBOOK in df.columns:
        for wb_name, group in df.groupby(LINEAGE_COL_WORKBOOK):
            _add_stats(str(wb_name), group)

    return pd.DataFrame(rows)[["source_workbook", "column", "metric", "value"]]


def _build_issues_log(ctx: MasterWorkbookContext) -> pd.DataFrame:
    """Auto-populate issues from profiling; leave room for manual entries."""
    base_cols = [
        "issue_id", "date_raised", "source_workbook", "source_sheet",
        "column", "issue_description", "severity", "status",
        "owner", "target_resolution_date", "notes",
    ]
    rows: list[dict] = []
    issue_id = 1

    for wm in ctx.workbook_metas:
        for sm in wm.sheets:
            # Formula cells
            if sm.formula_cell_count > 0:
                rows.append({
                    "issue_id": f"ISS-{issue_id:04d}",
                    "date_raised": str(date.today()),
                    "source_workbook": wm.workbook_name,
                    "source_sheet": sm.sheet_name,
                    "column": "",
                    "issue_description": (
                        f"{sm.formula_cell_count} formula cell(s) detected. "
                        "Verify that data was extracted as values, not formulas."
                    ),
                    "severity": "MEDIUM",
                    "status": "OPEN",
                    "owner": "",
                    "target_resolution_date": "",
                    "notes": "",
                })
                issue_id += 1

            for cm in sm.columns:
                # High null rate
                if cm.null_pct > 50 and not cm.is_empty:
                    rows.append({
                        "issue_id": f"ISS-{issue_id:04d}",
                        "date_raised": str(date.today()),
                        "source_workbook": wm.workbook_name,
                        "source_sheet": sm.sheet_name,
                        "column": cm.column_name,
                        "issue_description": (
                            f"High null rate: {cm.null_pct:.1f}% of values are missing."
                        ),
                        "severity": "HIGH",
                        "status": "OPEN",
                        "owner": "",
                        "target_resolution_date": "",
                        "notes": "",
                    })
                    issue_id += 1

                # Empty column
                if cm.is_empty:
                    rows.append({
                        "issue_id": f"ISS-{issue_id:04d}",
                        "date_raised": str(date.today()),
                        "source_workbook": wm.workbook_name,
                        "source_sheet": sm.sheet_name,
                        "column": cm.column_name,
                        "issue_description": "Column contains no data (100% null).",
                        "severity": "LOW",
                        "status": "OPEN",
                        "owner": "",
                        "target_resolution_date": "",
                        "notes": "",
                    })
                    issue_id += 1

    # Reconciliation failures
    for r in ctx.reconciliation_results:
        if r.status in ("fail", "warn"):
            rows.append({
                "issue_id": f"ISS-{issue_id:04d}",
                "date_raised": str(date.today()),
                "source_workbook": "RECONCILIATION",
                "source_sheet": "05_Reconciliation",
                "column": r.column,
                "issue_description": (
                    f"Reconciliation {r.status.upper()}: "
                    f"source={r.source_total:,.2f}, master={r.output_total:,.2f}, "
                    f"delta={r.delta_pct:.2f}% (tolerance={r.tolerance:.2f}%)."
                ),
                "severity": "HIGH" if r.status == "fail" else "MEDIUM",
                "status": "OPEN",
                "owner": "",
                "target_resolution_date": "",
                "notes": "",
            })
            issue_id += 1

    if not rows:
        # Placeholder row so the sheet isn't completely blank
        rows.append({k: "" for k in base_cols})
        rows[0]["issue_id"] = "ISS-0001"
        rows[0]["issue_description"] = "No automated issues detected."
        rows[0]["status"] = "CLOSED"

    return pd.DataFrame(rows)[base_cols]


def _build_sop() -> pd.DataFrame:
    """Return the standard operating procedure template."""
    rows = [
        {
            "step": 1,
            "action": "Collect source reports",
            "detail": "Gather all BAU Excel/CSV reports from report owners.",
            "frequency": "Monthly",
            "owner": "",
            "system": "SharePoint / Email",
            "last_completed": "",
            "notes": "",
        },
        {
            "step": 2,
            "action": "Upload to profiling tool",
            "detail": "Run the Upload & Profile page to generate profiling_summary.xlsx.",
            "frequency": "Monthly",
            "owner": "",
            "system": "Excel Rationalizer — Profile Reports",
            "last_completed": "",
            "notes": "",
        },
        {
            "step": 3,
            "action": "Review schema matches",
            "detail": "Check 03_Source_Mapping for any new or changed column mappings.",
            "frequency": "Monthly",
            "owner": "",
            "system": "Excel Rationalizer — Generate Master Workbook",
            "last_completed": "",
            "notes": "",
        },
        {
            "step": 4,
            "action": "Generate master workbook",
            "detail": "Run Generate Master Workbook to produce master_workbook.xlsx.",
            "frequency": "Monthly",
            "owner": "",
            "system": "Excel Rationalizer — Generate Master Workbook",
            "last_completed": "",
            "notes": "",
        },
        {
            "step": 5,
            "action": "Review reconciliation",
            "detail": "Open 05_Reconciliation. Investigate any FAIL or WARN rows.",
            "frequency": "Monthly",
            "owner": "",
            "system": "master_workbook.xlsx",
            "last_completed": "",
            "notes": "",
        },
        {
            "step": 6,
            "action": "Resolve issues",
            "detail": "Work through 06_Issues_Log. Update status and owner fields.",
            "frequency": "Monthly",
            "owner": "",
            "system": "master_workbook.xlsx",
            "last_completed": "",
            "notes": "",
        },
        {
            "step": 7,
            "action": "Distribute master workbook",
            "detail": "Share approved master_workbook.xlsx with stakeholders.",
            "frequency": "Monthly",
            "owner": "",
            "system": "SharePoint / Email",
            "last_completed": "",
            "notes": "",
        },
        {
            "step": 8,
            "action": "Review rationalization recommendations",
            "detail": "Check 07_Rationalization_Report for retire/merge candidates.",
            "frequency": "Quarterly",
            "owner": "",
            "system": "master_workbook.xlsx",
            "last_completed": "",
            "notes": "",
        },
    ]
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def _identify_numeric_columns(df: pd.DataFrame) -> list[str]:
    """Return canonical column names that contain numeric data, excluding lineage cols."""
    skip = {LINEAGE_COL_WORKBOOK, LINEAGE_COL_SHEET}
    result = []
    for col in df.columns:
        if col in skip:
            continue
        series = pd.to_numeric(df[col], errors="coerce")
        if series.notna().any():
            result.append(col)
    return result


def _reconciliation_status_col(results: list[ReconciliationResult]) -> str | None:
    """Return the Excel column letter of the 'status' column in the reconciliation sheet."""
    if not results:
        return None
    # reconciliation_summary returns: column, source_total, master_total, delta, delta_pct, tolerance_pct, status
    cols = ["column", "source_total", "master_total", "delta", "delta_pct", "tolerance_pct", "status"]
    idx = cols.index("status")
    from openpyxl.utils import get_column_letter
    return get_column_letter(idx + 1)
