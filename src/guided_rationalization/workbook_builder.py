"""Assemble the future-state rationalized workbook.

Produces an .xlsx workbook containing:
  - Master_Source_Data    — consolidated rows from all configured source tabs
  - One KPI tab per configured KPI tab (values + formula audit)
  - Column_Mapping        — canonical mapping report
  - KPI_Dependencies      — which source columns each KPI formula uses
  - Column_Coverage       — which source columns are / are not KPI-referenced
  - Documentation         — run parameters and lineage notes
"""

from __future__ import annotations

import io
import logging
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from src.guided_rationalization.config import RationalizationConfig
from src.guided_rationalization.source_analyzer import SourceAnalysisResult
from src.guided_rationalization.kpi_analyzer import KpiAnalysisResult

logger = logging.getLogger(__name__)

_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_ALT_FILL    = PatternFill("solid", fgColor="D6E4F0")


def _write_df_to_sheet(ws, df: pd.DataFrame, title: str | None = None) -> None:
    """Write a DataFrame to an openpyxl worksheet with styled headers."""
    start_row = 1
    if title:
        ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=12)
        start_row = 2

    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=start_row, column=col_idx, value=str(col_name))
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(wrap_text=True)

    for row_idx, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        fill = _ALT_FILL if (row_idx - start_row) % 2 == 0 else None
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if fill:
                cell.fill = fill

    # Auto-size columns (capped at 60)
    for col_idx, col_name in enumerate(df.columns, start=1):
        max_len = max(
            len(str(col_name)),
            df[df.columns[col_idx - 1]].astype(str).str.len().max() if not df.empty else 0,
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)


def build_master_source_df(
    bundles: list,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult,
) -> pd.DataFrame:
    """Consolidate rows from all configured source tabs into a single DataFrame.

    Adds three lineage columns: Source_Workbook, Source_Sheet, LOB_Identifier.
    Applies the canonical column mapping from *source_result*.
    If *config.remove_unused_columns* is True, drops columns not in the mapping
    (i.e., those in excluded_columns).
    """
    excl_set = set(source_result.excluded_columns)
    frames: list[pd.DataFrame] = []

    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        df = bundle.sheets.get(wb_cfg.source_tab)
        if df is None:
            logger.warning("Source tab '%s' missing from '%s'", wb_cfg.source_tab, bundle.file_name)
            continue

        # Rename source columns to canonical names
        rename_map: dict[str, str] = {}
        drop_cols: list[str] = []
        for col in df.columns:
            col_str = str(col)
            key = (bundle.file_name, wb_cfg.source_tab, col_str)
            if key in excl_set:
                drop_cols.append(col_str)
            elif key in source_result.column_mapping:
                canonical = source_result.column_mapping[key]
                rename_map[col_str] = canonical

        frame = df.copy()
        if drop_cols:
            frame = frame.drop(columns=[c for c in drop_cols if c in frame.columns], errors="ignore")
        frame = frame.rename(columns=rename_map)

        # Add lineage columns
        frame.insert(0, "LOB_Identifier", wb_cfg.lob_identifier or wb_cfg.workbook_name)
        frame.insert(0, "Source_Sheet", wb_cfg.source_tab)
        frame.insert(0, "Source_Workbook", bundle.file_name)

        frames.append(frame)

    if not frames:
        return pd.DataFrame(columns=["Source_Workbook", "Source_Sheet", "LOB_Identifier"])

    # Outer-concat so no data is lost when schemas differ
    master = pd.concat(frames, axis=0, ignore_index=True, sort=False)
    return master


def build_kpi_audit_df(
    bundle,
    kpi_tab: str,
    kpi_result: KpiAnalysisResult,
    future_source_tab: str,
) -> pd.DataFrame:
    """Build a KPI audit table for one KPI tab showing formula metadata."""
    deps = [
        d for d in kpi_result.dependencies
        if d.workbook_name == bundle.file_name and d.kpi_tab == kpi_tab
    ]
    if not deps:
        return pd.DataFrame(columns=[
            "kpi_label", "cell_address", "original_formula",
            "aggregate_function", "source_columns_used",
            "refs_source_tab_only", "suggested_formula_note",
        ])

    rows = []
    for dep in deps:
        note = (
            f"Update cross-sheet references to point to '{future_source_tab}' tab"
            if not dep.refs_source_tab_only
            else f"References '{future_source_tab}' — update column positions after merge"
        )
        rows.append({
            "kpi_label":            dep.kpi_label,
            "cell_address":         dep.cell_address,
            "original_formula":     dep.formula,
            "aggregate_function":   dep.aggregate_function,
            "source_columns_used":  ", ".join(dep.canonical_source_columns),
            "refs_source_tab_only": dep.refs_source_tab_only,
            "suggested_formula_note": note,
        })
    return pd.DataFrame(rows)


def build_rationalized_workbook_bytes(
    bundles: list,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult,
    kpi_result: KpiAnalysisResult,
) -> bytes:
    """Assemble the full future-state workbook and return as bytes."""
    wb = Workbook()
    wb.remove(wb.active)  # remove default empty sheet

    # 1. Master Source Data
    master_df = build_master_source_df(bundles, config, source_result)
    ws_master = wb.create_sheet(config.future_source_tab_name)
    _write_df_to_sheet(ws_master, master_df, title="Master Source Data")

    # 2. KPI audit tabs — one per configured KPI tab
    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        for kpi_tab in wb_cfg.kpi_tabs:
            audit_df = build_kpi_audit_df(bundle, kpi_tab, kpi_result, config.future_source_tab_name)
            tab_name = config.kpi_tab_name_for(bundle.file_name, kpi_tab)
            ws_kpi = wb.create_sheet(tab_name)
            _write_df_to_sheet(ws_kpi, audit_df, title=f"KPI Audit: {kpi_tab}")

    # 3. Column Mapping
    ws_map = wb.create_sheet("Column_Mapping")
    mapping_df = source_result.to_mapping_dataframe()
    _write_df_to_sheet(ws_map, mapping_df, title="Source Column → Canonical Mapping")

    # 4. KPI Dependencies
    ws_dep = wb.create_sheet("KPI_Dependencies")
    dep_df = kpi_result.to_dependency_dataframe()
    _write_df_to_sheet(ws_dep, dep_df, title="KPI Formula Dependencies")

    # 5. Column Coverage
    ws_cov = wb.create_sheet("Column_Coverage")
    all_canonicals: set[str] = {c for c in source_result.column_mapping.values()}
    cov_df = kpi_result.to_coverage_dataframe(all_canonicals)
    _write_df_to_sheet(ws_cov, cov_df, title="Source Column KPI Coverage")

    # 6. Documentation
    ws_doc = wb.create_sheet("Documentation")
    _write_documentation(ws_doc, config, source_result, kpi_result, master_df)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_rationalized_workbook(
    bundles: list,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult,
    kpi_result: KpiAnalysisResult,
    output_path: Path,
) -> Path:
    """Write the future-state workbook to *output_path* and return it."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    data = build_rationalized_workbook_bytes(bundles, config, source_result, kpi_result)
    output_path.write_bytes(data)
    logger.info("Future-state workbook written to %s", output_path)
    return output_path.resolve()


def _write_documentation(ws, config, source_result, kpi_result, master_df) -> None:
    """Populate the Documentation worksheet with run summary and lineage."""
    rows = [
        ("GUIDED WORKBOOK RATIONALIZATION — DOCUMENTATION", ""),
        ("", ""),
        ("Run Configuration", ""),
        ("Future Source Tab Name", config.future_source_tab_name),
        ("Matching Threshold", config.matching_threshold),
        ("Merge High-Confidence Fuzzy Matches", config.merge_high_confidence),
        ("Remove Unused Columns", config.remove_unused_columns),
        ("", ""),
        ("Source Summary", ""),
        ("Total master rows", len(master_df)),
        ("Total canonical columns (incl. lineage)", len(master_df.columns)),
        ("Common columns", len(source_result.common_columns)),
        ("Similar columns", len(source_result.similar_columns)),
        ("Unique columns", len(source_result.unique_columns)),
        ("Excluded columns (unused)", len(source_result.excluded_columns)),
        ("", ""),
        ("KPI Summary", ""),
        ("KPI formula cells analysed", len(kpi_result.dependencies)),
        ("Canonical source columns referenced by ≥1 KPI", len(kpi_result.referenced_canonicals)),
        ("Canonical source columns NOT referenced by any KPI", len(kpi_result.unreferenced_canonicals)),
        ("", ""),
        ("Configured Workbooks", ""),
    ]
    for wb_cfg in config.workbook_configs:
        rows.append((f"  {wb_cfg.workbook_name}", ""))
        rows.append(("    Source tab", wb_cfg.source_tab))
        rows.append(("    KPI tabs", ", ".join(wb_cfg.kpi_tabs)))
        rows.append(("    LOB identifier", wb_cfg.lob_identifier or "(none)"))

    rows += [
        ("", ""),
        ("Lineage Columns added to Master Source Data", ""),
        ("Source_Workbook", "Original workbook file name"),
        ("Source_Sheet", "Original source tab name"),
        ("LOB_Identifier", "Line-of-business label configured by user"),
        ("", ""),
        ("Notes", ""),
        ("KPI formulas reference the original source tab names.", ""),
        ("After adopting the future-state workbook, update formula cross-sheet references", ""),
        (f"to point to the '{config.future_source_tab_name}' tab and adjust column positions.", ""),
    ]

    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["B"].width = 40
    for r_idx, (label, value) in enumerate(rows, start=1):
        cell_a = ws.cell(row=r_idx, column=1, value=label)
        ws.cell(row=r_idx, column=2, value=value)
        if r_idx == 1:
            cell_a.font = Font(bold=True, size=13)
        elif value == "" and label and not label.startswith(" "):
            cell_a.font = Font(bold=True)
