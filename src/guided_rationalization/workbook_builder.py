"""Assemble the future-state rationalized workbook.

Output sheet layout (numbered tabs):
  01_Master_Source_Data          — consolidated rows from all configured source tabs
  02_KPI_Summary_<workbook>      — KPI formula audit per configured KPI tab
  03_Source_Mapping              — source column → canonical mapping
  04_Data_Dictionary             — one row per canonical column with profile
  05_Reconciliation              — row-count and column-count reconciliation
  06_Issues_Log                  — duplicate-column issues and warnings
  07_Duplicate_Column_Analysis   — per-duplicate detail (workbook / tab / column / positions)
  08_Workbook_Source_Analysis    — per-workbook column uniqueness summary
  09_Documentation               — run parameters, lineage, BAU update instructions

The workbook is assembled only after a duplicate-column pre-flight check.
If duplicates are detected the workbook is still produced, but:
  - sheet 06_Issues_Log lists every duplicate finding with severity HIGH
  - sheet 07_Duplicate_Column_Analysis provides actionable detail
  - build_master_source_df raises DuplicateColumnError so the caller can
    surface the issue before attempting pd.concat
"""

from __future__ import annotations

import io
import logging
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import NamedTuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.guided_rationalization.config import RationalizationConfig
from src.guided_rationalization.source_analyzer import SourceAnalysisResult
from src.guided_rationalization.kpi_analyzer import KpiAnalysisResult

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Styling constants
# ---------------------------------------------------------------------------

_HEADER_FILL      = PatternFill("solid", fgColor="1F4E79")   # dark blue
_HEADER_FONT      = Font(color="FFFFFF", bold=True)
_ALT_FILL         = PatternFill("solid", fgColor="D6E4F0")   # light blue
_WARN_FILL        = PatternFill("solid", fgColor="FFE699")   # amber
_ERROR_FILL       = PatternFill("solid", fgColor="FF7C80")   # red
_SECTION_FONT     = Font(bold=True, size=11)


# ---------------------------------------------------------------------------
# Public exception
# ---------------------------------------------------------------------------

class DuplicateColumnError(ValueError):
    """Raised when a source DataFrame has non-unique column names after renaming.

    Attributes:
        findings: List of :class:`DuplicateFinding` objects describing each
            duplicate cluster found across all configured source tabs.
        diagnostic_bytes: The diagnostic .xlsx workbook bytes produced before
            the error was raised (sheets 06–09 are populated; 01 is empty).
            May be ``None`` if the workbook could not be written.
    """

    def __init__(
        self,
        findings: list[DuplicateFinding],
        diagnostic_bytes: bytes | None = None,
    ) -> None:
        self.findings = findings
        self.diagnostic_bytes = diagnostic_bytes
        summary_lines = []
        for f in findings:
            summary_lines.append(
                f"  {f.workbook} / {f.source_tab}: column '{f.column_name}' "
                f"appears {f.duplicate_count} times at positions {f.column_positions} "
                f"(cause: {f.likely_cause})"
            )
        super().__init__(
            "Duplicate column names detected in source data — pd.concat cannot proceed.\n"
            + "\n".join(summary_lines)
            + "\n\nSee sheet '07_Duplicate_Column_Analysis' in the output workbook for details."
        )


# ---------------------------------------------------------------------------
# Duplicate-column data model
# ---------------------------------------------------------------------------

@dataclass
class DuplicateFinding:
    """One duplicate column occurrence within a single source frame."""

    workbook: str
    source_tab: str
    column_name: str          # the name after renaming (canonical or original)
    duplicate_count: int      # how many times this name appears
    column_positions: list[int]  # 0-based positions in the DataFrame
    original_names: list[str]    # original column names before renaming
    likely_cause: str            # human-readable cause classification


@dataclass
class SourceFrameProfile:
    """Column profile of one source frame before concat."""

    workbook: str
    source_tab: str
    shape: tuple[int, int]
    all_columns: list[str]        # after renaming / excluding
    duplicate_findings: list[DuplicateFinding] = field(default_factory=list)

    @property
    def total_columns(self) -> int:
        return len(self.all_columns)

    @property
    def unique_column_count(self) -> int:
        return len(set(self.all_columns))

    @property
    def duplicate_column_count(self) -> int:
        return self.total_columns - self.unique_column_count


# ---------------------------------------------------------------------------
# Duplicate-column detection
# ---------------------------------------------------------------------------

def inspect_frames_for_duplicates(
    frames_with_meta: list[tuple[pd.DataFrame, str, str, dict[str, str]]],
) -> list[SourceFrameProfile]:
    """Inspect every source frame and return per-frame profiles with duplicate findings.

    Args:
        frames_with_meta: Each element is
            ``(frame, workbook_name, source_tab, rename_map)`` where *rename_map*
            maps original column name → canonical name (before apply).

    Returns:
        One :class:`SourceFrameProfile` per frame.
    """
    profiles: list[SourceFrameProfile] = []

    for frame, wb_name, tab_name, rename_map in frames_with_meta:
        col_list = list(frame.columns)
        counts = Counter(col_list)
        findings: list[DuplicateFinding] = []

        for col_name, count in counts.items():
            if count < 2:
                continue
            positions = [i for i, c in enumerate(col_list) if c == col_name]

            # Identify original names that mapped to this canonical
            originals = [
                orig for orig, canon in rename_map.items() if canon == col_name
            ]
            if not originals:
                # Column was not renamed — original name IS the duplicate name
                originals = [col_name] * count

            # Classify the likely cause
            cause = _classify_duplicate_cause(col_name, originals, frame, positions)

            findings.append(DuplicateFinding(
                workbook=wb_name,
                source_tab=tab_name,
                column_name=col_name,
                duplicate_count=count,
                column_positions=positions,
                original_names=originals,
                likely_cause=cause,
            ))
            logger.warning(
                "DUPLICATE COLUMN: '%s' / '%s': column '%s' appears %d times at "
                "positions %s — original names: %s — likely cause: %s",
                wb_name, tab_name, col_name, count, positions, originals, cause,
            )

        profile = SourceFrameProfile(
            workbook=wb_name,
            source_tab=tab_name,
            shape=frame.shape,
            all_columns=col_list,
            duplicate_findings=findings,
        )
        profiles.append(profile)
        logger.info(
            "Frame profile: %s / %s — shape %s — %d total cols, %d unique, %d duplicate groups",
            wb_name, tab_name, frame.shape,
            profile.total_columns, profile.unique_column_count,
            len(findings),
        )

    return profiles


def _classify_duplicate_cause(
    canonical: str,
    original_names: list[str],
    frame: pd.DataFrame,
    positions: list[int],
) -> str:
    """Return a human-readable classification of why this duplicate exists."""
    # Case 1: The originals are all the same → raw duplicate in source sheet
    if len(set(original_names)) == 1:
        return (
            "repeated_source_field — identical column label appears multiple times "
            "in the source tab (possible repeated monthly/section blocks)"
        )

    # Case 2: Originals differ but all normalise to the same canonical
    # → schema-matching / normalisation collision
    if len(set(original_names)) > 1:
        return (
            f"normalisation_collision — distinct original names {original_names} "
            f"all normalise to the same canonical '{canonical}' "
            "(e.g. 'Policy Number' and 'Policy_Number' → 'policy_number')"
        )

    # Case 3: Only one original — canonical collision from fuzzy merge
    return (
        "fuzzy_merge_collision — two source columns were merged by schema matching "
        f"to the same canonical '{canonical}'"
    )


# ---------------------------------------------------------------------------
# Diagnostic DataFrames
# ---------------------------------------------------------------------------

def build_duplicate_column_analysis_df(profiles: list[SourceFrameProfile]) -> pd.DataFrame:
    """Build the Duplicate_Column_Analysis report DataFrame."""
    rows = []
    for profile in profiles:
        for finding in profile.duplicate_findings:
            rows.append({
                "Workbook":           finding.workbook,
                "Source Tab":         finding.source_tab,
                "Duplicate Column Name": finding.column_name,
                "Duplicate Count":    finding.duplicate_count,
                "Column Position(s)": ", ".join(str(p) for p in finding.column_positions),
                "Original Name(s)":   " | ".join(finding.original_names),
                "Likely Cause":       finding.likely_cause,
                "Recommended Action": _recommend_action(finding),
            })
    if not rows:
        return pd.DataFrame([{
            "Workbook": "(none)",
            "Source Tab": "(none)",
            "Duplicate Column Name": "(none)",
            "Duplicate Count": 0,
            "Column Position(s)": "",
            "Original Name(s)": "",
            "Likely Cause": "No duplicate columns detected",
            "Recommended Action": "",
        }])
    return pd.DataFrame(rows)


def build_workbook_source_analysis_df(profiles: list[SourceFrameProfile]) -> pd.DataFrame:
    """Build the Workbook_Source_Analysis summary DataFrame."""
    rows = []
    for profile in profiles:
        rows.append({
            "Workbook":              profile.workbook,
            "Source Tab":            profile.source_tab,
            "DataFrame Shape":       f"{profile.shape[0]} rows × {profile.shape[1]} cols",
            "Total Columns":         profile.total_columns,
            "Unique Columns":        profile.unique_column_count,
            "Duplicate Columns Count": profile.duplicate_column_count,
            "Has Duplicates":        "YES" if profile.duplicate_column_count > 0 else "no",
            "All Column Names":      " | ".join(profile.all_columns),
        })
    return pd.DataFrame(rows)


def _recommend_action(finding: DuplicateFinding) -> str:
    if "repeated_source_field" in finding.likely_cause:
        return (
            "Inspect the source tab. If these are genuinely the same metric repeated "
            "across sections (e.g. monthly columns), suffix them manually before upload "
            "(Jan_Revenue, Feb_Revenue). If they are truly redundant, remove the extra column."
        )
    if "normalisation_collision" in finding.likely_cause:
        return (
            "Two differently-labelled source columns map to the same standardised name. "
            "Determine whether they hold the same business data or different metrics. "
            "If different: rename one source column before upload. "
            "If same: remove the duplicate column from the source tab."
        )
    return (
        "Review the fuzzy schema-matching result. Disable 'merge high-confidence' "
        "or adjust the matching threshold to prevent these two columns being merged."
    )


# ---------------------------------------------------------------------------
# Issues log builder
# ---------------------------------------------------------------------------

def build_issues_log_df(profiles: list[SourceFrameProfile]) -> pd.DataFrame:
    rows = []
    for profile in profiles:
        for finding in profile.duplicate_findings:
            rows.append({
                "Severity":     "HIGH",
                "Issue Type":   "duplicate_column",
                "Workbook":     finding.workbook,
                "Source Tab":   finding.source_tab,
                "Detail":       (
                    f"Column '{finding.column_name}' appears {finding.duplicate_count} "
                    f"times at positions {finding.column_positions}. "
                    f"Original names: {finding.original_names}. "
                    f"Cause: {finding.likely_cause}"
                ),
                "Action Required": _recommend_action(finding),
            })
    if not rows:
        rows.append({
            "Severity": "INFO",
            "Issue Type": "none",
            "Workbook": "",
            "Source Tab": "",
            "Detail": "No issues detected",
            "Action Required": "",
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Data Dictionary builder
# ---------------------------------------------------------------------------

def build_data_dictionary_df(
    source_result: SourceAnalysisResult,
    kpi_result: KpiAnalysisResult,
) -> pd.DataFrame:
    rows = []
    all_canonicals = set(source_result.column_mapping.values())
    for canonical in sorted(all_canonicals):
        profile = next(
            (p for p in source_result.column_profiles if p.canonical_name == canonical), None
        )
        kpis_using = [
            d.kpi_label for d in kpi_result.dependencies
            if canonical in d.canonical_source_columns
        ]
        rows.append({
            "Canonical Column":       canonical,
            "Match Class":            profile.match_class if profile else "unknown",
            "Present In Workbooks":   profile.present_in_count if profile else 0,
            "Source Workbooks":       ", ".join(profile.present_in_workbooks) if profile else "",
            "Data Types":             " | ".join(
                f"{wb}:{dt}" for wb, dt in (profile.data_types.items() if profile else {}.items())
            ),
            "KPI Referenced":         "yes" if (profile and profile.is_kpi_referenced) else "no",
            "Used By KPIs":           ", ".join(kpis_using),
            "Similar To":             ", ".join(profile.similar_to) if profile else "",
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=[
        "Canonical Column", "Match Class", "Present In Workbooks", "Source Workbooks",
        "Data Types", "KPI Referenced", "Used By KPIs", "Similar To",
    ])


# ---------------------------------------------------------------------------
# Reconciliation builder
# ---------------------------------------------------------------------------

def build_reconciliation_df(
    bundles: list,
    config: RationalizationConfig,
    master_df: pd.DataFrame,
) -> pd.DataFrame:
    rows = []
    total_source_rows = 0
    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        df = bundle.sheets.get(wb_cfg.source_tab)
        if df is None:
            continue
        source_rows = len(df)
        total_source_rows += source_rows
        rows.append({
            "Check":        "source_row_count",
            "Workbook":     bundle.file_name,
            "Source Tab":   wb_cfg.source_tab,
            "Source Rows":  source_rows,
            "Master Rows":  len(master_df[master_df["Source_Workbook"] == bundle.file_name])
            if "Source_Workbook" in master_df.columns else "N/A",
            "Status":       "",
        })
    # Populate status
    for row in rows:
        if row["Master Rows"] == row["Source Rows"]:
            row["Status"] = "PASS — row counts match"
        elif row["Master Rows"] == "N/A":
            row["Status"] = "SKIP — master not built"
        else:
            row["Status"] = (
                f"WARN — source has {row['Source Rows']} rows; "
                f"master has {row['Master Rows']}"
            )
    rows.append({
        "Check": "total_source_rows",
        "Workbook": "(all)",
        "Source Tab": "",
        "Source Rows": total_source_rows,
        "Master Rows": len(master_df) if not master_df.empty else 0,
        "Status": (
            "PASS" if len(master_df) == total_source_rows
            else f"WARN — total source {total_source_rows} vs master {len(master_df)}"
        ),
    })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Master Source Data builder — with pre-flight duplicate check
# ---------------------------------------------------------------------------

def build_master_source_df(
    bundles: list,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult,
) -> tuple[pd.DataFrame, list[SourceFrameProfile]]:
    """Consolidate rows from all configured source tabs into a single DataFrame.

    Adds three lineage columns: Source_Workbook, Source_Sheet, LOB_Identifier.
    Applies the canonical column mapping from *source_result*.

    Returns:
        ``(master_df, profiles)`` where *profiles* contains per-frame
        duplicate analysis. If duplicates are detected, *master_df* will be
        a partial result and a :class:`DuplicateColumnError` is raised carrying
        the full *profiles* list so the caller can generate the diagnostic
        sheets before surfacing the error to the user.

    Raises:
        DuplicateColumnError: If any frame contains non-unique column names
            after applying the canonical mapping.
    """
    excl_set = set(source_result.excluded_columns)
    frames: list[pd.DataFrame] = []
    frames_with_meta: list[tuple[pd.DataFrame, str, str, dict[str, str]]] = []

    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        df = bundle.sheets.get(wb_cfg.source_tab)
        if df is None:
            logger.warning(
                "Source tab '%s' missing from '%s'", wb_cfg.source_tab, bundle.file_name
            )
            continue

        # Build rename map
        rename_map: dict[str, str] = {}
        drop_cols: list[str] = []
        for col in df.columns:
            col_str = str(col)
            key = (bundle.file_name, wb_cfg.source_tab, col_str)
            if key in excl_set:
                drop_cols.append(col_str)
            elif key in source_result.column_mapping:
                rename_map[col_str] = source_result.column_mapping[key]

        frame = df.copy()
        if drop_cols:
            frame = frame.drop(
                columns=[c for c in drop_cols if c in frame.columns], errors="ignore"
            )
        frame = frame.rename(columns=rename_map)

        # Add lineage columns
        frame.insert(0, "LOB_Identifier", wb_cfg.lob_identifier or wb_cfg.workbook_name)
        frame.insert(0, "Source_Sheet", wb_cfg.source_tab)
        frame.insert(0, "Source_Workbook", bundle.file_name)

        frames_with_meta.append((frame, bundle.file_name, wb_cfg.source_tab, rename_map))
        frames.append(frame)

    # --- Pre-flight duplicate-column check ----------------------------------
    profiles = inspect_frames_for_duplicates(frames_with_meta)
    all_findings = [f for p in profiles for f in p.duplicate_findings]

    if not frames:
        empty = pd.DataFrame(columns=["Source_Workbook", "Source_Sheet", "LOB_Identifier"])
        return empty, profiles

    if all_findings:
        # Attempt a best-effort concat for the diagnostic workbook, but also
        # raise so the caller knows the result is unreliable.
        try:
            master = pd.concat(frames, axis=0, ignore_index=True, sort=False)
        except Exception:
            master = pd.DataFrame(columns=["Source_Workbook", "Source_Sheet", "LOB_Identifier"])
        raise DuplicateColumnError(all_findings)

    master = pd.concat(frames, axis=0, ignore_index=True, sort=False)
    return master, profiles


# ---------------------------------------------------------------------------
# KPI audit tab builder
# ---------------------------------------------------------------------------

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
            else f"References configured source tab — update column positions after merge"
        )
        rows.append({
            "kpi_label":              dep.kpi_label,
            "cell_address":           dep.cell_address,
            "original_formula":       dep.formula,
            "aggregate_function":     dep.aggregate_function,
            "source_columns_used":    ", ".join(dep.canonical_source_columns),
            "refs_source_tab_only":   dep.refs_source_tab_only,
            "suggested_formula_note": note,
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Workbook assembly
# ---------------------------------------------------------------------------

def build_rationalized_workbook_bytes(
    bundles: list,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult,
    kpi_result: KpiAnalysisResult,
) -> bytes:
    """Assemble the full future-state workbook and return as bytes.

    If duplicate source columns are detected the workbook is still produced
    (so diagnostic sheets reach the user), but the function re-raises
    :class:`DuplicateColumnError` after saving so the caller can display the
    error message alongside the download.
    """
    wb = Workbook()
    wb.remove(wb.active)

    duplicate_error: DuplicateColumnError | None = None
    master_df = pd.DataFrame(columns=["Source_Workbook", "Source_Sheet", "LOB_Identifier"])
    profiles: list[SourceFrameProfile] = []

    # 1. Master Source Data (may raise DuplicateColumnError)
    try:
        master_df, profiles = build_master_source_df(bundles, config, source_result)
    except DuplicateColumnError as exc:
        duplicate_error = exc
        profiles = exc.findings  # type: ignore[assignment]
        # Rebuild profiles properly — exc.findings are DuplicateFinding objects
        # We need SourceFrameProfile objects; re-run inspect to get them
        profiles = _rebuild_profiles_from_findings(bundles, config, source_result)

    ws_master = wb.create_sheet("01_Master_Source_Data")
    _write_df_to_sheet(ws_master, master_df, title="Master Source Data")
    if duplicate_error:
        _write_warning_banner(ws_master, "⚠ Master data is incomplete — duplicate columns detected. See 07_Duplicate_Column_Analysis.")

    # 2. KPI audit tabs — numbered from 02
    kpi_tab_counter = 2
    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        for kpi_tab in wb_cfg.kpi_tabs:
            audit_df = build_kpi_audit_df(bundle, kpi_tab, kpi_result, config.future_source_tab_name)
            tab_name = f"{kpi_tab_counter:02d}_KPI_Summary_{config.kpi_tab_name_for(bundle.file_name, kpi_tab)}"[:31]
            ws_kpi = wb.create_sheet(tab_name)
            _write_df_to_sheet(ws_kpi, audit_df, title=f"KPI Audit: {kpi_tab}")
            kpi_tab_counter += 1

    # 3. Source Mapping
    ws_map = wb.create_sheet("03_Source_Mapping")
    mapping_df = source_result.to_mapping_dataframe()
    _write_df_to_sheet(ws_map, mapping_df, title="Source Column → Canonical Mapping")

    # 4. Data Dictionary
    ws_dd = wb.create_sheet("04_Data_Dictionary")
    dd_df = build_data_dictionary_df(source_result, kpi_result)
    _write_df_to_sheet(ws_dd, dd_df, title="Data Dictionary")

    # 5. Reconciliation
    ws_rec = wb.create_sheet("05_Reconciliation")
    rec_df = build_reconciliation_df(bundles, config, master_df)
    _write_df_to_sheet(ws_rec, rec_df, title="Row-Count Reconciliation")

    # 6. Issues Log
    ws_issues = wb.create_sheet("06_Issues_Log")
    issues_df = build_issues_log_df(profiles)
    _write_df_to_sheet(ws_issues, issues_df, title="Issues Log")
    _apply_severity_colours(ws_issues, issues_df)

    # 7. Duplicate Column Analysis
    ws_dup = wb.create_sheet("07_Duplicate_Column_Analysis")
    dup_df = build_duplicate_column_analysis_df(profiles)
    _write_df_to_sheet(ws_dup, dup_df, title="Duplicate Column Analysis")

    # 8. Workbook Source Analysis
    ws_src = wb.create_sheet("08_Workbook_Source_Analysis")
    src_df = build_workbook_source_analysis_df(profiles)
    _write_df_to_sheet(ws_src, src_df, title="Workbook Source Analysis")

    # 9. Documentation
    ws_doc = wb.create_sheet("09_Documentation")
    _write_documentation(ws_doc, config, source_result, kpi_result, master_df, profiles)

    buf = io.BytesIO()
    wb.save(buf)
    xlsx_bytes = buf.getvalue()

    if duplicate_error:
        duplicate_error.diagnostic_bytes = xlsx_bytes
        raise duplicate_error

    return xlsx_bytes


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


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _rebuild_profiles_from_findings(
    bundles: list,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult,
) -> list[SourceFrameProfile]:
    """Re-run frame inspection to get proper SourceFrameProfile objects."""
    excl_set = set(source_result.excluded_columns)
    frames_with_meta = []
    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        df = bundle.sheets.get(wb_cfg.source_tab)
        if df is None:
            continue
        rename_map: dict[str, str] = {}
        drop_cols: list[str] = []
        for col in df.columns:
            col_str = str(col)
            key = (bundle.file_name, wb_cfg.source_tab, col_str)
            if key in excl_set:
                drop_cols.append(col_str)
            elif key in source_result.column_mapping:
                rename_map[col_str] = source_result.column_mapping[key]
        frame = df.copy()
        if drop_cols:
            frame = frame.drop(
                columns=[c for c in drop_cols if c in frame.columns], errors="ignore"
            )
        frame = frame.rename(columns=rename_map)
        frame.insert(0, "LOB_Identifier", wb_cfg.lob_identifier or wb_cfg.workbook_name)
        frame.insert(0, "Source_Sheet", wb_cfg.source_tab)
        frame.insert(0, "Source_Workbook", bundle.file_name)
        frames_with_meta.append((frame, bundle.file_name, wb_cfg.source_tab, rename_map))
    return inspect_frames_for_duplicates(frames_with_meta)


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

    for col_idx, col_name in enumerate(df.columns, start=1):
        try:
            max_len = max(
                len(str(col_name)),
                df[df.columns[col_idx - 1]].astype(str).str.len().max()
                if not df.empty else 0,
            )
        except Exception:
            max_len = len(str(col_name))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)


def _write_warning_banner(ws, message: str) -> None:
    """Insert a warning row at the top of an existing worksheet."""
    ws.insert_rows(1)
    cell = ws.cell(row=1, column=1, value=message)
    cell.fill = _WARN_FILL
    cell.font = Font(bold=True, color="7B3F00")


def _apply_severity_colours(ws, issues_df: pd.DataFrame) -> None:
    """Colour Issues Log rows by severity."""
    if "Severity" not in issues_df.columns:
        return
    severity_col = list(issues_df.columns).index("Severity") + 1
    # Row 1 is title, row 2 is header, data starts at row 3
    for row_idx, severity in enumerate(issues_df["Severity"], start=3):
        if severity == "HIGH":
            fill = _ERROR_FILL
        elif severity == "WARN":
            fill = _WARN_FILL
        else:
            fill = None
        if fill:
            for col_idx in range(1, len(issues_df.columns) + 1):
                ws.cell(row=row_idx, column=col_idx).fill = fill


def _write_documentation(
    ws,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult,
    kpi_result: KpiAnalysisResult,
    master_df: pd.DataFrame,
    profiles: list[SourceFrameProfile],
) -> None:
    """Populate the Documentation worksheet."""
    total_duplicates = sum(len(p.duplicate_findings) for p in profiles)

    rows: list[tuple[str, str]] = [
        ("GUIDED WORKBOOK RATIONALIZATION — DOCUMENTATION", ""),
        ("", ""),
        ("1. Run Configuration", ""),
        ("Future Source Tab Name",          config.future_source_tab_name),
        ("Matching Threshold",              str(config.matching_threshold)),
        ("Merge High-Confidence Matches",   str(config.merge_high_confidence)),
        ("Remove Unused Columns",           str(config.remove_unused_columns)),
        ("", ""),
        ("2. Source Workbooks Analyzed", ""),
    ]
    for wb_cfg in config.workbook_configs:
        rows += [
            (f"  Workbook",        wb_cfg.workbook_name),
            (f"  Source Tab",      wb_cfg.source_tab),
            (f"  KPI Tabs",        ", ".join(wb_cfg.kpi_tabs) or "(none)"),
            (f"  LOB Identifier",  wb_cfg.lob_identifier or "(none)"),
            ("", ""),
        ]
    rows += [
        ("3. Source-to-Target Mappings", ""),
        ("Common columns (auto-merged)",   str(len(source_result.common_columns))),
        ("Similar columns (documented)",   str(len(source_result.similar_columns))),
        ("Unique columns (kept as-is)",    str(len(source_result.unique_columns))),
        ("", ""),
        ("4. Column Retention", ""),
        ("Columns retained",    str(len(set(source_result.column_mapping.values())))),
        ("Columns excluded",    str(len(source_result.excluded_columns))),
        ("Exclusion reason",    "Not referenced by any KPI formula" if config.remove_unused_columns else "N/A — remove_unused_columns=False"),
        ("", ""),
        ("5. KPI Dependencies", ""),
        ("KPI formula cells analyzed",              str(len(kpi_result.dependencies))),
        ("Source columns referenced by ≥1 KPI",    str(len(kpi_result.referenced_canonicals))),
        ("Source columns NOT referenced by any KPI", str(len(kpi_result.unreferenced_canonicals))),
        ("", ""),
        ("6. Duplicate Column Findings", ""),
        ("Total duplicate column groups found",     str(total_duplicates)),
        ("Status", "CLEAN — no duplicates detected" if total_duplicates == 0
         else f"ACTION REQUIRED — {total_duplicates} duplicate group(s) found. See sheet 07_Duplicate_Column_Analysis."),
        ("", ""),
        ("7. Master Source Data", ""),
        ("Total rows in master",   str(len(master_df))),
        ("Total columns (incl. lineage)", str(len(master_df.columns))),
        ("Lineage columns added",  "Source_Workbook, Source_Sheet, LOB_Identifier"),
        ("", ""),
        ("8. Formula Transformations", ""),
        ("KPI formula update required", "YES — after adopting the future-state workbook, update"),
        ("",                       f"all formula cross-sheet references to point to"),
        ("",                       f"the '{config.future_source_tab_name}' tab."),
        ("",                       "Column positions will change after consolidation."),
        ("", ""),
        ("9. Reconciliation", ""),
        ("See sheet", "05_Reconciliation for row-count verification per workbook"),
        ("", ""),
        ("10. BAU Update Instructions", ""),
        ("Step 1", "Copy new source data rows into the Master_Source_Data tab."),
        ("Step 2", "Ensure Source_Workbook, Source_Sheet, and LOB_Identifier are populated."),
        ("Step 3", "Update KPI formula range references to include new rows."),
        ("Step 4", "Validate KPI outputs against the original workbook totals."),
        ("Step 5", "Archive the original workbooks — do not delete until reconciled."),
    ]

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 55
    for r_idx, (label, value) in enumerate(rows, start=1):
        cell_a = ws.cell(row=r_idx, column=1, value=label)
        ws.cell(row=r_idx, column=2, value=value)
        if r_idx == 1:
            cell_a.font = Font(bold=True, size=13)
        elif label and label[0].isdigit() and ". " in label:
            cell_a.font = _SECTION_FONT
        elif label == "Status" and "ACTION REQUIRED" in str(value):
            cell_a.fill = _ERROR_FILL
            ws.cell(row=r_idx, column=2).fill = _ERROR_FILL
