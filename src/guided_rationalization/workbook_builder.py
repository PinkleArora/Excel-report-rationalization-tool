"""Assemble the future-state rationalized workbook.

Output sheet layout (numbered tabs):
  01_Master_Source_Data          — consolidated rows from all configured source tabs
  02_KPI_Summary_<workbook>      — KPI formula audit per configured KPI tab
  03_Source_Mapping              — source column → canonical mapping
  04_Data_Dictionary             — one row per canonical column with profile
  05_Reconciliation              — row-count and column-count reconciliation
  06_Issues_Log                  — warnings and manual-review items
  07_Duplicate_Column_Analysis   — per-duplicate detail with KPI usage and resolution
  08_Workbook_Source_Analysis    — per-workbook column uniqueness summary
  09_Documentation               — run parameters, lineage, BAU update instructions

Duplicate-column handling
-------------------------
Duplicate columns (two source columns that map to the same canonical name) are
evaluated against KPI usage rather than treated as unconditional errors.

  Scenario A — one position used by KPIs, other(s) not
              → keep the used position, remove the unused position(s), continue

  Scenario B — no position used by any KPI
              → remove all positions (none contributes to future-state reporting)

  Scenario C — multiple positions each used by different KPIs
              → flag for manual review; block only this case

  Scenario D — multiple positions used by the same KPI(s) with identical values
              → keep the first position, remove the rest, continue

Workbook generation is blocked only when Scenario C duplicates remain
after applying rules A, B, and D.
"""

from __future__ import annotations

import io
import logging
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

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

_HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
_HEADER_FONT  = Font(color="FFFFFF", bold=True)
_ALT_FILL     = PatternFill("solid", fgColor="D6E4F0")
_WARN_FILL    = PatternFill("solid", fgColor="FFE699")
_ERROR_FILL   = PatternFill("solid", fgColor="FF7C80")
_OK_FILL      = PatternFill("solid", fgColor="C6EFCE")
_SECTION_FONT = Font(bold=True, size=11)


# ---------------------------------------------------------------------------
# Public exception
# ---------------------------------------------------------------------------

class DuplicateColumnError(ValueError):
    """Raised only when Scenario C duplicates remain after automatic resolution.

    Scenario C: two or more positions of the same canonical name are each
    referenced by *different* KPI formulas — the tool cannot determine which
    to keep without human judgement.

    Attributes:
        resolutions: All :class:`DuplicateResolution` objects, including the
            blocking MANUAL_REVIEW entries.
        diagnostic_bytes: The diagnostic .xlsx workbook bytes (always populated
            when raised from :func:`build_rationalized_workbook_bytes`).
    """

    def __init__(
        self,
        resolutions: list[DuplicateResolution],
        diagnostic_bytes: bytes | None = None,
    ) -> None:
        self.resolutions = resolutions
        self.diagnostic_bytes = diagnostic_bytes
        blocking = [r for r in resolutions if r.scenario == "C"]
        lines = []
        for r in blocking:
            lines.append(
                f"  {r.workbook} / {r.source_tab}: column '{r.canonical_name}' "
                f"position {r.original_position} ('{r.original_col_name}') "
                f"used by KPIs: {r.kpis_using}"
            )
        super().__init__(
            "Duplicate columns require manual review — both positions are used by "
            "different KPI formulas and cannot be auto-resolved.\n"
            + "\n".join(lines)
            + "\n\nSee sheet '07_Duplicate_Column_Analysis' for full details."
        )


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass
class DuplicateFinding:
    """One duplicate-column group within a single source frame (pre-resolution)."""

    workbook: str
    source_tab: str
    # Canonical name that multiple original columns map to
    canonical_name: str
    duplicate_count: int
    # 0-based positions in the ORIGINAL source DataFrame (before rename/lineage)
    original_positions: list[int]
    original_col_names: list[str]
    likely_cause: str


@dataclass
class DuplicateResolution:
    """Resolution decision for one column position within a duplicate group."""

    workbook: str
    source_tab: str
    canonical_name: str
    original_position: int       # 0-based index in original source DataFrame
    original_col_name: str
    scenario: str                # A | B | C | D
    action: str                  # KEEP | REMOVE | MANUAL_REVIEW
    kpis_using: list[str]        # KPI cell addresses referencing this position
    reason: str


@dataclass
class SourceFrameProfile:
    """Column profile of one source frame (post-resolution, pre-concat)."""

    workbook: str
    source_tab: str
    shape: tuple[int, int]       # of the RESOLVED frame
    all_columns: list[str]       # column names after resolution + rename
    duplicate_findings: list[DuplicateFinding] = field(default_factory=list)
    resolutions: list[DuplicateResolution] = field(default_factory=list)

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
# Column-letter helper
# ---------------------------------------------------------------------------

def _index_to_col_letter(idx: int) -> str:
    """Convert 0-based column index to Excel column letter (A, B, …, AA, …)."""
    result = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        result = chr(rem + ord("A")) + result
    return result


# ---------------------------------------------------------------------------
# Step 1 — detect duplicates in the original (pre-rename) source frame
# ---------------------------------------------------------------------------

def _detect_raw_duplicates(
    df: pd.DataFrame,
    rename_map: dict[str, str],
    workbook: str,
    source_tab: str,
) -> list[DuplicateFinding]:
    """Find columns in *df* whose canonical names (after applying *rename_map*) collide.

    Works on the ORIGINAL source DataFrame before any renaming, so positions
    correspond directly to Excel column letters used in KPI raw_refs.
    """
    # Build list of (original_position, original_name, canonical_name)
    triples: list[tuple[int, str, str]] = []
    for pos, col in enumerate(df.columns):
        col_str = str(col)
        canonical = rename_map.get(col_str, col_str)
        triples.append((pos, col_str, canonical))

    # Group by canonical name
    from collections import defaultdict
    canonical_to_positions: dict[str, list[tuple[int, str]]] = defaultdict(list)
    for pos, orig, canon in triples:
        canonical_to_positions[canon].append((pos, orig))

    findings: list[DuplicateFinding] = []
    for canon, entries in canonical_to_positions.items():
        if len(entries) < 2:
            continue
        positions = [p for p, _ in entries]
        originals = [o for _, o in entries]
        cause = _classify_cause(canon, originals)
        findings.append(DuplicateFinding(
            workbook=workbook,
            source_tab=source_tab,
            canonical_name=canon,
            duplicate_count=len(entries),
            original_positions=positions,
            original_col_names=originals,
            likely_cause=cause,
        ))
        logger.warning(
            "Duplicate canonical '%s' in '%s'/'%s': positions %s, originals %s",
            canon, workbook, source_tab, positions, originals,
        )
    return findings


def _classify_cause(canonical: str, original_names: list[str]) -> str:
    if len(set(original_names)) == 1:
        return (
            "repeated_source_field — identical label appears multiple times in the "
            "source tab (possible repeated monthly/section blocks)"
        )
    return (
        f"normalisation_collision — distinct original names {original_names} "
        f"all normalise to the same canonical '{canonical}' "
        "(e.g. 'Policy Number' and 'Policy_Number' → 'policy_number')"
    )


# ---------------------------------------------------------------------------
# Step 2 — determine KPI usage per position
# ---------------------------------------------------------------------------

def _kpis_using_position(
    workbook: str,
    source_tab: str,
    original_position: int,
    kpi_result: KpiAnalysisResult,
) -> list[str]:
    """Return KPI cell addresses whose raw formula references include this column position.

    Converts the 0-based *original_position* to an Excel column letter and
    checks it against each dependency's raw_refs list.
    """
    col_letter = _index_to_col_letter(original_position)
    using: list[str] = []
    for dep in kpi_result.dependencies:
        if dep.workbook_name != workbook:
            continue
        for sheet, letter in dep.raw_refs:
            if sheet == source_tab and letter == col_letter:
                using.append(dep.cell_address)
                break  # count cell_address once per dependency
    return using


# ---------------------------------------------------------------------------
# Step 3 — apply resolution scenarios
# ---------------------------------------------------------------------------

def resolve_duplicate_group(
    finding: DuplicateFinding,
    kpi_result: KpiAnalysisResult,
    df: pd.DataFrame,
) -> list[DuplicateResolution]:
    """Apply scenario logic to one duplicate group and return per-position resolutions."""
    per_position: list[tuple[int, str, list[str]]] = []
    for pos, orig in zip(finding.original_positions, finding.original_col_names):
        kpis = _kpis_using_position(
            finding.workbook, finding.source_tab, pos, kpi_result
        )
        per_position.append((pos, orig, kpis))

    used_positions   = [(p, o, k) for p, o, k in per_position if k]
    unused_positions = [(p, o, k) for p, o, k in per_position if not k]

    # Scenario B — no position is used by any KPI
    if not used_positions:
        return [
            DuplicateResolution(
                workbook=finding.workbook,
                source_tab=finding.source_tab,
                canonical_name=finding.canonical_name,
                original_position=pos,
                original_col_name=orig,
                scenario="B",
                action="REMOVE",
                kpis_using=[],
                reason=(
                    "No KPI formula references this column. "
                    "Removed — does not contribute to future-state reporting."
                ),
            )
            for pos, orig, _ in per_position
        ]

    # Scenario A — exactly one position used, rest unused
    if len(used_positions) == 1:
        used_pos, used_orig, used_kpis = used_positions[0]
        resolutions = [
            DuplicateResolution(
                workbook=finding.workbook,
                source_tab=finding.source_tab,
                canonical_name=finding.canonical_name,
                original_position=used_pos,
                original_col_name=used_orig,
                scenario="A",
                action="KEEP",
                kpis_using=used_kpis,
                reason=(
                    f"This position is referenced by KPI(s): {', '.join(used_kpis)}. "
                    "Kept — required for future-state KPI reproduction."
                ),
            )
        ]
        for pos, orig, _ in unused_positions:
            resolutions.append(DuplicateResolution(
                workbook=finding.workbook,
                source_tab=finding.source_tab,
                canonical_name=finding.canonical_name,
                original_position=pos,
                original_col_name=orig,
                scenario="A",
                action="REMOVE",
                kpis_using=[],
                reason=(
                    "Not referenced by any KPI formula. "
                    f"Removed — position {used_pos} ('{used_orig}') is the KPI-referenced column."
                ),
            ))
        return resolutions

    # Multiple positions all used — check Scenario D first (same KPIs, identical values)
    # Scenario D: all used positions reference the same set of KPIs AND values are identical
    all_kpi_sets = [frozenset(k) for _, _, k in used_positions]
    same_kpis = len(set(all_kpi_sets)) == 1

    if same_kpis and len(used_positions) > 1:
        # Check if values are identical across positions
        try:
            value_arrays = [
                df.iloc[:, pos].reset_index(drop=True)
                for pos, _, _ in used_positions
            ]
            identical_values = all(
                value_arrays[0].equals(arr) for arr in value_arrays[1:]
            )
        except Exception:
            identical_values = False

        if identical_values:
            keep_pos, keep_orig, keep_kpis = used_positions[0]
            resolutions = [
                DuplicateResolution(
                    workbook=finding.workbook,
                    source_tab=finding.source_tab,
                    canonical_name=finding.canonical_name,
                    original_position=keep_pos,
                    original_col_name=keep_orig,
                    scenario="D",
                    action="KEEP",
                    kpis_using=keep_kpis,
                    reason=(
                        "Multiple positions reference the same KPI(s) with identical values. "
                        "First position kept; duplicates removed."
                    ),
                )
            ]
            for pos, orig, kpis in used_positions[1:]:
                resolutions.append(DuplicateResolution(
                    workbook=finding.workbook,
                    source_tab=finding.source_tab,
                    canonical_name=finding.canonical_name,
                    original_position=pos,
                    original_col_name=orig,
                    scenario="D",
                    action="REMOVE",
                    kpis_using=kpis,
                    reason=(
                        "Identical values to kept position. "
                        "Removed — redundant for KPI reproduction."
                    ),
                ))
            # Unused positions also get removed
            for pos, orig, _ in unused_positions:
                resolutions.append(DuplicateResolution(
                    workbook=finding.workbook,
                    source_tab=finding.source_tab,
                    canonical_name=finding.canonical_name,
                    original_position=pos,
                    original_col_name=orig,
                    scenario="D",
                    action="REMOVE",
                    kpis_using=[],
                    reason="Not referenced by any KPI. Removed as part of Scenario D resolution.",
                ))
            return resolutions

    # Scenario C — multiple positions used by different KPIs, cannot auto-resolve
    resolutions = []
    for pos, orig, kpis in per_position:
        resolutions.append(DuplicateResolution(
            workbook=finding.workbook,
            source_tab=finding.source_tab,
            canonical_name=finding.canonical_name,
            original_position=pos,
            original_col_name=orig,
            scenario="C",
            action="MANUAL_REVIEW",
            kpis_using=kpis,
            reason=(
                "Multiple positions are each referenced by different KPI formula(s). "
                "Cannot automatically determine which to keep. "
                "Manual review required — rename the source columns to give them "
                "distinct identities before re-running."
            ),
        ))
    return resolutions


# ---------------------------------------------------------------------------
# Diagnostic DataFrames
# ---------------------------------------------------------------------------

def build_duplicate_column_analysis_df(
    profiles: list[SourceFrameProfile],
) -> pd.DataFrame:
    """Build the Duplicate Column Analysis sheet with KPI usage and resolution decisions."""
    rows = []
    for profile in profiles:
        # Index resolutions by (canonical_name, original_position) for lookup
        res_index: dict[tuple[str, int], DuplicateResolution] = {
            (r.canonical_name, r.original_position): r
            for r in profile.resolutions
        }
        for finding in profile.duplicate_findings:
            for pos, orig in zip(finding.original_positions, finding.original_col_names):
                res = res_index.get((finding.canonical_name, pos))
                rows.append({
                    "Workbook":              finding.workbook,
                    "Source Tab":            finding.source_tab,
                    "Duplicate Column Name": finding.canonical_name,
                    "Column Position":       pos,
                    "Original Column Name":  orig,
                    "Column Letter (Excel)": _index_to_col_letter(pos),
                    "KPI(s) Using Column":   ", ".join(res.kpis_using) if res else "",
                    "Usage Count":           len(res.kpis_using) if res else 0,
                    "Scenario":              res.scenario if res else "?",
                    "Recommended Action":    res.action if res else "UNKNOWN",
                    "Reason":                res.reason if res else "",
                    "Likely Cause":          finding.likely_cause,
                })
    if not rows:
        return pd.DataFrame([{
            "Workbook": "(none)",
            "Source Tab": "(none)",
            "Duplicate Column Name": "(none)",
            "Column Position": "",
            "Original Column Name": "",
            "Column Letter (Excel)": "",
            "KPI(s) Using Column": "",
            "Usage Count": 0,
            "Scenario": "",
            "Recommended Action": "No duplicate columns detected",
            "Reason": "",
            "Likely Cause": "",
        }])
    return pd.DataFrame(rows)


def build_workbook_source_analysis_df(profiles: list[SourceFrameProfile]) -> pd.DataFrame:
    rows = []
    for profile in profiles:
        auto_resolved = sum(
            1 for r in profile.resolutions
            if r.scenario in ("A", "B", "D") and r.action == "REMOVE"
        )
        manual_review = sum(1 for r in profile.resolutions if r.scenario == "C")
        rows.append({
            "Workbook":                   profile.workbook,
            "Source Tab":                 profile.source_tab,
            "DataFrame Shape (resolved)": f"{profile.shape[0]} rows × {profile.shape[1]} cols",
            "Total Columns (resolved)":   profile.total_columns,
            "Unique Columns":             profile.unique_column_count,
            "Duplicate Groups Found":     len(profile.duplicate_findings),
            "Auto-Resolved (removed)":    auto_resolved,
            "Manual Review Required":     manual_review,
            "Status":                     (
                "MANUAL REVIEW REQUIRED" if manual_review > 0
                else ("AUTO-RESOLVED" if auto_resolved > 0 else "CLEAN")
            ),
            "All Column Names (resolved)": " | ".join(profile.all_columns),
        })
    return pd.DataFrame(rows)


def build_issues_log_df(profiles: list[SourceFrameProfile]) -> pd.DataFrame:
    rows = []
    for profile in profiles:
        for r in profile.resolutions:
            if r.scenario == "C":
                severity = "HIGH"
                issue = "duplicate_column_manual_review"
            elif r.action == "REMOVE":
                severity = "INFO"
                issue = f"duplicate_column_auto_resolved_scenario_{r.scenario}"
            else:
                continue  # KEEP entries are not issues

            rows.append({
                "Severity":        severity,
                "Issue Type":      issue,
                "Workbook":        r.workbook,
                "Source Tab":      r.source_tab,
                "Canonical Name":  r.canonical_name,
                "Original Column": r.original_col_name,
                "Position":        r.original_position,
                "KPIs Using":      ", ".join(r.kpis_using),
                "Action":          r.action,
                "Detail":          r.reason,
            })
    if not rows:
        rows.append({
            "Severity": "INFO", "Issue Type": "none",
            "Workbook": "", "Source Tab": "", "Canonical Name": "",
            "Original Column": "", "Position": "", "KPIs Using": "",
            "Action": "", "Detail": "No issues detected",
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Data Dictionary and Reconciliation (unchanged logic, updated signatures)
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
            "Canonical Column":    canonical,
            "Match Class":         profile.match_class if profile else "unknown",
            "Present In Workbooks": profile.present_in_count if profile else 0,
            "Source Workbooks":    ", ".join(profile.present_in_workbooks) if profile else "",
            "Data Types":          " | ".join(
                f"{wb}:{dt}" for wb, dt in (profile.data_types.items() if profile else {})
            ),
            "KPI Referenced":      "yes" if (profile and profile.is_kpi_referenced) else "no",
            "Used By KPIs":        ", ".join(kpis_using),
            "Similar To":          ", ".join(profile.similar_to) if profile else "",
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=[
        "Canonical Column", "Match Class", "Present In Workbooks", "Source Workbooks",
        "Data Types", "KPI Referenced", "Used By KPIs", "Similar To",
    ])


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
        master_rows = (
            len(master_df[master_df["Source_Workbook"] == bundle.file_name])
            if "Source_Workbook" in master_df.columns else "N/A"
        )
        status = (
            "PASS — row counts match" if master_rows == source_rows
            else "SKIP — master not built" if master_rows == "N/A"
            else f"WARN — source {source_rows} rows vs master {master_rows}"
        )
        rows.append({
            "Check": "source_row_count",
            "Workbook": bundle.file_name,
            "Source Tab": wb_cfg.source_tab,
            "Source Rows": source_rows,
            "Master Rows": master_rows,
            "Status": status,
        })
    master_total = len(master_df) if not master_df.empty else 0
    rows.append({
        "Check": "total_source_rows",
        "Workbook": "(all)",
        "Source Tab": "",
        "Source Rows": total_source_rows,
        "Master Rows": master_total,
        "Status": (
            "PASS" if master_total == total_source_rows
            else f"WARN — total source {total_source_rows} vs master {master_total}"
        ),
    })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Master Source Data builder — resolve duplicates, then concat
# ---------------------------------------------------------------------------

def build_master_source_df(
    bundles: list,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult,
    kpi_result: KpiAnalysisResult,
) -> tuple[pd.DataFrame, list[SourceFrameProfile]]:
    """Consolidate rows from all configured source tabs into one DataFrame.

    For each bundle:
    1. Build the rename map from the canonical column mapping.
    2. Detect duplicate canonical names within the original source frame.
    3. Resolve each duplicate group using KPI-usage scenarios (A/B/C/D).
    4. Drop columns resolved as REMOVE; keep columns resolved as KEEP.
    5. Rename remaining columns to canonical names.
    6. Prepend lineage columns (Source_Workbook, Source_Sheet, LOB_Identifier).

    Returns:
        ``(master_df, profiles)`` — profiles contain per-frame duplicate
        findings and resolution decisions.

    Raises:
        DuplicateColumnError: Only when Scenario C duplicates remain (both
            positions referenced by different KPIs — cannot auto-resolve).
            The caller can still access ``exc.resolutions`` for reporting.
    """
    excl_set = set(source_result.excluded_columns)
    frames: list[pd.DataFrame] = []
    profiles: list[SourceFrameProfile] = []
    all_resolutions: list[DuplicateResolution] = []

    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        df = bundle.sheets.get(wb_cfg.source_tab)
        if df is None:
            logger.warning("Source tab '%s' missing from '%s'", wb_cfg.source_tab, bundle.file_name)
            continue

        # Build rename map (original_col → canonical) and excluded set
        rename_map: dict[str, str] = {}
        drop_cols: set[str] = set()
        for col in df.columns:
            col_str = str(col)
            key = (bundle.file_name, wb_cfg.source_tab, col_str)
            if key in excl_set:
                drop_cols.add(col_str)
            elif key in source_result.column_mapping:
                rename_map[col_str] = source_result.column_mapping[key]

        # Detect duplicates on the ORIGINAL frame (positions match KPI raw_refs letters)
        findings = _detect_raw_duplicates(df, rename_map, bundle.file_name, wb_cfg.source_tab)

        # Resolve each duplicate group via KPI usage
        bundle_resolutions: list[DuplicateResolution] = []
        for finding in findings:
            resolutions = resolve_duplicate_group(finding, kpi_result, df)
            bundle_resolutions.extend(resolutions)
            all_resolutions.extend(resolutions)

        # Build the extra drop set from resolution decisions (REMOVE actions)
        for r in bundle_resolutions:
            if r.action == "REMOVE":
                orig_col = r.original_col_name
                drop_cols.add(orig_col)
                logger.info(
                    "Scenario %s — removing '%s' (pos %d) from '%s'/'%s': %s",
                    r.scenario, orig_col, r.original_position,
                    bundle.file_name, wb_cfg.source_tab, r.reason,
                )

        # Apply drops and renames
        frame = df.copy()
        cols_to_drop = [c for c in frame.columns if str(c) in drop_cols]
        if cols_to_drop:
            frame = frame.drop(columns=cols_to_drop, errors="ignore")
        frame = frame.rename(columns=rename_map)

        # Prepend lineage columns
        frame.insert(0, "LOB_Identifier", wb_cfg.lob_identifier or wb_cfg.workbook_name)
        frame.insert(0, "Source_Sheet", wb_cfg.source_tab)
        frame.insert(0, "Source_Workbook", bundle.file_name)

        profiles.append(SourceFrameProfile(
            workbook=bundle.file_name,
            source_tab=wb_cfg.source_tab,
            shape=frame.shape,
            all_columns=list(frame.columns),
            duplicate_findings=findings,
            resolutions=bundle_resolutions,
        ))
        frames.append(frame)

    # Check for unresolvable Scenario C duplicates
    blocking = [r for r in all_resolutions if r.scenario == "C"]

    if not frames:
        return pd.DataFrame(columns=["Source_Workbook", "Source_Sheet", "LOB_Identifier"]), profiles

    if blocking:
        raise DuplicateColumnError(all_resolutions)

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
            else "References configured source tab — update column positions after merge"
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
    """Assemble the future-state workbook and return as bytes.

    Applies KPI-usage-based duplicate resolution (Scenarios A/B/C/D).
    Only raises :class:`DuplicateColumnError` for Scenario C cases.
    In that case, ``exc.diagnostic_bytes`` contains the workbook with
    sheets 06–09 populated so the user can review the findings.
    """
    wb = Workbook()
    wb.remove(wb.active)

    blocking_error: DuplicateColumnError | None = None
    master_df = pd.DataFrame(columns=["Source_Workbook", "Source_Sheet", "LOB_Identifier"])
    profiles: list[SourceFrameProfile] = []

    try:
        master_df, profiles = build_master_source_df(bundles, config, source_result, kpi_result)
    except DuplicateColumnError as exc:
        blocking_error = exc
        # Rebuild profiles using the resolutions the error carries
        profiles = _profiles_from_resolutions(bundles, config, source_result, kpi_result)

    # 01 — Master Source Data
    ws_master = wb.create_sheet("01_Master_Source_Data")
    _write_df_to_sheet(ws_master, master_df, title="Master Source Data")
    if blocking_error:
        _write_warning_banner(
            ws_master,
            "⚠ Incomplete — Scenario C duplicates require manual review. "
            "See 07_Duplicate_Column_Analysis.",
        )

    # 02+ — KPI audit tabs
    kpi_tab_counter = 2
    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        for kpi_tab in wb_cfg.kpi_tabs:
            audit_df = build_kpi_audit_df(
                bundle, kpi_tab, kpi_result, config.future_source_tab_name
            )
            tab_name = (
                f"{kpi_tab_counter:02d}_KPI_Summary_"
                f"{config.kpi_tab_name_for(bundle.file_name, kpi_tab)}"
            )[:31]
            ws_kpi = wb.create_sheet(tab_name)
            _write_df_to_sheet(ws_kpi, audit_df, title=f"KPI Audit: {kpi_tab}")
            kpi_tab_counter += 1

    # 03 — Source Mapping
    ws_map = wb.create_sheet("03_Source_Mapping")
    _write_df_to_sheet(
        ws_map, source_result.to_mapping_dataframe(),
        title="Source Column → Canonical Mapping",
    )

    # 04 — Data Dictionary
    ws_dd = wb.create_sheet("04_Data_Dictionary")
    _write_df_to_sheet(ws_dd, build_data_dictionary_df(source_result, kpi_result),
                       title="Data Dictionary")

    # 05 — Reconciliation
    ws_rec = wb.create_sheet("05_Reconciliation")
    _write_df_to_sheet(ws_rec, build_reconciliation_df(bundles, config, master_df),
                       title="Row-Count Reconciliation")

    # 06 — Issues Log
    ws_issues = wb.create_sheet("06_Issues_Log")
    issues_df = build_issues_log_df(profiles)
    _write_df_to_sheet(ws_issues, issues_df, title="Issues Log")
    _apply_severity_colours(ws_issues, issues_df)

    # 07 — Duplicate Column Analysis
    ws_dup = wb.create_sheet("07_Duplicate_Column_Analysis")
    dup_df = build_duplicate_column_analysis_df(profiles)
    _write_df_to_sheet(ws_dup, dup_df, title="Duplicate Column Analysis")
    _apply_scenario_colours(ws_dup, dup_df)

    # 08 — Workbook Source Analysis
    ws_src = wb.create_sheet("08_Workbook_Source_Analysis")
    _write_df_to_sheet(ws_src, build_workbook_source_analysis_df(profiles),
                       title="Workbook Source Analysis")

    # 09 — Documentation
    ws_doc = wb.create_sheet("09_Documentation")
    _write_documentation(ws_doc, config, source_result, kpi_result, master_df, profiles)

    buf = io.BytesIO()
    wb.save(buf)
    xlsx_bytes = buf.getvalue()

    if blocking_error:
        blocking_error.diagnostic_bytes = xlsx_bytes
        raise blocking_error

    return xlsx_bytes


def build_rationalized_workbook(
    bundles: list,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult,
    kpi_result: KpiAnalysisResult,
    output_path: Path,
) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    data = build_rationalized_workbook_bytes(bundles, config, source_result, kpi_result)
    output_path.write_bytes(data)
    logger.info("Future-state workbook written to %s", output_path)
    return output_path.resolve()


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _profiles_from_resolutions(
    bundles: list,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult,
    kpi_result: KpiAnalysisResult,
) -> list[SourceFrameProfile]:
    """Rebuild SourceFrameProfile objects after a DuplicateColumnError."""
    excl_set = set(source_result.excluded_columns)
    profiles: list[SourceFrameProfile] = []
    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        df = bundle.sheets.get(wb_cfg.source_tab)
        if df is None:
            continue
        rename_map: dict[str, str] = {}
        drop_cols: set[str] = set()
        for col in df.columns:
            col_str = str(col)
            key = (bundle.file_name, wb_cfg.source_tab, col_str)
            if key in excl_set:
                drop_cols.add(col_str)
            elif key in source_result.column_mapping:
                rename_map[col_str] = source_result.column_mapping[key]
        findings = _detect_raw_duplicates(df, rename_map, bundle.file_name, wb_cfg.source_tab)
        resolutions: list[DuplicateResolution] = []
        for finding in findings:
            resolutions.extend(resolve_duplicate_group(finding, kpi_result, df))
        frame = df.copy()
        all_drop = drop_cols | {r.original_col_name for r in resolutions if r.action == "REMOVE"}
        cols_to_drop = [c for c in frame.columns if str(c) in all_drop]
        if cols_to_drop:
            frame = frame.drop(columns=cols_to_drop, errors="ignore")
        frame = frame.rename(columns=rename_map)
        frame.insert(0, "LOB_Identifier", wb_cfg.lob_identifier or wb_cfg.workbook_name)
        frame.insert(0, "Source_Sheet", wb_cfg.source_tab)
        frame.insert(0, "Source_Workbook", bundle.file_name)
        profiles.append(SourceFrameProfile(
            workbook=bundle.file_name,
            source_tab=wb_cfg.source_tab,
            shape=frame.shape,
            all_columns=list(frame.columns),
            duplicate_findings=findings,
            resolutions=resolutions,
        ))
    return profiles


def _write_df_to_sheet(ws, df: pd.DataFrame, title: str | None = None) -> None:
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
    ws.insert_rows(1)
    cell = ws.cell(row=1, column=1, value=message)
    cell.fill = _WARN_FILL
    cell.font = Font(bold=True, color="7B3F00")


def _apply_severity_colours(ws, df: pd.DataFrame) -> None:
    if "Severity" not in df.columns:
        return
    for row_idx, severity in enumerate(df["Severity"], start=3):
        fill = _ERROR_FILL if severity == "HIGH" else (_WARN_FILL if severity == "WARN" else None)
        if fill:
            for col_idx in range(1, len(df.columns) + 1):
                ws.cell(row=row_idx, column=col_idx).fill = fill


def _apply_scenario_colours(ws, df: pd.DataFrame) -> None:
    """Colour the Duplicate Column Analysis rows by scenario/action."""
    if "Scenario" not in df.columns or "Recommended Action" not in df.columns:
        return
    scenario_col = list(df.columns).index("Scenario") + 1
    action_col = list(df.columns).index("Recommended Action") + 1
    for row_idx, (scenario, action) in enumerate(
        zip(df["Scenario"], df["Recommended Action"]), start=3
    ):
        if action == "MANUAL_REVIEW":
            fill = _ERROR_FILL
        elif action == "KEEP":
            fill = _OK_FILL
        elif action == "REMOVE":
            fill = _ALT_FILL
        else:
            fill = None
        if fill:
            for col_idx in range(1, len(df.columns) + 1):
                ws.cell(row=row_idx, column=col_idx).fill = fill


def _write_documentation(
    ws,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult,
    kpi_result: KpiAnalysisResult,
    master_df: pd.DataFrame,
    profiles: list[SourceFrameProfile],
) -> None:
    all_resolutions = [r for p in profiles for r in p.resolutions]
    auto_removed  = sum(1 for r in all_resolutions if r.action == "REMOVE" and r.scenario != "C")
    manual_review = sum(1 for r in all_resolutions if r.scenario == "C")

    rows: list[tuple[str, str]] = [
        ("GUIDED WORKBOOK RATIONALIZATION — DOCUMENTATION", ""),
        ("", ""),
        ("1. Run Configuration", ""),
        ("Future Source Tab Name",         config.future_source_tab_name),
        ("Matching Threshold",             str(config.matching_threshold)),
        ("Merge High-Confidence Matches",  str(config.merge_high_confidence)),
        ("Remove Unused Columns",          str(config.remove_unused_columns)),
        ("", ""),
        ("2. Source Workbooks Analyzed", ""),
    ]
    for wb_cfg in config.workbook_configs:
        rows += [
            ("  Workbook",        wb_cfg.workbook_name),
            ("  Source Tab",      wb_cfg.source_tab),
            ("  KPI Tabs",        ", ".join(wb_cfg.kpi_tabs) or "(none)"),
            ("  LOB Identifier",  wb_cfg.lob_identifier or "(none)"),
            ("", ""),
        ]
    rows += [
        ("3. Source-to-Target Mappings", ""),
        ("Common columns (auto-merged)",  str(len(source_result.common_columns))),
        ("Similar columns (documented)",  str(len(source_result.similar_columns))),
        ("Unique columns (kept as-is)",   str(len(source_result.unique_columns))),
        ("", ""),
        ("4. Duplicate Column Resolution", ""),
        ("Scenario A (keep KPI-used, remove unused)", str(sum(
            1 for r in all_resolutions if r.scenario == "A" and r.action == "REMOVE"
        ))),
        ("Scenario B (both unused — removed)", str(sum(
            1 for r in all_resolutions if r.scenario == "B"
        ))),
        ("Scenario C (both used — manual review)", str(manual_review)),
        ("Scenario D (identical values — deduplicated)", str(sum(
            1 for r in all_resolutions if r.scenario == "D" and r.action == "REMOVE"
        ))),
        ("Total columns auto-removed",    str(auto_removed)),
        ("", ""),
        ("5. Column Retention", ""),
        ("Columns retained",  str(len(set(source_result.column_mapping.values())))),
        ("Columns excluded (not KPI-referenced)",
         str(len(source_result.excluded_columns))),
        ("", ""),
        ("6. KPI Dependencies", ""),
        ("KPI formula cells analyzed",              str(len(kpi_result.dependencies))),
        ("Source columns referenced by ≥1 KPI",    str(len(kpi_result.referenced_canonicals))),
        ("Source columns not referenced by any KPI", str(len(kpi_result.unreferenced_canonicals))),
        ("", ""),
        ("7. Master Source Data", ""),
        ("Total rows",                str(len(master_df))),
        ("Total columns (incl. lineage)", str(len(master_df.columns))),
        ("Lineage columns",           "Source_Workbook, Source_Sheet, LOB_Identifier"),
        ("", ""),
        ("8. BAU Update Instructions", ""),
        ("Step 1", "Copy new source data rows into the Master_Source_Data tab."),
        ("Step 2", "Ensure Source_Workbook, Source_Sheet, and LOB_Identifier are populated."),
        ("Step 3", "Update KPI formula range references to cover new rows."),
        ("Step 4", "Validate KPI outputs against original workbook totals."),
        ("Step 5", "Archive original workbooks — do not delete until reconciled."),
    ]
    if manual_review > 0:
        rows += [
            ("", ""),
            ("⚠ Action Required", ""),
            ("Manual review needed", (
                f"{manual_review} duplicate column position(s) are each referenced by "
                "different KPI formulas. See 07_Duplicate_Column_Analysis. "
                "Rename the conflicting source columns to give them distinct identities, "
                "then re-run the guided rationalization."
            )),
        ]

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 60
    for r_idx, (label, value) in enumerate(rows, start=1):
        cell_a = ws.cell(row=r_idx, column=1, value=label)
        ws.cell(row=r_idx, column=2, value=value)
        if r_idx == 1:
            cell_a.font = Font(bold=True, size=13)
        elif label and label[0].isdigit() and ". " in label:
            cell_a.font = _SECTION_FONT
        elif label.startswith("⚠"):
            cell_a.fill = _ERROR_FILL
            ws.cell(row=r_idx, column=2).fill = _ERROR_FILL
