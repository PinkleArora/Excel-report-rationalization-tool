"""Analyse source data tabs across workbooks.

Classifies every column as COMMON, SIMILAR, or UNIQUE:

COMMON  — exact normalized name match in ≥ 2 workbooks → auto-merged.
SIMILAR — fuzzy score ≥ threshold but not exact → documented for review,
           merged only if RationalizationConfig.merge_high_confidence=True
           and score ≥ 95.
UNIQUE  — appears in only one workbook → kept as-is under canonical name.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field

import pandas as pd
from rapidfuzz import fuzz

from src.guided_rationalization.config import RationalizationConfig, WorkbookTabConfig
from src.schema_matching.normalizer import normalize_column_name

logger = logging.getLogger(__name__)


@dataclass
class ColumnProfile:
    """Analysis of a single canonical column across all source workbooks."""

    canonical_name: str
    match_class: str                   # "common" | "similar" | "unique"
    source_entries: list[tuple[str, str, str]] = field(default_factory=list)
    # (workbook_name, tab_name, original_col_name)
    similar_to: list[str] = field(default_factory=list)  # other canonical names it resembles
    data_types: dict[str, str] = field(default_factory=dict)  # workbook → pandas dtype str
    is_kpi_referenced: bool = False
    referenced_by_kpis: list[str] = field(default_factory=list)  # KPI labels

    @property
    def present_in_workbooks(self) -> list[str]:
        return sorted({wb for wb, _, _ in self.source_entries})

    @property
    def present_in_count(self) -> int:
        return len(self.present_in_workbooks)


@dataclass
class SourceAnalysisResult:
    """Outcome of source data analysis across all configured workbooks."""

    column_profiles: list[ColumnProfile]
    # (workbook, source_tab, original_col) → canonical_col
    column_mapping: dict[tuple[str, str, str], str]
    # columns excluded from master (not referenced by any KPI, if configured)
    excluded_columns: list[tuple[str, str, str]]

    @property
    def common_columns(self) -> list[ColumnProfile]:
        return [p for p in self.column_profiles if p.match_class == "common"]

    @property
    def similar_columns(self) -> list[ColumnProfile]:
        return [p for p in self.column_profiles if p.match_class == "similar"]

    @property
    def unique_columns(self) -> list[ColumnProfile]:
        return [p for p in self.column_profiles if p.match_class == "unique"]

    @property
    def kpi_referenced_columns(self) -> list[ColumnProfile]:
        return [p for p in self.column_profiles if p.is_kpi_referenced]

    def to_dataframe(self) -> pd.DataFrame:
        """Tabular view of the source rationalization analysis."""
        rows = []
        for p in self.column_profiles:
            rows.append({
                "canonical_column":      p.canonical_name,
                "match_class":           p.match_class,
                "present_in_workbooks":  p.present_in_count,
                "workbooks":             ", ".join(p.present_in_workbooks),
                "original_columns":      " | ".join(
                    f"{wb}/{tab}/{col}" for wb, tab, col in p.source_entries
                ),
                "data_types":            " | ".join(
                    f"{wb}:{dt}" for wb, dt in p.data_types.items()
                ),
                "similar_to":            ", ".join(p.similar_to),
                "is_kpi_referenced":     p.is_kpi_referenced,
                "referenced_by_kpis":    ", ".join(p.referenced_by_kpis),
                "in_master_source":      p.canonical_name not in {
                    c for _, _, c in []  # placeholder — set after exclusion list
                },
            })
        return pd.DataFrame(rows)

    def to_mapping_dataframe(self) -> pd.DataFrame:
        """One row per source column showing the canonical mapping and transformation applied."""
        rows = []
        excl_set = set(self.excluded_columns)
        for (wb, tab, orig_col), canonical in self.column_mapping.items():
            profile = next(
                (p for p in self.column_profiles if p.canonical_name == canonical), None
            )
            rows.append({
                "source_workbook":      wb,
                "source_tab":           tab,
                "source_column":        orig_col,
                "canonical_column":     canonical,
                "match_class":          profile.match_class if profile else "unknown",
                "transformation":       (
                    "renamed" if normalize_column_name(orig_col) != canonical
                    else "no change"
                ),
                "in_master_source":     (wb, tab, orig_col) not in excl_set,
                "exclusion_reason":     (
                    "Not referenced by any KPI formula"
                    if (wb, tab, orig_col) in excl_set else ""
                ),
            })
        return pd.DataFrame(rows)

    def to_enriched_mapping_dataframe(
        self,
        kpi_result,
        config,
    ) -> pd.DataFrame:
        """Source mapping enriched with KPI dependency information.

        Adds columns:
        - kpi_tabs_using_column  — semicolon-separated future output tab names
        - kpi_labels_using_column — semicolon-separated KPI labels
        - usage_count             — number of KPI cells referencing this canonical
        - retention_reason        — why the column is in Master Source Data
        """
        # Build lookup: canonical → list of (future_tab_name, kpi_label) per dependency
        from collections import defaultdict
        canonical_to_deps: dict[str, list[tuple[str, str]]] = defaultdict(list)
        for dep in kpi_result.dependencies:
            future_name = config.kpi_tab_name_for(dep.workbook_name, dep.kpi_tab)
            for canonical in dep.canonical_source_columns:
                canonical_to_deps[canonical].append((future_name, dep.kpi_label))

        _LINEAGE_COLS = {"Source_Workbook", "Source_Sheet"}
        excl_set = set(self.excluded_columns)
        rows = []
        for (wb, tab, orig_col), canonical in self.column_mapping.items():
            profile = next(
                (p for p in self.column_profiles if p.canonical_name == canonical), None
            )
            is_excluded = (wb, tab, orig_col) in excl_set
            deps = canonical_to_deps.get(canonical, [])

            # Deduplicate while preserving order
            seen_tabs: dict[str, None] = {}
            seen_labels: dict[str, None] = {}
            for tab_name, label in deps:
                seen_tabs[tab_name] = None
                seen_labels[label] = None
            kpi_tabs_str = "; ".join(seen_tabs)
            kpi_labels_str = "; ".join(seen_labels)
            usage_count = len(deps)

            if canonical in _LINEAGE_COLS:
                retention_reason = "Lineage column"
            elif deps:
                retention_reason = "Used by KPI calculations"
            elif is_excluded:
                retention_reason = "Not referenced by any KPI formula"
            else:
                retention_reason = "Retained (unreferenced)"

            rows.append({
                "source_workbook":          wb,
                "source_tab":               tab,
                "source_column":            orig_col,
                "canonical_column":         canonical,
                "match_class":              profile.match_class if profile else "unknown",
                "transformation":           (
                    "renamed" if normalize_column_name(orig_col) != canonical
                    else "no change"
                ),
                "in_master_source":         not is_excluded,
                "kpi_tabs_using_column":    kpi_tabs_str,
                "kpi_labels_using_column":  kpi_labels_str,
                "usage_count":              usage_count,
                "retention_reason":         retention_reason,
            })
        return pd.DataFrame(rows)



def analyze_source_data(
    bundles: list,
    config: RationalizationConfig,
    kpi_referenced_canonicals: set[str] | None = None,
) -> SourceAnalysisResult:
    """Classify every source column as COMMON, SIMILAR, or UNIQUE.

    Args:
        bundles: All loaded WorkbookBundle objects.
        config: Full rationalization configuration (which tab is source per workbook).
        kpi_referenced_canonicals: Canonical names referenced by at least one KPI formula.
            Used to populate ``is_kpi_referenced`` and — when
            ``config.remove_unused_columns`` is True — to determine exclusions.

    Returns:
        :class:`SourceAnalysisResult` with profiles, column map, and exclusions.
    """
    # 1. Collect (workbook, tab, original_col, normalized_col, dtype) entries
    entries: list[tuple[str, str, str, str, str]] = []
    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        source_tab = wb_cfg.source_tab
        df = bundle.sheets.get(source_tab)
        if df is None:
            logger.warning("Source tab '%s' not found in '%s' — skipping",
                           source_tab, bundle.file_name)
            continue
        for col in df.columns:
            col_str = str(col)
            norm = normalize_column_name(col_str)
            dtype_str = str(df[col].dtype)
            entries.append((bundle.file_name, source_tab, col_str, norm, dtype_str))

    # 2. Group by normalized name → find COMMON columns (same norm in ≥2 workbooks)
    from collections import defaultdict
    norm_to_entries: dict[str, list[tuple]] = defaultdict(list)
    for wb, tab, orig, norm, dtype in entries:
        norm_to_entries[norm].append((wb, tab, orig, norm, dtype))

    column_profiles: list[ColumnProfile] = []
    column_mapping: dict[tuple[str, str, str], str] = {}

    # Track which norms have been claimed (to avoid double-processing in similarity pass)
    claimed_norms: set[str] = set()

    # 2a. COMMON: same normalized name in ≥2 workbooks
    for norm, group in norm_to_entries.items():
        workbooks_present = {wb for wb, _, _, _, _ in group}
        if len(workbooks_present) >= 2:
            profile = ColumnProfile(
                canonical_name=norm,
                match_class="common",
                source_entries=[(wb, tab, orig) for wb, tab, orig, _, _ in group],
                data_types={wb: dtype for wb, _, _, _, dtype in group},
            )
            column_profiles.append(profile)
            for wb, tab, orig, _, _ in group:
                column_mapping[(wb, tab, orig)] = norm
            claimed_norms.add(norm)

    # 2b. SIMILAR: fuzzy score ≥ threshold between unclaimed norms
    unclaimed = {norm: grp for norm, grp in norm_to_entries.items()
                 if norm not in claimed_norms}
    unclaimed_norms = list(unclaimed.keys())

    merged_in_similar: set[str] = set()
    for i, norm_a in enumerate(unclaimed_norms):
        if norm_a in merged_in_similar:
            continue
        similar_group = [norm_a]
        for norm_b in unclaimed_norms[i + 1:]:
            if norm_b in merged_in_similar:
                continue
            score = fuzz.WRatio(norm_a, norm_b)
            if score >= config.matching_threshold:
                similar_group.append(norm_b)

        if len(similar_group) > 1:
            # All are similar to each other; canonical = shortest/earliest
            canonical = min(similar_group, key=lambda s: (len(s), s))
            all_entries = [e for n in similar_group for e in unclaimed[n]]
            # Check: are these genuinely different workbooks?
            workbooks_present = {wb for wb, _, _, _, _ in all_entries}
            is_cross_workbook = len(workbooks_present) >= 2

            # Choose match_class:
            # "similar" if cross-workbook and not merged, "common" if we actually merge
            do_merge = (
                is_cross_workbook and
                config.merge_high_confidence and
                all(fuzz.WRatio(norm_a, n) >= 95 for n in similar_group[1:])
            )
            match_class = "common" if do_merge else "similar"
            similar_to_list = [n for n in similar_group if n != canonical]

            profile = ColumnProfile(
                canonical_name=canonical,
                match_class=match_class,
                source_entries=[(wb, tab, orig) for wb, tab, orig, _, _ in all_entries],
                data_types={wb: dtype for wb, _, _, _, dtype in all_entries},
                similar_to=similar_to_list if match_class == "similar" else [],
            )
            column_profiles.append(profile)
            if do_merge:
                # All columns are genuinely equivalent — map everything to one canonical.
                for norm in similar_group:
                    for wb, tab, orig, _, _ in unclaimed[norm]:
                        column_mapping[(wb, tab, orig)] = canonical
                    merged_in_similar.add(norm)
            else:
                # SIMILAR but NOT merged — each column keeps its own normalised name
                # so that distinct business fields (e.g. "GAAP Reserve - ADB" and
                # "GAAP Reserve - WPA") remain separate columns in the output.
                # The profile documents the similarity relationship for the user.
                for norm in similar_group:
                    for wb, tab, orig, _, _ in unclaimed[norm]:
                        column_mapping[(wb, tab, orig)] = norm
                    merged_in_similar.add(norm)
        else:
            # UNIQUE — single workbook or no similar found
            group = unclaimed[norm_a]
            profile = ColumnProfile(
                canonical_name=norm_a,
                match_class="unique",
                source_entries=[(wb, tab, orig) for wb, tab, orig, _, _ in group],
                data_types={wb: dtype for wb, _, _, _, dtype in group},
            )
            column_profiles.append(profile)
            for wb, tab, orig, _, _ in group:
                column_mapping[(wb, tab, orig)] = norm_a
            merged_in_similar.add(norm_a)

    # 3. Mark KPI-referenced columns
    if kpi_referenced_canonicals:
        for profile in column_profiles:
            if profile.canonical_name in kpi_referenced_canonicals:
                profile.is_kpi_referenced = True

    # 4. Determine excluded columns
    excluded: list[tuple[str, str, str]] = []
    if config.remove_unused_columns and kpi_referenced_canonicals is not None:
        for (wb, tab, orig), canonical in column_mapping.items():
            if canonical not in kpi_referenced_canonicals:
                excluded.append((wb, tab, orig))

    logger.info(
        "Source analysis: %d common, %d similar, %d unique columns; %d excluded",
        len([p for p in column_profiles if p.match_class == "common"]),
        len([p for p in column_profiles if p.match_class == "similar"]),
        len([p for p in column_profiles if p.match_class == "unique"]),
        len(excluded),
    )
    return SourceAnalysisResult(
        column_profiles=column_profiles,
        column_mapping=column_mapping,
        excluded_columns=excluded,
    )
