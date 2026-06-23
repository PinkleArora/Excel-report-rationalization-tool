"""Consolidation Intelligence Agent — grain detection, pairwise scoring, grouping."""

from __future__ import annotations

import itertools
from dataclasses import dataclass, field

from src.guided_rationalization.config import RationalizationConfig
from src.agentic_rationalization.models import AgentDecision, AgentResult
from src.agentic_rationalization.grain_detector import detect_grain, GrainDetectionResult
from src.agentic_rationalization.compatibility_scorer import score_compatibility, CompatibilityScore
from src.schema_matching.normalizer import normalize_column_name

_AGENT_NAME = "ConsolidationIntelligenceAgent"

_TIME_KEYWORDS = {
    "date", "period", "month", "year", "quarter",
    "effective_date", "as_of_date", "asofdate", "effectivedate",
}


@dataclass
class FileProfile:
    file_name: str
    row_count: int
    column_count: int
    distinct_columns: int
    inferred_grain: str
    grain_confidence: float
    source_tab: str
    kpi_referenced_column_count: int
    non_kpi_column_count: int
    primary_key_candidates: list[str]
    has_time_dimension: bool


@dataclass
class ConsolidationGroup:
    group_id: int
    group_label: str
    grain_category: str
    file_names: list[str]
    common_column_count: int
    unique_column_count: int
    kpi_required_column_count: int
    discardable_column_count: int
    recommendation: str  # "Consolidate" | "Keep Separate"
    expected_master_columns: int
    reasoning: str


@dataclass
class ColumnDiscardRec:
    canonical_name: str
    used_by_kpi: bool
    kpi_names: list[str]
    recommendation: str  # "Keep" | "Discard"
    confidence: str  # "High" | "Medium" | "Low"
    reason_retained: str


@dataclass
class DiscardAnalysis:
    total_original_columns: int
    kpi_used_columns: int
    non_kpi_columns: int
    potential_reduction_pct: float
    column_recs: list[ColumnDiscardRec]


@dataclass
class FinalRecommendation:
    summary: str
    business_reasoning: str
    consolidate_files: list[str]
    separate_model_files: list[str]
    exclude_files: list[str]


@dataclass
class ConsolidationIntelligenceResult:
    file_profiles: list[FileProfile]
    pairwise_scores: list[CompatibilityScore]
    groups: list[ConsolidationGroup]
    discard_analysis: DiscardAnalysis
    final_recommendation: FinalRecommendation
    visual_map: str


# ---------------------------------------------------------------------------
# Union-Find helpers
# ---------------------------------------------------------------------------

def _build_union_find(files: list[str]) -> dict[str, str]:
    """Initialise union-find parent map."""
    return {f: f for f in files}


def _find(parent: dict[str, str], x: str) -> str:
    while parent[x] != x:
        parent[x] = parent[parent[x]]
        x = parent[x]
    return x


def _union(parent: dict[str, str], x: str, y: str) -> None:
    rx, ry = _find(parent, x), _find(parent, y)
    if rx != ry:
        parent[ry] = rx


def _build_groups(
    files: list[str],
    pairwise: list[CompatibilityScore],
    grain_map: dict[str, GrainDetectionResult],
    col_map: dict[str, list[str]],
    kpi_analysis,
) -> list[ConsolidationGroup]:
    """Cluster files into ConsolidationGroups using union-find."""
    parent = _build_union_find(files)

    for sc in pairwise:
        if sc.overall_score >= 0.50 and sc.grain_compatible:
            _union(parent, sc.file_a, sc.file_b)

    # Build clusters
    clusters: dict[str, list[str]] = {}
    for f in files:
        root = _find(parent, f)
        clusters.setdefault(root, []).append(f)

    kpi_referenced: set[str] = set()
    if kpi_analysis is not None:
        kpi_referenced = getattr(kpi_analysis, "referenced_canonicals", set())

    groups: list[ConsolidationGroup] = []
    for group_id, (root, members) in enumerate(sorted(clusters.items()), start=1):
        # Columns
        all_cols: list[set[str]] = []
        for fname in members:
            normed = {normalize_column_name(c) for c in col_map.get(fname, [])}
            all_cols.append(normed)

        union_cols: set[str] = set().union(*all_cols) if all_cols else set()
        common_cols: set[str] = all_cols[0].copy() if all_cols else set()
        for s in all_cols[1:]:
            common_cols &= s

        kpi_required = len(union_cols & kpi_referenced)
        discardable = len(union_cols - kpi_referenced - common_cols)

        # Grain
        grain_cat = grain_map[root].grain_category if root in grain_map else "unknown"
        # Most common grain among members
        from collections import Counter
        grain_counts = Counter(grain_map[f].grain_category for f in members if f in grain_map)
        if grain_counts:
            grain_cat = grain_counts.most_common(1)[0][0]

        recommendation = "Consolidate" if len(members) > 1 else "Keep Separate"

        reasoning = (
            f"{len(members)} file(s) share '{grain_cat}' grain. "
            f"Common columns: {len(common_cols)}, total union columns: {len(union_cols)}. "
            f"KPI-required: {kpi_required}, discardable: {discardable}."
        )

        groups.append(ConsolidationGroup(
            group_id=group_id,
            group_label=f"{grain_cat.replace('_', ' ').title()} Group {group_id}",
            grain_category=grain_cat,
            file_names=members,
            common_column_count=len(common_cols),
            unique_column_count=len(union_cols),
            kpi_required_column_count=kpi_required,
            discardable_column_count=discardable,
            recommendation=recommendation,
            expected_master_columns=len(union_cols),
            reasoning=reasoning,
        ))

    return groups


# ---------------------------------------------------------------------------
# Main run() function
# ---------------------------------------------------------------------------

def run(
    bundles: list,
    config: RationalizationConfig,
    kpi_analysis=None,
    source_analysis=None,
) -> AgentResult:
    """Run Consolidation Intelligence analysis.

    Args:
        bundles: Loaded WorkbookBundle objects.
        config: RationalizationConfig.
        kpi_analysis: KpiAnalysisResult from KPI Agent (may be None).
        source_analysis: SourceAnalysisResult from Schema Agent (may be None).

    Returns:
        AgentResult whose ``output`` is a ConsolidationIntelligenceResult.
    """
    decisions: list[AgentDecision] = []
    warnings: list[str] = []

    kpi_referenced: set[str] = set()
    kpi_label_map: dict[str, list[str]] = {}
    if kpi_analysis is not None:
        kpi_referenced = getattr(kpi_analysis, "referenced_canonicals", set())
        for dep in getattr(kpi_analysis, "dependencies", []):
            for canonical in getattr(dep, "canonical_source_columns", []):
                kpi_label_map.setdefault(canonical, [])
                label = getattr(dep, "kpi_label", "")
                if label and label not in kpi_label_map[canonical]:
                    kpi_label_map[canonical].append(label)

    # --- Build FileProfiles ---
    file_profiles: list[FileProfile] = []
    grain_map: dict[str, GrainDetectionResult] = {}
    col_map: dict[str, list[str]] = {}

    for bundle in bundles:
        file_name = getattr(bundle, "file_name", str(bundle))
        wb_cfg = config.config_for(file_name)
        if wb_cfg is None:
            warnings.append(f"No config found for '{file_name}'; skipping.")
            continue
        source_tab = wb_cfg.source_tab
        df = bundle.sheets.get(source_tab)

        if df is None:
            warnings.append(f"Bundle '{file_name}' has no sheet '{source_tab}'; skipping grain detection.")
            continue

        cols = list(df.columns)
        col_map[file_name] = cols
        grain = detect_grain(df, kpi_analysis)
        grain_map[file_name] = grain

        normed_cols = {normalize_column_name(c) for c in cols}
        kpi_col_count = len(normed_cols & kpi_referenced)
        non_kpi_col_count = len(normed_cols) - kpi_col_count

        has_time = any(
            any(kw in normalize_column_name(c) for kw in _TIME_KEYWORDS) for c in cols
        )

        fp = FileProfile(
            file_name=file_name,
            row_count=grain.row_count,
            column_count=grain.column_count,
            distinct_columns=len(normed_cols),
            inferred_grain=grain.grain_label,
            grain_confidence=grain.confidence,
            source_tab=source_tab,
            kpi_referenced_column_count=kpi_col_count,
            non_kpi_column_count=non_kpi_col_count,
            primary_key_candidates=grain.primary_key_candidates,
            has_time_dimension=has_time,
        )
        file_profiles.append(fp)

        decisions.append(AgentDecision(
            subject=file_name,
            decision=f"grain={grain.grain_category}",
            confidence=grain.confidence,
            reasoning=grain.reasoning,
            signals={"row_count": grain.row_count, "column_count": grain.column_count},
        ))

    # --- Pairwise compatibility ---
    file_names = list(col_map.keys())
    pairwise_scores: list[CompatibilityScore] = []

    for fa, fb in itertools.combinations(file_names, 2):
        sc = score_compatibility(
            fa, col_map[fa], grain_map[fa],
            fb, col_map[fb], grain_map[fb],
            kpi_analysis,
        )
        pairwise_scores.append(sc)
        decisions.append(AgentDecision(
            subject=f"{fa} vs {fb}",
            decision=sc.recommendation,
            confidence=sc.recommendation_confidence,
            reasoning=sc.reasoning,
            signals={"overall_score": sc.overall_score, "grain_compatible": sc.grain_compatible},
        ))

    # --- Groups ---
    groups = _build_groups(file_names, pairwise_scores, grain_map, col_map, kpi_analysis)

    # --- Discard Analysis ---
    all_normed: set[str] = set()
    for cols in col_map.values():
        for c in cols:
            all_normed.add(normalize_column_name(c))

    column_recs: list[ColumnDiscardRec] = []
    for canonical in sorted(all_normed):
        used = canonical in kpi_referenced
        kpi_names = kpi_label_map.get(canonical, [])
        if used:
            rec = ColumnDiscardRec(
                canonical_name=canonical,
                used_by_kpi=True,
                kpi_names=kpi_names,
                recommendation="Keep",
                confidence="High",
                reason_retained="Referenced by KPI formula.",
            )
        else:
            rec = ColumnDiscardRec(
                canonical_name=canonical,
                used_by_kpi=False,
                kpi_names=[],
                recommendation="Discard",
                confidence="Medium",
                reason_retained="Not referenced by any KPI formula.",
            )
        column_recs.append(rec)

    total_cols = len(all_normed)
    kpi_used = len([r for r in column_recs if r.used_by_kpi])
    non_kpi = total_cols - kpi_used
    reduction_pct = (non_kpi / total_cols * 100) if total_cols > 0 else 0.0

    discard_analysis = DiscardAnalysis(
        total_original_columns=total_cols,
        kpi_used_columns=kpi_used,
        non_kpi_columns=non_kpi,
        potential_reduction_pct=reduction_pct,
        column_recs=column_recs,
    )

    # --- Final Recommendation ---
    consolidate_files: list[str] = []
    separate_files: list[str] = []
    for group in groups:
        if group.recommendation == "Consolidate":
            consolidate_files.extend(group.file_names)
        else:
            separate_files.extend(group.file_names)

    consolidate_count = len([g for g in groups if g.recommendation == "Consolidate"])
    if consolidate_count > 0:
        summary = f"{consolidate_count} group(s) recommended for consolidation."
        business_reasoning = (
            "Files sharing the same grain and overlapping KPI columns can be merged "
            "into a single master source, reducing duplication and maintenance overhead."
        )
    else:
        summary = "No files are candidates for consolidation based on current analysis."
        business_reasoning = (
            "Files operate at different grain levels or have insufficient column overlap "
            "to justify consolidation."
        )

    # Build ASCII visual map
    map_lines: list[str] = []
    for group in groups:
        icon = "+" if group.recommendation == "Consolidate" else "-"
        map_lines.append(f"[{icon}] Group {group.group_id}: {group.group_label}")
        for fname in group.file_names:
            map_lines.append(f"    |-- {fname}")
    visual_map = "\n".join(map_lines)

    final_rec = FinalRecommendation(
        summary=summary,
        business_reasoning=business_reasoning,
        consolidate_files=consolidate_files,
        separate_model_files=separate_files,
        exclude_files=[],
    )

    result = ConsolidationIntelligenceResult(
        file_profiles=file_profiles,
        pairwise_scores=pairwise_scores,
        groups=groups,
        discard_analysis=discard_analysis,
        final_recommendation=final_rec,
        visual_map=visual_map,
    )

    return AgentResult(
        agent_name=_AGENT_NAME,
        decisions=decisions,
        output=result,
        warnings=warnings,
    )
