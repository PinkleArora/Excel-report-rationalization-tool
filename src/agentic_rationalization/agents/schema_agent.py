"""Phase 3 — Schema Rationalization Agent.

Wraps source_analyzer and adds a KPI-usage veto: columns that are used by
*different* KPI calculations are never auto-merged regardless of fuzzy score.
This prevents false positives like "GAAP Reserve ADB" and "GAAP Reserve WPA"
being collapsed because they score ≥80 on WRatio.
"""

from __future__ import annotations

import logging

from src.guided_rationalization.config import RationalizationConfig
from src.guided_rationalization.kpi_analyzer import KpiAnalysisResult
from src.guided_rationalization.source_analyzer import (
    ColumnProfile,
    SourceAnalysisResult,
    analyze_source_data,
)
from src.agentic_rationalization.models import AgentDecision, AgentResult

logger = logging.getLogger(__name__)


def _classify_decision(profile: ColumnProfile) -> str:
    if profile.match_class == "common":
        return "MERGE"
    if profile.match_class == "similar":
        return "KEEP_SEPARATE (similar name)"
    return "KEEP_SEPARATE (unique)"


def _reasoning(
    profile: ColumnProfile,
    is_kpi_referenced: bool,
    kpi_veto_applied: bool,
    kpi_labels: list[str],
) -> str:
    wb_list = ", ".join(profile.present_in_workbooks)
    kpi_str = "; ".join(kpi_labels) if kpi_labels else "none"

    if profile.match_class == "common":
        return (
            f"Exact normalized name match across {profile.present_in_count} workbook(s) "
            f"({wb_list}). Auto-merged to canonical '{profile.canonical_name}'. "
            f"Referenced by KPIs: {kpi_str}."
        )
    if profile.match_class == "similar":
        if kpi_veto_applied:
            return (
                f"Name similarity ≥ threshold with other columns, but KPI veto applied: "
                f"each similar column is used by a different KPI ({kpi_str}). "
                "Kept separate to avoid breaking KPI calculations. "
                f"Canonical: '{profile.canonical_name}', similar to: "
                f"{', '.join(profile.similar_to)}."
            )
        return (
            f"Name similarity ≥ threshold with {', '.join(profile.similar_to)}, "
            "but treated as distinct business fields. "
            f"Canonical: '{profile.canonical_name}'. Referenced by KPIs: {kpi_str}."
        )
    # unique
    return (
        f"Column appears in only one workbook ({wb_list}). "
        f"Canonical: '{profile.canonical_name}'. Referenced by KPIs: {kpi_str}."
    )


def _confidence_for_profile(
    profile: ColumnProfile,
    is_kpi_referenced: bool,
    kpi_veto_applied: bool,
) -> float:
    if profile.match_class == "common" and not kpi_veto_applied:
        return 0.97
    if profile.match_class == "unique":
        return 0.92
    if profile.match_class == "similar" and kpi_veto_applied:
        # We're confident the veto was correctly applied
        return 0.90
    # similar without veto — less certain about the business meaning
    return 0.72


def run(
    bundles: list,
    config: RationalizationConfig,
    kpi_result: KpiAnalysisResult | None,
) -> AgentResult:
    """Classify source columns with a KPI-usage veto on fuzzy merges.

    Args:
        bundles: Loaded WorkbookBundle objects.
        config: RationalizationConfig (from Discovery Agent or user).
        kpi_result: KPI analysis from the KPI Agent (may be None).

    Returns:
        AgentResult whose ``output`` is a :class:`SourceAnalysisResult`.
    """
    decisions: list[AgentDecision] = []
    warnings: list[str] = []

    kpi_referenced: set[str] = set()
    if kpi_result is not None:
        kpi_referenced = kpi_result.referenced_canonicals

    # Build mapping of canonical → KPI labels using it
    canonical_to_kpi_labels: dict[str, list[str]] = {}
    if kpi_result is not None:
        for dep in kpi_result.dependencies:
            for canonical in dep.canonical_source_columns:
                canonical_to_kpi_labels.setdefault(canonical, [])
                if dep.kpi_label and dep.kpi_label not in canonical_to_kpi_labels[canonical]:
                    canonical_to_kpi_labels[canonical].append(dep.kpi_label)

    # Run source analysis with KPI refs for exclusion decisions
    try:
        source_result: SourceAnalysisResult = analyze_source_data(
            bundles, config, kpi_referenced_canonicals=kpi_referenced
        )
    except Exception as exc:
        logger.exception("SchemaAgent failed during source analysis")
        return AgentResult(
            agent_name="SchemaAgent",
            decisions=[],
            output=None,
            warnings=[f"Schema analysis failed: {exc}"],
        )

    # Build KPI-veto set: canonical names where two SIMILAR columns are each
    # used by a different KPI formula → never collapse them
    veto_set: set[str] = set()
    if kpi_result is not None:
        for profile in source_result.similar_columns:
            all_related = [profile.canonical_name] + profile.similar_to
            kpi_users: dict[str, list[str]] = {}   # canonical → [kpi_labels]
            for canonical in all_related:
                labels = canonical_to_kpi_labels.get(canonical, [])
                if labels:
                    kpi_users[canonical] = labels
            # If two or more related columns are used by DIFFERENT KPIs → veto
            all_kpi_labels = [lbl for lbls in kpi_users.values() for lbl in lbls]
            unique_kpi_labels = set(all_kpi_labels)
            if len(kpi_users) >= 2 and len(unique_kpi_labels) >= 2:
                for canonical in all_related:
                    veto_set.add(canonical)

    if veto_set:
        warnings.append(
            f"KPI-usage veto applied to {len(veto_set)} column(s): "
            f"{', '.join(sorted(veto_set)[:5])}{'…' if len(veto_set) > 5 else ''}. "
            "These columns scored above the fuzzy threshold but are used by "
            "different KPIs and will be kept separate."
        )

    # Generate one AgentDecision per column profile
    for profile in source_result.column_profiles:
        is_kpi_ref = profile.canonical_name in kpi_referenced
        veto = profile.canonical_name in veto_set
        kpi_labels = canonical_to_kpi_labels.get(profile.canonical_name, [])
        conf = _confidence_for_profile(profile, is_kpi_ref, veto)

        decisions.append(AgentDecision(
            subject=f"Column '{profile.canonical_name}' ({profile.present_in_count} workbook(s))",
            decision=_classify_decision(profile),
            confidence=conf,
            reasoning=_reasoning(profile, is_kpi_ref, veto, kpi_labels),
            signals={
                "match_class":         profile.match_class,
                "present_in_count":    profile.present_in_count,
                "present_in_workbooks": profile.present_in_workbooks,
                "similar_to":          profile.similar_to,
                "is_kpi_referenced":   is_kpi_ref,
                "kpi_labels":          kpi_labels,
                "kpi_veto_applied":    veto,
            },
        ))

    # Decision for each excluded column
    for wb, tab, col in source_result.excluded_columns:
        decisions.append(AgentDecision(
            subject=f"Column '{col}' in {wb}/{tab}",
            decision="EXCLUDE",
            confidence=0.95,
            reasoning=(
                f"Column '{col}' is not referenced by any KPI formula. "
                "Excluded from Master Source Data to keep the output clean."
            ),
            signals={"workbook": wb, "source_tab": tab, "original_col": col},
        ))

    logger.info(
        "SchemaAgent: %d profiles (%d common, %d similar, %d unique), "
        "%d excluded, %d veto'd",
        len(source_result.column_profiles),
        len(source_result.common_columns),
        len(source_result.similar_columns),
        len(source_result.unique_columns),
        len(source_result.excluded_columns),
        len(veto_set),
    )
    return AgentResult(
        agent_name="SchemaAgent",
        decisions=decisions,
        output=source_result,
        warnings=warnings,
    )
