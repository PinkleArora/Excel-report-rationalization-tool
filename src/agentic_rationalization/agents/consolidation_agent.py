"""Consolidation Agent — strategy decisions + grain detection, pairwise scoring, and grouping.

Merges the logic from the former consolidation_intelligence_agent.py into a single agent.

Phase A — Intelligence:
  1. Grain detection (via grain_detector.py) on each file's source DataFrame
  2. Pairwise compatibility scoring (via compatibility_scorer.py)
  3. Union-find grouping into ConsolidationGroups
  4. Column overlap stats per group
  5. Discard recommendations

Phase B — Strategy:
  6. Evaluates whether source tabs can be appended (APPEND / MANUAL_REVIEW)
  7. Checks KPI column coverage fraction and numeric distribution compatibility

Criteria evaluated for strategy:
1. Schema overlap — do the workbooks share enough KPI-required columns?
2. Row grain compatibility — do numeric columns have compatible distributions?
3. Common KPI column coverage — fraction of KPI-needed columns that are common.
4. Structural conflicts — duplicate canonicals that could not be auto-resolved.
"""

from __future__ import annotations

import itertools
import logging
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

try:
    from rapidfuzz import fuzz as _fuzz
    _RAPIDFUZZ_AVAILABLE = True
except ImportError:
    _RAPIDFUZZ_AVAILABLE = False

from src.guided_rationalization.config import RationalizationConfig
from src.guided_rationalization.kpi_analyzer import KpiAnalysisResult
from src.guided_rationalization.source_analyzer import SourceAnalysisResult
from src.agentic_rationalization.models import AgentDecision, AgentResult
from src.agentic_rationalization.grain_detector import detect_grain, GrainDetectionResult
from src.agentic_rationalization.compatibility_scorer import score_compatibility, CompatibilityScore
from src.schema_matching.normalizer import normalize_column_name

logger = logging.getLogger(__name__)

_APPEND_COMMON_THRESHOLD = 0.5
_MIN_OVERLAP = 0.3

_TIME_KEYWORDS = {
    "date", "period", "month", "year", "quarter",
    "effective_date", "as_of_date", "asofdate", "effectivedate",
}


# ---------------------------------------------------------------------------
# Dataclasses (from former consolidation_intelligence_agent)
# ---------------------------------------------------------------------------

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
class FileContribution:
    file_name: str
    total_columns: int
    kpi_columns: int
    common_columns_contributed: int
    unique_columns_contributed: int
    discardable_columns: int


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
    recommendation: str
    expected_master_columns: int
    reasoning: str
    incompatibility_detail: str = ""
    file_contributions: list = field(default_factory=list)  # list[FileContribution]


@dataclass
class ColumnDiscardRec:
    canonical_name: str
    used_by_kpi: bool
    kpi_names: list[str]
    recommendation: str
    confidence: str
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
class KpiAlignment:
    canonical_kpi: str
    kpi_labels: list[str]
    per_file_match: dict  # file_name → matched raw column name ("" = not found)
    per_file_confidence: dict  # file_name → confidence (1.0 = exact, <1.0 = fuzzy)
    status: str  # "aligned" | "partial" | "missing"
    suggestion: str


@dataclass
class ConsolidationIntelligenceResult:
    file_profiles: list[FileProfile]
    pairwise_scores: list[CompatibilityScore]
    groups: list[ConsolidationGroup]
    discard_analysis: DiscardAnalysis
    final_recommendation: FinalRecommendation
    visual_map: str
    kpi_alignments: list = field(default_factory=list)  # list[KpiAlignment]


# ---------------------------------------------------------------------------
# Strategy helpers
# ---------------------------------------------------------------------------

def _column_overlap(cols_a: set[str], cols_b: set[str]) -> float:
    if not cols_a or not cols_b:
        return 0.0
    return len(cols_a & cols_b) / len(cols_a | cols_b)


def _numeric_distribution_compatible(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    common_cols: list[str],
) -> tuple[bool, str]:
    issues = []
    for col in common_cols:
        if col not in df_a.columns or col not in df_b.columns:
            continue
        s_a = pd.to_numeric(df_a[col], errors="coerce").dropna()
        s_b = pd.to_numeric(df_b[col], errors="coerce").dropna()
        if s_a.empty or s_b.empty:
            continue
        max_a, max_b = s_a.abs().max(), s_b.abs().max()
        if max_a > 0 and max_b > 0 and (max_a / max_b > 100 or max_b / max_a > 100):
            issues.append(
                f"'{col}' has very different magnitudes "
                f"(max {max_a:,.0f} vs {max_b:,.0f}) — possible grain mismatch."
            )
    return (len(issues) == 0), "; ".join(issues)


# ---------------------------------------------------------------------------
# Union-Find helpers
# ---------------------------------------------------------------------------

def _build_union_find(files: list[str]) -> dict[str, str]:
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
    kpi_referenced: set | None = None,
) -> list[ConsolidationGroup]:
    parent = _build_union_find(files)

    for sc in pairwise:
        if sc.overall_score >= 0.50 and sc.grain_compatible:
            _union(parent, sc.file_a, sc.file_b)

    clusters: dict[str, list[str]] = {}
    for f in files:
        root = _find(parent, f)
        clusters.setdefault(root, []).append(f)

    _kpi_referenced: set[str] = kpi_referenced or set()
    if not _kpi_referenced and kpi_analysis is not None:
        _kpi_referenced = getattr(kpi_analysis, "referenced_canonicals", set())

    # Build lookup for pairwise scores
    pair_score_map: dict[tuple[str, str], CompatibilityScore] = {}
    for sc in pairwise:
        pair_score_map[(sc.file_a, sc.file_b)] = sc
        pair_score_map[(sc.file_b, sc.file_a)] = sc

    groups: list[ConsolidationGroup] = []
    for group_id, (root, members) in enumerate(sorted(clusters.items()), start=1):
        normed_by_file: dict[str, set[str]] = {}
        for fname in members:
            normed_by_file[fname] = {normalize_column_name(c) for c in col_map.get(fname, [])}

        all_cols_list = list(normed_by_file.values())
        union_cols: set[str] = set().union(*all_cols_list) if all_cols_list else set()
        common_cols: set[str] = all_cols_list[0].copy() if all_cols_list else set()
        for s in all_cols_list[1:]:
            common_cols &= s

        kpi_required = len(union_cols & _kpi_referenced)
        discardable = len(union_cols - _kpi_referenced - common_cols)

        grain_cat = grain_map[root].grain_category if root in grain_map else "unknown"
        grain_counts = Counter(grain_map[f].grain_category for f in members if f in grain_map)
        if grain_counts:
            grain_cat = grain_counts.most_common(1)[0][0]

        recommendation = "Consolidate" if len(members) > 1 else "Keep Separate"

        reasoning = (
            f"{len(members)} file(s) share '{grain_cat}' grain. "
            f"Common columns: {len(common_cols)}, total union columns: {len(union_cols)}. "
            f"KPI-required: {kpi_required}, discardable: {discardable}."
        )

        # Incompatibility detail for standalone files
        incompatibility_detail = ""
        if len(members) == 1 and files:
            fname = members[0]
            grain_info = grain_map.get(fname)
            grain_label = grain_info.grain_label if grain_info else "Unknown"
            reasons = []
            for other in files:
                if other == fname:
                    continue
                sc = pair_score_map.get((fname, other))
                if sc is None:
                    continue
                if not sc.grain_compatible:
                    other_grain = grain_map.get(other)
                    other_label = other_grain.grain_label if other_grain else sc.grain_b
                    reasons.append(
                        f"• **Grain mismatch with `{other}`**: this file is *{grain_label}* "
                        f"while `{other}` is *{other_label}*. "
                        f"Consolidating would mix {grain_cat}-level and {sc.grain_b}-level rows, "
                        f"making aggregation unreliable."
                    )
                elif sc.overall_score < 0.50:
                    reasons.append(
                        f"• **Low compatibility with `{other}`** (score {sc.overall_score:.0%}): "
                        f"{sc.reasoning}"
                    )
            incompatibility_detail = "\n".join(reasons) if reasons else ""

        # Per-file contributions
        file_contributions = []
        for fname in members:
            fnormed = normed_by_file[fname]
            # unique = in this file but not in any other file in the group
            others_union: set[str] = set()
            for f2, cols2 in normed_by_file.items():
                if f2 != fname:
                    others_union |= cols2
            unique_to_file = fnormed - others_union
            common_in_file = fnormed & common_cols
            kpi_in_file = len(fnormed & _kpi_referenced)
            discardable_in_file = len(fnormed - _kpi_referenced)

            file_contributions.append(FileContribution(
                file_name=fname,
                total_columns=len(fnormed),
                kpi_columns=kpi_in_file,
                common_columns_contributed=len(common_in_file),
                unique_columns_contributed=len(unique_to_file),
                discardable_columns=discardable_in_file,
            ))

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
            incompatibility_detail=incompatibility_detail,
            file_contributions=file_contributions,
        ))

    return groups


def _compute_kpi_alignments(
    file_names: list[str],
    col_map: dict[str, list[str]],
    kpi_referenced: set[str],
    kpi_label_map: dict[str, list[str]],
) -> list[KpiAlignment]:
    """Compute per-file KPI column alignment using exact canonical matching + fuzzy fallback."""
    alignments: list[KpiAlignment] = []

    # Build per-file normalized column map
    normed_by_file: dict[str, dict[str, str]] = {}  # file → {canonical → raw_col}
    for fname in file_names:
        mapping: dict[str, str] = {}
        for raw_col in col_map.get(fname, []):
            canonical = normalize_column_name(raw_col)
            mapping[canonical] = raw_col
        normed_by_file[fname] = mapping

    for canonical in sorted(kpi_referenced):
        labels = kpi_label_map.get(canonical, [])
        per_file_match: dict[str, str] = {}
        per_file_confidence: dict[str, float] = {}

        for fname in file_names:
            mapping = normed_by_file.get(fname, {})
            if canonical in mapping:
                per_file_match[fname] = mapping[canonical]
                per_file_confidence[fname] = 1.0
            else:
                # Fuzzy match against all raw columns
                best_col = ""
                best_score = 0.0
                if _RAPIDFUZZ_AVAILABLE:
                    for raw_col in col_map.get(fname, []):
                        score = _fuzz.token_sort_ratio(canonical, normalize_column_name(raw_col)) / 100.0
                        if score > best_score:
                            best_score = score
                            best_col = raw_col
                if best_score >= 0.70:
                    per_file_match[fname] = best_col
                    per_file_confidence[fname] = best_score
                else:
                    per_file_match[fname] = ""
                    per_file_confidence[fname] = 0.0

        files_with_match = [f for f in file_names if per_file_match.get(f)]
        if len(files_with_match) == len(file_names):
            status = "aligned"
            suggestion = "All files cover this KPI column."
        elif files_with_match:
            missing = [f for f in file_names if not per_file_match.get(f)]
            suggestion = f"Present in {len(files_with_match)} of {len(file_names)} files. Missing from: {', '.join(missing)}."
            status = "partial"
        else:
            status = "missing"
            suggestion = "KPI column not found in any file."

        alignments.append(KpiAlignment(
            canonical_kpi=canonical,
            kpi_labels=labels,
            per_file_match=per_file_match,
            per_file_confidence=per_file_confidence,
            status=status,
            suggestion=suggestion,
        ))

    return alignments


# ---------------------------------------------------------------------------
# Main run()
# ---------------------------------------------------------------------------

def run(
    bundles: list,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult | None = None,
    kpi_result: KpiAnalysisResult | None = None,
) -> AgentResult:
    """Run the merged Consolidation Agent.

    Phase A: grain detection, pairwise scoring, grouping, discard analysis.
    Phase B: APPEND / MANUAL_REVIEW recommendation (requires source_result + kpi_result).

    Returns AgentResult whose output is a dict::

        {
            "intelligence": ConsolidationIntelligenceResult,
            "recommendation": str,
            "confidence": float,
            "reason": str,
            "per_pair": list[dict],
        }
    """
    decisions: list[AgentDecision] = []
    warnings: list[str] = []

    # ── Phase A: Intelligence ──────────────────────────────────────────────

    kpi_referenced: set[str] = set()
    kpi_label_map: dict[str, list[str]] = {}
    if kpi_result is not None:
        kpi_referenced = getattr(kpi_result, "referenced_canonicals", set())
        for dep in getattr(kpi_result, "dependencies", []):
            for canonical in getattr(dep, "canonical_source_columns", []):
                kpi_label_map.setdefault(canonical, [])
                label = getattr(dep, "kpi_label", "")
                if label and label not in kpi_label_map[canonical]:
                    kpi_label_map[canonical].append(label)

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
            warnings.append(
                f"Bundle '{file_name}' has no sheet '{source_tab}'; skipping grain detection."
            )
            continue

        cols = list(df.columns)
        col_map[file_name] = cols
        grain = detect_grain(df, kpi_result)
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

    file_names = list(col_map.keys())
    pairwise_scores: list[CompatibilityScore] = []

    for fa, fb in itertools.combinations(file_names, 2):
        sc = score_compatibility(
            fa, col_map[fa], grain_map[fa],
            fb, col_map[fb], grain_map[fb],
            kpi_result,
        )
        pairwise_scores.append(sc)
        decisions.append(AgentDecision(
            subject=f"{fa} vs {fb}",
            decision=sc.recommendation,
            confidence=sc.recommendation_confidence,
            reasoning=sc.reasoning,
            signals={"overall_score": sc.overall_score, "grain_compatible": sc.grain_compatible},
        ))

    groups = _build_groups(file_names, pairwise_scores, grain_map, col_map, kpi_result, kpi_referenced)

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

    consolidate_files: list[str] = []
    separate_files: list[str] = []
    for group in groups:
        if group.recommendation == "Consolidate":
            consolidate_files.extend(group.file_names)
        else:
            separate_files.extend(group.file_names)

    consolidate_count = len([g for g in groups if g.recommendation == "Consolidate"])
    if consolidate_count > 0:
        intel_summary = f"{consolidate_count} group(s) recommended for consolidation."
        intel_reasoning = (
            "Files sharing the same grain and overlapping KPI columns can be merged "
            "into a single master source, reducing duplication and maintenance overhead."
        )
    else:
        intel_summary = "No files are candidates for consolidation based on current analysis."
        intel_reasoning = (
            "Files operate at different grain levels or have insufficient column overlap "
            "to justify consolidation."
        )

    map_lines: list[str] = []
    for group in groups:
        icon = "+" if group.recommendation == "Consolidate" else "-"
        map_lines.append(f"[{icon}] Group {group.group_id}: {group.group_label}")
        for fname in group.file_names:
            map_lines.append(f"    |-- {fname}")
    visual_map = "\n".join(map_lines)

    final_rec = FinalRecommendation(
        summary=intel_summary,
        business_reasoning=intel_reasoning,
        consolidate_files=consolidate_files,
        separate_model_files=separate_files,
        exclude_files=[],
    )

    kpi_alignments = _compute_kpi_alignments(file_names, col_map, kpi_referenced, kpi_label_map)

    intelligence_result = ConsolidationIntelligenceResult(
        file_profiles=file_profiles,
        pairwise_scores=pairwise_scores,
        groups=groups,
        discard_analysis=discard_analysis,
        final_recommendation=final_rec,
        visual_map=visual_map,
        kpi_alignments=kpi_alignments,
    )

    # ── Phase B: Strategy ──────────────────────────────────────────────────

    if source_result is None or kpi_result is None:
        warnings.append(
            "Source or KPI analysis was not available — cannot evaluate consolidation strategy."
        )
        output = {
            "intelligence": intelligence_result,
            "recommendation": "MANUAL_REVIEW",
            "confidence": 0.0,
            "reason": "Upstream analysis unavailable.",
            "per_pair": [],
        }
        return AgentResult(
            agent_name="ConsolidationAgent",
            decisions=decisions,
            output=output,
            warnings=warnings,
        )

    active_bundles = [b for b in bundles if config.config_for(b.file_name) is not None]

    if len(active_bundles) < 2:
        result = {
            "recommendation": "N/A",
            "confidence": 1.0,
            "reason": "Only one workbook configured — no consolidation needed.",
            "per_pair": [],
        }
        decisions.insert(0, AgentDecision(
            subject="Consolidation strategy",
            decision="N/A",
            confidence=1.0,
            reasoning=result["reason"],
            signals={"workbook_count": len(active_bundles)},
        ))
        output = {"intelligence": intelligence_result, **result}
        return AgentResult(
            agent_name="ConsolidationAgent",
            decisions=decisions,
            output=output,
            warnings=warnings,
        )

    kpi_canonicals = kpi_result.referenced_canonicals
    common_kpi_canonicals = {
        p.canonical_name for p in source_result.common_columns
        if p.canonical_name in kpi_canonicals
    }
    total_kpi_canonicals = len(kpi_canonicals) or 1
    common_kpi_fraction = len(common_kpi_canonicals) / total_kpi_canonicals

    wb_canonicals: dict[str, set[str]] = defaultdict(set)
    for (wb, tab, orig_col), canonical in source_result.column_mapping.items():
        wb_canonicals[wb].add(canonical)

    wb_names = list(wb_canonicals.keys())
    per_pair: list[dict] = []
    pair_verdicts: list[str] = []
    grain_issues: list[str] = []

    for i, wb_a in enumerate(wb_names):
        for wb_b in wb_names[i + 1:]:
            cols_a = wb_canonicals[wb_a]
            cols_b = wb_canonicals[wb_b]
            overlap = _column_overlap(cols_a, cols_b)

            bundle_a = next((b for b in bundles if b.file_name == wb_a), None)
            bundle_b = next((b for b in bundles if b.file_name == wb_b), None)
            cfg_a = config.config_for(wb_a)
            cfg_b = config.config_for(wb_b)
            grain_ok, grain_msg = True, ""
            if bundle_a and bundle_b and cfg_a and cfg_b:
                df_a = bundle_a.sheets.get(cfg_a.source_tab, pd.DataFrame())
                df_b = bundle_b.sheets.get(cfg_b.source_tab, pd.DataFrame())
                common_cols_strat = list(cols_a & cols_b)
                grain_ok, grain_msg = _numeric_distribution_compatible(df_a, df_b, common_cols_strat)

            if grain_msg:
                grain_issues.append(f"{wb_a} ↔ {wb_b}: {grain_msg}")

            verdict = "APPEND" if (overlap >= _MIN_OVERLAP and grain_ok) else "MANUAL_REVIEW"
            pair_verdicts.append(verdict)
            per_pair.append({
                "workbook_a": wb_a,
                "workbook_b": wb_b,
                "column_overlap_pct": round(overlap * 100, 1),
                "grain_compatible": grain_ok,
                "grain_issue": grain_msg,
                "verdict": verdict,
            })
            decisions.append(AgentDecision(
                subject=f"Pair: {wb_a} ↔ {wb_b}",
                decision=verdict,
                confidence=0.85 if grain_ok else 0.60,
                reasoning=(
                    f"Column overlap: {overlap*100:.1f}% of combined schema. "
                    + ("Grain compatible. " if grain_ok else f"Grain warning: {grain_msg} ")
                    + f"Verdict: {verdict}."
                ),
                signals={
                    "column_overlap_pct": round(overlap * 100, 1),
                    "grain_compatible":   grain_ok,
                    "common_kpi_fraction": round(common_kpi_fraction, 3),
                },
            ))

    if all(v == "APPEND" for v in pair_verdicts) and common_kpi_fraction >= _APPEND_COMMON_THRESHOLD:
        overall = "APPEND"
        conf = 0.85 + min(common_kpi_fraction * 0.1, 0.10)
        reason = (
            f"All workbook pairs have sufficient column overlap and compatible row grain. "
            f"{len(common_kpi_canonicals)} of {len(kpi_canonicals)} KPI-required columns "
            f"({common_kpi_fraction*100:.0f}%) are common across workbooks. "
            "Recommend appending source tabs into a single Master Source Data."
        )
    elif grain_issues:
        overall = "MANUAL_REVIEW"
        conf = 0.70
        reason = (
            "Grain compatibility issues detected between workbooks: "
            + "; ".join(grain_issues)
            + ". Manual review required before consolidation."
        )
        warnings.extend(grain_issues)
    else:
        overall = "MANUAL_REVIEW"
        conf = 0.65
        reason = (
            f"Only {common_kpi_fraction*100:.0f}% of KPI-required columns are common "
            f"across workbooks (threshold: {_APPEND_COMMON_THRESHOLD*100:.0f}%). "
            "Some workbook pairs have low schema overlap. "
            "Review the per-pair analysis before deciding."
        )

    decisions.insert(0, AgentDecision(
        subject="Overall consolidation strategy",
        decision=overall,
        confidence=round(conf, 3),
        reasoning=reason,
        signals={
            "workbook_count":         len(active_bundles),
            "common_kpi_fraction":    round(common_kpi_fraction, 3),
            "common_kpi_canonicals":  len(common_kpi_canonicals),
            "total_kpi_canonicals":   len(kpi_canonicals),
            "pair_count":             len(per_pair),
        },
    ))

    output = {
        "intelligence": intelligence_result,
        "recommendation": overall,
        "confidence": round(conf, 3),
        "reason": reason,
        "per_pair": per_pair,
    }
    logger.info(
        "ConsolidationAgent: recommendation=%s conf=%.2f common_kpi_fraction=%.2f",
        overall, conf, common_kpi_fraction,
    )
    return AgentResult(
        agent_name="ConsolidationAgent",
        decisions=decisions,
        output=output,
        warnings=warnings,
    )
