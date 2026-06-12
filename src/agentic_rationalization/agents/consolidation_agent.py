"""Phase 4 — Consolidation Strategy Agent.

Decides whether source datasets can be safely appended into a single
Master Source Data tab, and explains why.

Criteria evaluated:
1. Schema overlap — do the workbooks share enough KPI-required columns?
2. Row grain compatibility — do numeric columns have compatible distributions?
3. Common KPI column coverage — fraction of KPI-needed columns that are common.
4. Structural conflicts — duplicate canonicals that could not be auto-resolved.
"""

from __future__ import annotations

import logging
from collections import defaultdict

import pandas as pd

from src.guided_rationalization.config import RationalizationConfig
from src.guided_rationalization.kpi_analyzer import KpiAnalysisResult
from src.guided_rationalization.source_analyzer import SourceAnalysisResult
from src.agentic_rationalization.models import AgentDecision, AgentResult

logger = logging.getLogger(__name__)

# Fraction of KPI-referenced columns that must be common for APPEND recommendation
_APPEND_COMMON_THRESHOLD = 0.5
# Minimum column overlap fraction between any two workbooks to even consider appending
_MIN_OVERLAP = 0.3


def _column_overlap(cols_a: set[str], cols_b: set[str]) -> float:
    if not cols_a or not cols_b:
        return 0.0
    return len(cols_a & cols_b) / len(cols_a | cols_b)


def _numeric_distribution_compatible(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    common_cols: list[str],
) -> tuple[bool, str]:
    """Check whether numeric columns in two frames have broadly compatible ranges."""
    issues = []
    for col in common_cols:
        if col not in df_a.columns or col not in df_b.columns:
            continue
        s_a = pd.to_numeric(df_a[col], errors="coerce").dropna()
        s_b = pd.to_numeric(df_b[col], errors="coerce").dropna()
        if s_a.empty or s_b.empty:
            continue
        # Flag if one frame's max is > 100× the other's max (order-of-magnitude mismatch)
        max_a, max_b = s_a.abs().max(), s_b.abs().max()
        if max_a > 0 and max_b > 0 and (max_a / max_b > 100 or max_b / max_a > 100):
            issues.append(
                f"'{col}' has very different magnitudes "
                f"(max {max_a:,.0f} vs {max_b:,.0f}) — possible grain mismatch."
            )
    return (len(issues) == 0), "; ".join(issues)


def run(
    bundles: list,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult | None,
    kpi_result: KpiAnalysisResult | None,
) -> AgentResult:
    """Evaluate whether source tabs can be appended and explain the reasoning.

    Returns:
        AgentResult whose ``output`` is a dict::

            {
                "recommendation": "APPEND" | "SEPARATE" | "MANUAL_REVIEW",
                "confidence":      float,
                "reason":          str,
                "per_pair":        list[dict],  # workbook-pair level detail
            }
    """
    decisions: list[AgentDecision] = []
    warnings: list[str] = []

    if source_result is None or kpi_result is None:
        return AgentResult(
            agent_name="ConsolidationAgent",
            decisions=[],
            output={"recommendation": "MANUAL_REVIEW",
                    "confidence": 0.0,
                    "reason": "Upstream analysis unavailable.",
                    "per_pair": []},
            warnings=["Source or KPI analysis was not available — cannot evaluate consolidation."],
        )

    active_bundles = [b for b in bundles if config.config_for(b.file_name) is not None]
    if len(active_bundles) < 2:
        result = {
            "recommendation": "N/A",
            "confidence": 1.0,
            "reason": "Only one workbook configured — no consolidation needed.",
            "per_pair": [],
        }
        decisions.append(AgentDecision(
            subject="Consolidation strategy",
            decision="N/A",
            confidence=1.0,
            reasoning=result["reason"],
            signals={"workbook_count": len(active_bundles)},
        ))
        return AgentResult(
            agent_name="ConsolidationAgent",
            decisions=decisions,
            output=result,
            warnings=warnings,
        )

    kpi_canonicals = kpi_result.referenced_canonicals
    common_kpi_canonicals = {
        p.canonical_name for p in source_result.common_columns
        if p.canonical_name in kpi_canonicals
    }
    total_kpi_canonicals = len(kpi_canonicals) or 1
    common_kpi_fraction = len(common_kpi_canonicals) / total_kpi_canonicals

    # Per-workbook source column sets (canonical names after mapping)
    wb_canonicals: dict[str, set[str]] = defaultdict(set)
    for (wb, tab, orig_col), canonical in source_result.column_mapping.items():
        wb_canonicals[wb].add(canonical)

    # Evaluate all workbook pairs
    wb_names = list(wb_canonicals.keys())
    per_pair: list[dict] = []
    pair_verdicts: list[str] = []
    grain_issues: list[str] = []

    for i, wb_a in enumerate(wb_names):
        for wb_b in wb_names[i + 1:]:
            cols_a = wb_canonicals[wb_a]
            cols_b = wb_canonicals[wb_b]
            overlap = _column_overlap(cols_a, cols_b)

            # Check numeric distribution compatibility on common columns
            bundle_a = next((b for b in bundles if b.file_name == wb_a), None)
            bundle_b = next((b for b in bundles if b.file_name == wb_b), None)
            cfg_a = config.config_for(wb_a)
            cfg_b = config.config_for(wb_b)
            grain_ok, grain_msg = True, ""
            if bundle_a and bundle_b and cfg_a and cfg_b:
                df_a = bundle_a.sheets.get(cfg_a.source_tab, pd.DataFrame())
                df_b = bundle_b.sheets.get(cfg_b.source_tab, pd.DataFrame())
                common_cols = list(cols_a & cols_b)
                grain_ok, grain_msg = _numeric_distribution_compatible(df_a, df_b, common_cols)

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
                    + (f"Grain compatible. " if grain_ok else f"Grain warning: {grain_msg} ")
                    + f"Verdict: {verdict}."
                ),
                signals={
                    "column_overlap_pct": round(overlap * 100, 1),
                    "grain_compatible":   grain_ok,
                    "common_kpi_fraction": round(common_kpi_fraction, 3),
                },
            ))

    # Overall recommendation
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
