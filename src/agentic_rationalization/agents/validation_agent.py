"""Phase 6 — Validation Agent.

Compares numeric column totals between the original source tabs and the
generated Master Source Data to confirm KPI inputs are preserved.

Uses the existing reconciliation/reconciler.py machinery.
"""

from __future__ import annotations

import logging

import pandas as pd

from src.guided_rationalization.config import RationalizationConfig
from src.guided_rationalization.kpi_analyzer import KpiAnalysisResult
from src.reconciliation.reconciler import ReconciliationResult, reconcile
from src.schema_matching.normalizer import normalize_column_name
from src.agentic_rationalization.models import AgentDecision, AgentResult

logger = logging.getLogger(__name__)


def _status_reasoning(result: ReconciliationResult) -> str:
    if result.status == "pass":
        return (
            f"Column '{result.column}': source total {result.source_total:,.2f}, "
            f"master total {result.output_total:,.2f}. "
            f"Variance {result.delta_pct:.4f}% ≤ tolerance {result.tolerance}%. PASS."
        )
    if result.status == "warn":
        return (
            f"Column '{result.column}': variance {result.delta_pct:.4f}% exceeds "
            f"tolerance {result.tolerance}% but within warning threshold. "
            f"Delta: {result.delta:+,.2f}. WARN — review recommended."
        )
    return (
        f"Column '{result.column}': variance {result.delta_pct:.4f}% exceeds "
        f"tolerance {result.tolerance}%. Source total {result.source_total:,.2f}, "
        f"master total {result.output_total:,.2f}, delta {result.delta:+,.2f}. FAIL."
    )


def _confidence_for_status(status: str) -> float:
    return {"pass": 0.99, "warn": 0.80, "fail": 0.99}.get(status, 0.90)


def run(
    bundles: list,
    config: RationalizationConfig,
    kpi_result: KpiAnalysisResult | None,
    master_df: pd.DataFrame | None,
    tolerance: float = 0.01,
) -> AgentResult:
    """Reconcile numeric totals: original source tabs vs Master Source Data.

    Args:
        bundles: Original loaded bundles.
        config: RationalizationConfig used during generation.
        kpi_result: KPI analysis (to know which columns to reconcile).
        master_df: The generated Master Source Data DataFrame.
        tolerance: Acceptable relative difference (default 1%).

    Returns:
        AgentResult whose ``output`` is a list of :class:`ReconciliationResult`.
    """
    decisions: list[AgentDecision] = []
    warnings: list[str] = []

    if master_df is None or kpi_result is None:
        return AgentResult(
            agent_name="ValidationAgent",
            decisions=[],
            output=[],
            warnings=["Validation skipped — master data or KPI result unavailable."],
        )

    # Collect original source frames
    source_frames: list[pd.DataFrame] = []
    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        df = bundle.sheets.get(wb_cfg.source_tab)
        if df is not None:
            source_frames.append(df)

    if not source_frames:
        return AgentResult(
            agent_name="ValidationAgent",
            decisions=[],
            output=[],
            warnings=["No source frames found for validation."],
        )

    # Columns to reconcile: KPI-referenced canonicals that are numeric
    numeric_cols = [
        c for c in kpi_result.referenced_canonicals
        if c in master_df.columns
        and pd.to_numeric(master_df[c], errors="coerce").notna().any()
    ]

    if not numeric_cols:
        warnings.append(
            "No numeric KPI-referenced columns found in Master Source Data. "
            "Validation skipped."
        )
        return AgentResult(
            agent_name="ValidationAgent",
            decisions=decisions,
            output=[],
            warnings=warnings,
        )

    results: list[ReconciliationResult] = reconcile(
        source_frames=source_frames,
        output_frame=master_df,
        numeric_columns=numeric_cols,
        tolerance=tolerance,
    )

    fails = [r for r in results if r.status == "fail"]
    warns = [r for r in results if r.status == "warn"]
    passes = [r for r in results if r.status == "pass"]

    for r in results:
        decisions.append(AgentDecision(
            subject=f"Reconciliation: '{r.column}'",
            decision=r.status.upper(),
            confidence=_confidence_for_status(r.status),
            reasoning=_status_reasoning(r),
            signals={
                "source_total":   r.source_total,
                "master_total":   r.output_total,
                "delta":          r.delta,
                "delta_pct":      r.delta_pct,
                "tolerance_pct":  r.tolerance,
            },
        ))

    if fails:
        warnings.append(
            f"{len(fails)} column(s) FAILED reconciliation: "
            + ", ".join(r.column for r in fails)
        )
    if warns:
        warnings.append(
            f"{len(warns)} column(s) have reconciliation WARNINGs: "
            + ", ".join(r.column for r in warns)
        )

    summary_decision = "ALL_PASS" if not fails and not warns else (
        "HAS_FAILURES" if fails else "HAS_WARNINGS"
    )
    decisions.insert(0, AgentDecision(
        subject="Overall validation",
        decision=summary_decision,
        confidence=0.99,
        reasoning=(
            f"Reconciled {len(numeric_cols)} numeric KPI column(s): "
            f"{len(passes)} PASS, {len(warns)} WARN, {len(fails)} FAIL."
        ),
        signals={
            "total_columns": len(results),
            "pass_count":    len(passes),
            "warn_count":    len(warns),
            "fail_count":    len(fails),
        },
        overridable=False,
    ))

    logger.info(
        "ValidationAgent: %d pass, %d warn, %d fail across %d columns",
        len(passes), len(warns), len(fails), len(results),
    )
    return AgentResult(
        agent_name="ValidationAgent",
        decisions=decisions,
        output=results,
        warnings=warnings,
    )
