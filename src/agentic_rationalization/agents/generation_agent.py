"""Phase 5 — Workbook Generation Agent.

Thin wrapper around the existing build_rationalized_workbook_pair.
Records which tab names were chosen and why.
"""

from __future__ import annotations

import logging

from src.guided_rationalization.config import RationalizationConfig
from src.guided_rationalization.kpi_analyzer import KpiAnalysisResult
from src.guided_rationalization.source_analyzer import SourceAnalysisResult
from src.guided_rationalization.workbook_builder import (
    DuplicateColumnError,
    build_rationalized_workbook_pair,
)
from src.agentic_rationalization.models import AgentDecision, AgentResult

logger = logging.getLogger(__name__)


def run(
    bundles: list,
    config: RationalizationConfig,
    source_result: SourceAnalysisResult | None,
    kpi_result: KpiAnalysisResult | None,
) -> AgentResult:
    """Generate the two output workbooks and record tab-naming decisions.

    Returns:
        AgentResult whose ``output`` is ``(future_state_bytes, analysis_pack_bytes)``
        or ``(None, diagnostic_bytes)`` on a blocking duplicate error.
    """
    decisions: list[AgentDecision] = []
    warnings: list[str] = []

    if source_result is None or kpi_result is None:
        return AgentResult(
            agent_name="GenerationAgent",
            decisions=[],
            output=(None, None),
            warnings=["Cannot generate workbook: upstream analysis unavailable."],
        )

    # Document master source data tab name decision
    decisions.append(AgentDecision(
        subject="Master Source Data tab name",
        decision=config.future_source_tab_name,
        confidence=1.0,
        reasoning=(
            f"Using user-specified name '{config.future_source_tab_name}' "
            "for the consolidated source data worksheet."
        ),
        signals={"tab_name": config.future_source_tab_name},
        overridable=False,
    ))

    # Document KPI output tab name decisions
    for wb_cfg in config.workbook_configs:
        for kpi_tab in wb_cfg.kpi_tabs:
            output_name = config.kpi_tab_name_for(wb_cfg.workbook_name, kpi_tab)
            has_explicit = config.kpi_tab_has_explicit_name(wb_cfg.workbook_name)
            decisions.append(AgentDecision(
                subject=f"KPI output tab: {wb_cfg.workbook_name} / {kpi_tab}",
                decision=output_name,
                confidence=1.0 if has_explicit else 0.80,
                reasoning=(
                    f"Output tab named '{output_name}' "
                    + ("(user-specified)." if has_explicit
                       else f"(auto-generated from workbook stem and original tab name '{kpi_tab}').")
                ),
                signals={
                    "source_workbook":  wb_cfg.workbook_name,
                    "original_kpi_tab": kpi_tab,
                    "output_tab_name":  output_name,
                    "user_defined":     has_explicit,
                },
                overridable=False,
            ))

    # Call the existing builder
    fs_bytes: bytes | None = None
    ap_bytes: bytes | None = None

    try:
        fs_bytes, ap_bytes = build_rationalized_workbook_pair(
            bundles, config, source_result, kpi_result
        )
        decisions.append(AgentDecision(
            subject="Workbook generation",
            decision="SUCCESS",
            confidence=1.0,
            reasoning="Both output workbooks generated successfully.",
            signals={},
            overridable=False,
        ))
    except DuplicateColumnError as dup_exc:
        ap_bytes = dup_exc.diagnostic_bytes
        blocking = [r for r in dup_exc.resolutions if r.scenario == "C"]
        msg = (
            f"{len(blocking)} unresolvable duplicate column(s) require manual review. "
            + "; ".join(
                f"{r.workbook}/{r.source_tab}: '{r.canonical_name}' "
                f"used by KPIs {r.kpis_using}"
                for r in blocking
            )
        )
        warnings.append(msg)
        decisions.append(AgentDecision(
            subject="Workbook generation",
            decision="BLOCKED",
            confidence=1.0,
            reasoning=msg,
            signals={"blocking_count": len(blocking)},
            overridable=False,
        ))
        logger.warning("GenerationAgent blocked: %s", msg)
    except Exception as exc:
        warnings.append(f"Generation failed: {exc}")
        logger.exception("GenerationAgent failed")

    return AgentResult(
        agent_name="GenerationAgent",
        decisions=decisions,
        output=(fs_bytes, ap_bytes),
        warnings=warnings,
    )
