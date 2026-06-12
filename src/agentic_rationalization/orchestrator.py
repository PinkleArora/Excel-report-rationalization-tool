"""Agentic Rationalization Orchestrator.

Runs all agents sequentially and wires outputs between them.
Returns a PipelineResult with every agent's decisions + final workbook bytes.
"""

from __future__ import annotations

import io
import logging

import pandas as pd

from src.agentic_rationalization.models import PipelineResult
from src.agentic_rationalization import decision_log
from src.agentic_rationalization.agents import (
    discovery_agent,
    kpi_agent,
    schema_agent,
    consolidation_agent,
    generation_agent,
    validation_agent,
)

logger = logging.getLogger(__name__)


def run_pipeline(
    bundles: list,
    discovery_overrides: dict | None = None,
    tolerance: float = 0.01,
    decision_overrides: dict[str, dict[str, str]] | None = None,
) -> PipelineResult:
    """Execute all 6 agents in order and return a PipelineResult.

    Args:
        bundles: Loaded WorkbookBundle objects.
        discovery_overrides: Optional dict of tab-role overrides for DiscoveryAgent.
        tolerance: Reconciliation tolerance for ValidationAgent (default 1%).
        decision_overrides: Optional nested dict ``{agent_name: {subject: override_value}}``.
            Where a match is found the agent's decision value is replaced with the
            user-supplied value before the decision log is built.
    """
    agent_results = []
    errors: list[str] = []
    _overrides: dict[str, dict[str, str]] = decision_overrides or {}

    # Phase 1 — Discovery
    discovery_result = discovery_agent.run(bundles, overrides=discovery_overrides or {})
    agent_results.append(discovery_result)
    config = discovery_result.output
    if config is None:
        errors.append("DiscoveryAgent failed — cannot continue pipeline.")
        return PipelineResult(
            agent_results=agent_results,
            future_state_bytes=None,
            analysis_pack_bytes=None,
            decision_log_bytes=_build_log(agent_results),
            errors=errors,
        )

    # Phase 2 — KPI Understanding
    kpi_result_obj = kpi_agent.run(bundles, config)
    _apply_overrides(kpi_result_obj, _overrides)
    agent_results.append(kpi_result_obj)
    kpi_analysis = kpi_result_obj.output

    # Phase 3 — Schema Rationalization
    schema_result_obj = schema_agent.run(bundles, config, kpi_analysis)
    _apply_overrides(schema_result_obj, _overrides)
    agent_results.append(schema_result_obj)
    source_analysis = schema_result_obj.output

    # Phase 4 — Consolidation Strategy
    consol_result = consolidation_agent.run(bundles, config, source_analysis, kpi_analysis)
    _apply_overrides(consol_result, _overrides)
    agent_results.append(consol_result)

    # Phase 5 — Generation
    gen_result = generation_agent.run(bundles, config, source_analysis, kpi_analysis)
    _apply_overrides(gen_result, _overrides)
    agent_results.append(gen_result)
    fs_bytes, ap_bytes = gen_result.output if isinstance(gen_result.output, tuple) else (None, None)

    # Phase 6 — Validation (requires master DataFrame)
    master_df: pd.DataFrame | None = None
    if fs_bytes:
        try:
            from openpyxl import load_workbook
            wb_check = load_workbook(io.BytesIO(fs_bytes), data_only=True)
            master_tab = config.future_source_tab_name
            if master_tab in wb_check.sheetnames:
                ws = wb_check[master_tab]
                rows = ws.values
                headers = next(rows, None)
                if headers:
                    master_df = pd.DataFrame(rows, columns=headers)
        except Exception as exc:
            logger.warning("Could not extract master DataFrame for validation: %s", exc)

    val_result = validation_agent.run(bundles, config, kpi_analysis, master_df, tolerance=tolerance)
    _apply_overrides(val_result, _overrides)
    agent_results.append(val_result)

    # Build decision log
    log_bytes = _build_log(agent_results)

    logger.info(
        "Pipeline complete: %d agents, %d errors",
        len(agent_results),
        len(errors),
    )
    return PipelineResult(
        agent_results=agent_results,
        future_state_bytes=fs_bytes,
        analysis_pack_bytes=ap_bytes,
        decision_log_bytes=log_bytes,
        errors=errors,
    )


def _apply_overrides(
    result: "AgentResult",
    overrides: dict[str, dict[str, str]],
) -> None:
    """Mutate agent decisions in-place where the user supplied an override value.

    Matches on ``agent_name`` + ``decision.subject``.  The original agent
    decision is preserved in ``decision.signals["agent_decision"]`` so the
    decision log can show both values.
    """
    from src.agentic_rationalization.models import AgentResult  # local to avoid circular
    agent_overrides = overrides.get(result.agent_name, {})
    if not agent_overrides:
        return
    for d in result.decisions:
        if d.subject in agent_overrides:
            override_val = agent_overrides[d.subject]
            if override_val and override_val != d.decision:
                d.signals["agent_decision"] = d.decision
                d.decision = f"[USER OVERRIDE] {override_val}"
                d.reasoning = (
                    f"User overrode agent recommendation '{d.signals['agent_decision']}' "
                    f"with '{override_val}'. Original reasoning: {d.reasoning}"
                )
                logger.info(
                    "Override applied: %s / %s → %s",
                    result.agent_name, d.subject, override_val,
                )


def _build_log(agent_results: list) -> bytes | None:
    try:
        return decision_log.build(agent_results)
    except Exception as exc:
        logger.warning("Could not build decision log: %s", exc)
        return None
