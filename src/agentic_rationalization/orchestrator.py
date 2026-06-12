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
) -> PipelineResult:
    """Execute all 6 agents in order and return a PipelineResult.

    Args:
        bundles: Loaded WorkbookBundle objects.
        discovery_overrides: Optional dict of tab-role overrides for DiscoveryAgent.
        tolerance: Reconciliation tolerance for ValidationAgent (default 1%).
    """
    agent_results = []
    errors: list[str] = []

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
    agent_results.append(kpi_result_obj)
    kpi_analysis = kpi_result_obj.output

    # Phase 3 — Schema Rationalization
    schema_result_obj = schema_agent.run(bundles, config, kpi_analysis)
    agent_results.append(schema_result_obj)
    source_analysis = schema_result_obj.output

    # Phase 4 — Consolidation Strategy
    consol_result = consolidation_agent.run(bundles, config, source_analysis, kpi_analysis)
    agent_results.append(consol_result)

    # Phase 5 — Generation
    gen_result = generation_agent.run(bundles, config, source_analysis, kpi_analysis)
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


def _build_log(agent_results: list) -> bytes | None:
    try:
        return decision_log.build(agent_results)
    except Exception as exc:
        logger.warning("Could not build decision log: %s", exc)
        return None
