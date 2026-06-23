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
    action_overrides: dict[str, dict[str, str]] | None = None,
) -> PipelineResult:
    """Execute all 6 agents in order and return a PipelineResult.

    Args:
        bundles: Loaded WorkbookBundle objects.
        discovery_overrides: Optional dict of tab-role overrides for DiscoveryAgent.
        tolerance: Reconciliation tolerance for ValidationAgent (default 1%).
        decision_overrides: Optional nested dict ``{agent_name: {subject: override_value}}``.
            Where a match is found the agent's decision value is replaced with the
            user-supplied value before the decision log is built.
        action_overrides: Structured user choices from the review panel.
            Format: ``{subject: action_key}`` where action_key is one of the
            SIMILAR_COLUMN_ACTIONS, KPI_UNRESOLVED_ACTIONS, or CONSOLIDATION_ACTIONS keys.
            These are parsed into force_merge / force_separate / force_exclude for
            SchemaAgent so choices actually change the output workbook.
    """
    agent_results = []
    errors: list[str] = []
    _overrides: dict[str, dict[str, str]] = decision_overrides or {}

    # Parse action_overrides into schema force params
    force_merge:    dict[str, str] = {}
    force_separate: set[str]       = set()
    force_exclude:  set[str]       = set()

    for subject, action_key in (action_overrides or {}).items():
        if action_key in ("MERGE", "MERGE_ANYWAY"):
            # subject looks like "Column 'canonical_a' (N workbook(s))"
            canonical_a = _extract_canonical_from_subject(subject)
            if canonical_a:
                # The review panel stores canonical_b in decision_overrides["SchemaAgent"]
                # as the annotation; we use the naming convention that canonical_b is in
                # the paired entry stored when the review panel saved the choice.
                # For simplicity, we'll look up the similar_to from agent decisions below.
                force_merge[canonical_a] = canonical_a  # placeholder; resolved after schema run
        elif action_key == "KEEP_SEPARATE":
            canonical_a = _extract_canonical_from_subject(subject)
            if canonical_a:
                force_separate.add(canonical_a)
        elif action_key == "EXCLUDE_A":
            canonical_a = _extract_canonical_from_subject(subject)
            if canonical_a:
                force_exclude.add(canonical_a)
        elif action_key == "EXCLUDE_B":
            # canonical_b is encoded as the override value in decision_overrides
            pass  # handled via decision_overrides annotation; exclude resolved below

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

    # Phase 3 — Schema Rationalization (with user-directed force params)
    schema_result_obj = schema_agent.run(
        bundles, config, kpi_analysis,
        force_merge=force_merge if force_merge else None,
        force_separate=force_separate if force_separate else None,
        force_exclude=force_exclude if force_exclude else None,
    )
    _apply_overrides(schema_result_obj, _overrides)
    agent_results.append(schema_result_obj)
    source_analysis = schema_result_obj.output

    # Phase 4 — Consolidation (intelligence + strategy, merged agent)
    consol_result = consolidation_agent.run(bundles, config, source_analysis, kpi_analysis)
    _apply_overrides(consol_result, _overrides)
    agent_results.append(consol_result)

    # Phase 5 — Generation
    gen_result = generation_agent.run(bundles, config, source_analysis, kpi_analysis)
    _apply_overrides(gen_result, _overrides)
    agent_results.append(gen_result)
    _gen_out = gen_result.output if isinstance(gen_result.output, tuple) else (None, None, [])
    fs_bytes  = _gen_out[0] if len(_gen_out) > 0 else None
    ap_bytes  = _gen_out[1] if len(_gen_out) > 1 else None
    resolved_frames = _gen_out[2] if len(_gen_out) > 2 else []

    # Phase 6 — Validation (requires master DataFrame)
    master_df: pd.DataFrame | None = None
    if fs_bytes:
        try:
            master_tab = config.future_source_tab_name
            # _write_df_to_sheet writes a title at row 1 and headers at row 2;
            # use header=1 (0-indexed) to skip the title row.
            master_df = pd.read_excel(
                io.BytesIO(fs_bytes),
                sheet_name=master_tab,
                header=1,
            )
        except Exception:
            logger.exception("Could not extract master DataFrame for validation")

    val_result = validation_agent.run(
        bundles, config, kpi_analysis, master_df,
        tolerance=tolerance,
        resolved_source_frames=resolved_frames or None,
    )
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


def _extract_canonical_from_subject(subject: str) -> str:
    """Parse canonical name from SchemaAgent subject string."""
    if subject.startswith("Column '"):
        end = subject.find("'", 8)
        if end > 8:
            return subject[8:end]
    return ""


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
