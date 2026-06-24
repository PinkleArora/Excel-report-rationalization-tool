"""Pipeline phase functions — standalone entry points for the step-by-step UI.

Each function runs exactly one phase of the agentic pipeline and returns an
AgentResult.  The existing run_pipeline() in orchestrator.py calls these in
sequence to preserve full backward compatibility.
"""

from __future__ import annotations

import io
import logging

import pandas as pd

from src.agentic_rationalization.models import AgentResult
from src.agentic_rationalization.agents import (
    discovery_agent,
    kpi_agent,
    schema_agent,
    consolidation_agent,
    generation_agent,
    validation_agent,
)

logger = logging.getLogger(__name__)


def run_discovery_phase(
    bundles: list,
    config=None,  # config unused — discovery produces the config; kept for API symmetry
    overrides: dict | None = None,
    analyses: list | None = None,
) -> AgentResult:
    """Phase 1: Workbook discovery — classify tabs and build RationalizationConfig.

    Args:
        bundles: Loaded WorkbookBundle objects.
        config: Ignored (discovery produces the config). Kept for a consistent signature.
        overrides: Optional ``{workbook_name: {"source_tab": "...", "kpi_tabs": [...]}}``
            overrides for the discovery agent.
        analyses: Pre-computed WorkbookAnalysis objects.  When provided the
            expensive ``analyze_many`` scan is skipped entirely; the caller
            caches these between override-only re-runs to eliminate redundant
            workbook re-scanning.

    Returns:
        AgentResult whose ``output`` is a RationalizationConfig.
    """
    return discovery_agent.run(bundles, overrides=overrides or {}, analyses=analyses)


def run_consolidation_phase(
    bundles: list,
    config,
    kpi_result=None,
    source_result=None,
) -> AgentResult:
    """Phase 2: Grain detection, pairwise compatibility, grouping, and discard analysis.

    Args:
        bundles: Loaded WorkbookBundle objects.
        config: RationalizationConfig produced by Phase 1.
        kpi_result: Optional KpiAnalysisResult from Phase 4 (may be None on first pass).
        source_result: Optional SourceAnalysisResult from Phase 3 (may be None on first pass).

    Returns:
        AgentResult (agent_name="ConsolidationAgent") whose ``output`` contains
        both intelligence (ConsolidationIntelligenceResult) and strategy fields.
    """
    return consolidation_agent.run(bundles, config, source_result, kpi_result)


def run_schema_phase(
    bundles: list,
    config,
    kpi_result,
    group_file_names: list[str],
    force_merge: dict[str, str] | None = None,
    force_separate: set[str] | None = None,
    force_exclude: set[str] | None = None,
) -> AgentResult:
    """Phase 3: Schema rationalization for a specific group of files.

    Only the bundles whose file_name appears in *group_file_names* are processed.
    A filtered subset of *config* is constructed so only those workbooks are analysed.

    Args:
        bundles: All loaded WorkbookBundle objects (filtered internally).
        config: Full RationalizationConfig.
        kpi_result: KpiAnalysisResult from Phase 4.
        group_file_names: The file names belonging to the selected consolidation group.
        force_merge: Optional map of canonical_a → canonical_b to force-merge.
        force_separate: Optional set of canonicals to keep separate.
        force_exclude: Optional set of canonicals to exclude.

    Returns:
        AgentResult (agent_name="SchemaAgent") whose ``output`` is a SourceAnalysisResult.
    """
    from src.guided_rationalization.config import RationalizationConfig

    # Filter bundles and config to the selected group
    selected_bundles = [b for b in bundles if b.file_name in group_file_names]
    selected_cfgs = [
        cfg for cfg in config.workbook_configs
        if cfg.workbook_name in group_file_names
    ]
    filtered_config = RationalizationConfig(
        workbook_configs=selected_cfgs,
        future_source_tab_name=config.future_source_tab_name,
        future_kpi_tab_names=config.future_kpi_tab_names,
        matching_threshold=config.matching_threshold,
        merge_high_confidence=config.merge_high_confidence,
        remove_unused_columns=config.remove_unused_columns,
    )

    return schema_agent.run(
        selected_bundles,
        filtered_config,
        kpi_result,
        force_merge=force_merge,
        force_separate=force_separate,
        force_exclude=force_exclude,
    )


def run_kpi_phase(
    bundles: list,
    config,
) -> AgentResult:
    """Phase 4: KPI dependency analysis — parse formulas and map to source columns.

    Args:
        bundles: Loaded WorkbookBundle objects.
        config: RationalizationConfig produced by Phase 1.

    Returns:
        AgentResult (agent_name="KpiAgent") whose ``output`` is a KpiAnalysisResult.
    """
    return kpi_agent.run(bundles, config)


def run_validation_phase(
    bundles: list,
    config,
    kpi_result,
    master_df: "pd.DataFrame | None",
    tolerance: float = 0.01,
    resolved_source_frames=None,
) -> AgentResult:
    """Phase 5: Validation — reconcile numeric KPI column totals.

    Args:
        bundles: Loaded WorkbookBundle objects.
        config: RationalizationConfig.
        kpi_result: KpiAnalysisResult.
        master_df: Master Source Data DataFrame extracted from the generated workbook.
        tolerance: Acceptable fractional difference (default 1%).
        resolved_source_frames: Optional list of resolved source DataFrames from generation.

    Returns:
        AgentResult (agent_name="ValidationAgent") whose decisions are per-column.
    """
    return validation_agent.run(
        bundles,
        config,
        kpi_result,
        master_df,
        tolerance=tolerance,
        resolved_source_frames=resolved_source_frames or None,
    )


def run_generation_phase(
    bundles: list,
    config,
    source_result,
    kpi_result,
) -> AgentResult:
    """Phase 6: Workbook generation — build Future State Workbook and Analysis Pack.

    Args:
        bundles: Loaded WorkbookBundle objects.
        config: RationalizationConfig.
        source_result: SourceAnalysisResult from Phase 3.
        kpi_result: KpiAnalysisResult from Phase 4.

    Returns:
        AgentResult (agent_name="GenerationAgent") whose ``output`` is a
        ``(future_state_bytes, analysis_pack_bytes, resolved_frames)`` tuple.
    """
    return generation_agent.run(bundles, config, source_result, kpi_result)
