"""Backwards-compatibility re-export shim.

The Consolidation Intelligence Agent has been merged into consolidation_agent.py.
This module re-exports everything from there so existing imports keep working.
"""

from src.agentic_rationalization.agents.consolidation_agent import (
    FileProfile,
    FileContribution,
    ConsolidationGroup,
    ColumnDiscardRec,
    DiscardAnalysis,
    FinalRecommendation,
    KpiAlignment,
    WorkbookKpiStats,
    ConsolidationIntelligenceResult,
    _build_union_find,
    _find,
    _union,
    _build_groups,
    _compute_kpi_alignments,
    compute_file_group_compatibility,
    run,
)

__all__ = [
    "FileProfile",
    "FileContribution",
    "ConsolidationGroup",
    "ColumnDiscardRec",
    "DiscardAnalysis",
    "FinalRecommendation",
    "KpiAlignment",
    "WorkbookKpiStats",
    "ConsolidationIntelligenceResult",
    "_build_union_find",
    "_find",
    "_union",
    "_build_groups",
    "_compute_kpi_alignments",
    "compute_file_group_compatibility",
    "run",
]
