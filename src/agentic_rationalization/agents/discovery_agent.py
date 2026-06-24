"""Phase 1 — Workbook Discovery Agent.

Wraps the existing workbook_analyzer classifier to produce structured
AgentDecision objects for every tab in every uploaded workbook.

Output: a RationalizationConfig skeleton ready for downstream agents,
plus a list of decisions explaining every tab assignment.
"""

from __future__ import annotations

import logging

from src.ingestion.workbook_analyzer import (
    TabType,
    WorkbookAnalysis,
    analyze_many,
)
from src.guided_rationalization.config import RationalizationConfig, WorkbookTabConfig
from src.agentic_rationalization.models import AgentDecision, AgentResult

logger = logging.getLogger(__name__)


def _get_pivot_infos_for_wb(analysis: WorkbookAnalysis):
    """Return a list of (pivot_sheet, source_sheets) pairs for this workbook."""
    return [
        tab for tab in analysis.tab_analyses if tab.is_pivot_sheet
    ]


# Confidence is derived from the score gap between the winning TabType and the
# runner-up.  A large gap → high confidence; a close contest → lower confidence.
_HIGH_CONF_GAP = 1.5       # winning score this much ahead of runner-up → ≥ 0.90
_MED_CONF_GAP  = 0.5       # → 0.70–0.89
# Tabs classified as these types are not interesting to downstream agents
_SKIP_TYPES = {TabType.UNKNOWN, TabType.REFERENCE_DATA, TabType.MAPPING_TABLE}

# Types that qualify a tab as a KPI/summary tab for rationalization
_KPI_TYPES = {TabType.KPI_SUMMARY, TabType.DASHBOARD, TabType.OUTPUT}


def _score_gap_to_confidence(signal_scores: dict[str, float], winner: str) -> float:
    """Convert the score gap between winner and runner-up into a 0–1 confidence."""
    if not signal_scores:
        return 0.5
    sorted_scores = sorted(signal_scores.values(), reverse=True)
    winner_score = signal_scores.get(winner, 0.0)
    runner_up = sorted_scores[1] if len(sorted_scores) > 1 else 0.0
    gap = winner_score - runner_up
    if gap >= _HIGH_CONF_GAP:
        conf = 0.90 + min((gap - _HIGH_CONF_GAP) / 5.0, 0.09)
    elif gap >= _MED_CONF_GAP:
        conf = 0.70 + (gap - _MED_CONF_GAP) / (_HIGH_CONF_GAP - _MED_CONF_GAP) * 0.20
    else:
        conf = 0.50 + (gap / _MED_CONF_GAP) * 0.20
    return round(min(max(conf, 0.50), 0.99), 3)


def _human_reasoning(tab_type: TabType, signals: dict) -> str:
    """Generate a plain-English explanation for a tab classification."""
    fd = signals.get("formula_density_pct", 0)
    rows = signals.get("row_count", 0)
    refs = signals.get("inter_sheet_refs", 0)
    agg = signals.get("agg_func_count", 0)

    if tab_type == TabType.SOURCE_DATA:
        return (
            f"Classified as source data: {rows:,} rows with low formula density "
            f"({fd:.1f}%) and {refs} inter-sheet references. "
            "This tab contains raw transactional records."
        )
    if tab_type == TabType.KPI_SUMMARY:
        return (
            f"Classified as KPI summary: formula density {fd:.1f}%, "
            f"{agg} aggregate function(s), {refs} inter-sheet reference(s). "
            "This tab produces computed KPI outputs from source data."
        )
    if tab_type == TabType.DASHBOARD:
        return (
            f"Classified as dashboard: references {refs} cells across multiple sheets, "
            f"formula density {fd:.1f}%. Suitable as KPI output tab."
        )
    if tab_type == TabType.OUTPUT:
        return (
            f"Classified as output tab: dense formulas ({fd:.1f}%) referencing "
            f"multiple sheets. Treated as KPI output."
        )
    if tab_type == TabType.CALCULATION:
        return (
            f"Classified as calculation helper: formula density {fd:.1f}%, "
            "intermediate calculations only — not a source or KPI tab."
        )
    if tab_type == TabType.MIXED:
        return (
            f"Classified as mixed: contains both raw data ({rows:,} rows) and "
            f"formulas ({fd:.1f}%). Manual review recommended."
        )
    return f"Classified as {tab_type.value}: insufficient signal to be more specific."


def run(
    bundles: list,
    overrides: dict[str, dict[str, str]] | None = None,
    analyses: "list[WorkbookAnalysis] | None" = None,
) -> AgentResult:
    """Analyse every bundle and return a RationalizationConfig + AgentDecisions.

    Args:
        bundles: Loaded WorkbookBundle objects.
        overrides: Optional user overrides mapping
            ``{workbook_name: {"source_tab": "...", "kpi_tabs": ["..."]}}``
            Any value here takes precedence over the agent's recommendation.
        analyses: Pre-computed WorkbookAnalysis list.  When supplied, the
            expensive ``analyze_many`` call is skipped — the caller is
            responsible for passing up-to-date analyses.  Pass ``None``
            (default) to always re-analyse.

    Returns:
        AgentResult whose ``output`` is a :class:`RationalizationConfig`.
    """
    overrides = overrides or {}
    if analyses is None:
        analyses = analyze_many(bundles)
    decisions: list[AgentDecision] = []
    warnings: list[str] = []
    wb_tab_configs: list[WorkbookTabConfig] = []

    for analysis in analyses:
        wb_name = analysis.workbook_name
        wb_override = overrides.get(wb_name, {})
        detection_type = analysis.detection_type()

        # Per-tab decisions
        tab_decisions: list[AgentDecision] = []
        for tab in analysis.tab_analyses:
            if tab.tab_type in _SKIP_TYPES:
                continue

            # Pivot-detected tabs get maximum confidence; signal-only tabs use gap heuristic
            if tab.is_pivot_sheet:
                conf = 0.99
            elif tab.tab_type == TabType.SOURCE_DATA and any(
                tab.tab_name in p.pivot_source_tabs
                for p in _get_pivot_infos_for_wb(analysis)
            ):
                conf = 0.97  # pivot-traced source tab → very high confidence
            else:
                conf = _score_gap_to_confidence(tab.signal_scores, tab.tab_type.value)

            signals = {
                "tab_type":              tab.tab_type.value,
                "row_count":             tab.row_count,
                "col_count":             tab.col_count,
                "formula_density_pct":   round(tab.formula_density * 100, 2),
                "inter_sheet_refs":      tab.inter_sheet_ref_count,
                "agg_func_count":        tab.formula_cell_count,
                "referenced_sheets":     tab.referenced_sheets,
                "is_pivot_sheet":        tab.is_pivot_sheet,
                "pivot_source_tabs":     tab.pivot_source_tabs,
                "pivot_value_fields":    [m.display_name for m in tab.pivot_value_fields],
                "detection_type":        detection_type,
            }
            # Use pivot-aware reasoning when applicable; fall back to signal-based
            if tab.classification_reason.startswith("PIVOT"):
                reasoning = tab.classification_reason
            else:
                reasoning = _human_reasoning(tab.tab_type, signals)

            tab_decisions.append(AgentDecision(
                subject=f"{wb_name} / {tab.tab_name}",
                decision=tab.tab_type.value,
                confidence=conf,
                reasoning=reasoning,
                signals=signals,
            ))

        decisions.extend(tab_decisions)

        # ── Source tab selection ─────────────────────────────────────────────
        # Priority: user override > pivot-traced source > signal-based winner
        source_candidates = [
            d for d in tab_decisions if d.decision == TabType.SOURCE_DATA.value
        ]
        source_candidates.sort(key=lambda d: d.confidence, reverse=True)

        if wb_override.get("source_tab"):
            chosen_source = wb_override["source_tab"]
        elif source_candidates:
            chosen_source = source_candidates[0].subject.split(" / ", 1)[-1]
        else:
            chosen_source = ""
            warnings.append(
                f"No source data tab found in '{wb_name}'. "
                "Please configure manually in Guided Rationalization."
            )

        # ── KPI tab selection ────────────────────────────────────────────────
        # Priority: user override > pivot KPI tabs > formula-based KPI tabs
        kpi_candidates = [
            d for d in tab_decisions if d.decision in {t.value for t in _KPI_TYPES}
        ]
        if wb_override.get("kpi_tabs"):
            chosen_kpi_tabs = list(wb_override["kpi_tabs"])
        else:
            chosen_kpi_tabs = [d.subject.split(" / ", 1)[-1] for d in kpi_candidates]

        if not chosen_kpi_tabs:
            warnings.append(
                f"No KPI/summary tabs found in '{wb_name}'. "
                "Workbook may not contain formula-driven outputs."
            )

        if chosen_source:
            wb_tab_configs.append(WorkbookTabConfig(
                workbook_name=wb_name,
                source_tab=chosen_source,
                kpi_tabs=chosen_kpi_tabs,
            ))

    config = RationalizationConfig(
        workbook_configs=wb_tab_configs,
        remove_unused_columns=True,
    )

    logger.info(
        "DiscoveryAgent: %d workbook(s), %d decisions, %d warnings",
        len(analyses), len(decisions), len(warnings),
    )
    return AgentResult(
        agent_name="DiscoveryAgent",
        decisions=decisions,
        output=config,
        warnings=warnings,
    )
