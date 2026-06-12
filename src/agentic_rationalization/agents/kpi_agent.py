"""Phase 2 — KPI Understanding Agent.

Wraps the existing kpi_analyzer to extract structured KPI descriptions
with plain-English explanations of what each formula measures.
"""

from __future__ import annotations

import logging
import re

from src.guided_rationalization.config import RationalizationConfig
from src.guided_rationalization.kpi_analyzer import (
    KpiAnalysisResult,
    KpiDependency,
    analyze_kpi_dependencies,
)
from src.guided_rationalization.source_analyzer import analyze_source_data
from src.agentic_rationalization.models import AgentDecision, AgentResult

logger = logging.getLogger(__name__)

# Functions that aggregate numeric data across many rows
_NUMERIC_AGGS = {"SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "AVERAGE", "AVERAGEIF", "AVERAGEIFS"}
_COUNT_AGGS   = {"COUNT", "COUNTA", "COUNTIF", "COUNTIFS"}
_LOOKUP_FUNCS = {"VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "XLOOKUP"}
_EXTREMA_FUNCS = {"MAX", "MIN", "MEDIAN", "LARGE", "SMALL"}

# Filter condition patterns inside SUMIFS / COUNTIFS
_CRITERIA_RE = re.compile(r'"([^"]*)"')


def _describe_kpi(dep: KpiDependency) -> str:
    """Generate a plain-English description of what a KPI formula measures."""
    func = dep.aggregate_function.upper()
    cols = dep.canonical_source_columns
    col_str = ", ".join(f"'{c}'" for c in cols) if cols else "unknown source columns"

    if func in _NUMERIC_AGGS:
        criteria = _CRITERIA_RE.findall(dep.formula)
        if criteria:
            filter_str = "; ".join(f'filter: "{c}"' for c in criteria[:3])
            return (
                f"Aggregates {func} of {col_str} "
                f"with conditions ({filter_str})."
            )
        return f"Sums / aggregates {col_str} across all rows."

    if func in _COUNT_AGGS:
        criteria = _CRITERIA_RE.findall(dep.formula)
        if criteria:
            filter_str = "; ".join(f'"{c}"' for c in criteria[:3])
            return f"Counts rows in {col_str} matching {filter_str}."
        return f"Counts non-empty rows in {col_str}."

    if func in _LOOKUP_FUNCS:
        return (
            f"Looks up a value in {col_str} using {func}. "
            "This is a point lookup, not an aggregation."
        )

    if func in _EXTREMA_FUNCS:
        return f"Finds the {func.lower()} value in {col_str}."

    if func == "OTHER":
        return (
            f"Complex formula referencing {col_str}. "
            "Exact measure requires manual inspection."
        )

    return f"Uses {func} on {col_str}."


def _kpi_confidence(dep: KpiDependency) -> float:
    """Confidence that we correctly identified the KPI's source columns."""
    if not dep.canonical_source_columns:
        return 0.40   # formula found but no source columns resolved
    if not dep.refs_source_tab_only:
        return 0.65   # references tabs we don't fully control
    if dep.aggregate_function == "OTHER":
        return 0.70   # formula parsed but function unknown
    return 0.92       # clean, resolvable dependency


def run(
    bundles: list,
    config: RationalizationConfig,
) -> AgentResult:
    """Extract KPI dependencies and annotate each with a plain-English description.

    Args:
        bundles: Loaded WorkbookBundle objects.
        config: RationalizationConfig from the Discovery Agent (or user).

    Returns:
        AgentResult whose ``output`` is a :class:`KpiAnalysisResult`.
    """
    decisions: list[AgentDecision] = []
    warnings: list[str] = []

    # Need a column mapping first (pass 1 — no KPI refs yet)
    try:
        source_result_pass1 = analyze_source_data(bundles, config)
        kpi_result: KpiAnalysisResult = analyze_kpi_dependencies(
            bundles, config, source_result_pass1.column_mapping
        )
    except Exception as exc:
        logger.exception("KpiAgent failed during analysis")
        return AgentResult(
            agent_name="KpiAgent",
            decisions=[],
            output=None,
            warnings=[f"KPI analysis failed: {exc}"],
        )

    for dep in kpi_result.dependencies:
        conf = _kpi_confidence(dep)
        description = _describe_kpi(dep)
        decisions.append(AgentDecision(
            subject=f"{dep.workbook_name} / {dep.kpi_tab} / {dep.cell_address}",
            decision=dep.kpi_label or dep.cell_address,
            confidence=conf,
            reasoning=description,
            signals={
                "formula":                dep.formula,
                "aggregate_function":     dep.aggregate_function,
                "canonical_source_cols":  dep.canonical_source_columns,
                "refs_source_tab_only":   dep.refs_source_tab_only,
            },
        ))

    if not kpi_result.dependencies:
        warnings.append(
            "No KPI formulas were extracted. "
            "Verify that the uploaded files are .xlsx (not .csv) and were not "
            "saved in data-only mode."
        )

    unresolved = [
        d for d in kpi_result.dependencies if not d.canonical_source_columns
    ]
    if unresolved:
        warnings.append(
            f"{len(unresolved)} KPI formula(s) could not be resolved to source "
            "columns — they may reference tabs not configured as source data."
        )

    logger.info(
        "KpiAgent: %d dependencies, %d referenced canonicals, %d unresolved",
        len(kpi_result.dependencies),
        len(kpi_result.referenced_canonicals),
        len(unresolved),
    )
    return AgentResult(
        agent_name="KpiAgent",
        decisions=decisions,
        output=kpi_result,
        warnings=warnings,
    )
