"""Actionable decision classification for the agentic review panel.

Classifies AgentDecision objects from a PipelineResult into typed
ActionableDecision objects that the UI can render as concrete choices
(radio buttons, dropdowns) rather than plain text fields.

Three action types:
  SIMILAR_COLUMN    — SchemaAgent found two columns with similar names;
                      user decides whether to merge, keep separate, or exclude.
  KPI_UNRESOLVED    — KpiAgent could not resolve a formula to source columns;
                      user can skip or specify the source column manually.
  CONSOLIDATION_PAIR— ConsolidationAgent flagged a workbook pair as uncertain;
                      user decides whether to append or keep separate.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from src.agentic_rationalization.models import PipelineResult


# ── Available actions per type ────────────────────────────────────────────────

SIMILAR_COLUMN_ACTIONS: list[tuple[str, str]] = [
    ("ACCEPT",      "Accept agent recommendation"),
    ("MERGE",       "Merge — treat as identical columns (same business concept)"),
    ("KEEP_SEPARATE", "Keep separate — distinct business fields, do not merge"),
    ("EXCLUDE_A",   "Exclude column A from Master Source Data"),
    ("EXCLUDE_B",   "Exclude column B from Master Source Data"),
]

KPI_UNRESOLVED_ACTIONS: list[tuple[str, str]] = [
    ("ACCEPT",        "Accept as-is — agent will note unresolved dependency"),
    ("SKIP",          "Skip — do not include this formula in dependency mapping"),
    ("MANUAL_COLUMN", "Specify source column name manually (enter below)"),
]

CONSOLIDATION_ACTIONS: list[tuple[str, str]] = [
    ("ACCEPT",   "Accept agent recommendation"),
    ("APPEND",   "Consolidate — append source tabs into Master Source Data"),
    ("SEPARATE", "Keep separate — do not consolidate these workbooks"),
]


ActionType = Literal["SIMILAR_COLUMN", "KPI_UNRESOLVED", "CONSOLIDATION_PAIR"]


@dataclass
class ActionableDecision:
    """A single agent decision that requires explicit user choice."""

    action_type: ActionType
    agent_name: str
    subject: str                    # matches AgentDecision.subject exactly
    description: str                # plain-English explanation for the user
    current_recommendation: str     # what the agent decided
    confidence: float
    available_actions: list[tuple[str, str]]  # [(key, label), ...]

    # Column-level context (SIMILAR_COLUMN)
    canonical_a: str = ""
    canonical_b: str = ""           # similar-to canonical (may be empty)
    original_names_a: list[str] = field(default_factory=list)
    original_names_b: list[str] = field(default_factory=list)
    workbooks_a: list[str] = field(default_factory=list)
    workbooks_b: list[str] = field(default_factory=list)

    # KPI context (KPI_UNRESOLVED)
    formula: str = ""
    kpi_tab: str = ""
    cell_address: str = ""

    # Consolidation context (CONSOLIDATION_PAIR)
    workbook_pair: tuple[str, str] = ("", "")
    overlap_pct: float = 0.0
    grain_issue: str = ""

    # KPI labels affected by this decision
    kpi_labels: list[str] = field(default_factory=list)


def classify_actionable_decisions(pipeline: PipelineResult) -> list[ActionableDecision]:
    """Extract all decisions that need explicit user review from a PipelineResult.

    Returns a list of :class:`ActionableDecision` objects sorted by confidence
    (lowest first — most urgent review at the top).
    """
    actionable: list[ActionableDecision] = []

    # ── Source for original column names ─────────────────────────────────────
    schema_result_obj = pipeline.get_result("SchemaAgent")
    source_result = schema_result_obj.output if schema_result_obj else None

    # Build reverse mapping: canonical → [(wb, tab, original)] for lookup
    reverse_map: dict[str, list[tuple[str, str, str]]] = {}
    if source_result is not None:
        for (wb, tab, orig), canonical in source_result.column_mapping.items():
            reverse_map.setdefault(canonical, []).append((wb, tab, orig))

    # ── SchemaAgent — similar columns ────────────────────────────────────────
    schema_agent_result = pipeline.get_result("SchemaAgent")
    if schema_agent_result:
        for d in schema_agent_result.decisions:
            if not d.overridable:
                continue
            signals = d.signals
            if signals.get("match_class") != "similar":
                continue
            # Only surface decisions that are genuinely uncertain (< 0.90)
            if d.confidence >= 0.90:
                continue

            canonical_a = signals.get("canonical_name") or _extract_canonical(d.subject)
            similar_to: list[str] = signals.get("similar_to", [])
            canonical_b = similar_to[0] if similar_to else ""

            # Resolve original names via reverse_map or signals["source_entries"]
            entries_a: list[tuple[str, str, str]] = (
                signals.get("source_entries", []) or
                reverse_map.get(canonical_a, [])
            )
            entries_b: list[tuple[str, str, str]] = reverse_map.get(canonical_b, [])

            orig_a = sorted({orig for _, _, orig in entries_a})
            orig_b = sorted({orig for _, _, orig in entries_b})
            wbs_a  = sorted({wb for wb, _, _ in entries_a})
            wbs_b  = sorted({wb for wb, _, _ in entries_b})

            kpi_labels: list[str] = signals.get("kpi_labels", [])

            desc_parts = [
                f"Columns '{canonical_a}' and '{canonical_b}' have similar names "
                f"(confidence: {d.confidence:.0%})."
            ]
            if kpi_labels:
                desc_parts.append(f"Affected KPIs: {', '.join(kpi_labels)}.")
            if signals.get("kpi_veto_applied"):
                desc_parts.append(
                    "⚠ KPI veto was applied — these columns are used by different "
                    "KPI formulas, so the agent kept them separate."
                )

            # Build dynamic action list — if no canonical_b, can't exclude B
            actions = list(SIMILAR_COLUMN_ACTIONS)
            if not canonical_b:
                actions = [(k, l) for k, l in actions if k != "EXCLUDE_B"]

            actionable.append(ActionableDecision(
                action_type="SIMILAR_COLUMN",
                agent_name="SchemaAgent",
                subject=d.subject,
                description=" ".join(desc_parts),
                current_recommendation=d.decision,
                confidence=d.confidence,
                available_actions=actions,
                canonical_a=canonical_a,
                canonical_b=canonical_b,
                original_names_a=orig_a,
                original_names_b=orig_b,
                workbooks_a=wbs_a,
                workbooks_b=wbs_b,
                kpi_labels=kpi_labels,
            ))

    # ── KpiAgent — unresolved formulas ────────────────────────────────────────
    kpi_agent_result = pipeline.get_result("KpiAgent")
    if kpi_agent_result:
        for d in kpi_agent_result.decisions:
            if not d.overridable:
                continue
            signals = d.signals
            canonical_cols: list[str] = signals.get("canonical_source_cols", [])
            formula_type: str = signals.get("formula_type", "SOURCE_BACKED")
            # Only flag SOURCE_BACKED formulas that couldn't resolve their source columns.
            # DERIVED/ROLLUP/VALIDATION formulas inherit lineage via tracing — never unresolved.
            if formula_type != "SOURCE_BACKED":
                continue
            if d.confidence > 0.65:
                continue

            parts = d.subject.split(" / ", maxsplit=2)
            kpi_tab_name = parts[1] if len(parts) > 1 else ""
            cell_addr    = parts[2] if len(parts) > 2 else d.subject
            formula      = signals.get("formula", "")
            refs_only    = signals.get("refs_source_tab_only", True)

            if not canonical_cols:
                desc = (
                    f"Formula in {d.subject} could not be resolved to any source column. "
                    "The formula may reference a tab not configured as source data."
                )
            else:
                desc = (
                    f"Formula in {d.subject} references sheet(s) outside the configured "
                    "source tab — partial mapping only. "
                    f"Resolved columns: {', '.join(canonical_cols)}."
                )

            actionable.append(ActionableDecision(
                action_type="KPI_UNRESOLVED",
                agent_name="KpiAgent",
                subject=d.subject,
                description=desc,
                current_recommendation=d.decision,
                confidence=d.confidence,
                available_actions=KPI_UNRESOLVED_ACTIONS,
                formula=formula,
                kpi_tab=kpi_tab_name,
                cell_address=cell_addr,
                kpi_labels=[d.decision],
            ))

    # ── ConsolidationAgent — uncertain workbook pairs ────────────────────────
    consol_result = pipeline.get_result("ConsolidationAgent")
    if consol_result:
        for d in consol_result.decisions:
            if not d.overridable:
                continue
            if not d.subject.startswith("Pair:"):
                continue
            if d.confidence >= 0.80:
                continue

            signals = d.signals
            overlap = signals.get("column_overlap_pct", 0.0)
            grain_msg = signals.get("grain_issue", "")

            # Extract workbook names from subject "Pair: WbA ↔ WbB"
            pair_part = d.subject.removeprefix("Pair: ")
            pair_names = [p.strip() for p in pair_part.split("↔")]
            wb_a = pair_names[0] if len(pair_names) > 0 else ""
            wb_b = pair_names[1] if len(pair_names) > 1 else ""

            desc_parts = [
                f"Workbooks '{wb_a}' and '{wb_b}' have {overlap:.1f}% column overlap.",
            ]
            if grain_msg:
                desc_parts.append(f"Grain issue: {grain_msg}")
            else:
                desc_parts.append(
                    "Low overlap makes automatic consolidation uncertain — "
                    "KPI formulas may not work correctly if these are appended."
                )

            actionable.append(ActionableDecision(
                action_type="CONSOLIDATION_PAIR",
                agent_name="ConsolidationAgent",
                subject=d.subject,
                description=" ".join(desc_parts),
                current_recommendation=d.decision,
                confidence=d.confidence,
                available_actions=CONSOLIDATION_ACTIONS,
                workbook_pair=(wb_a, wb_b),
                overlap_pct=overlap,
                grain_issue=grain_msg,
            ))

    # Sort by confidence ascending (most uncertain first)
    actionable.sort(key=lambda a: a.confidence)
    return actionable


def _extract_canonical(subject: str) -> str:
    """Extract canonical name from subject like \"Column 'my_col' (2 workbook(s))\"."""
    if subject.startswith("Column '"):
        end = subject.find("'", 8)
        if end > 8:
            return subject[8:end]
    return subject
