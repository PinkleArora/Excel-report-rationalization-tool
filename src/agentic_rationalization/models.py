"""Shared dataclasses for the agentic rationalization pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class AgentDecision:
    """One discrete decision made by an agent, with supporting evidence."""

    subject: str            # e.g. "SQL_data tab in workbook_a.xlsx"
    decision: str           # e.g. "SOURCE_DATA"
    confidence: float       # 0.0–1.0
    reasoning: str          # plain-English explanation for the user
    signals: dict[str, Any] = field(default_factory=dict)   # raw numeric evidence
    overridable: bool = True   # False for decisions the user cannot meaningfully change


@dataclass
class AgentResult:
    """Everything an agent returns: decisions, a typed output payload, and warnings."""

    agent_name: str
    decisions: list[AgentDecision]
    output: Any             # agent-specific payload (RationalizationConfig, bytes, …)
    warnings: list[str] = field(default_factory=list)

    @property
    def low_confidence_decisions(self) -> list[AgentDecision]:
        return [d for d in self.decisions if d.confidence < 0.70]

    @property
    def medium_confidence_decisions(self) -> list[AgentDecision]:
        return [d for d in self.decisions if 0.70 <= d.confidence < 0.90]

    @property
    def high_confidence_decisions(self) -> list[AgentDecision]:
        return [d for d in self.decisions if d.confidence >= 0.90]


@dataclass
class PipelineResult:
    """Everything produced by a complete agentic pipeline run."""

    agent_results: list[AgentResult]
    future_state_bytes: bytes | None        # Future_State_Workbook.xlsx
    analysis_pack_bytes: bytes | None       # Rationalization_Analysis_Pack.xlsx
    decision_log_bytes: bytes | None        # Agent_Decision_Log.xlsx
    errors: list[str] = field(default_factory=list)

    @property
    def succeeded(self) -> bool:
        return self.future_state_bytes is not None and not self.errors

    def get_result(self, agent_name: str) -> AgentResult | None:
        return next((r for r in self.agent_results if r.agent_name == agent_name), None)
