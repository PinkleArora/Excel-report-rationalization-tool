"""Define and evaluate standardised KPIs across consolidated data."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

import pandas as pd


@dataclass
class KPIDefinition:
    name: str
    description: str
    formula: str  # human-readable expression, e.g. "revenue / headcount"
    unit: str
    compute: Callable[[pd.DataFrame], pd.Series] | None = None


class KPIRegistry:
    """Central store of :class:`KPIDefinition` objects."""

    def __init__(self) -> None:
        self._kpis: dict[str, KPIDefinition] = {}

    def register(self, kpi: KPIDefinition) -> None:
        self._kpis[kpi.name] = kpi

    def get(self, name: str) -> KPIDefinition:
        return self._kpis[name]

    def names(self) -> list[str]:
        return list(self._kpis.keys())


def evaluate_kpis(
    df: pd.DataFrame,
    registry: KPIRegistry,
    kpi_names: list[str] | None = None,
) -> pd.DataFrame:
    """Compute all (or selected) KPIs against *df*.

    Args:
        df: Consolidated master dataset.
        registry: Populated :class:`KPIRegistry`.
        kpi_names: Subset of KPI names to evaluate; ``None`` evaluates all.

    Returns:
        DataFrame with one column per evaluated KPI.
    """
    raise NotImplementedError
