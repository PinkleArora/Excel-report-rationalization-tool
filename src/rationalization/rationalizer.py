"""Identify redundant reports and generate a rationalisation recommendation."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

import pandas as pd


class ReportStatus(str, Enum):
    KEEP = "keep"
    MERGE = "merge"
    RETIRE = "retire"
    REVIEW = "review"


@dataclass
class ReportRecommendation:
    report_name: str
    status: ReportStatus
    rationale: str
    merge_target: str | None = None
    duplicate_columns: list[str] = field(default_factory=list)


def assess_redundancy(
    profiles: dict,
    matches: list,
) -> list[ReportRecommendation]:
    """Evaluate profiling and match data to recommend keep / merge / retire per report.

    Args:
        profiles: Output from :mod:`src.profiling.profiler`.
        matches: Output from :mod:`src.schema_matching.matcher`.

    Returns:
        One :class:`ReportRecommendation` per report.
    """
    raise NotImplementedError


def build_inventory(recommendations: list[ReportRecommendation]) -> pd.DataFrame:
    """Convert recommendations to a tidy inventory DataFrame for display or export."""
    raise NotImplementedError
