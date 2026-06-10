"""Rationalization module — identify redundant reports and recommend actions."""

from src.rationalization.rationalizer import (
    ReportRecommendation,
    ReportStatus,
    assess_redundancy,
    build_inventory,
)

__all__ = [
    "ReportStatus",
    "ReportRecommendation",
    "assess_redundancy",
    "build_inventory",
]
