"""Identify redundant reports and generate rationalisation recommendations."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from enum import Enum

import pandas as pd

from src.schema_matching.matcher import ColumnMatch

logger = logging.getLogger(__name__)

# Thresholds for recommendation scoring
_RETIRE_THRESHOLD = 0.80   # >= 80 % column overlap → retire
_MERGE_THRESHOLD = 0.50    # >= 50 % column overlap → merge
_REVIEW_THRESHOLD = 0.25   # >= 25 % column overlap → review


class ReportStatus(str, Enum):
    KEEP = "keep"
    MERGE = "merge"
    RETIRE = "retire"
    REVIEW = "review"


@dataclass
class ReportRecommendation:
    """Recommendation for a single source workbook."""

    report_name: str
    sheet_count: int
    total_columns: int
    matched_columns: int
    overlap_pct: float               # 0–100
    status: ReportStatus
    rationale: str
    merge_target: str | None = None
    duplicate_columns: list[str] = field(default_factory=list)


def assess_redundancy(
    bundles: list,
    matches: list[ColumnMatch],
) -> list[ReportRecommendation]:
    """Evaluate match data to recommend keep / merge / retire per workbook.

    Overlap percentage = (columns in this workbook that have at least one match
    in any *other* workbook) / (total columns in this workbook).

    Args:
        bundles: :class:`~src.ingestion.loader.WorkbookBundle` objects.
        matches: Column matches from :func:`~src.schema_matching.matcher.match_all_bundles`.

    Returns:
        One :class:`ReportRecommendation` per bundle, sorted by overlap descending.
    """
    # Build a quick lookup: workbook_name → set of source_col names that have matches
    matched_cols_by_wb: dict[str, set[str]] = {b.file_name: set() for b in bundles}
    best_partner: dict[str, tuple[str, int]] = {}  # workbook → (best_match_partner, count)

    for m in matches:
        matched_cols_by_wb.setdefault(m.source_workbook, set()).add(m.source_col)
        # Also count reverse direction
        matched_cols_by_wb.setdefault(m.target_workbook, set()).add(m.target_col)

    # Count partner match frequency to identify the best merge target
    partner_counts: dict[str, dict[str, int]] = {b.file_name: {} for b in bundles}
    for m in matches:
        partner_counts[m.source_workbook][m.target_workbook] = (
            partner_counts[m.source_workbook].get(m.target_workbook, 0) + 1
        )
        partner_counts[m.target_workbook][m.source_workbook] = (
            partner_counts[m.target_workbook].get(m.source_workbook, 0) + 1
        )

    recommendations: list[ReportRecommendation] = []

    for bundle in bundles:
        wb_name = bundle.file_name
        total_cols = sum(len(df.columns) for df in bundle.sheets.values())
        sheet_count = len(bundle.sheets)
        matched = len(matched_cols_by_wb.get(wb_name, set()))
        overlap = (matched / total_cols * 100) if total_cols > 0 else 0.0

        # Best merge target (workbook with the most matching columns)
        partners = partner_counts.get(wb_name, {})
        merge_target = max(partners, key=partners.get) if partners else None

        # Determine status
        overlap_ratio = overlap / 100.0
        if overlap_ratio >= _RETIRE_THRESHOLD:
            status = ReportStatus.RETIRE
            rationale = (
                f"{overlap:.0f}% of columns exist in other reports. "
                "This report is substantially redundant and can be retired."
            )
        elif overlap_ratio >= _MERGE_THRESHOLD:
            status = ReportStatus.MERGE
            rationale = (
                f"{overlap:.0f}% of columns overlap with other reports. "
                f"Recommended merge target: {merge_target or 'N/A'}."
            )
        elif overlap_ratio >= _REVIEW_THRESHOLD:
            status = ReportStatus.REVIEW
            rationale = (
                f"{overlap:.0f}% of columns overlap. "
                "Review whether partial consolidation is appropriate."
            )
        else:
            status = ReportStatus.KEEP
            rationale = (
                f"Low overlap ({overlap:.0f}%). "
                "This report appears to contain unique data — keep as-is."
            )

        duplicate_columns = sorted(matched_cols_by_wb.get(wb_name, set()))

        recommendations.append(ReportRecommendation(
            report_name=wb_name,
            sheet_count=sheet_count,
            total_columns=total_cols,
            matched_columns=matched,
            overlap_pct=round(overlap, 2),
            status=status,
            rationale=rationale,
            merge_target=merge_target,
            duplicate_columns=duplicate_columns,
        ))
        logger.debug(
            "Rationalization '%s': overlap=%.0f%%, status=%s",
            wb_name, overlap, status.value,
        )

    recommendations.sort(key=lambda r: r.overlap_pct, reverse=True)
    return recommendations


def build_inventory(recommendations: list[ReportRecommendation]) -> pd.DataFrame:
    """Convert recommendations to a tidy inventory DataFrame."""
    rows = []
    for r in recommendations:
        rows.append({
            "report_name": r.report_name,
            "sheet_count": r.sheet_count,
            "total_columns": r.total_columns,
            "matched_columns": r.matched_columns,
            "overlap_pct": r.overlap_pct,
            "recommendation": r.status.value.upper(),
            "merge_target": r.merge_target or "",
            "rationale": r.rationale,
            "duplicate_columns": ", ".join(r.duplicate_columns),
            "action_owner": "",
            "target_date": "",
            "notes": "",
        })
    return pd.DataFrame(rows)
