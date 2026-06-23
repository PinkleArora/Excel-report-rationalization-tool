"""Score pairwise compatibility between two WorkbookBundles."""

from __future__ import annotations

from dataclasses import dataclass

from src.schema_matching.normalizer import normalize_column_name
from src.agentic_rationalization.grain_detector import GrainDetectionResult

_TIME_KEYWORDS = {
    "date", "period", "month", "year", "quarter",
    "effective_date", "as_of_date", "asofdate", "effectivedate",
}


@dataclass
class CompatibilityScore:
    file_a: str
    file_b: str
    schema_overlap_pct: float
    kpi_overlap_pct: float
    grain_compatible: bool
    grain_a: str
    grain_b: str
    key_compatible: bool
    time_dimension_compatible: bool
    overall_score: float
    recommendation: str   # "Consolidate" | "Conditionally Consolidate" | "Do Not Consolidate"
    recommendation_confidence: float
    reasoning: str


def _has_time_dimension(columns: list[str]) -> bool:
    for col in columns:
        normed = normalize_column_name(col)
        if any(kw in normed for kw in _TIME_KEYWORDS):
            return True
    return False


def _canonical_columns(columns: list[str]) -> set[str]:
    return {normalize_column_name(c) for c in columns}


def score_compatibility(
    file_a: str,
    cols_a: list[str],
    grain_a: GrainDetectionResult,
    file_b: str,
    cols_b: list[str],
    grain_b: GrainDetectionResult,
    kpi_result=None,
) -> CompatibilityScore:
    """Compute a pairwise compatibility score between two files."""
    can_a = _canonical_columns(cols_a)
    can_b = _canonical_columns(cols_b)

    # Schema overlap
    union = can_a | can_b
    intersection = can_a & can_b
    schema_overlap_pct = len(intersection) / len(union) if union else 0.0

    # KPI overlap
    kpi_overlap_pct = 0.0
    if kpi_result is not None:
        kpi_cols: set[str] = getattr(kpi_result, "referenced_canonicals", set())
        kpi_in_a = kpi_cols & can_a
        kpi_in_b = kpi_cols & can_b
        kpi_union = kpi_in_a | kpi_in_b
        kpi_intersection = kpi_in_a & kpi_in_b
        kpi_overlap_pct = len(kpi_intersection) / len(kpi_union) if kpi_union else 0.0

    # Grain compatibility
    grain_compatible = (
        grain_a.grain_category == grain_b.grain_category
        and grain_a.grain_category != "unknown"
    )

    # Key compatibility: both have primary key candidates of similar type
    pk_a_normed = {normalize_column_name(c) for c in grain_a.primary_key_candidates}
    pk_b_normed = {normalize_column_name(c) for c in grain_b.primary_key_candidates}
    key_compatible = bool(pk_a_normed & pk_b_normed) or (
        bool(pk_a_normed) and bool(pk_b_normed)
        and grain_compatible
    )

    # Time dimension compatibility
    time_a = _has_time_dimension(cols_a)
    time_b = _has_time_dimension(cols_b)
    time_dimension_compatible = time_a == time_b

    # Weighted overall score
    grain_score = 1.0 if grain_compatible else 0.0
    key_score = 1.0 if key_compatible else 0.0
    time_score = 1.0 if time_dimension_compatible else 0.0

    overall_score = (
        schema_overlap_pct * 0.30
        + kpi_overlap_pct * 0.35
        + grain_score * 0.20
        + key_score * 0.10
        + time_score * 0.05
    )

    # Recommendation
    if overall_score >= 0.75:
        recommendation = "Consolidate"
        rec_confidence = min(0.90, 0.70 + overall_score * 0.20)
    elif overall_score >= 0.50:
        recommendation = "Conditionally Consolidate"
        rec_confidence = 0.65
    else:
        recommendation = "Do Not Consolidate"
        rec_confidence = min(0.90, 0.50 + (1.0 - overall_score) * 0.40)

    # Reasoning
    reasoning_parts = [
        f"Schema overlap: {schema_overlap_pct:.1%}.",
        f"KPI column overlap: {kpi_overlap_pct:.1%}.",
        f"Grain: {grain_a.grain_category} vs {grain_b.grain_category} "
        f"({'compatible' if grain_compatible else 'incompatible'}).",
    ]
    if not grain_compatible:
        reasoning_parts.append(
            f"Grain mismatch is the primary reason for low score — "
            f"{file_a} is {grain_a.grain_label} while {file_b} is {grain_b.grain_label}."
        )
    if not time_dimension_compatible:
        reasoning_parts.append(
            "One file has a time/period dimension and the other does not."
        )

    return CompatibilityScore(
        file_a=file_a,
        file_b=file_b,
        schema_overlap_pct=schema_overlap_pct,
        kpi_overlap_pct=kpi_overlap_pct,
        grain_compatible=grain_compatible,
        grain_a=grain_a.grain_category,
        grain_b=grain_b.grain_category,
        key_compatible=key_compatible,
        time_dimension_compatible=time_dimension_compatible,
        overall_score=overall_score,
        recommendation=recommendation,
        recommendation_confidence=rec_confidence,
        reasoning=" ".join(reasoning_parts),
    )
