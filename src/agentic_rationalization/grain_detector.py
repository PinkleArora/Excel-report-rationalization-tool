"""Detect the data grain (level of detail) of a source DataFrame."""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from src.schema_matching.normalizer import normalize_column_name


@dataclass
class GrainDetectionResult:
    grain_label: str
    grain_category: str  # "policy"|"certificate"|"coverage"|"rider"|"claim"|"transaction"|"plan_code"|"product"|"aggregated"|"unknown"
    confidence: float
    reasoning: str
    primary_key_candidates: list[str]
    row_count: int
    column_count: int
    distinct_value_counts: dict[str, int]


# Grain indicator keywords mapped to category. More specific categories first.
_GRAIN_INDICATORS: dict[str, list[str]] = {
    "claim":       ["claim_number", "claim_id", "claim_no", "clm_num"],
    "transaction": ["transaction_id", "trans_id", "txn_id", "transaction_date"],
    "coverage":    ["coverage_code", "coverage_id", "cov_code", "coverage_type"],
    "rider":       ["rider_code", "rider_id", "rider_type"],
    "certificate": ["cert_number", "cert_no", "certificate_id", "certificate_number"],
    "policy":      ["policy_number", "policy_no", "pol_num", "policy_id", "policy_nbr"],
    "plan_code":   ["plan_code", "plan_id", "plan_type", "product_code"],
    "product":     ["product_id", "product_name", "product_type"],
    "aggregated":  ["total", "subtotal", "grand_total", "summary"],
}

_GRAIN_LABELS: dict[str, str] = {
    "claim":       "Claim Level",
    "transaction": "Transaction Level",
    "coverage":    "Coverage Level",
    "rider":       "Rider Level",
    "certificate": "Certificate Level",
    "policy":      "Policy Level",
    "plan_code":   "Plan Code Level",
    "product":     "Product Level",
    "aggregated":  "Aggregated / Summary Level",
    "unknown":     "Unknown Level",
}

# Specificity order (most specific first)
_SPECIFICITY: list[str] = [
    "claim", "transaction", "coverage", "rider", "certificate",
    "policy", "plan_code", "product", "aggregated",
]


def detect_grain(df: pd.DataFrame, kpi_result=None) -> GrainDetectionResult:
    """Detect the data grain of *df*.

    Args:
        df: Source DataFrame to analyse.
        kpi_result: Optional KpiAnalysisResult; if provided and all KPIs are
                    SUMIFS/SUMPRODUCT aggregations with no unique-key column
                    found, infer "aggregated".

    Returns:
        GrainDetectionResult with grain label, category, confidence, etc.
    """
    row_count = len(df)
    column_count = len(df.columns)

    # Compute distinct value counts per column
    distinct_value_counts: dict[str, int] = {}
    for col in df.columns:
        try:
            distinct_value_counts[col] = int(df[col].nunique())
        except Exception:
            distinct_value_counts[col] = 0

    # Primary key candidates: uniqueness ratio > 0.95
    primary_key_candidates: list[str] = []
    if row_count > 0:
        for col, dcount in distinct_value_counts.items():
            if dcount / row_count > 0.95:
                primary_key_candidates.append(col)

    # Normalize all column names for matching
    col_normalized: dict[str, str] = {col: normalize_column_name(col) for col in df.columns}

    # Find matching grain indicators
    matched_categories: list[str] = []
    matched_columns: dict[str, list[str]] = {}  # category -> list of matched column names

    for category, indicators in _GRAIN_INDICATORS.items():
        for col, normed in col_normalized.items():
            if normed in indicators:
                if category not in matched_columns:
                    matched_columns[category] = []
                matched_columns[category].append(col)
                if category not in matched_categories:
                    matched_categories.append(category)

    # Pick most specific category
    grain_category = "unknown"
    grain_label = _GRAIN_LABELS["unknown"]
    confidence = 0.40
    reasoning_parts: list[str] = []

    for specific_cat in _SPECIFICITY:
        if specific_cat in matched_categories:
            grain_category = specific_cat
            grain_label = _GRAIN_LABELS[specific_cat]
            cols_matched = matched_columns.get(specific_cat, [])
            confidence = 0.85
            if primary_key_candidates:
                confidence = min(confidence + 0.10, 0.97)
            reasoning_parts.append(
                f"Column(s) {cols_matched} match '{specific_cat}' grain indicators."
            )
            break

    if grain_category == "unknown" and primary_key_candidates:
        # Heuristic: has high-cardinality columns but no indicator words
        reasoning_parts.append(
            f"No grain indicator columns found but high-cardinality column(s) "
            f"{primary_key_candidates[:3]} suggest a row-level (not aggregated) grain."
        )
        confidence = 0.30

    # KPI-based aggregation inference
    if kpi_result is not None and grain_category == "unknown":
        _AGG_FUNCS = {"SUMIFS", "SUMPRODUCT", "SUMIF", "AVERAGEIF", "AVERAGEIFS"}
        all_agg = all(
            getattr(dep, "aggregate_function", "").upper() in _AGG_FUNCS
            for dep in getattr(kpi_result, "dependencies", [])
        )
        if all_agg and not primary_key_candidates:
            grain_category = "aggregated"
            grain_label = _GRAIN_LABELS["aggregated"]
            confidence = 0.65
            reasoning_parts.append(
                "All KPI formulas use aggregation functions (SUMIFS/SUMPRODUCT) "
                "and no unique-key column found — inferred aggregated/summary grain."
            )

    if not reasoning_parts:
        reasoning_parts.append(
            f"No grain indicator columns detected in {column_count} columns. "
            "Grain cannot be reliably determined."
        )

    if primary_key_candidates:
        reasoning_parts.append(
            f"Primary key candidates (uniqueness > 95%): {primary_key_candidates[:5]}."
        )

    return GrainDetectionResult(
        grain_label=grain_label,
        grain_category=grain_category,
        confidence=confidence,
        reasoning=" ".join(reasoning_parts),
        primary_key_candidates=primary_key_candidates,
        row_count=row_count,
        column_count=column_count,
        distinct_value_counts=distinct_value_counts,
    )
