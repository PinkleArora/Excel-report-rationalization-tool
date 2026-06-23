"""Tests for Consolidation Intelligence components."""

from __future__ import annotations

import pandas as pd
import pytest

from src.agentic_rationalization.grain_detector import detect_grain
from src.agentic_rationalization.compatibility_scorer import score_compatibility


# ---------------------------------------------------------------------------
# Test 1: Policy-level grain detection
# ---------------------------------------------------------------------------

def test_grain_detection_policy_level():
    df = pd.DataFrame({
        "policy_number": [f"POL{i:04d}" for i in range(100)],
        "insured_name": ["Alice"] * 50 + ["Bob"] * 50,
        "premium": [1000.0] * 100,
        "effective_date": ["2024-01-01"] * 100,
    })
    result = detect_grain(df)
    assert result.grain_category == "policy"
    assert result.grain_label == "Policy Level"
    assert result.confidence >= 0.85
    assert "policy_number" in result.primary_key_candidates


# ---------------------------------------------------------------------------
# Test 2: Plan-code-level grain detection
# ---------------------------------------------------------------------------

def test_grain_detection_plan_code_level():
    df = pd.DataFrame({
        "plan_code": [f"PLAN{i:02d}" for i in range(20)],
        "plan_name": [f"Plan {i}" for i in range(20)],
        "monthly_rate": [float(i * 100) for i in range(20)],
    })
    result = detect_grain(df)
    assert result.grain_category == "plan_code"
    assert result.grain_label == "Plan Code Level"
    assert result.confidence >= 0.85


# ---------------------------------------------------------------------------
# Test 3: Pairwise compatibility — same grain (should be compatible)
# ---------------------------------------------------------------------------

def test_pairwise_compatibility_same_grain():
    df_a = pd.DataFrame({
        "policy_number": [f"POL{i:04d}" for i in range(100)],
        "premium": [1000.0] * 100,
        "coverage_amount": [50000.0] * 100,
    })
    df_b = pd.DataFrame({
        "policy_number": [f"POL{i:04d}" for i in range(100)],
        "premium": [1000.0] * 100,
        "benefit_amount": [25000.0] * 100,
    })
    grain_a = detect_grain(df_a)
    grain_b = detect_grain(df_b)
    score = score_compatibility(
        "file_a.xlsx", list(df_a.columns), grain_a,
        "file_b.xlsx", list(df_b.columns), grain_b,
    )
    assert score.grain_compatible is True
    assert score.grain_a == "policy"
    assert score.grain_b == "policy"
    assert score.overall_score > 0.20  # Some overlap from policy_number + premium


# ---------------------------------------------------------------------------
# Test 4: Pairwise compatibility — incompatible grains (policy vs claim)
# ---------------------------------------------------------------------------

def test_pairwise_compatibility_incompatible_grains():
    df_a = pd.DataFrame({
        "policy_number": [f"POL{i:04d}" for i in range(100)],
        "premium": [1000.0] * 100,
    })
    df_b = pd.DataFrame({
        "claim_number": [f"CLM{i:04d}" for i in range(200)],
        "claim_amount": [500.0] * 200,
        "claim_date": ["2024-01-01"] * 200,
    })
    grain_a = detect_grain(df_a)
    grain_b = detect_grain(df_b)
    score = score_compatibility(
        "policy.xlsx", list(df_a.columns), grain_a,
        "claims.xlsx", list(df_b.columns), grain_b,
    )
    assert score.grain_compatible is False
    assert score.grain_a == "policy"
    assert score.grain_b == "claim"
    assert score.recommendation == "Do Not Consolidate"


# ---------------------------------------------------------------------------
# Test 5: Grouping — 3 files where 2 compatible, 1 standalone
# ---------------------------------------------------------------------------

def test_grouping_two_compatible_one_standalone():
    from src.agentic_rationalization.agents.consolidation_intelligence_agent import (
        _build_groups, _build_union_find
    )
    from src.agentic_rationalization.grain_detector import GrainDetectionResult
    from src.agentic_rationalization.compatibility_scorer import CompatibilityScore

    def make_grain(category):
        return GrainDetectionResult(
            grain_label=f"{category} Level",
            grain_category=category,
            confidence=0.85,
            reasoning="test",
            primary_key_candidates=[],
            row_count=100,
            column_count=5,
            distinct_value_counts={},
        )

    files = ["policy_a.xlsx", "policy_b.xlsx", "claim_c.xlsx"]
    grain_map = {
        "policy_a.xlsx": make_grain("policy"),
        "policy_b.xlsx": make_grain("policy"),
        "claim_c.xlsx": make_grain("claim"),
    }
    col_map = {
        "policy_a.xlsx": ["policy_number", "premium", "coverage"],
        "policy_b.xlsx": ["policy_number", "premium", "benefit"],
        "claim_c.xlsx": ["claim_number", "claim_amount"],
    }

    def make_score(fa, fb, grain_compat, overall):
        rec = "Consolidate" if overall >= 0.75 else ("Conditionally Consolidate" if overall >= 0.50 else "Do Not Consolidate")
        return CompatibilityScore(
            file_a=fa, file_b=fb,
            schema_overlap_pct=0.5, kpi_overlap_pct=0.0,
            grain_compatible=grain_compat,
            grain_a=grain_map[fa].grain_category,
            grain_b=grain_map[fb].grain_category,
            key_compatible=grain_compat,
            time_dimension_compatible=True,
            overall_score=overall,
            recommendation=rec,
            recommendation_confidence=0.80,
            reasoning="test",
        )

    pairwise = [
        make_score("policy_a.xlsx", "policy_b.xlsx", True, 0.80),
        make_score("policy_a.xlsx", "claim_c.xlsx", False, 0.10),
        make_score("policy_b.xlsx", "claim_c.xlsx", False, 0.10),
    ]

    groups = _build_groups(files, pairwise, grain_map, col_map, None)

    # Should be 2 groups: [policy_a, policy_b] and [claim_c]
    assert len(groups) == 2
    group_sizes = sorted([len(g.file_names) for g in groups], reverse=True)
    assert group_sizes == [2, 1]

    policy_group = next(g for g in groups if len(g.file_names) == 2)
    assert policy_group.recommendation == "Consolidate"
    assert policy_group.grain_category == "policy"

    standalone = next(g for g in groups if len(g.file_names) == 1)
    assert standalone.recommendation == "Keep Separate"
