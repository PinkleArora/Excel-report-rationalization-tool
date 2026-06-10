"""Column-name normalisation utilities shared across matching and consolidation."""

from __future__ import annotations

import re


def normalize_column_name(name: str) -> str:
    """Return a stable, lowercase, alphanumeric-underscore identifier.

    Examples::

        normalize_column_name("Revenue (£)")  -> "revenue"
        normalize_column_name("TOTAL  COST")  -> "total_cost"
        normalize_column_name("  Q1 Sales ")  -> "q1_sales"
    """
    s = str(name).strip().lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = s.strip("_")
    return s or "unnamed"


def canonical_for_group(names: list[str]) -> str:
    """Pick the canonical representative from a set of equivalent column names.

    Selection rule: shortest normalised name; ties broken alphabetically.
    """
    normed = sorted({normalize_column_name(n) for n in names})
    return min(normed, key=lambda s: (len(s), s))
