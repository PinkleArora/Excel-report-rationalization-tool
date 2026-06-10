"""Merge aligned reports into a single master dataset."""

from __future__ import annotations

import pandas as pd

from src.schema_matching.matcher import ColumnMatch


def consolidate(
    bundles: list[dict[str, pd.DataFrame]],
    matches: list[ColumnMatch],
    strategy: str = "union",
) -> pd.DataFrame:
    """Merge workbook data guided by *matches*.

    Args:
        bundles: List of sheet-name → DataFrame mappings (one per workbook).
        matches: Column match list from :mod:`src.schema_matching.matcher`.
        strategy: ``"union"`` — keep all columns; ``"intersection"`` — keep only
            matched columns.

    Returns:
        A single consolidated :class:`pandas.DataFrame`.
    """
    raise NotImplementedError


def resolve_conflicts(master: pd.DataFrame, strategy: str = "last_wins") -> pd.DataFrame:
    """Resolve duplicate rows or conflicting values in *master*.

    Args:
        master: The raw consolidated DataFrame.
        strategy: Conflict resolution strategy (``"last_wins"`` | ``"first_wins"`` |
            ``"raise"``).

    Returns:
        Cleaned DataFrame with conflicts resolved.
    """
    raise NotImplementedError
