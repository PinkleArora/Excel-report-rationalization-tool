"""Detect semantically equivalent columns across workbooks using fuzzy matching."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field

import pandas as pd
from rapidfuzz import fuzz, process

from src.schema_matching.normalizer import normalize_column_name

logger = logging.getLogger(__name__)


@dataclass
class ColumnMatch:
    """A scored mapping between one source column and one target column."""

    source_workbook: str
    source_sheet: str
    source_col: str
    target_workbook: str
    target_sheet: str
    target_col: str
    score: float          # 0–100 (RapidFuzz WRatio or 100 for exact)
    match_type: str       # "exact" | "fuzzy"


def match_columns(
    source: dict[str, pd.DataFrame],
    target: dict[str, pd.DataFrame],
    threshold: float = 80.0,
    source_workbook: str = "",
    target_workbook: str = "",
) -> list[ColumnMatch]:
    """Find column-level matches between *source* and *target* sheet dicts.

    For each column in every *source* sheet, the best-scoring column from any
    *target* sheet that meets *threshold* is returned.  Exact normalised-name
    matches always score 100 regardless of the RapidFuzz result.

    Args:
        source: Sheet-name → DataFrame mapping for the source workbook.
        target: Sheet-name → DataFrame mapping for the target workbook.
        threshold: Minimum score (0–100) to accept a match.
        source_workbook: Optional workbook label for the source side.
        target_workbook: Optional workbook label for the target side.

    Returns:
        Sorted list of :class:`ColumnMatch` objects, highest score first.
        Each source (sheet, col) pair appears at most once.
    """
    # Collect all target (sheet, col, normalised) tuples
    target_entries: list[tuple[str, str, str]] = []
    for sheet_name, df in target.items():
        for col in df.columns:
            target_entries.append((sheet_name, str(col), normalize_column_name(str(col))))

    if not target_entries:
        return []

    target_norms = [e[2] for e in target_entries]

    matches: list[ColumnMatch] = []
    seen_source: set[tuple[str, str]] = set()

    for src_sheet, src_df in source.items():
        for src_col in src_df.columns:
            src_col_str = str(src_col)
            src_norm = normalize_column_name(src_col_str)
            src_key = (src_sheet, src_col_str)

            if src_key in seen_source:
                continue

            # --- exact normalised match (score 100) ---
            exact_hits = [
                (tgt_sheet, tgt_col)
                for tgt_sheet, tgt_col, tgt_norm in target_entries
                if tgt_norm == src_norm
            ]
            if exact_hits:
                tgt_sheet, tgt_col = exact_hits[0]
                matches.append(ColumnMatch(
                    source_workbook=source_workbook,
                    source_sheet=src_sheet,
                    source_col=src_col_str,
                    target_workbook=target_workbook,
                    target_sheet=tgt_sheet,
                    target_col=tgt_col,
                    score=100.0,
                    match_type="exact",
                ))
                seen_source.add(src_key)
                continue

            # --- fuzzy match ---
            results = process.extractOne(
                src_norm,
                target_norms,
                scorer=fuzz.WRatio,
                score_cutoff=threshold,
            )
            if results is not None:
                best_norm, best_score, best_idx = results
                tgt_sheet, tgt_col, _ = target_entries[best_idx]
                matches.append(ColumnMatch(
                    source_workbook=source_workbook,
                    source_sheet=src_sheet,
                    source_col=src_col_str,
                    target_workbook=target_workbook,
                    target_sheet=tgt_sheet,
                    target_col=tgt_col,
                    score=float(best_score),
                    match_type="fuzzy",
                ))
                seen_source.add(src_key)

    matches.sort(key=lambda m: m.score, reverse=True)
    logger.debug(
        "match_columns: %d source cols → %d matches (threshold=%.0f)",
        sum(len(df.columns) for df in source.values()),
        len(matches),
        threshold,
    )
    return matches


def match_workbooks(
    source_bundle,
    target_bundle,
    threshold: float = 80.0,
) -> list[ColumnMatch]:
    """Cross-match all sheets of two :class:`~src.ingestion.loader.WorkbookBundle` objects.

    Args:
        source_bundle: Source :class:`WorkbookBundle`.
        target_bundle: Target :class:`WorkbookBundle`.
        threshold: Minimum RapidFuzz score.

    Returns:
        List of :class:`ColumnMatch` with workbook names populated.
    """
    return match_columns(
        source=source_bundle.sheets,
        target=target_bundle.sheets,
        threshold=threshold,
        source_workbook=source_bundle.file_name,
        target_workbook=target_bundle.file_name,
    )


def match_all_bundles(
    bundles: list,
    threshold: float = 80.0,
) -> list[ColumnMatch]:
    """Run pairwise matching across every combination of bundles.

    Each ordered pair (i, j) where i < j is matched once, with bundle i as
    source and bundle j as target.

    Args:
        bundles: List of :class:`~src.ingestion.loader.WorkbookBundle` objects.
        threshold: Minimum RapidFuzz score.

    Returns:
        Flat, deduplicated list of all :class:`ColumnMatch` objects.
    """
    all_matches: list[ColumnMatch] = []
    for i in range(len(bundles)):
        for j in range(i + 1, len(bundles)):
            pair = match_workbooks(bundles[i], bundles[j], threshold=threshold)
            all_matches.extend(pair)
            logger.info(
                "Matched '%s' ↔ '%s': %d column match(es)",
                bundles[i].file_name,
                bundles[j].file_name,
                len(pair),
            )
    return all_matches


def build_match_matrix(
    bundles: list,
    threshold: float = 80.0,
) -> pd.DataFrame:
    """Build a pairwise match-count matrix for all provided bundles.

    Returns:
        Square DataFrame indexed and columned by workbook names;
        cell value = number of matched columns between that pair.
    """
    names = [b.file_name for b in bundles]
    import numpy as np
    matrix = pd.DataFrame(np.zeros((len(names), len(names)), dtype=int),
                          index=names, columns=names)

    for i in range(len(bundles)):
        for j in range(i + 1, len(bundles)):
            count = len(match_workbooks(bundles[i], bundles[j], threshold=threshold))
            matrix.loc[names[i], names[j]] = count
            matrix.loc[names[j], names[i]] = count

    return matrix
