"""Detect semantically equivalent columns across workbooks.

Matching is split from merging.  A match documents a *relationship*
between two columns; whether that relationship justifies automatic column
unification is controlled by the match's ``confidence`` level:

``EXACT``   — normalised names are identical (score = 100).
              Always safe to merge automatically.
``HIGH``    — fuzzy score ≥ 95 and dtypes are compatible.
              Merged only if the caller passes ``merge_high_confidence=True``.
``REVIEW``  — fuzzy score 80–94.
              Never merged automatically; surfaced in the Source Mapping sheet
              as a suggestion requiring human confirmation.

Additionally, ``detect_many_to_one_mappings()`` finds cases where N ≥ 2
distinct source columns from *different* columns all match the same target
canonical name.  Those matches are downgraded to REVIEW regardless of score,
because automatic merging would silently destroy distinct business metrics.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from enum import Enum

import pandas as pd
from rapidfuzz import fuzz, process

from src.schema_matching.normalizer import normalize_column_name

logger = logging.getLogger(__name__)


class MatchConfidence(str, Enum):
    EXACT  = "exact"   # normalized names identical — always safe to merge
    HIGH   = "high"    # score ≥ 95 — merge only when explicitly opted-in
    REVIEW = "review"  # score 80–94 — document only, never auto-merge


# Score threshold above which a fuzzy match is classified HIGH instead of REVIEW
_HIGH_CONFIDENCE_SCORE = 95.0


@dataclass
class ColumnMatch:
    """A scored mapping between one source column and one target column."""

    source_workbook: str
    source_sheet: str
    source_col: str
    target_workbook: str
    target_sheet: str
    target_col: str
    score: float           # 0–100 (RapidFuzz WRatio or 100 for exact)
    match_type: str        # "exact" | "fuzzy"
    confidence: MatchConfidence = MatchConfidence.REVIEW
    safe_to_merge: bool = False  # set by classify_match(); drives consolidation

    def __post_init__(self) -> None:
        # Derive confidence and safe_to_merge from match_type / score if not set
        if self.match_type == "exact":
            object.__setattr__(self, "confidence", MatchConfidence.EXACT)
            object.__setattr__(self, "safe_to_merge", True)
        elif self.score >= _HIGH_CONFIDENCE_SCORE:
            object.__setattr__(self, "confidence", MatchConfidence.HIGH)
            # safe_to_merge stays False until caller opts in


def match_columns(
    source: dict[str, pd.DataFrame],
    target: dict[str, pd.DataFrame],
    threshold: float = 80.0,
    source_workbook: str = "",
    target_workbook: str = "",
    merge_high_confidence: bool = False,
) -> list[ColumnMatch]:
    """Find column-level matches between *source* and *target* sheet dicts.

    **Exact** matches (identical normalised names) are always returned and
    marked ``safe_to_merge=True``.

    **Fuzzy** matches (RapidFuzz WRatio ≥ *threshold*) are returned for
    documentation purposes.  They are marked ``safe_to_merge=True`` only
    when *merge_high_confidence=True* **and** the score is ≥ 95.

    Args:
        source: Sheet-name → DataFrame mapping for the source workbook.
        target: Sheet-name → DataFrame mapping for the target workbook.
        threshold: Minimum score (0–100) to include a fuzzy match in output.
            Does not affect ``safe_to_merge`` classification.
        source_workbook: Label for the source side.
        target_workbook: Label for the target side.
        merge_high_confidence: If True, matches scoring ≥ 95 are also
            marked ``safe_to_merge=True``.

    Returns:
        Sorted list of :class:`ColumnMatch` objects, highest score first.
        Each source (sheet, col) pair appears at most once.
    """
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

            # --- exact normalised match ---
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
                    confidence=MatchConfidence.EXACT,
                    safe_to_merge=True,
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
                score_f = float(best_score)
                is_high = score_f >= _HIGH_CONFIDENCE_SCORE
                confidence = MatchConfidence.HIGH if is_high else MatchConfidence.REVIEW
                safe = merge_high_confidence and is_high
                matches.append(ColumnMatch(
                    source_workbook=source_workbook,
                    source_sheet=src_sheet,
                    source_col=src_col_str,
                    target_workbook=target_workbook,
                    target_sheet=tgt_sheet,
                    target_col=tgt_col,
                    score=score_f,
                    match_type="fuzzy",
                    confidence=confidence,
                    safe_to_merge=safe,
                ))
                seen_source.add(src_key)

    matches.sort(key=lambda m: m.score, reverse=True)
    logger.debug(
        "match_columns '%s'↔'%s': %d source cols → %d matches "
        "(%d exact, %d high, %d review)",
        source_workbook, target_workbook,
        sum(len(df.columns) for df in source.values()),
        len(matches),
        sum(1 for m in matches if m.confidence == MatchConfidence.EXACT),
        sum(1 for m in matches if m.confidence == MatchConfidence.HIGH),
        sum(1 for m in matches if m.confidence == MatchConfidence.REVIEW),
    )
    return matches


def match_workbooks(
    source_bundle,
    target_bundle,
    threshold: float = 80.0,
    merge_high_confidence: bool = False,
) -> list[ColumnMatch]:
    """Cross-match all sheets of two :class:`~src.ingestion.loader.WorkbookBundle` objects."""
    return match_columns(
        source=source_bundle.sheets,
        target=target_bundle.sheets,
        threshold=threshold,
        source_workbook=source_bundle.file_name,
        target_workbook=target_bundle.file_name,
        merge_high_confidence=merge_high_confidence,
    )


def match_all_bundles(
    bundles: list,
    threshold: float = 80.0,
    merge_high_confidence: bool = False,
) -> list[ColumnMatch]:
    """Run pairwise matching across every combination of bundles.

    Returns all matches — exact, high-confidence, and review-only — so the
    full relationship map can be documented.  The ``safe_to_merge`` flag on
    each match controls which relationships drive column unification in
    :func:`~src.consolidation.consolidator.consolidate`.
    """
    all_matches: list[ColumnMatch] = []
    for i in range(len(bundles)):
        for j in range(i + 1, len(bundles)):
            pair = match_workbooks(
                bundles[i], bundles[j],
                threshold=threshold,
                merge_high_confidence=merge_high_confidence,
            )
            all_matches.extend(pair)
            exact_n  = sum(1 for m in pair if m.confidence == MatchConfidence.EXACT)
            high_n   = sum(1 for m in pair if m.confidence == MatchConfidence.HIGH)
            review_n = sum(1 for m in pair if m.confidence == MatchConfidence.REVIEW)
            logger.info(
                "Matched '%s' ↔ '%s': %d match(es) [exact=%d high=%d review=%d]",
                bundles[i].file_name, bundles[j].file_name,
                len(pair), exact_n, high_n, review_n,
            )
    return all_matches


def detect_many_to_one_mappings(
    matches: list[ColumnMatch],
    bundles: list | None = None,
) -> list[dict]:
    """Detect cases where ≥ 2 distinct source columns map to the same target canonical.

    A many-to-one mapping is dangerous: it means N different business metrics
    from different workbooks would all be renamed to the same column name,
    silently destroying distinct information.

    Any match involved in a many-to-one situation is downgraded to
    ``safe_to_merge=False`` and ``confidence=REVIEW``, in-place.

    Args:
        matches: All matches from :func:`match_all_bundles`.
        bundles: Optional list of bundles (unused, reserved for future use).

    Returns:
        List of dicts describing each detected many-to-one situation, suitable
        for inclusion in the Issues Log.
    """
    from collections import defaultdict

    # Group by target canonical name
    target_canonical_to_sources: dict[str, list[ColumnMatch]] = defaultdict(list)
    for m in matches:
        if not m.safe_to_merge:
            continue
        target_norm = normalize_column_name(m.target_col)
        target_canonical_to_sources[target_norm].append(m)

    issues: list[dict] = []
    for canonical, group in target_canonical_to_sources.items():
        # Check if there are genuinely different source columns (not just same col different workbook)
        distinct_source_norms = {normalize_column_name(m.source_col) for m in group}
        if len(distinct_source_norms) <= 1:
            continue  # same column name from multiple workbooks — that's fine

        # Multiple distinct source column names all map to same canonical → many-to-one
        for m in group:
            # Downgrade in-place
            object.__setattr__(m, "safe_to_merge", False)
            object.__setattr__(m, "confidence", MatchConfidence.REVIEW)

        source_summaries = [
            f"{m.source_workbook}/{m.source_sheet}/{m.source_col} (score={m.score:.0f})"
            for m in group
        ]
        issues.append({
            "issue_type": "many_to_one_mapping",
            "target_canonical": canonical,
            "source_columns": source_summaries,
            "match_count": len(group),
            "description": (
                f"Many-to-one mapping detected: {len(group)} distinct source columns "
                f"all match canonical '{canonical}'. "
                f"These columns may represent different business metrics. "
                f"Automatic merging has been suppressed. Review the Source Mapping "
                f"sheet and confirm which columns are truly equivalent."
            ),
        })
        logger.warning(
            "Many-to-one mapping suppressed for canonical '%s': %s",
            canonical, source_summaries,
        )

    return issues


def build_match_matrix(
    bundles: list,
    threshold: float = 80.0,
) -> pd.DataFrame:
    """Build a pairwise exact-match count matrix for all provided bundles."""
    names = [b.file_name for b in bundles]
    import numpy as np
    matrix = pd.DataFrame(
        np.zeros((len(names), len(names)), dtype=int),
        index=names, columns=names,
    )
    for i in range(len(bundles)):
        for j in range(i + 1, len(bundles)):
            count = sum(
                1 for m in match_workbooks(bundles[i], bundles[j], threshold=threshold)
                if m.confidence == MatchConfidence.EXACT
            )
            matrix.loc[names[i], names[j]] = count
            matrix.loc[names[j], names[i]] = count
    return matrix
