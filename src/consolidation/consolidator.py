"""Merge aligned reports into a single master dataset with full source lineage."""

from __future__ import annotations

import logging

import pandas as pd

from src.schema_matching.matcher import ColumnMatch
from src.schema_matching.normalizer import normalize_column_name

logger = logging.getLogger(__name__)

# Reserved lineage column names prepended to every row in the master dataset
LINEAGE_COL_WORKBOOK = "_source_workbook"
LINEAGE_COL_SHEET = "_source_sheet"


def consolidate(
    bundles: list,
    matches: list[ColumnMatch],
    strategy: str = "union",
    collision_log: list | None = None,
) -> pd.DataFrame:
    """Merge workbook data guided by *matches*.

    Only matches with ``safe_to_merge=True`` drive column unification.
    Fuzzy REVIEW-level matches are ignored during consolidation — they are
    preserved in the Source Mapping sheet for human review.

    Every row in the output carries two lineage columns:
    ``_source_workbook`` and ``_source_sheet``.

    Args:
        bundles: List of :class:`~src.ingestion.loader.WorkbookBundle` objects.
        matches: All matches from :func:`~src.schema_matching.matcher.match_all_bundles`.
            Only ``safe_to_merge=True`` matches are applied.
        strategy: ``"union"`` — keep all columns (default);
            ``"intersection"`` — keep only columns confirmed by a safe match.
        collision_log: Optional mutable list; collision/skip entries are appended.

    Returns:
        Consolidated :class:`pandas.DataFrame` with lineage columns first.
    """
    safe_matches = [m for m in matches if m.safe_to_merge]
    canonical_map = _build_canonical_map(bundles, safe_matches, collision_log=collision_log)
    all_frames: list[pd.DataFrame] = []

    for bundle in bundles:
        for sheet_name, df in bundle.sheets.items():
            if df.empty and len(df.columns) == 0:
                continue
            frame = df.copy()
            # Rename every column to its canonical form
            rename = {
                str(col): canonical_map.get((bundle.file_name, sheet_name, str(col)),
                                            normalize_column_name(str(col)))
                for col in frame.columns
            }
            frame.rename(columns=rename, inplace=True)
            frame = _disambiguate_columns(frame)
            # Prepend lineage columns
            frame.insert(0, LINEAGE_COL_SHEET, sheet_name)
            frame.insert(0, LINEAGE_COL_WORKBOOK, bundle.file_name)
            all_frames.append(frame)
            logger.debug("Consolidated sheet '%s' from '%s': %d rows", sheet_name, bundle.file_name, len(frame))

    if not all_frames:
        logger.warning("consolidate: no frames produced — returning empty DataFrame")
        return pd.DataFrame()

    master = pd.concat(all_frames, ignore_index=True, sort=False)

    if strategy == "intersection":
        matched_canonical = {
            canonical_map.get((m.source_workbook, m.source_sheet, m.source_col),
                               normalize_column_name(m.source_col))
            for m in matches
        }
        keep = [LINEAGE_COL_WORKBOOK, LINEAGE_COL_SHEET] + [
            c for c in master.columns
            if c not in (LINEAGE_COL_WORKBOOK, LINEAGE_COL_SHEET) and c in matched_canonical
        ]
        master = master[keep]

    logger.info(
        "Consolidated %d bundle(s) → %d rows × %d cols (strategy=%s)",
        len(bundles), len(master), len(master.columns), strategy,
    )
    return master


def resolve_conflicts(master: pd.DataFrame, strategy: str = "last_wins") -> pd.DataFrame:
    """Remove duplicate rows or resolve conflicting values in *master*.

    Duplicates are detected on all non-lineage columns.  Lineage columns are
    preserved on the surviving row.

    Args:
        master: The consolidated DataFrame (output of :func:`consolidate`).
        strategy: ``"last_wins"`` — keep the last occurrence (default);
            ``"first_wins"`` — keep the first occurrence;
            ``"raise"`` — raise :class:`ValueError` if duplicates exist.

    Returns:
        Cleaned :class:`pandas.DataFrame`.
    """
    data_cols = [c for c in master.columns if c not in (LINEAGE_COL_WORKBOOK, LINEAGE_COL_SHEET)]
    if not data_cols:
        return master

    dupe_mask = master.duplicated(subset=data_cols, keep=False)
    n_dupes = int(dupe_mask.sum())

    if n_dupes == 0:
        return master

    if strategy == "raise":
        raise ValueError(f"resolve_conflicts: {n_dupes} duplicate row(s) found.")

    keep = "last" if strategy == "last_wins" else "first"
    result = master.drop_duplicates(subset=data_cols, keep=keep).reset_index(drop=True)
    logger.info("resolve_conflicts: removed %d duplicate row(s) (strategy=%s)", n_dupes - (n_dupes // 2), strategy)
    return result


def build_source_mapping_df(
    bundles: list,
    matches: list[ColumnMatch],
    collision_log: list | None = None,
) -> pd.DataFrame:
    """Build a tidy Source Mapping table showing all matches.

    ALL match confidence levels are included so the human reviewer can see
    both confirmed merges and suggestions requiring review.

    Columns:
        source_workbook, source_sheet, source_column,
        canonical_column (applied in master data),
        match_type, match_score, match_confidence,
        safe_to_merge (True = applied automatically),
        review_required (True = human confirmation needed),
        matched_to_workbook, matched_to_sheet, matched_to_column,
        is_matched.
    """
    safe_matches = [m for m in matches if m.safe_to_merge]
    canonical_map = _build_canonical_map(bundles, safe_matches, collision_log=collision_log)

    # Index ALL matches (including review-only) for documentation
    match_index: dict[tuple, ColumnMatch] = {}
    for m in matches:
        key = (m.source_workbook, m.source_sheet, m.source_col)
        # Prefer higher-confidence match if multiple exist for same source col
        existing = match_index.get(key)
        if existing is None or m.score > existing.score:
            match_index[key] = m

    rows: list[dict] = []
    for bundle in bundles:
        for sheet_name, df in bundle.sheets.items():
            for col in df.columns:
                col_str = str(col)
                key = (bundle.file_name, sheet_name, col_str)
                canonical = canonical_map.get(key, normalize_column_name(col_str))
                match = match_index.get(key)
                confidence_str = match.confidence.value if match else "none"
                rows.append({
                    "source_workbook": bundle.file_name,
                    "source_sheet": sheet_name,
                    "source_column": col_str,
                    "canonical_column": canonical,
                    "match_type": match.match_type if match else "none",
                    "match_score": round(match.score, 1) if match else None,
                    "match_confidence": confidence_str,
                    "safe_to_merge": match.safe_to_merge if match else False,
                    "review_required": (
                        match is not None and not match.safe_to_merge
                    ),
                    "matched_to_workbook": match.target_workbook if match else "",
                    "matched_to_sheet": match.target_sheet if match else "",
                    "matched_to_column": match.target_col if match else "",
                    "is_matched": match is not None,
                })

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _detect_intra_frame_collisions(
    canon_map: dict[tuple[str, str, str], str],
) -> list[tuple[str, str, list[str]]]:
    """Return list of (workbook, sheet, [colliding_originals]) where ≥2 cols share a canonical."""
    from collections import defaultdict
    by_frame: dict[tuple[str, str], dict[str, list[str]]] = defaultdict(lambda: defaultdict(list))
    for (wb, sh, col), can in canon_map.items():
        by_frame[(wb, sh)][can].append(col)
    collisions = []
    for (wb, sh), can_to_cols in by_frame.items():
        for can, cols in can_to_cols.items():
            if len(cols) > 1:
                collisions.append((wb, sh, cols))
    return collisions


def _disambiguate_columns(frame: pd.DataFrame) -> pd.DataFrame:
    """Rename duplicate column names in-place with _dup2, _dup3 suffixes."""
    seen: dict[str, int] = {}
    new_cols = []
    for col in frame.columns:
        if col not in seen:
            seen[col] = 1
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}_dup{seen[col]}")
    frame.columns = new_cols
    return frame


def _build_canonical_map(
    bundles: list,
    matches: list[ColumnMatch],
    collision_log: list | None = None,
) -> dict[tuple[str, str, str], str]:
    """Return mapping (workbook, sheet, original_col) → canonical_col_name.

    Propagation is skipped for any match that would create intra-frame collisions
    (two distinct columns in the same sheet resolving to the same canonical name).
    Skipped matches are appended to *collision_log* if provided.
    """
    col_to_canonical: dict[tuple[str, str, str], str] = {}

    for bundle in bundles:
        for sheet_name, df in bundle.sheets.items():
            for col in df.columns:
                key = (bundle.file_name, sheet_name, str(col))
                col_to_canonical[key] = normalize_column_name(str(col))

    for m in matches:
        src_key = (m.source_workbook, m.source_sheet, m.source_col)
        tgt_key = (m.target_workbook, m.target_sheet, m.target_col)

        src_current = col_to_canonical.get(src_key, normalize_column_name(m.source_col))
        tgt_current = col_to_canonical.get(tgt_key, normalize_column_name(m.target_col))

        if src_current == tgt_current:
            continue  # already unified

        canonical = min(src_current, tgt_current, key=lambda s: (len(s), s))

        # Simulate the propagation
        candidate = dict(col_to_canonical)
        for key, val in candidate.items():
            if val in (src_current, tgt_current):
                candidate[key] = canonical

        collisions = _detect_intra_frame_collisions(candidate)
        if collisions:
            for wb, sh, cols in collisions:
                logger.warning(
                    "Skipping match %s→%s: would create intra-frame collision in '%s/%s': %s",
                    m.source_col, m.target_col, wb, sh, cols,
                )
                if collision_log is not None:
                    collision_log.append({
                        "source_col": m.source_col,
                        "target_col": m.target_col,
                        "workbook": wb,
                        "sheet": sh,
                        "colliding_columns": cols,
                        "canonical_attempted": canonical,
                    })
        else:
            col_to_canonical = candidate

    return col_to_canonical
