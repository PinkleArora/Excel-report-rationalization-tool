"""Schema matching — detect semantically equivalent columns across workbooks."""

from src.schema_matching.matcher import (
    ColumnMatch,
    build_match_matrix,
    match_all_bundles,
    match_columns,
    match_workbooks,
)
from src.schema_matching.normalizer import canonical_for_group, normalize_column_name

__all__ = [
    "ColumnMatch",
    "match_columns",
    "match_workbooks",
    "match_all_bundles",
    "build_match_matrix",
    "normalize_column_name",
    "canonical_for_group",
]
