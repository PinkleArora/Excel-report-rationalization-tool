"""Dataclasses that represent extracted workbook/sheet/column metadata."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class ColumnMetadata:
    """Column-level profile extracted from a single sheet."""

    workbook_name: str
    sheet_name: str
    column_index: int          # 0-based position in the sheet
    column_name: str
    inferred_dtype: str        # pandas dtype string, e.g. "int64", "object"
    semantic_type: str         # coarser label: "integer", "float", "text", "datetime", "boolean", "empty"
    non_null_count: int
    null_count: int
    null_pct: float            # 0-100
    unique_count: int
    is_empty: bool             # True when null_pct == 100
    sample_values: list[Any] = field(default_factory=list)   # up to 5 non-null values
    min_value: Any = None
    max_value: Any = None
    mean_value: float | None = None


@dataclass
class SheetMetadata:
    """Sheet-level profile."""

    workbook_name: str
    sheet_name: str
    row_count: int             # data rows (excluding header)
    col_count: int
    empty_col_count: int
    formula_cell_count: int    # cells whose raw value starts with "="
    column_names: list[str] = field(default_factory=list)
    columns: list[ColumnMetadata] = field(default_factory=list)


@dataclass
class WorkbookMetadata:
    """Workbook-level profile."""

    workbook_name: str
    source_path: Path
    file_extension: str
    sheet_count: int
    total_rows: int            # sum across all sheets
    total_cols: int            # sum across all sheets
    total_formula_cells: int
    total_empty_cols: int
    sheets: list[SheetMetadata] = field(default_factory=list)
