"""Produce clean, formatted Excel output workbooks from consolidated data."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook as openpyxl_load

from src.workbook_generator.styles import apply_full_style, auto_column_width, style_header_row

logger = logging.getLogger(__name__)


@dataclass
class SheetSpec:
    """Specification for a single output sheet."""

    name: str
    data: pd.DataFrame
    freeze_panes: str = "A2"
    auto_filter: bool = True
    column_widths: dict[str, int] = field(default_factory=dict)
    status_col: str | None = None   # column letter that holds PASS/WARN/FAIL values


def generate_workbook(
    sheets: list[SheetSpec],
    output_path: str | Path,
    include_cover_sheet: bool = False,
) -> Path:
    """Write *sheets* to a formatted ``.xlsx`` file at *output_path*.

    Args:
        sheets: Ordered list of :class:`SheetSpec` objects.
        output_path: Destination path for the generated file.
        include_cover_sheet: Prepend a cover/index sheet listing all sheets.

    Returns:
        Resolved absolute :class:`pathlib.Path` of the written file.

    Raises:
        ValueError: If *sheets* is empty.
    """
    if not sheets:
        raise ValueError("generate_workbook: at least one SheetSpec is required.")

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if include_cover_sheet:
            _write_cover_sheet(writer, sheets)
        for spec in sheets:
            spec.data.to_excel(writer, sheet_name=spec.name, index=False)

    # Post-process with openpyxl for styling
    apply_house_style(output_path, sheets)

    logger.info("Workbook written: %s (%d sheet(s))", output_path, len(sheets))
    return output_path.resolve()


def apply_house_style(
    workbook_path: str | Path,
    specs: list[SheetSpec] | None = None,
) -> None:
    """Apply corporate colour palette, fonts, and filters to *workbook_path* in-place.

    Args:
        workbook_path: Path to the ``.xlsx`` file to style.
        specs: Optional list of :class:`SheetSpec` — used to apply per-sheet
            options (column widths, status column highlighting).
    """
    workbook_path = Path(workbook_path)
    wb = openpyxl_load(workbook_path)

    spec_by_name: dict[str, SheetSpec] = {s.name: s for s in (specs or [])}

    for ws in wb.worksheets:
        spec = spec_by_name.get(ws.title)
        freeze = spec.freeze_panes if spec else "A2"
        status_col = spec.status_col if spec else None
        apply_full_style(ws, freeze_cell=freeze, status_col_letter=status_col)

        # Override explicit column widths from spec
        if spec and spec.column_widths:
            for col_letter, width in spec.column_widths.items():
                ws.column_dimensions[col_letter].width = width

    wb.save(workbook_path)


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _write_cover_sheet(writer: pd.ExcelWriter, sheets: list[SheetSpec]) -> None:
    cover = pd.DataFrame(
        {"Sheet": [s.name for s in sheets], "Description": ["" for _ in sheets]},
    )
    cover.to_excel(writer, sheet_name="Contents", index=False)
