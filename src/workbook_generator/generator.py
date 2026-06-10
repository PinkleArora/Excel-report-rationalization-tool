"""Produce clean, formatted Excel output workbooks from consolidated data."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


@dataclass
class SheetSpec:
    """Specification for a single output sheet."""

    name: str
    data: pd.DataFrame
    freeze_panes: str = "A2"
    auto_filter: bool = True
    column_widths: dict[str, int] = field(default_factory=dict)


def generate_workbook(
    sheets: list[SheetSpec],
    output_path: Path,
    include_summary_sheet: bool = True,
) -> Path:
    """Write *sheets* to an ``.xlsx`` file at *output_path*.

    Args:
        sheets: Ordered list of sheet specifications.
        output_path: Destination path for the generated workbook.
        include_summary_sheet: Prepend a cover/summary sheet if ``True``.

    Returns:
        The resolved :class:`pathlib.Path` of the written file.
    """
    raise NotImplementedError


def apply_house_style(workbook_path: Path) -> None:
    """Apply corporate colour palette, fonts, and border styles to *workbook_path* in-place."""
    raise NotImplementedError
