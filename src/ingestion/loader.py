"""Load Excel workbooks into a structured in-memory representation."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


@dataclass
class WorkbookBundle:
    """Holds all sheets parsed from a single Excel file."""

    source_path: Path
    sheets: dict[str, pd.DataFrame] = field(default_factory=dict)

    @property
    def sheet_names(self) -> list[str]:
        return list(self.sheets.keys())


def load_workbook(path: str | Path) -> WorkbookBundle:
    """Parse every sheet of *path* into a :class:`WorkbookBundle`.

    Args:
        path: Absolute or relative path to an ``.xlsx`` / ``.xls`` file.

    Returns:
        A :class:`WorkbookBundle` containing one :class:`pandas.DataFrame`
        per sheet, keyed by sheet name.

    Raises:
        FileNotFoundError: If *path* does not exist.
        ValueError: If the file extension is not supported.
    """
    raise NotImplementedError


def load_many(paths: list[str | Path]) -> list[WorkbookBundle]:
    """Load multiple workbooks and return one bundle per file."""
    raise NotImplementedError
