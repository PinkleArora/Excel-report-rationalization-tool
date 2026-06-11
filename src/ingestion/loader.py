"""Load Excel and CSV workbooks into a structured in-memory representation."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".csv"}


@dataclass
class WorkbookBundle:
    """All sheets parsed from a single file, keyed by sheet name."""

    source_path: Path
    sheets: dict[str, pd.DataFrame] = field(default_factory=dict)
    # openpyxl Workbook object retained for formula / style inspection (None for CSV/xls)
    _raw_wb: object = field(default=None, repr=False, compare=False)

    @property
    def sheet_names(self) -> list[str]:
        return list(self.sheets.keys())

    @property
    def file_name(self) -> str:
        return self.source_path.name


def load_workbook(path: str | Path) -> WorkbookBundle:
    """Parse every sheet of *path* into a :class:`WorkbookBundle`.

    Supported formats: ``.xlsx``, ``.xlsm``, ``.xls``, ``.csv``.

    Args:
        path: Absolute or relative path to the file.

    Returns:
        A :class:`WorkbookBundle` with one DataFrame per sheet.

    Raises:
        FileNotFoundError: If *path* does not exist.
        ValueError: If the file extension is not supported.
    """
    path = Path(path)

    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    ext = path.suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported file type '{ext}'. "
            f"Supported: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
        )

    logger.info("Loading workbook: %s", path)

    try:
        if ext == ".csv":
            return _load_csv(path)
        elif ext == ".xls":
            return _load_xls(path)
        else:
            return _load_xlsx(path)
    except Exception as exc:
        logger.error("Failed to load %s: %s", path, exc)
        raise


def load_workbook_from_bytes(data: bytes, file_name: str) -> WorkbookBundle:
    """Load an Excel workbook from raw bytes (e.g. a Streamlit upload).

    Args:
        data: Raw bytes of the .xlsx / .xlsm file.
        file_name: Original file name (used as the bundle identifier).

    Returns:
        A :class:`WorkbookBundle` with one DataFrame per sheet.
    """
    import io
    import openpyxl

    buf = io.BytesIO(data)
    raw_wb = openpyxl.load_workbook(buf, data_only=False)
    sheets: dict[str, pd.DataFrame] = {}
    for sheet_name in raw_wb.sheetnames:
        buf.seek(0)
        df = pd.read_excel(buf, sheet_name=sheet_name, engine="openpyxl", header=0)
        sheets[sheet_name] = df

    # Use a synthetic Path so file_name is preserved
    fake_path = Path(file_name)
    logger.info("Loaded %d sheet(s) from bytes (%s)", len(sheets), file_name)
    return WorkbookBundle(source_path=fake_path, sheets=sheets, _raw_wb=raw_wb)


def load_many(paths: list[str | Path]) -> list[WorkbookBundle]:
    """Load multiple files, skipping those that fail with a logged warning.

    Args:
        paths: Iterable of file paths.

    Returns:
        List of successfully loaded :class:`WorkbookBundle` objects.
    """
    bundles: list[WorkbookBundle] = []
    for p in paths:
        try:
            bundles.append(load_workbook(p))
        except Exception as exc:
            logger.warning("Skipping %s — %s", p, exc)
    return bundles


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _load_xlsx(path: Path) -> WorkbookBundle:
    """Load .xlsx / .xlsm using openpyxl (preserves formula metadata)."""
    import openpyxl  # local import keeps module importable without openpyxl installed

    raw_wb = openpyxl.load_workbook(path, data_only=False)
    sheets: dict[str, pd.DataFrame] = {}

    for sheet_name in raw_wb.sheetnames:
        df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl", header=0)
        sheets[sheet_name] = df
        logger.debug("  Sheet '%s': %d rows × %d cols", sheet_name, len(df), len(df.columns))

    logger.info("Loaded %d sheet(s) from %s", len(sheets), path.name)
    return WorkbookBundle(source_path=path, sheets=sheets, _raw_wb=raw_wb)


def _load_xls(path: Path) -> WorkbookBundle:
    """Load legacy .xls using xlrd."""
    xl = pd.ExcelFile(path, engine="xlrd")
    sheets = {
        name: xl.parse(name)
        for name in xl.sheet_names
    }
    logger.info("Loaded %d sheet(s) from %s (xls)", len(sheets), path.name)
    return WorkbookBundle(source_path=path, sheets=sheets)


def _load_csv(path: Path) -> WorkbookBundle:
    """Load a CSV as a single-sheet bundle (sheet name = file stem)."""
    df = pd.read_csv(path)
    sheet_name = path.stem
    logger.info("Loaded CSV '%s': %d rows × %d cols", path.name, len(df), len(df.columns))
    return WorkbookBundle(source_path=path, sheets={sheet_name: df})
