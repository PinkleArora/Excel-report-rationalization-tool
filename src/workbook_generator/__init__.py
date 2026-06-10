"""Workbook generator — produce clean, formatted Excel workbooks."""

from src.workbook_generator.generator import SheetSpec, apply_house_style, generate_workbook
from src.workbook_generator.master_builder import (
    ALL_SHEET_NAMES,
    MasterWorkbookContext,
    build_master_workbook,
    build_master_workbook_bytes,
)

__all__ = [
    "SheetSpec",
    "generate_workbook",
    "apply_house_style",
    "build_master_workbook",
    "build_master_workbook_bytes",
    "MasterWorkbookContext",
    "ALL_SHEET_NAMES",
]
