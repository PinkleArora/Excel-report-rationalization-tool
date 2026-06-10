"""Profiling module — summarise column statistics and data quality metrics.

Public surface:
    - :class:`~src.profiling.metadata.ColumnMetadata`
    - :class:`~src.profiling.metadata.SheetMetadata`
    - :class:`~src.profiling.metadata.WorkbookMetadata`
    - :func:`~src.profiling.profiler.profile_workbook`
    - :func:`~src.profiling.profiler.profile_sheet`
    - :func:`~src.profiling.profiler.profile_many`
    - :func:`~src.profiling.profiler.build_summary_dataframes`
    - :func:`~src.profiling.exporter.export_profiling_results`
"""

from src.profiling.exporter import export_profiling_results
from src.profiling.metadata import ColumnMetadata, SheetMetadata, WorkbookMetadata
from src.profiling.profiler import (
    build_summary_dataframes,
    profile_many,
    profile_sheet,
    profile_workbook,
)

__all__ = [
    "ColumnMetadata",
    "SheetMetadata",
    "WorkbookMetadata",
    "profile_workbook",
    "profile_sheet",
    "profile_many",
    "build_summary_dataframes",
    "export_profiling_results",
]
