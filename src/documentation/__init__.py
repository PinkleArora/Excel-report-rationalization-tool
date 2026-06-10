"""Documentation module — data dictionaries and lineage metadata."""

from src.documentation.generator import (
    DataDictionaryEntry,
    build_data_dictionary,
    build_lineage_map,
    data_dictionary_to_df,
    export_to_excel,
)

__all__ = [
    "DataDictionaryEntry",
    "build_data_dictionary",
    "data_dictionary_to_df",
    "build_lineage_map",
    "export_to_excel",
]
