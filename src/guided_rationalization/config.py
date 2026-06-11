"""Data models for the Guided Rationalization workflow."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class WorkbookTabConfig:
    """User-specified tab assignments for a single uploaded workbook."""

    workbook_name: str     # matches bundle.file_name
    source_tab: str        # the one tab that holds raw, transactional records
    kpi_tabs: list[str]    # one or more formula/summary tabs
    lob_identifier: str = ""   # free-text LOB label added to every row in master source


@dataclass
class RationalizationConfig:
    """Complete user configuration for one rationalization run."""

    workbook_configs: list[WorkbookTabConfig]

    # --- Output tab naming ---
    future_source_tab_name: str = "Master_Source_Data"
    # workbook_name → desired output KPI tab name; missing entries get auto-names
    future_kpi_tab_names: dict[str, str] = field(default_factory=dict)

    # --- Merge / filter options ---
    matching_threshold: float = 80.0
    # Only exact-normalized matches are merged by default.
    # Set True to also merge HIGH-confidence fuzzy matches (score ≥ 95).
    merge_high_confidence: bool = False
    # If True, columns not referenced by any KPI formula are dropped from master.
    remove_unused_columns: bool = True

    def kpi_tab_name_for(self, workbook_name: str, tab_name: str) -> str:
        """Return the configured output tab name for a KPI tab, or an auto-generated one."""
        explicit = self.future_kpi_tab_names.get(workbook_name)
        if explicit:
            return explicit
        stem = workbook_name.rsplit(".", 1)[0]
        return f"{stem}_{tab_name}"[:31]  # Excel tab name max 31 chars

    def config_for(self, workbook_name: str) -> WorkbookTabConfig | None:
        for cfg in self.workbook_configs:
            if cfg.workbook_name == workbook_name:
                return cfg
        return None
