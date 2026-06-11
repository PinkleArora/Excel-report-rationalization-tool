"""Analyse KPI formulas to determine source-column dependencies.

For each formula in a configured KPI tab:
- Extracts the aggregate function, referenced sheet, and column letter(s)
- Resolves column letters to canonical source-column names using the source
  column mapping produced by :mod:`source_analyzer`
- Reports which source columns are used by ≥1 KPI and which are never referenced
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field

import pandas as pd

from src.guided_rationalization.config import RationalizationConfig
from src.schema_matching.normalizer import normalize_column_name

logger = logging.getLogger(__name__)

# Regex for cross-sheet formula references  SheetName!ColLetters[RowNum]
_SHEET_REF_RE = re.compile(
    r"'?([A-Za-z0-9_][\w\s]*?)'?\s*!\s*\$?([A-Z]+)\$?(\d*)"
)

_AGGREGATE_FUNCS = frozenset({
    "SUM", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS",
    "AVERAGE", "AVERAGEIF", "AVERAGEIFS",
    "MAX", "MIN", "MEDIAN",
    "SUMIF", "SUMIFS",
    "VLOOKUP", "HLOOKUP", "INDEX", "MATCH",
})


@dataclass
class KpiDependency:
    """A single KPI formula with its resolved source-column dependencies."""

    workbook_name: str
    kpi_tab: str
    cell_address: str
    kpi_label: str
    formula: str
    aggregate_function: str
    # Raw (sheet_name, col_letter) tuples from the formula
    raw_refs: list[tuple[str, str]] = field(default_factory=list)
    # Canonical source-column names resolved from raw_refs
    canonical_source_columns: list[str] = field(default_factory=list)
    # True if every referenced sheet is the configured source tab for this workbook
    refs_source_tab_only: bool = True


@dataclass
class KpiAnalysisResult:
    """Aggregated KPI dependency analysis across all configured workbooks."""

    dependencies: list[KpiDependency]
    # All canonical column names referenced by ≥1 KPI formula
    referenced_canonicals: set[str]
    # Canonical names that appear in source data but are never used by any KPI
    unreferenced_canonicals: set[str]

    def to_dependency_dataframe(self) -> pd.DataFrame:
        rows = []
        for dep in self.dependencies:
            rows.append({
                "workbook_name":         dep.workbook_name,
                "kpi_tab":               dep.kpi_tab,
                "kpi_label":             dep.kpi_label,
                "cell_address":          dep.cell_address,
                "formula":               dep.formula,
                "aggregate_function":    dep.aggregate_function,
                "source_columns_used":   ", ".join(dep.canonical_source_columns),
                "refs_source_tab_only":  dep.refs_source_tab_only,
                "raw_refs":              " | ".join(
                    f"{s}!{c}" for s, c in dep.raw_refs
                ),
            })
        if not rows:
            return pd.DataFrame(columns=[
                "workbook_name", "kpi_tab", "kpi_label", "cell_address", "formula",
                "aggregate_function", "source_columns_used", "refs_source_tab_only",
                "raw_refs",
            ])
        return pd.DataFrame(rows)

    def to_coverage_dataframe(self, all_canonicals: set[str]) -> pd.DataFrame:
        rows = []
        for canonical in sorted(all_canonicals):
            kpis_using = [
                d.kpi_label for d in self.dependencies
                if canonical in d.canonical_source_columns
            ]
            rows.append({
                "canonical_column":    canonical,
                "is_kpi_referenced":   canonical in self.referenced_canonicals,
                "used_by_n_kpis":      len(kpis_using),
                "used_by_kpis":        ", ".join(kpis_using),
                "recommendation":      (
                    "KEEP — required by KPI(s)"
                    if canonical in self.referenced_canonicals
                    else "REVIEW — not referenced by any KPI formula"
                ),
            })
        return pd.DataFrame(rows)


def _col_letter_to_index(letters: str) -> int:
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def analyze_kpi_dependencies(
    bundles: list,
    config: RationalizationConfig,
    source_column_mapping: dict[tuple[str, str, str], str],
) -> KpiAnalysisResult:
    """Extract KPI formulas and resolve their source-column dependencies.

    Args:
        bundles: All loaded :class:`~src.ingestion.loader.WorkbookBundle` objects.
        config: Rationalization configuration.
        source_column_mapping: ``(workbook, source_tab, original_col)`` → canonical name,
            produced by :func:`~source_analyzer.analyze_source_data`.

    Returns:
        :class:`KpiAnalysisResult` with dependency list and referenced-column sets.
    """
    # Build: workbook → source_tab → list of (col_letter_0_based, canonical_name)
    # so we can resolve formula column letters to canonical names
    header_map: dict[tuple[str, str], dict[str, str]] = {}
    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        source_tab = wb_cfg.source_tab
        df = bundle.sheets.get(source_tab)
        if df is None:
            continue
        col_letter_to_canonical: dict[str, str] = {}
        for col_idx, col_name in enumerate(df.columns):
            letter = _index_to_col_letter(col_idx)
            canonical = source_column_mapping.get(
                (bundle.file_name, source_tab, str(col_name)),
                normalize_column_name(str(col_name)),
            )
            col_letter_to_canonical[letter] = canonical
        header_map[(bundle.file_name, source_tab)] = col_letter_to_canonical

    dependencies: list[KpiDependency] = []

    for bundle in bundles:
        wb_cfg = config.config_for(bundle.file_name)
        if wb_cfg is None:
            continue
        raw_wb = getattr(bundle, "_raw_wb", None)

        for kpi_tab in wb_cfg.kpi_tabs:
            raw_ws = _get_raw_ws(raw_wb, kpi_tab)
            df = bundle.sheets.get(kpi_tab)
            if raw_ws is None and df is None:
                logger.warning("KPI tab '%s' not found in '%s'", kpi_tab, bundle.file_name)
                continue

            if raw_ws is not None:
                # Extract from openpyxl worksheet (has formula strings)
                for row in raw_ws.iter_rows():
                    for cell in row:
                        if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                            continue
                        formula = cell.value
                        label = _find_label(raw_ws, cell.row, cell.column)
                        dep = _parse_formula_dependency(
                            formula=formula,
                            workbook_name=bundle.file_name,
                            kpi_tab=kpi_tab,
                            cell_address=f"{cell.column_letter}{cell.row}",
                            kpi_label=label,
                            configured_source_tab=wb_cfg.source_tab,
                            header_map=header_map,
                        )
                        dependencies.append(dep)
            else:
                # No formula metadata — log warning, no KPI dependencies extractable
                logger.warning(
                    "No formula metadata for '%s/%s' — KPI dependency extraction skipped",
                    bundle.file_name, kpi_tab,
                )

    referenced: set[str] = {
        canonical
        for dep in dependencies
        for canonical in dep.canonical_source_columns
    }

    all_source_canonicals: set[str] = {
        canonical
        for (_, _, _), canonical in source_column_mapping.items()
    }
    unreferenced = all_source_canonicals - referenced

    logger.info(
        "KPI analysis: %d KPI formula(s), %d source canonical(s) referenced, "
        "%d unreferenced",
        len(dependencies), len(referenced), len(unreferenced),
    )
    return KpiAnalysisResult(
        dependencies=dependencies,
        referenced_canonicals=referenced,
        unreferenced_canonicals=unreferenced,
    )


def _parse_formula_dependency(
    formula: str,
    workbook_name: str,
    kpi_tab: str,
    cell_address: str,
    kpi_label: str,
    configured_source_tab: str,
    header_map: dict[tuple[str, str], dict[str, str]],
) -> KpiDependency:
    agg_func = _detect_agg(formula)
    raw_refs: list[tuple[str, str]] = []
    canonical_cols: list[str] = []
    refs_source_only = True

    for m in _SHEET_REF_RE.finditer(formula):
        sheet = m.group(1).strip()
        col_letter = m.group(2).upper()
        raw_refs.append((sheet, col_letter))

        if sheet != configured_source_tab:
            refs_source_only = False

        col_map = header_map.get((workbook_name, configured_source_tab), {})
        canonical = col_map.get(col_letter)
        if canonical and canonical not in canonical_cols:
            canonical_cols.append(canonical)

    return KpiDependency(
        workbook_name=workbook_name,
        kpi_tab=kpi_tab,
        cell_address=cell_address,
        kpi_label=kpi_label,
        formula=formula,
        aggregate_function=agg_func,
        raw_refs=raw_refs,
        canonical_source_columns=canonical_cols,
        refs_source_tab_only=refs_source_only,
    )


def _detect_agg(formula: str) -> str:
    m = re.match(r"=\s*([A-Z]+)\s*\(", formula.upper())
    if m and m.group(1) in _AGGREGATE_FUNCS:
        return m.group(1)
    return "OTHER"


def _find_label(raw_ws, row_idx: int, col_idx: int) -> str:
    if col_idx > 1:
        left = raw_ws.cell(row=row_idx, column=col_idx - 1)
        if left.value and not (isinstance(left.value, str) and left.value.startswith("=")):
            return str(left.value).strip()
    if row_idx > 1:
        above = raw_ws.cell(row=row_idx - 1, column=col_idx)
        if above.value and not (isinstance(above.value, str) and above.value.startswith("=")):
            return str(above.value).strip()
    return f"KPI_R{row_idx}C{col_idx}"


def _index_to_col_letter(idx: int) -> str:
    """Convert 0-based column index to Excel column letter (A, B, …, AA, …)."""
    result = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        result = chr(rem + ord("A")) + result
    return result


def _get_raw_ws(raw_wb, sheet_name: str):
    if raw_wb is None:
        return None
    try:
        return raw_wb[sheet_name]
    except (KeyError, TypeError):
        return None
