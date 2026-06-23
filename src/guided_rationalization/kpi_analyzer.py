"""Analyse KPI formulas to determine source-column dependencies.

For each formula in a configured KPI tab:
- Classifies the formula as SOURCE_BACKED, DERIVED, ROLLUP, or VALIDATION
- For SOURCE_BACKED formulas: resolves cross-sheet column references to canonical
  source-column names using the mapping produced by :mod:`source_analyzer`
- For DERIVED / ROLLUP / VALIDATION formulas: traces intra-sheet cell references
  back to SOURCE_BACKED cells and inherits their canonical columns automatically
- Reports which source columns are used by ≥1 KPI and which are never referenced

Formula types
-------------
SOURCE_BACKED  cross-sheet reference(s) pointing at the configured source tab
               e.g. =SUMIFS(SQL_data!C:C, SQL_data!E:E, ...)
DERIVED        arithmetic / logical combination of other KPI cells in the same sheet
               e.g. =B6+C6-D6,  =IF(C15>0, C15-D15, 0)
ROLLUP         aggregation of a same-sheet range
               e.g. =SUM(B6:B9),  =AVERAGE(C6:C10)
VALIDATION     cross-check / reconciliation formula referencing other KPI cells
               e.g. =E8-F8,  =ABS(D15-E15)
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Literal

import pandas as pd

from src.guided_rationalization.config import RationalizationConfig
from src.schema_matching.normalizer import normalize_column_name

logger = logging.getLogger(__name__)

# ── Regex patterns ────────────────────────────────────────────────────────────

# Cross-sheet reference:  'SheetName'!$COL$ROW  or  SheetName!COL
_SHEET_REF_RE = re.compile(
    r"'?([A-Za-z0-9_][\w\s]*?)'?\s*!\s*\$?([A-Z]+)\$?(\d*)"
)

# Intra-sheet cell reference after cross-sheet refs have been stripped:
# matches COL_LETTERS + ROW_DIGITS, e.g. B6, $C$15, AA100
# Requires a digit so bare function names (SUM, IF) don't match.
_INTRASHEET_CELL_RE = re.compile(r"\$?([A-Z]{1,3})\$?(\d{1,7})")

# Range separator inside a function arg, e.g. B6:B10
_RANGE_RE = re.compile(
    r"\$?([A-Z]{1,3})\$?(\d{1,7})\s*:\s*\$?([A-Z]{1,3})\$?(\d{1,7})"
)

_AGGREGATE_FUNCS = frozenset({
    "SUM", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS",
    "AVERAGE", "AVERAGEIF", "AVERAGEIFS",
    "MAX", "MIN", "MEDIAN",
    "SUMIF", "SUMIFS",
    "VLOOKUP", "HLOOKUP", "INDEX", "MATCH",
})

_ROLLUP_FUNCS = frozenset({"SUM", "AVERAGE", "MAX", "MIN", "MEDIAN", "COUNT", "COUNTA"})

FormulaType = Literal["SOURCE_BACKED", "DERIVED", "ROLLUP", "VALIDATION"]

# GETPIVOTDATA("FieldName", PivotRef!$A$1, ...)
_GETPIVOTDATA_RE = re.compile(
    r'GETPIVOTDATA\s*\(\s*"([^"]+)"\s*,\s*\'?([A-Za-z0-9_][\w\s]*?)\'?\s*!',
    re.IGNORECASE,
)


# ── Data models ───────────────────────────────────────────────────────────────

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
    # Canonical source-column names resolved from raw_refs (or inherited via tracing)
    canonical_source_columns: list[str] = field(default_factory=list)
    # True if every referenced sheet is the configured source tab for this workbook
    refs_source_tab_only: bool = True
    # Formula classification
    formula_type: FormulaType = "SOURCE_BACKED"
    # Intra-sheet cells this formula was traced through (DERIVED / ROLLUP)
    traced_from_cells: list[str] = field(default_factory=list)
    # KPI type: "FORMULA" | "GETPIVOTDATA" | "PIVOT" | "DERIVED"
    kpi_type: str = "FORMULA"
    # For GETPIVOTDATA: the sheet containing the pivot table
    pivot_tab: str = ""


@dataclass
class KpiLineage:
    """Data lineage path for a single KPI."""

    workbook_name: str
    kpi_tab: str
    kpi_label: str
    source_tab: str
    intermediate_tabs: list[str]   # e.g. ["Pivot_Reserve"]
    kpi_type: str                   # "FORMULA" | "GETPIVOTDATA"
    lineage_path: str               # "SQL_data -> Pivot_Reserve -> Summary"


@dataclass
class KpiAnalysisResult:
    """Aggregated KPI dependency analysis across all configured workbooks."""

    dependencies: list[KpiDependency]
    # All canonical column names referenced by >=1 KPI formula
    referenced_canonicals: set[str]
    # Canonical names that appear in source data but are never used by any KPI
    unreferenced_canonicals: set[str]
    # Data lineage paths
    lineage: list[KpiLineage] = field(default_factory=list)

    def to_dependency_dataframe(self) -> pd.DataFrame:
        rows = []
        for dep in self.dependencies:
            rows.append({
                "workbook_name":         dep.workbook_name,
                "kpi_tab":               dep.kpi_tab,
                "kpi_label":             dep.kpi_label,
                "cell_address":          dep.cell_address,
                "formula":               dep.formula,
                "formula_type":          dep.formula_type,
                "aggregate_function":    dep.aggregate_function,
                "source_columns_used":   ", ".join(dep.canonical_source_columns),
                "refs_source_tab_only":  dep.refs_source_tab_only,
                "traced_from_cells":     ", ".join(dep.traced_from_cells),
                "raw_refs":              " | ".join(
                    f"{s}!{c}" for s, c in dep.raw_refs
                ),
            })
        if not rows:
            return pd.DataFrame(columns=[
                "workbook_name", "kpi_tab", "kpi_label", "cell_address", "formula",
                "formula_type", "aggregate_function", "source_columns_used",
                "refs_source_tab_only", "traced_from_cells", "raw_refs",
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


# ── Column letter helpers ─────────────────────────────────────────────────────

def _col_letter_to_index(letters: str) -> int:
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _index_to_col_letter(idx: int) -> str:
    """Convert 0-based column index to Excel column letter (A, B, …, AA, …)."""
    result = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        result = chr(rem + ord("A")) + result
    return result


# ── Formula classification ────────────────────────────────────────────────────

def _classify_formula(formula: str, configured_source_tab: str) -> FormulaType:
    """Determine the formula type based on its reference pattern."""
    has_any_sheet_ref = bool(_SHEET_REF_RE.search(formula))

    if has_any_sheet_ref:
        # Check whether any cross-sheet ref points at the source tab
        for m in _SHEET_REF_RE.finditer(formula):
            sheet = m.group(1).strip()
            if sheet == configured_source_tab:
                return "SOURCE_BACKED"
        # Has cross-sheet refs but none to the source tab — treat as DERIVED
        # (e.g. references another KPI tab or an external lookup sheet)
        return "DERIVED"

    # No cross-sheet refs — pure intra-sheet formula
    upper = formula.upper()
    agg_match = re.match(r"=\s*([A-Z]+)\s*\(", upper)
    if agg_match and agg_match.group(1) in _ROLLUP_FUNCS:
        return "ROLLUP"

    return "DERIVED"


# ── Intra-sheet cell extraction ───────────────────────────────────────────────

def _extract_intrasheet_cells(formula: str) -> list[str]:
    """Return distinct same-sheet cell addresses referenced by *formula*.

    Cross-sheet references are stripped first so only bare cell addresses
    (B6, $C$15, etc.) are matched.  Range notation ``B6:B10`` is expanded
    up to a maximum of 20 cells; wider ranges use just the endpoints.
    """
    # Strip cross-sheet references to avoid mis-matching their column letters
    clean = _SHEET_REF_RE.sub(" ", formula)

    cells: list[str] = []
    seen: set[str] = set()

    def _add(col: str, row: str) -> None:
        addr = f"{col.upper()}{row}"
        if addr not in seen:
            seen.add(addr)
            cells.append(addr)

    # Expand ranges first
    for m in _RANGE_RE.finditer(clean):
        col_a, row_a, col_b, row_b = (
            m.group(1).upper(), int(m.group(2)),
            m.group(3).upper(), int(m.group(4)),
        )
        if col_a == col_b:
            # Same-column range: expand rows (cap at 20)
            r_min, r_max = min(row_a, row_b), max(row_a, row_b)
            for r in range(r_min, min(r_max + 1, r_min + 20)):
                _add(col_a, str(r))
        elif row_a == row_b:
            # Same-row range: expand columns (cap at 20)
            c_min = _col_letter_to_index(col_a)
            c_max = _col_letter_to_index(col_b)
            for c in range(c_min, min(c_max + 1, c_min + 20)):
                _add(_index_to_col_letter(c), str(row_a))
        else:
            _add(col_a, str(row_a))
            _add(col_b, str(row_b))

    # Remove range notation so individual cell re doesn't double-match
    clean_no_ranges = _RANGE_RE.sub(" ", clean)

    for m in _INTRASHEET_CELL_RE.finditer(clean_no_ranges):
        col, row = m.group(1).upper(), m.group(2)
        _add(col, row)

    return cells


# ── Lineage tracing for DERIVED / ROLLUP formulas ────────────────────────────

def _trace_derived_lineage(
    dependencies: list[KpiDependency],
    workbook_name: str,
    kpi_tab: str,
) -> None:
    """Propagate source columns from SOURCE_BACKED cells to DERIVED/ROLLUP cells.

    Iterates until stable (handles chains like E6→B6→SUMIFS) with a cap of
    10 rounds to guard against circular references.
    """
    # Build address → dependency map for cells in this workbook/tab
    cell_map: dict[str, KpiDependency] = {
        dep.cell_address: dep
        for dep in dependencies
        if dep.workbook_name == workbook_name and dep.kpi_tab == kpi_tab
    }

    changed = True
    rounds = 0
    while changed and rounds < 10:
        changed = False
        rounds += 1
        for dep in dependencies:
            if dep.workbook_name != workbook_name or dep.kpi_tab != kpi_tab:
                continue
            if dep.formula_type not in ("DERIVED", "ROLLUP", "VALIDATION"):
                continue

            referenced_cells = _extract_intrasheet_cells(dep.formula)
            inherited: list[str] = list(dep.canonical_source_columns)
            traced: list[str] = list(dep.traced_from_cells)

            for cell_addr in referenced_cells:
                ref_dep = cell_map.get(cell_addr)
                if ref_dep is None:
                    continue
                for col in ref_dep.canonical_source_columns:
                    if col not in inherited:
                        inherited.append(col)
                        changed = True
                if cell_addr not in traced:
                    traced.append(cell_addr)

            dep.canonical_source_columns = inherited
            dep.traced_from_cells = traced

    if rounds > 1:
        logger.debug(
            "Lineage tracing for %s/%s completed in %d round(s)",
            workbook_name, kpi_tab, rounds,
        )


# ── Main analysis function ────────────────────────────────────────────────────

def analyze_kpi_dependencies(
    bundles: list,
    config: RationalizationConfig,
    source_column_mapping: dict[tuple[str, str, str], str],
) -> KpiAnalysisResult:
    """Extract KPI formulas, classify them, and resolve source-column dependencies.

    SOURCE_BACKED formulas are resolved directly from cross-sheet column refs.
    DERIVED / ROLLUP / VALIDATION formulas have their source columns inherited
    automatically by tracing intra-sheet cell references back to SOURCE_BACKED cells.

    Args:
        bundles: All loaded WorkbookBundle objects.
        config: Rationalization configuration.
        source_column_mapping: ``(workbook, source_tab, original_col)`` → canonical name.

    Returns:
        KpiAnalysisResult with classified dependency list and referenced-column sets.
    """
    # Build: (workbook, source_tab) → col_letter → canonical_name
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
                tab_deps: list[KpiDependency] = []
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
                        tab_deps.append(dep)

                # Trace DERIVED/ROLLUP lineage for this tab
                _trace_derived_lineage(tab_deps, bundle.file_name, kpi_tab)
                dependencies.extend(tab_deps)
            else:
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

    n_source  = sum(1 for d in dependencies if d.formula_type == "SOURCE_BACKED")
    n_derived = sum(1 for d in dependencies if d.formula_type == "DERIVED")
    n_rollup  = sum(1 for d in dependencies if d.formula_type == "ROLLUP")
    n_valid   = sum(1 for d in dependencies if d.formula_type == "VALIDATION")

    logger.info(
        "KPI analysis: %d formula(s) [%d source-backed, %d derived, %d rollup, %d validation], "
        "%d source canonical(s) referenced, %d unreferenced",
        len(dependencies), n_source, n_derived, n_rollup, n_valid,
        len(referenced), len(unreferenced),
    )

    # Build lineage
    lineage: list[KpiLineage] = []
    seen_lineage: set[tuple[str, str, str]] = set()
    for dep in dependencies:
        key = (dep.workbook_name, dep.kpi_tab, dep.kpi_label)
        if key in seen_lineage:
            continue
        seen_lineage.add(key)
        wb_cfg = config.config_for(dep.workbook_name)
        source_tab = wb_cfg.source_tab if wb_cfg else ""
        intermediates: list[str] = []
        if dep.kpi_type == "GETPIVOTDATA" and dep.pivot_tab and dep.pivot_tab != source_tab:
            intermediates = [dep.pivot_tab]
        parts = [p for p in ([source_tab] + intermediates + [dep.kpi_tab]) if p]
        path = " -> ".join(parts)
        lineage.append(KpiLineage(
            workbook_name=dep.workbook_name,
            kpi_tab=dep.kpi_tab,
            kpi_label=dep.kpi_label,
            source_tab=source_tab,
            intermediate_tabs=intermediates,
            kpi_type=dep.kpi_type,
            lineage_path=path,
        ))

    return KpiAnalysisResult(
        dependencies=dependencies,
        referenced_canonicals=referenced,
        unreferenced_canonicals=unreferenced,
        lineage=lineage,
    )


# ── Formula parser ────────────────────────────────────────────────────────────

def _parse_formula_dependency(
    formula: str,
    workbook_name: str,
    kpi_tab: str,
    cell_address: str,
    kpi_label: str,
    configured_source_tab: str,
    header_map: dict[tuple[str, str], dict[str, str]],
) -> KpiDependency:
    formula_type = _classify_formula(formula, configured_source_tab)
    agg_func     = _detect_agg(formula)
    raw_refs: list[tuple[str, str]] = []
    canonical_cols: list[str] = []
    refs_source_only = True

    for m in _SHEET_REF_RE.finditer(formula):
        sheet = m.group(1).strip()
        col_letter = m.group(2).upper()
        raw_refs.append((sheet, col_letter))

        # A self-reference (formula on kpi_tab referencing kpi_tab, e.g.
        # Summary!$A6 used as a SUMIFS filter criterion) is NOT an external
        # dependency — it must not lower refs_source_only.
        if sheet != configured_source_tab and sheet != kpi_tab:
            refs_source_only = False

        if formula_type == "SOURCE_BACKED" and sheet == configured_source_tab:
            col_map = header_map.get((workbook_name, configured_source_tab), {})
            canonical = col_map.get(col_letter)
            if canonical and canonical not in canonical_cols:
                canonical_cols.append(canonical)

    # DERIVED / ROLLUP / VALIDATION: no cross-sheet source refs — vacuously fine.
    if formula_type != "SOURCE_BACKED":
        refs_source_only = True

    dep = KpiDependency(
        workbook_name=workbook_name,
        kpi_tab=kpi_tab,
        cell_address=cell_address,
        kpi_label=kpi_label,
        formula=formula,
        aggregate_function=agg_func,
        raw_refs=raw_refs,
        canonical_source_columns=canonical_cols,
        refs_source_tab_only=refs_source_only,
        formula_type=formula_type,
    )

    # Enrich with GETPIVOTDATA field extraction
    if "GETPIVOTDATA" in formula.upper():
        m = _GETPIVOTDATA_RE.search(formula)
        if m:
            field_name = m.group(1).strip()
            pivot_ref  = m.group(2).strip()
            dep.pivot_tab = pivot_ref
            dep.kpi_type  = "GETPIVOTDATA"
            if not dep.canonical_source_columns:
                # Add the field name as a canonical if not already resolved
                dep.canonical_source_columns.append(normalize_column_name(field_name))

    return dep


# ── Auxiliary helpers ─────────────────────────────────────────────────────────

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


def _get_raw_ws(raw_wb, sheet_name: str):
    if raw_wb is None:
        return None
    try:
        return raw_wb[sheet_name]
    except (KeyError, TypeError):
        return None
