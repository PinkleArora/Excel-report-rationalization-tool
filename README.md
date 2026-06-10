# Excel Report Rationalization & BAU Standardization Tool

An AI-assisted platform for analysing, deduplicating, and standardising Excel-based
Business-As-Usual (BAU) reports across an organisation.

## Features

- **Ingestion** — parse `.xlsx`, `.xlsm`, `.xls`, and `.csv` files into structured DataFrames
- **Profiling** — column statistics, null-rate analysis, semantic type detection, formula counting
- **Schema Matching** — fuzzy + AI-powered detection of equivalent columns across reports
- **Rationalization** — per-report recommendations: keep, merge, or retire
- **Consolidation** — merge aligned sheets into a single authoritative dataset
- **KPI Engine** — define and evaluate standardised metrics across consolidated data
- **Workbook Generator** — produce clean, styled `.xlsx` outputs ready for distribution
- **Documentation** — auto-generated data dictionaries and column-level lineage maps
- **Reconciliation** — numeric cross-check validating output totals against source files

---

## Quick Start

```bash
python3.12 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Open [http://localhost:8501](http://localhost:8501) in your browser, then navigate to
**Profile Reports** or **Workbook Generator** in the sidebar.

---

## Phase 3: Master Workbook Generator

### What it produces

`output/master_workbook.xlsx` — the team's consolidated BAU workbook:

| Sheet | Contents |
|-------|---------|
| `01_Master_Data` | All source rows with `_source_workbook` / `_source_sheet` lineage columns |
| `02_KPI_Summary` | Numeric aggregates (count, sum, mean, min, max) per source and overall |
| `03_Source_Mapping` | Canonical column mapping for every source column, with match type and score |
| `04_Data_Dictionary` | Auto-generated: types, sample values, null rates, source workbooks |
| `05_Reconciliation` | Numeric totals cross-check: source vs master — PASS / WARN / FAIL with traffic-light styling |
| `06_Issues_Log` | Auto-detected issues (high nulls, empty columns, formula cells, recon failures) |
| `07_Rationalization_Report` | Keep / Merge / Retire recommendation per source workbook |
| `08_Update_SOP` | Standard Operating Procedure for monthly master workbook maintenance |

### Run via Streamlit UI

```bash
streamlit run app.py
# Navigate to "Workbook Generator", upload files, click "Generate Master Workbook"
```

### Run via Python API

```python
from src.ingestion.loader import load_many
from src.workbook_generator.master_builder import build_master_workbook

bundles = load_many(["sales_report.xlsx", "hr_report.xlsx", "finance.csv"])
path = build_master_workbook(
    bundles,
    output_path="output/master_workbook.xlsx",
    matching_threshold=80.0,       # RapidFuzz score cutoff for column matching
    reconciliation_tolerance=0.01, # acceptable delta = 1%
)
print(f"Master workbook: {path}")
```

### Schema Matching

Columns are matched across workbooks using a two-pass strategy:

1. **Exact match** — normalised names (`Revenue (£)` → `revenue`) compared for equality → score 100
2. **Fuzzy match** — RapidFuzz `WRatio` on remaining columns above the configurable threshold (default 80)

Matched columns are unified under a single canonical name in `01_Master_Data` and documented in `03_Source_Mapping`.

### Rationalization Logic

| Overlap % | Recommendation |
|-----------|---------------|
| ≥ 80 % | **RETIRE** — substantially redundant |
| 50–79 % | **MERGE** — consolidate with best-matched partner |
| 25–49 % | **REVIEW** — partial consolidation warranted |
| < 25 % | **KEEP** — unique data, retain as-is |

---

## Phase 1: Workbook Profiling & Metadata Extraction

### What it does

Accepts one or more Excel / CSV files and extracts:

| Metric | Scope |
|--------|-------|
| Row count, column count | Sheet |
| Column names | Sheet |
| Inferred dtype + semantic type (`integer`, `float`, `text`, `datetime`, `boolean`, `empty`) | Column |
| Non-null count, null count, null % | Column |
| Unique count | Column |
| Sample values (up to 5 non-null) | Column |
| Min / max / mean (numeric columns) | Column |
| Formula cell count (`.xlsx`/`.xlsm` only) | Sheet |
| Empty column count | Sheet |

Results are written to **`output/profiling_summary.xlsx`** with three sheets:

- `Workbook_Summary` — one row per input file
- `Sheet_Summary` — one row per worksheet
- `Column_Summary` — one row per column across all sheets

### Run via Streamlit UI

```bash
streamlit run app.py
# Navigate to "Profile Reports", upload files, click "Run Profiling"
```

### Run via CLI

```bash
python -m src.profiling.cli path/to/report1.xlsx path/to/report2.csv \
    --output output/profiling_summary.xlsx \
    --log-level INFO
```

### Run via Python API

```python
from src.ingestion.loader import load_many
from src.profiling.profiler import profile_many
from src.profiling.exporter import export_profiling_results

bundles = load_many(["report1.xlsx", "report2.csv"])
metas   = profile_many(bundles)
path    = export_profiling_results(metas, output_path="output/profiling_summary.xlsx")
print(f"Written to: {path}")
```

---

## Running Tests

```bash
pytest tests/ -v --cov=src
```

All 183 tests pass across Phase 1 and Phase 3:

```
tests/test_ingestion.py        16 passed   Phase 1
tests/test_profiling.py        40 passed   Phase 1
tests/test_exporter.py          8 passed   Phase 1
tests/test_schema_matching.py  37 passed   Phase 3
tests/test_consolidation.py    22 passed   Phase 3
tests/test_rationalization.py  16 passed   Phase 3
tests/test_reconciliation.py   18 passed   Phase 3
tests/test_master_builder.py   26 passed   Phase 3
```

---

## Project Structure

```
.
├── app.py                          # Streamlit UI entry point
├── requirements.txt                # Python dependencies
├── pytest.ini
├── output/                         # Generated outputs (git-ignored)
├── sample_data/                    # Synthetic example files
├── src/
│   ├── ingestion/
│   │   └── loader.py                   # load_workbook, load_many, WorkbookBundle
│   ├── profiling/
│   │   ├── metadata.py                 # ColumnMetadata, SheetMetadata, WorkbookMetadata
│   │   ├── profiler.py                 # profile_workbook, profile_sheet, build_summary_dataframes
│   │   ├── exporter.py                 # export_profiling_results → profiling_summary.xlsx
│   │   └── cli.py                      # python -m src.profiling.cli
│   ├── schema_matching/
│   │   ├── normalizer.py               # normalize_column_name, canonical_for_group
│   │   └── matcher.py                  # ColumnMatch, match_columns, match_all_bundles
│   ├── consolidation/
│   │   └── consolidator.py             # consolidate, resolve_conflicts, build_source_mapping_df
│   ├── rationalization/
│   │   └── rationalizer.py             # assess_redundancy, build_inventory, ReportRecommendation
│   ├── documentation/
│   │   └── generator.py               # build_data_dictionary, build_lineage_map
│   ├── reconciliation/
│   │   └── reconciler.py              # reconcile, reconciliation_summary, ReconciliationResult
│   ├── workbook_generator/
│   │   ├── styles.py                   # Shared openpyxl styling helpers
│   │   ├── generator.py                # SheetSpec, generate_workbook, apply_house_style
│   │   └── master_builder.py           # build_master_workbook — full 8-sheet pipeline
│   └── ui/
│       └── pages/
│           ├── profile_reports.py      # Streamlit Phase 1 page
│           └── generate_master.py      # Streamlit Phase 3 page
├── tests/
│   ├── test_ingestion.py
│   ├── test_profiling.py
│   └── test_exporter.py
└── docs/
    ├── architecture.md
    └── developer_guide.md
```

---

## Tech Stack

| Library | Version | Purpose |
|---------|---------|---------|
| Python | 3.12 | Runtime |
| Streamlit | ≥ 1.35 | Interactive UI |
| Pandas | ≥ 2.2 | Tabular data manipulation |
| OpenPyXL | ≥ 3.1 | Excel read / write / style |
| xlrd | ≥ 2.0 | Legacy `.xls` support |
| RapidFuzz | ≥ 3.9 | Fuzzy column matching (Phase 2+) |
| Anthropic SDK | ≥ 0.28 | AI-assisted matching & descriptions (Phase 2+) |

---

## Contributing

See [docs/developer_guide.md](docs/developer_guide.md) for setup and contribution workflow.

## License

See [LICENSE](LICENSE).
