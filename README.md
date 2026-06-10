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
**Upload & Ingest** or **Profile Reports** in the sidebar.

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

All 64 Phase 1 tests pass:

```
tests/test_ingestion.py   16 passed
tests/test_profiling.py   40 passed
tests/test_exporter.py     8 passed
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
│   │   └── loader.py               # load_workbook, load_many, WorkbookBundle
│   ├── profiling/
│   │   ├── metadata.py             # ColumnMetadata, SheetMetadata, WorkbookMetadata
│   │   ├── profiler.py             # profile_workbook, profile_sheet, build_summary_dataframes
│   │   ├── exporter.py             # export_profiling_results → profiling_summary.xlsx
│   │   └── cli.py                  # python -m src.profiling.cli
│   ├── ui/
│   │   └── pages/
│   │       └── profile_reports.py  # Streamlit Upload & Profile page
│   ├── schema_matching/
│   ├── rationalization/
│   ├── consolidation/
│   ├── kpi_engine/
│   ├── documentation/
│   ├── reconciliation/
│   └── workbook_generator/
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
