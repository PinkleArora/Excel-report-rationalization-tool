# Developer Guide

## Setup

```bash
# 1. Clone and enter the repository
git clone <repo-url>
cd Excel-report-rationalization-tool

# 2. Create a virtual environment
python3.12 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
streamlit run app.py
```

## Running Tests

```bash
pytest tests/ -v --cov=src
```

## Project Layout

```
.
├── app.py                    # Streamlit entry point
├── requirements.txt
├── sample_data/              # Example Excel files for development
├── src/
│   ├── ingestion/            # load_workbook, WorkbookBundle
│   ├── profiling/            # profile_sheet, SheetProfile
│   ├── schema_matching/      # match_columns, ColumnMatch
│   ├── rationalization/      # assess_redundancy, ReportRecommendation
│   ├── consolidation/        # consolidate, resolve_conflicts
│   ├── kpi_engine/           # KPIRegistry, evaluate_kpis
│   ├── documentation/        # build_data_dictionary, build_lineage_map
│   ├── reconciliation/       # reconcile, ReconciliationResult
│   └── workbook_generator/   # generate_workbook, SheetSpec
├── tests/                    # pytest test suite
└── docs/                     # Architecture and developer docs
```

## Contribution Workflow

1. Create a feature branch from `main`.
2. Implement business logic inside the appropriate `src/` module.
3. Add or update tests in `tests/`.
4. Run `ruff check .` and `pytest` before committing.
5. Open a pull request with a clear description.

## Environment Variables

| Variable | Purpose |
|----------|---------|
| `ANTHROPIC_API_KEY` | API key for Claude AI features |
