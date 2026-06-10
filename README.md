# Excel Report Rationalization & BAU Standardization Tool

An AI-assisted platform for analysing, deduplicating, and standardising Excel-based
Business-As-Usual (BAU) reports across an organisation.

## Features

- **Ingestion** — parse one or many `.xlsx` / `.xls` workbooks into structured DataFrames
- **Profiling** — automatic column statistics, null-rate analysis, and cardinality summaries
- **Schema Matching** — fuzzy + AI-powered detection of equivalent columns across reports
- **Rationalization** — per-report recommendations: keep, merge, or retire
- **Consolidation** — merge aligned sheets into a single authoritative dataset
- **KPI Engine** — define and evaluate standardised metrics across consolidated data
- **Workbook Generator** — produce clean, styled `.xlsx` outputs ready for distribution
- **Documentation** — auto-generated data dictionaries and column-level lineage maps
- **Reconciliation** — numeric cross-check validating output totals against source files

## Quick Start

```bash
python3.12 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Open [http://localhost:8501](http://localhost:8501) in your browser.

## Running Tests

```bash
pytest tests/ -v --cov=src
```

## Project Structure

```
.
├── app.py                    # Streamlit UI entry point
├── requirements.txt          # Python dependencies
├── src/
│   ├── ingestion/
│   ├── profiling/
│   ├── schema_matching/
│   ├── rationalization/
│   ├── consolidation/
│   ├── kpi_engine/
│   ├── documentation/
│   ├── reconciliation/
│   └── workbook_generator/
├── tests/
├── docs/
│   ├── architecture.md
│   └── developer_guide.md
└── sample_data/
```

## Tech Stack

| Library | Version | Purpose |
|---------|---------|---------|
| Python | 3.12 | Runtime |
| Streamlit | ≥ 1.35 | Interactive UI |
| Pandas | ≥ 2.2 | Tabular data |
| OpenPyXL | ≥ 3.1 | Excel I/O |
| RapidFuzz | ≥ 3.9 | Fuzzy column matching |
| Anthropic SDK | ≥ 0.28 | AI-assisted matching & descriptions |

## Contributing

See [docs/developer_guide.md](docs/developer_guide.md) for setup and contribution instructions.

## License

See [LICENSE](LICENSE).
