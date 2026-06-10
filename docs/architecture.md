# Architecture Overview

## High-Level Design

```
┌─────────────────────────────────────────────────────┐
│                   Streamlit UI (app.py)              │
└──────┬──────────────────────────────────────────────┘
       │
       ▼
┌──────────────┐   ┌───────────────┐   ┌──────────────────┐
│  Ingestion   │──▶│   Profiling   │──▶│ Schema Matching  │
│  (loader.py) │   │ (profiler.py) │   │  (matcher.py)    │
└──────────────┘   └───────────────┘   └────────┬─────────┘
                                                 │
                        ┌────────────────────────┤
                        ▼                        ▼
               ┌─────────────────┐   ┌──────────────────┐
               │ Rationalization │   │  Consolidation   │
               │(rationalizer.py)│   │(consolidator.py) │
               └─────────────────┘   └────────┬─────────┘
                                               │
                   ┌───────────────────────────┤
                   ▼                           ▼
          ┌─────────────────┐      ┌───────────────────┐
          │   KPI Engine    │      │  Documentation    │
          │   (engine.py)   │      │  (generator.py)   │
          └────────┬────────┘      └───────────────────┘
                   │
                   ▼
          ┌─────────────────┐      ┌───────────────────┐
          │    Workbook     │      │  Reconciliation   │
          │   Generator     │      │  (reconciler.py)  │
          │ (generator.py)  │      └───────────────────┘
          └─────────────────┘
```

## Module Responsibilities

| Module | Package | Responsibility |
|--------|---------|----------------|
| Ingestion | `src/ingestion` | Parse `.xlsx`/`.xls` files into `WorkbookBundle` objects |
| Profiling | `src/profiling` | Column statistics, null rates, cardinality |
| Schema Matching | `src/schema_matching` | Fuzzy + AI column-name matching across workbooks |
| Rationalization | `src/rationalization` | Report-level keep/merge/retire recommendations |
| Consolidation | `src/consolidation` | Merge aligned sheets into a master DataFrame |
| KPI Engine | `src/kpi_engine` | Define and evaluate standardised metrics |
| Documentation | `src/documentation` | Data dictionary & lineage metadata generation |
| Reconciliation | `src/reconciliation` | Numeric totals cross-check (source vs output) |
| Workbook Generator | `src/workbook_generator` | Styled `.xlsx` output production |

## Technology Stack

- **Python 3.12**
- **Streamlit** — interactive UI
- **Pandas** — tabular data manipulation
- **OpenPyXL** — Excel read/write
- **RapidFuzz** — fast fuzzy string matching
- **Anthropic Claude API** — AI-assisted column matching and description generation
