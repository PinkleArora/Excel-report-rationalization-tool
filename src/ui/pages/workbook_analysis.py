"""Streamlit page: Workbook Analysis Report."""

from __future__ import annotations

import logging
import tempfile
from pathlib import Path

import streamlit as st

logger = logging.getLogger(__name__)


def render() -> None:
    st.header("Workbook Analysis")
    st.caption(
        "Upload one or more workbooks to generate a detailed analysis report — "
        "**workbook_analysis.xlsx** — showing exactly how the application interpreted "
        "each file: tab classification, KPI extraction, schema mapping, collision "
        "detection, and rationalization verdicts."
    )

    with st.expander("What this report contains", expanded=False):
        st.markdown(
            """
            | Sheet | Contents |
            |-------|---------|
            | **01_Workbook_Inventory** | One row per uploaded workbook — sheet counts, row counts, tab type breakdown |
            | **02_Source_Tab_Detection** | Tabs classified as Source Data, Reference Data, or Mapping Table, with signal scores |
            | **03_Summary_Tab_Detection** | Tabs classified as KPI/Summary, Dashboard, Output, or Calculation, with signal scores |
            | **04_KPI_Inventory** | All formulas extracted from formula tabs — label, function, referenced source fields, rationalization verdict |
            | **05_Formula_Inventory** | Cell-level formula listing from all formula-bearing tabs |
            | **06_Schema_Mapping** | Every source column with match confidence (EXACT / HIGH / REVIEW), merge verdict, and matched column |
            | **07_Collision_Analysis** | Intra-frame collisions and many-to-one mapping flags with severity and remediation guidance |
            | **08_Rationalization_Candidates** | KPIs confirmed as consolidation candidates (same label + function + shared source fields) |
            | **09_Rationalization_Exclusions** | KPIs appearing in multiple workbooks but excluded — with the specific reason |

            > **No data is sent externally.** All analysis runs locally on uploaded files.
            """
        )

    # --- File upload ---
    uploaded = st.file_uploader(
        "Upload workbooks to analyse",
        type=["xlsx", "xlsm", "xls", "csv"],
        accept_multiple_files=True,
        key="analysis_uploader",
    )

    with st.expander("Advanced options", expanded=False):
        threshold = st.slider(
            "Schema match threshold",
            min_value=50, max_value=100, value=80, step=5,
            help=(
                "Minimum RapidFuzz WRatio score to include a fuzzy column match "
                "in the Schema Mapping sheet. Does NOT affect which columns are "
                "merged — only EXACT matches (score=100) drive column unification."
            ),
        )
        min_source_rows = st.number_input(
            "Minimum rows for source-data classification",
            min_value=1, max_value=500, value=10, step=5,
            help="Sheets with fewer rows than this threshold will not be classified as Source Data.",
        )
        kpi_density = st.slider(
            "KPI formula density threshold",
            min_value=0.05, max_value=0.80, value=0.25, step=0.05,
            help="Formula density (formula cells / non-empty cells) above which a tab is a KPI/Summary candidate.",
        )

    if not uploaded:
        st.info("Upload at least one file to begin.")
        _render_sample_explanation()
        return

    st.divider()

    if st.button("Run Analysis", type="primary", use_container_width=True):
        _run_and_display(uploaded, threshold, int(min_source_rows), float(kpi_density))


def _run_and_display(
    uploaded_files,
    threshold: int,
    min_source_rows: int,
    kpi_density: float,
) -> None:
    from src.ingestion.loader import load_workbook
    from src.ingestion.workbook_analyzer import ClassificationConfig
    from src.workbook_generator.analysis_builder import (
        AnalysisContext,
        _run_analysis,
        _sheet_collision_analysis,
        _sheet_formula_inventory,
        _sheet_kpi_inventory,
        _sheet_rationalization_candidates,
        _sheet_rationalization_exclusions,
        _sheet_schema_mapping,
        _sheet_source_tabs,
        _sheet_summary_tabs,
        _sheet_workbook_inventory,
        build_analysis_workbook_bytes,
    )

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        bundles = []
        load_errors = []

        with st.spinner("Loading files…"):
            for uf in uploaded_files:
                dest = tmp / uf.name
                dest.write_bytes(uf.read())
                try:
                    bundles.append(load_workbook(dest))
                except Exception as exc:
                    load_errors.append(f"**{uf.name}**: {exc}")

        if load_errors:
            for err in load_errors:
                st.error(err)
        if not bundles:
            st.error("No files could be loaded.")
            return

        cfg = ClassificationConfig(
            min_source_rows=min_source_rows,
            kpi_formula_density_min=kpi_density,
        )

        with st.spinner("Analysing workbook structure…"):
            try:
                ctx = _run_analysis(bundles, threshold, cfg)
            except Exception as exc:
                st.error(f"Analysis failed: {exc}")
                logger.exception("Analysis error")
                return

        # --- Summary metrics ---
        total_tabs = sum(len(a.tab_analyses) for a in ctx.analyses)
        source_tabs = sum(len(a.source_tabs) for a in ctx.analyses)
        summary_tabs = sum(len(a.summary_tabs) for a in ctx.analyses)
        total_kpis = sum(len(a.kpi_definitions) for a in ctx.analyses)
        safe_matches = sum(1 for m in ctx.column_matches if m.safe_to_merge)
        review_matches = sum(1 for m in ctx.column_matches if not m.safe_to_merge)
        collisions = len(ctx.collision_log) + len(ctx.many_to_one_issues)

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Workbooks", len(bundles))
        col2.metric("Tabs analysed", total_tabs,
                    delta=f"{source_tabs} source / {summary_tabs} summary",
                    delta_color="off")
        col3.metric("KPIs extracted", total_kpis)
        col4.metric("Column matches", f"{safe_matches} safe / {review_matches} review",
                    delta_color="off")
        col5.metric("Collision flags", collisions,
                    delta_color="inverse" if collisions > 0 else "off")

        st.divider()

        # --- Inline previews ---
        _preview_section(
            "Workbook Inventory",
            _sheet_workbook_inventory(ctx),
            "tab_type_detail",
        )
        _preview_section(
            "Source Tab Detection",
            _sheet_source_tabs(ctx),
        )
        _preview_section(
            "Summary / KPI Tab Detection",
            _sheet_summary_tabs(ctx),
        )
        _preview_section(
            "KPI Inventory",
            _sheet_kpi_inventory(ctx),
            help_text=(
                "**rationalization_candidate = True** means this KPI shares label, "
                "aggregate function, and referenced source columns with a KPI in "
                "another workbook — safe to consider for consolidation.  "
                "**rationalization_candidate = False** means it was excluded — "
                "see the Exclusions tab for the reason."
            ),
        )
        _preview_section(
            "Schema Mapping",
            _sheet_schema_mapping(ctx),
            help_text=(
                "**safe_to_merge = True** (EXACT confidence) → column renamed to canonical name in master data.  "
                "**review_required = True** → documented only; human confirmation needed before merging."
            ),
        )
        _preview_section(
            "Collision Analysis",
            _sheet_collision_analysis(ctx),
        )
        _preview_section(
            "Rationalization Candidates",
            _sheet_rationalization_candidates(ctx),
        )
        _preview_section(
            "Rationalization Exclusions",
            _sheet_rationalization_exclusions(ctx),
        )

        # --- Download button ---
        st.divider()
        with st.spinner("Generating workbook_analysis.xlsx…"):
            try:
                xlsx_bytes = build_analysis_workbook_bytes(
                    bundles,
                    matching_threshold=threshold,
                    classification_config=cfg,
                )
            except Exception as exc:
                st.error(f"Failed to generate workbook: {exc}")
                logger.exception("Workbook generation error")
                return

        st.download_button(
            label="Download Analysis Report (workbook_analysis.xlsx)",
            data=xlsx_bytes,
            file_name="workbook_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.success(
            f"Analysis complete. {total_tabs} tabs classified, {total_kpis} KPIs extracted, "
            f"{safe_matches} safe column merges, {collisions} collision flag(s)."
        )


def _preview_section(
    title: str,
    df,
    detail: str = "",
    help_text: str = "",
) -> None:
    with st.expander(f"Preview: {title}  ({len(df)} row(s))", expanded=False):
        if help_text:
            st.info(help_text)
        if df.empty or (len(df.columns) == 1 and "message" in df.columns):
            val = df.iloc[0, 0] if not df.empty else "No data."
            st.caption(str(val))
        else:
            st.dataframe(df, use_container_width=True, height=300)


def _render_sample_explanation() -> None:
    st.markdown(
        """
        ---
        ### How the analysis works

        **Step 1 — Tab Classification**
        Every worksheet is scored against seven structural signals:
        formula density, inter-sheet reference count, aggregation function presence,
        row count, average unique-value ratio, dtype homogeneity, and column count.
        The tab type with the highest weighted score wins.
        No sheet names are inspected.

        **Step 2 — KPI Extraction**
        Tabs classified as KPI/Summary, Dashboard, Output, or Calculation are scanned
        for formula cells. Each formula is parsed for its aggregate function and
        cross-sheet column references.

        **Step 3 — Schema Matching**
        Columns from data-bearing tabs are matched pairwise across workbooks.
        Only exact normalized-name matches are marked `safe_to_merge`.
        Fuzzy matches (score 80–94) are documented for review but never rename columns.

        **Step 4 — Collision Detection**
        Many-to-one mappings (two distinct source columns matching the same canonical target)
        are detected and flagged HIGH severity. Each is listed in sheet 07 with remediation guidance.

        **Step 5 — KPI Rationalization**
        Two KPIs are only candidates for consolidation if they share the same label,
        the same aggregate function, **and** at least one overlapping referenced source column.
        """
    )
