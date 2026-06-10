"""Streamlit page: Generate Master Workbook (Phase 3)."""

from __future__ import annotations

import io
import logging
import tempfile
from pathlib import Path

import streamlit as st

logger = logging.getLogger(__name__)


def render() -> None:
    st.header("Generate Master Workbook")
    st.caption(
        "Upload source reports to consolidate them into a single BAU master workbook "
        "(**master_workbook.xlsx**) containing data, KPIs, source mapping, data dictionary, "
        "reconciliation checks, issues log, rationalization report, and update SOP."
    )

    with st.expander("ℹ️ What this generates", expanded=False):
        st.markdown(
            """
            | Sheet | Contents |
            |-------|---------|
            | **01_Master_Data** | All source rows consolidated, with `_source_workbook` / `_source_sheet` lineage columns |
            | **02_KPI_Summary** | Numeric column aggregates (count, sum, mean, min, max) per source |
            | **03_Source_Mapping** | Canonical column name mapping for every source column |
            | **04_Data_Dictionary** | Auto-generated column definitions, types, and sample values |
            | **05_Reconciliation** | Numeric totals cross-check: source vs master (PASS / WARN / FAIL) |
            | **06_Issues_Log** | Auto-detected data quality issues (high nulls, formula cells, recon failures) |
            | **07_Rationalization_Report** | Keep / merge / retire recommendation per source workbook |
            | **08_Update_SOP** | Standard Operating Procedure for monthly master workbook maintenance |
            """
        )

    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded = st.file_uploader(
            "Upload source reports",
            type=["xlsx", "xlsm", "xls", "csv"],
            accept_multiple_files=True,
            key="master_uploader",
        )
    with col2:
        threshold = st.slider(
            "Schema match threshold",
            min_value=50, max_value=100, value=80, step=5,
            help="Minimum fuzzy-match score (0–100) to link columns across workbooks.",
        )
        tolerance = st.slider(
            "Reconciliation tolerance (%)",
            min_value=0, max_value=10, value=1, step=1,
            help="Acceptable % difference between source and master totals.",
        ) / 100.0

    if not uploaded:
        st.info("Upload at least one file to begin.")
        return

    if st.button("Generate Master Workbook", type="primary", use_container_width=True):
        _run_build(uploaded, threshold=float(threshold), tolerance=tolerance)


def _run_build(uploaded_files, threshold: float, tolerance: float) -> None:
    from src.ingestion.loader import load_workbook
    from src.workbook_generator.master_builder import (
        MasterWorkbookContext,
        _stage_data_dictionary,
        _stage_profile,
        _stage_rationalization,
        _stage_reconciliation,
        _stage_schema_matching,
        _stage_source_mapping,
        build_master_workbook_bytes,
    )
    from src.workbook_generator.master_builder import _stage_consolidate as _stage_cons

    bundles = []
    progress = st.progress(0, text="Loading files…")

    with tempfile.TemporaryDirectory() as tmp:
        for i, uf in enumerate(uploaded_files):
            tmp_path = Path(tmp) / uf.name
            tmp_path.write_bytes(uf.read())
            try:
                bundles.append(load_workbook(tmp_path))
            except Exception as exc:
                st.warning(f"Could not load **{uf.name}**: {exc}")
            progress.progress((i + 1) / len(uploaded_files), text=f"Loaded {uf.name}")

    if not bundles:
        st.error("No files could be loaded.")
        return

    status_box = st.empty()
    stage_progress = st.progress(0, text="Starting pipeline…")

    def update(msg: str, pct: int) -> None:
        status_box.info(msg)
        stage_progress.progress(pct, text=msg)

    try:
        update("Profiling workbooks…", 10)
        from src.workbook_generator.master_builder import MasterWorkbookContext, _stage_profile
        ctx = MasterWorkbookContext(bundles=bundles)
        _stage_profile(ctx)

        update("Matching column schemas…", 25)
        _stage_schema_matching(ctx, threshold)

        update("Consolidating master data…", 40)
        _stage_cons(ctx)

        update("Building source mapping…", 55)
        _stage_source_mapping(ctx)

        update("Building data dictionary…", 65)
        _stage_data_dictionary(ctx)

        update("Running reconciliation checks…", 75)
        _stage_reconciliation(ctx, tolerance)

        update("Assessing rationalization…", 85)
        _stage_rationalization(ctx)

        update("Generating Excel workbook…", 90)
        from src.workbook_generator.master_builder import build_master_workbook_bytes
        workbook_bytes = build_master_workbook_bytes(
            bundles,
            matching_threshold=threshold,
            reconciliation_tolerance=tolerance,
        )

        stage_progress.progress(100, text="Done.")
        status_box.success(
            f"Master workbook built from **{len(bundles)}** source file(s) — "
            f"{len(ctx.column_matches)} column match(es) found."
        )

    except Exception as exc:
        st.error(f"Build failed: {exc}")
        logger.exception("Master workbook build error")
        return

    _render_preview(ctx)

    st.download_button(
        label="⬇ Download master_workbook.xlsx",
        data=workbook_bytes,
        file_name="master_workbook.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )


def _render_preview(ctx) -> None:
    """Show tabbed previews of key sheets."""
    from src.consolidation.consolidator import LINEAGE_COL_SHEET, LINEAGE_COL_WORKBOOK
    from src.documentation.generator import data_dictionary_to_df
    from src.rationalization.rationalizer import build_inventory
    from src.reconciliation.reconciler import reconciliation_summary
    from src.workbook_generator.master_builder import _build_kpi_summary, _build_issues_log

    st.subheader("Preview")

    tabs = st.tabs([
        "01 Master Data",
        "02 KPI Summary",
        "03 Source Mapping",
        "04 Data Dictionary",
        "05 Reconciliation",
        "06 Issues Log",
        "07 Rationalization",
    ])

    def _show(tab, df, max_rows: int = 200) -> None:
        with tab:
            st.dataframe(df.head(max_rows), use_container_width=True)
            st.caption(f"{len(df):,} row(s)")

    _show(tabs[0], ctx.master_df)
    _show(tabs[1], _build_kpi_summary(ctx))
    _show(tabs[2], ctx.source_mapping_df)
    _show(tabs[3], data_dictionary_to_df(ctx.data_dict_entries))
    _show(tabs[4], reconciliation_summary(ctx.reconciliation_results))
    _show(tabs[5], _build_issues_log(ctx))
    _show(tabs[6], build_inventory(ctx.rationalization_recs))
