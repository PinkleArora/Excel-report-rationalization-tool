"""Streamlit page: Upload & Profile Reports (Phase 1)."""

from __future__ import annotations

import io
import logging
import tempfile
from pathlib import Path

import streamlit as st

logger = logging.getLogger(__name__)


def render() -> None:
    st.header("Upload & Profile Reports")
    st.caption("Upload one or more Excel (.xlsx, .xlsm, .xls) or CSV files to generate a profiling report.")

    uploaded = st.file_uploader(
        "Choose files",
        type=["xlsx", "xlsm", "xls", "csv"],
        accept_multiple_files=True,
    )

    if not uploaded:
        st.info("Upload at least one file to begin.")
        return

    if st.button("Run Profiling", type="primary"):
        _run_profiling(uploaded)


def _run_profiling(uploaded_files) -> None:
    from src.ingestion.loader import load_workbook
    from src.profiling.exporter import export_profiling_results
    from src.profiling.profiler import profile_many, profile_workbook

    bundles = []
    progress = st.progress(0, text="Loading files…")

    with tempfile.TemporaryDirectory() as tmp:
        for i, uf in enumerate(uploaded_files):
            tmp_path = Path(tmp) / uf.name
            tmp_path.write_bytes(uf.read())
            try:
                bundle = load_workbook(tmp_path)
                bundles.append(bundle)
            except Exception as exc:
                st.warning(f"Could not load **{uf.name}**: {exc}")
            progress.progress((i + 1) / len(uploaded_files), text=f"Loaded {uf.name}")

    if not bundles:
        st.error("No files were loaded successfully.")
        return

    with st.spinner("Profiling workbooks…"):
        metas = profile_many(bundles)

    st.success(f"Profiled **{len(metas)}** workbook(s).")

    from src.profiling.profiler import build_summary_dataframes
    wb_df, sheet_df, col_df = build_summary_dataframes(metas)

    tab_wb, tab_sh, tab_col = st.tabs(["Workbook Summary", "Sheet Summary", "Column Summary"])
    with tab_wb:
        st.dataframe(wb_df, use_container_width=True)
    with tab_sh:
        st.dataframe(sheet_df, use_container_width=True)
    with tab_col:
        st.dataframe(col_df, use_container_width=True)

    # In-memory export for download
    buf = io.BytesIO()
    import pandas as pd
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        wb_df.to_excel(writer, sheet_name="Workbook_Summary", index=False)
        sheet_df.to_excel(writer, sheet_name="Sheet_Summary", index=False)
        col_df.to_excel(writer, sheet_name="Column_Summary", index=False)
    buf.seek(0)

    st.download_button(
        label="⬇ Download profiling_summary.xlsx",
        data=buf,
        file_name="profiling_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
