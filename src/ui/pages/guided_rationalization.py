"""Streamlit page: Guided Workbook Rationalization.

Multi-step wizard:
  Step 1 — Upload workbooks
  Step 2 — Configure source tab + KPI tabs + LOB per workbook
  Step 3 — Naming & options
  Step 4 — Review analysis summary
  Step 5 — Generate & download future-state workbook
"""

from __future__ import annotations

import logging

import streamlit as st

from src.ingestion.loader import load_workbook_from_bytes
from src.guided_rationalization.config import RationalizationConfig, WorkbookTabConfig
from src.guided_rationalization.source_analyzer import analyze_source_data
from src.guided_rationalization.kpi_analyzer import analyze_kpi_dependencies
from src.guided_rationalization.workbook_builder import (
    build_rationalized_workbook_pair,
    DuplicateColumnError,
)
from src.guided_rationalization.lineage_builder import (
    build_data_lineage_df,
    build_column_lineage_df,
    build_kpi_dependencies_df,
    build_lineage_workbook,
)
from src.guided_rationalization.diagram_builder import (
    build_rationalization_flow,
    build_dependency_diagram,
)

logger = logging.getLogger(__name__)

_STEPS = ["Upload", "Configure Tabs", "Naming & Options", "Review Analysis", "Generate"]


def _step_indicator(current: int) -> None:
    cols = st.columns(len(_STEPS))
    for i, (col, label) in enumerate(zip(cols, _STEPS)):
        if i < current:
            col.markdown(f"✅ **{label}**")
        elif i == current:
            col.markdown(f"🔵 **{label}**")
        else:
            col.markdown(f"⬜ {label}")
    st.divider()


def render() -> None:
    st.title("Guided Workbook Rationalization")
    st.caption(
        "A user-driven wizard that consolidates source data and updates KPI tab "
        "dependencies into a single future-state workbook."
    )

    # ── session state bootstrap ──────────────────────────────────────────────
    if "gr_step" not in st.session_state:
        st.session_state.gr_step = 0
    if "gr_bundles" not in st.session_state:
        st.session_state.gr_bundles = []
    if "gr_wb_configs" not in st.session_state:
        st.session_state.gr_wb_configs = {}   # file_name → dict of config values
    if "gr_options" not in st.session_state:
        st.session_state.gr_options = {}

    step = st.session_state.gr_step
    _step_indicator(step)

    # ── Step 0: Upload ────────────────────────────────────────────────────────
    if step == 0:
        _step_upload()

    # ── Step 1: Configure per-workbook tabs ──────────────────────────────────
    elif step == 1:
        _step_configure_tabs()

    # ── Step 2: Naming & options ─────────────────────────────────────────────
    elif step == 2:
        _step_naming_options()

    # ── Step 3: Review analysis ──────────────────────────────────────────────
    elif step == 3:
        _step_review_analysis()

    # ── Step 4: Generate ─────────────────────────────────────────────────────
    elif step == 4:
        _step_generate()


# ---------------------------------------------------------------------------
# Step implementations
# ---------------------------------------------------------------------------

def _step_upload() -> None:
    st.subheader("Step 1 — Upload Workbooks")
    uploaded = st.file_uploader(
        "Upload one or more Excel workbooks (.xlsx / .xlsm)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        key="gr_uploader",
    )

    if uploaded:
        bundles = []
        errors = []
        for f in uploaded:
            try:
                bundle = load_workbook_from_bytes(f.read(), file_name=f.name)
                bundles.append(bundle)
            except Exception as exc:
                errors.append(f"{f.name}: {exc}")

        if errors:
            for e in errors:
                st.error(e)

        if bundles:
            st.success(f"Loaded {len(bundles)} workbook(s).")
            for b in bundles:
                with st.expander(b.file_name):
                    st.write(f"Sheets: {', '.join(b.sheets.keys())}")

            if st.button("Next →", type="primary"):
                st.session_state.gr_bundles = bundles
                # pre-populate wb_configs skeleton
                for b in bundles:
                    if b.file_name not in st.session_state.gr_wb_configs:
                        st.session_state.gr_wb_configs[b.file_name] = {
                            "source_tab": "",
                            "kpi_tabs": [],
                        }
                st.session_state.gr_step = 1
                st.rerun()
    else:
        st.info("Upload at least one workbook to begin.")


def _step_configure_tabs() -> None:
    st.subheader("Step 2 — Configure Tabs")
    st.caption(
        "For each workbook, select the **source data tab** (raw transactional records) "
        "and one or more **KPI / summary tabs** (formula-driven outputs)."
    )

    bundles = st.session_state.gr_bundles
    wb_configs = st.session_state.gr_wb_configs

    all_valid = True
    for bundle in bundles:
        sheet_names = list(bundle.sheets.keys())
        cfg = wb_configs[bundle.file_name]

        with st.expander(f"📂 {bundle.file_name}", expanded=True):
            source_tab = st.selectbox(
                "Source data tab",
                options=["(select)"] + sheet_names,
                index=(["(select)"] + sheet_names).index(cfg["source_tab"])
                if cfg["source_tab"] in sheet_names else 0,
                key=f"gr_src_{bundle.file_name}",
            )
            cfg["source_tab"] = source_tab if source_tab != "(select)" else ""

            remaining = [s for s in sheet_names if s != cfg["source_tab"]]
            kpi_tabs = st.multiselect(
                "KPI / summary tab(s)",
                options=remaining,
                default=[t for t in cfg["kpi_tabs"] if t in remaining],
                key=f"gr_kpi_{bundle.file_name}",
            )
            cfg["kpi_tabs"] = kpi_tabs

        if not cfg["source_tab"]:
            all_valid = False

    col1, col2 = st.columns([1, 5])
    with col1:
        if st.button("← Back"):
            st.session_state.gr_step = 0
            st.rerun()
    with col2:
        if not all_valid:
            st.warning("Select a source tab for every workbook before continuing.")
        if st.button("Next →", type="primary", disabled=not all_valid):
            st.session_state.gr_wb_configs = wb_configs
            st.session_state.gr_step = 2
            st.rerun()


def _step_naming_options() -> None:
    st.subheader("Step 3 — Naming & Options")

    opts = st.session_state.gr_options
    bundles = st.session_state.gr_bundles
    wb_configs = st.session_state.gr_wb_configs

    opts["future_source_tab_name"] = st.text_input(
        "Future master source tab name",
        value=opts.get("future_source_tab_name", "Master_Source_Data"),
    )

    st.markdown("**Future KPI tab names** (leave blank for auto-generated names):")
    future_kpi_names: dict[str, str] = opts.get("future_kpi_tab_names", {})
    for bundle in bundles:
        cfg = wb_configs[bundle.file_name]
        for kpi_tab in cfg.get("kpi_tabs", []):
            key = f"{bundle.file_name}::{kpi_tab}"
            default_name = future_kpi_names.get(key, "")
            future_kpi_names[key] = st.text_input(
                f"{bundle.file_name} / {kpi_tab}",
                value=default_name,
                key=f"gr_kpiname_{key}",
            )
    opts["future_kpi_tab_names"] = future_kpi_names

    st.divider()
    opts["matching_threshold"] = st.slider(
        "Column name matching threshold (fuzzy similarity %)",
        min_value=50, max_value=100,
        value=int(opts.get("matching_threshold", 80)),
        help="Columns with similarity ≥ threshold are flagged as SIMILAR (documented, not auto-merged).",
    )
    opts["merge_high_confidence"] = st.checkbox(
        "Auto-merge high-confidence fuzzy matches (score ≥ 95%)",
        value=opts.get("merge_high_confidence", False),
    )
    # Always enforce KPI-only columns — removed columns are reported in the Analysis Pack.
    opts["remove_unused_columns"] = True

    col1, col2 = st.columns([1, 5])
    with col1:
        if st.button("← Back"):
            st.session_state.gr_step = 1
            st.rerun()
    with col2:
        if st.button("Next →", type="primary"):
            st.session_state.gr_options = opts
            st.session_state.gr_step = 3
            st.rerun()


def _build_config() -> RationalizationConfig:
    opts = st.session_state.gr_options
    wb_configs = st.session_state.gr_wb_configs

    # Flatten future_kpi_tab_names: key is "wb::tab", value maps workbook → name
    raw_kpi_names = opts.get("future_kpi_tab_names", {})
    future_kpi_tab_names: dict[str, str] = {}
    for key, name in raw_kpi_names.items():
        if name.strip():
            wb_name, _ = key.split("::", 1)
            future_kpi_tab_names[wb_name] = name.strip()

    workbook_configs = [
        WorkbookTabConfig(
            workbook_name=wb_name,
            source_tab=cfg["source_tab"],
            kpi_tabs=cfg["kpi_tabs"],
        )
        for wb_name, cfg in wb_configs.items()
        if cfg["source_tab"]
    ]

    return RationalizationConfig(
        workbook_configs=workbook_configs,
        future_source_tab_name=opts.get("future_source_tab_name", "Master_Source_Data"),
        future_kpi_tab_names=future_kpi_tab_names,
        matching_threshold=float(opts.get("matching_threshold", 80)),
        merge_high_confidence=bool(opts.get("merge_high_confidence", False)),
        remove_unused_columns=bool(opts.get("remove_unused_columns", True)),
    )


def _step_review_analysis() -> None:
    st.subheader("Step 4 — Review Analysis")

    bundles = st.session_state.gr_bundles
    config = _build_config()

    with st.spinner("Running source and KPI analysis…"):
        # First pass: source analysis without KPI refs
        source_result_pass1 = analyze_source_data(bundles, config)
        # KPI analysis
        kpi_result = analyze_kpi_dependencies(
            bundles, config, source_result_pass1.column_mapping
        )
        # Second pass: source analysis with KPI refs for exclusions
        source_result = analyze_source_data(
            bundles, config, kpi_referenced_canonicals=kpi_result.referenced_canonicals
        )

    st.session_state.gr_source_result = source_result
    st.session_state.gr_kpi_result = kpi_result

    # ── Metrics ──────────────────────────────────────────────────────────────
    # Count source column *entries* (one per source workbook × column), not
    # de-duplicated canonicals, so the numbers reconcile to a total.
    excl_set = set(source_result.excluded_columns)
    common_kpi = sum(
        1 for (wb, tab, col), canonical in source_result.column_mapping.items()
        if (wb, tab, col) not in excl_set
        and canonical in kpi_result.referenced_canonicals
        and next((p for p in source_result.column_profiles if p.canonical_name == canonical), None) is not None
        and next((p for p in source_result.column_profiles if p.canonical_name == canonical)).match_class == "common"
    )
    unique_kpi = sum(
        1 for (wb, tab, col), canonical in source_result.column_mapping.items()
        if (wb, tab, col) not in excl_set
        and canonical in kpi_result.referenced_canonicals
        and (
            next((p for p in source_result.column_profiles if p.canonical_name == canonical), None) is None
            or next((p for p in source_result.column_profiles if p.canonical_name == canonical)).match_class != "common"
        )
    )
    excluded_count = len(source_result.excluded_columns)
    total_source = len(source_result.column_mapping)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Source Columns", total_source)
    c2.metric("Referenced by KPIs", common_kpi + unique_kpi)
    c3.metric("Common — KPI Required", common_kpi,
              help="Columns present in >1 source workbook that are needed by KPI formulas")
    c4.metric("Unique — KPI Required", unique_kpi,
              help="Columns present in only 1 source workbook that are needed by KPI formulas")
    c5.metric("Excluded (Not Required)", excluded_count,
              help="Columns not referenced by any KPI — omitted from Master Source Data")

    # Reconciliation check
    reconciles = (common_kpi + unique_kpi + excluded_count) == total_source
    if reconciles:
        st.success(
            f"✓ Metrics reconcile: {common_kpi} common + {unique_kpi} unique "
            f"+ {excluded_count} excluded = {total_source} total"
        )
    else:
        st.warning(
            f"⚠ Reconciliation gap: {common_kpi} + {unique_kpi} + {excluded_count} "
            f"= {common_kpi + unique_kpi + excluded_count} ≠ {total_source} total  "
            "(some columns may map to canonicals not in column_profiles)"
        )

    # ── Source column profiles ────────────────────────────────────────────────
    with st.expander("Source Column Profiles", expanded=True):
        df = source_result.to_dataframe()
        if not df.empty:
            st.dataframe(df, use_container_width=True)
        else:
            st.info("No source columns found.")

    # ── KPI dependencies ─────────────────────────────────────────────────────
    with st.expander("KPI Formula Dependencies"):
        dep_df = kpi_result.to_dependency_dataframe()
        if not dep_df.empty:
            st.dataframe(dep_df, use_container_width=True)
        else:
            st.info("No KPI formulas found. Ensure uploaded files are .xlsx with formulas (not data_only).")

    # ── Column coverage ───────────────────────────────────────────────────────
    with st.expander("Column Coverage (KPI usage)"):
        all_canonicals: set[str] = {c for c in source_result.column_mapping.values()}
        cov_df = kpi_result.to_coverage_dataframe(all_canonicals)
        if not cov_df.empty:
            st.dataframe(cov_df, use_container_width=True)

    # ── Data Lineage preview ──────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📐 Data Lineage")
    st.caption(
        "Full traceability from source columns to KPI outputs.  "
        "These reports are also included in the Rationalization Analysis Pack."
    )

    lineage_tab1, lineage_tab2, lineage_tab3 = st.tabs(
        ["KPI → Source (Data Lineage)", "Column Provenance (Column Lineage)", "Dependency Map"]
    )

    with lineage_tab1:
        lin_df = build_data_lineage_df(source_result, kpi_result, config)
        if lin_df.empty:
            st.info("No KPI dependencies found.")
        else:
            st.caption(f"{len(lin_df)} KPI formula(s) traced to source columns.")
            st.dataframe(lin_df, use_container_width=True, hide_index=True)

    with lineage_tab2:
        col_lin_df = build_column_lineage_df(source_result, kpi_result, config)
        if col_lin_df.empty:
            st.info("No column lineage data available.")
        else:
            st.caption(f"{len(col_lin_df)} column source entries.")
            st.dataframe(col_lin_df, use_container_width=True, hide_index=True)

    with lineage_tab3:
        dep_df = build_kpi_dependencies_df(source_result, kpi_result, config)
        if dep_df.empty:
            st.info("No KPI dependencies found.")
        else:
            st.caption(f"{len(dep_df)} KPI dependency rows.")
            st.dataframe(dep_df, use_container_width=True, hide_index=True)

    with st.spinner("Building lineage workbook…"):
        lineage_bytes = build_lineage_workbook(source_result, kpi_result, config)

    st.download_button(
        label="⬇ Download Data_Lineage.xlsx (standalone)",
        data=lineage_bytes,
        file_name="Data_Lineage.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_lineage_preview",
    )

    # ── Visual Diagrams ───────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("🗺️ Visual Rationalization Diagrams")
    st.caption(
        "Business-friendly visuals embedded in the Analysis Pack.  "
        "Download PNG or SVG for presentations and documentation."
    )

    diag_flow_png = diag_flow_svg = diag_dep_png = diag_dep_svg = None
    try:
        with st.spinner("Generating diagrams…"):
            diag_flow_png, diag_flow_svg = build_rationalization_flow(
                source_result, kpi_result, config
            )
            diag_dep_png, diag_dep_svg = build_dependency_diagram(
                source_result, kpi_result, config
            )
    except Exception as _diag_err:
        st.warning(f"Diagram generation encountered an issue: {_diag_err}")

    if diag_flow_png:
        diag_t1, diag_t2 = st.tabs(
            ["🔵 Rationalization Flow", "🔗 Column Dependency"]
        )
        with diag_t1:
            st.image(diag_flow_png, use_container_width=True)
            dcol1, dcol2 = st.columns(2)
            dcol1.download_button(
                "⬇ Flow Diagram — PNG",
                data=diag_flow_png,
                file_name="Rationalization_Flow.png",
                mime="image/png",
                key="dl_flow_png",
            )
            dcol2.download_button(
                "⬇ Flow Diagram — SVG",
                data=diag_flow_svg,
                file_name="Rationalization_Flow.svg",
                mime="image/svg+xml",
                key="dl_flow_svg",
            )
        with diag_t2:
            st.image(diag_dep_png, use_container_width=True)
            dcol3, dcol4 = st.columns(2)
            dcol3.download_button(
                "⬇ Dependency Diagram — PNG",
                data=diag_dep_png,
                file_name="Dependency_Diagram.png",
                mime="image/png",
                key="dl_dep_png",
            )
            dcol4.download_button(
                "⬇ Dependency Diagram — SVG",
                data=diag_dep_svg,
                file_name="Dependency_Diagram.svg",
                mime="image/svg+xml",
                key="dl_dep_svg",
            )
        st.caption(
            "Both diagrams are also embedded in the Rationalization Analysis Pack "
            "(Rationalization_Flow and Dependency_Diagram sheets)."
        )

    # Store for re-use in Step 5
    st.session_state["gr_flow_png"]  = diag_flow_png
    st.session_state["gr_flow_svg"]  = diag_flow_svg
    st.session_state["gr_dep_png"]   = diag_dep_png
    st.session_state["gr_dep_svg"]   = diag_dep_svg

    col1, col2 = st.columns([1, 5])
    with col1:
        if st.button("← Back"):
            st.session_state.gr_step = 2
            st.rerun()
    with col2:
        if st.button("Generate Future-State Workbook →", type="primary"):
            st.session_state.gr_step = 4
            st.rerun()


def _step_generate() -> None:
    st.subheader("Step 5 — Generate Workbooks")

    bundles = st.session_state.gr_bundles
    config = _build_config()
    source_result = st.session_state.get("gr_source_result")
    kpi_result = st.session_state.get("gr_kpi_result")

    if source_result is None or kpi_result is None:
        st.error("Analysis results missing — please go back to Step 4.")
        if st.button("← Back to Analysis"):
            st.session_state.gr_step = 3
            st.rerun()
        return

    fs_bytes: bytes | None = None
    ap_bytes: bytes | None = None
    has_blocking = False

    with st.spinner("Building workbooks…"):
        try:
            fs_bytes, ap_bytes = build_rationalized_workbook_pair(
                bundles, config, source_result, kpi_result
            )
        except DuplicateColumnError as dup_exc:
            has_blocking = True
            ap_bytes = dup_exc.diagnostic_bytes
            blocking = [r for r in dup_exc.resolutions if r.scenario == "C"]
            st.error(
                "⚠ Manual review required — duplicate columns cannot be auto-resolved "
                f"({len(blocking)} position(s) each used by different KPI formulas)."
            )
            st.markdown(
                "All other duplicate columns were **automatically resolved** "
                "(Scenarios A, B, D).  \n"
                "The following positions could **not** be resolved automatically because "
                "each is referenced by a different KPI formula, and the tool cannot "
                "determine which one to retain without business input.\n"
            )
            for r in blocking:
                st.markdown(
                    f"- **{r.workbook} / {r.source_tab}**: "
                    f"canonical `{r.canonical_name}` — "
                    f"position {r.original_position} (`{r.original_col_name}`) "
                    f"used by KPIs: `{', '.join(r.kpis_using)}`"
                )
            st.info(
                "**To resolve:** rename the conflicting source columns in your workbook "
                "to give them distinct names (e.g. 'Reserve_Amount_Gross' and "
                "'Reserve_Amount_Net'), then re-run.  \n"
                "Download the Analysis Pack below for full detail."
            )
        except Exception as exc:
            st.error(f"Generation failed: {exc}")
            logger.exception("Workbook generation error")
            return

    if has_blocking:
        if ap_bytes:
            st.download_button(
                label="⬇ Download Diagnostic Analysis Pack",
                data=ap_bytes,
                file_name="Rationalization_Analysis_Pack_diagnostic.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.success("Workbooks generated successfully.")

        # Build standalone lineage workbook
        source_result_gen = st.session_state.get("gr_source_result")
        kpi_result_gen = st.session_state.get("gr_kpi_result")
        lineage_bytes_gen: bytes | None = None
        if source_result_gen and kpi_result_gen:
            try:
                lineage_bytes_gen = build_lineage_workbook(source_result_gen, kpi_result_gen, config)
            except Exception as _exc:
                logger.warning("Could not build lineage workbook: %s", _exc)

        col_dl1, col_dl2, col_dl3 = st.columns(3)
        with col_dl1:
            st.markdown("### 📊 Future State Workbook")
            st.caption(
                "Production-ready workbook for BAU use.  \n"
                "Contains: Master Source Data + recreated KPI Summary tabs."
            )
            st.download_button(
                label="⬇ Download Future_State_Workbook.xlsx",
                data=fs_bytes,
                file_name="Future_State_Workbook.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="dl_fs",
            )
            st.markdown("""
**Contents:**
- Master Source Data tab — KPI-referenced columns only, with lineage
- One recreated Summary tab per configured KPI tab
""")

        with col_dl2:
            st.markdown("### 🔍 Rationalization Analysis Pack")
            st.caption(
                "Diagnostic and project documentation workbook.  \n"
                "Contains: mappings, audit, reconciliation, lineage, documentation."
            )
            st.download_button(
                label="⬇ Download Rationalization_Analysis_Pack.xlsx",
                data=ap_bytes,
                file_name="Rationalization_Analysis_Pack.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_ap",
            )
            st.markdown("""
**Contents:**
- `Source_Mapping` — source column → canonical mapping
- `Reconciliation` — row-count verification
- `Issues_Log` — auto-resolved decisions and warnings
- `Documentation` — run configuration and BAU update instructions
- `Formula_Validation` — generated formula audit
- `KPI_Audit_...` — per-KPI formula details
- `Removed_Source_Columns` — columns excluded from Master Source Data
- `Duplicate_Column_Analysis` — KPI usage and scenario (when applicable)
- `Data_Lineage` — KPI → source column traceability
- `Column_Lineage` — master column provenance
- `KPI_Dependencies` — compact dependency map
""")

        with col_dl3:
            st.markdown("### 📐 Data Lineage")
            st.caption(
                "Standalone lineage workbook for sharing with business stakeholders.  \n"
                "Contains: Data Lineage, Column Lineage, KPI Dependencies, Legend."
            )
            if lineage_bytes_gen:
                st.download_button(
                    label="⬇ Download Data_Lineage.xlsx",
                    data=lineage_bytes_gen,
                    file_name="Data_Lineage.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_lineage_gen",
                )
            else:
                st.info("Lineage workbook unavailable.")
            st.markdown("""
**Contents:**
- `Data_Lineage` — KPI formula → source columns
- `Column_Lineage` — original column → master column
- `KPI_Dependencies` — compact dependency map
- `Legend` — confidence colour key and status definitions
""")

    # ── Diagram exports ───────────────────────────────────────────────────────
    flow_png = st.session_state.get("gr_flow_png")
    flow_svg = st.session_state.get("gr_flow_svg")
    dep_png  = st.session_state.get("gr_dep_png")
    dep_svg  = st.session_state.get("gr_dep_svg")

    if flow_png or dep_png:
        with st.expander("📊 Export Visual Diagrams", expanded=False):
            st.caption("Download PNG or SVG diagrams for presentations and documentation.")
            exp_cols = st.columns(4)
            if flow_png:
                exp_cols[0].download_button("⬇ Flow PNG",  flow_png, "Rationalization_Flow.png",  "image/png", key="dl_s5_flow_png")
                exp_cols[1].download_button("⬇ Flow SVG",  flow_svg, "Rationalization_Flow.svg",  "image/svg+xml", key="dl_s5_flow_svg")
            if dep_png:
                exp_cols[2].download_button("⬇ Dependency PNG", dep_png, "Dependency_Diagram.png", "image/png", key="dl_s5_dep_png")
                exp_cols[3].download_button("⬇ Dependency SVG", dep_svg, "Dependency_Diagram.svg", "image/svg+xml", key="dl_s5_dep_svg")

    col1, col2 = st.columns([1, 5])
    with col1:
        if st.button("← Back"):
            st.session_state.gr_step = 3
            st.rerun()
    with col2:
        if st.button("🔄 Start New Rationalization"):
            for key in list(st.session_state.keys()):
                if key.startswith("gr_"):
                    del st.session_state[key]
            st.rerun()
