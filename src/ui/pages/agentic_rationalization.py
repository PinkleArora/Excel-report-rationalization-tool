"""Streamlit page — Agentic Rationalization step-by-step consulting workflow.

The pipeline runs in discrete user-confirmed phases (steps 1-8).
Session state tracks the current step and intermediate results so each phase
stays visible as a collapsed summary once confirmed.
"""

from __future__ import annotations

import io
import traceback
import zipfile
from datetime import datetime
from typing import Any

import openpyxl
import pandas as pd
import streamlit as st

from src.ingestion.loader import load_workbook_from_bytes
from src.ingestion.workbook_analyzer import scan_kpi_blocks, _get_raw_ws, build_kpi_blocks_from_pivot
from src.agentic_rationalization.orchestrator import run_pipeline
from src.agentic_rationalization.pipeline_phases import (
    run_discovery_phase,
    run_consolidation_phase,
    run_schema_phase,
    run_kpi_phase,
    run_validation_phase,
    run_generation_phase,
)
from src.agentic_rationalization.models import AgentResult, PipelineResult
from src.agentic_rationalization.review_panel import (
    ActionableDecision,
    classify_actionable_decisions,
    SIMILAR_COLUMN_ACTIONS,
    SIMILAR_COLUMN_ACTIONS_VETOED,
    KPI_UNRESOLVED_ACTIONS,
    CONSOLIDATION_ACTIONS,
)

# ── Session state keys ────────────────────────────────────────────────────────

_SS_STEP           = "ar_step"
_SS_BUNDLES        = "ar_bundles"
_SS_CONFIG         = "ar_config"
_SS_DISCOVERY      = "ar_discovery"
_SS_WB_ANALYSES    = "ar_wb_analyses"   # cached WorkbookAnalysis list — survives override changes
_SS_CONSOLIDATION  = "ar_consolidation"
_SS_SELECTED_GROUP = "ar_selected_group"
_SS_SCHEMA         = "ar_schema"
_SS_KPI            = "ar_kpi"
_SS_VALIDATION     = "ar_validation"
_SS_GENERATION     = "ar_generation"
_SS_SOURCE_TAB_OVERRIDES  = "ar_src_overrides"
_SS_GROUP_OVERRIDES       = "ar_group_overrides"   # dict[str, int|None] file→group_id (None=excluded)
_SS_KPI_ALIGN_DECISIONS   = "ar_kpi_align"         # dict[canonical, "accepted"|"rejected"|pending]

# UI render caches — survive re-renders, cleared on new upload
_SS_STRUCT_CACHE  = "ar_struct_cache"   # (wb_name, sheet) → structure label string
_SS_BLOCKS_CACHE  = "ar_blocks_cache"   # (wb_name, sheet) → list[KpiBlock]

# Legacy keys kept for the override review panel
_OVERRIDES_KEY = "agentic_decision_overrides"
_ACTION_KEY    = "agentic_action_overrides"


# ── Step helpers ──────────────────────────────────────────────────────────────

def _current_step() -> int:
    return st.session_state.get(_SS_STEP, 1)


def _set_step(n: int) -> None:
    st.session_state[_SS_STEP] = n


def _reset_state() -> None:
    for key in [
        _SS_STEP, _SS_BUNDLES, _SS_CONFIG, _SS_DISCOVERY, _SS_WB_ANALYSES,
        _SS_CONSOLIDATION, _SS_SELECTED_GROUP, _SS_SCHEMA, _SS_KPI,
        _SS_VALIDATION, _SS_GENERATION,
        _SS_SOURCE_TAB_OVERRIDES, _SS_GROUP_OVERRIDES, _SS_KPI_ALIGN_DECISIONS,
        _SS_STRUCT_CACHE, _SS_BLOCKS_CACHE,
        _OVERRIDES_KEY, _ACTION_KEY, "ar_kpi_tab_overrides", "ar_detection_type_overrides",
        "ar_disc_tags", "ar_disc_explorer_tab", "ar_disc_selected_block",
    ]:
        st.session_state.pop(key, None)


# ── Shared visual helpers ─────────────────────────────────────────────────────

def _conf_color(conf: float) -> str:
    if conf >= 0.90:
        return "🟢"
    if conf >= 0.70:
        return "🟡"
    return "🔴"


def _conf_label(conf: float) -> str:
    if conf >= 0.90:
        return "High"
    if conf >= 0.70:
        return "Medium"
    return "Low"


def _agent_badge(result: AgentResult) -> str:
    n_warn = len(result.warnings)
    n_low  = len(result.low_confidence_decisions)
    if n_warn or n_low:
        return f"⚠️ {result.agent_name}"
    return f"✅ {result.agent_name}"


def _render_agent_card(result: AgentResult) -> None:
    badge = _agent_badge(result)
    with st.expander(badge, expanded=False):
        if result.warnings:
            for w in result.warnings:
                st.warning(w)
        if not result.decisions:
            st.info("No decisions recorded for this agent.")
            return
        rows = []
        for d in result.decisions:
            rows.append({
                "": _conf_color(d.confidence),
                "Subject": d.subject,
                "Decision": d.decision,
                "Confidence": f"{d.confidence:.0%}",
                "Level": _conf_label(d.confidence),
                "Overridable": "Yes" if d.overridable else "No",
                "Reasoning": d.reasoning,
            })
        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, hide_index=True)


# ── Sidebar step navigation ───────────────────────────────────────────────────

_STEP_LABELS = [
    "1. Upload",
    "2. Discovery",
    "3. Consolidation",
    "4. Group Select",
    "5. Schema",
    "6. KPI",
    "7. Validation",
    "8. Generate",
]


def _render_sidebar_nav() -> None:
    """Render the persistent left-side workflow navigation panel."""
    step = _current_step()
    with st.sidebar:
        st.divider()
        st.markdown("**Agentic Rationalization**")
        for i, label in enumerate(_STEP_LABELS, start=1):
            if i < step:
                # Completed — clickable to navigate back
                if st.button(
                    f"✅  {label}",
                    key=f"ar_nav_{i}",
                    use_container_width=True,
                    help="Navigate back to this step",
                ):
                    _set_step(i)
                    st.rerun()
            elif i == step:
                # Current step — highlighted
                st.markdown(
                    f"<div style='"
                    f"background:#1e40af;color:white;"
                    f"padding:0.35rem 0.75rem;"
                    f"border-radius:0.375rem;"
                    f"font-size:0.85rem;font-weight:600;"
                    f"margin:0.15rem 0;"
                    f"'>▶  {label}</div>",
                    unsafe_allow_html=True,
                )
            else:
                # Future step — disabled
                st.button(
                    f"◦  {label}",
                    key=f"ar_nav_{i}",
                    use_container_width=True,
                    disabled=True,
                )


# ── Progress indicator (removed — replaced by sidebar nav) ───────────────────

def _render_progress_bar() -> None:
    pass  # kept for any external callers; sidebar nav is used instead


# ── Override panel helpers (kept from original) ───────────────────────────────

def _widget_key(prefix: str, subject: str) -> str:
    safe = subject.replace(" ", "_").replace("/", "_").replace("'", "")
    return f"{prefix}_{safe}"[:120]


def _render_similar_column_card(ad: ActionableDecision, idx: int) -> None:
    saved_actions: dict[str, str] = st.session_state.get(_ACTION_KEY, {})
    current_action = saved_actions.get(ad.subject, "ACCEPT")

    icon = _conf_color(ad.confidence)
    with st.container(border=True):
        c1, c2 = st.columns([3, 1])
        with c1:
            st.markdown(f"**{icon} Similar Column Match** — confidence `{ad.confidence:.0%}`")
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**Column A**")
                st.code(ad.canonical_a, language=None)
                if ad.original_names_a:
                    st.caption("Original name(s): " + " · ".join(ad.original_names_a[:3]))
                if ad.workbooks_a:
                    st.caption("Workbook: " + ", ".join(ad.workbooks_a[:2]))
            with col_b:
                st.markdown("**Column B**")
                st.code(ad.canonical_b or "—", language=None)
                if ad.original_names_b:
                    st.caption("Original name(s): " + " · ".join(ad.original_names_b[:3]))
                if ad.workbooks_b:
                    st.caption("Workbook: " + ", ".join(ad.workbooks_b[:2]))
            if ad.kpi_labels:
                st.caption(f"Affected KPIs: {', '.join(ad.kpi_labels[:4])}")
            if ad.description and "KPI veto" in ad.description:
                st.warning("KPI veto applied — these columns are used by different KPI formulas.")
        with c2:
            st.markdown("**Agent recommendation:**")
            st.info(ad.current_recommendation)

        key = _widget_key("sim", ad.subject)
        default_idx = next(
            (i for i, (k, _) in enumerate(ad.available_actions) if k == current_action), 0
        )
        labels = [label for _, label in ad.available_actions]
        chosen_label = st.radio("Your decision:", labels, index=default_idx, key=key, horizontal=False)
        chosen_key = next(k for k, l in ad.available_actions if l == chosen_label)

        if _ACTION_KEY not in st.session_state:
            st.session_state[_ACTION_KEY] = {}
        st.session_state[_ACTION_KEY][ad.subject] = chosen_key
        if chosen_key != "ACCEPT":
            if _OVERRIDES_KEY not in st.session_state:
                st.session_state[_OVERRIDES_KEY] = {}
            st.session_state[_OVERRIDES_KEY].setdefault(ad.agent_name, {})[ad.subject] = chosen_key


def _render_kpi_unresolved_card(ad: ActionableDecision, idx: int) -> None:
    saved_actions: dict[str, str] = st.session_state.get(_ACTION_KEY, {})
    current_action = saved_actions.get(ad.subject, "ACCEPT")

    with st.container(border=True):
        c1, c2 = st.columns([3, 1])
        with c1:
            st.markdown(f"**🔴 Unresolved KPI Formula** — `{ad.cell_address}` in `{ad.kpi_tab}`")
            st.caption(ad.description)
            if ad.formula:
                st.code(ad.formula, language=None)
            if ad.kpi_labels:
                st.caption(f"KPI label: **{ad.kpi_labels[0]}**")
        with c2:
            st.markdown("**Confidence:**")
            st.error(f"{ad.confidence:.0%}")

        key_radio = _widget_key("kpi", ad.subject)
        default_idx = next(
            (i for i, (k, _) in enumerate(ad.available_actions) if k == current_action), 0
        )
        labels = [label for _, label in ad.available_actions]
        chosen_label = st.radio("Your decision:", labels, index=default_idx, key=key_radio, horizontal=False)
        chosen_key = next(k for k, l in ad.available_actions if l == chosen_label)

        selected_col = ""
        if chosen_key == "MANUAL_COLUMN":
            saved_manual = (
                st.session_state.get(_OVERRIDES_KEY, {}).get(ad.agent_name, {}).get(ad.subject, "")
            )
            canonicals = ad.available_canonicals or []
            if canonicals:
                default_col = saved_manual if saved_manual in canonicals else (canonicals[0] if canonicals else "")
                default_col_idx = canonicals.index(default_col) if default_col in canonicals else 0
                st.markdown("**Select source column:**")
                st.caption(
                    "Choose the canonical source column this formula should map to. "
                    "Column names shown are the normalised names used in Master Source Data."
                )
                selected_col = st.selectbox(
                    "Source column:", options=canonicals, index=default_col_idx,
                    key=_widget_key("kpi_col", ad.subject),
                )
            else:
                selected_col = st.text_input(
                    "Source column name:",
                    value=saved_manual if saved_manual not in ("SKIP", "ACCEPT") else "",
                    key=_widget_key("kpi_manual", ad.subject),
                    placeholder="e.g. statutory_reserves_total",
                )

        if _ACTION_KEY not in st.session_state:
            st.session_state[_ACTION_KEY] = {}
        st.session_state[_ACTION_KEY][ad.subject] = chosen_key
        if chosen_key != "ACCEPT":
            annotation = (
                selected_col if chosen_key == "MANUAL_COLUMN" and selected_col else chosen_key
            )
            if _OVERRIDES_KEY not in st.session_state:
                st.session_state[_OVERRIDES_KEY] = {}
            st.session_state[_OVERRIDES_KEY].setdefault(ad.agent_name, {})[ad.subject] = annotation


def _render_consolidation_card(ad: ActionableDecision, idx: int) -> None:
    saved_actions: dict[str, str] = st.session_state.get(_ACTION_KEY, {})
    current_action = saved_actions.get(ad.subject, "ACCEPT")

    icon = _conf_color(ad.confidence)
    wb_a, wb_b = ad.workbook_pair
    with st.container(border=True):
        c1, c2 = st.columns([3, 1])
        with c1:
            st.markdown(f"**{icon} Consolidation Uncertainty** — `{wb_a}` ↔ `{wb_b}`")
            st.caption(f"Column overlap: **{ad.overlap_pct:.1f}%**")
            if ad.grain_issue:
                st.warning(f"Grain issue: {ad.grain_issue}")
            else:
                st.caption(ad.description)
        with c2:
            st.markdown("**Agent recommendation:**")
            st.info(ad.current_recommendation)

        key = _widget_key("consol", ad.subject)
        default_idx = next(
            (i for i, (k, _) in enumerate(ad.available_actions) if k == current_action), 0
        )
        labels = [label for _, label in ad.available_actions]
        chosen_label = st.radio("Your decision:", labels, index=default_idx, key=key, horizontal=False)
        chosen_key = next(k for k, l in ad.available_actions if l == chosen_label)

        if _ACTION_KEY not in st.session_state:
            st.session_state[_ACTION_KEY] = {}
        st.session_state[_ACTION_KEY][ad.subject] = chosen_key
        if chosen_key != "ACCEPT":
            if _OVERRIDES_KEY not in st.session_state:
                st.session_state[_OVERRIDES_KEY] = {}
            st.session_state[_OVERRIDES_KEY].setdefault(ad.agent_name, {})[ad.subject] = chosen_key


# ── Per-step renderers ────────────────────────────────────────────────────────

def _render_step_1() -> None:
    if _current_step() != 1:
        return
    st.subheader("Step 1 — Upload Workbooks")
    uploaded = st.file_uploader(
        "Upload one or more Excel workbooks (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="ar_upload",
    )

    if not uploaded:
        st.info("Upload at least one workbook to begin.")
        return

    with st.spinner("Loading workbooks…"):
        try:
            bundles = [load_workbook_from_bytes(f.read(), file_name=f.name) for f in uploaded]
        except Exception as exc:
            st.error(f"Failed to load workbooks: `{type(exc).__name__}: {exc}`")
            st.code(traceback.format_exc(), language="python")
            return
    st.session_state[_SS_BUNDLES] = bundles
    st.session_state.pop(_SS_DISCOVERY, None)
    _set_step(2)
    st.rerun()


def _tab_structure_label(bundle, sheet_name: str, discovery_result: "AgentResult | None") -> str:
    """Return 'Raw Data', 'Pivot Table', or 'Formula Based' for a tab (result cached)."""
    cache_key = (getattr(bundle, "file_name", ""), sheet_name)
    struct_cache: dict = st.session_state.setdefault(_SS_STRUCT_CACHE, {})
    if cache_key in struct_cache:
        return struct_cache[cache_key]
    label = _tab_structure_label_compute(bundle, sheet_name, discovery_result)
    struct_cache[cache_key] = label
    return label


def _tab_structure_label_compute(bundle, sheet_name: str, discovery_result: "AgentResult | None") -> str:
    """Internal: compute structure label without caching."""
    # Check pivot signal from discovery decisions
    if discovery_result:
        for d in discovery_result.decisions:
            if d.subject.endswith(f" / {sheet_name}") and d.signals:
                if d.signals.get("is_pivot_sheet"):
                    return "Pivot Table"
                fd = d.signals.get("formula_density_pct", 0)
                if fd >= 20:
                    return "Formula Based"
    # Fallback: scan the DataFrame
    df = bundle.sheets.get(sheet_name)
    if df is None:
        return "—"
    raw_wb = getattr(bundle, "_raw_wb", None)
    if raw_wb is not None:
        try:
            raw_ws = raw_wb[sheet_name]
            if getattr(raw_ws, "_pivots", []):
                return "Pivot Table"
        except Exception:
            pass
    return "Raw Data"


def _infer_tab_tag(sheet_name: str, source_tab: str, kpi_tabs: list[str]) -> str:
    """Return 'Source', 'KPI', or 'Ignore' for a tab based on discovery output."""
    if sheet_name == source_tab:
        return "Source"
    if sheet_name in kpi_tabs:
        return "KPI"
    return "Ignore"


def _render_step_2() -> None:
    step = _current_step()
    if step != 2:
        return

    bundles = st.session_state.get(_SS_BUNDLES, [])
    if not bundles:
        st.warning("No bundles loaded. Return to Step 1.")
        return

    # -- Phase 1: structural analysis (cached) --------------------------------
    from src.ingestion.workbook_analyzer import analyze_many
    wb_analyses = st.session_state.get(_SS_WB_ANALYSES)
    if wb_analyses is None:
        with st.spinner("Analysing workbooks…"):
            try:
                wb_analyses = analyze_many(bundles)
            except Exception as exc:
                st.error(f"Workbook analysis failed: `{type(exc).__name__}: {exc}`")
                st.code(traceback.format_exc(), language="python")
                return
        st.session_state[_SS_WB_ANALYSES] = wb_analyses

    # -- Phase 2: build RationalizationConfig (fast, uses cached analyses) ----
    discovery_result: AgentResult | None = st.session_state.get(_SS_DISCOVERY)
    if discovery_result is None:
        src_overrides = st.session_state.get(_SS_SOURCE_TAB_OVERRIDES, {})
        kpi_overrides = st.session_state.get("ar_kpi_tab_overrides", {})
        discovery_overrides: dict[str, dict] = {
            fname: {"source_tab": tab} for fname, tab in src_overrides.items()
        }
        for fname, kpi_tabs_list in kpi_overrides.items():
            if fname not in discovery_overrides:
                discovery_overrides[fname] = {}
            discovery_overrides[fname]["kpi_tabs"] = kpi_tabs_list
        try:
            discovery_result = run_discovery_phase(
                bundles,
                overrides=discovery_overrides,
                analyses=wb_analyses,
            )
        except Exception as exc:
            st.error(f"Discovery failed: `{type(exc).__name__}: {exc}`")
            st.code(traceback.format_exc(), language="python")
            return
        st.session_state[_SS_DISCOVERY] = discovery_result
        st.session_state[_SS_CONFIG] = discovery_result.output

    config = discovery_result.output
    if config is None:
        st.error("Discovery did not produce a configuration.")
        return

    # -- Pending tag edits keyed by (wb_name, sheet_name) ---------------------
    _SS_DISC_TAGS = "ar_disc_tags"
    all_tags: dict[str, dict[str, str]] = st.session_state.get(_SS_DISC_TAGS, {})

    for bundle in bundles:
        wb_name = bundle.file_name
        if wb_name not in all_tags:
            wb_cfg = config.config_for(wb_name)
            src  = wb_cfg.source_tab if wb_cfg else ""
            kpis = wb_cfg.kpi_tabs if wb_cfg else []
            all_tags[wb_name] = {s: _infer_tab_tag(s, src, kpis) for s in bundle.sheets}
    st.session_state[_SS_DISC_TAGS] = all_tags

    def _committed(wb_name: str) -> dict[str, str]:
        wb_cfg = config.config_for(wb_name)
        src  = wb_cfg.source_tab if wb_cfg else ""
        kpis = wb_cfg.kpi_tabs if wb_cfg else []
        bnd  = next(b for b in bundles if b.file_name == wb_name)
        return {s: _infer_tab_tag(s, src, kpis) for s in bnd.sheets}

    has_pending = any(
        all_tags.get(b.file_name, {}) != _committed(b.file_name)
        for b in bundles
    )

    st.subheader("Step 2 — Discovery")

    _TAG_OPTIONS = ["Source", "KPI", "Ignore"]

    # -- Workbook cards --------------------------------------------------------
    for bundle in bundles:
        wb_name    = bundle.file_name
        all_sheets = list(bundle.sheets.keys())
        cur_tags   = all_tags.get(wb_name, {})
        n_source   = sum(1 for t in cur_tags.values() if t == "Source")
        n_kpi      = sum(1 for t in cur_tags.values() if t == "KPI")

        with st.expander(
            f"**{wb_name}**  —  {len(all_sheets)} sheets · {n_source} source · {n_kpi} KPI",
            expanded=True,
        ):
            sc1, sc2, sc3 = st.columns(3)
            sc1.metric("Sheets",      len(all_sheets))
            sc2.metric("Source Tabs", n_source)
            sc3.metric("KPI Tabs",    n_kpi)

            st.markdown("")
            h_tab, h_tag, h_struct, h_rows, h_cols = st.columns([3, 2, 2, 1, 1])
            h_tab.markdown("**Tab Name**")
            h_tag.markdown("**Tag**")
            h_struct.markdown("**Structure**")
            h_rows.markdown("**Rows**")
            h_cols.markdown("**Cols**")

            for sheet_name in all_sheets:
                df_sheet = bundle.sheets.get(sheet_name)
                n_rows  = len(df_sheet) if df_sheet is not None else 0
                n_cols  = len(df_sheet.columns) if df_sheet is not None else 0
                struct  = _tab_structure_label(bundle, sheet_name, discovery_result)

                r_tab, r_tag, r_struct, r_rows, r_cols = st.columns([3, 2, 2, 1, 1])
                with r_tab:
                    if st.button(
                        sheet_name,
                        key=f"ar_disc_sel_{wb_name}_{sheet_name}",
                        use_container_width=True,
                    ):
                        st.session_state["ar_disc_explorer_tab"] = (wb_name, sheet_name)
                        st.session_state.pop("ar_disc_selected_block", None)
                        st.rerun()
                with r_tag:
                    current_tag = cur_tags.get(sheet_name, "Ignore")
                    new_tag = st.selectbox(
                        "tag",
                        _TAG_OPTIONS,
                        index=_TAG_OPTIONS.index(current_tag),
                        key=f"ar_disc_tag_{wb_name}_{sheet_name}",
                        label_visibility="collapsed",
                    )
                    if new_tag != current_tag:
                        cur_tags[sheet_name] = new_tag
                        all_tags[wb_name] = cur_tags
                        st.session_state[_SS_DISC_TAGS] = all_tags
                        st.rerun()
                r_struct.markdown(struct)
                r_rows.markdown(str(n_rows))
                r_cols.markdown(str(n_cols))

    # -- Tab Explorer ----------------------------------------------------------
    explorer_target = st.session_state.get("ar_disc_explorer_tab")
    if explorer_target:
        exp_wb_name, exp_sheet = explorer_target
        exp_bundle = next((b for b in bundles if b.file_name == exp_wb_name), None)
        if exp_bundle is not None:
            exp_df       = exp_bundle.sheets.get(exp_sheet)
            exp_tag      = all_tags.get(exp_wb_name, {}).get(exp_sheet, "Ignore")
            struct_label = _tab_structure_label(exp_bundle, exp_sheet, discovery_result)

            st.markdown("---")
            st.markdown(f"### {exp_sheet}")

            if exp_tag in ("Source", "Ignore"):
                n_rows = len(exp_df) if exp_df is not None else 0
                n_cols = len(exp_df.columns) if exp_df is not None else 0
                mc1, mc2, mc3 = st.columns(3)
                mc1.metric("Rows",      f"{n_rows:,}")
                mc2.metric("Columns",   n_cols)
                mc3.metric("Structure", struct_label)

                if exp_df is not None:
                    col_rows_data = []
                    for col in exp_df.columns:
                        col_str = str(col)
                        if col_str.startswith("Unnamed:"):
                            continue
                        dtype = exp_df[col].dtype
                        if "datetime" in str(dtype):
                            dtype_label = "Date"
                        elif "bool" in str(dtype):
                            dtype_label = "Boolean"
                        elif str(dtype).startswith(("int", "float")):
                            dtype_label = "Numeric"
                        else:
                            col_lower = col_str.lower()
                            if any(kw in col_lower for kw in ("date", "period", "month", "year", "day")):
                                dtype_label = "Date"
                            else:
                                dtype_label = "Text"
                        col_rows_data.append({"Column Name": col_str, "Data Type": dtype_label})
                    if col_rows_data:
                        st.dataframe(
                            pd.DataFrame(col_rows_data),
                            use_container_width=True,
                            hide_index=True,
                        )
                    else:
                        st.info("No readable columns found in this tab.")

            elif exp_tag == "KPI":
                blocks_cache: dict = st.session_state.setdefault(_SS_BLOCKS_CACHE, {})
                cache_key = (exp_wb_name, exp_sheet)
                if cache_key not in blocks_cache:
                    raw_wb = getattr(exp_bundle, "_raw_wb", None)
                    raw_ws = _get_raw_ws(raw_wb, exp_sheet)
                    if raw_ws is not None and getattr(raw_ws, "_pivots", []):
                        blocks_cache[cache_key] = build_kpi_blocks_from_pivot(raw_wb, exp_sheet)
                    else:
                        blocks_cache[cache_key] = scan_kpi_blocks(raw_ws) if raw_ws is not None else []
                kpi_blocks = blocks_cache[cache_key]

                n_rows = len(exp_df) if exp_df is not None else 0
                mc1, mc2, mc3 = st.columns(3)
                mc1.metric("Rows",       f"{n_rows:,}")
                mc2.metric("KPI Blocks", len(kpi_blocks))
                mc3.metric("Structure",  struct_label)

                if kpi_blocks:
                    _SS_SEL_BLOCK = "ar_disc_selected_block"
                    selected_block_name = st.session_state.get(_SS_SEL_BLOCK)
                    block_names = [blk.block_name for blk in kpi_blocks]
                    if selected_block_name not in block_names:
                        selected_block_name = block_names[0]
                        st.session_state[_SS_SEL_BLOCK] = selected_block_name

                    st.markdown("**KPI Blocks**")
                    n_chips   = len(kpi_blocks)
                    chip_cols = st.columns(min(n_chips, 8))
                    for i, blk in enumerate(kpi_blocks):
                        with chip_cols[i % len(chip_cols)]:
                            is_sel = blk.block_name == selected_block_name
                            if st.button(
                                blk.block_name,
                                key=f"ar_kpi_chip_{exp_wb_name}_{exp_sheet}_{i}",
                                type="primary" if is_sel else "secondary",
                                use_container_width=True,
                            ):
                                st.session_state[_SS_SEL_BLOCK] = blk.block_name
                                st.rerun()

                    sel_blk = next(
                        (b for b in kpi_blocks if b.block_name == selected_block_name),
                        kpi_blocks[0],
                    )
                    st.markdown("")
                    dim_col, meas_col = st.columns(2)
                    with dim_col:
                        st.markdown("**Dimensions**")
                        if sel_blk.dimensions:
                            for dim in sel_blk.dimensions:
                                st.markdown(f"- {dim}")
                        else:
                            st.caption("—")
                    with meas_col:
                        st.markdown("**Measures**")
                        if sel_blk.kpi_measures:
                            for m in sel_blk.kpi_measures:
                                st.markdown(f"- {m.display_name}")
                        else:
                            st.caption("—")
                else:
                    st.info("No KPI blocks detected in this tab.")

    # -- Action buttons -------------------------------------------------------
    st.markdown("---")
    if has_pending:
        if st.button("Apply Changes", type="primary", key="ar_btn_apply_changes"):
            new_src_overrides: dict[str, str] = {}
            new_kpi_overrides: dict[str, list[str]] = {}
            for wb_n, tag_map in all_tags.items():
                src  = next((s for s, t in tag_map.items() if t == "Source"), "")
                kpis = [s for s, t in tag_map.items() if t == "KPI"]
                if src:
                    new_src_overrides[wb_n] = src
                if kpis:
                    new_kpi_overrides[wb_n] = kpis
            st.session_state[_SS_SOURCE_TAB_OVERRIDES] = new_src_overrides
            st.session_state["ar_kpi_tab_overrides"]   = new_kpi_overrides
            st.session_state.pop(_SS_DISCOVERY, None)
            st.session_state.pop(_SS_CONSOLIDATION, None)
            st.session_state.pop(_SS_SCHEMA, None)
            st.session_state.pop(_SS_BLOCKS_CACHE, None)
            st.session_state.pop(_SS_STRUCT_CACHE, None)
            st.session_state.pop(_SS_DISC_TAGS, None)
            st.rerun()
    else:
        if step == 2:
            if st.button(
                "Confirm Discovery & Continue",
                type="primary",
                key="ar_btn_confirm_discovery",
            ):
                with st.spinner("Analysing KPI formulas and grouping files…"):
                    try:
                        kpi_result_obj = run_kpi_phase(bundles, config)
                        st.session_state[_SS_KPI] = kpi_result_obj
                        kpi_analysis = kpi_result_obj.output if kpi_result_obj else None
                        consol = run_consolidation_phase(bundles, config, kpi_result=kpi_analysis)
                        st.session_state[_SS_CONSOLIDATION] = consol
                    except Exception as exc:
                        st.error(f"Analysis failed: `{type(exc).__name__}: {exc}`")
                        st.code(traceback.format_exc(), language="python")
                        return
                _set_step(3)
                st.rerun()


def _build_column_mapping_xlsx(intel) -> bytes:
    """Column Mapping: canonical → per-file raw column name."""
    from src.schema_matching.normalizer import normalize_column_name
    # Gather all canonicals across all files
    file_names = [fp.file_name for fp in intel.file_profiles]
    # col_map: file → list of raw cols (reconstructed from file profiles via group data)
    # We'll use the group file names and collect all unique canonicals from kpi_alignments + groups
    canonicals: dict[str, dict[str, str]] = {}  # canonical → {file_name → raw_col}
    # Use kpi_alignments first
    for aln in getattr(intel, "kpi_alignments", []):
        canonicals.setdefault(aln.canonical_kpi, {})
        for fn, raw in aln.per_file_match.items():
            canonicals[aln.canonical_kpi][fn] = raw or ""
    rows = []
    for canonical, per_file in sorted(canonicals.items()):
        row = {"Canonical Column": canonical}
        for fn in file_names:
            row[fn] = per_file.get(fn, "")
        rows.append(row)
    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    df.to_excel(buf, index=False, sheet_name="Column Mapping")
    return buf.getvalue()


def _build_kpi_mapping_xlsx(intel) -> bytes:
    """KPI Mapping: canonical KPI → per-file KPI label/column."""
    file_names = [fp.file_name for fp in intel.file_profiles]
    rows = []
    for aln in getattr(intel, "kpi_alignments", []):
        row = {
            "Canonical KPI Column": aln.canonical_kpi,
            "KPI Label(s)":         ", ".join(aln.kpi_labels),
            "Alignment Status":     aln.status.title(),
            "Notes":                aln.suggestion,
        }
        for fn in file_names:
            match = aln.per_file_match.get(fn, "")
            conf  = aln.per_file_confidence.get(fn, 0.0)
            row[fn] = match if match else "(not found)"
            row[f"{fn} Confidence"] = f"{conf:.0%}" if match else ""
        rows.append(row)
    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    df.to_excel(buf, index=False, sheet_name="KPI Mapping")
    return buf.getvalue()


def _build_consolidation_assessment_xlsx(intel) -> bytes:
    """Consolidation Assessment: 5-dimension pairwise scores."""
    rows = []
    for sc in intel.pairwise_scores:
        rows.append({
            "File A":                  sc.file_a,
            "File B":                  sc.file_b,
            "Grain Compatibility":     "Yes" if sc.grain_compatible else "No",
            "Grain A":                 sc.grain_a,
            "Grain B":                 sc.grain_b,
            "Schema Similarity %":     f"{sc.schema_overlap_pct:.1%}",
            "KPI Similarity %":        f"{sc.kpi_overlap_pct:.1%}",
            "Key Compatible":          "Yes" if sc.key_compatible else "No",
            "Time Dimension Match":    "Yes" if sc.time_dimension_compatible else "No",
            "Overall Score":           f"{sc.overall_score:.1%}",
            "Recommendation":          sc.recommendation,
            "Confidence":              f"{sc.recommendation_confidence:.0%}",
            "Reasoning":               sc.reasoning,
        })
    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    df.to_excel(buf, index=False, sheet_name="Consolidation Assessment")
    return buf.getvalue()


def _build_group_analysis_xlsx(intel) -> bytes:
    """Group Analysis: group-level summary + per-file contributions."""
    rows_group = []
    rows_contrib = []
    for group in intel.groups:
        rows_group.append({
            "Group ID":          group.group_id,
            "Group Label":       group.group_label,
            "Grain Category":    group.grain_category,
            "Files":             ", ".join(group.file_names),
            "Recommendation":    group.recommendation,
            "Common Columns":    group.common_column_count,
            "Total Union Cols":  group.unique_column_count,
            "KPI Required":      group.kpi_required_column_count,
            "Discardable":       group.discardable_column_count,
            "Expected Master Cols": group.expected_master_columns,
            "Reasoning":         group.reasoning,
        })
        for fc in getattr(group, "file_contributions", []):
            rows_contrib.append({
                "Group ID":                  group.group_id,
                "File Name":                 fc.file_name,
                "Total Columns":             fc.total_columns,
                "KPI Columns":               fc.kpi_columns,
                "Common Columns Contributed": fc.common_columns_contributed,
                "Unique Columns Contributed": fc.unique_columns_contributed,
                "Discardable Columns":        fc.discardable_columns,
            })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(rows_group).to_excel(writer, index=False, sheet_name="Group Summary")
        pd.DataFrame(rows_contrib).to_excel(writer, index=False, sheet_name="File Contributions")
    return buf.getvalue()


def _build_compatibility_matrix_xlsx(intel) -> bytes:
    """Compatibility Matrix: file × group with 5-dimension scores."""
    from src.agentic_rationalization.agents.consolidation_agent import compute_file_group_compatibility
    rows = []
    for fp in intel.file_profiles:
        compat = compute_file_group_compatibility(fp.file_name, intel.groups, intel.pairwise_scores)
        for gid, scores in sorted(compat.items()):
            rows.append({
                "File":            fp.file_name,
                "Group":           f"Group {gid}",
                "Grain Score":     f"{scores['grain_score']:.0%}",
                "KPI Score":       f"{scores['kpi_score']:.0%}",
                "Schema Score":    f"{scores['schema_score']:.0%}",
                "Key Score":       f"{scores['key_score']:.0%}",
                "Time Score":      f"{scores['time_score']:.0%}",
                "Overall Score":   f"{scores['overall_score']:.0%}",
                "Recommendation":  scores["recommendation"],
            })
    df = pd.DataFrame(rows) if rows else pd.DataFrame()
    buf = io.BytesIO()
    df.to_excel(buf, index=False, sheet_name="Compatibility Matrix")
    return buf.getvalue()


def _apply_group_overrides(groups, overrides: dict) -> list:
    """Return modified groups list after applying manual overrides."""
    if not overrides:
        return groups

    # Build a dict of file_name → current group_id
    file_to_group: dict[str, int] = {}
    for group in groups:
        for fname in group.file_names:
            file_to_group[fname] = group.group_id

    # Apply overrides
    for fname, target in overrides.items():
        if target is not None:
            file_to_group[fname] = target
        else:
            file_to_group.pop(fname, None)  # excluded

    # Rebuild groups
    new_members: dict[int, list[str]] = {}
    for fname, gid in file_to_group.items():
        new_members.setdefault(gid, []).append(fname)

    # Map old group_id → original group for metadata
    orig_map = {g.group_id: g for g in groups}
    result = []
    for gid, members in sorted(new_members.items()):
        orig = orig_map.get(gid)
        if orig:
            from dataclasses import replace as dc_replace
            try:
                new_g = dc_replace(orig, file_names=members)
            except Exception:
                new_g = orig
                new_g.file_names = members
        else:
            from src.agentic_rationalization.agents.consolidation_agent import ConsolidationGroup
            new_g = ConsolidationGroup(
                group_id=gid,
                group_label=f"Custom Group {gid}",
                grain_category="custom",
                file_names=members,
                common_column_count=0,
                unique_column_count=0,
                kpi_required_column_count=0,
                discardable_column_count=0,
                recommendation="Consolidate" if len(members) > 1 else "Keep Separate",
                expected_master_columns=0,
                reasoning="Manual group created by user.",
            )
        result.append(new_g)
    return result


def _kpis_for_group(group_files: list[str], kpi_alignments: list, max_show: int = 8) -> list[str]:
    """KPI names present in ≥2 files of the group (or any file if single-file)."""
    if not kpi_alignments:
        return []
    threshold = min(2, len(group_files))
    result: list[str] = []
    for aln in kpi_alignments:
        matches = sum(1 for fn in group_files if aln.per_file_match.get(fn))
        if matches >= threshold:
            result.append(aln.canonical_kpi)
        if len(result) >= max_show:
            break
    return result


def _unique_kpis_for_file(fname: str, all_files: list[str], kpi_alignments: list, max_show: int = 8) -> list[str]:
    """KPI names present only in *fname* and no other file."""
    if not kpi_alignments:
        return []
    result: list[str] = []
    for aln in kpi_alignments:
        if not aln.per_file_match.get(fname):
            continue
        other_matches = sum(1 for fn in all_files if fn != fname and aln.per_file_match.get(fn))
        if other_matches == 0:
            result.append(aln.canonical_kpi)
        if len(result) >= max_show:
            break
    return result


def _render_step_3() -> None:
    step = _current_step()
    if step != 3:
        return

    st.subheader("Step 3 — Consolidation")

    consol_result: AgentResult | None = st.session_state.get(_SS_CONSOLIDATION)

    if consol_result is None or consol_result.output is None:
        st.warning("Consolidation analysis not yet available.")
        return

    intel = consol_result.output.get("intelligence")
    if intel is None:
        st.info("No intelligence output available.")
        return

    _missing = _validate_pipeline_inputs(
        "Step 3 — Consolidation",
        {
            "ConsolidationIntelligenceResult": (
                intel,
                ["file_profiles", "groups", "pairwise_scores", "final_recommendation"],
            ),
        },
    )
    if _missing:
        _render_validation_error("Step 3 — Consolidation", _missing)
        return

    group_overrides: dict = st.session_state.get(_SS_GROUP_OVERRIDES, {})
    effective_groups = _apply_group_overrides(intel.groups, group_overrides)
    file_names = [fp.file_name for fp in intel.file_profiles]
    kpi_alignments = getattr(intel, "kpi_alignments", [])
    kpi_redundancy = getattr(intel, "kpi_redundancy", [])

    # Accepted-state tracking (purely visual, no pipeline effect)
    _ACCEPTED_KEY = "ar_consol_accepted"
    accepted: set = st.session_state.setdefault(_ACCEPTED_KEY, set())
    _OVERRIDE_OPEN_KEY = "ar_consol_override_open"
    override_open: set = st.session_state.setdefault(_OVERRIDE_OPEN_KEY, set())

    # ── Build governance recommendations (KPI-driven) ────────────────────────
    redundancy_map = {r.file_name: r for r in kpi_redundancy}

    def _kpi_coverage_pct(fname: str) -> float:
        r = redundancy_map.get(fname)
        if r is None or r.total_kpi_columns == 0:
            return 0.0
        shared = r.total_kpi_columns - r.unique_kpi_columns
        return round(shared / r.total_kpi_columns * 100, 1)

    def _kpi_uniqueness_pct(fname: str) -> float:
        r = redundancy_map.get(fname)
        if r is None or r.total_kpi_columns == 0:
            return 100.0
        return round(r.unique_kpi_columns / r.total_kpi_columns * 100, 1)

    # Decommission: fully redundant files (all KPIs exist elsewhere)
    decommission_names: set[str] = {r.file_name for r in kpi_redundancy if r.redundant}

    # Consolidation groups: multi-file groups excluding decommissioned files
    consol_groups = [
        g for g in effective_groups
        if g.recommendation == "Consolidate" and len(g.file_names) > 1
    ]

    # Within each consolidation group find canonical (most KPI cols) and
    # non-canonical files → CONSOLIDATE & MERGE
    canonical_names: set[str] = set()
    merge_entries: list[dict] = []  # {fname, canonical_fname, group}
    keep_certify_from_groups: list[dict] = []  # {fname, group}

    for group in consol_groups:
        active_files = [f for f in group.file_names if f not in decommission_names]
        if not active_files:
            continue
        canonical = max(
            active_files,
            key=lambda f: (redundancy_map[f].total_kpi_columns if f in redundancy_map else 0),
        )
        canonical_names.add(canonical)
        keep_certify_from_groups.append({"file_name": canonical, "group": group})
        for fname in active_files:
            if fname != canonical:
                merge_entries.append({"file_name": fname, "canonical": canonical, "group": group})

    # Files not in any multi-file consolidation group and not decommissioned → KEEP & CERTIFY
    grouped_files: set[str] = set()
    for group in consol_groups:
        grouped_files |= set(group.file_names)

    standalone_keep: list[str] = [
        fp.file_name for fp in intel.file_profiles
        if fp.file_name not in grouped_files and fp.file_name not in decommission_names
    ]

    keep_certify_files = [e["file_name"] for e in keep_certify_from_groups] + standalone_keep
    decommission_list  = [r for r in kpi_redundancy if r.redundant]

    # ── Summary bar ──────────────────────────────────────────────────────────
    s1, s2, s3 = st.columns(3)
    s1.metric("Keep & Certify",      len(keep_certify_files),  help="Canonical workbooks — retain as authoritative sources")
    s2.metric("Consolidate & Merge", len(merge_entries),        help="Files whose KPIs are covered by a canonical workbook")
    s3.metric("Decommission",        len(decommission_list),    help="Fully redundant files — all KPIs exist elsewhere")

    st.markdown("---")

    # ── KEEP & CERTIFY ───────────────────────────────────────────────────────
    if keep_certify_files:
        st.markdown("### Keep & Certify")
        st.caption("These workbooks are the authoritative sources for their KPI domains.")
        for fname in keep_certify_files:
            card_key  = f"gov_keep_{fname}"
            is_accepted = card_key in accepted
            coverage  = _kpi_coverage_pct(fname)
            uniqueness = _kpi_uniqueness_pct(fname)
            r = redundancy_map.get(fname)
            total_kpis = r.total_kpi_columns if r else 0

            with st.container(border=True):
                h1, h2 = st.columns([3, 1])
                with h1:
                    st.markdown(
                        f"{'✅ ' if is_accepted else ''}**{fname}**  "
                        f"<span style='background:#d4edda;color:#155724;padding:2px 8px;"
                        f"border-radius:4px;font-size:0.85em;'>KEEP & CERTIFY</span>",
                        unsafe_allow_html=True,
                    )
                with h2:
                    st.metric("Total KPIs", total_kpis)

                mc1, mc2 = st.columns(2)
                mc1.metric("KPI Coverage",   f"{coverage:.0f}%",   help="% of KPIs shared with at least one other file")
                mc2.metric("KPI Uniqueness", f"{uniqueness:.0f}%", help="% of KPIs found only in this file")

                # Governance rationale
                if uniqueness >= 70:
                    rationale = (
                        f"This workbook owns {uniqueness:.0f}% of its KPIs exclusively — "
                        "no other file covers this domain. Retain as the canonical source."
                    )
                elif fname in canonical_names:
                    rationale = (
                        f"Designated canonical within its consolidation group: "
                        f"broadest KPI set ({total_kpis} KPIs) and highest coverage breadth."
                    )
                else:
                    rationale = (
                        f"Standalone workbook with {total_kpis} KPIs. "
                        "No overlap sufficient to justify merging into another file."
                    )
                st.markdown(f"**Governance Rationale:** {rationale}")

                common_kpis = _kpis_for_group(
                    [fname] + [e["file_name"] for e in merge_entries if e["canonical"] == fname],
                    kpi_alignments,
                )
                if common_kpis:
                    st.markdown("**Common KPIs:**")
                    for kpi in common_kpis:
                        st.markdown(f"- {kpi}")
                    if len(common_kpis) == 8:
                        st.caption("(showing first 8 — see View Analysis for full list)")

                if st.button(
                    "✅ Accept" if not is_accepted else "✅ Accepted",
                    key=f"accept_{card_key}",
                    type="primary" if not is_accepted else "secondary",
                ):
                    accepted.add(card_key)
                    st.session_state[_ACCEPTED_KEY] = accepted
                    st.rerun()

        st.markdown("---")

    # ── CONSOLIDATE & MERGE ──────────────────────────────────────────────────
    if merge_entries:
        st.markdown("### Consolidate & Merge")
        st.caption("These workbooks can be merged into a canonical file — their KPIs are substantially covered by the target.")
        for entry in merge_entries:
            fname     = entry["file_name"]
            canonical = entry["canonical"]
            card_key  = f"gov_merge_{fname}"
            is_accepted = card_key in accepted
            is_override = card_key in override_open
            coverage   = _kpi_coverage_pct(fname)
            uniqueness = _kpi_uniqueness_pct(fname)
            r = redundancy_map.get(fname)
            total_kpis = r.total_kpi_columns if r else 0

            with st.container(border=True):
                h1, h2 = st.columns([3, 1])
                with h1:
                    st.markdown(
                        f"{'✅ ' if is_accepted else ''}**{fname}**  "
                        f"<span style='background:#fff3cd;color:#856404;padding:2px 8px;"
                        f"border-radius:4px;font-size:0.85em;'>CONSOLIDATE & MERGE</span>",
                        unsafe_allow_html=True,
                    )
                with h2:
                    st.metric("Total KPIs", total_kpis)

                st.markdown(f"**Merge Into:** `{canonical}`")

                mc1, mc2 = st.columns(2)
                mc1.metric("KPI Coverage",   f"{coverage:.0f}%",   help="% of KPIs shared with the target file")
                mc2.metric("KPI Uniqueness", f"{uniqueness:.0f}%", help="% of KPIs found only in this file")

                if coverage >= 80:
                    rationale = (
                        f"{coverage:.0f}% of this workbook's KPIs are already present in `{canonical}`. "
                        "Merge to eliminate duplication and consolidate reporting."
                    )
                else:
                    rationale = (
                        f"Significant KPI overlap ({coverage:.0f}%) with `{canonical}`. "
                        f"The {uniqueness:.0f}% unique KPIs should be migrated before decommissioning."
                    )
                st.markdown(f"**Governance Rationale:** {rationale}")

                common_kpis = _kpis_for_group([fname, canonical], kpi_alignments)
                if common_kpis:
                    st.markdown("**Common KPIs:**")
                    for kpi in common_kpis:
                        st.markdown(f"- {kpi}")
                    if len(common_kpis) == 8:
                        st.caption("(showing first 8 — see View Analysis for full list)")

                unique_kpis = _unique_kpis_for_file(fname, file_names, kpi_alignments)
                if unique_kpis:
                    st.markdown("**Unique KPIs (must migrate before merge):**")
                    for kpi in unique_kpis:
                        st.markdown(f"- {kpi}")

                if not is_override:
                    ba, bb = st.columns([1, 1])
                    with ba:
                        if st.button(
                            "✅ Accept" if not is_accepted else "✅ Accepted",
                            key=f"accept_{card_key}",
                            type="primary" if not is_accepted else "secondary",
                        ):
                            accepted.add(card_key)
                            st.session_state[_ACCEPTED_KEY] = accepted
                            st.rerun()
                    with bb:
                        if st.button("✏️ Override Target", key=f"override_open_{card_key}"):
                            override_open.add(card_key)
                            st.session_state[_OVERRIDE_OPEN_KEY] = override_open
                            st.rerun()
                else:
                    st.markdown("**Override — reassign merge target:**")
                    group_ids      = sorted({g.group_id for g in intel.groups})
                    max_gid        = max(group_ids) if group_ids else 1
                    options_labels = [f"Group {gid}" for gid in group_ids] + [f"New Group {max_gid + 1}", "Keep Separate"]
                    options_ids    = group_ids + [max_gid + 1, None]
                    current_gid    = next((g.group_id for g in intel.groups if fname in g.file_names), group_ids[0] if group_ids else 0)
                    override_val   = group_overrides.get(fname, current_gid)
                    safe_val       = override_val if override_val in options_ids else current_gid
                    current_idx    = options_ids.index(safe_val) if safe_val in options_ids else 0
                    chosen_label   = st.selectbox(
                        "Assign to",
                        options_labels,
                        index=current_idx,
                        key=f"grp_ov_merge_{fname}",
                        label_visibility="collapsed",
                    )
                    new_gid = options_ids[options_labels.index(chosen_label)]
                    oa, ob  = st.columns([1, 1])
                    with oa:
                        if st.button("Apply", key=f"override_apply_{card_key}", type="primary"):
                            group_overrides[fname] = new_gid
                            st.session_state[_SS_GROUP_OVERRIDES] = group_overrides
                            override_open.discard(card_key)
                            st.session_state[_OVERRIDE_OPEN_KEY] = override_open
                            st.rerun()
                    with ob:
                        if st.button("Cancel", key=f"override_cancel_{card_key}"):
                            override_open.discard(card_key)
                            st.session_state[_OVERRIDE_OPEN_KEY] = override_open
                            st.rerun()

        st.markdown("---")

    # ── DECOMMISSION ─────────────────────────────────────────────────────────
    if decommission_list:
        st.markdown("### Decommission")
        st.caption("These workbooks are fully redundant — every KPI they contain already exists in another file.")
        for r in decommission_list:
            card_key   = f"gov_decomm_{r.file_name}"
            is_accepted = card_key in accepted
            is_override = card_key in override_open
            coverage    = _kpi_coverage_pct(r.file_name)
            uniqueness  = _kpi_uniqueness_pct(r.file_name)

            with st.container(border=True):
                h1, h2 = st.columns([3, 1])
                with h1:
                    st.markdown(
                        f"{'✅ ' if is_accepted else ''}**{r.file_name}**  "
                        f"<span style='background:#f8d7da;color:#721c24;padding:2px 8px;"
                        f"border-radius:4px;font-size:0.85em;'>DECOMMISSION</span>",
                        unsafe_allow_html=True,
                    )
                with h2:
                    st.metric("Total KPIs", r.total_kpi_columns)

                if r.covered_by:
                    st.markdown("**Covered By:** " + ", ".join(f"`{f}`" for f in r.covered_by))

                mc1, mc2 = st.columns(2)
                mc1.metric("KPI Coverage",   f"{coverage:.0f}%",   help="% of KPIs shared with other files")
                mc2.metric("KPI Uniqueness", f"{uniqueness:.0f}%", help="% of KPIs found only in this file")

                rationale = (
                    f"All {r.total_kpi_columns} KPIs in this workbook are already present in "
                    + (", ".join(f"`{f}`" for f in r.covered_by) if r.covered_by else "other files")
                    + ". Decommission to remove duplication from the reporting estate."
                )
                st.markdown(f"**Governance Rationale:** {rationale}")

                if not is_override:
                    ba, bb = st.columns([1, 1])
                    with ba:
                        if st.button(
                            "✅ Accept" if not is_accepted else "✅ Accepted",
                            key=f"accept_{card_key}",
                            type="primary" if not is_accepted else "secondary",
                        ):
                            accepted.add(card_key)
                            st.session_state[_ACCEPTED_KEY] = accepted
                            st.rerun()
                    with bb:
                        if st.button("✏️ Override — Keep This File", key=f"override_open_{card_key}"):
                            override_open.add(card_key)
                            st.session_state[_OVERRIDE_OPEN_KEY] = override_open
                            st.rerun()
                else:
                    st.markdown("**Override — assign to a consolidation group instead:**")
                    group_ids      = sorted({g.group_id for g in intel.groups})
                    max_gid        = max(group_ids) if group_ids else 1
                    options_labels = [f"Group {gid}" for gid in group_ids] + [f"New Group {max_gid + 1}"]
                    options_ids    = group_ids + [max_gid + 1]
                    chosen_label   = st.selectbox(
                        "Assign to group:",
                        options_labels,
                        key=f"grp_ov_decomm_{r.file_name}",
                        label_visibility="collapsed",
                    )
                    new_gid = options_ids[options_labels.index(chosen_label)]
                    oa, ob  = st.columns([1, 1])
                    with oa:
                        if st.button("Apply", key=f"override_apply_{card_key}", type="primary"):
                            group_overrides[r.file_name] = new_gid
                            st.session_state[_SS_GROUP_OVERRIDES] = group_overrides
                            override_open.discard(card_key)
                            st.session_state[_OVERRIDE_OPEN_KEY] = override_open
                            st.rerun()
                    with ob:
                        if st.button("Cancel", key=f"override_cancel_{card_key}"):
                            override_open.discard(card_key)
                            st.session_state[_OVERRIDE_OPEN_KEY] = override_open
                            st.rerun()

    # ── View Analysis (hidden by default) ───────────────────────────────────
    with st.expander("View Analysis", expanded=False):
        from src.agentic_rationalization.agents.consolidation_agent import compute_file_group_compatibility

        ana_tab_groups, ana_tab_pairwise, ana_tab_kpi = st.tabs([
            "Consolidation Groups", "Pairwise Scores", "KPI Redundancy"
        ])

        with ana_tab_groups:
            for group in effective_groups:
                icon = "🟢" if group.recommendation == "Consolidate" else "🟡"
                with st.expander(
                    f"{icon} Group {group.group_id} — {group.group_label}  [{group.recommendation}]",
                    expanded=False,
                ):
                    st.caption(group.reasoning)
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Common Columns",  group.common_column_count)
                    c2.metric("Union Cols",       group.unique_column_count)
                    c3.metric("KPI Required",     group.kpi_required_column_count)
                    c4.metric("Discardable",      group.discardable_column_count)
                    contributions = getattr(group, "file_contributions", [])
                    if contributions:
                        contrib_rows = []
                        for fc in contributions:
                            contrib_rows.append({
                                "File":                fc.file_name,
                                "Total Cols":          fc.total_columns,
                                "KPI Cols":            fc.kpi_columns,
                                "Common Contributed":  fc.common_columns_contributed,
                                "Unique Contributed":  fc.unique_columns_contributed,
                                "Discardable":         fc.discardable_columns,
                            })
                        st.dataframe(pd.DataFrame(contrib_rows), use_container_width=True, hide_index=True)
                    else:
                        for fname in group.file_names:
                            st.markdown(f"- `{fname}`")
                    if group.recommendation == "Keep Separate" and group.incompatibility_detail:
                        st.info(group.incompatibility_detail)

            # KPI redundancy table
            if kpi_redundancy:
                st.markdown("---")
                st.markdown("**KPI Redundancy**")
                for r in kpi_redundancy:
                    icon = "🗑" if r.redundant else ("⚠️" if r.recommendation.startswith("Partial") else "✅")
                    with st.expander(
                        f"{icon} {r.file_name} — {r.recommendation} "
                        f"({r.total_kpi_columns} KPI cols, {r.unique_kpi_columns} unique)",
                        expanded=False,
                    ):
                        st.markdown(r.reason)
                        if r.covered_by:
                            st.caption("Covered by: " + ", ".join(r.covered_by))

            # Downloads
            st.markdown("---")
            st.markdown("**Download Analysis**")
            dl1, dl2, dl3, dl4, dl5 = st.columns(5)
            with dl1:
                try:
                    st.download_button(
                        "Column Mapping.xlsx", data=_build_column_mapping_xlsx(intel),
                        file_name="Column_Mapping.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_col_map",
                    )
                except Exception as e:
                    st.caption(f"Unavailable: {e}")
            with dl2:
                try:
                    st.download_button(
                        "KPI Mapping.xlsx", data=_build_kpi_mapping_xlsx(intel),
                        file_name="KPI_Mapping.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_kpi_map",
                    )
                except Exception as e:
                    st.caption(f"Unavailable: {e}")
            with dl3:
                try:
                    st.download_button(
                        "Consolidation Assessment.xlsx",
                        data=_build_consolidation_assessment_xlsx(intel),
                        file_name="Consolidation_Assessment.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_consol_assess",
                    )
                except Exception as e:
                    st.caption(f"Unavailable: {e}")
            with dl4:
                try:
                    st.download_button(
                        "Group Analysis.xlsx", data=_build_group_analysis_xlsx(intel),
                        file_name="Group_Analysis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_group_analysis",
                    )
                except Exception as e:
                    st.caption(f"Unavailable: {e}")
            with dl5:
                try:
                    st.download_button(
                        "Compatibility Matrix.xlsx",
                        data=_build_compatibility_matrix_xlsx(intel),
                        file_name="Compatibility_Matrix.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_compat_matrix",
                    )
                except Exception as e:
                    st.caption(f"Unavailable: {e}")

        with ana_tab_pairwise:
            if not intel.pairwise_scores:
                st.info("No pairwise scores (single file or no comparison possible).")
            else:
                for sc in intel.pairwise_scores:
                    with st.expander(
                        f"`{sc.file_a}` vs `{sc.file_b}` — {sc.recommendation}",
                        expanded=False,
                    ):
                        g1, g2, g3, g4, g5, g6 = st.columns(6)
                        g1.metric("Grain Compatible", "✓" if sc.grain_compatible else "✗")
                        g2.metric("Schema Similarity", f"{sc.schema_overlap_pct:.0%}")
                        g3.metric("KPI Similarity",    f"{sc.kpi_overlap_pct:.0%}")
                        g4.metric("Key Compatible",    "Yes" if sc.key_compatible else "No")
                        g5.metric("Time Dimension",    "Match" if sc.time_dimension_compatible else "Mismatch")
                        g6.metric("Overall Score",     f"{sc.overall_score:.0%}")
                        st.caption(sc.reasoning)
                        if not sc.grain_compatible:
                            st.info(
                                f"Grain mismatch: `{sc.file_a}` is {sc.grain_a}-level, "
                                f"`{sc.file_b}` is {sc.grain_b}-level."
                            )

        with ana_tab_kpi:
            kpi_alignments_full = getattr(intel, "kpi_alignments", [])
            if not kpi_alignments_full:
                st.info("No KPI alignment data available.")
            else:
                aligned = [a for a in kpi_alignments_full if a.status == "aligned"]
                partial = [a for a in kpi_alignments_full if a.status == "partial"]
                missing = [a for a in kpi_alignments_full if a.status == "missing"]
                st.metric("Fully Aligned", len(aligned))
                if aligned:
                    with st.expander(f"Fully Aligned ({len(aligned)})", expanded=False):
                        rows = []
                        for a in aligned:
                            row = {"Canonical KPI": a.canonical_kpi, "KPI Label(s)": ", ".join(a.kpi_labels)}
                            for fn in file_names:
                                row[fn] = a.per_file_match.get(fn, "")
                            rows.append(row)
                        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
                if partial:
                    with st.expander(f"Partially Aligned ({len(partial)})", expanded=False):
                        rows = []
                        for a in partial:
                            row = {"Canonical KPI": a.canonical_kpi, "Status": a.suggestion}
                            for fn in file_names:
                                conf = a.per_file_confidence.get(fn, 0.0)
                                raw  = a.per_file_match.get(fn, "")
                                row[fn] = f"{raw} ({conf:.0%})" if raw else "—"
                            rows.append(row)
                        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
                if missing:
                    with st.expander(f"Missing ({len(missing)})", expanded=False):
                        st.dataframe(
                            pd.DataFrame([{"Canonical": a.canonical_kpi, "Labels": ", ".join(a.kpi_labels)} for a in missing]),
                            use_container_width=True, hide_index=True,
                        )

    # ── Continue ─────────────────────────────────────────────────────────────
    st.markdown("---")
    if step == 3:
        if st.button("✅ Confirm Groups & Continue", type="primary", key="ar_btn_confirm_consol"):
            if group_overrides:
                st.session_state[_SS_GROUP_OVERRIDES] = group_overrides
            _set_step(4)
            st.rerun()




def _render_step_4() -> None:
    step = _current_step()
    if step != 4:
        return

    st.subheader("Step 4 — Select Group to Process")

    consol_result: AgentResult | None = st.session_state.get(_SS_CONSOLIDATION)
    if consol_result is None or consol_result.output is None:
        st.warning("Consolidation analysis not available.")
        return

    intel = consol_result.output.get("intelligence")
    if intel is None or not intel.groups:
        st.info("No groups available — nothing to select.")
        return

    # Defensive validation
    _missing = _validate_pipeline_inputs(
        "Step 4 — Select Group",
        {
            "ConsolidationIntelligenceResult": (
                intel,
                ["groups", "file_profiles"],
            ),
        },
    )
    if _missing:
        _render_validation_error("Step 4 — Select Group", _missing)
        return

    group_overrides: dict = st.session_state.get(_SS_GROUP_OVERRIDES, {})
    effective_groups = _apply_group_overrides(intel.groups, group_overrides)

    options = []
    for group in effective_groups:
        files_str = ", ".join(group.file_names[:3])
        if len(group.file_names) > 3:
            files_str += f" + {len(group.file_names) - 3} more"
        rec_tag = "Recommended: Consolidate" if group.recommendation == "Consolidate" else "Keep Separate"
        options.append(
            f"Group {group.group_id} — {group.group_label} ({len(group.file_names)} file(s): {files_str})  [{rec_tag}]"
        )

    prev_group = st.session_state.get(_SS_SELECTED_GROUP)
    prev_idx = 0
    if prev_group:
        for i, group in enumerate(effective_groups):
            if set(group.file_names) == set(prev_group):
                prev_idx = i
                break

    chosen_label = st.radio(
        "Select the group to process through Schema and KPI analysis:",
        options,
        index=prev_idx,
        key="ar_group_radio",
    )
    chosen_idx = options.index(chosen_label)
    chosen_group = effective_groups[chosen_idx]

    st.markdown(f"**Selected:** {len(chosen_group.file_names)} file(s)")
    for f in chosen_group.file_names:
        st.markdown(f"  - `{f}`")

    if step == 4:
        if st.button("▶ Process Selected Group", type="primary", key="ar_btn_process_group"):
            st.session_state[_SS_SELECTED_GROUP] = chosen_group.file_names
            st.session_state.pop(_SS_SCHEMA, None)
            _set_step(5)
            st.rerun()


def _render_step_5() -> None:
    step = _current_step()
    if step != 5:
        return

    st.subheader("Step 5 — Schema Rationalization")

    bundles       = st.session_state.get(_SS_BUNDLES, [])
    config        = st.session_state.get(_SS_CONFIG)
    kpi_result_obj: AgentResult | None = st.session_state.get(_SS_KPI)
    kpi_analysis  = kpi_result_obj.output if kpi_result_obj else None
    group_files   = st.session_state.get(_SS_SELECTED_GROUP, [])

    if not group_files:
        st.warning("No group selected. Return to Step 4.")
        return

    # Parse action overrides into force params
    force_merge:    dict[str, str] = {}
    force_separate: set[str]       = set()
    force_exclude:  set[str]       = set()
    for subject, action_key in (st.session_state.get(_ACTION_KEY) or {}).items():
        if action_key in ("MERGE", "MERGE_ANYWAY"):
            from src.agentic_rationalization.orchestrator import _extract_canonical_from_subject
            canonical_a = _extract_canonical_from_subject(subject)
            if canonical_a:
                force_merge[canonical_a] = canonical_a
        elif action_key == "KEEP_SEPARATE":
            from src.agentic_rationalization.orchestrator import _extract_canonical_from_subject
            canonical_a = _extract_canonical_from_subject(subject)
            if canonical_a:
                force_separate.add(canonical_a)
        elif action_key == "EXCLUDE_A":
            from src.agentic_rationalization.orchestrator import _extract_canonical_from_subject
            canonical_a = _extract_canonical_from_subject(subject)
            if canonical_a:
                force_exclude.add(canonical_a)

    schema_result: AgentResult | None = st.session_state.get(_SS_SCHEMA)
    if schema_result is None:
        with st.spinner("Running Schema Agent…"):
            try:
                schema_result = run_schema_phase(
                    bundles, config, kpi_analysis, group_files,
                    force_merge=force_merge if force_merge else None,
                    force_separate=force_separate if force_separate else None,
                    force_exclude=force_exclude if force_exclude else None,
                )
                st.session_state[_SS_SCHEMA] = schema_result
            except Exception as exc:
                st.error(f"Schema Agent failed: `{type(exc).__name__}: {exc}`")
                st.code(traceback.format_exc(), language="python")
                return

    source_analysis = schema_result.output
    if source_analysis is None:
        for w in schema_result.warnings:
            st.warning(w)
        st.info("Schema analysis did not produce results.")
        return

    # Defensive validation — ensure SourceAnalysisResult has the attributes we need
    _missing = _validate_pipeline_inputs(
        "Step 5 — Schema Rationalization",
        {
            "SourceAnalysisResult": (
                source_analysis,
                ["column_profiles", "common_columns", "similar_columns", "unique_columns"],
            ),
        },
    )
    if _missing:
        _render_validation_error("Step 5 — Schema Rationalization", _missing)
        return

    # Schema Dashboard
    all_profiles = source_analysis.column_profiles
    common_count = len(source_analysis.common_columns)
    unique_count = len([p for p in all_profiles if p.match_class == "unique"])
    similar_count = len([p for p in all_profiles if p.match_class == "similar"])
    merged_count = len([p for p in all_profiles if p.match_class == "common"])
    discarded_count = len([p for p in all_profiles if getattr(p, "excluded", False)])
    kpi_referenced_set = getattr(kpi_analysis, "referenced_canonicals", set()) if kpi_analysis else set()
    kpi_ref_count = len([p for p in all_profiles if p.canonical_name in kpi_referenced_set])

    st.markdown("**Schema Dashboard**")
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Total Source Columns", len(all_profiles))
    m2.metric("Common Columns", common_count)
    m3.metric("Unique Columns", unique_count)
    m4.metric("Discarded Columns", discarded_count)
    m5.metric("Merged Columns", merged_count)
    m6.metric("KPI Referenced", kpi_ref_count)

    # Similar column pairs highlighted section
    similar_profiles = [p for p in all_profiles if p.match_class == "similar"]
    if similar_profiles:
        with st.expander(f"Similar Column Pairs — {len(similar_profiles)} pair(s)", expanded=False):
            st.caption(
                "These columns were identified as similar across workbooks. "
                "The decision below reflects agent analysis + any overrides you applied."
            )
            for p in similar_profiles:
                # similar_to is list[str] per ColumnProfile contract
                similar_to_val = getattr(p, "similar_to", None)
                if isinstance(similar_to_val, list):
                    similar_to_names: list[str] = similar_to_val
                elif isinstance(similar_to_val, str):
                    similar_to_names = [similar_to_val]
                else:
                    similar_to_names = []
                kpi_vetoed = (
                    p.canonical_name in kpi_referenced_set
                    and any(s in kpi_referenced_set for s in similar_to_names)
                )
                verdict_color = "orange"
                verdict_text  = "KEPT SEPARATE"
                if kpi_vetoed:
                    verdict_color = "red"
                    verdict_text  = "BLOCKED BY KPI VETO"
                elif getattr(p, "excluded", False):
                    verdict_color = "gray"
                    verdict_text  = "EXCLUDED"

                similar_to_display = ", ".join(similar_to_names) if similar_to_names else "—"
                c1, c2, c3 = st.columns([2, 2, 3])
                c1.markdown(f"**`{p.canonical_name}`**")
                c2.markdown(f"Similar to: `{similar_to_display}`")
                conf = getattr(p, "similarity_score", None) or getattr(p, "confidence", 0.0)
                c3.markdown(f":{verdict_color}[{verdict_text}]  similarity: {conf:.0%}")
                if kpi_vetoed:
                    st.caption(
                        "KPI veto applied — both columns are referenced by KPI formulas "
                        "and cannot be merged without breaking KPI calculations."
                    )

    # Duplicate Column Decisions
    st.markdown("**Column Decisions**")

    _DECISION_COLORS = {
        "common":   ("green",  "MERGED"),
        "similar":  ("orange", "KEPT SEPARATE (similar)"),
        "unique":   ("blue",   "UNIQUE"),
    }
    kpi_label_map: dict[str, list[str]] = {}
    if kpi_analysis:
        for dep in getattr(kpi_analysis, "dependencies", []):
            for canonical in getattr(dep, "canonical_source_columns", []):
                label = getattr(dep, "kpi_label", "")
                if label:
                    kpi_label_map.setdefault(canonical, []).append(label)

    visible_profiles = [p for p in all_profiles if p.match_class in ("common", "similar")]
    if visible_profiles:
        rows = []
        for p in visible_profiles:
            color, decision_label = _DECISION_COLORS.get(p.match_class, ("grey", p.match_class))
            kpi_names = kpi_label_map.get(p.canonical_name, [])
            kpi_usage = ("Yes — used by: " + ", ".join(kpi_names[:3])) if kpi_names else "No"
            orig_names = getattr(p, "original_names", [])
            rows.append({
                "Column":         p.canonical_name,
                "Decision":       decision_label,
                "Original Names": " | ".join(orig_names[:4]) if orig_names else "—",
                "KPI Usage":      kpi_usage,
                "Reasoning":      getattr(p, "reasoning", "—"),
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    else:
        st.info("No similar or common columns found.")

    # Override section (kept from original review panel)
    st.markdown("**Override Column Decisions**")
    st.caption(
        "These column pairs have similar names. Decide whether they represent the same "
        "business concept (merge) or different concepts (keep separate)."
    )

    # Build a temporary pipeline-result-like object to reuse classify_actionable_decisions
    # We synthesise a minimal PipelineResult substitute using a local helper
    _fake_pipeline_results = [schema_result]
    if kpi_result_obj:
        _fake_pipeline_results.append(kpi_result_obj)

    class _FakePipeline:
        def __init__(self, results):
            self.agent_results = results
        def get_result(self, name):
            return next((r for r in self.agent_results if r.agent_name == name), None)

    fake_pipeline = _FakePipeline(_fake_pipeline_results)
    actionable = classify_actionable_decisions(fake_pipeline)
    sim_items = [a for a in actionable if a.action_type == "SIMILAR_COLUMN"]
    if sim_items:
        for i, ad in enumerate(sim_items):
            _render_similar_column_card(ad, i)
        if st.button("🔄 Re-run Schema with Overrides", key="ar_schema_rerun"):
            st.session_state.pop(_SS_SCHEMA, None)
            st.rerun()
    else:
        st.success("No ambiguous column pairs require manual review.")

    with st.expander("Agent Details", expanded=False):
        _render_agent_card(schema_result)

    if step == 5:
        if st.button("✅ Confirm Schema & Continue", type="primary", key="ar_btn_confirm_schema"):
            # Run KPI phase now
            if st.session_state.get(_SS_KPI) is None:
                with st.spinner("Running KPI Agent…"):
                    try:
                        kpi_result = run_kpi_phase(bundles, config)
                        st.session_state[_SS_KPI] = kpi_result
                    except Exception as exc:
                        st.error(f"KPI Agent failed: `{type(exc).__name__}: {exc}`")
                        st.code(traceback.format_exc(), language="python")
                        return
            _set_step(6)
            st.rerun()


def _render_step_6() -> None:
    step = _current_step()
    if step != 6:
        return

    st.subheader("Step 6 — KPI Analysis")

    bundles = st.session_state.get(_SS_BUNDLES, [])
    config  = st.session_state.get(_SS_CONFIG)
    kpi_result_obj: AgentResult | None = st.session_state.get(_SS_KPI)

    if kpi_result_obj is None:
        with st.spinner("Running KPI Agent…"):
            try:
                kpi_result_obj = run_kpi_phase(bundles, config)
                st.session_state[_SS_KPI] = kpi_result_obj
            except Exception as exc:
                st.error(f"KPI Agent failed: `{type(exc).__name__}: {exc}`")
                st.code(traceback.format_exc(), language="python")
                return

    kpi_analysis = kpi_result_obj.output

    # Defensive validation (only when kpi_analysis is present)
    if kpi_analysis is not None:
        _missing = _validate_pipeline_inputs(
            "Step 6 — KPI Analysis",
            {
                "KpiAnalysisResult": (
                    kpi_analysis,
                    ["dependencies", "referenced_canonicals"],
                ),
            },
        )
        if _missing:
            _render_validation_error("Step 6 — KPI Analysis", _missing)
            return

    if kpi_result_obj.warnings:
        for w in kpi_result_obj.warnings:
            st.warning(w)

    if kpi_analysis is None:
        st.info("KPI analysis did not produce results.")
        with st.expander("Agent Details", expanded=False):
            _render_agent_card(kpi_result_obj)
        if step == 6:
            if st.button("✅ Confirm KPI Analysis & Continue", type="primary", key="ar_btn_confirm_kpi"):
                _set_step(7)
                st.rerun()
        return

    # Formula type breakdown
    deps = getattr(kpi_analysis, "dependencies", [])
    if deps:
        type_counts: dict[str, int] = {}
        for dep in deps:
            ftype = getattr(dep, "formula_type", "UNKNOWN")
            type_counts[ftype] = type_counts.get(ftype, 0) + 1

        st.markdown("**Formula Type Breakdown**")
        cols = st.columns(max(len(type_counts), 1))
        for col, (ftype, count) in zip(cols, type_counts.items()):
            col.metric(ftype, count)

    # KPI dependency table
    if deps:
        st.markdown("**KPI Dependencies**")
        rows = []
        for dep in deps:
            rows.append({
                "KPI Label":        getattr(dep, "kpi_label", "—"),
                "Formula Type":     getattr(dep, "formula_type", "—"),
                "Cell":             getattr(dep, "cell_address", "—"),
                "Source Columns":   ", ".join(getattr(dep, "canonical_source_columns", [])[:5]),
                "Resolved":         "Yes" if getattr(dep, "is_resolved", True) else "No",
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # Warn about unresolved SOURCE_BACKED KPIs only if genuinely ambiguous
    unresolved = [
        dep for dep in deps
        if not getattr(dep, "is_resolved", True)
        and getattr(dep, "formula_type", "") == "SOURCE_BACKED"
    ]
    if unresolved:
        st.warning(
            f"{len(unresolved)} SOURCE_BACKED KPI formula(s) could not be automatically "
            "resolved to source columns. Check the KPI Unresolved cards below."
        )
        class _FakePipeline:
            def __init__(self, results):
                self.agent_results = results
            def get_result(self, name):
                return next((r for r in self.agent_results if r.agent_name == name), None)

        schema_result = st.session_state.get(_SS_SCHEMA)
        fake_results = [kpi_result_obj]
        if schema_result:
            fake_results.append(schema_result)
        fake_pipeline = _FakePipeline(fake_results)
        actionable = classify_actionable_decisions(fake_pipeline)
        kpi_items = [a for a in actionable if a.action_type == "KPI_UNRESOLVED"]
        for i, ad in enumerate(kpi_items):
            _render_kpi_unresolved_card(ad, i)

    with st.expander("Agent Details", expanded=False):
        _render_agent_card(kpi_result_obj)

    if step == 6:
        if st.button("✅ Confirm KPI Analysis & Continue", type="primary", key="ar_btn_confirm_kpi"):
            _set_step(7)
            st.rerun()


def _render_step_7() -> None:
    step = _current_step()
    if step != 7:
        return

    st.subheader("Step 7 — Validation")

    bundles        = st.session_state.get(_SS_BUNDLES, [])
    config         = st.session_state.get(_SS_CONFIG)
    kpi_result_obj = st.session_state.get(_SS_KPI)
    kpi_analysis   = kpi_result_obj.output if kpi_result_obj else None
    gen_result     = st.session_state.get(_SS_GENERATION)
    val_result: AgentResult | None = st.session_state.get(_SS_VALIDATION)

    # Defensive validation
    if val_result is None and config is not None:
        _missing = _validate_pipeline_inputs(
            "Step 7 — Validation",
            {
                "RationalizationConfig": (
                    config,
                    ["future_source_tab_name", "workbook_configs"],
                ),
            },
        )
        if _missing:
            _render_validation_error("Step 7 — Validation", _missing)
            return

    if val_result is None:
        # Need master_df from generation output
        master_df = None
        resolved_frames = None
        if gen_result and isinstance(gen_result.output, tuple):
            fs_bytes = gen_result.output[0] if len(gen_result.output) > 0 else None
            resolved_frames = gen_result.output[2] if len(gen_result.output) > 2 else []
            if fs_bytes and config:
                try:
                    master_tab = config.future_source_tab_name
                    master_df = pd.read_excel(io.BytesIO(fs_bytes), sheet_name=master_tab, header=1)
                except Exception as exc:
                    st.warning(f"Could not extract Master Source Data for validation: {exc}")

        with st.spinner("Running Validation Agent…"):
            try:
                val_result = run_validation_phase(
                    bundles, config, kpi_analysis, master_df,
                    resolved_source_frames=resolved_frames,
                )
                st.session_state[_SS_VALIDATION] = val_result
            except Exception as exc:
                st.error(f"Validation failed: `{type(exc).__name__}: {exc}`")
                st.code(traceback.format_exc(), language="python")
                return

    if val_result.warnings:
        for w in val_result.warnings:
            st.info(w)

    if not val_result.decisions:
        st.info(
            "Validation skipped: No numeric KPI-referenced columns found in Master Source Data. "
            "This typically occurs when source tabs contain non-numeric data or when KPI formulas "
            "reference columns that could not be resolved. "
            "To resolve: verify the source tab contains numeric data and KPI formulas reference "
            "the correct column names."
        )
        with st.expander("Agent Details", expanded=False):
            _render_agent_card(val_result)
        if step == 7:
            if st.button("✅ Confirm Validation & Continue", type="primary", key="ar_btn_confirm_val"):
                _set_step(8)
                st.rerun()
        return

    overall = next((d for d in val_result.decisions if d.subject == "Overall validation"), None)
    if overall:
        sig    = overall.signals
        total  = sig.get("total_columns", 0)
        passed = sig.get("pass_count", 0)
        warned = sig.get("warn_count", 0)
        failed = sig.get("fail_count", 0)

        if failed == 0 and warned == 0:
            st.success(f"All {total} KPI column(s) reconciled successfully within tolerance.")
        elif failed > 0:
            st.error(f"{failed} column(s) FAILED reconciliation — source totals do not match Master.")
        else:
            st.warning(f"{warned} column(s) have reconciliation warnings.")

        c1, c2, c3 = st.columns(3)
        c1.metric("✅ Pass", passed)
        c2.metric("⚠️ Warn", warned)
        c3.metric("❌ Fail", failed)

    detail_rows = []
    for d in val_result.decisions:
        if d.subject == "Overall validation":
            continue
        s = d.signals
        status_icon = "✅" if d.decision == "PASS" else ("⚠️" if d.decision == "WARN" else "❌")
        detail_rows.append({
            "Status":       f"{status_icon} {d.decision}",
            "Column":       d.subject.replace("Reconciliation: '", "").rstrip("'"),
            "Source Total": f"{s.get('source_total', 0):,.2f}",
            "Master Total": f"{s.get('master_total', 0):,.2f}",
            "Delta":        f"{s.get('delta', 0):+,.2f}",
            "Variance %":   f"{s.get('delta_pct', 0):.4f}%",
            "Tolerance":    f"{s.get('tolerance_pct', 0):.2f}%",
        })
    if detail_rows:
        st.dataframe(pd.DataFrame(detail_rows), use_container_width=True, hide_index=True)

    with st.expander("Agent Details", expanded=False):
        _render_agent_card(val_result)

    if step == 7:
        if st.button("✅ Confirm Validation & Continue", type="primary", key="ar_btn_confirm_val"):
            _set_step(8)
            st.rerun()


def _render_step_8() -> None:
    step = _current_step()
    if step != 8:
        return

    st.subheader("Step 8 — Generate & Download")

    bundles       = st.session_state.get(_SS_BUNDLES, [])
    config        = st.session_state.get(_SS_CONFIG)
    schema_result = st.session_state.get(_SS_SCHEMA)
    kpi_result_obj = st.session_state.get(_SS_KPI)
    kpi_analysis  = kpi_result_obj.output if kpi_result_obj else None
    source_analysis = schema_result.output if schema_result else None
    gen_result: AgentResult | None = st.session_state.get(_SS_GENERATION)

    if gen_result is None:
        if source_analysis is None:
            st.error("Schema analysis is required before generation. Complete Step 5 first.")
            return
        # Defensive validation
        _missing = _validate_pipeline_inputs(
            "Step 8 — Generate",
            {
                "SourceAnalysisResult": (
                    source_analysis,
                    ["column_profiles", "column_mapping", "excluded_columns"],
                ),
                "RationalizationConfig": (
                    config,
                    ["future_source_tab_name", "workbook_configs"],
                ),
            },
        )
        if _missing:
            _render_validation_error("Step 8 — Generate", _missing)
            return
        with st.spinner("Running Generation Agent…"):
            try:
                gen_result = run_generation_phase(bundles, config, source_analysis, kpi_analysis)
                st.session_state[_SS_GENERATION] = gen_result

                # Trigger validation now that we have the workbook
                val_result = run_validation_phase(
                    bundles, config, kpi_analysis,
                    master_df=_extract_master_df(gen_result, config),
                    resolved_source_frames=_extract_resolved_frames(gen_result),
                )
                st.session_state[_SS_VALIDATION] = val_result
            except Exception as exc:
                st.error(f"Generation failed: `{type(exc).__name__}: {exc}`")
                st.code(traceback.format_exc(), language="python")
                return

    if gen_result.warnings:
        for w in gen_result.warnings:
            st.warning(w)

    gen_out = gen_result.output if isinstance(gen_result.output, tuple) else (None, None, [])
    fs_bytes  = gen_out[0] if len(gen_out) > 0 else None
    ap_bytes  = gen_out[1] if len(gen_out) > 1 else None

    if not fs_bytes and not ap_bytes:
        st.warning(
            "No output workbooks were generated. "
            "Check the agent recommendations above for blocking issues."
        )
        with st.expander("Agent Details", expanded=False):
            _render_agent_card(gen_result)
        return

    st.success("Workbooks generated successfully.")

    col1, col2, col3, col4 = st.columns(4)

    if fs_bytes:
        col1.download_button(
            label="📥 Future State Workbook",
            data=fs_bytes,
            file_name="Future_State_Workbook.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if ap_bytes:
        col2.download_button(
            label="📥 Analysis Pack",
            data=ap_bytes,
            file_name="Rationalization_Analysis_Pack.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # Decision log from all available agent results
    all_agent_results = [
        r for r in [
            st.session_state.get(_SS_DISCOVERY),
            st.session_state.get(_SS_CONSOLIDATION),
            st.session_state.get(_SS_KPI),
            st.session_state.get(_SS_SCHEMA),
            gen_result,
            st.session_state.get(_SS_VALIDATION),
        ]
        if r is not None
    ]
    log_bytes = None
    try:
        from src.agentic_rationalization import decision_log
        log_bytes = decision_log.build(all_agent_results)
    except Exception:
        pass

    if log_bytes:
        col3.download_button(
            label="📥 Agent Decision Log",
            data=log_bytes,
            file_name="Agent_Decision_Log.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if sum(1 for x in [fs_bytes, ap_bytes, log_bytes] if x) > 1:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            if fs_bytes:
                zf.writestr("Future_State_Workbook.xlsx", fs_bytes)
            if ap_bytes:
                zf.writestr("Rationalization_Analysis_Pack.xlsx", ap_bytes)
            if log_bytes:
                zf.writestr("Agent_Decision_Log.xlsx", log_bytes)
        col4.download_button(
            label="📦 Download All (ZIP)",
            data=zip_buf.getvalue(),
            file_name=f"Agentic_Rationalization_{datetime.now():%Y%m%d_%H%M}.zip",
            mime="application/zip",
        )

    with st.expander("Agent Details", expanded=False):
        _render_agent_card(gen_result)

    st.markdown("---")
    if st.button("🔄 Start New Analysis", key="ar_btn_reset"):
        _reset_state()
        st.rerun()


# ── Helpers for generation step ───────────────────────────────────────────────

def _extract_master_df(gen_result: AgentResult, config) -> "pd.DataFrame | None":
    if gen_result is None or not isinstance(gen_result.output, tuple):
        return None
    fs_bytes = gen_result.output[0] if len(gen_result.output) > 0 else None
    if not fs_bytes or config is None:
        return None
    try:
        master_tab = config.future_source_tab_name
        return pd.read_excel(io.BytesIO(fs_bytes), sheet_name=master_tab, header=1)
    except Exception:
        return None


def _extract_resolved_frames(gen_result: AgentResult):
    if gen_result is None or not isinstance(gen_result.output, tuple):
        return []
    return gen_result.output[2] if len(gen_result.output) > 2 else []


# ── Pipeline input validation ────────────────────────────────────────────────

def _validate_pipeline_inputs(
    step_name: str,
    required: "dict[str, tuple[Any, list[str]]]",
) -> "list[str]":
    """Check that all required objects exist and expose the expected attributes.

    Args:
        step_name: Human-readable name of the step (used in error messages).
        required: Mapping of ``description → (obj, [attr, ...])`` pairs.
            *description* is shown to the user when an object is missing.
            *obj* is the actual object to inspect (may be None).
            *[attr, ...]* is the list of attribute names that must exist on *obj*.

    Returns:
        A list of human-readable missing-attribute descriptions.  Empty when
        all checks pass.
    """
    missing: list[str] = []
    for description, (obj, attrs) in required.items():
        if obj is None:
            missing.append(f"{description} is None (not yet produced)")
            continue
        for attr in attrs:
            if not hasattr(obj, attr):
                actual = [a for a in dir(obj) if not a.startswith("_")]
                missing.append(
                    f"{description}.{attr} — "
                    f"actual attributes: {actual}"
                )
    return missing


def _render_validation_error(step_name: str, missing: "list[str]") -> None:
    """Render a styled error panel listing missing pipeline inputs."""
    lines = [f"⚠️ {step_name} cannot start.", ""]
    for m in missing:
        lines.append(f"Missing: {m}")
    st.error("\n".join(lines))


# ── Main page ─────────────────────────────────────────────────────────────────

def render() -> None:
    _render_sidebar_nav()

    _render_step_1()
    _render_step_2()
    _render_step_3()
    _render_step_4()
    _render_step_5()
    _render_step_6()
    _render_step_7()
    _render_step_8()
