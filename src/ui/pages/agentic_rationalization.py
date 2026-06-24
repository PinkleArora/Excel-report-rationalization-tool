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
from src.ingestion.workbook_analyzer import scan_kpi_blocks, _get_raw_ws
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

    if st.button("▶ Run Discovery", type="primary", key="ar_btn_discovery"):
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

    st.subheader("Step 2 — Discovery")

    bundles = st.session_state.get(_SS_BUNDLES, [])
    if not bundles:
        st.warning("No bundles loaded. Return to Step 1.")
        return

    # ── Phase 1: workbook structural analysis (cached — only runs once per upload) ──
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

    # ── Phase 2: build RationalizationConfig from cached analyses (fast) ─────
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

    # ── Workbook selector ─────────────────────────────────────────────────────
    wb_names = [b.file_name for b in bundles]
    selected_wb = st.selectbox(
        "Workbook",
        wb_names,
        key="ar_disc_selected_wb",
        label_visibility="visible",
    )
    bundle = next(b for b in bundles if b.file_name == selected_wb)
    wb_cfg = config.config_for(selected_wb)
    source_tab = (st.session_state.get(_SS_SOURCE_TAB_OVERRIDES, {}).get(selected_wb)
                  or (wb_cfg.source_tab if wb_cfg else ""))
    kpi_tabs   = (st.session_state.get("ar_kpi_tab_overrides", {}).get(selected_wb)
                  or (wb_cfg.kpi_tabs if wb_cfg else []))
    all_sheets = list(bundle.sheets.keys())

    # ── Pending tag edits — stored per workbook ───────────────────────────────
    # ar_disc_tags: {wb_name: {sheet_name: "Source"|"KPI"|"Ignore"}}
    _SS_DISC_TAGS = "ar_disc_tags"
    all_tags: dict[str, dict[str, str]] = st.session_state.get(_SS_DISC_TAGS, {})
    # Initialise tags for this workbook if not set
    if selected_wb not in all_tags:
        all_tags[selected_wb] = {
            s: _infer_tab_tag(s, source_tab, kpi_tabs) for s in all_sheets
        }
        st.session_state[_SS_DISC_TAGS] = all_tags
    current_tags: dict[str, str] = all_tags[selected_wb]

    # Committed tags (what discovery was actually run with)
    committed_tags: dict[str, str] = {
        s: _infer_tab_tag(s, source_tab, kpi_tabs) for s in all_sheets
    }
    has_pending = current_tags != committed_tags

    # ── Tabs table ────────────────────────────────────────────────────────────
    st.markdown("**Tabs**")
    _TAG_OPTIONS = ["Source", "KPI", "Ignore"]
    col_tab, col_tag, col_struct, col_rows, col_cols = st.columns([3, 2, 2, 1, 1])
    col_tab.markdown("**Tab Name**")
    col_tag.markdown("**Tag**")
    col_struct.markdown("**Structure**")
    col_rows.markdown("**Rows**")
    col_cols.markdown("**Cols**")

    for sheet_name in all_sheets:
        df_sheet = bundle.sheets.get(sheet_name)
        n_rows = len(df_sheet) if df_sheet is not None else 0
        n_cols = len(df_sheet.columns) if df_sheet is not None else 0
        struct = _tab_structure_label(bundle, sheet_name, discovery_result)

        c_tab, c_tag, c_struct, c_rows, c_cols = st.columns([3, 2, 2, 1, 1])
        with c_tab:
            # Clicking the tab name sets it as the column-explorer target
            if st.button(sheet_name, key=f"ar_disc_sel_{selected_wb}_{sheet_name}", use_container_width=True):
                st.session_state["ar_disc_explorer_tab"] = (selected_wb, sheet_name)
        with c_tag:
            current_tag = current_tags.get(sheet_name, "Ignore")
            new_tag = st.selectbox(
                "tag",
                _TAG_OPTIONS,
                index=_TAG_OPTIONS.index(current_tag),
                key=f"ar_disc_tag_{selected_wb}_{sheet_name}",
                label_visibility="collapsed",
            )
            if new_tag != current_tag:
                current_tags[sheet_name] = new_tag
                all_tags[selected_wb] = current_tags
                st.session_state[_SS_DISC_TAGS] = all_tags
                has_pending = True
                st.rerun()
        c_struct.markdown(struct)
        c_rows.markdown(str(n_rows))
        c_cols.markdown(str(n_cols))

    # ── Column Explorer ───────────────────────────────────────────────────────
    explorer_target = st.session_state.get("ar_disc_explorer_tab")
    # Default to source tab on first render
    if explorer_target is None or explorer_target[0] != selected_wb:
        if source_tab in all_sheets:
            explorer_target = (selected_wb, source_tab)
            st.session_state["ar_disc_explorer_tab"] = explorer_target

    if explorer_target and explorer_target[0] == selected_wb:
        explorer_sheet = explorer_target[1]
        df_exp = bundle.sheets.get(explorer_sheet)
        if df_exp is not None:
            st.markdown("---")
            st.markdown(f"**Selected Tab: `{explorer_sheet}`**")
            explorer_tag = current_tags.get(explorer_sheet, "Ignore")

            # For KPI tabs, attempt block-level structure detection (result cached)
            kpi_blocks = []
            if explorer_tag == "KPI":
                blocks_cache: dict = st.session_state.setdefault(_SS_BLOCKS_CACHE, {})
                cache_key = (selected_wb, explorer_sheet)
                if cache_key not in blocks_cache:
                    # Prefer pivot metadata when available — it gives authoritative
                    # Row/Col/Filter = dimensions, Data = KPI measures classification.
                    wb_analyses = st.session_state.get(_SS_WB_ANALYSES) or []
                    tab_analysis = next(
                        (
                            t
                            for a in wb_analyses
                            if a.workbook_name == selected_wb
                            for t in a.tab_analyses
                            if t.tab_name == explorer_sheet
                        ),
                        None,
                    )
                    if tab_analysis is not None and tab_analysis.is_pivot_sheet and tab_analysis.pivot_value_fields:
                        from src.ingestion.workbook_analyzer import KpiBlock, KpiMeasure  # noqa: F401
                        all_dim_fields = (
                            tab_analysis.pivot_filter_fields
                            + tab_analysis.pivot_row_fields
                            + tab_analysis.pivot_col_fields
                        )
                        # Deduplicate while preserving order
                        seen: set[str] = set()
                        dims: list[str] = []
                        for f in all_dim_fields:
                            if f not in seen:
                                seen.add(f)
                                dims.append(f)
                        blocks_cache[cache_key] = [
                            KpiBlock(
                                block_name=explorer_sheet,
                                dimensions=dims,
                                kpi_measures=tab_analysis.pivot_value_fields,
                                header_row=1,
                                data_start_row=2,
                                data_end_row=len(df_exp) + 1,
                            )
                        ]
                    else:
                        raw_wb = getattr(bundle, "_raw_wb", None)
                        raw_ws = _get_raw_ws(raw_wb, explorer_sheet)
                        blocks_cache[cache_key] = scan_kpi_blocks(raw_ws) if raw_ws is not None else []
                kpi_blocks = blocks_cache[cache_key]

            struct_label = _tab_structure_label(bundle, explorer_sheet, discovery_result)
            if kpi_blocks:
                struct_label = "Multi-Block KPI"

            ec1, ec2, ec3 = st.columns(3)
            ec1.metric("Rows", f"{len(df_exp):,}")
            if kpi_blocks:
                total_measures = sum(len(b.kpi_measures) for b in kpi_blocks)
                ec2.metric("KPI Blocks", len(kpi_blocks))
                ec3.metric("KPI Measures", total_measures)
            else:
                ec2.metric("Columns", len(df_exp.columns))
                ec3.metric("Structure", struct_label)

            if kpi_blocks:
                # ── Block-level KPI structure display ────────────────────────
                st.caption(f"{len(kpi_blocks)} KPI block(s) detected — expand to see dimensions and measures.")
                for blk in kpi_blocks:
                    n_kpi = len(blk.kpi_measures)
                    n_dim = len(blk.dimensions)
                    with st.expander(f"{blk.block_name}  ({n_kpi} KPI{'s' if n_kpi != 1 else ''})", expanded=False):
                        if blk.dimensions:
                            st.markdown("**Dimensions**")
                            for dim in blk.dimensions:
                                st.markdown(f"- {dim}")
                        st.markdown("**KPI Measures**")
                        for measure in blk.kpi_measures:
                            agg = getattr(measure, "aggregation", "")
                            agg_tag = f" `{agg}`" if agg else ""
                            st.markdown(f"- {measure.display_name}{agg_tag}")
            else:
                # ── Flat column list (source tab or KPI tab with no blocks) ──
                col_rows_data = []
                for col in df_exp.columns:
                    col_str = str(col)
                    # Skip unnamed columns for KPI tabs
                    if explorer_tag == "KPI" and col_str.startswith("Unnamed:"):
                        continue
                    dtype = df_exp[col].dtype
                    if str(dtype).startswith("int") or str(dtype).startswith("float"):
                        dtype_label = "Numeric"
                    elif str(dtype) in ("object", "string"):
                        col_lower = col_str.lower()
                        if any(kw in col_lower for kw in ("date", "period", "month", "year", "day")):
                            dtype_label = "Date"
                        else:
                            dtype_label = "Text"
                    elif "datetime" in str(dtype):
                        dtype_label = "Date"
                    elif "bool" in str(dtype):
                        dtype_label = "Boolean"
                    else:
                        dtype_label = str(dtype)
                    col_rows_data.append({"Column Name": col_str, "Data Type": dtype_label})
                if col_rows_data:
                    st.dataframe(
                        pd.DataFrame(col_rows_data),
                        use_container_width=True,
                        hide_index=True,
                    )
                else:
                    st.info("No readable columns found in this tab.")

    # ── Action buttons ────────────────────────────────────────────────────────
    st.markdown("---")
    btn_c1, btn_c2 = st.columns([1, 1])

    with btn_c1:
        if st.button(
            "🔄 Update Discovery",
            disabled=not has_pending,
            key="ar_btn_update_discovery",
        ):
            # Apply pending tags as overrides and re-run discovery
            new_src_overrides: dict[str, str] = {}
            new_kpi_overrides: dict[str, list[str]] = {}
            for wb_n, tag_map in all_tags.items():
                src = next((s for s, t in tag_map.items() if t == "Source"), "")
                kpis = [s for s, t in tag_map.items() if t == "KPI"]
                if src:
                    new_src_overrides[wb_n] = src
                if kpis:
                    new_kpi_overrides[wb_n] = kpis
            st.session_state[_SS_SOURCE_TAB_OVERRIDES] = new_src_overrides
            st.session_state["ar_kpi_tab_overrides"] = new_kpi_overrides
            # Clear discovery + downstream; keep _SS_WB_ANALYSES so re-run is fast
            st.session_state.pop(_SS_DISCOVERY, None)
            st.session_state.pop(_SS_CONSOLIDATION, None)
            st.session_state.pop(_SS_SCHEMA, None)
            st.session_state.pop(_SS_BLOCKS_CACHE, None)
            st.session_state.pop(_SS_STRUCT_CACHE, None)
            st.session_state.pop(_SS_DISC_TAGS, None)
            st.rerun()

    with btn_c2:
        if step == 2:
            if st.button(
                "✅ Continue",
                type="primary",
                disabled=has_pending,
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


# ── Consolidation workspace helpers ──────────────────────────────────────────

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


def _render_step_3() -> None:
    from src.agentic_rationalization.agents.consolidation_agent import compute_file_group_compatibility

    step = _current_step()
    if step != 3:
        return

    st.subheader("Step 3 — Consolidation Workspace")

    config = st.session_state.get(_SS_CONFIG)
    consol_result: AgentResult | None = st.session_state.get(_SS_CONSOLIDATION)

    if consol_result is None or consol_result.output is None:
        st.warning("Consolidation analysis not yet available.")
        return

    intel = consol_result.output.get("intelligence")
    if intel is None:
        st.info("No intelligence output available.")
        return

    # Defensive validation
    _missing = _validate_pipeline_inputs(
        "Step 3 — Consolidation Workspace",
        {
            "ConsolidationIntelligenceResult": (
                intel,
                ["file_profiles", "groups", "pairwise_scores", "final_recommendation"],
            ),
        },
    )
    if _missing:
        _render_validation_error("Step 3 — Consolidation Workspace", _missing)
        return

    group_overrides: dict = st.session_state.get(_SS_GROUP_OVERRIDES, {})
    effective_groups = _apply_group_overrides(intel.groups, group_overrides)
    file_names = [fp.file_name for fp in intel.file_profiles]

    (tab_profiles, tab_kpi_disc, tab_groups, tab_pairwise,
     tab_kpi_align, tab_overrides) = st.tabs([
        "File Profiles",
        "KPI Discovery",
        "Consolidation Groups",
        "5-Dimension Scores",
        "KPI Alignment",
        "Manual Overrides",
    ])

    # ── Tab 1: File Profiles ──────────────────────────────────────────────────
    with tab_profiles:
        if intel.file_profiles:
            # Build per-workbook KPI tab list from workbook_kpi_stats
            kpi_tabs_by_file: dict[str, str] = {}
            for wks in getattr(intel, "workbook_kpi_stats", []):
                kpi_tabs_by_file[wks.workbook_name] = ", ".join(wks.kpi_tabs) if wks.kpi_tabs else "—"

            rows = []
            for fp in intel.file_profiles:
                pk_str = ", ".join(fp.primary_key_candidates) if fp.primary_key_candidates else "—"
                config_ss = st.session_state.get(_SS_CONFIG)
                kpi_tabs_str = "—"
                if config_ss:
                    wb_cfg_fp = config_ss.config_for(fp.file_name)
                    if wb_cfg_fp and hasattr(wb_cfg_fp, "kpi_tabs") and wb_cfg_fp.kpi_tabs:
                        kpi_tabs_str = ", ".join(wb_cfg_fp.kpi_tabs)
                rows.append({
                    "File":              fp.file_name,
                    "Source Tab":        fp.source_tab,
                    "KPI Tabs":          kpi_tabs_by_file.get(fp.file_name, "—"),
                    "Rows":              fp.row_count,
                    "Total Columns":     fp.column_count,
                    "Distinct Columns":  fp.distinct_columns,
                    "KPI Columns":       fp.kpi_referenced_column_count,
                    "Non-KPI Columns":   fp.non_kpi_column_count,
                    "Inferred Grain":    fp.inferred_grain,
                    "Grain Confidence":  f"{fp.grain_confidence:.0%}",
                    "Primary Key(s)":    pk_str,
                    "Time Dimension":    "Yes" if fp.has_time_dimension else "No",
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        else:
            st.info("No file profiles available.")

    # ── Tab 2: KPI Discovery ─────────────────────────────────────────────────
    with tab_kpi_disc:
        wb_kpi_stats = getattr(intel, "workbook_kpi_stats", [])
        if not wb_kpi_stats:
            st.info("KPI Discovery data not available — re-run from Step 2.")
        else:
            for wks in wb_kpi_stats:
                conf_icon = _conf_color(wks.detection_confidence)
                dt = getattr(wks, "detection_type", "Formula-Based")
                dt_icon = {"Formula-Based": "🧮", "Pivot-Based": "🔄", "Mixed": "⚡"}.get(dt, "🧮")
                with st.expander(
                    f"{conf_icon} **{wks.workbook_name}** — "
                    f"{wks.kpi_column_count} KPI column(s), "
                    f"{wks.formula_count} formula(s)  {dt_icon} {dt}",
                    expanded=True,
                ):
                    c1, c2, c3, c4, c5 = st.columns(5)
                    c1.metric("KPI Tabs Found", len(wks.kpi_tabs))
                    c2.metric("KPI Formulas", wks.formula_count)
                    c3.metric("KPI Columns", wks.kpi_column_count)
                    c4.metric("Detection Confidence", f"{wks.detection_confidence:.0%}")
                    c5.metric("Detection Type", dt)
                    if wks.kpi_tabs:
                        st.markdown(
                            "**KPI Tabs Identified:** " +
                            " · ".join(f"`{t}`" for t in wks.kpi_tabs)
                        )
                    if wks.reason_if_empty:
                        st.warning(f"**Why no KPIs detected:** {wks.reason_if_empty}")
                        st.markdown(
                            "**Troubleshooting:**\n"
                            "1. Verify the correct KPI tab is selected in Step 1.\n"
                            "2. Check that KPI formulas reference source-data columns "
                            "(e.g. `=SUMIFS(SQL_data!D:D, SQL_data!B:B, A1)`).\n"
                            "3. Formulas that only reference other KPI-tab cells are "
                            "classified as DERIVED and may not resolve to source columns."
                        )

        # KPI Lineage
        kpi_result_obj_ln = st.session_state.get(_SS_KPI)
        kpi_analysis_ln = kpi_result_obj_ln.output if kpi_result_obj_ln else None
        lineage = getattr(kpi_analysis_ln, "lineage", []) if kpi_analysis_ln else []
        if lineage:
            st.markdown("**KPI Lineage**")
            lineage_by_wb: dict[str, list] = {}
            for ln in lineage:
                lineage_by_wb.setdefault(ln.workbook_name, []).append(ln)
            for wb_name, wb_lineage in lineage_by_wb.items():
                with st.expander(
                    f"🔗 {wb_name} — {len(wb_lineage)} KPI lineage path(s)", expanded=False
                ):
                    unique_paths = list({ln.lineage_path for ln in wb_lineage})
                    for path in sorted(unique_paths):
                        st.code(path, language=None)
                    type_counts: dict[str, int] = {}
                    for ln in wb_lineage:
                        type_counts[ln.kpi_type] = type_counts.get(ln.kpi_type, 0) + 1
                    for ktype, cnt in type_counts.items():
                        st.caption(f"{ktype}: {cnt}")

    # ── Tab 3: Consolidation Groups ───────────────────────────────────────────
    with tab_groups:
        if not effective_groups:
            st.info("No groups detected.")
        else:
            for group in effective_groups:
                if group.recommendation == "Consolidate":
                    icon, rec_tag = "🟢", "[Consolidate ✓]"
                else:
                    icon, rec_tag = "🟡", "[Keep Separate]"
                with st.expander(
                    f"{icon} Group {group.group_id} — {group.group_label}  {rec_tag}",
                    expanded=True,
                ):
                    st.caption(group.reasoning)
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Common Columns", group.common_column_count)
                    c2.metric("Total Union Cols", group.unique_column_count)
                    c3.metric("KPI Required", group.kpi_required_column_count)
                    c4.metric("Discardable", group.discardable_column_count)

                    contributions = getattr(group, "file_contributions", [])
                    if contributions:
                        st.markdown("**Per-File Contribution Breakdown**")
                        contrib_rows = []
                        for fc in contributions:
                            contrib_rows.append({
                                "File":                   fc.file_name,
                                "Total Cols":             fc.total_columns,
                                "KPI Cols":               fc.kpi_columns,
                                "Key/Grain Cols":         fc.key_columns,
                                "Common Contributed":     fc.common_columns_contributed,
                                "Unique Contributed":     fc.unique_columns_contributed,
                                "Discardable":            fc.discardable_columns,
                            })
                        st.dataframe(
                            pd.DataFrame(contrib_rows),
                            use_container_width=True, hide_index=True,
                        )
                    else:
                        st.markdown("**Files in Group:**")
                        for fname in group.file_names:
                            st.markdown(f"- `{fname}`")

                    detail = getattr(group, "incompatibility_detail", "")
                    if group.recommendation == "Keep Separate" and detail:
                        with st.expander("Why can't this file be combined?", expanded=False):
                            st.markdown(detail)

        # ── KPI Redundancy Analysis ───────────────────────────────────────────
        kpi_redundancy = getattr(intel, "kpi_redundancy", [])
        if kpi_redundancy:
            st.markdown("---")
            st.markdown("### KPI Redundancy Analysis")
            st.caption(
                "Redundancy is determined **only by KPI coverage**. "
                "A file is redundant when every KPI-referenced column it contains "
                "is already present in at least one other file. "
                "Non-KPI columns (dimensions, metadata, unused attributes) are ignored."
            )

            n_redundant = sum(1 for r in kpi_redundancy if r.redundant)
            n_partial   = sum(1 for r in kpi_redundancy if r.recommendation.startswith("Partial"))
            n_retain    = sum(1 for r in kpi_redundancy if r.recommendation == "Retain")

            rc1, rc2, rc3 = st.columns(3)
            rc1.metric("🗑 Redundant (safe to discard)", n_redundant)
            rc2.metric("⚠️ Partial — review required",  n_partial)
            rc3.metric("✅ Retain (unique KPI coverage)", n_retain)

            for r in kpi_redundancy:
                if r.redundant:
                    icon, border_color = "🗑", "red"
                elif r.recommendation.startswith("Partial"):
                    icon, border_color = "⚠️", "orange"
                else:
                    icon, border_color = "✅", "green"

                with st.expander(
                    f"{icon} **{r.file_name}** — {r.recommendation}  "
                    f"({r.total_kpi_columns} KPI cols, {r.unique_kpi_columns} unique)",
                    expanded=r.redundant,
                ):
                    col_a, col_b, col_c = st.columns(3)
                    col_a.metric("Total KPI Columns", r.total_kpi_columns)
                    col_b.metric("Unique KPI Columns", r.unique_kpi_columns,
                                 delta=None,
                                 help="KPI columns not present in any other file")
                    col_c.metric("Covered By", len(r.covered_by))

                    if r.recommendation == "Discard":
                        st.error(f"**Recommendation: Discard**")
                    elif r.recommendation.startswith("Partial"):
                        st.warning(f"**Recommendation: {r.recommendation}**")
                    else:
                        st.success(f"**Recommendation: Retain**")

                    st.markdown(r.reason)

                    if r.covered_by:
                        st.markdown(
                            "**Covered by:** " + ", ".join(f"`{f}`" for f in r.covered_by)
                        )

        fr = intel.final_recommendation
        if fr.consolidate_files:
            st.success(f"**{fr.summary}**  {fr.business_reasoning}")
        else:
            st.info(f"**{fr.summary}**  {fr.business_reasoning}")

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

    # ── Tab 4: 5-Dimension Scores ─────────────────────────────────────────────
    with tab_pairwise:
        if not intel.pairwise_scores:
            st.info("No pairwise scores (single file or no comparison possible).")
        else:
            for sc in intel.pairwise_scores:
                with st.expander(
                    f"`{sc.file_a}` vs `{sc.file_b}` — {sc.recommendation}",
                    expanded=False,
                ):
                    g1, g2, g3, g4, g5, g6 = st.columns(6)
                    g1.metric("Grain Compatibility", "✓" if sc.grain_compatible else "✗ Mismatch")
                    g2.metric("Schema Similarity", f"{sc.schema_overlap_pct:.0%}")
                    g3.metric("KPI Similarity", f"{sc.kpi_overlap_pct:.0%}")
                    g4.metric("Key Compatible", "Yes" if sc.key_compatible else "No")
                    g5.metric("Time Dimension", "Match" if sc.time_dimension_compatible else "Mismatch")
                    g6.metric("Overall Score", f"{sc.overall_score:.0%}")

                    rec_color = (
                        "green" if sc.recommendation == "Consolidate"
                        else ("orange" if sc.recommendation == "Conditionally Consolidate" else "red")
                    )
                    st.markdown(
                        f"**Recommendation:** :{rec_color}[{sc.recommendation}]  "
                        f"(confidence: {sc.recommendation_confidence:.0%})"
                    )
                    st.caption(sc.reasoning)

                    if not sc.grain_compatible:
                        st.markdown(
                            f"> **Grain mismatch**: `{sc.file_a}` is *{sc.grain_a}*-level while "
                            f"`{sc.file_b}` is *{sc.grain_b}*-level. "
                            f"{sc.grain_a.title()}-level records cannot be directly merged with "
                            f"{sc.grain_b.title()}-level records without aggregation — doing so "
                            f"would produce incorrect totals."
                        )
                    elif sc.overall_score < 0.50:
                        st.markdown(
                            f"> **Low compatibility** (score {sc.overall_score:.0%}): Files are "
                            f"grain-compatible but have insufficient KPI/schema overlap to justify "
                            f"consolidation. They likely support different reporting objectives."
                        )

    # ── Tab 5: KPI Alignment ─────────────────────────────────────────────────
    with tab_kpi_align:
        kpi_alignments = getattr(intel, "kpi_alignments", [])
        kpi_decisions: dict = st.session_state.get(_SS_KPI_ALIGN_DECISIONS, {})

        if not kpi_alignments:
            st.warning(
                "No KPI columns identified. This usually means KPI analysis was not run before "
                "consolidation, or no KPI formulas reference source-data columns."
            )
            st.markdown(
                "**Troubleshooting:**\n"
                "1. Return to Step 1 and verify KPI tabs are correctly identified.\n"
                "2. Check the **KPI Discovery** tab for per-workbook detection details.\n"
                "3. Ensure KPI formulas reference the source tab "
                "(e.g. `=SUMIFS(SQL_data!D:D,SQL_data!B:B,A1)`)."
            )
        else:
            partial = [a for a in kpi_alignments if a.status == "partial"]
            aligned = [a for a in kpi_alignments if a.status == "aligned"]
            missing = [a for a in kpi_alignments if a.status == "missing"]
            all_canonicals = [a.canonical_kpi for a in kpi_alignments]

            # ── Summary counts ────────────────────────────────────────────────
            n_accepted = sum(1 for c in all_canonicals if kpi_decisions.get(c, "pending") == "accepted")
            n_rejected = sum(1 for c in all_canonicals if kpi_decisions.get(c, "pending") == "rejected")
            n_pending  = sum(1 for c in all_canonicals if kpi_decisions.get(c, "pending") not in ("accepted", "rejected"))

            m1, m2, m3, m4, m5, m6 = st.columns(6)
            m1.metric("Fully Aligned",      len(aligned), help="KPI column present in all files")
            m2.metric("Partially Aligned",  len(partial), help="KPI column present in some files only")
            m3.metric("Missing",            len(missing), help="KPI column not found in any file")
            m4.metric("✅ Accepted",         n_accepted)
            m5.metric("❌ Rejected",         n_rejected)
            m6.metric("❓ Pending Review",   n_pending)

            # ── Global bulk actions ───────────────────────────────────────────
            st.markdown("---")
            st.markdown("**Bulk Actions**")
            gc1, gc2, gc3 = st.columns(3)
            with gc1:
                if st.button("✅ Accept All Mappings", key="kpi_bulk_accept_all"):
                    new = dict(kpi_decisions)
                    for c in all_canonicals:
                        new[c] = "accepted"
                    st.session_state[_SS_KPI_ALIGN_DECISIONS] = new
                    st.rerun()
            with gc2:
                if st.button("❌ Reject All Mappings", key="kpi_bulk_reject_all"):
                    new = dict(kpi_decisions)
                    for c in all_canonicals:
                        new[c] = "rejected"
                    st.session_state[_SS_KPI_ALIGN_DECISIONS] = new
                    st.rerun()
            with gc3:
                if st.button("🔄 Reset All Decisions", key="kpi_bulk_reset_all"):
                    new = {k: v for k, v in kpi_decisions.items() if k not in all_canonicals}
                    st.session_state[_SS_KPI_ALIGN_DECISIONS] = new
                    st.rerun()

            st.markdown("---")

            if partial:
                partial_canonicals = [a.canonical_kpi for a in partial]
                pc1, pc2, pc3, _ = st.columns([1, 1, 1, 3])
                with pc1:
                    if st.button("✅ Accept Section", key="kpi_sec_accept_partial"):
                        new = dict(kpi_decisions)
                        for c in partial_canonicals:
                            new[c] = "accepted"
                        st.session_state[_SS_KPI_ALIGN_DECISIONS] = new
                        st.rerun()
                with pc2:
                    if st.button("❌ Reject Section", key="kpi_sec_reject_partial"):
                        new = dict(kpi_decisions)
                        for c in partial_canonicals:
                            new[c] = "rejected"
                        st.session_state[_SS_KPI_ALIGN_DECISIONS] = new
                        st.rerun()
                with pc3:
                    if st.button("🔄 Reset Section", key="kpi_sec_reset_partial"):
                        new = {k: v for k, v in kpi_decisions.items() if k not in partial_canonicals}
                        st.session_state[_SS_KPI_ALIGN_DECISIONS] = new
                        st.rerun()
                st.markdown("**Partially Aligned KPIs — Review suggested matches**")
                for aln in partial:
                    dec = kpi_decisions.get(aln.canonical_kpi, "pending")
                    status_icon = {"accepted": "✅", "rejected": "❌", "pending": "❓"}.get(dec, "❓")
                    with st.expander(
                        f"{status_icon} `{aln.canonical_kpi}` — {aln.suggestion}",
                        expanded=(dec == "pending"),
                    ):
                        if aln.kpi_labels:
                            st.caption(f"KPI formula label(s): {', '.join(aln.kpi_labels)}")

                        match_rows = []
                        for fn in file_names:
                            raw = aln.per_file_match.get(fn, "")
                            conf = aln.per_file_confidence.get(fn, 0.0)
                            if conf == 1.0:
                                match_type = "Exact"
                            elif conf >= 0.90:
                                match_type = "Column match"
                            elif raw:
                                match_type = f"Fuzzy ({conf:.0%})"
                            else:
                                match_type = "—"
                            rec = "Match" if conf >= 0.90 else ("Possible Match" if conf >= 0.70 else "Review")
                            match_rows.append({
                                "Source Workbook":  fn,
                                "Canonical KPI":    aln.canonical_kpi,
                                "Source KPI Name":  ", ".join(aln.kpi_labels) if aln.kpi_labels else aln.canonical_kpi,
                                "Mapped Column":    raw if raw else "(not found)",
                                "Confidence":       f"{conf:.0%}" if raw else "—",
                                "Match Type":       match_type,
                                "Recommendation":   rec if raw else "No match",
                            })
                        st.dataframe(pd.DataFrame(match_rows), use_container_width=True, hide_index=True)

                        ba, br, bm_col = st.columns([1, 1, 2])
                        with ba:
                            if st.button("✅ Accept Mapping", key=f"kpi_accept_{aln.canonical_kpi}"):
                                kpi_decisions[aln.canonical_kpi] = "accepted"
                                st.session_state[_SS_KPI_ALIGN_DECISIONS] = kpi_decisions
                                st.rerun()
                        with br:
                            if st.button("❌ Reject Mapping", key=f"kpi_reject_{aln.canonical_kpi}"):
                                kpi_decisions[aln.canonical_kpi] = "rejected"
                                st.session_state[_SS_KPI_ALIGN_DECISIONS] = kpi_decisions
                                st.rerun()
                        with bm_col:
                            manual_val = st.text_input(
                                "Manual map to canonical:",
                                key=f"kpi_manual_{aln.canonical_kpi}",
                                placeholder="e.g. gross_reserve",
                                label_visibility="collapsed",
                            )
                            if manual_val and st.button("Apply Manual", key=f"kpi_mapply_{aln.canonical_kpi}"):
                                kpi_decisions[aln.canonical_kpi] = f"manual:{manual_val}"
                                st.session_state[_SS_KPI_ALIGN_DECISIONS] = kpi_decisions
                                st.rerun()

            if aligned:
                aligned_canonicals = [a.canonical_kpi for a in aligned]
                ac1, ac2, ac3, _ = st.columns([1, 1, 1, 3])
                with ac1:
                    if st.button("✅ Accept Section", key="kpi_sec_accept_aligned"):
                        new = dict(kpi_decisions)
                        for c in aligned_canonicals:
                            new[c] = "accepted"
                        st.session_state[_SS_KPI_ALIGN_DECISIONS] = new
                        st.rerun()
                with ac2:
                    if st.button("❌ Reject Section", key="kpi_sec_reject_aligned"):
                        new = dict(kpi_decisions)
                        for c in aligned_canonicals:
                            new[c] = "rejected"
                        st.session_state[_SS_KPI_ALIGN_DECISIONS] = new
                        st.rerun()
                with ac3:
                    if st.button("🔄 Reset Section", key="kpi_sec_reset_aligned"):
                        new = {k: v for k, v in kpi_decisions.items() if k not in aligned_canonicals}
                        st.session_state[_SS_KPI_ALIGN_DECISIONS] = new
                        st.rerun()
                with st.expander(f"✅ Fully Aligned KPI Columns ({len(aligned)})", expanded=False):
                    aligned_rows = []
                    for a in aligned:
                        dec = kpi_decisions.get(a.canonical_kpi, "pending")
                        status_icon = {"accepted": "✅", "rejected": "❌", "pending": "❓"}.get(dec, "❓")
                        row = {
                            "": status_icon,
                            "Canonical KPI": a.canonical_kpi,
                            "KPI Label(s)":  ", ".join(a.kpi_labels),
                            "Decision":      dec,
                        }
                        for fn in file_names:
                            row[fn] = a.per_file_match.get(fn, "")
                        aligned_rows.append(row)
                    st.dataframe(pd.DataFrame(aligned_rows), use_container_width=True, hide_index=True)

            if missing:
                missing_canonicals = [a.canonical_kpi for a in missing]
                mc1, mc2, mc3, _ = st.columns([1, 1, 1, 3])
                with mc1:
                    if st.button("✅ Accept Section", key="kpi_sec_accept_missing"):
                        new = dict(kpi_decisions)
                        for c in missing_canonicals:
                            new[c] = "accepted"
                        st.session_state[_SS_KPI_ALIGN_DECISIONS] = new
                        st.rerun()
                with mc2:
                    if st.button("❌ Reject Section", key="kpi_sec_reject_missing"):
                        new = dict(kpi_decisions)
                        for c in missing_canonicals:
                            new[c] = "rejected"
                        st.session_state[_SS_KPI_ALIGN_DECISIONS] = new
                        st.rerun()
                with mc3:
                    if st.button("🔄 Reset Section", key="kpi_sec_reset_missing"):
                        new = {k: v for k, v in kpi_decisions.items() if k not in missing_canonicals}
                        st.session_state[_SS_KPI_ALIGN_DECISIONS] = new
                        st.rerun()
                with st.expander(f"⚠️ Missing KPI Columns ({len(missing)})", expanded=False):
                    st.caption("Not found in any source file.")
                    st.dataframe(
                        pd.DataFrame([{"Canonical": a.canonical_kpi, "KPI Label(s)": ", ".join(a.kpi_labels)} for a in missing]),
                        use_container_width=True, hide_index=True,
                    )

    # ── Tab 6: Manual Group Overrides ────────────────────────────────────────
    with tab_overrides:
        st.markdown("**File-to-Group Compatibility Matrix**")
        st.caption("Shows how compatible each file would be if moved to each group.")

        matrix_rows = []
        for fp in intel.file_profiles:
            compat = compute_file_group_compatibility(
                fp.file_name, intel.groups, intel.pairwise_scores
            )
            current_gid = next(
                (g.group_id for g in intel.groups if fp.file_name in g.file_names), None
            )
            best_gid, best_score = None, -1.0
            row: dict = {"File": fp.file_name, "Current Group": f"Group {current_gid}" if current_gid else "—"}
            for gid, scores in sorted(compat.items()):
                row[f"Grp {gid} Score"] = f"{scores['overall_score']:.0%}"
                if scores["overall_score"] > best_score:
                    best_score, best_gid = scores["overall_score"], gid
            row["Recommended"] = f"Group {best_gid}" if best_gid else "—"
            matrix_rows.append(row)
        if matrix_rows:
            st.dataframe(pd.DataFrame(matrix_rows), use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("**Per-File Decision Support**")
        for fp in intel.file_profiles:
            compat = compute_file_group_compatibility(
                fp.file_name, intel.groups, intel.pairwise_scores
            )
            with st.expander(f"📊 `{fp.file_name}`", expanded=False):
                for gid, scores in sorted(compat.items()):
                    st.markdown(f"**Group {gid}:** Overall {scores['overall_score']:.0%} → {scores['recommendation']}")
                    dc1, dc2, dc3, dc4, dc5 = st.columns(5)
                    dc1.metric("Grain", f"{scores['grain_score']:.0%}")
                    dc2.metric("KPI Sim", f"{scores['kpi_score']:.0%}")
                    dc3.metric("Schema", f"{scores['schema_score']:.0%}")
                    dc4.metric("Key", f"{scores['key_score']:.0%}")
                    dc5.metric("Time", f"{scores['time_score']:.0%}")

        st.markdown("---")
        st.markdown(
            "**Move files between groups, exclude files, or create new groups.**  "
            "Click **Apply Overrides** when done."
        )

        group_overrides = st.session_state.get(_SS_GROUP_OVERRIDES, {})
        group_ids   = sorted({g.group_id for g in intel.groups})
        max_group_id = max(group_ids) if group_ids else 0

        for fname in file_names:
            current_group = next(
                (g.group_id for g in intel.groups if fname in g.file_names), 0
            )
            override_val = group_overrides.get(fname, current_group)
            exclude_now  = override_val is None

            fc1, fc2, fc3 = st.columns([3, 2, 1])
            with fc1:
                st.markdown(f"`{fname}`")
            with fc2:
                options_labels = [f"Group {gid}" for gid in group_ids] + [f"New Group {max_group_id + 1}"]
                options_ids    = group_ids + [max_group_id + 1]
                safe_val = override_val if override_val in options_ids else current_group
                current_idx = options_ids.index(safe_val) if safe_val in options_ids else 0
                chosen_label = st.selectbox(
                    "Move to group:",
                    options_labels,
                    index=current_idx,
                    key=f"grp_override_{fname}",
                    label_visibility="collapsed",
                )
                new_gid = options_ids[options_labels.index(chosen_label)]
                group_overrides[fname] = new_gid
            with fc3:
                if st.checkbox("Exclude", value=exclude_now, key=f"grp_exclude_{fname}"):
                    group_overrides[fname] = None

        col_apply, col_reset = st.columns(2)
        with col_apply:
            if st.button("Apply Overrides", key="ar_btn_apply_overrides"):
                st.session_state[_SS_GROUP_OVERRIDES] = group_overrides
                st.rerun()
        with col_reset:
            if st.button("Reset to Agent Defaults", key="ar_btn_reset_overrides"):
                st.session_state.pop(_SS_GROUP_OVERRIDES, None)
                st.rerun()

    # ── Agent details + continue ──────────────────────────────────────────────
    with st.expander("Agent Details", expanded=False):
        _render_agent_card(consol_result)

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
