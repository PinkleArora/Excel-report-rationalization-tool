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

import pandas as pd
import streamlit as st

from src.ingestion.loader import load_workbook_from_bytes
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
_SS_CONSOLIDATION  = "ar_consolidation"
_SS_SELECTED_GROUP = "ar_selected_group"
_SS_SCHEMA         = "ar_schema"
_SS_KPI            = "ar_kpi"
_SS_VALIDATION     = "ar_validation"
_SS_GENERATION     = "ar_generation"
_SS_SOURCE_TAB_OVERRIDES = "ar_src_overrides"

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
        _SS_STEP, _SS_BUNDLES, _SS_CONFIG, _SS_DISCOVERY, _SS_CONSOLIDATION,
        _SS_SELECTED_GROUP, _SS_SCHEMA, _SS_KPI, _SS_VALIDATION, _SS_GENERATION,
        _SS_SOURCE_TAB_OVERRIDES, _OVERRIDES_KEY, _ACTION_KEY,
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


# ── Progress indicator ────────────────────────────────────────────────────────

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


def _render_progress_bar() -> None:
    step = _current_step()
    cols = st.columns(len(_STEP_LABELS))
    for i, (col, label) in enumerate(zip(cols, _STEP_LABELS), start=1):
        if i < step:
            col.markdown(f"<div style='text-align:center;color:#22c55e;font-size:0.8em'>✅ {label}</div>", unsafe_allow_html=True)
        elif i == step:
            col.markdown(f"<div style='text-align:center;color:#3b82f6;font-weight:bold;font-size:0.8em'>▶ {label}</div>", unsafe_allow_html=True)
        else:
            col.markdown(f"<div style='text-align:center;color:#9ca3af;font-size:0.8em'>{label}</div>", unsafe_allow_html=True)
    st.markdown("---")


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


def _render_step_2() -> None:
    step = _current_step()
    if step < 2:
        return

    st.subheader("Step 2 — Discovery")

    bundles = st.session_state.get(_SS_BUNDLES, [])
    if not bundles:
        st.warning("No bundles loaded. Return to Step 1.")
        return

    # Run discovery if not already done
    discovery_result: AgentResult | None = st.session_state.get(_SS_DISCOVERY)
    if discovery_result is None:
        overrides = st.session_state.get(_SS_SOURCE_TAB_OVERRIDES, {})
        # Convert flat overrides {file_name: tab_name} to discovery_agent format
        discovery_overrides = {
            fname: {"source_tab": tab} for fname, tab in overrides.items()
        }
        with st.spinner("Running Discovery Agent…"):
            try:
                discovery_result = run_discovery_phase(bundles, overrides=discovery_overrides)
            except Exception as exc:
                st.error(f"Discovery failed: `{type(exc).__name__}: {exc}`")
                st.code(traceback.format_exc(), language="python")
                return
        st.session_state[_SS_DISCOVERY] = discovery_result
        config = discovery_result.output
        st.session_state[_SS_CONFIG] = config

    config = discovery_result.output
    if config is None:
        st.error("Discovery Agent did not produce a configuration.")
        for w in discovery_result.warnings:
            st.warning(w)
        return

    st.warning(
        "Source tab selection affects all downstream analysis. "
        "Confirm before proceeding."
    )

    # Per-workbook display
    changed = False
    src_overrides: dict[str, str] = st.session_state.get(_SS_SOURCE_TAB_OVERRIDES, {})

    for bundle in bundles:
        fname = bundle.file_name
        wb_cfg = config.config_for(fname)
        source_tab = src_overrides.get(fname) or (wb_cfg.source_tab if wb_cfg else "")
        kpi_tabs   = wb_cfg.kpi_tabs if wb_cfg else []
        all_sheets = list(bundle.sheets.keys())

        # Find confidence for the source tab from decisions
        tab_conf = 0.0
        for d in discovery_result.decisions:
            if d.subject == f"{fname} / {source_tab}":
                tab_conf = d.confidence
                break

        # Grain info if available
        consol_result: AgentResult | None = st.session_state.get(_SS_CONSOLIDATION)
        grain_label = "—"
        row_count = col_count = 0
        if consol_result and consol_result.output:
            intel = consol_result.output.get("intelligence")
            if intel:
                fp = next((p for p in intel.file_profiles if p.file_name == fname), None)
                if fp:
                    grain_label = fp.inferred_grain
                    row_count   = fp.row_count
                    col_count   = fp.column_count

        with st.container(border=True):
            c1, c2 = st.columns([3, 1])
            with c1:
                st.markdown(f"**📁 {fname}**")
                conf_icon = _conf_color(tab_conf)
                conf_pct  = f"{tab_conf:.0%}" if tab_conf > 0 else "—"
                st.caption(
                    f"Source Tab: `{source_tab}`  {conf_icon} confidence: {conf_pct}   |   "
                    f"KPI Tabs: {', '.join(kpi_tabs) if kpi_tabs else '—'}"
                )
                if row_count:
                    st.caption(f"Rows: {row_count:,} | Columns: {col_count} | Grain: {grain_label}")
            with c2:
                if all_sheets:
                    default_idx = all_sheets.index(source_tab) if source_tab in all_sheets else 0
                    new_tab = st.selectbox(
                        "Change source tab:",
                        all_sheets,
                        index=default_idx,
                        key=f"ar_src_tab_{fname}",
                    )
                    if new_tab != source_tab:
                        src_overrides[fname] = new_tab
                        changed = True

    if changed:
        st.session_state[_SS_SOURCE_TAB_OVERRIDES] = src_overrides
        st.session_state.pop(_SS_DISCOVERY, None)
        st.session_state.pop(_SS_CONSOLIDATION, None)
        st.session_state.pop(_SS_SCHEMA, None)
        st.rerun()

    with st.expander("Agent Details", expanded=False):
        _render_agent_card(discovery_result)

    if step == 2:
        if st.button("✅ Confirm Source Tabs & Continue", type="primary", key="ar_btn_confirm_discovery"):
            # Run consolidation now so step 3 has data immediately
            with st.spinner("Running Consolidation Intelligence…"):
                try:
                    consol = run_consolidation_phase(bundles, config)
                    st.session_state[_SS_CONSOLIDATION] = consol
                except Exception as exc:
                    st.error(f"Consolidation analysis failed: `{type(exc).__name__}: {exc}`")
                    st.code(traceback.format_exc(), language="python")
                    return
            _set_step(3)
            st.rerun()


def _render_step_3() -> None:
    step = _current_step()
    if step < 3:
        return

    st.subheader("Step 3 — Consolidation Analysis")

    bundles = st.session_state.get(_SS_BUNDLES, [])
    config  = st.session_state.get(_SS_CONFIG)
    consol_result: AgentResult | None = st.session_state.get(_SS_CONSOLIDATION)

    if consol_result is None or consol_result.output is None:
        st.warning("Consolidation analysis not yet available.")
        return

    intel = consol_result.output.get("intelligence")
    if intel is None:
        st.info("No intelligence output available.")
        return

    # File profile table
    if intel.file_profiles:
        st.markdown("**File Profiles**")
        rows = []
        for fp in intel.file_profiles:
            rows.append({
                "File":          fp.file_name,
                "Source Tab":    fp.source_tab,
                "Rows":          fp.row_count,
                "Cols":          fp.column_count,
                "Grain":         fp.inferred_grain,
                "Grain Conf":    f"{fp.grain_confidence:.0%}",
                "KPI Cols":      fp.kpi_referenced_column_count,
                "Non-KPI Cols":  fp.non_kpi_column_count,
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.markdown("**Consolidation Groups**")

    if not intel.groups:
        st.info("No groups detected.")
    else:
        for group in intel.groups:
            icon = "🟢" if group.recommendation == "Consolidate" else "🟡"
            rec_tag = f"[Consolidate ✓]" if group.recommendation == "Consolidate" else "[Keep Separate]"
            with st.expander(
                f"{icon} Group {group.group_id} — {group.group_label}  {rec_tag}",
                expanded=True,
            ):
                st.caption(group.reasoning)
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Common Columns", group.common_column_count)
                c2.metric("Unique Columns", group.unique_column_count)
                c3.metric("KPI Columns", group.kpi_required_column_count)
                c4.metric("Discardable", group.discardable_column_count)
                st.markdown("**Files:**")
                for fname in group.file_names:
                    col_file, col_action = st.columns([4, 1])
                    col_file.markdown(f"- `{fname}`")

    fr = intel.final_recommendation
    if fr.consolidate_files:
        st.success(f"**{fr.summary}**  {fr.business_reasoning}")
    else:
        st.info(f"**{fr.summary}**  {fr.business_reasoning}")

    with st.expander("Agent Details", expanded=False):
        _render_agent_card(consol_result)

    if step == 3:
        if st.button("✅ Confirm Groups & Continue", type="primary", key="ar_btn_confirm_consol"):
            _set_step(4)
            st.rerun()


def _render_step_4() -> None:
    step = _current_step()
    if step < 4:
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

    options = []
    for group in intel.groups:
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
        for i, group in enumerate(intel.groups):
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
    chosen_group = intel.groups[chosen_idx]

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
    if step < 5:
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

    # Schema Dashboard
    all_profiles = source_analysis.all_columns
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
    if step < 6:
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
    if step < 7:
        return

    st.subheader("Step 7 — Validation")

    bundles        = st.session_state.get(_SS_BUNDLES, [])
    config         = st.session_state.get(_SS_CONFIG)
    kpi_result_obj = st.session_state.get(_SS_KPI)
    kpi_analysis   = kpi_result_obj.output if kpi_result_obj else None
    gen_result     = st.session_state.get(_SS_GENERATION)
    val_result: AgentResult | None = st.session_state.get(_SS_VALIDATION)

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
    if step < 8:
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


# ── Main page ─────────────────────────────────────────────────────────────────

def render() -> None:
    st.title("Agentic Rationalization")
    st.caption(
        "A step-by-step consulting workflow — AI agents analyse your workbooks, "
        "explain every decision with confidence scores, "
        "and let you confirm before each phase proceeds."
    )

    _render_progress_bar()

    _render_step_1()
    _render_step_2()
    _render_step_3()
    _render_step_4()
    _render_step_5()
    _render_step_6()
    _render_step_7()
    _render_step_8()
