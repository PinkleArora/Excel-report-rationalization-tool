"""Streamlit page — Agentic Rationalization workflow."""

from __future__ import annotations

import io
import traceback
import zipfile
from datetime import datetime
from itertools import groupby

import pandas as pd
import streamlit as st

from src.ingestion.loader import load_workbook_from_bytes
from src.agentic_rationalization.orchestrator import run_pipeline
from src.agentic_rationalization.models import AgentResult, PipelineResult
from src.agentic_rationalization.review_panel import (
    ActionableDecision,
    classify_actionable_decisions,
    SIMILAR_COLUMN_ACTIONS,
    SIMILAR_COLUMN_ACTIONS_VETOED,
    KPI_UNRESOLVED_ACTIONS,
    CONSOLIDATION_ACTIONS,
)


# ── Helpers ──────────────────────────────────────────────────────────────────

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
    n_low = len(result.low_confidence_decisions)
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


# ── Review panel ─────────────────────────────────────────────────────────────

_OVERRIDES_KEY     = "agentic_decision_overrides"    # generic text overrides
_ACTION_KEY        = "agentic_action_overrides"      # structured action choices


def _widget_key(prefix: str, subject: str) -> str:
    safe = subject.replace(" ", "_").replace("/", "_").replace("'", "")
    return f"{prefix}_{safe}"[:120]


# ── Per-type card renderers ───────────────────────────────────────────────────

def _render_similar_column_card(ad: ActionableDecision, idx: int) -> None:
    saved_actions: dict[str, str] = st.session_state.get(_ACTION_KEY, {})
    current_action = saved_actions.get(ad.subject, "ACCEPT")

    icon = _conf_color(ad.confidence)
    with st.container(border=True):
        c1, c2 = st.columns([3, 1])
        with c1:
            st.markdown(
                f"**{icon} Similar Column Match** — confidence `{ad.confidence:.0%}`"
            )
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
        chosen_label = st.radio(
            "Your decision:",
            labels,
            index=default_idx,
            key=key,
            horizontal=False,
        )
        chosen_key = next(k for k, l in ad.available_actions if l == chosen_label)

        # Persist immediately
        if _ACTION_KEY not in st.session_state:
            st.session_state[_ACTION_KEY] = {}
        st.session_state[_ACTION_KEY][ad.subject] = chosen_key

        # Record as annotation override too (for decision log)
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
            st.markdown(
                f"**🔴 Unresolved KPI Formula** — `{ad.cell_address}` in `{ad.kpi_tab}`"
            )
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
        chosen_label = st.radio(
            "Your decision:",
            labels,
            index=default_idx,
            key=key_radio,
            horizontal=False,
        )
        chosen_key = next(k for k, l in ad.available_actions if l == chosen_label)

        # Dropdown replaces free-text: user picks from all known canonical names
        selected_col = ""
        if chosen_key == "MANUAL_COLUMN":
            saved_manual = (
                st.session_state.get(_OVERRIDES_KEY, {})
                .get(ad.agent_name, {})
                .get(ad.subject, "")
            )
            canonicals = ad.available_canonicals or []
            if canonicals:
                # Determine default selection
                default_col = saved_manual if saved_manual in canonicals else (
                    canonicals[0] if canonicals else ""
                )
                default_col_idx = canonicals.index(default_col) if default_col in canonicals else 0
                st.markdown("**Select source column:**")
                st.caption(
                    "Choose the canonical source column this formula should map to. "
                    "Column names shown are the normalised names used in Master Source Data."
                )
                selected_col = st.selectbox(
                    "Source column:",
                    options=canonicals,
                    index=default_col_idx,
                    key=_widget_key("kpi_col", ad.subject),
                )
            else:
                # Fallback: no schema available yet — show text input
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
                selected_col if chosen_key == "MANUAL_COLUMN" and selected_col
                else chosen_key
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
            st.markdown(
                f"**{icon} Consolidation Uncertainty** — `{wb_a}` ↔ `{wb_b}`"
            )
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
        chosen_label = st.radio(
            "Your decision:",
            labels,
            index=default_idx,
            key=key,
            horizontal=False,
        )
        chosen_key = next(k for k, l in ad.available_actions if l == chosen_label)

        if _ACTION_KEY not in st.session_state:
            st.session_state[_ACTION_KEY] = {}
        st.session_state[_ACTION_KEY][ad.subject] = chosen_key
        if chosen_key != "ACCEPT":
            if _OVERRIDES_KEY not in st.session_state:
                st.session_state[_OVERRIDES_KEY] = {}
            st.session_state[_OVERRIDES_KEY].setdefault(ad.agent_name, {})[ad.subject] = chosen_key


# ── Validation summary ───────────────────────────────────────────────────────

def _render_validation_summary(pipeline: PipelineResult) -> None:
    """Structured validation results card showing KPI reconciliation status."""
    val_result = pipeline.get_result("ValidationAgent")
    if val_result is None:
        return
    if not val_result.decisions:
        if val_result.warnings:
            st.info(f"Validation: {val_result.warnings[0]}")
        return

    overall = next((d for d in val_result.decisions if d.subject == "Overall validation"), None)
    if overall is None:
        return

    sig = overall.signals
    total  = sig.get("total_columns", 0)
    passed = sig.get("pass_count", 0)
    warned = sig.get("warn_count", 0)
    failed = sig.get("fail_count", 0)

    expanded = failed > 0
    with st.expander("Validation — KPI Column Reconciliation", expanded=expanded):
        if total == 0:
            st.info("No numeric KPI-referenced columns were found to reconcile.")
            if val_result.warnings:
                for w in val_result.warnings:
                    st.warning(w)
            return

        if failed == 0 and warned == 0:
            st.success(
                f"All {total} KPI column(s) reconciled — source totals match "
                "Master Source Data within tolerance."
            )
        elif failed > 0:
            st.error(
                f"{failed} of {total} column(s) FAILED reconciliation. "
                "Source totals do not match Master Source Data."
            )
        else:
            st.warning(f"{warned} of {total} column(s) have reconciliation warnings.")

        cp, cw, cf = st.columns(3)
        cp.metric("✅ Pass", passed)
        cw.metric("⚠️ Warn", warned)
        cf.metric("❌ Fail", failed)

        detail_rows = []
        for d in val_result.decisions:
            if d.subject == "Overall validation":
                continue
            s = d.signals
            detail_rows.append({
                "Status":      d.decision,
                "Column":      d.subject.replace("Reconciliation: '", "").rstrip("'"),
                "Source Total": f"{s.get('source_total', 0):,.2f}",
                "Master Total": f"{s.get('master_total', 0):,.2f}",
                "Delta":       f"{s.get('delta', 0):+,.2f}",
                "Variance %":  f"{s.get('delta_pct', 0):.4f}%",
                "Tolerance":   f"{s.get('tolerance_pct', 0):.2f}%",
            })
        if detail_rows:
            st.dataframe(pd.DataFrame(detail_rows), use_container_width=True, hide_index=True)

        if val_result.warnings:
            for w in val_result.warnings:
                st.warning(w)


def _render_validation_summary(pipeline: PipelineResult) -> None:
    """Show a structured validation results card if validation ran."""
    val_result = pipeline.get_result("ValidationAgent")
    if val_result is None:
        return
    if not val_result.decisions:
        if val_result.warnings:
            st.info(f"Validation: {val_result.warnings[0]}")
        return

    # Find overall summary decision
    overall = next((d for d in val_result.decisions if d.subject == "Overall validation"), None)
    if overall is None:
        return

    sig = overall.signals
    total = sig.get("total_columns", 0)
    passed = sig.get("pass_count", 0)
    warned = sig.get("warn_count", 0)
    failed = sig.get("fail_count", 0)

    with st.expander("✅ Validation Summary — KPI Column Reconciliation", expanded=(failed > 0)):
        if failed == 0 and warned == 0:
            st.success(f"All {total} KPI column(s) reconciled successfully (totals match within tolerance).")
        elif failed > 0:
            st.error(f"{failed} column(s) FAILED reconciliation — source totals do not match Master Source Data.")
        else:
            st.warning(f"{warned} column(s) have reconciliation warnings.")

        col_p, col_w, col_f = st.columns(3)
        col_p.metric("✅ Pass", passed)
        col_w.metric("⚠️ Warn", warned)
        col_f.metric("❌ Fail", failed)

        # Build per-column table from remaining decisions
        detail_rows = []
        for d in val_result.decisions:
            if d.subject == "Overall validation":
                continue
            sig_d = d.signals
            detail_rows.append({
                "Status": d.decision,
                "Column": d.subject.replace("Reconciliation: '", "").rstrip("'"),
                "Source Total": f"{sig_d.get('source_total', 0):,.2f}",
                "Master Total": f"{sig_d.get('master_total', 0):,.2f}",
                "Delta": f"{sig_d.get('delta', 0):+,.2f}",
                "Variance %": f"{sig_d.get('delta_pct', 0):.4f}%",
                "Tolerance": f"{sig_d.get('tolerance_pct', 0):.2f}%",
            })
        if detail_rows:
            df_val = pd.DataFrame(detail_rows)
            st.dataframe(df_val, use_container_width=True, hide_index=True)

        if val_result.warnings:
            for w in val_result.warnings:
                st.warning(w)


# ── Main review panel ─────────────────────────────────────────────────────────

def _render_review_panel(pipeline: PipelineResult) -> None:
    """Structured review panel — surfaces specific low-confidence decisions
    as typed cards with radio-button action choices.
    """
    st.subheader("Step 4 — Review & Resolve Decisions")
    st.caption(
        "Decisions below require your input. Each card shows exactly what the agent "
        "is uncertain about and offers concrete choices. "
        "Click **Re-run with overrides** after making your selections."
    )

    actionable = classify_actionable_decisions(pipeline)

    if not actionable:
        st.success(
            "All decisions have high confidence (≥ 90%). "
            "No manual review required — you can proceed to download the outputs."
        )
        return

    # Count by type
    n_sim   = sum(1 for a in actionable if a.action_type == "SIMILAR_COLUMN")
    n_kpi   = sum(1 for a in actionable if a.action_type == "KPI_UNRESOLVED")
    n_con   = sum(1 for a in actionable if a.action_type == "CONSOLIDATION_PAIR")

    tab_labels = []
    if n_sim:
        tab_labels.append(f"🔀 Similar Columns ({n_sim})")
    if n_kpi:
        tab_labels.append(f"❓ Unresolved KPI Mappings ({n_kpi})")
    if n_con:
        tab_labels.append(f"⚖️ Consolidation Decisions ({n_con})")

    tabs = st.tabs(tab_labels) if len(tab_labels) > 1 else [st.container()]
    tab_iter = iter(tabs)

    if n_sim:
        with next(tab_iter):
            st.caption(
                "These column pairs have similar names. "
                "Decide whether they represent the same business concept (merge) "
                "or different concepts (keep separate)."
            )
            sim_items = [a for a in actionable if a.action_type == "SIMILAR_COLUMN"]
            for i, ad in enumerate(sim_items):
                _render_similar_column_card(ad, i)

    if n_kpi:
        with next(tab_iter):
            st.caption(
                "These source-backed KPI formulas could not be automatically resolved "
                "to a source column. Select the correct source column from the dropdown, "
                "or skip to exclude the formula from dependency mapping."
            )
            kpi_items = [a for a in actionable if a.action_type == "KPI_UNRESOLVED"]
            for i, ad in enumerate(kpi_items):
                _render_kpi_unresolved_card(ad, i)

    if n_con:
        with next(tab_iter):
            st.caption(
                "These workbooks have low schema overlap or grain compatibility issues. "
                "The agent recommends manual review. "
                "For each pair, choose whether to consolidate into Master Source Data or keep separate."
            )
            con_items = [a for a in actionable if a.action_type == "CONSOLIDATION_PAIR"]
            for i, ad in enumerate(con_items):
                _render_consolidation_card(ad, i)


# ── Main page ─────────────────────────────────────────────────────────────────

def render() -> None:
    st.title("Agentic Rationalization")
    st.caption(
        "AI agents analyse your workbooks, explain every decision with confidence scores, "
        "and let you override before generating the final output."
    )

    # ── Step 1: Upload ────────────────────────────────────────────────────────
    st.header("Step 1 — Upload Workbooks")
    uploaded = st.file_uploader(
        "Upload one or more Excel workbooks (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="agentic_upload",
    )

    if not uploaded:
        st.info("Upload at least one workbook to begin.")
        return

    # ── Step 2: Run Pipeline ──────────────────────────────────────────────────
    st.header("Step 2 — Run Agentic Analysis")

    with st.expander("Advanced options", expanded=False):
        tolerance = st.slider(
            "Reconciliation tolerance (%)",
            min_value=0.0, max_value=5.0, value=1.0, step=0.1,
        ) / 100.0

    run_key = "agentic_pipeline_result"

    col_run, col_rerun = st.columns([1, 1])
    run_clicked = col_run.button("▶ Run Agents", type="primary")
    rerun_clicked = col_rerun.button(
        "🔄 Re-run with overrides",
        disabled=st.session_state.get(run_key) is None,
        help="Apply any overrides entered below and regenerate the output.",
    )

    def _execute_pipeline(decision_overrides: dict | None, action_overrides: dict | None) -> None:
        with st.spinner("Loading workbooks…"):
            try:
                bundles = [load_workbook_from_bytes(f.read(), file_name=f.name) for f in uploaded]
            except Exception as exc:
                tb = traceback.format_exc()
                import logging
                logging.getLogger(__name__).error("Failed to load workbooks:\n%s", tb)
                st.error(f"**Failed to load workbooks:** `{type(exc).__name__}: {exc}`")
                st.code(tb, language="python")
                return
        with st.spinner("Running agentic pipeline (this may take a moment)…"):
            try:
                result = run_pipeline(
                    bundles,
                    tolerance=tolerance,
                    decision_overrides=decision_overrides,
                    action_overrides=action_overrides,
                )
                st.session_state[run_key] = result
                st.session_state["agentic_bundles"] = bundles
            except Exception as exc:
                tb = traceback.format_exc()
                import logging
                logging.getLogger(__name__).error("Pipeline error:\n%s", tb)
                st.error(f"**Pipeline error:** `{type(exc).__name__}: {exc}`")
                st.code(tb, language="python")

    if run_clicked:
        st.session_state.pop(_OVERRIDES_KEY, None)
        st.session_state.pop(_ACTION_KEY, None)
        _execute_pipeline(None, None)
    elif rerun_clicked:
        _execute_pipeline(
            st.session_state.get(_OVERRIDES_KEY),
            st.session_state.get(_ACTION_KEY),
        )

    pipeline: PipelineResult | None = st.session_state.get(run_key)
    if pipeline is None:
        return

    # ── Step 3: Agent Recommendations ────────────────────────────────────────
    st.header("Step 3 — Agent Recommendations")

    if pipeline.errors:
        for err in pipeline.errors:
            st.error(err)

    total_decisions = sum(len(r.decisions) for r in pipeline.agent_results)
    high = sum(len(r.high_confidence_decisions) for r in pipeline.agent_results)
    med = sum(len(r.medium_confidence_decisions) for r in pipeline.agent_results)
    low = sum(len(r.low_confidence_decisions) for r in pipeline.agent_results)
    total_warnings = sum(len(r.warnings) for r in pipeline.agent_results)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Decisions", total_decisions)
    c2.metric("🟢 High Confidence", high)
    c3.metric("🟡 Medium Confidence", med)
    c4.metric("🔴 Low Confidence", low)
    c5.metric("⚠️ Warnings", total_warnings)

    # Decision Explorer — filter decisions by confidence level
    with st.expander("🔍 Decision Explorer", expanded=False):
        conf_tab_all, conf_tab_high, conf_tab_med, conf_tab_low, conf_tab_warn = st.tabs([
            "All",
            f"🟢 High ({high})",
            f"🟡 Medium ({med})",
            f"🔴 Low ({low})",
            f"⚠️ Warnings ({total_warnings})",
        ])

        all_rows = []
        for ar in pipeline.agent_results:
            for d in ar.decisions:
                all_rows.append({
                    "Agent":      ar.agent_name,
                    "":           _conf_color(d.confidence),
                    "Subject":    d.subject[:80],
                    "Decision":   d.decision[:60],
                    "Confidence": f"{d.confidence:.0%}",
                    "Overridable": "Yes" if d.overridable else "No",
                    "Reasoning":  d.reasoning[:120],
                })

        df_all  = pd.DataFrame(all_rows)
        with conf_tab_all:
            st.dataframe(df_all, use_container_width=True, hide_index=True)
        with conf_tab_high:
            df_h = df_all[df_all["Confidence"].apply(lambda x: float(x.rstrip("%")) >= 90)]
            st.dataframe(df_h, use_container_width=True, hide_index=True) if not df_h.empty else st.info("No high-confidence decisions.")
        with conf_tab_med:
            df_m = df_all[df_all["Confidence"].apply(lambda x: 70 <= float(x.rstrip("%")) < 90)]
            st.dataframe(df_m, use_container_width=True, hide_index=True) if not df_m.empty else st.info("No medium-confidence decisions.")
        with conf_tab_low:
            df_l = df_all[df_all["Confidence"].apply(lambda x: float(x.rstrip("%")) < 70)]
            st.dataframe(df_l, use_container_width=True, hide_index=True) if not df_l.empty else st.info("No low-confidence decisions.")
        with conf_tab_warn:
            warn_agents = {ar.agent_name for ar in pipeline.agent_results if ar.warnings}
            df_w = df_all[df_all["Agent"].isin(warn_agents)]
            st.dataframe(df_w, use_container_width=True, hide_index=True) if not df_w.empty else st.info("No warnings.")

    for agent_result in pipeline.agent_results:
        _render_agent_card(agent_result)

    _render_validation_summary(pipeline)

    # ── Step 4: Structured Review ─────────────────────────────────────────────
    _render_review_panel(pipeline)

    # ── Step 5: Downloads ─────────────────────────────────────────────────────
    st.header("Step 5 — Download Outputs")

    any_output = (
        pipeline.future_state_bytes
        or pipeline.analysis_pack_bytes
        or pipeline.decision_log_bytes
    )

    if not any_output:
        st.warning(
            "No output workbooks were generated. "
            "Check the agent recommendations above for blocking issues."
        )
        return

    col1, col2, col3, col4 = st.columns(4)

    if pipeline.future_state_bytes:
        col1.download_button(
            label="📥 Future State Workbook",
            data=pipeline.future_state_bytes,
            file_name="Future_State_Workbook.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if pipeline.analysis_pack_bytes:
        col2.download_button(
            label="📥 Analysis Pack",
            data=pipeline.analysis_pack_bytes,
            file_name="Rationalization_Analysis_Pack.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if pipeline.decision_log_bytes:
        col3.download_button(
            label="📥 Agent Decision Log",
            data=pipeline.decision_log_bytes,
            file_name="Agent_Decision_Log.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ZIP of all outputs
    if sum(1 for x in [pipeline.future_state_bytes, pipeline.analysis_pack_bytes, pipeline.decision_log_bytes] if x) > 1:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            if pipeline.future_state_bytes:
                zf.writestr("Future_State_Workbook.xlsx", pipeline.future_state_bytes)
            if pipeline.analysis_pack_bytes:
                zf.writestr("Rationalization_Analysis_Pack.xlsx", pipeline.analysis_pack_bytes)
            if pipeline.decision_log_bytes:
                zf.writestr("Agent_Decision_Log.xlsx", pipeline.decision_log_bytes)
        col4.download_button(
            label="📦 Download All (ZIP)",
            data=zip_buf.getvalue(),
            file_name=f"Agentic_Rationalization_{datetime.now():%Y%m%d_%H%M}.zip",
            mime="application/zip",
        )
