"""Streamlit page — Agentic Rationalization workflow."""

from __future__ import annotations

import io
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

from src.workbook_loader.loader import load_workbooks
from src.agentic_rationalization.orchestrator import run_pipeline
from src.agentic_rationalization.models import AgentResult, PipelineResult


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


# ── Override UI ───────────────────────────────────────────────────────────────

def _render_overrides(pipeline: PipelineResult) -> dict:
    """Render the user-override section. Returns override dict for re-run."""
    st.subheader("User Overrides")
    st.caption(
        "Decisions marked 'Overridable' can be changed below. "
        "Re-run the pipeline after saving changes."
    )

    discovery_result = next(
        (r for r in pipeline.agent_results if r.agent_name == "DiscoveryAgent"), None
    )
    config = discovery_result.output if discovery_result else None

    overrides: dict = {}

    if config is None:
        st.info("No config available for overrides.")
        return overrides

    st.markdown("**Tab Role Assignments** (from Discovery Agent)")
    for wb_cfg in config.workbook_configs:
        st.markdown(f"*{wb_cfg.workbook_name}*")
        col1, col2 = st.columns(2)
        with col1:
            st.text_input(
                "Source tab",
                value=wb_cfg.source_tab,
                key=f"src_{wb_cfg.workbook_name}",
                disabled=True,
            )
        with col2:
            st.text_input(
                "KPI tabs",
                value=", ".join(wb_cfg.kpi_tabs),
                key=f"kpi_{wb_cfg.workbook_name}",
                disabled=True,
            )

    st.markdown("**Output Tab Names**")
    new_master = st.text_input(
        "Master Source Data tab name",
        value=config.future_source_tab_name,
        key="override_master_tab",
    )
    if new_master and new_master != config.future_source_tab_name:
        overrides["future_source_tab_name"] = new_master[:31]

    return overrides


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
    if st.button("▶ Run Agents", type="primary"):
        file_objects = {f.name: io.BytesIO(f.read()) for f in uploaded}
        with st.spinner("Loading workbooks…"):
            try:
                bundles = load_workbooks(file_objects)
            except Exception as exc:
                st.error(f"Failed to load workbooks: {exc}")
                return

        with st.spinner("Running agentic pipeline (this may take a moment)…"):
            try:
                result = run_pipeline(bundles, tolerance=tolerance)
                st.session_state[run_key] = result
                st.session_state["agentic_bundles"] = bundles
            except Exception as exc:
                st.error(f"Pipeline error: {exc}")
                return

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

    for agent_result in pipeline.agent_results:
        _render_agent_card(agent_result)

    # ── Step 4: Overrides ─────────────────────────────────────────────────────
    st.header("Step 4 — Review & Override")
    _render_overrides(pipeline)

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
