"""Streamlit page — Agentic Rationalization workflow."""

from __future__ import annotations

import io
import zipfile
from datetime import datetime
from itertools import groupby

import pandas as pd
import streamlit as st

from src.ingestion.loader import load_workbook_from_bytes
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

_OVERRIDES_KEY = "agentic_decision_overrides"


def _override_field_key(agent_name: str, subject: str) -> str:
    """Stable widget key for a single override field."""
    safe = (agent_name + "__" + subject).replace(" ", "_").replace("/", "_")
    return f"ovr_{safe}"[:128]


def _render_overrides(pipeline: PipelineResult) -> None:
    """Display every overridable agent decision as an editable field.

    Values are stored in st.session_state[_OVERRIDES_KEY] so they survive
    re-runs and are passed back into the pipeline.
    """
    st.subheader("User Overrides")
    st.caption(
        "All decisions marked **Overridable = Yes** are shown below.  "
        "Leave a field blank (or unchanged) to accept the agent recommendation.  "
        "Click **Re-run with overrides** to regenerate the output."
    )

    # Collect every overridable decision across all agents
    overridable = [
        (result.agent_name, d)
        for result in pipeline.agent_results
        for d in result.decisions
        if d.overridable
    ]

    if not overridable:
        st.info("No overridable decisions found in this pipeline run.")
        return

    # Group by agent for readability
    overridable_sorted = sorted(overridable, key=lambda x: x[0])
    saved: dict[str, dict[str, str]] = st.session_state.get(_OVERRIDES_KEY, {})

    for agent_name, group in groupby(overridable_sorted, key=lambda x: x[0]):
        decisions_in_group = [d for _, d in group]
        # Skip agents whose decisions are all high-confidence (≥ 0.95) — read-only
        all_high = all(d.confidence >= 0.95 for d in decisions_in_group)
        if all_high:
            with st.expander(
                f"🔒 {agent_name} — {len(decisions_in_group)} decision(s), all high-confidence (read-only)",
                expanded=False,
            ):
                st.caption("These decisions have ≥ 95% confidence and are shown for information only.")
                for d in decisions_in_group:
                    st.markdown(
                        f"**{d.subject}** → `{d.decision}` "
                        f"({d.confidence:.0%} confidence)"
                    )
            continue

        with st.expander(
            f"✏️ {agent_name} — {len(decisions_in_group)} overridable decision(s)",
            expanded=True,
        ):
            agent_saved = saved.get(agent_name, {})
            for d in decisions_in_group:
                conf_icon = _conf_color(d.confidence)
                st.markdown(
                    f"{conf_icon} **{d.subject}** &nbsp;&nbsp; "
                    f"Confidence: `{d.confidence:.0%}`"
                )
                st.caption(d.reasoning)
                field_key = _override_field_key(agent_name, d.subject)
                existing = agent_saved.get(d.subject, "")
                new_val = st.text_input(
                    label=f"Agent decision: `{d.decision}`",
                    value=existing,
                    placeholder=f"Leave blank to accept: {d.decision}",
                    key=field_key,
                )
                # Persist to session_state immediately on each interaction
                if new_val != existing:
                    if _OVERRIDES_KEY not in st.session_state:
                        st.session_state[_OVERRIDES_KEY] = {}
                    st.session_state[_OVERRIDES_KEY].setdefault(agent_name, {})[d.subject] = new_val
                st.divider()


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

    def _execute_pipeline(overrides: dict | None) -> None:
        with st.spinner("Loading workbooks…"):
            try:
                bundles = [load_workbook_from_bytes(f.read(), file_name=f.name) for f in uploaded]
            except Exception as exc:
                st.error(f"Failed to load workbooks: {exc}")
                return
        with st.spinner("Running agentic pipeline (this may take a moment)…"):
            try:
                result = run_pipeline(
                    bundles,
                    tolerance=tolerance,
                    decision_overrides=overrides,
                )
                st.session_state[run_key] = result
                st.session_state["agentic_bundles"] = bundles
            except Exception as exc:
                st.error(f"Pipeline error: {exc}")

    if run_clicked:
        st.session_state.pop(_OVERRIDES_KEY, None)  # clear any stale overrides on fresh run
        _execute_pipeline(None)
    elif rerun_clicked:
        _execute_pipeline(st.session_state.get(_OVERRIDES_KEY))

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
    st.caption(
        "Saved overrides persist across re-runs. "
        "Click **Re-run with overrides** (Step 2) after making changes."
    )
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
