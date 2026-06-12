"""Build the Agent Decision Log workbook from all AgentResult objects."""

from __future__ import annotations

import io
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from src.agentic_rationalization.models import AgentDecision, AgentResult

# Confidence colour bands
_FILL_HIGH   = PatternFill("solid", fgColor="C6EFCE")   # green ≥ 0.90
_FILL_MED    = PatternFill("solid", fgColor="FFEB9C")   # amber 0.70–0.89
_FILL_LOW    = PatternFill("solid", fgColor="FFC7CE")   # red   < 0.70
_FILL_HEADER = PatternFill("solid", fgColor="1F4E79")   # dark blue


def _conf_fill(conf: float) -> PatternFill:
    if conf >= 0.90:
        return _FILL_HIGH
    if conf >= 0.70:
        return _FILL_MED
    return _FILL_LOW


def _write_header(ws, headers: list[str]) -> None:
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = _FILL_HEADER
        cell.alignment = Alignment(wrap_text=True)
    ws.row_dimensions[1].height = 20


def _auto_width(ws) -> None:
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)


def _write_df(ws, df: pd.DataFrame, title: str | None = None) -> None:
    row_offset = 0
    if title:
        ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=13)
        row_offset = 2
    _write_header(ws, list(df.columns))
    # Shift headers down if title was written
    if row_offset:
        # Move header row down
        for col_idx, h in enumerate(df.columns, start=1):
            ws.cell(row=row_offset + 1, column=col_idx, value=h).font = Font(bold=True, color="FFFFFF")
            ws.cell(row=row_offset + 1, column=col_idx).fill = _FILL_HEADER
        for r_idx, row in enumerate(df.itertuples(index=False), start=row_offset + 2):
            for c_idx, val in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=val)
    else:
        for r_idx, row in enumerate(df.itertuples(index=False), start=2):
            for c_idx, val in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=val)


def build(agent_results: list[AgentResult]) -> bytes:
    """Assemble Agent_Decision_Log.xlsx and return bytes."""
    wb = Workbook()
    wb.remove(wb.active)

    # ── 01 Agent Summary ─────────────────────────────────────────────────────
    ws_summary = wb.create_sheet("01_Agent_Summary")
    summary_rows = []
    for result in agent_results:
        total = len(result.decisions)
        high   = len(result.high_confidence_decisions)
        med    = len(result.medium_confidence_decisions)
        low    = len(result.low_confidence_decisions)
        summary_rows.append({
            "Agent":             result.agent_name,
            "Total Decisions":   total,
            "High Conf (≥90%)":  high,
            "Med Conf (70–90%)": med,
            "Low Conf (<70%)":   low,
            "Warnings":          len(result.warnings),
            "Warning Detail":    " | ".join(result.warnings),
        })
    summary_df = pd.DataFrame(summary_rows)
    _write_df(ws_summary, summary_df, title=f"Agent Decision Log — {datetime.now():%Y-%m-%d %H:%M}")
    _auto_width(ws_summary)

    # ── Per-agent decision sheets ─────────────────────────────────────────────
    for i, result in enumerate(agent_results, start=2):
        ws = wb.create_sheet(f"{i:02d}_{result.agent_name}")
        rows = []
        for d in result.decisions:
            rows.append({
                "Subject":    d.subject,
                "Decision":   d.decision,
                "Confidence": f"{d.confidence:.0%}",
                "Reasoning":  d.reasoning,
                "Overridable": "Yes" if d.overridable else "No",
                "Key Signals": "; ".join(
                    f"{k}={v}" for k, v in list(d.signals.items())[:5]
                ),
            })
        if not rows:
            ws.cell(row=1, column=1, value="No decisions recorded.")
            continue
        df = pd.DataFrame(rows)
        headers = list(df.columns)
        _write_header(ws, headers)
        for r_idx, decision in enumerate(result.decisions, start=2):
            row_data = rows[r_idx - 2]
            for c_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=row_data[h])
                if h == "Confidence":
                    cell.fill = _conf_fill(decision.confidence)
                cell.alignment = Alignment(wrap_text=True)
        _auto_width(ws)

    # ── Confidence Legend ─────────────────────────────────────────────────────
    ws_legend = wb.create_sheet("Legend")
    ws_legend.cell(row=1, column=1, value="Confidence Colour Key").font = Font(bold=True)
    for row_idx, (label, fill, threshold) in enumerate([
        ("High (≥ 90%) — decision applied automatically",   _FILL_HIGH, ""),
        ("Medium (70–89%) — flagged for user awareness",    _FILL_MED,  ""),
        ("Low (< 70%) — user confirmation recommended",     _FILL_LOW,  ""),
    ], start=3):
        ws_legend.cell(row=row_idx, column=1, value=label).fill = fill

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
