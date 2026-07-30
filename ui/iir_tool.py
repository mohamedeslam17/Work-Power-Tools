"""IIR Review — 'Is this incoming-inspection report wrong anywhere, and where?'"""
import io
import os
import tempfile
from itertools import groupby
from pathlib import Path

import streamlit as st

import iir_review

from ui import components

_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_CHOICE_LABEL = {iir_review.FAIL: "🔴 Fail", iir_review.WARN: "🟠 Warn",
                iir_review.INFO: "🔵 Info", iir_review.OFF: "⚪ Off"}
# Preset severity bumps, applied to each check's *default* severity.
_STRICT = {iir_review.INFO: iir_review.WARN, iir_review.WARN: iir_review.FAIL,
          iir_review.FAIL: iir_review.FAIL}
_LENIENT = {iir_review.FAIL: iir_review.WARN, iir_review.WARN: iir_review.INFO,
           iir_review.INFO: iir_review.OFF}


@st.cache_data(show_spinner=False)
def _cached_parse(name, data):
    """Parse one IIR workbook, cached on the file bytes so tweaking check
    severities re-checks instantly without re-reading the file."""
    with tempfile.TemporaryDirectory() as tmp:
        # Fixed temp name — never join the upload's raw name to a path.
        path = os.path.join(tmp, "iir.xlsx")
        with open(path, "wb") as fh:
            fh.write(data)
        return iir_review.parse_iir(path)


def _apply_preset(mapping):
    for title, default in iir_review.DEFAULT_SEVERITY.items():
        st.session_state[f"iir_sev::{title}"] = mapping[default] if mapping else default


def _reset_overrides():
    for title in iir_review.DEFAULT_SEVERITY:
        st.session_state.pop(f"iir_sev::{title}", None)


def _check_settings():
    """Severity popover: presets first, per-check tuning below. Returns
    {check_title: severity|OFF}."""
    with st.popover("⚙️ Check severities"):
        st.caption("A preset sets every check at once; tune individual checks below.")
        b1, b2, b3 = st.columns(3)
        if b1.button("Strict", width="stretch"):
            _apply_preset(_STRICT)
        if b2.button("Default", width="stretch"):
            _reset_overrides()
        if b3.button("Lenient", width="stretch"):
            _apply_preset(_LENIENT)
        st.divider()
        chosen = {}
        for cat, items in groupby(iir_review.CHECK_CATALOG, key=lambda t: t[0]):
            st.markdown(f"**{cat}**")
            cols = st.columns(2)
            for j, (_, title, default) in enumerate(items):
                key = f"iir_sev::{title}"
                st.session_state.setdefault(key, default)   # seed once; presets/user own it after
                chosen[title] = cols[j % 2].selectbox(
                    title, iir_review.SEVERITY_CHOICES,
                    format_func=lambda s: _CHOICE_LABEL[s], key=key)

    overridden = sum(1 for title, default in iir_review.DEFAULT_SEVERITY.items()
                     if chosen.get(title, default) != default)
    if overridden:
        c1, c2 = st.columns([6, 1])
        c1.caption(f"⚙️ {overridden} check{'s' if overridden != 1 else ''} overridden")
        if c2.button("Reset", key="iir_reset_line"):
            _reset_overrides()
            st.rerun()
    return chosen


def render():
    files = st.file_uploader(
        "Upload IIR workbook(s) to review (.xlsx)",
        type=["xlsx"], accept_multiple_files=True, key="iir_files")
    if not files:
        return

    parsed_items, perrs = [], []
    with st.spinner(f"Reading {len(files)} report(s)…"):
        for f in files:
            try:
                parsed_items.append({'src': f.name, 'data': _cached_parse(f.name, f.getvalue())})
            except Exception as e:
                perrs.append(f"{f.name}: {e}")
    for e in perrs:
        st.error(f"Could not read — {e}")
    if not parsed_items:
        return

    overrides = _check_settings()

    # Re-run the checks for every report with the current severity settings
    # (cheap — no re-parsing) so tuning applies instantly.
    results = []
    for item in parsed_items:
        data = item['data']
        findings = iir_review.run_checks(data, overrides)
        results.append({
            'src': item['src'], 'data': data, 'findings': findings,
            'ident': data['ident'], 'rp': data['received_parts'],
            'template': data.get('template', '?'),
        })

    if len(results) > 1:
        items = []
        for r in results:
            rows = components.normalize_iir(r['findings'])
            rp = r['rp']
            items.append({
                'counts': components.count_by_severity(rows), 'Report': r['src'],
                'Doc No': r['ident'].get('doc_no', ''), 'Component': r['ident'].get('component', ''),
                'Recv': rp.get('received') if rp.get('found') else None,
                'Top issue': iir_review.top_issue(r['findings']), '_ref': r,
            })
        ranked, idx = components.batch_table(items, key="iir_batch")

        records = [iir_review.record_of(r['data'], r['findings']) for r in results]
        buf = io.BytesIO()
        iir_review.build_batch_summary(records, buf)
        st.download_button(
            "⬇ Batch summary (.xlsx)", data=buf.getvalue(),
            file_name="IIR_Batch_Summary.xlsx", mime=_XLSX_MIME, key="iir_batch_dl")

        pooled = [row for it in ranked
                  for row in components.normalize_iir(it['_ref']['findings'], report=it['_ref']['src'])]
        if pooled:
            with st.expander(f"All findings across the batch ({len(pooled)})"):
                components.findings_table(pooled, key="iir_pool", show_report=True)
        selected = ranked[idx]['_ref']
    else:
        selected = results[0]

    _report_card(selected)


def _report_card(r):
    data, ident = r['data'], r['ident']
    rows = components.normalize_iir(r['findings'])
    counts = components.count_by_severity(rows)
    with st.container(border=True):
        tag = f"Template {r['template']}" if r['template'] not in ('unknown', '?') else None
        facts = (f"{ident.get('doc_no', '?')} · {ident.get('customer', '?')} · "
                 f"{ident.get('component', '?')}  —  prepared by "
                 f"{ident.get('preparer') or ident.get('author', '?')}, "
                 f"reviewed by {ident.get('reviewer', '?')}")
        components.report_header(r['src'], tag=tag, facts=facts)
        if r['template'] == 'unknown':
            st.warning("Unrecognized IIR layout — didn't match a known template, so the review "
                       "couldn't be fully scored. Confirm it's a Detailed Assessment Customer "
                       "Report `.xlsx`.")
        components.severity_readout(counts)

        buf = io.BytesIO()
        iir_review.build_checklist(data, r['findings'], buf)
        st.download_button(
            "⬇ Findings checklist (.xlsx)", data=buf.getvalue(),
            file_name=f"IIR_Review_{Path(r['src']).stem}.xlsx",
            mime=_XLSX_MIME, key=f"iir_dl_{r['src']}")

        components.findings_table(rows, key=f"iir_{r['src']}")
        _render_evidence(r)


@st.fragment
def _render_evidence(r):
    choice = st.segmented_control(
        "Evidence", ["Protocol", "Extracted data"],
        key=f"iirev_{r['src']}", label_visibility="collapsed")
    if choice == "Protocol":
        _protocol_view(r['data'])
    elif choice == "Extracted data":
        _extracted_view(r['data'])


def _protocol_view(data):
    """The per-position serial-number protocol grid the checks are derived from."""
    rows = data.get('sn_rows') or []
    if not rows:
        st.caption("No serial-number protocol rows were parsed for this report.")
        return
    st.caption(f"{len(rows)} position(s) — the per-position grid the checks are derived from.")
    st.dataframe([{
        "Pos": r.get('pos'), "Part No": r.get('pn', ''), "Serial": r.get('sn', ''),
        "Scope": r.get('scope', ''), "Scrap": "✓" if r.get('scrap') else "",
        "Defects": ", ".join(r.get('defects') or []),
    } for r in rows], width="stretch", hide_index=True,
        column_config={"Defects": st.column_config.TextColumn(width="large")})
    sums = data.get('sn_sumrow') or {}
    if sums:
        st.caption("Protocol sum row — " + " · ".join(f"{k}: **{v}**" for k, v in sums.items()))


def _extracted_view(data):
    """The rich parsed context that used to be Excel-only."""
    shown = False
    rp = data.get('received_parts') or {}
    if rp.get('found'):
        shown = True
        st.markdown("**Received-parts table**")
        st.caption(
            f"Rows **{rp.get('rows', '—')}** · Required **{rp.get('required', '—')}** · "
            f"Received **{rp.get('received', '—')}** · Scrap **{rp.get('scrap', '—')}** · "
            f"Reconditionable **{rp.get('reconditionable', '—')}**")
    ft = data.get('findings_tbl') or {}
    if ft:
        shown = True
        st.markdown("**Summary of damages**")
        st.dataframe([{"Finding": k, "Count": v} for k, v in ft.items()],
                     width="stretch", hide_index=True)
    op = data.get('operating') or {}
    if op:
        shown = True
        st.markdown("**Operating data**")
        st.caption(" · ".join(f"{k}: **{v}**" for k, v in op.items()))
    sm = data.get('spares_matrix') or []
    if sm:
        shown = True
        st.markdown("**Expected replacement components**")
        st.dataframe([{
            "Pos": r.get('pos'), "Scrap": "✓" if r.get('scrap') else "",
            "Components": ", ".join(r.get('comps') or []),
        } for r in sm], width="stretch", hide_index=True,
            column_config={"Components": st.column_config.TextColumn(width="large")})
    sl = data.get('spares_list') or []
    if sl:
        shown = True
        st.markdown("**Spare parts list**")
        st.dataframe([{"Part": r.get('part'), "Qty": r.get('qty')} for r in sl],
                     width="stretch", hide_index=True)
    photos = data.get('photos') or []
    if photos:
        shown = True
        st.markdown("**Photo sheets**")
        st.dataframe([{
            "Sheet": p.get('sheet'), "Captions": len(p.get('captions') or []),
            "Images": p.get('images'),
        } for p in photos], width="stretch", hide_index=True)
    footers = data.get('footers') or []
    if footers:
        shown = True
        with st.expander(f"Page footers ({len(footers)})"):
            frows = []
            for fr in footers:
                fr = list(fr) + [None, None, None]
                frows.append({"Sheet": fr[0], "Page": fr[1], "Of": fr[2]})
            st.dataframe(frows, width="stretch", hide_index=True)
    if not shown:
        st.caption("No additional extracted data available for this report.")
