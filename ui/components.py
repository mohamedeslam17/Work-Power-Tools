"""Shared UI building blocks used by both reviewers (Lab + IIR).

One severity vocabulary, one findings list, one batch table, one report
header — this module exists so the two reviewers say things once instead of
each growing their own near-identical version.
"""
import streamlit as st

# ── severity: one vocabulary for both reviewers ────────────────────────────
# Lab findings arrive as ('critical'|'warning'|'info'|'pass', category, msg).
# IIR findings arrive as {'severity': 'FAIL'|'WARN'|'INFO'|'PASS', ...}. Both
# get normalized to this one set of codes so there is exactly one severity
# model, one readout and one findings table shared by both tools.
ORDER = ('fail', 'warn', 'info', 'pass')
RANK = {s: i for i, s in enumerate(ORDER)}
META = {
    'fail': {'label': 'Fail',    'icon': '🔴', 'color': 'red'},
    'warn': {'label': 'Warning', 'icon': '🟠', 'color': 'orange'},
    'info': {'label': 'Note',    'icon': '🔵', 'color': 'blue'},
    'pass': {'label': 'Pass',    'icon': '🟢', 'color': 'green'},
}
_LAB_SEV = {'critical': 'fail', 'warning': 'warn', 'info': 'info', 'pass': 'pass'}
_IIR_SEV = {'FAIL': 'fail', 'WARN': 'warn', 'INFO': 'info', 'PASS': 'pass'}


def normalize_lab(findings):
    """Lab (severity, category, message) triples -> common finding rows."""
    return [{'severity': _LAB_SEV.get(sev, 'info'), 'category': cat, 'detail': msg}
            for sev, cat, msg in findings]


def normalize_iir(findings, report=None):
    """IIR finding dicts -> common finding rows."""
    rows = []
    for f in findings:
        row = {'severity': _IIR_SEV.get(f['severity'], 'info'), 'category': f['category'],
               'check': f['check'], 'sheet': f['sheet'], 'detail': f['detail']}
        if report is not None:
            row['report'] = report
        rows.append(row)
    return rows


def count_by_severity(rows):
    counts = {s: 0 for s in ORDER}
    for r in rows:
        counts[r['severity']] = counts.get(r['severity'], 0) + 1
    return counts


def verdict(counts):
    """The worst non-zero severity in counts; 'pass' if none."""
    for s in ORDER:
        if counts.get(s):
            return s
    return 'pass'


def status_text(counts):
    """One-line 'icon label' for the worst severity present — for table cells."""
    if not any(counts.values()):
        return "No findings"
    s = verdict(counts)
    return f"{META[s]['icon']} {META[s]['label']}"


def severity_readout(counts):
    """The one severity readout, shown on every report card: a row of native
    badges, worst severity first, zero counts omitted."""
    present = [s for s in ORDER if counts.get(s)]
    if not present:
        st.badge("No findings", icon="⚪", color="gray")
        return
    cols = st.columns(len(present))
    for col, s in zip(cols, present):
        m = META[s]
        col.badge(f"{counts[s]} {m['label']}", icon=m['icon'], color=m['color'])


def report_header(name, tag=None, facts=None):
    """One report header: name, an optional type tag, an optional facts line."""
    st.markdown(f"**{name}**" + (f"  ·  _{tag}_" if tag else ""))
    if facts:
        st.caption(facts)


@st.fragment
def findings_table(rows, key, show_report=False):
    """The one findings list shared by both reviewers.

    rows: common finding dicts (severity, category, detail, and optionally
    check / sheet / report). Severity + category filter and free-text search
    live in this fragment so using them doesn't re-run the whole page (and,
    on a batch screen, doesn't re-loop every report in it).
    """
    if not rows:
        st.caption("No findings.")
        return
    counts = count_by_severity(rows)
    present = [s for s in ORDER if counts[s]]
    default_sev = [s for s in present if s != 'pass'] or present
    cats = sorted({r['category'] for r in rows})

    c1, c2, c3 = st.columns([3, 2, 2])
    picked = c1.pills(
        "Severity", present, selection_mode="multi", default=default_sev,
        format_func=lambda s: f"{META[s]['icon']} {META[s]['label']} ({counts[s]})",
        label_visibility="collapsed", key=f"sev_{key}")
    cat_pick = c2.multiselect(
        "Category", cats, placeholder="All categories",
        label_visibility="collapsed", key=f"cat_{key}")
    query = c3.text_input(
        "Search", placeholder="🔍 search findings…",
        label_visibility="collapsed", key=f"q_{key}")

    picked = picked or present
    q = (query or "").lower().strip()
    shown = [r for r in rows if r['severity'] in picked
             and (not cat_pick or r['category'] in cat_pick)
             and (not q or q in " ".join(str(r.get(k, "")) for k in
                                          ('category', 'check', 'sheet', 'detail', 'report')).lower())]
    if not shown:
        st.caption("No findings match the filter.")
        return
    shown = sorted(shown, key=lambda r: RANK[r['severity']])

    show_check = any(r.get('check') for r in rows)
    show_sheet = any(r.get('sheet') for r in rows)
    table = []
    for r in shown:
        row = {}
        if show_report:
            row["Report"] = r.get('report', '')
        row["Sev"] = f"{META[r['severity']]['icon']} {META[r['severity']]['label']}"
        row["Category"] = r['category']
        if show_check:
            row["Check"] = r.get('check', '')
        if show_sheet:
            row["Sheet"] = r.get('sheet', '')
        row["Detail"] = r['detail']
        table.append(row)
    st.dataframe(table, hide_index=True, width="stretch",
                 column_config={"Detail": st.column_config.TextColumn(width="large")})


def batch_table(items, key, need_label="need attention"):
    """The one batch table shared by both reviewers.

    items: dicts with a 'counts' key (common severity counts) plus whatever
    display columns the caller wants (dict order = column order). Keys
    starting with '_' are carried through but never shown as a column — use
    them to stash a reference back to the caller's full report object.
    Worst-first; click a row to select. Returns (sorted_items, selected_index).
    """
    ranked = sorted(items, key=lambda it: RANK[verdict(it['counts'])])
    need = sum(1 for it in ranked if verdict(it['counts']) != 'pass')
    st.caption(f"**{len(ranked)} reports** — {need} {need_label}  ·  click a row to open it below")
    rows = []
    for it in ranked:
        row = {"Status": status_text(it['counts'])}
        row.update({k: v for k, v in it.items() if k != 'counts' and not k.startswith('_')})
        rows.append(row)
    event = st.dataframe(rows, hide_index=True, width="stretch",
                         on_select="rerun", selection_mode="single-row", key=key)
    idx = event.selection.rows[0] if event and event.selection.rows else 0
    return ranked, idx
