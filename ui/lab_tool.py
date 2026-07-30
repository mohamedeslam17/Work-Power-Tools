"""Lab Report Review — 'Is this .xlsx lab report wrong anywhere, and where?'"""
import io
from pathlib import Path

import streamlit as st
from openpyxl.utils import get_column_letter

import report_render
from lab_review import review_report, collect_highlights
from photo_lib import add_to_library

try:
    from lab_review import COMP_WARN_REL, COMP_CRIT_REL
except Exception:     # display-only tolerance hint; authoritative logic lives in lab_review
    COMP_WARN_REL, COMP_CRIT_REL = 10.0, 25.0

try:
    import pandas as pd
except Exception:     # pandas ships with Streamlit; guard just in case
    pd = None

from ui import components
from ui import photo_tool

_TYPE_LABEL = {'metallurgical': 'Metallurgical', 'coating': 'Coating', 'unknown': 'Unknown type'}


@st.cache_data(show_spinner=False)
def _cached_review(name, data, ocr):
    """Parse + rule-check one report. Cached on the file bytes so filtering
    findings afterwards never re-parses the workbook."""
    return review_report(name, data, ocr=ocr)


@st.cache_data(show_spinner=False)
def _cached_micrographs(name, data, ocr):
    _, parsed, _ = _cached_review(name, data, ocr)
    try:
        return report_render.annotate_micrographs(data, parsed)
    except Exception:
        return []


@st.cache_data(show_spinner=False)
def _cached_faithful(name, data, ocr):
    """The pixel-faithful LibreOffice render. Returns (png_or_None, status)."""
    _, parsed, findings = _cached_review(name, data, ocr)
    try:
        return report_render.render_report_faithful(data, parsed, findings=findings, filename=name)
    except Exception as e:
        return None, f'{type(e).__name__}: {e}'


@st.cache_data(show_spinner=False)
def _ocr_available():
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def _key_facts(rtype, parsed):
    """A compact 'Alloy · Job · Component · S/N' line for the report header."""
    if rtype == 'metallurgical':
        hdr, smp = parsed.get('header', {}) or {}, parsed.get('sample', {}) or {}
        bits = [('Alloy', smp.get('material')), ('Job', hdr.get('job')),
                ('Component', smp.get('description')), ('S/N', smp.get('serial'))]
    elif rtype == 'coating':
        bits = [('Report', parsed.get('report_no')), ('Component', parsed.get('component'))]
    else:
        return ""
    shown = [f"{k}: **{v}**" for k, v in bits if v and str(v).strip()]
    return "  ·  ".join(shown)


def _findings_csv(findings):
    rows = [{'Severity': sev, 'Category': cat, 'Finding': msg} for sev, cat, msg in findings]
    if pd is not None:
        return pd.DataFrame(rows).to_csv(index=False).encode('utf-8')
    import csv
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=['Severity', 'Category', 'Finding'])
    w.writeheader()
    w.writerows(rows)
    return buf.getvalue().encode('utf-8')


def render():
    files = st.file_uploader(
        "Upload lab report(s) to check for issues (.xlsx)",
        type=["xlsx"], accept_multiple_files=True, key="lab_files")
    if not files:
        return

    ocr_ok = _ocr_available()
    ocr = st.toggle(
        "Also read micrograph legends & etch contrast (slower)", value=False,
        disabled=not ocr_ok,
        help="Cross-checks each micrograph's burned-in legend and etch contrast via OCR."
             + ("" if ocr_ok else " Unavailable — Tesseract isn't installed in this environment."))

    reviewed = []
    for f in files:
        try:
            with st.spinner(f"Reviewing {f.name}…"):
                rtype, parsed, findings = _cached_review(f.name, f.getvalue(), ocr)
        except Exception as e:
            reviewed.append({'name': f.name, 'error': str(e)})
            continue
        rows = components.normalize_lab(findings)
        reviewed.append({
            'f': f, 'name': f.name, 'rtype': rtype, 'parsed': parsed, 'findings': findings,
            'rows': rows, 'counts': components.count_by_severity(rows),
            'facts': _key_facts(rtype, parsed),
        })

    for r in [x for x in reviewed if 'error' in x]:
        st.error(f"Could not read **{r['name']}** — {r['error']}")
    ok = [r for r in reviewed if 'error' not in r]
    if not ok:
        return

    if len(ok) > 1:
        items = [{'counts': r['counts'], 'Report': r['name'], 'Type': _TYPE_LABEL[r['rtype']],
                  '_ref': r} for r in ok]
        ranked, idx = components.batch_table(items, key="lab_batch")

        allrows = [{'Report': r['name'], 'Severity': row['severity'],
                    'Category': row['category'], 'Finding': row['detail']}
                   for r in ok for row in r['rows']]
        if pd is not None and allrows:
            st.download_button(
                "⬇ All findings (.csv)",
                data=pd.DataFrame(allrows).to_csv(index=False).encode('utf-8'),
                file_name="lab_review_all_findings.csv", mime="text/csv", key="lab_batch_csv")
        selected = ranked[idx]['_ref']
    else:
        selected = ok[0]

    _render_detail(selected, ocr)


def _render_detail(r, ocr):
    with st.container(border=True):
        components.report_header(r['name'], tag=_TYPE_LABEL[r['rtype']], facts=r['facts'])
        if r['rtype'] == 'unknown':
            st.warning("This workbook didn't match a metallurgical or coating layout, so only "
                       "a limited review ran. Check it's an AEG lab report `.xlsx`.")
        components.severity_readout(r['counts'])

        with st.popover("Actions"):
            if r['rtype'] in ('metallurgical', 'coating'):
                if st.button("📁 Add to photo library", key=f"add_{r['name']}", width="stretch"):
                    added = add_to_library(r['name'], r['f'].getvalue(), r['parsed'], r['rtype'])
                    if added:
                        photo_tool.invalidate()
                    st.toast(f"Added {added} micrograph(s) to the library." if added
                             else "No new micrographs (already in library).")
            st.download_button(
                "⬇ Findings (.csv)", data=_findings_csv(r['findings']),
                file_name=f"{Path(r['name']).stem}_findings.csv", mime="text/csv",
                key=f"labcsv_{r['name']}", width="stretch")

        components.findings_table(r['rows'], key=f"lab_{r['name']}")
        _render_evidence(r, ocr)


@st.fragment
def _render_evidence(r, ocr):
    choice = st.segmented_control(
        "Evidence", ["Annotated view", "Extracted data"],
        key=f"labev_{r['name']}", label_visibility="collapsed")
    if choice == "Annotated view":
        _annotated_view(r, ocr)
    elif choice == "Extracted data":
        _extracted_view(r['rtype'], r['parsed'])


def _annotated_view(r, ocr):
    """Pixel-faithful LibreOffice render + annotated micrographs — built only
    now, on demand, so opening this is what pays the LibreOffice/OCR cost."""
    f, rtype, parsed = r['f'], r['rtype'], r['parsed']
    data = f.getvalue()

    if report_render.libreoffice_available():
        with st.status("Rendering the exact workbook with LibreOffice…", expanded=False) as status:
            png, stat = _cached_faithful(f.name, data, ocr)
            status.update(
                label=("Pixel-faithful render ready" if png
                       else f"Pixel-faithful render unavailable — {stat}"),
                state=("complete" if png else "error"))
        if png:
            st.image(png, width="stretch",
                     caption="Pixel-faithful render — original fonts, layout and embedded "
                             "micrographs, with flagged cells boxed and numbered to the legend.")
            st.download_button(
                "⬇ Annotated view (.png)", data=png,
                file_name=f"{Path(f.name).stem}_annotated.png",
                mime="image/png", key=f"fpng_{f.name}")
    elif rtype in ('metallurgical', 'coating'):
        st.caption("Annotated view unavailable — LibreOffice isn't installed in this environment.")

    with st.spinner("Reading micrographs…"):
        micrographs = _cached_micrographs(f.name, data, ocr)
    if micrographs:
        st.markdown("**Annotated micrographs** — legend / scale-bar regions boxed; "
                    "contrast and any burned-in thickness flagged.")
        cols = st.columns(3)
        for i, (mname, mbytes, mcap) in enumerate(micrographs):
            cols[i % 3].image(mbytes, caption=mcap, width="stretch")


def _extracted_view(rtype, parsed):
    """The facts the reviewer extracted, for transparency."""
    if rtype == 'metallurgical':
        hdr = parsed.get('header', {})
        smp = parsed.get('sample', {})
        c1, c2 = st.columns(2)
        c1.write(f"**Job:** {hdr.get('job') or '—'}")
        c1.write(f"**Customer:** {hdr.get('customer') or '—'}")
        c1.write(f"**Machine:** {hdr.get('machine') or '—'}")
        c2.write(f"**Material:** {smp.get('material') or '—'}")
        c2.write(f"**Description:** {smp.get('description') or '—'}")
        c2.write(f"**S/N:** {smp.get('serial') or '—'}")
        coat = parsed.get('coating') or {}
        coat_str = (coat.get('type') or coat.get('received') or coat.get('outgoing')
                    or coat.get('present') or '—')
        c1.write(f"**Coating:** {coat_str}")

        nom, act = parsed.get('nominal', {}), parsed.get('actual', {})
        if nom or act:
            st.markdown("**Composition — Nominal vs Actual (wt%)**")
            rows = []
            for el in sorted(set(nom) | set(act)):
                n, a = nom.get(el), act.get(el)
                if n not in (None, 0) and a is not None:
                    dev_pct = (a - n) / abs(n) * 100
                    dev = f"{dev_pct:+.0f}%"
                    flag = ("🔴" if abs(dev_pct) >= COMP_CRIT_REL
                            else "🟠" if abs(dev_pct) >= COMP_WARN_REL else "")
                else:
                    dev, flag = "—", ""
                rows.append({
                    "": flag, "Element": el,
                    "Nominal": f"{n:g}" if n is not None else "—",
                    "Actual":  f"{a:g}" if a is not None else "—",
                    "Δ":       dev,
                })
            st.dataframe(rows, width="stretch", hide_index=True,
                         column_config={"": st.column_config.TextColumn(width="small")})
            st.caption("Colour is an at-a-glance deviation hint (🟠 ≥ %g%%, 🔴 ≥ %g%%); "
                       "the findings above hold the authoritative result."
                       % (COMP_WARN_REL, COMP_CRIT_REL))

    elif rtype == 'coating':
        st.write(f"**Report No.:** {parsed.get('report_no') or '—'}")
        st.write(f"**Title:** {parsed.get('title') or '—'}")
        rows = parsed.get('rows', [])
        if rows:
            lo, hi = rows[0].get('min'), rows[0].get('max')
            if lo is not None and hi is not None:
                st.write(f"**Design limit:** {lo:g} – {hi:g} mm")
            st.dataframe([
                {"Row": e['row'], "Measurements (mm)": ", ".join(f"{v:g}" for v in e['values'])}
                for e in rows
            ], width="stretch", hide_index=True,
                column_config={"Measurements (mm)": st.column_config.TextColumn(width="large")})

    legends = parsed.get('legends') or []
    if legends:
        st.markdown("**Micrograph legends — read from the images**")
        st.dataframe([
            {"Image": l.get('image', '—'), "Magnification": l.get('mag', '—'),
             "Scale": l.get('scale', '—'), "Legend ID": l.get('id', '—')}
            for l in legends
        ], width="stretch", hide_index=True)

    _flagged_cells(parsed)


def _flagged_cells(parsed):
    """Cell-anchored findings (from collect_highlights) as spreadsheet refs."""
    try:
        hi = collect_highlights(parsed)
    except Exception:
        return
    if not hi:
        return
    rows = []
    for h in hi:
        cell = h.get('cell') or (None, None)
        row, col = (cell + (None, None))[:2]
        ref = f"{get_column_letter(col)}{row}" if (row and col) else "—"
        rows.append({'Cell': ref, 'Severity': h.get('severity') or '',
                     'Category': h.get('category', ''), 'Note': h.get('note', '')})
    with st.expander(f"📍 Flagged cells ({len(rows)})"):
        st.caption("Where each cell-anchored finding sits in the workbook — matches the "
                   "boxed & numbered cells in the annotated view.")
        st.dataframe(rows, width="stretch", hide_index=True,
                     column_config={"Note": st.column_config.TextColumn(width="large")})
