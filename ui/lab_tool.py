"""Lab Report Review — 'Is this .xlsx lab report wrong anywhere, and where?'"""
import io
from pathlib import Path

import streamlit as st

import report_render
from lab_review import review_report
from photo_lib import add_to_library

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
def _cached_annotated_report(name, data, ocr):
    """Return one report-centric review canvas.

    The exact LibreOffice rendering is preferred. A spreadsheet-style
    annotated fallback keeps the report visible if exact rendering is not
    available in the deployment environment.
    """
    rtype, parsed, findings = _cached_review(name, data, ocr)
    exact_status = 'LibreOffice not installed'
    if report_render.libreoffice_available():
        try:
            png, exact_status = report_render.render_report_faithful(
                data, parsed, findings=findings, filename=name)
            if png:
                return png, 'exact'
        except Exception as e:
            exact_status = f'{type(e).__name__}: {e}'

    try:
        png = report_render.render_report_image(
            data, parsed, findings=findings, rtype=rtype, filename=name)
    except Exception as e:
        return None, f'{exact_status}; fallback failed: {type(e).__name__}: {e}'
    return (png, f'fallback: {exact_status}') if png else (
        None, f'{exact_status}; annotated fallback unavailable')


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
        "Read micrograph identity, legends & etch contrast (slower)", value=ocr_ok,
        disabled=not ocr_ok,
        help="Cross-checks every micrograph's burned-in job number, magnification, "
             "scale and etch contrast via OCR."
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
        ranked = sorted(ok, key=lambda r: components.RANK[components.verdict(r['counts'])])
        idx = st.selectbox(
            "Report",
            range(len(ranked)),
            format_func=lambda i: (
                f"{components.status_text(ranked[i]['counts'])} · {ranked[i]['name']}"),
            key="lab_report_selector")
        selected = ranked[idx]
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

        _annotated_report(r, ocr)


def _annotated_report(r, ocr):
    """Make the report itself the review UI.

    The image contains the original sheet, highlighted issue locations,
    numbered markers and the matching explanations. No separate findings
    table or evidence-mode selection is required.
    """
    f = r['f']
    with st.spinner(f"Building annotated report for {f.name}…"):
        png, mode = _cached_annotated_report(f.name, f.getvalue(), ocr)

    if not png:
        st.error(f"Could not build the annotated report view — {mode}.")
        return

    if mode != 'exact':
        st.warning(
            "The exact Excel renderer is unavailable, so this is a simplified "
            "annotated sheet view. Issue markers and explanations are still included.")

    st.image(
        png,
        width="stretch",
        caption=("Annotated report — numbered markers identify the affected locations; "
                 "their explanations are included in the same image."))
    st.download_button(
        "⬇ Download annotated report (.png)",
        data=png,
        file_name=f"{Path(f.name).stem}_annotated.png",
        mime="image/png",
        key=f"fpng_{f.name}",
        width="stretch")
