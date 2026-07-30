"""Lab Report Review — 'Is this .xlsx lab report wrong anywhere, and where?'"""
import html
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
    """Return a page-oriented report-centric review package.

    The exact LibreOffice rendering is preferred. A spreadsheet-style
    annotated fallback keeps the report visible if exact rendering is not
    available in the deployment environment.
    """
    rtype, parsed, findings = _cached_review(name, data, ocr)
    exact_status = 'LibreOffice not installed'
    if report_render.libreoffice_available():
        try:
            view, exact_status = report_render.render_report_faithful_view(
                data, parsed, findings=findings, filename=name)
            if view:
                return view, 'exact'
        except Exception as e:
            exact_status = f'{type(e).__name__}: {e}'

    try:
        png = report_render.render_report_image(
            data, parsed, findings=findings, rtype=rtype, filename=name)
    except Exception as e:
        return None, f'{exact_status}; fallback failed: {type(e).__name__}: {e}'
    if not png:
        return None, f'{exact_status}; annotated fallback unavailable'
    issues, extras = report_render.build_issue_index(parsed, findings)
    for issue in issues:
        issue['pages'] = [1]
    return {
        'filename': name,
        'pages': [{'number': 1, 'png': png,
                   'issue_nums': [issue['num'] for issue in issues]}],
        'issues': issues,
        'extras': extras,
        'annotated_pdf': None,
        'combined_png': png,
    }, f'fallback: {exact_status}'


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
        view, mode = _cached_annotated_report(f.name, f.getvalue(), ocr)

    if not view:
        st.error(f"Could not build the annotated report view — {mode}.")
        return

    if mode != 'exact':
        st.warning(
            "The exact Excel renderer is unavailable, so this is a simplified "
            "annotated sheet view. Issue markers and explanations are still included.")

    pages = view.get('pages') or []
    if not pages:
        st.error("The workbook rendered, but no report pages were produced.")
        return

    st.markdown("### Annotated report")
    st.caption(
        "Open one page at a time. Matching numbers can appear in more than one "
        "location when a single issue affects several report fields.")

    labels = {
        page['number']: (
            f"Page {page['number']}"
            + (f" · {len(page.get('issue_nums') or [])} issue"
               f"{'s' if len(page.get('issue_nums') or []) != 1 else ''}"
               if page.get('issue_nums') else " · clear")
        )
        for page in pages
    }
    numbers = [page['number'] for page in pages]
    if len(numbers) <= 6:
        selected_number = st.pills(
            "Report page",
            numbers,
            default=numbers[0],
            selection_mode="single",
            format_func=lambda number: labels[number],
            key=f"lab_page_{f.name}",
            label_visibility="collapsed",
        ) or numbers[0]
    else:
        selected_number = st.selectbox(
            "Report page",
            numbers,
            format_func=lambda number: labels[number],
            key=f"lab_page_{f.name}",
        )
    page = next(item for item in pages if item['number'] == selected_number)

    report_col, issue_col = st.columns([3.25, 1.35], gap="large")
    with report_col:
        st.markdown(
            f'<div class="aeg-page-kicker">PAGE {page["number"]} OF {len(pages)}</div>',
            unsafe_allow_html=True,
        )
        st.image(page['png'], width="stretch")
    with issue_col:
        _page_findings(view, page)

    download_col, _ = st.columns([1.7, 3.3])
    with download_col:
        if view.get('annotated_pdf'):
            st.download_button(
                "⬇ Download annotated report (.pdf)",
                data=view['annotated_pdf'],
                file_name=f"{Path(f.name).stem}_annotated.pdf",
                mime="application/pdf",
                key=f"fpdf_{f.name}",
                width="stretch",
            )
        else:
            st.download_button(
                "⬇ Download annotated report (.png)",
                data=view['combined_png'],
                file_name=f"{Path(f.name).stem}_annotated.png",
                mime="image/png",
                key=f"fpng_{f.name}",
                width="stretch",
            )


def _page_findings(view, page):
    """Render concise issue cards beside the selected report page."""
    page_numbers = set(page.get('issue_nums') or [])
    page_issues = [
        issue for issue in view.get('issues') or []
        if issue['num'] in page_numbers
    ]
    st.markdown("#### Issues on this page")
    if page_issues:
        for issue in page_issues:
            _issue_card(issue)
    else:
        st.markdown(
            '<div class="aeg-clear-card">'
            '<div class="aeg-clear-title">No marked issues</div>'
            '<div class="aeg-clear-copy">Nothing on this page needs a location marker.</div>'
            '</div>',
            unsafe_allow_html=True,
        )

    unplaced = [
        issue for issue in view.get('issues') or []
        if not issue.get('pages')
    ]
    extras = list(view.get('extras') or [])
    if unplaced or extras:
        st.markdown("#### Report-level checks")
        for issue in unplaced:
            _issue_card(issue, show_number=False)
        for extra in extras:
            _issue_card(extra, show_number=False)


def _issue_card(issue, show_number=True):
    severity = issue.get('severity', 'warning')
    label = {
        'critical': 'Fail',
        'warning': 'Warning',
        'info': 'Note',
        'pass': 'Pass',
    }.get(severity, 'Review')
    number = (
        f'<span class="aeg-issue-number">{issue.get("num")}</span>'
        if show_number and issue.get('num') is not None else ''
    )
    refs = issue.get('refs') or []
    meta = f"{html.escape(issue.get('category') or 'Review')}"
    if refs:
        meta += " · " + html.escape(", ".join(refs))
    st.markdown(
        f'<div class="aeg-issue-card aeg-{severity}">'
        f'<div class="aeg-issue-head">{number}'
        f'<span class="aeg-issue-label">{html.escape(label)}</span></div>'
        f'<div class="aeg-issue-meta">{meta}</div>'
        f'<div class="aeg-issue-copy">{html.escape(issue.get("note") or "")}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
