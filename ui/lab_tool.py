"""Lab Report Review — 'Is this report releasable, and if not, where exactly?'"""
import html
import io
from pathlib import Path

import streamlit as st

import report_render
from lab_review import add_version_findings, finding_stem, partition_by_scope, review_report
from photo_lib import add_to_library

try:
    from lab_review import COMP_WARN_REL, COMP_CRIT_REL
except Exception:     # display-only hint; authoritative logic lives in lab_review
    COMP_WARN_REL, COMP_CRIT_REL = 10.0, 25.0

try:
    import pandas as pd
except Exception:     # pandas ships with Streamlit; guard just in case
    pd = None

try:
    from PIL import Image
except Exception:
    Image = None

from ui import components
from ui import photo_tool

_TYPE_LABEL = {'metallurgical': 'Metallurgical', 'coating': 'Coating', 'unknown': 'Unknown type'}


@st.cache_data(show_spinner=False)
def _cached_review(name, data, ocr):
    """Parse + rule-check one report. Cached on the file bytes so filtering
    findings afterwards never re-parses the workbook."""
    return review_report(name, data, ocr=ocr)


@st.cache_data(show_spinner=False)
def _cached_annotated_report(name, data, ocr, extra_findings=()):
    """Return a page-oriented report-centric review package.

    The exact LibreOffice rendering is preferred. A spreadsheet-style
    annotated fallback keeps the report visible if exact rendering is not
    available in the deployment environment.

    Deliberately built from EVERY finding, not just the ones triage state
    currently wants shown — this stays cached on (name, data, ocr) alone so
    dismissing or restoring a finding never re-runs LibreOffice/Pillow. The
    triage state (scope partition, accept/dismiss) is applied afterwards, as
    a cheap filter over the already-rendered issue list; see `_render_detail`.
    """
    rtype, parsed, findings = _cached_review(name, data, ocr)
    findings = list(findings) + list(extra_findings or ())
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


# ── Triage state — session-local, never cached, keyed on (file, category, stem)
# so it survives a message's variable detail (an element list, a page label)
# changing between reviews of the same report family. ─────────────────────
def _dismissed_store():
    return st.session_state.setdefault('lab_dismissed', {})   # {key: reason}


def _accepted_store():
    return st.session_state.setdefault('lab_accepted', set())  # {key}


def _tri_key(name, category, message):
    return (name, category, finding_stem(message))


def _dismiss(name, category, message, reason):
    _dismissed_store()[_tri_key(name, category, message)] = reason.strip() or '(no reason given)'


def _restore(name, category, message):
    _dismissed_store().pop(_tri_key(name, category, message), None)


def _is_dismissed(name, category, message):
    return _tri_key(name, category, message) in _dismissed_store()


def _toggle_accept(name, category, message):
    key = _tri_key(name, category, message)
    store = _accepted_store()
    (store.discard if key in store else store.add)(key)


def _is_accepted(name, category, message):
    return _tri_key(name, category, message) in _accepted_store()


def _triaged_findings(name, findings):
    """Partition + dismiss-filter one report's raw findings.

    Returns (active, dismissed, template). `active` is what drives the
    verdict and the severity counts; `dismissed` is restorable and never
    counted; `template` is said once, in its own section, and — the whole
    point of D11 — never enters the verdict either.
    """
    report, template = partition_by_scope(findings)
    active, dismissed = [], []
    for f in report:
        _sev, cat, msg = f
        (dismissed if _is_dismissed(name, cat, msg) else active).append(f)
    return active, dismissed, template


def _verdict(active):
    """(tier, label, reason) — the one line at the top of a report."""
    criticals = [f for f in active if f[0] == 'critical']
    warnings = [f for f in active if f[0] == 'warning']
    if criticals:
        reason = criticals[0][2]
        if len(criticals) > 1:
            more = len(criticals) - 1
            reason += f' (+{more} more critical finding{"s" if more != 1 else ""}.)'
        return 'hold', 'HOLD — do not release', reason
    if warnings:
        n = len(warnings)
        return 'correction', 'NEEDS CORRECTION', (
            f'{n} warning{"s" if n != 1 else ""} should be resolved or dismissed with a '
            f'reason before release.')
    return 'release', 'RELEASE', 'No unresolved findings block release.'


def _verdict_banner(tier, label, reason):
    st.markdown(
        f'<div class="aeg-verdict aeg-verdict-{tier}">'
        f'<span class="aeg-verdict-label">{html.escape(label)}</span>'
        f'<span class="aeg-verdict-reason">{html.escape(reason)}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )


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
        reviewed.append({
            'f': f, 'name': f.name, 'rtype': rtype, 'parsed': parsed, 'findings': findings,
            'facts': _key_facts(rtype, parsed),
        })

    add_version_findings(reviewed)
    for report in reviewed:
        if 'error' in report:
            continue
        active, dismissed, template = _triaged_findings(report['name'], report['findings'])
        report['active'], report['dismissed'], report['template'] = active, dismissed, template
        report['rows'] = components.normalize_lab(report['findings'])
        report['counts'] = components.count_by_severity(components.normalize_lab(active))

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
    name = r['name']
    with st.container(border=True):
        components.report_header(name, tag=_TYPE_LABEL[r['rtype']], facts=r['facts'])
        if r['rtype'] == 'unknown':
            st.warning("This workbook didn't match a metallurgical or coating layout, so only "
                       "a limited review ran. Check it's an AEG lab report `.xlsx`.")

        tier, label, reason = _verdict(r['active'])
        _verdict_banner(tier, label, reason)
        components.severity_readout(r['counts'])

        if r['dismissed']:
            with st.expander(f"🗑 {len(r['dismissed'])} dismissed — click to restore"):
                _dismissed_list(name, r['dismissed'])

        with st.popover("Actions"):
            if r['rtype'] in ('metallurgical', 'coating'):
                if st.button("📁 Add to photo library", key=f"add_{name}", width="stretch"):
                    added = add_to_library(name, r['f'].getvalue(), r['parsed'], r['rtype'])
                    if added:
                        photo_tool.invalidate()
                    st.toast(f"Added {added} micrograph(s) to the library." if added
                             else "No new micrographs (already in library).")
            st.download_button(
                "⬇ Findings (.csv)", data=_findings_csv(r['findings']),
                file_name=f"{Path(name).stem}_findings.csv", mime="text/csv",
                key=f"labcsv_{name}", width="stretch")

        with st.expander("Extracted information", expanded=True):
            _extracted_view(r['rtype'], r['parsed'])

        _template_section(name, r['template'])

        _annotated_report(r, ocr)


def _dismissed_list(name, dismissed):
    for sev, cat, msg in dismissed:
        reason = _dismissed_store().get(_tri_key(name, cat, msg), '')
        cols = st.columns([5, 1], vertical_alignment="center")
        cols[0].markdown(f"**{cat}** — {msg}  \n*Dismissed: {html.escape(reason)}*")
        if cols[1].button("Restore", key=f"restore_{name}_{cat}_{finding_stem(msg)}",
                          width="stretch"):
            _restore(name, cat, msg)
            st.rerun()


def _template_section(name, template_findings):
    """D11's other half: template-scoped findings, said once, collapsed."""
    if not template_findings:
        return
    with st.expander(f"About this template ({len(template_findings)})"):
        st.markdown(
            '<div class="aeg-template-note">These describe a gap in the AEG report '
            'template itself — every report on this template shows the same items — '
            'not something specific to this report, so they never affect the verdict '
            'above and are listed once here instead of repeating per report.</div>',
            unsafe_allow_html=True)
        components.findings_table(
            components.normalize_lab(template_findings), key=f"template_{name}")


def _shown(value):
    return value if value is not None and str(value).strip() else '—'


def _field_html(label, field):
    """One 'Label: value' line, with a status badge distinguishing a field the
    parser never located from one that is genuinely blank (Phase 1's D1 fix,
    finally visible) — a `not_located` field looks nothing like an `empty`
    one, on purpose. Falls back to a plain value where no status is tracked
    (only the seven header fields carry one so far)."""
    field = field or {}
    status = field.get('status')
    if status == 'found':
        return f"<b>{html.escape(label)}:</b> {html.escape(str(field.get('value')))}"
    if status == 'empty':
        return (f"<b>{html.escape(label)}:</b> —"
                '<span class="aeg-field-status aeg-field-empty">blank in report</span>')
    if status == 'not_located':
        return (f"<b>{html.escape(label)}:</b> "
                '<span class="aeg-field-status aeg-field-notlocated">⚠ not read</span>')
    return f"<b>{html.escape(label)}:</b> {html.escape(str(_shown(field.get('value'))))}"


def _hardness_text(entry):
    entry = entry or {}
    raw = entry.get('raw')
    if raw is not None and str(raw).strip():
        return str(raw)
    value = entry.get('value')
    if value is None:
        return '—'
    return f"{value:g} {entry.get('unit') or ''}".strip()


def _extracted_view(rtype, parsed):
    """Show the facts used by the reviewer without replacing the report view."""
    st.caption(
        "Values read from the workbook. The annotated report remains the main "
        "review view below.")

    if rtype == 'metallurgical':
        fields = parsed.get('fields') or {}
        sample = parsed.get('sample') or {}
        hardness = parsed.get('hardness') or {}
        coating = parsed.get('coating') or {}
        left, right = st.columns(2, gap="large")
        with left:
            st.markdown("**Report identity**")
            if fields:
                st.caption("A ⚠ badge means the field's label wasn't found on the "
                           "sheet — an extraction gap, not a blank report field.")
            lines = [
                _field_html('Job', fields.get('job')),
                _field_html('AEG reference', fields.get('aeg_ref')),
                _field_html('Customer', fields.get('customer')),
                _field_html('Customer reference', fields.get('customer_ref')),
                _field_html('Machine', fields.get('machine')),
                _field_html('Quantity', fields.get('qty')),
                _field_html('EOH', fields.get('eoh')),
            ]
            st.markdown("<br>".join(lines), unsafe_allow_html=True)
        with right:
            st.markdown("**Sample identity**")
            st.write(f"**Sample number:** {_shown(sample.get('sample_no'))}")
            st.write(f"**Component:** {_shown(sample.get('description'))}")
            st.write(f"**Serial number:** {_shown(sample.get('serial'))}")
            st.write(f"**Location:** {_shown(sample.get('location'))}")
            st.write(f"**Material:** {_shown(sample.get('material'))}")
            st.write(f"**Result:** {_shown(sample.get('result'))}")

        samples = parsed.get('samples') or []
        if len(samples) > 1:
            st.caption(
                f"{len(samples)} samples on this report (shown: the first). "
                f"See the annotated report below for findings on the others.")

        coat_text = next(
            (coating.get(key) for key in ('type', 'received', 'outgoing', 'present')
             if coating.get(key) is not None and str(coating.get(key)).strip()),
            '—',
        )
        details_left, details_right = st.columns(2, gap="large")
        with details_left:
            st.markdown("**Test details**")
            st.write(f"**Pre-solution hardness:** {_hardness_text(hardness.get('pre'))}")
            st.write(f"**Post-solution hardness:** {_hardness_text(hardness.get('post'))}")
            st.write(f"**Coating:** {coat_text}")
        with details_right:
            st.markdown("**Evidence extracted**")
            caption_count = len(parsed.get('pictures') or [])
            image_count = parsed.get('micrograph_count')
            st.write(f"**Picture captions:** {caption_count}")
            st.write(
                f"**Embedded micrographs:** "
                f"{_shown(image_count) if image_count is not None else '—'}")

        comment = parsed.get('comment')
        if comment is not None and str(comment).strip():
            st.markdown("**Reported comment**")
            st.write(str(comment))

        nominal = parsed.get('nominal') or {}
        actual = parsed.get('actual') or {}
        actual_entries = (
            ((parsed.get('composition_meta') or {}).get('actual') or {})
            .get('entries') or []
        )
        actual_raw = {}
        for entry in actual_entries:
            element = entry.get('element')
            if element:
                actual_raw.setdefault(element, []).append(entry.get('raw') or '—')
        if nominal or actual or actual_raw:
            st.markdown("**Composition — nominal vs actual (wt%)**")
            rows = []
            for element in sorted(set(nominal) | set(actual) | set(actual_raw)):
                expected, measured = nominal.get(element), actual.get(element)
                if expected not in (None, 0) and measured is not None:
                    deviation_pct = (measured - expected) / abs(expected) * 100
                    deviation = f"{deviation_pct:+.0f}%"
                    flag = (
                        "🔴" if abs(deviation_pct) >= COMP_CRIT_REL
                        else "🟠" if abs(deviation_pct) >= COMP_WARN_REL
                        else ""
                    )
                else:
                    deviation, flag = "—", ""
                rows.append({
                    "": flag,
                    "Element": element,
                    "Nominal": f"{expected:g}" if expected is not None else "—",
                    "Actual": (
                        f"{measured:g}" if measured is not None
                        else " / ".join(actual_raw.get(element) or ["—"])
                    ),
                    "Δ": deviation,
                })
            st.dataframe(
                rows,
                width="stretch",
                hide_index=True,
                column_config={"": st.column_config.TextColumn(width="small")},
            )
            st.caption(
                "Deviation hint only: 🟠 ≥ %g%% and 🔴 ≥ %g%%. The numbered "
                "review comments remain authoritative."
                % (COMP_WARN_REL, COMP_CRIT_REL)
            )

    elif rtype == 'coating':
        left, right = st.columns(2, gap="large")
        with left:
            st.write(f"**Report number:** {_shown(parsed.get('report_no'))}")
            st.write(f"**Title:** {_shown(parsed.get('title'))}")
        with right:
            st.write(f"**Component:** {_shown(parsed.get('component'))}")
            st.write(f"**Embedded images:** {_shown(parsed.get('media'))}")

        rows = parsed.get('rows') or []
        if rows:
            st.markdown("**Coating measurements**")
            st.dataframe([
                {
                    "Workbook row": entry.get('row'),
                    "Design minimum (mm)": entry.get('min'),
                    "Design maximum (mm)": entry.get('max'),
                    "Measurements (mm)": ", ".join(
                        f"{value:g}" for value in entry.get('values') or []),
                }
                for entry in rows
            ], width="stretch", hide_index=True)
    else:
        st.info("Only limited information could be extracted from this workbook.")

    legends = parsed.get('legends') or []
    if legends:
        st.markdown("**Micrograph legends read from images**")
        st.dataframe([
            {
                "Image": legend.get('image', '—'),
                "Magnification": legend.get('mag', '—'),
                "Scale": legend.get('scale', '—'),
                "Legend ID": legend.get('id', '—'),
            }
            for legend in legends
        ], width="stretch", hide_index=True)


def _cropped_detail(data, parsed, page_png, cell, pad=44, min_side=280):
    """A zoomed PNG crop around `cell`, or None if it can't be located.

    Only valid for the 'fallback' annotated view — its coordinate space is
    the custom grid `report_render.render_report_image` draws, which
    `cell_pixel_rect` mirrors exactly. The 'exact' (LibreOffice) pages are a
    different rendering pipeline entirely, so this must never be called
    against them.
    """
    if Image is None:
        return None
    rect = report_render.cell_pixel_rect(data, parsed, cell)
    if not rect:
        return None
    try:
        img = Image.open(io.BytesIO(page_png))
    except Exception:
        return None
    x0, y0, x1, y1 = rect
    cx0, cy0 = max(0, x0 - pad), max(0, y0 - pad)
    cx1, cy1 = min(img.width, x1 + pad), min(img.height, y1 + pad)
    if cx1 - cx0 < min_side:
        extra = (min_side - (cx1 - cx0)) // 2
        cx0, cx1 = max(0, cx0 - extra), min(img.width, cx1 + extra)
    if cy1 - cy0 < min_side:
        extra = (min_side - (cy1 - cy0)) // 2
        cy0, cy1 = max(0, cy0 - extra), min(img.height, cy1 + extra)
    if cx1 <= cx0 or cy1 <= cy0:
        return None
    buf = io.BytesIO()
    img.crop((cx0, cy0, cx1, cy1)).save(buf, format='PNG')
    return buf.getvalue()


def _annotated_report(r, ocr):
    """Make the report itself the review UI.

    The image contains the original sheet, highlighted issue locations,
    numbered markers and the matching explanations. Findings and the report
    picture are linked both ways: opening a finding jumps the view to its
    page and (in fallback mode) zooms its cell; the list itself only ever
    shows findings the active triage state (scope + dismiss) still counts.
    """
    f = r['f']
    name = f.name
    with st.spinner(f"Building annotated report for {name}…"):
        view, mode = _cached_annotated_report(
            name, f.getvalue(), ocr, tuple(r.get('version_findings') or ()))

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
    omitted = view.get('omitted_blank_pages') or []
    page_note = (
        f" {len(omitted)} blank trailing source page(s) were omitted: "
        f"{', '.join(map(str, omitted))}."
        if omitted else ""
    )
    st.caption(
        "Open one effective report page at a time. Click a finding to jump the "
        f"view to it. Dismissed and template-level findings are hidden here — see "
        f"the sections above.{page_note}")

    active_keys = {(cat, msg) for _sev, cat, msg in r['active']}
    all_issues = [issue for issue in (view.get('issues') or [])
                  if (issue['category'], issue['note']) in active_keys]
    all_extras = [extra for extra in (view.get('extras') or [])
                  if (extra['category'], extra['note']) in active_keys]
    visible_nums = {issue['num'] for issue in all_issues}

    def _count_on(page):
        return len([n for n in (page.get('issue_nums') or []) if n in visible_nums])

    labels = {
        page['number']: (
            f"Page {page['number']}"
            + (f" · {_count_on(page)} issue{'s' if _count_on(page) != 1 else ''}"
               if _count_on(page) else " · clear")
        )
        for page in pages
    }
    numbers = [page['number'] for page in pages]

    focus_key = f'lab_focus_{name}'
    focus_num = st.session_state.get(focus_key)
    focused = next((i for i in all_issues if i['num'] == focus_num), None)
    if focused and focused.get('pages') and st.session_state.get(f'lab_page_{name}') not in focused['pages']:
        st.session_state[f'lab_page_{name}'] = focused['pages'][0]

    if len(numbers) <= 6:
        selected_number = st.pills(
            "Report page",
            numbers,
            default=numbers[0],
            selection_mode="single",
            format_func=lambda number: labels[number],
            key=f"lab_page_{name}",
            label_visibility="collapsed",
        ) or numbers[0]
    else:
        selected_number = st.selectbox(
            "Report page",
            numbers,
            format_func=lambda number: labels[number],
            key=f"lab_page_{name}",
        )
    page = next(item for item in pages if item['number'] == selected_number)

    report_col, issue_col = st.columns([3.25, 1.35], gap="large")
    with report_col:
        st.markdown(
            f'<div class="aeg-page-kicker">PAGE {page["number"]} OF {len(pages)}</div>',
            unsafe_allow_html=True,
        )
        if mode.startswith('fallback') and focused and focused.get('cells'):
            crop = _cropped_detail(f.getvalue(), r['parsed'], page['png'], focused['cells'][0])
            if crop:
                refs = ", ".join(focused.get('refs') or [])
                st.caption(f"🔍 Zoomed to issue #{focused['num']}" + (f" · {refs}" if refs else ""))
                st.image(crop, width="stretch")
        st.image(page['png'], width="stretch")
    with issue_col:
        page_numbers = set(page.get('issue_nums') or [])
        page_issues = [i for i in all_issues if i['num'] in page_numbers]
        unplaced = [i for i in all_issues if not i.get('pages')]
        _page_findings(name, focus_key, page_issues, unplaced + all_extras)

    download_col, _ = st.columns([1.7, 3.3])
    with download_col:
        if view.get('annotated_pdf'):
            st.download_button(
                "⬇ Download annotated report (.pdf)",
                data=view['annotated_pdf'],
                file_name=f"{Path(name).stem}_annotated.pdf",
                mime="application/pdf",
                key=f"fpdf_{name}",
                width="stretch",
            )
        else:
            st.download_button(
                "⬇ Download annotated report (.png)",
                data=view['combined_png'],
                file_name=f"{Path(name).stem}_annotated.png",
                mime="image/png",
                key=f"fpng_{name}",
                width="stretch",
            )


def _page_findings(name, focus_key, page_issues, report_level):
    """Render concise issue cards beside the selected report page."""
    st.markdown("#### Issues on this page")
    if page_issues:
        for issue in page_issues:
            _issue_card(name, focus_key, issue)
    else:
        st.markdown(
            '<div class="aeg-clear-card">'
            '<div class="aeg-clear-title">No marked issues</div>'
            '<div class="aeg-clear-copy">Nothing on this page needs a location marker.</div>'
            '</div>',
            unsafe_allow_html=True,
        )

    if report_level:
        st.markdown("#### Report-level checks")
        for issue in report_level:
            _issue_card(name, focus_key, issue, show_number=False)


def _issue_card(name, focus_key, issue, show_number=True):
    severity = issue.get('severity', 'warning')
    label = {
        'critical': 'Fail',
        'warning': 'Warning',
        'info': 'Note',
        'pass': 'Pass',
    }.get(severity, 'Review')
    category = issue.get('category') or 'Review'
    note = issue.get('note') or ''
    is_focused = focus_key and st.session_state.get(focus_key) == issue.get('num')
    is_accepted = _is_accepted(name, category, note)
    classes = ['aeg-issue-card', f'aeg-{severity}']
    if is_focused:
        classes.append('aeg-focused')
    number = (
        f'<span class="aeg-issue-number">{issue.get("num")}</span>'
        if show_number and issue.get('num') is not None else ''
    )
    refs = issue.get('refs') or []
    meta = html.escape(category)
    if refs:
        meta += " · " + html.escape(", ".join(refs))
    ack = ' · ✓ acknowledged' if is_accepted else ''
    st.markdown(
        f'<div class="{" ".join(classes)}">'
        f'<div class="aeg-issue-head">{number}'
        f'<span class="aeg-issue-label">{html.escape(label)}</span></div>'
        f'<div class="aeg-issue-meta">{meta}{ack}</div>'
        f'<div class="aeg-issue-copy">{html.escape(note)}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if severity not in ('critical', 'warning'):
        return  # triage actions apply to things that could block release

    stem = finding_stem(note)
    slug = f"{name}_{category}_{stem}_{issue.get('num')}"
    can_locate = focus_key and issue.get('cells')
    cols = st.columns([1, 1, 1] if can_locate else [1, 1])
    i = 0
    if can_locate:
        if cols[i].button("📍 Locate", key=f"loc_{slug}", width="stretch",
                          help="Jump the report view to this finding's cell."):
            st.session_state[focus_key] = issue.get('num')
            st.rerun()
        i += 1
    if cols[i].button("✓ Un-ack" if is_accepted else "✓ Acknowledge",
                      key=f"ack_{slug}", width="stretch"):
        _toggle_accept(name, category, note)
        st.rerun()
    i += 1
    with cols[i].popover("✕ Dismiss", width="stretch"):
        st.caption("Dismissing removes this from the count and the verdict. "
                   "It stays restorable above.")
        reason = st.text_input("Reason", key=f"reason_{slug}",
                               placeholder="Why doesn't this apply?")
        if st.button("Confirm dismiss", key=f"dismiss_{slug}", width="stretch"):
            _dismiss(name, category, note, reason)
            st.rerun()
