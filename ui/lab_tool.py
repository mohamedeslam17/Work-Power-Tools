"""Lab Report Review — 'Is this report releasable, and if not, where exactly?'"""
import base64
import html
import io
import time
from pathlib import Path

import streamlit as st

import report_render
import share
from lab_review import (add_duplicate_composition_findings, add_version_findings,
                        finding_stem, partition_by_scope, review_report)
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
def _cached_annotated_report(name, data, ocr, extra_findings=(), fit_width=True,
                             dpi=150):
    """Return a page-oriented report-centric review package.

    The exact LibreOffice rendering is preferred. A spreadsheet-style
    annotated fallback keeps the report visible if exact rendering is not
    available in the deployment environment.

    Deliberately built from EVERY finding, not just the ones triage state
    currently wants shown — this stays cached on the arguments alone so
    dismissing or restoring a finding never re-runs LibreOffice/Pillow. The
    triage state (scope partition, accept/dismiss) is applied afterwards, as
    a cheap filter over the already-rendered issue list; see `_render_detail`.

    fit_width defaults True here even though report_render's own default is
    False: the AEG template prints narrower than its sheet, so a literally
    faithful render cuts the sample table off mid-row (Result and Remarks
    land on later, near-empty pages). A reviewer checking data needs the whole
    row; someone who wants the true printed pagination can turn it off.

    dpi is the render resolution, raised when the reviewer zooms in: a page
    rasterised for fit-width viewing has no detail left to magnify, which is
    what made zooming useless. Each resolution is one more cached render.
    """
    rtype, parsed, findings = _cached_review(name, data, ocr)
    findings = list(findings) + list(extra_findings or ())
    exact_status = 'LibreOffice not installed'
    if report_render.libreoffice_available():
        try:
            # with_pdf=False: the sendable document is composed per triage
            # state and per comment by _cached_shared_report, so composing a
            # second, comment-less one here would be thrown away.
            view, exact_status = report_render.render_report_faithful_view(
                data, parsed, findings=findings, filename=name,
                dpi=dpi, fit_width=fit_width, with_pdf=False)
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
def _cached_shared_report(name, data, ocr, extra_findings, fit_width, dpi,
                          issue_states, extra_states, comments):
    """The sendable report: pages with comment callouts, plus the PDF.

    Cached on the triage state and the comment list rather than rebuilt on every
    interaction, and deliberately layered on top of `_cached_annotated_report`
    so this only ever costs a Pillow pass — writing a comment must not re-run
    LibreOffice.

    `issue_states` / `extra_states` carry how each finding should be presented
    (live, template-scoped, or dismissed-with-reason) instead of a filtered
    list, because every numbered marker already drawn on the page needs a
    matching callout: a badge with no comment beside it reads as a bug to
    whoever receives the document.
    """
    view, _mode = _cached_annotated_report(
        name, data, ocr, extra_findings, fit_width, dpi)
    if not view:
        return None

    states = {num: (state, reason) for num, state, reason in issue_states}
    issues = []
    for issue in view.get('issues') or ():
        state, reason = states.get(issue['num'], ('active', ''))
        record = dict(issue)
        if state != 'active':
            record['state'], record['reason'] = state, reason
        issues.append(record)

    extra_lookup = {(cat, note): (state, reason)
                    for cat, note, state, reason in extra_states}
    extras = []
    for extra in view.get('extras') or ():
        state, reason = extra_lookup.get(
            (extra['category'], extra['note']), ('active', ''))
        record = dict(extra)
        if state != 'active':
            record['state'], record['reason'] = state, reason
        extras.append(record)

    return report_render.compose_shared_report(
        view, issues=issues, extras=extras,
        comments=[{'text': text, 'author': author, 'page': page, 'issue': issue}
                  for text, author, page, issue in comments],
        filename=name, dpi=dpi)


@st.cache_data(show_spinner=False)
def _viewport_image(png):
    """(bytes, subtype) for inlining one page into the zoom viewport.

    The viewport is an <img> with a data URI — that is what lets the page be
    scaled and panned — so its bytes cross the websocket on every rerun. A
    large page render is re-encoded as JPEG for the screen only; the download
    and the sent PDF always keep the lossless render.
    """
    if Image is None or len(png) <= 700_000:
        return png, 'png'
    try:
        buffer = io.BytesIO()
        Image.open(io.BytesIO(png)).convert('RGB').save(
            buffer, format='JPEG', quality=84, optimize=True, progressive=True)
        return buffer.getvalue(), 'jpeg'
    except Exception:
        return png, 'png'


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


def _dismissal_reason(name, category, message):
    return _dismissed_store().get(_tri_key(name, category, message), '')


# ── Reviewer comments — the reviewer's own words, not the rule engine's.
# Session-local like triage state, keyed by filename, and carried onto the
# report page as callouts so they travel with the document that gets sent. ──
def _comments_store():
    return st.session_state.setdefault('lab_comments', {})   # {file: [comment]}


def _comments(name):
    return _comments_store().setdefault(name, [])


def _add_comment(name, text, author, page=None, issue=None):
    text = (text or '').strip()
    if not text:
        return False
    _comments(name).append({
        'id': f'{int(time.time() * 1000)}-{len(_comments(name))}',
        'text': text,
        'author': (author or '').strip(),
        'page': page,
        'issue': issue,
    })
    return True


def _remove_comment(name, comment_id):
    _comments_store()[name] = [
        comment for comment in _comments(name) if comment['id'] != comment_id]


def _comment_key(name):
    """The comment list as a hashable value, for the render cache."""
    return tuple((comment['text'], comment['author'], comment['page'],
                  comment['issue']) for comment in _comments(name))


def _presentation_states(r, view):
    """How every marked finding should be presented on the sendable report.

    Matched on the finding stem, the same key triage state uses, because an
    anchored highlight's note can be the leading clause of the longer message
    the findings list carries — comparing the full strings would classify the
    same finding differently in the two places.
    """
    name = r['name']
    template = {(cat, finding_stem(msg)) for _sev, cat, msg in r['template']}
    dismissed = {(cat, finding_stem(msg)): _dismissal_reason(name, cat, msg)
                 for _sev, cat, msg in r['dismissed']}

    def state_for(category, note):
        key = (category, finding_stem(note))
        if key in dismissed:
            return 'dismissed', dismissed[key]
        if key in template:
            return 'template', ''
        return 'active', ''

    issue_states, extra_states = [], []
    for issue in view.get('issues') or ():
        state, reason = state_for(issue['category'], issue['note'])
        issue_states.append((issue['num'], state, reason))
    for extra in view.get('extras') or ():
        state, reason = state_for(extra['category'], extra['note'])
        extra_states.append((extra['category'], extra['note'], state, reason))
    return tuple(issue_states), tuple(extra_states)


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
    # File picking and settings live in the sidebar, not above the report:
    # once a report is open the main column is a workspace, and an uploader
    # parked at the top of it just pushes the thing you came to look at
    # below the fold on every rerun.
    ocr_ok = _ocr_available()
    with st.sidebar:
        st.divider()
        files = st.file_uploader(
            "Lab report(s) (.xlsx)",
            type=["xlsx"], accept_multiple_files=True, key="lab_files")
        ocr = st.toggle(
            "Read micrograph legends (slower)", value=ocr_ok,
            disabled=not ocr_ok,
            help="Cross-checks every micrograph's burned-in job number, magnification, "
                 "scale and etch contrast via OCR."
                 + ("" if ocr_ok else
                    " Unavailable — Tesseract isn't installed in this environment."))
    if not files:
        st.info("Upload one or more AEG lab reports (.xlsx) from the sidebar to review them.")
        return

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
    # Cross-report: an Actual composition that is identical under two different
    # job numbers. Only visible when more than one report is uploaded, which is
    # why it sits here rather than inside review_report().
    add_duplicate_composition_findings(reviewed)
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
    """The triage workspace.

    Deliberately NOT a stack of expanders: the reviewer's job is "is this
    releasable, and if not where", so the verdict and the report-plus-findings
    panes own the top of the screen and are reachable without scrolling.
    Everything that is reference material rather than a decision — the
    extracted field dump, the template-level notes, the micrograph legends,
    the exports — sits in tabs underneath, one click away and out of the path.
    """
    name = r['name']

    _identity_bar(r)
    if r['rtype'] == 'unknown':
        st.warning("This workbook didn't match a metallurgical or coating layout, so only "
                   "a limited review ran. Check it's an AEG lab report `.xlsx`.")

    tier, label, reason = _verdict(r['active'])
    _verdict_banner(tier, label, reason)

    strip, restore = st.columns([3, 1], vertical_alignment="center")
    with strip:
        components.severity_readout(r['counts'])
    with restore:
        if r['dismissed']:
            with st.popover(f"🗑 {len(r['dismissed'])} dismissed", width="stretch"):
                _dismissed_list(name, r['dismissed'])

    # ── the workspace itself ──────────────────────────────────────────────
    package, mode, pages, issues = _annotated_report(r, ocr)
    if pages:
        _comments_and_sending(r, pages, issues, package, mode)

    # ── reference material, out of the triage path ────────────────────────
    labels = ["Extracted data", f"About this template ({len(r['template'])})", "Export"]
    tab_extracted, tab_template, tab_export = st.tabs(labels)
    with tab_extracted:
        _extracted_view(r['rtype'], r['parsed'])
    with tab_template:
        _template_body(name, r['template'])
    with tab_export:
        if r['rtype'] in ('metallurgical', 'coating'):
            if st.button("📁 Add micrographs to photo library", key=f"add_{name}"):
                added = add_to_library(name, r['f'].getvalue(), r['parsed'], r['rtype'])
                if added:
                    photo_tool.invalidate()
                st.toast(f"Added {added} micrograph(s) to the library." if added
                         else "No new micrographs (already in library).")
        st.download_button(
            "⬇ Findings (.csv)", data=_findings_csv(r['findings']),
            file_name=f"{Path(name).stem}_findings.csv", mime="text/csv",
            key=f"labcsv_{name}")


def _identity_bar(r):
    """Filename, type and the key facts on one compact line — the report's
    identity is context for the verdict, not a section of its own."""
    st.markdown(
        f'<div class="aeg-idbar">'
        f'<span class="aeg-idbar-name">{html.escape(r["name"])}</span>'
        f'<span class="aeg-idbar-tag">{html.escape(_TYPE_LABEL[r["rtype"]])}</span>'
        f'</div>',
        unsafe_allow_html=True)
    if r['facts']:
        st.caption(r['facts'])


def _dismissed_list(name, dismissed):
    for sev, cat, msg in dismissed:
        reason = _dismissed_store().get(_tri_key(name, cat, msg), '')
        cols = st.columns([5, 1], vertical_alignment="center")
        cols[0].markdown(f"**{cat}** — {msg}  \n*Dismissed: {html.escape(reason)}*")
        if cols[1].button("Restore", key=f"restore_{name}_{cat}_{finding_stem(msg)}",
                          width="stretch"):
            _restore(name, cat, msg)
            st.rerun()


def _template_body(name, template_findings):
    """D11's other half: template-scoped findings, said once, in their own tab."""
    if not template_findings:
        st.caption("Nothing template-level to report for this layout.")
        return
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


_ZOOM = {'Fit': 0, '100%': 100, '150%': 150, '200%': 200, '300%': 300}


def _zoom_dpi(zoom):
    """Render resolution for a zoom level.

    Magnifying a fit-width raster only magnifies its pixels, so beyond 150% the
    page is re-rendered at a higher resolution instead. Two resolutions, not
    one per level: each is a separate LibreOffice render.
    """
    return 240 if _ZOOM.get(zoom, 0) >= 200 else 150


def _zoom_pane(png, zoom):
    """Show a report page in a scrollable, zoomable viewport.

    `st.image` always fits its column, which is why the annotated report could
    not be zoomed at all: at fit width a spreadsheet page lands in ~900 px and
    the values in it are a few pixels tall. An <img> inside an overflow:auto box
    can be drawn at any scale and panned by scrolling — which is what someone
    checking a number against the sheet actually needs.
    """
    data, subtype = _viewport_image(png)
    percent = _ZOOM.get(zoom, 0)
    sizing = ('width:100%;' if not percent
              else f'width:{percent}%;max-width:none;')
    encoded = base64.b64encode(data).decode('ascii')
    st.markdown(
        f'<div class="aeg-zoom-pane"><img alt="Annotated report page" '
        f'style="{sizing}height:auto;display:block" '
        f'src="data:image/{subtype};base64,{encoded}"></div>',
        unsafe_allow_html=True,
    )


def _annotated_report(r, ocr):
    """Make the report itself the review UI, and the thing that gets sent.

    The page carries the original sheet, the highlighted issue locations, the
    numbered markers and — in the callout view — every comment written out
    beside the cell it is about, joined to it by a leader line. That composite
    is exactly what the PDF download and the e-mail attachment contain, so what
    the reviewer approves on screen is what the recipient opens.

    Returns (package, mode, pages, live issues) — the composed pages and PDF
    plus what they contain, so the comment and sending sections below work off
    the same artefact rather than rebuilding their own.
    """
    f = r['f']
    name = f.name
    fit_width = st.session_state.get(f'lab_fit_{name}', True)
    zoom = st.session_state.get(f'lab_zoom_{name}', 'Fit')
    dpi = _zoom_dpi(zoom)
    with st.spinner(f"Building annotated report for {name}…"):
        view, mode = _cached_annotated_report(
            name, f.getvalue(), ocr, tuple(r.get('version_findings') or ()),
            fit_width, dpi)

    if not view:
        st.error(f"Could not build the annotated report view — {mode}.")
        return None, mode, [], []

    if mode != 'exact':
        st.warning(
            "The exact Excel renderer is unavailable, so this is a simplified "
            "annotated sheet view. Issue markers and explanations are still included.")

    pages = view.get('pages') or []
    if not pages:
        st.error("The workbook rendered, but no report pages were produced.")
        return None, mode, [], []

    omitted = view.get('omitted_blank_pages') or []
    if omitted:
        st.caption(f"{len(omitted)} blank trailing source page(s) omitted: "
                   f"{', '.join(map(str, omitted))}.")

    issue_states, extra_states = _presentation_states(r, view)
    live = {num for num, state, _reason in issue_states if state == 'active'}
    all_issues = [issue for issue in (view.get('issues') or [])
                  if issue['num'] in live]
    live_extras = {(cat, note) for cat, note, state, _reason in extra_states
                   if state == 'active'}
    all_extras = [extra for extra in (view.get('extras') or [])
                  if (extra['category'], extra['note']) in live_extras]

    with st.spinner("Drawing the comment callouts…"):
        package = _cached_shared_report(
            name, f.getvalue(), ocr, tuple(r.get('version_findings') or ()),
            fit_width, dpi, issue_states, extra_states, _comment_key(name))
    if not (package or {}).get('pages'):
        # Composition is Pillow-only and should not fail, but the marked pages
        # themselves already exist — never leave the reviewer with no report and
        # nothing to take away.
        package = {'pages': [item['png'] for item in pages], 'pdf': None}

    def _count_on(page):
        return len([n for n in (page.get('issue_nums') or []) if n in live])

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

    pager, viewer, zoomer, options = st.columns(
        [2.1, 1.25, 1.5, 0.75], vertical_alignment="center")
    with pager:
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
                label_visibility="collapsed",
            )
    with viewer:
        show_callouts = st.segmented_control(
            "View", ["Callouts", "Page"], default="Callouts",
            key=f"lab_view_{name}", label_visibility="collapsed",
            help="Callouts: the page as it will be sent — every comment written "
                 "beside the cell it is about. Page: the report and its markers "
                 "only, at full width.") != "Page"
    with zoomer:
        st.segmented_control(
            "Zoom", list(_ZOOM), default="Fit", key=f"lab_zoom_{name}",
            label_visibility="collapsed",
            help="Scroll inside the page to pan. From 200% the page is "
                 "re-rendered at a higher resolution, which takes a few seconds "
                 "the first time.")
    with options:
        with st.popover("⚙", width="stretch"):
            if mode == 'exact':
                st.toggle(
                    "Fit page width", key=f'lab_fit_{name}', value=fit_width,
                    help="On: rescale so every column of a row is visible together. "
                         "Off: the workbook's own print layout, which on this template "
                         "splits wide rows across pages.")
            st.caption(f"Rendered at {dpi} dpi.")
    page = next(item for item in pages if item['number'] == selected_number)
    composed = (package or {}).get('pages') or []
    shown = page['png']
    if show_callouts and len(composed) >= selected_number:
        shown = composed[selected_number - 1]

    report_col, issue_col = st.columns([3.25, 1.35], gap="large")
    with report_col:
        st.markdown(
            f'<div class="aeg-page-kicker">PAGE {page["number"]} OF {len(pages)}'
            f'{" · WITH COMMENT CALLOUTS" if show_callouts else ""}</div>',
            unsafe_allow_html=True,
        )
        if focused:
            crop = _focus_crop(r, f, page, focused, mode)
            if crop:
                refs = ", ".join(focused.get('refs') or [])
                st.caption(f"🔍 Zoomed to issue #{focused['num']}"
                           + (f" · {refs}" if refs else ""))
                st.image(crop, width="stretch")
        _zoom_pane(shown, zoom)
    with issue_col:
        page_numbers = set(page.get('issue_nums') or [])
        page_issues = [i for i in all_issues if i['num'] in page_numbers]
        unplaced = [i for i in all_issues if not i.get('pages')]
        _page_findings(name, focus_key, page_issues, unplaced + all_extras)

    return package, mode, pages, all_issues


def _focus_crop(r, f, page, focused, mode):
    """A zoomed crop of the focused finding's cell, whichever renderer is live.

    The fallback grid is measured from the workbook, while an exact page knows
    where its markers actually landed — so 'Locate' can now zoom on both,
    instead of only on the fallback the reviewer is least likely to see.
    """
    if mode == 'exact':
        return report_render.crop_issue_detail(
            page['png'], page.get('placements'), focused['num'])
    if focused.get('cells'):
        return _cropped_detail(
            f.getvalue(), r['parsed'], page['png'], focused['cells'][0])
    return None


def _comments_and_sending(r, view_pages, issues, package, mode):
    """Write comments, then send the report — the two halves of one job.

    Side by side on purpose: the reviewer's comments are what makes the document
    worth sending, and the send panel has to show, at the moment of sending,
    exactly which comments are going with it.
    """
    name = r['name']
    st.divider()
    writing, sending = st.columns([1.45, 1], gap="large")
    with writing:
        _comment_editor(name, view_pages, issues)
    with sending:
        _send_panel(r, package, mode)


def _comment_targets(view_pages, issues):
    """[(label, page, issue)] — what a new comment can be attached to.

    Attaching to a finding is the useful case: the comment inherits that
    finding's cell, so it is drawn against the value it is about rather than
    floating at the top of the page.
    """
    options = [('The whole report', None, None)]
    for page in view_pages:
        options.append((f'Page {page["number"]}', page['number'], None))
    for issue in issues:
        note = issue.get('note') or ''
        summary = note if len(note) <= 46 else note[:45].rstrip() + '…'
        options.append(
            (f'Finding #{issue["num"]} · {issue.get("category")} — {summary}',
             None, issue['num']))
    return options


def _comment_editor(name, view_pages, issues):
    st.markdown("#### Review comments")
    st.caption(
        "Written here, drawn on the report as a callout, and included in the PDF "
        "you download or send — the recipient reads your comment on the page, "
        "beside the cell it is about.")

    options = _comment_targets(view_pages, issues)
    with st.form(f"comment_form_{name}", clear_on_submit=True, border=False):
        text = st.text_area(
            "Comment", height=110, label_visibility="collapsed",
            placeholder="e.g. Confirm with the workshop which bucket was "
                        "sectioned before this goes to the customer.")
        fields = st.columns([2.4, 1.4, 1])
        target = fields[0].selectbox(
            "Attach to", range(len(options)),
            format_func=lambda index: options[index][0])
        author = fields[1].text_input(
            "Your name", value=st.session_state.get('lab_author', ''),
            placeholder="Your name")
        fields[2].markdown('<div class="aeg-form-spacer"></div>',
                           unsafe_allow_html=True)
        added = fields[2].form_submit_button("＋ Add comment", width="stretch")
    if added:
        st.session_state['lab_author'] = (author or '').strip()
        _, page, issue = options[target]
        if _add_comment(name, text, author, page, issue):
            st.rerun()
        st.warning("Write something first — an empty comment isn't added.")

    comments = _comments(name)
    if not comments:
        st.caption("No comments yet. The report can still be sent without any.")
        return
    for index, comment in enumerate(comments, start=1):
        if comment['issue']:
            where = f"on finding #{comment['issue']}"
        elif comment['page']:
            where = f"page {comment['page']}"
        else:
            where = "whole report"
        meta = f"{comment['author'] or 'Reviewer'} · {where}"
        row = st.columns([6, 1], vertical_alignment="center")
        row[0].markdown(
            f'<div class="aeg-comment-card">'
            f'<div class="aeg-comment-head">R{index}'
            f'<span class="aeg-comment-meta">{html.escape(meta)}</span></div>'
            f'<div class="aeg-issue-copy">{html.escape(comment["text"])}</div>'
            f'</div>',
            unsafe_allow_html=True)
        if row[1].button("Remove", key=f"delc_{name}_{comment['id']}",
                         width="stretch"):
            _remove_comment(name, comment['id'])
            st.rerun()


def _share_body(r, comments):
    """The message that goes with the attachment.

    Written so the mail is useful on a phone, before the PDF is opened: the
    verdict, what was found, and the reviewer's own comments in full.
    """
    _tier, label, reason = _verdict(r['active'])
    lines = [f"Report: {r['name']}"]
    if r['facts']:
        # Packed workbook cells carry their own line breaks (four serials in one
        # cell); flattened here so the header stays one line in a mail client.
        lines.append(" ".join(r['facts'].replace('**', '').split()))
    lines += [f"Review verdict: {label} — {reason}", ""]

    if comments:
        lines.append("Reviewer comments")
        for index, comment in enumerate(comments, start=1):
            if comment['issue']:
                where = f" (on finding #{comment['issue']})"
            elif comment['page']:
                where = f" (page {comment['page']})"
            else:
                where = ""
            who = f" — {comment['author']}" if comment['author'] else ""
            lines.append(f"  R{index}{where}{who}: {comment['text']}")
        lines.append("")

    if r['active']:
        lines.append("Findings from this review")
        for severity, category, message in sorted(
                r['active'], key=lambda finding: components.RANK[
                    components.normalize_lab([finding])[0]['severity']]):
            tag = {'critical': 'FAIL', 'warning': 'WARNING',
                   'info': 'NOTE', 'pass': 'PASS'}.get(severity, 'REVIEW')
            lines.append(f"  [{tag}] {category}: {message}")
        lines.append("")

    lines.append(
        "Attached: the report itself, with every finding marked on the page and "
        "the comments written beside it.")
    return "\n".join(lines)


def _send_panel(r, package, mode):
    """Send the annotated report to people, or take it away as a file."""
    name = r['name']
    stem = Path(name).stem
    pdf = (package or {}).get('pdf')
    comments = _comments(name)
    transports = share.transports()

    st.markdown("#### Send this report")
    if pdf:
        st.caption(
            f"One PDF, {len(package['pages'])} page(s): the report with "
            f"{len(comments)} comment(s) called out on it.")
        st.download_button(
            "⬇ Download annotated report (.pdf)", data=pdf,
            file_name=f"{stem}_annotated.pdf", mime="application/pdf",
            key=f"fpdf_{name}", width="stretch")
    elif (package or {}).get('pages'):
        st.caption("The PDF could not be built here, so the marked first page "
                   "is offered as an image instead.")
        st.download_button(
            "⬇ Download annotated page (.png)", data=package['pages'][0],
            file_name=f"{stem}_annotated.png", mime="image/png",
            key=f"fpng_{name}", width="stretch")

    if not any(transports.values()):
        with st.expander("Send it from here instead of attaching it yourself"):
            st.markdown(
                "Sending isn't configured for this deployment yet. Add either "
                "set of secrets and the send button appears here:\n\n"
                "- **E-mail** — `smtp_host`, `smtp_from`, and usually "
                "`smtp_user` / `smtp_password`.\n"
                "- **Google Drive** — the same `drive_client_id`, "
                "`drive_client_secret` and `drive_refresh_token` the photo "
                "library uses; recipients get read access to the PDF.\n\n"
                "See `share.py` for the full list.")
        return
    if not pdf:
        st.info("Nothing to send yet — the annotated PDF could not be built.")
        return

    recipients_raw = st.text_input(
        "Recipients", key=f"to_{name}",
        placeholder="name@company.com, second@company.com",
        help="Comma or space separated.")
    recipients, rejected = share.parse_recipients(recipients_raw)
    if rejected:
        st.warning("Not an e-mail address: " + ", ".join(rejected))

    subject = st.text_input(
        "Subject", key=f"subj_{name}", value=f"Lab report review — {stem}")

    # The body is written from the review, so it has to keep up with it: a
    # comment added after the panel first rendered must appear in the message.
    # A widget only reads `value` once, so the generated text is pushed into
    # session state instead — and left alone the moment the reviewer types over
    # it, since silently discarding what someone wrote is worse than a stale
    # summary they can rebuild.
    body_key, generated_key = f"msg_{name}", f"msgauto_{name}"
    generated = _share_body(r, comments)
    untouched = st.session_state.get(body_key) == st.session_state.get(generated_key)
    if body_key not in st.session_state or untouched:
        st.session_state[body_key] = generated
    st.session_state[generated_key] = generated
    edited = st.session_state[body_key] != generated
    with st.expander("Message" + (" · edited" if edited else ""), expanded=False):
        if edited:
            if st.button("↺ Rebuild from the review", key=f"rebuild_{name}"):
                st.session_state[body_key] = generated
                st.rerun()
            st.caption("Your edits are kept, so new comments are not added "
                       "automatically. Rebuild to pick them up.")
        message = st.text_area(
            "Message", key=body_key, height=220, label_visibility="collapsed")

    if mode != 'exact':
        st.caption("Note: this is the simplified annotated view, not the exact "
                   "Excel rendering.")

    attachment = (f"{stem}_annotated.pdf", pdf, "application/pdf")
    buttons = st.columns(2 if all(transports.values()) else 1)
    index = 0
    if transports['email']:
        label = (f"✉ Send to {len(recipients)} recipient(s)" if recipients
                 else "✉ Send by e-mail")
        if buttons[index].button(label, key=f"mail_{name}", width="stretch",
                                 disabled=not recipients, type="primary"):
            try:
                sent = share.send_email(
                    recipients, subject, message, attachments=[attachment])
                st.success(f"Sent to {', '.join(sent)}.")
            except Exception as error:
                st.error(f"Could not send — {type(error).__name__}: {error}")
        index += 1
    if transports['drive']:
        if buttons[index].button("☁ Share via Drive", key=f"drive_{name}",
                                 width="stretch", disabled=not recipients,
                                 help="Uploads the PDF and gives each recipient "
                                      "read access — no public link."):
            try:
                result = share.share_via_drive(
                    f"{stem}_annotated.pdf", pdf, recipients, message=message)
                if result['granted']:
                    st.success("Shared with " + ", ".join(result['granted']))
                if result['failed']:
                    st.error("Could not share with: " + "; ".join(
                        f"{address} ({reason})" for address, reason in result['failed']))
                if result['link']:
                    st.markdown(f"[Open in Drive]({result['link']})")
            except Exception as error:
                st.error(f"Drive upload failed — {type(error).__name__}: {error}")
    if not recipients:
        st.caption("Add at least one recipient to enable sending.")


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
    # Two rows, not three squeezed columns: this sidebar column is narrow
    # enough that even short labels ("Drop") wrapped mid-word there —
    # confirmed both by independent verification (FEEDBACK.md entry 008) and
    # by re-screenshotting after the first, still-too-narrow fix. Dismiss
    # gets the full row since its popover needs to anchor to something wide
    # enough not to feel cramped.
    can_locate = focus_key and issue.get('cells')
    top = st.columns([1, 1] if can_locate else [1])
    i = 0
    if can_locate:
        if top[i].button("📍 Locate", key=f"loc_{slug}", width="stretch",
                         help="Jump the report view to this finding's cell."):
            st.session_state[focus_key] = issue.get('num')
            st.rerun()
        i += 1
    if top[i].button("↺ Ack" if is_accepted else "✓ Ack",
                     key=f"ack_{slug}", width="stretch",
                     help="Click to un-acknowledge." if is_accepted
                     else "Acknowledge — mark reviewed; does not affect the verdict."):
        _toggle_accept(name, category, note)
        st.rerun()
    with st.popover("✕ Dismiss", width="stretch"):
        st.caption("Dismissing removes this from the count and the verdict. "
                   "It stays restorable above.")
        reason = st.text_input("Reason", key=f"reason_{slug}",
                               placeholder="Why doesn't this apply?")
        if st.button("Confirm dismiss", key=f"dismiss_{slug}", width="stretch"):
            _dismiss(name, category, note, reason)
            st.rerun()
