import streamlit as st
import tempfile
import os
from pathlib import Path
from sem_convert import parse, extract_figures, build
from lab_review import review_report, summarize
try:
    from lab_review import HT_ORDER
except ImportError:   # tolerate a stale/cached lab_review module on redeploy
    HT_ORDER = ['As-received', 'Post-solution', 'Post stress-relief', 'Post-ageing', 'Unspecified']
import report_render
from photo_lib import (add_to_library, alloy_counts, photos_for,
                       get_image_bytes, backend_name, LIBRARY_DIR)


@st.cache_data(show_spinner=False)
def _lib_image(path, drive_id):
    """Cached fetch of one library image (local file / GitHub / Drive)."""
    return get_image_bytes({"path": path, "drive_id": drive_id})


@st.cache_data(show_spinner=False)
def _review_and_render(name, data, ocr, faithful=True):
    """Review a report and build its annotated images, cached on the bytes so
    re-runs (and the check toggles) don't re-parse or re-render. The annotated
    view is the pixel-faithful LibreOffice render when available, else the
    drawn-grid fallback; `mode`/`status` say which (and why)."""
    rtype, parsed, findings = review_report(name, data, ocr=ocr)
    annotated, mode, status = None, 'none', ''
    if faithful:
        annotated, status = report_render.render_report_faithful(data, parsed, filename=name)
        if annotated:
            mode = 'faithful'
    if not annotated:
        annotated = report_render.render_report_image(data, parsed, findings, rtype, filename=name)
        mode = 'grid' if annotated else 'none'
    micrographs = report_render.annotate_micrographs(data, parsed)
    return rtype, parsed, findings, annotated, micrographs, mode, status

st.set_page_config(
    page_title="AEG Materials Tools",
    page_icon="🔬",
    layout="centered"
)


# ════════════════════════════════════════════════════════════════════════
# TAB 1 — SEM Metallurgical Report Converter (vendor PDF → Ansaldo Word)
# ════════════════════════════════════════════════════════════════════════
def render_converter():
    st.markdown(
        "Upload one or more vendor SEM PDFs and fill in the fields below, "
        "then click **Generate** to build the Word report(s)."
    )

    st.divider()

    vendor_files = st.file_uploader(
        "Vendor PDF(s) *(required)*",
        type=["pdf"],
        accept_multiple_files=True,
    )

    st.subheader("Report Fields")
    col1, col2 = st.columns(2)
    ht_input = col1.selectbox(
        "Heat Treatment Condition",
        [
            "Full reheat treated condition, including aging.",
            "Solution treated",
            "Aged",
            "Over-aged",
            "As-received",
        ],
        index=0
    )
    ia_input = col2.selectbox(
        "Incoming Assessment",
        ["Medium Repair", "Heavy Repair", "Light Repair"],
        index=0
    )
    col3, col4 = st.columns(2)
    mat_input = col3.text_input(
        "Material / Alloy",
        value="IN738",
        help="Alloy designation extracted from the PDF — edit if needed (e.g. IN738LC, IN-738).",
    )
    conclusion_input = st.text_area(
        "Conclusion",
        placeholder="Enter the conclusion text for the report...",
        height=160
    )

    st.divider()

    if 'results' not in st.session_state:
        st.session_state.results = []

    btn_col1, btn_col2 = st.columns(2)

    if btn_col1.button("▶ Generate Reports", type="primary", disabled=not vendor_files):
        results = []
        errors  = []

        with st.spinner(f"Processing {len(vendor_files)} PDF(s)..."):
            for vendor_file in vendor_files:
                with tempfile.TemporaryDirectory() as tmp:
                    vendor_path = os.path.join(tmp, vendor_file.name)
                    with open(vendor_path, "wb") as f:
                        f.write(vendor_file.getvalue())

                    out_name = f"Ansaldo_{Path(vendor_file.name).stem}.docx"
                    out_path = os.path.join(tmp, out_name)

                    try:
                        info = parse(vendor_path)
                        info['ht'] = ht_input
                        info['ia'] = ia_input
                        if mat_input.strip():
                            info['material'] = mat_input.strip()
                        if conclusion_input.strip():
                            info['conclusion'] = conclusion_input.strip()

                        figs = extract_figures(vendor_path)
                        build(info, figs, out_path)

                        with open(out_path, "rb") as f:
                            docx_bytes = f.read()

                        results.append({
                            'name':      out_name,
                            'bytes':     docx_bytes,
                            'info':      info,
                            'fig_count': len(figs),
                        })
                    except Exception as e:
                        errors.append(f"{vendor_file.name}: {e}")

        if errors:
            for err in errors:
                st.error(f"Conversion failed — {err}")

        st.session_state.results = results
        if results:
            st.success(f"Generated {len(results)} report(s).")

    # ── Download buttons appear in btn_col2 side-by-side with Generate ──
    results = st.session_state.results
    if results:
        for i, r in enumerate(results):
            btn_col2.download_button(
                label=f"⬇ {Path(r['name']).stem[-28:]}",
                data=r['bytes'],
                file_name=r['name'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key=f"dl_{i}",
            )

        st.divider()
        for r in results:
            st.success(f"**{r['name']}** — {r['fig_count']} figures extracted")
            m1, m2 = st.columns(2)
            m1.metric("Job Number", f"JC. {r['info']['job']}")
            m2.metric("Stage", r['info']['stage'])
            m1.metric("γ′ Size — Location 1", f"{r['info']['l1']} µm")
            m2.metric("γ′ Size — Location 2", f"{r['info']['l2']} µm")
            m1.metric("Heat Treatment", r['info']['ht'])
            m2.metric("Material", r['info']['material'])

    if not vendor_files:
        st.info("Upload one or more vendor PDFs above to get started.")


# ════════════════════════════════════════════════════════════════════════
# TAB 2 — Lab Report Review (rule-based QA on AEG Excel reports)
# ════════════════════════════════════════════════════════════════════════
# severity → (streamlit render fn, icon, label)
_SEV = {
    'critical': ('error',   '🔴', 'Fail'),
    'warning':  ('warning', '🟠', 'Warning'),
    'info':     ('info',    '🔵', 'Note'),
    'pass':     ('success', '🟢', 'Pass'),
}
_SEV_ORDER = ['critical', 'warning', 'info', 'pass']


def _render_parsed(rtype, parsed):
    """Show the facts the reviewer extracted, for transparency."""
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
                dev = f"{(a - n) / abs(n) * 100:+.0f}%" if (n not in (None, 0) and a is not None) else "—"
                rows.append({
                    "Element": el,
                    "Nominal": f"{n:g}" if n is not None else "—",
                    "Actual":  f"{a:g}" if a is not None else "—",
                    "Δ":       dev,
                })
            st.table(rows)

    elif rtype == 'coating':
        st.write(f"**Report No.:** {parsed.get('report_no') or '—'}")
        st.write(f"**Title:** {parsed.get('title') or '—'}")
        rows = parsed.get('rows', [])
        if rows:
            lo, hi = rows[0].get('min'), rows[0].get('max')
            if lo is not None and hi is not None:
                st.write(f"**Design limit:** {lo:g} – {hi:g} mm")
            st.table([
                {"Row": e['row'], "Measurements (mm)": ", ".join(f"{v:g}" for v in e['values'])}
                for e in rows
            ])

    legends = parsed.get('legends') or []
    if legends:
        st.markdown("**Micrograph legends — read from the images**")
        st.table([
            {"Image": l.get('image', '—'), "Magnification": l.get('mag', '—'),
             "Scale": l.get('scale', '—'), "Legend ID": l.get('id', '—')}
            for l in legends
        ])


def render_reviewer():
    st.markdown(
        "Upload one or more AEG lab reports (`.xlsx`) — **metallurgical** or "
        "**coating** — for an automated QA review. Each report is checked against "
        "its own stated spec (Nominal composition / design limits), plus hardness, "
        "completeness and sign-off."
    )

    st.divider()

    files = st.file_uploader(
        "Lab report(s) *(.xlsx)*",
        type=["xlsx"],
        accept_multiple_files=True,
        key="lab_files",
    )

    oc1, oc2 = st.columns(2)
    ocr = oc1.checkbox(
        "🔍 Analyse micrographs (legends, etch, thickness)",
        value=True,
        help="Reads each micrograph's burned-in legend (magnification / scale), gauges "
             "etched-vs-low-contrast, and reads any burned-in thickness measurements — "
             "cross-checking against the captions and comment. Requires the Tesseract OCR engine.",
    )
    faithful = oc2.checkbox(
        "🖼 Pixel-faithful report image (LibreOffice)",
        value=True,
        help="Renders the real workbook with LibreOffice — original fonts, column widths, "
             "borders and embedded micrographs — and overlays the flagged cells. Falls back "
             "to a fast drawn grid when LibreOffice isn't available (the first render of a "
             "report can take a few seconds).",
    )

    if not files:
        st.info("Upload one or more `.xlsx` lab reports above to review.")
        return

    for f in files:
        st.divider()
        st.subheader(f.name)

        try:
            with st.spinner("Reviewing…"):
                rtype, parsed, findings, annotated, micrographs, mode, status = \
                    _review_and_render(f.name, f.getvalue(), ocr, faithful)
        except Exception as e:
            st.error(f"Could not read report — {e}")
            continue

        counts = summarize(findings)
        st.caption(f"Detected report type: **{rtype}**")

        m = st.columns(4)
        m[0].metric("🔴 Fail", counts['critical'])
        m[1].metric("🟠 Warning", counts['warning'])
        m[2].metric("🔵 Note", counts['info'])
        m[3].metric("🟢 Pass", counts['pass'])

        if counts['critical']:
            st.error("**Needs attention** — critical findings present.")
        elif counts['warning']:
            st.warning("**Review recommended** — warnings present.")
        else:
            st.success("**Looks good** — no warnings or failures.")

        for sev in _SEV_ORDER:
            items = [(cat, msg) for s, cat, msg in findings if s == sev]
            if not items:
                continue
            fn, icon, _ = _SEV[sev]
            render = getattr(st, fn)
            for cat, msg in items:
                render(f"{icon} **{cat}** — {msg}")

        with st.expander("Extracted data"):
            _render_parsed(rtype, parsed)

        if annotated or micrographs:
            flagged = bool(counts['critical'] or counts['warning'])
            with st.expander("🖼 Annotated report view — issue areas highlighted",
                             expanded=flagged):
                if annotated:
                    if mode == 'faithful':
                        note = "Pixel-faithful LibreOffice render — flagged cells highlighted, numbered to the legend."
                    else:
                        note = "Drawn-grid view — flagged cells boxed and numbered to the legend."
                        if faithful and status:
                            note += f"  (Pixel-faithful render unavailable: {status}.)"
                    st.image(annotated, use_container_width=True, caption=note)
                    st.download_button(
                        "⬇ Download annotated report (.png)", data=annotated,
                        file_name=f"{Path(f.name).stem}_annotated.png",
                        mime="image/png", key=f"annpng_{f.name}")
                elif rtype in ('metallurgical', 'coating'):
                    st.caption("Annotated view unavailable (image libraries not installed).")
                if micrographs:
                    st.markdown("**Annotated micrographs** — legend / scale-bar regions "
                                "boxed, contrast and any burned-in thickness flagged.")
                    mcols = st.columns(2)
                    for i, (mname, mbytes, mcap) in enumerate(micrographs):
                        mcols[i % 2].image(mbytes, caption=mcap, use_container_width=True)

        if rtype in ('metallurgical', 'coating'):
            if st.button("📁 Add this report's micrographs to the library",
                         key=f"add_{f.name}"):
                added = add_to_library(f.name, f.getvalue(), parsed, rtype)
                st.success(f"Added {added} micrograph(s) to the library."
                           if added else "No new micrographs added (already in library).")


# ════════════════════════════════════════════════════════════════════════
# TAB 3 — Photo Library (per-alloy micrograph gallery)
# ════════════════════════════════════════════════════════════════════════
def render_gallery():
    st.markdown(
        "Browse stored micrographs **by alloy**, with the data of the report they "
        "came from. Add micrographs from the **Lab Report Review** tab."
    )
    st.caption(f"Storage backend: **{backend_name()}**")
    counts = alloy_counts()
    total = sum(counts.values())
    if not total:
        st.info("The library is empty. Review a report and click "
                "“Add this report's micrographs to the library”.")
        return

    st.caption(f"{total} micrograph(s) across {len(counts)} alloy(s)")
    pick = st.selectbox("Alloy", [f"{a} ({n})" for a, n in sorted(counts.items())])
    alloy = pick.rsplit(" (", 1)[0]
    recs = photos_for(alloy)

    # Segregate within the alloy by either heat-treatment condition or etchant.
    dim = st.selectbox("Segregate by", ["Heat treatment", "Etchant"])
    key = "ht" if dim == "Heat treatment" else "etchant"
    other = "etchant" if key == "ht" else "ht"
    icon = "🔥" if key == "ht" else "🧪"

    groups = {}
    for r in recs:
        groups.setdefault(r.get(key) or "Unspecified", []).append(r)
    if key == "ht":          # process sequence for HT, alphabetical for etchant
        order = [g for g in HT_ORDER if g in groups] + sorted(g for g in groups if g not in HT_ORDER)
    else:
        order = sorted(groups)

    flt = st.selectbox(dim, ["All"] + [f"{g} ({len(groups[g])})" for g in order])
    shown = order if flt == "All" else [flt.rsplit(" (", 1)[0]]

    for g in shown:
        st.divider()
        st.subheader(f"{icon} {g}  ·  {len(groups[g])}")
        cols = st.columns(3)
        for i, r in enumerate(groups[g]):
            c = cols[i % 3]
            img = _lib_image(r.get("path"), r.get("drive_id"))
            contrast = "etched" if r.get("etched") else "low-contrast"
            cap = f"Job {r.get('job') or '—'} · {r.get('mag') or '?'} · {contrast}"
            if img:
                c.image(img, caption=cap, width="stretch")
            else:
                c.warning(f"missing: {r.get('path') or r.get('drive_id')}")
            bits = [f"set: {(r.get('source') or '')[:30]}"]
            if r.get(other):
                bits.append(str(r.get(other)))
            meas = r.get("measurements") or []
            if meas:
                bits.append(", ".join(f"{m}µm" for m in meas))
            c.caption(" · ".join(bits))


# ════════════════════════════════════════════════════════════════════════
# TAB 4 — IIR Quality Review (Incoming Inspection Report consistency/completeness QA)
def render_iir_tool():
    import io
    import iir_review
    from itertools import groupby

    XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    BADGE = {iir_review.FAIL: "🔴", iir_review.WARN: "🟠",
             iir_review.INFO: "🔵", iir_review.PASS: "🟢"}
    CHOICE_LABEL = {iir_review.FAIL: "🔴 Fail", iir_review.WARN: "🟠 Warn",
                    iir_review.INFO: "🔵 Info", iir_review.OFF: "⚪ Off"}

    def check_settings():
        """Severity dropdown per check; returns {check_title: severity|OFF}."""
        with st.expander("⚙️ Check settings — set the severity of each check"):
            st.caption("Change any check's level, or turn it Off. The overview and "
                       "verdicts below update immediately — no need to re-upload.")
            if st.button("↺ Reset to defaults", key="iir_sev_reset"):
                for _, title, _ in iir_review.CHECK_CATALOG:
                    st.session_state.pop(f"iir_sev::{title}", None)
            chosen = {}
            for cat, items in groupby(iir_review.CHECK_CATALOG, key=lambda t: t[0]):
                st.markdown(f"**{cat}**")
                items = list(items)
                cols = st.columns(2)
                for j, (_, title, default) in enumerate(items):
                    chosen[title] = cols[j % 2].selectbox(
                        title, iir_review.SEVERITY_CHOICES,
                        index=iir_review.SEVERITY_CHOICES.index(default),
                        format_func=lambda s: CHOICE_LABEL[s],
                        key=f"iir_sev::{title}",
                    )
        return chosen

    st.title("IIR Quality Review")
    st.markdown(
        "Upload one or more **Incoming Inspection Reports** (Detailed Assessment "
        "Customer Reports) as `.xlsx`. Each workbook is checked for internal "
        "consistency and completeness, and a downloadable findings checklist is produced."
    )
    with st.expander("What does it check?"):
        st.markdown(
            "- **Identity / metadata** — doc number, customer, component, preparer, "
            "reviewer, approver and PO# present and well-formed\n"
            "- **Quantities** — Received = Scrap + Reconditionable; positions listed = "
            "received; serial-number scope totals reconcile; the stated sum row matches "
            "the scopes actually marked; Received-Parts table vs Serial-Number protocol agree\n"
            "- **Integrity** — unique & contiguous positions, serial numbers present and "
            "unique, valid repair-scope values, scrap mark ↔ scope 'S', every part scoped\n"
            "- **Consistency** — Summary-of-Damages finding counts match the protocol "
            "defect marks; executive-summary received count and scrap positions agree\n"
            "- **Completeness** — incoming photos embedded for each caption; page numbering\n\n"
            "Review **several reports at once** for a combined **Batch Summary** workbook "
            "(one row per report + a pooled findings tab)."
        )

    st.divider()

    iir_files = st.file_uploader(
        "IIR workbook(s) *(required)*",
        type=["xlsx"],
        accept_multiple_files=True,
        key="iir_uploader",
    )

    if 'iir_data' not in st.session_state:
        st.session_state.iir_data = []   # [{'src': name, 'data': parsed}, ...]

    # Parse on demand; cached so changing the check settings re-checks instantly.
    if st.button("▶ Run Review", type="primary", disabled=not iir_files):
        parsed, errors = [], []
        with st.spinner(f"Reading {len(iir_files)} report(s)..."):
            with tempfile.TemporaryDirectory() as tmp:
                for f in iir_files:
                    path = os.path.join(tmp, f.name)
                    with open(path, "wb") as fh:
                        fh.write(f.getvalue())
                    try:
                        parsed.append({'src': f.name, 'data': iir_review.parse_iir(path)})
                    except Exception as e:
                        errors.append(f"{f.name}: {e}")
        for err in errors:
            st.error(f"Review failed — {err}")
        st.session_state.iir_data = parsed

    parsed = st.session_state.iir_data
    if not parsed:
        st.info("Upload IIR Excel file(s) and click ▶ Run Review." if iir_files
                else "Upload one or more IIR Excel files above to run the review.")
        return

    overrides = check_settings()

    # Re-run the checks for every report with the current severity settings.
    results = []
    for item in parsed:
        data = item['data']
        findings = iir_review.run_checks(data, overrides)
        results.append({
            'src': item['src'], 'data': data, 'findings': findings,
            'counts': iir_review.count_severities(findings),
            'ident': data['ident'], 'rp': data['received_parts'],
            'npos': len(data['sn_rows']),
        })

    # ── Batch overview (one row per report) ──────────────────────────────────
    st.divider()
    tot = {k: sum(r['counts'][k] for r in results)
           for k in (iir_review.FAIL, iir_review.WARN, iir_review.INFO, iir_review.PASS)}
    st.subheader(f"Batch overview — {len(results)} report(s)")
    overview = []
    for r in results:
        c = r['counts']
        sev, label = iir_review.verdict_of(c)
        rp = r['rp']
        overview.append({
            "Verdict": f"{BADGE[sev]} {label.split(' — ')[0]}",
            "Report": r['src'],
            "Doc No": r['ident'].get('doc_no', ''),
            "Component": r['ident'].get('component', ''),
            "Recv": rp.get('received') if rp.get('found') else None,
            "Scrap": rp.get('scrap') if rp.get('found') else None,
            "Recond": rp.get('reconditionable') if rp.get('found') else None,
            "🔴": c[iir_review.FAIL], "🟠": c[iir_review.WARN],
            "🔵": c[iir_review.INFO], "🟢": c[iir_review.PASS],
        })
    st.dataframe(overview, use_container_width=True, hide_index=True)

    m = st.columns(4)
    m[0].metric("🔴 Fail", tot[iir_review.FAIL])
    m[1].metric("🟠 Warn", tot[iir_review.WARN])
    m[2].metric("🔵 Info", tot[iir_review.INFO])
    m[3].metric("🟢 Pass", tot[iir_review.PASS])

    if len(results) > 1:
        records = [iir_review.record_of(r['data'], r['findings']) for r in results]
        buf = io.BytesIO()
        iir_review.build_batch_summary(records, buf)
        st.download_button(
            "⬇ Download batch summary (.xlsx)", data=buf.getvalue(),
            file_name="IIR_Batch_Summary.xlsx", mime=XLSX_MIME,
            key="iir_batch_dl", type="primary",
        )

    # ── Per-report detail (expanders) ────────────────────────────────────────
    st.divider()
    st.subheader("Report details")
    for i, r in enumerate(results):
        c = r['counts']
        sev, _ = iir_review.verdict_of(c)
        ident = r['ident']
        header = (f"{BADGE[sev]} {r['src']}  —  {c[iir_review.FAIL]}F / "
                  f"{c[iir_review.WARN]}W / {c[iir_review.INFO]}I / {c[iir_review.PASS]}P")
        with st.expander(header, expanded=(sev != iir_review.PASS)):
            st.caption(
                f"**{ident.get('doc_no','?')}** · {ident.get('customer','?')} · "
                f"{ident.get('component','?')}  —  prepared by "
                f"{ident.get('preparer') or ident.get('author','?')}, "
                f"reviewed by {ident.get('reviewer','?')}"
            )
            rp = r['rp']
            if rp.get('found'):
                st.caption(
                    f"Received **{rp['received']}** · Scrap **{rp['scrap']}** · "
                    f"Reconditionable **{rp['reconditionable']}** · "
                    f"Positions in protocol **{r['npos']}**"
                )
            rows = [{
                "Severity": f"{iir_review.SEV_ICON[f['severity']]} {f['severity']}",
                "Category": f['category'],
                "Check": f['check'],
                "Sheet": f['sheet'],
                "Detail": f['detail'],
            } for f in r['findings']]
            st.dataframe(rows, use_container_width=True, hide_index=True)
            buf = io.BytesIO()
            iir_review.build_checklist(r['data'], r['findings'], buf)
            st.download_button(
                label="⬇ Download findings checklist (.xlsx)",
                data=buf.getvalue(),
                file_name=f"IIR_Review_{Path(r['src']).stem}.xlsx",
                mime=XLSX_MIME, key=f"iir_dl_{i}",
            )


# ════════════════════════════════════════════════════════════════════════
st.title("AEG Materials Engineering Tools")

tab_conv, tab_review, tab_gallery, tab_iir = st.tabs(
    ["🔬 SEM Report Converter", "🧪 Lab Report Review", "🖼️ Photo Library",
     "🛠️ IIR Review"])
with tab_conv:
    render_converter()
with tab_review:
    render_reviewer()
with tab_gallery:
    render_gallery()
with tab_iir:
    render_iir_tool()
