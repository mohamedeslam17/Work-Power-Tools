import streamlit as st
import tempfile
import os
from pathlib import Path
from sem_convert import parse, extract_figures, build
from lab_review import review_report, summarize

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

    ocr = st.checkbox(
        "🔍 Read micrograph legends (OCR)",
        value=True,
        help="Reads the magnification / scale-bar burned into each micrograph and "
             "cross-checks it against the written captions. Requires the Tesseract OCR engine.",
    )

    if not files:
        st.info("Upload one or more `.xlsx` lab reports above to review.")
        return

    for f in files:
        st.divider()
        st.subheader(f.name)

        try:
            with st.spinner("Reviewing…"):
                rtype, parsed, findings = review_report(f.name, f.getvalue(), ocr=ocr)
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


# ════════════════════════════════════════════════════════════════════════
st.title("AEG Materials Engineering Tools")

tab_conv, tab_review = st.tabs(["🔬 SEM Report Converter", "🧪 Lab Report Review"])
with tab_conv:
    render_converter()
with tab_review:
    render_reviewer()
