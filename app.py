import streamlit as st
import tempfile
import os
from pathlib import Path

st.set_page_config(
    page_title="Work Power Tools",
    page_icon="⚙️",
    layout="centered",
)

# ── Sidebar tool selector ──────────────────────────────────────────────────
st.sidebar.title("⚙️ Work Power Tools")
TOOL = st.sidebar.radio(
    "Select a tool",
    ["🔬 SEM Report Converter", "🛠️ IIR Review"],
)
st.sidebar.markdown("---")
st.sidebar.caption("Ansaldo Energia · Materials Engineering")


# ════════════════════════════════════════════════════════════════════════════
# TOOL 1 — SEM Metallurgical Report Converter
# ════════════════════════════════════════════════════════════════════════════
def render_sem_tool():
    from sem_convert import parse, extract_figures, build

    st.title("SEM Metallurgical Report Converter")
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


# ════════════════════════════════════════════════════════════════════════════
# TOOL 2 — IIR Quality Review
# ════════════════════════════════════════════════════════════════════════════
def render_iir_tool():
    import iir_review

    SEV_ORDER = [iir_review.FAIL, iir_review.WARN, iir_review.INFO, iir_review.PASS]

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
            "received; serial-number scope totals reconcile; table vs protocol agree\n"
            "- **Integrity** — unique & contiguous positions, serial numbers present and "
            "unique, every part has a repair scope or scrap mark\n"
            "- **Consistency** — executive-summary received count and scrap positions "
            "match the protocol; finding counts ≤ received\n"
            "- **Completeness** — incoming photos embedded for each caption; page "
            "numbering consistent"
        )

    st.divider()

    iir_files = st.file_uploader(
        "IIR workbook(s) *(required)*",
        type=["xlsx"],
        accept_multiple_files=True,
        key="iir_uploader",
    )

    if 'iir_results' not in st.session_state:
        st.session_state.iir_results = []

    if st.button("▶ Run Review", type="primary", disabled=not iir_files):
        results, errors = [], []
        with st.spinner(f"Reviewing {len(iir_files)} report(s)..."):
            for f in iir_files:
                with tempfile.TemporaryDirectory() as tmp:
                    path = os.path.join(tmp, f.name)
                    with open(path, "wb") as fh:
                        fh.write(f.getvalue())
                    out_name = f"IIR_Review_{Path(f.name).stem}.xlsx"
                    out_path = os.path.join(tmp, out_name)
                    try:
                        data = iir_review.parse_iir(path)
                        findings = iir_review.run_checks(data)
                        counts, verdict = iir_review.build_checklist(data, findings, out_path)
                        with open(out_path, "rb") as fh:
                            xbytes = fh.read()
                        results.append({
                            'name': out_name, 'src': f.name, 'bytes': xbytes,
                            'ident': data['ident'], 'rp': data['received_parts'],
                            'npos': len(data['sn_rows']),
                            'findings': findings, 'counts': counts, 'verdict': verdict,
                        })
                    except Exception as e:
                        errors.append(f"{f.name}: {e}")
        for err in errors:
            st.error(f"Review failed — {err}")
        st.session_state.iir_results = results

    results = st.session_state.iir_results
    if not results and not iir_files:
        st.info("Upload one or more IIR Excel files above to run the review.")
        return

    for i, r in enumerate(results):
        st.divider()
        c = r['counts']
        verdict_kind = (iir_review.FAIL if c[iir_review.FAIL]
                        else iir_review.WARN if c[iir_review.WARN] else iir_review.PASS)
        badge = {iir_review.FAIL: "🔴", iir_review.WARN: "🟠", iir_review.PASS: "🟢"}[verdict_kind]
        st.subheader(f"{badge} {r['src']}")

        ident = r['ident']
        st.caption(
            f"**{ident.get('doc_no','?')}** · {ident.get('customer','?')} · "
            f"{ident.get('component','?')}  —  prepared by "
            f"{ident.get('preparer') or ident.get('author','?')}, "
            f"reviewed by {ident.get('reviewer','?')}"
        )

        m = st.columns(4)
        m[0].metric("🔴 Fail", c[iir_review.FAIL])
        m[1].metric("🟠 Warn", c[iir_review.WARN])
        m[2].metric("🔵 Info", c[iir_review.INFO])
        m[3].metric("🟢 Pass", c[iir_review.PASS])

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

        st.download_button(
            label="⬇ Download findings checklist (.xlsx)",
            data=r['bytes'],
            file_name=r['name'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"iir_dl_{i}",
        )


# ── Router ──────────────────────────────────────────────────────────────────
if TOOL.endswith("IIR Review"):
    render_iir_tool()
else:
    render_sem_tool()
