import streamlit as st
import tempfile
import os
from pathlib import Path
from sem_convert import parse, extract_figures, build

st.set_page_config(
    page_title="SEM Report Converter — Ansaldo Energia",
    page_icon="🔬",
    layout="centered"
)

st.title("SEM Metallurgical Report Converter")
st.markdown(
    "Upload one or more vendor SEM PDFs and fill in the fields below, "
    "then click **Generate** to build the Ansaldo Energia Word report(s)."
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
