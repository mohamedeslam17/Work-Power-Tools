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
    "Upload the vendor SEM PDF and fill in the fields below, "
    "then click **Convert** to download the formatted Ansaldo Energia Word report."
)

st.divider()

vendor_file = st.file_uploader("Vendor PDF *(required)*", type=["pdf"])

st.subheader("Report Fields")
col1, col2 = st.columns(2)
ht_input = col1.text_input(
    "Heat Treatment Condition",
    placeholder="e.g. Aged / Solution Treated / Over-aged"
)
ia_input = col2.text_input(
    "Incoming Assessment",
    placeholder="e.g. Heavy Repair / Light Repair / Serviceable"
)
conclusion_input = st.text_area(
    "Conclusion",
    placeholder="Enter the conclusion text for the report...",
    height=160
)

st.divider()

if vendor_file:
    if st.button("Convert to Word Report", type="primary", use_container_width=True):
        with st.spinner("Parsing PDF and extracting figures..."):
            with tempfile.TemporaryDirectory() as tmp:
                vendor_path = os.path.join(tmp, vendor_file.name)
                with open(vendor_path, "wb") as f:
                    f.write(vendor_file.read())

                out_name = f"Ansaldo_{Path(vendor_file.name).stem}.docx"
                out_path = os.path.join(tmp, out_name)

                try:
                    info = parse(vendor_path)

                    # Manual fields override PDF-extracted values
                    if ht_input.strip():
                        info['ht'] = ht_input.strip()
                    if ia_input.strip():
                        info['ia'] = ia_input.strip()
                    if conclusion_input.strip():
                        info['conclusion'] = conclusion_input.strip()

                    figs = extract_figures(vendor_path)
                    build(info, figs, out_path)

                    with open(out_path, "rb") as f:
                        docx_bytes = f.read()

                    st.success(f"Done! Extracted **{len(figs)} figures**.")

                    col1, col2 = st.columns(2)
                    col1.metric("Job Number", f"JC. {info['job']}")
                    col2.metric("Stage", info['stage'])
                    col1.metric("γ′ Size — Location 1", f"{info['l1']} µm")
                    col2.metric("γ′ Size — Location 2", f"{info['l2']} µm")
                    col1.metric("Heat Treatment", info['ht'])
                    col2.metric("Material", info['material'])

                    st.divider()
                    st.download_button(
                        label="⬇ Download Word Report",
                        data=docx_bytes,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        type="primary"
                    )
                except Exception as e:
                    st.error(f"Conversion failed: {e}")
else:
    st.info("Upload a vendor PDF above to get started.")
