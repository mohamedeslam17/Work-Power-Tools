"""SEM Converter — 'Turn this vendor SEM PDF into our Word report format.'"""
import os
import tempfile
from pathlib import Path

import streamlit as st

from sem_convert import parse, extract_figures, build

_DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def render():
    vendor_files = st.file_uploader(
        "Upload vendor SEM PDF(s) to convert",
        type=["pdf"], accept_multiple_files=True, key="sem_files")
    if not vendor_files:
        return

    st.subheader("Report fields")
    col1, col2 = st.columns(2)
    ht_input = col1.selectbox(
        "Heat Treatment Condition",
        ["Full reheat treated condition, including aging.", "Solution treated",
         "Aged", "Over-aged", "As-received"], index=0)
    ia_input = col2.selectbox(
        "Incoming Assessment", ["Medium Repair", "Heavy Repair", "Light Repair"], index=0)
    col3, col4 = st.columns(2)
    mat_input = col3.text_input(
        "Material / Alloy", value="IN738",
        help="Alloy designation extracted from the PDF — edit if needed (e.g. IN738LC, IN-738).")
    conclusion_input = st.text_area(
        "Conclusion", placeholder="Enter the conclusion text for the report...", height=160)

    if 'sem_results' not in st.session_state:
        st.session_state.sem_results = []

    if st.button("▶ Generate reports", type="primary"):
        results, errors = [], []
        with st.spinner(f"Processing {len(vendor_files)} PDF(s)..."):
            for vendor_file in vendor_files:
                with tempfile.TemporaryDirectory() as tmp:
                    # Fixed temp name — never join an upload's raw name to a path
                    # (a crafted "../.." name would be a path-traversal write).
                    vendor_path = os.path.join(tmp, "vendor.pdf")
                    with open(vendor_path, "wb") as fh:
                        fh.write(vendor_file.getvalue())

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

                        with open(out_path, "rb") as fh:
                            docx_bytes = fh.read()

                        results.append({'name': out_name, 'bytes': docx_bytes,
                                        'info': info, 'fig_count': len(figs)})
                    except Exception as e:
                        errors.append(f"{vendor_file.name}: {e}")

        for err in errors:
            st.error(f"Conversion failed — {err}")
        st.session_state.sem_results = results
        if results:
            st.toast(f"Generated {len(results)} report(s).")

    results = st.session_state.sem_results
    if results:
        st.divider()
        for i, r in enumerate(results):
            with st.container(border=True):
                c1, c2 = st.columns([4, 1])
                c1.markdown(f"**{r['name']}**  ·  {r['fig_count']} figures extracted")
                c1.caption(
                    f"Job JC.{r['info']['job']} · Stage {r['info']['stage']} · "
                    f"γ′ {r['info']['l1']}/{r['info']['l2']} µm · "
                    f"{r['info']['ht']} · {r['info']['material']}")
                c2.download_button(
                    "⬇ Download", data=r['bytes'], file_name=r['name'],
                    mime=_DOCX_MIME, width="stretch", key=f"dl_{i}")
