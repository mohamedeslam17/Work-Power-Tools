import streamlit as st
import tempfile
import os
import zipfile
import io
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

if vendor_files:
    btn_col1, btn_col2 = st.columns(2)

    if btn_col1.button("⚙ Generate Report(s)", type="primary"):
        results = []
        errors  = []

        with st.spinner(f"Processing {len(vendor_files)} PDF(s)…"):
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

        st.session_state['results'] = results

    # ── Download button (appears in btn_col2 once results are ready) ──
    results = st.session_state.get('results', [])

    if results:
        if len(results) == 1:
            r = results[0]
            btn_col2.download_button(
                label="⬇ Download Report",
                data=r['bytes'],
                file_name=r['name'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
        else:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                for r in results:
                    zf.writestr(r['name'], r['bytes'])
            zip_buf.seek(0)
            btn_col2.download_button(
                label=f"⬇ Download All ({len(results)} reports)",
                data=zip_buf.getvalue(),
                file_name="Ansaldo_Reports.zip",
                mime="application/zip",
                type="primary",
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

else:
    st.info("Upload one or more vendor PDFs above to get started.")
