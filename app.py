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
def _review(name, data, ocr):
    """Parse + rule-check one report (fast). Cached on the file bytes."""
    return review_report(name, data, ocr=ocr)


@st.cache_data(show_spinner=False)
def _grid_and_micros(name, data, ocr):
    """The instant annotated view: the drawn-grid image + annotated micrographs."""
    rtype, parsed, findings = _review(name, data, ocr)
    try:
        grid = report_render.render_report_image(data, parsed, findings, rtype, filename=name)
    except Exception:
        grid = None
    try:
        micros = report_render.annotate_micrographs(data, parsed)
    except Exception:
        micros = []
    return grid, micros


@st.cache_data(show_spinner=False)
def _faithful_image(name, data, ocr):
    """The slow pixel-faithful LibreOffice render — built only on demand, so a
    big report never blocks the review. Returns (png_or_None, status)."""
    _, parsed, findings = _review(name, data, ocr)
    try:
        return report_render.render_report_faithful(data, parsed, findings=findings, filename=name)
    except TypeError:                       # tolerate a stale report_render on redeploy
        try:
            return report_render.render_report_faithful(data, parsed, filename=name)
        except Exception as e:
            return None, f'{type(e).__name__}: {e}'
    except Exception as e:
        return None, f'{type(e).__name__}: {e}'


@st.cache_data(show_spinner=False, ttl=120)
def _gallery_counts():
    """Alloy → micrograph count from the (possibly remote) library index, cached
    so the backend isn't hit on every rerun. Raises on backend failure — the
    caller catches it so a flaky backend can't take the whole app down."""
    return alloy_counts()


@st.cache_data(show_spinner=False, ttl=120)
def _gallery_photos(alloy):
    return photos_for(alloy)


# Complete UI redesign: Inter type, a light sidebar-nav dashboard shell, a
# branded page header, and designed cards / inputs / metrics / controls.
_CSS = """
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
  html, body, .stApp, button, input, textarea, select,
  [class*="css"] { font-family:'Inter',-apple-system,system-ui,Segoe UI,sans-serif !important; }
  .stApp { background:#eef2f8; }
  .block-container { max-width:1140px; padding-top:1.3rem; padding-bottom:4rem; }
  h1,h2,h3,h4 { letter-spacing:-.018em; color:#0f172a; font-weight:700; }
  p,li,label,.stMarkdown { color:#334155; }
  #MainMenu, footer, [data-testid="stDecoration"] { display:none; }
  [data-testid="stHeader"] { background:transparent; height:0; }

  /* Sidebar */
  [data-testid="stSidebar"] { background:#ffffff; border-right:1px solid #e6eaf2; }
  [data-testid="stSidebar"] > div { padding-top:.6rem; }
  [data-testid="stSidebar"] [role="radiogroup"] { gap:.18rem; }
  [data-testid="stSidebar"] [role="radiogroup"] label {
    display:flex; align-items:center; width:100%; padding:.62rem .85rem; margin:0;
    border-radius:11px; cursor:pointer; font-weight:600; font-size:.95rem;
    color:#475569; transition:all .12s; }
  [data-testid="stSidebar"] [role="radiogroup"] label:hover { background:#f1f5f9; }
  [data-testid="stSidebar"] [role="radiogroup"] label > div:first-child { display:none; }
  [data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {
    background:#eaf1ff; color:#1d4ed8; box-shadow:inset 3px 0 0 #2563eb; }

  /* Branded page header */
  .pagehead { display:flex; align-items:center; gap:.9rem; margin:.1rem 0 1.2rem; }
  .pagehead .ic { width:48px; height:48px; flex:none; border-radius:14px;
    display:flex; align-items:center; justify-content:center; font-size:1.5rem;
    background:linear-gradient(135deg,#2563eb,#4f46e5); color:#fff;
    box-shadow:0 8px 18px rgba(37,99,235,.28); }
  .pagehead h2 { margin:0; font-size:1.55rem; }
  .pagehead .sub { color:#64748b; font-size:.92rem; margin-top:.12rem; }

  /* Cards */
  [data-testid="stVerticalBlockBorderWrapper"] {
    background:#fff; border:1px solid #e8ecf4 !important; border-radius:18px;
    box-shadow:0 1px 2px rgba(15,23,42,.04), 0 12px 32px rgba(15,23,42,.06); }

  /* Inputs */
  .stTextInput input, .stTextArea textarea,
  div[data-baseweb="select"] > div, div[data-baseweb="input"] > div {
    background:#fff !important; border:1px solid #dbe2ec !important; border-radius:11px !important; }
  .stTextInput input:focus, .stTextArea textarea:focus {
    border-color:#2563eb !important; box-shadow:0 0 0 3px rgba(37,99,235,.13) !important; }

  /* Buttons */
  .stButton button, .stDownloadButton button {
    border-radius:11px; font-weight:600; border:1px solid #dbe2ec; padding:.45rem 1rem; transition:all .12s; }
  .stButton button:hover, .stDownloadButton button:hover {
    border-color:#2563eb; color:#2563eb; box-shadow:0 3px 10px rgba(37,99,235,.13); }
  .stButton button[kind="primary"], .stDownloadButton button[kind="primary"] {
    background:#2563eb; border-color:#2563eb; color:#fff; }
  .stButton button[kind="primary"]:hover { background:#1d4ed8; color:#fff; }

  /* Metrics */
  [data-testid="stMetric"] {
    background:#f8fafc; border:1px solid #eef2f8; border-radius:14px; padding:.55rem .9rem; }
  [data-testid="stMetricValue"] { font-size:1.6rem; font-weight:800; }
  [data-testid="stMetricLabel"] p { font-size:.78rem; opacity:.85; }

  /* Sub-tabs */
  .stTabs [data-baseweb="tab-list"] { gap:.15rem; border-bottom:1px solid #e8ecf4; }
  .stTabs [data-baseweb="tab"] { padding:.45rem 1.05rem; font-weight:600; color:#64748b; }
  .stTabs [aria-selected="true"] { color:#2563eb; }

  /* Uploader / expander / divider */
  [data-testid="stFileUploaderDropzone"] { background:#f8fafe; border:1.5px dashed #c6d2e6; border-radius:14px; }
  [data-testid="stExpander"] details { border:1px solid #e8ecf4 !important; border-radius:13px; background:#fff; }
  hr { border-color:#e8ecf4; margin:.7rem 0; }
</style>
"""

_BRAND = (
    '<div style="display:flex;align-items:center;gap:.6rem;padding:.4rem .35rem 1rem;">'
    '<div style="width:40px;height:40px;border-radius:12px;'
    'background:linear-gradient(135deg,#2563eb,#4f46e5);display:flex;align-items:center;'
    'justify-content:center;font-size:1.35rem;box-shadow:0 8px 18px rgba(37,99,235,.3);">🔬</div>'
    '<div><div style="font-weight:800;font-size:1.05rem;color:#0f172a;line-height:1.1;">AEG Tools</div>'
    '<div style="font-size:.74rem;color:#94a3b8;">Materials Engineering</div></div></div>'
)


def _page_header(icon, title, sub):
    return (f'<div class="pagehead"><div class="ic">{icon}</div>'
            f'<div><h2>{title}</h2><div class="sub">{sub}</div></div></div>')


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
                    # Fixed temp name — never join an upload's raw name to a path
                    # (a crafted "../.." name would be a path-traversal write).
                    vendor_path = os.path.join(tmp, "vendor.pdf")
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

# severity → (accent colour, light tint, label) for the compact HTML findings list
_SEV_STYLE = {
    'critical': ('#d62d38', '#fdecee', 'Fail'),
    'warning':  ('#e07b16', '#fdf2e3', 'Warning'),
    'info':     ('#1a6ed6', '#e9f1fc', 'Note'),
    'pass':     ('#1f9e50', '#e8f6ee', 'Pass'),
}


def _findings_html(items, sev):
    """Compact colour-coded rows for one severity (items = [(category, msg), …])."""
    import html as _html
    color, tint, label = _SEV_STYLE[sev]
    rows = [
        f'<div style="display:flex;align-items:baseline;gap:.6rem;'
        f'padding:.42rem .75rem;margin:.3rem 0;border-left:4px solid {color};'
        f'background:{tint};border-radius:8px;">'
        f'<span style="color:{color};font-weight:700;font-size:.7rem;letter-spacing:.03em;'
        f'text-transform:uppercase;min-width:56px;">{label}</span>'
        f'<span style="line-height:1.4;color:#28323f;"><b>{_html.escape(cat)}</b> — '
        f'{_html.escape(msg)}</span></div>'
        for cat, msg in items]
    return '<div>' + ''.join(rows) + '</div>'


def _render_findings(findings):
    """Triage: fails + warnings up front; notes and passes tucked behind
    disclosures so the things that need attention aren't buried."""
    by = {s: [(c, m) for ss, c, m in findings if ss == s] for s in _SEV_ORDER}
    if by['critical']:
        st.markdown(_findings_html(by['critical'], 'critical'), unsafe_allow_html=True)
    if by['warning']:
        st.markdown(_findings_html(by['warning'], 'warning'), unsafe_allow_html=True)
    if not by['critical'] and not by['warning']:
        st.markdown(_findings_html([('Result', 'No fails or warnings.')], 'pass'),
                    unsafe_allow_html=True)
    if by['info']:
        with st.expander(f"🔵  {len(by['info'])} note{'s' if len(by['info']) != 1 else ''}"):
            st.markdown(_findings_html(by['info'], 'info'), unsafe_allow_html=True)
    if by['pass']:
        with st.expander(f"🟢  {len(by['pass'])} check{'s' if len(by['pass']) != 1 else ''} passed"):
            st.markdown(_findings_html(by['pass'], 'pass'), unsafe_allow_html=True)


def _key_facts(rtype, parsed):
    """A compact 'Alloy · Job · Component · S/N' line for the report header."""
    if rtype == 'metallurgical':
        hdr, smp = parsed.get('header', {}) or {}, parsed.get('sample', {}) or {}
        bits = [('Alloy', smp.get('material')), ('Job', hdr.get('job')),
                ('Component', smp.get('description')), ('S/N', smp.get('serial'))]
    elif rtype == 'coating':
        bits = [('Report', parsed.get('report_no')), ('Component', parsed.get('component'))]
    else:
        return f"Detected report type: **{rtype}**"
    shown = [f"{k}: **{v}**" for k, v in bits if v and str(v).strip()]
    return f"*{rtype}*  ·  " + "  ·  ".join(shown) if shown else f"*{rtype}*"


def _render_annotated(f, rtype, parsed, ocr):
    """Annotated view: an instant drawn-grid render + annotated micrographs, with
    the heavy pixel-faithful LibreOffice render available on demand."""
    data = f.getvalue()
    grid, micrographs = _grid_and_micros(f.name, data, ocr)
    if grid:
        st.image(grid, width="stretch",
                 caption="Quick annotated view — flagged cells boxed and numbered to the legend.")
        st.download_button(
            "⬇ Download annotated view (.png)", data=grid,
            file_name=f"{Path(f.name).stem}_annotated.png",
            mime="image/png", key=f"gridpng_{f.name}")
    elif rtype in ('metallurgical', 'coating'):
        st.caption("Annotated view unavailable (image libraries not installed).")

    # The pixel-faithful render runs LibreOffice (seconds, more on a big report),
    # so build it only when asked — the review above is already on screen.
    if report_render.libreoffice_available():
        fkey = f"faithful_{f.name}"
        if st.button("🖼 Render pixel-faithful view  ·  exact workbook look (slower)",
                     key=f"fbtn_{f.name}"):
            st.session_state[fkey] = True
        if st.session_state.get(fkey):
            with st.spinner("Rendering the exact workbook with LibreOffice…"):
                png, status = _faithful_image(f.name, data, ocr)
            if png:
                st.image(png, width="stretch",
                         caption="Pixel-faithful render — original fonts, layout and embedded micrographs.")
                st.download_button(
                    "⬇ Download pixel-faithful (.png)", data=png,
                    file_name=f"{Path(f.name).stem}_faithful.png",
                    mime="image/png", key=f"fpng_{f.name}")
            else:
                st.caption(f"Pixel-faithful render unavailable — {status}")

    if micrographs:
        st.markdown("**Annotated micrographs** — legend / scale-bar regions boxed; "
                    "contrast and any burned-in thickness flagged.")
        mcols = st.columns(3)
        for i, (mname, mbytes, mcap) in enumerate(micrographs):
            mcols[i % 3].image(mbytes, caption=mcap, width="stretch")

    if rtype in ('metallurgical', 'coating'):
        st.divider()
        if st.button("📁 Add this report's micrographs to the library",
                     key=f"add_{f.name}"):
            added = add_to_library(f.name, data, parsed, rtype)
            if added:
                _gallery_counts.clear()      # let the gallery reflect the add now
                _gallery_photos.clear()
            st.success(f"Added {added} micrograph(s) to the library."
                       if added else "No new micrographs added (already in library).")


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

    ocr = st.checkbox(
        "🔍 Analyse micrographs (legends, etch, thickness)",
        value=True,
        help="Reads each micrograph's burned-in legend (magnification / scale), gauges "
             "etched-vs-low-contrast, and reads any burned-in thickness measurements — "
             "cross-checking against the captions and comment. Requires the Tesseract OCR engine.",
    )

    if not files:
        st.info("Upload one or more `.xlsx` lab reports above to review.")
        return

    for f in files:
        try:
            with st.spinner(f"Reviewing {f.name}…"):
                rtype, parsed, findings = _review(f.name, f.getvalue(), ocr)
        except Exception as e:
            st.error(f"Could not read **{f.name}** — {e}")
            continue

        counts = summarize(findings)
        verdict = ('critical' if counts['critical'] else
                   'warning' if counts['warning'] else 'pass')
        vcolor, vtint, vlabel = _SEV_STYLE[verdict]
        vtext = ("Needs attention — critical findings present." if counts['critical']
                 else "Review recommended — warnings present." if counts['warning']
                 else "Looks good — no warnings or failures.")

        st.write("")
        with st.container(border=True):
            st.markdown(f"#### {f.name}")
            facts = _key_facts(rtype, parsed)
            if facts:
                st.caption(facts)

            # One clear verdict banner (replaces the old pill + duplicate banner).
            vicon = {'critical': '🔴', 'warning': '🟠', 'pass': '🟢'}[verdict]
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:.55rem;padding:.65rem 1rem;'
                f'border-radius:12px;background:{vtint};border:1px solid {vcolor}2e;'
                f'margin:.1rem 0 .7rem;"><span style="font-size:1.1rem;">{vicon}</span>'
                f'<span style="color:{vcolor};font-weight:700;">{vlabel}</span>'
                f'<span style="color:#475569;"> — {vtext}</span></div>',
                unsafe_allow_html=True)

            m = st.columns(4)
            m[0].metric("🔴 Fail", counts['critical'])
            m[1].metric("🟠 Warning", counts['warning'])
            m[2].metric("🔵 Note", counts['info'])
            m[3].metric("🟢 Pass", counts['pass'])

            ftab, atab, dtab = st.tabs(
                ["📋 Findings", "🖼 Annotated view", "🔬 Extracted data"])
            with ftab:
                _render_findings(findings)
            with atab:
                _render_annotated(f, rtype, parsed, ocr)
            with dtab:
                _render_parsed(rtype, parsed)


# ════════════════════════════════════════════════════════════════════════
# TAB 3 — Photo Library (per-alloy micrograph gallery)
# ════════════════════════════════════════════════════════════════════════
def render_gallery():
    st.markdown(
        "Browse stored micrographs **by alloy**, with the data of the report they "
        "came from. Add micrographs from the **Lab Report Review** tab."
    )
    st.caption(f"Storage backend: **{backend_name()}**")
    try:
        counts = _gallery_counts()
    except Exception as e:
        st.error(f"Couldn't reach the photo-library backend (**{backend_name()}**) — "
                 f"the library is unavailable right now. The other tools are unaffected.")
        st.caption(f"{type(e).__name__}: {e}")
        if st.button("↻ Retry", key="gallery_retry"):
            _gallery_counts.clear()
            _gallery_photos.clear()
            st.rerun()
        return
    total = sum(counts.values())
    if not total:
        st.info("The library is empty. Review a report and click "
                "“Add this report's micrographs to the library”.")
        return

    st.caption(f"{total} micrograph(s) across {len(counts)} alloy(s)")
    pick = st.selectbox("Alloy", [f"{a} ({n})" for a, n in sorted(counts.items())])
    alloy = pick.rsplit(" (", 1)[0]
    recs = _gallery_photos(alloy)

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
                for i, f in enumerate(iir_files):
                    # Index-prefixed fixed name — never join the upload's raw name
                    # (path-traversal safety); f.name is kept only as a label.
                    path = os.path.join(tmp, f"iir_{i}.xlsx")
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
_TOOLS = [
    ("🧪", "Lab Report Review", "Automated QA of AEG metallurgical & coating reports",
     render_reviewer),
    ("🔬", "SEM Converter", "Vendor SEM PDF → formatted Ansaldo Word report",
     render_converter),
    ("🖼️", "Photo Library", "Browse stored micrographs by alloy", render_gallery),
    ("🛠️", "IIR Review", "Incoming-inspection report consistency QA", render_iir_tool),
]


def main():
    st.set_page_config(page_title="AEG Materials Tools", page_icon="🔬",
                       layout="wide", initial_sidebar_state="expanded")
    st.markdown(_CSS, unsafe_allow_html=True)

    labels = [f"{ic}  {name}" for ic, name, _, _ in _TOOLS]
    with st.sidebar:
        st.markdown(_BRAND, unsafe_allow_html=True)
        choice = st.radio("nav", labels, label_visibility="collapsed")
    icon, name, sub, fn = _TOOLS[labels.index(choice)]

    st.markdown(_page_header(icon, name, sub), unsafe_allow_html=True)
    try:
        fn()
    except Exception as e:
        st.error("This tool hit an error and couldn't finish.")
        st.caption(f"{type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
