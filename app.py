import streamlit as st
import tempfile
import os
import io
import html as _html
from pathlib import Path
from openpyxl.utils import get_column_letter
from sem_convert import parse, extract_figures, build
from lab_review import review_report, summarize
try:
    from lab_review import HT_ORDER
except ImportError:   # tolerate a stale/cached lab_review module on redeploy
    HT_ORDER = ['As-received', 'Post-solution', 'Post stress-relief', 'Post-ageing', 'Unspecified']
try:
    from lab_review import COMP_WARN_REL, COMP_CRIT_REL
except Exception:     # display-only tolerance hint; authoritative logic lives in lab_review
    COMP_WARN_REL, COMP_CRIT_REL = 10.0, 25.0
import report_render
import iir_review
from photo_lib import (add_to_library, alloy_counts, photos_for,
                       get_image_bytes, backend_name, LIBRARY_DIR)

try:
    import pandas as pd
except Exception:     # pandas ships with Streamlit; guard just in case
    pd = None


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


@st.cache_data(show_spinner=False)
def _parse_iir(name, data):
    """Parse one IIR workbook, cached on the file bytes so tweaking the check
    severities re-checks instantly without re-reading the file."""
    with tempfile.TemporaryDirectory() as tmp:
        # Fixed temp name — never join the upload's raw name to a path.
        path = os.path.join(tmp, "iir.xlsx")
        with open(path, "wb") as fh:
            fh.write(data)
        return iir_review.parse_iir(path)


@st.cache_data(show_spinner=False)
def _ocr_available():
    """Whether the Tesseract engine is actually installed at runtime."""
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


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

  /* Filter pills / segmented control */
  [data-testid="stPills"] button, [data-testid="stButtonGroup"] button {
    border-radius:999px !important; font-weight:600 !important; }
  /* Dataframes read as part of the card */
  [data-testid="stDataFrame"] { border-radius:12px; }
  [data-testid="stPopover"] button { border-radius:11px; font-weight:600; }
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


# ── Shared UI helpers (used by both revamped tools) ──────────────────────
def _chip(label, color, tint, icon=""):
    """A small rounded status pill (returns HTML)."""
    ic = f'{icon} ' if icon else ''
    return (f'<span style="display:inline-flex;align-items:center;gap:.3rem;'
            f'padding:.18rem .62rem;border-radius:999px;background:{tint};color:{color};'
            f'font-weight:700;font-size:.78rem;border:1px solid {color}33;">{ic}{label}</span>')


def _chips(items):
    """items: list of (label, color, tint, icon) → one inline chip row."""
    inner = ' '.join(_chip(l, c, t, i) for l, c, t, i in items)
    return f'<div style="display:flex;gap:.45rem;flex-wrap:wrap;margin:.15rem 0 .5rem;">{inner}</div>'


def _sevbar(segments):
    """A slim proportional severity bar. segments: list of (count, color)."""
    total = sum(c for c, _ in segments) or 1
    cells = ''.join(
        f'<div style="width:{c / total * 100:.4f}%;background:{col};"></div>'
        for c, col in segments if c)
    return ('<div style="display:flex;height:8px;width:100%;border-radius:999px;'
            f'overflow:hidden;background:#eef2f8;margin:.2rem 0 .6rem;">{cells}</div>')


def _finding_rows_html(rows):
    """rows: list of (color, tint, label, category, message) → full-wrap colored rows."""
    out = []
    for color, tint, label, cat, msg in rows:
        out.append(
            f'<div style="display:flex;align-items:baseline;gap:.6rem;'
            f'padding:.45rem .8rem;margin:.32rem 0;border-left:4px solid {color};'
            f'background:{tint};border-radius:8px;">'
            f'<span style="color:{color};font-weight:700;font-size:.68rem;'
            f'letter-spacing:.04em;text-transform:uppercase;min-width:54px;">{_html.escape(label)}</span>'
            f'<span style="line-height:1.45;color:#28323f;">'
            f'<b>{_html.escape(str(cat))}</b> — {_html.escape(str(msg))}</span></div>')
    return '<div>' + ''.join(out) + '</div>'


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
# severity → (accent colour, light tint, label)
_SEV_ORDER = ['critical', 'warning', 'info', 'pass']
_SEV_STYLE = {
    'critical': ('#d62d38', '#fdecee', 'Fail'),
    'warning':  ('#e07b16', '#fdf2e3', 'Warning'),
    'info':     ('#1a6ed6', '#e9f1fc', 'Note'),
    'pass':     ('#1f9e50', '#e8f6ee', 'Pass'),
}
_SEV_LABELS = {'critical': 'Fail', 'warning': 'Warning', 'info': 'Note', 'pass': 'Pass'}
_LAB_EMOJI = {'critical': '🔴', 'warning': '🟠', 'info': '🔵', 'pass': '🟢'}
_LAB_VERDICT = {
    'critical': ('#d62d38', '#fdecee', 'Needs attention'),
    'warning':  ('#e07b16', '#fdf2e3', 'Review recommended'),
    'pass':     ('#1f9e50', '#e8f6ee', 'Looks good'),
}


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


def _lab_rows(findings):
    """Lab (sev, cat, msg) tuples → styled rows, sorted by severity."""
    rank = {s: i for i, s in enumerate(_SEV_ORDER)}
    out = []
    for sev, cat, msg in sorted(findings, key=lambda t: rank.get(t[0], 9)):
        color, tint, label = _SEV_STYLE.get(sev, ('#64748b', '#f1f5f9', sev))
        out.append((color, tint, label, cat, msg))
    return out


def _lab_findings_csv(findings):
    """Findings → CSV bytes (Severity / Category / Finding)."""
    rows = [{'Severity': _SEV_LABELS.get(s, s), 'Category': c, 'Finding': m}
            for s, c, m in findings]
    if pd is not None:
        return pd.DataFrame(rows).to_csv(index=False).encode('utf-8')
    import csv
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=['Severity', 'Category', 'Finding'])
    w.writeheader()
    w.writerows(rows)
    return buf.getvalue().encode('utf-8')


def _lab_findings_tab(findings, name):
    """Filterable findings: severity pills + text search, full-wrap colored rows."""
    if not findings:
        st.caption("No findings.")
        return
    counts = {s: sum(1 for x, _, _ in findings if x == s) for s in _SEV_ORDER}
    sev_opts = [s for s in _SEV_ORDER if counts[s]]
    default_sev = [s for s in ('critical', 'warning', 'info') if counts[s]] or sev_opts

    c1, c2 = st.columns([3, 2])
    with c1:
        picked = st.pills(
            "Severity", sev_opts, selection_mode="multi", default=default_sev,
            format_func=lambda s: f"{_LAB_EMOJI[s]} {_SEV_STYLE[s][2]} ({counts[s]})",
            label_visibility="collapsed", key=f"labsev_{name}")
    query = c2.text_input("Search findings", key=f"labq_{name}",
                          placeholder="🔍 search findings…", label_visibility="collapsed")

    picked = picked or sev_opts
    q = (query or "").lower().strip()
    shown = [(s, c, m) for s, c, m in findings
             if s in picked and (not q or q in f"{c} {m}".lower())]
    if not shown:
        st.caption("No findings match the filter.")
        return
    st.markdown(_finding_rows_html(_lab_rows(shown)), unsafe_allow_html=True)


def _flagged_cells(parsed):
    """Cell-anchored findings (from collect_highlights) as spreadsheet refs."""
    try:
        from lab_review import collect_highlights
        hi = collect_highlights(parsed)
    except Exception:
        return
    if not hi:
        return
    rows = []
    for h in hi:
        cell = h.get('cell') or (None, None)
        r, c = (cell + (None, None))[:2]
        ref = f"{get_column_letter(c)}{r}" if (r and c) else "—"
        rows.append({'Cell': ref,
                     'Severity': _SEV_LABELS.get(h.get('severity'), h.get('severity') or ''),
                     'Category': h.get('category', ''),
                     'Note': h.get('note', '')})
    with st.expander(f"📍 Flagged cells ({len(rows)})"):
        st.caption("Where each cell-anchored finding sits in the workbook — matches the "
                   "boxed & numbered cells in the annotated view.")
        st.dataframe(rows, width="stretch", hide_index=True,
                     column_config={"Note": st.column_config.TextColumn(width="large")})


def _render_annotated(r, ocr):
    """Annotated view: instant drawn-grid render + annotated micrographs, with the
    heavy pixel-faithful LibreOffice render available on demand."""
    f, rtype, parsed = r['f'], r['rtype'], r['parsed']
    data = f.getvalue()
    grid, micrographs = _grid_and_micros(f.name, data, ocr)
    if grid:
        st.image(grid, width="stretch",
                 caption="Quick annotated view — flagged cells boxed and numbered to the legend.")
        st.download_button(
            "⬇ Annotated view (.png)", data=grid,
            file_name=f"{Path(f.name).stem}_annotated.png",
            mime="image/png", key=f"gridpng_{f.name}")
    elif rtype in ('metallurgical', 'coating'):
        st.caption("Annotated view unavailable (image libraries not installed).")

    # The pixel-faithful render runs LibreOffice (seconds, more on a big report),
    # so build it only when asked — the quick view above is already on screen.
    if report_render.libreoffice_available():
        fkey = f"faithful_{f.name}"
        if st.button("🖼 Render pixel-faithful view — exact workbook look (slower)",
                     key=f"fbtn_{f.name}"):
            st.session_state[fkey] = True
        if st.session_state.get(fkey):
            with st.status("Rendering the exact workbook with LibreOffice…", expanded=False) as status:
                png, stat = _faithful_image(f.name, data, ocr)
                status.update(
                    label=("Pixel-faithful render ready" if png
                           else f"Pixel-faithful render unavailable — {stat}"),
                    state=("complete" if png else "error"))
            if png:
                st.image(png, width="stretch",
                         caption="Pixel-faithful render — original fonts, layout and embedded micrographs.")
                st.download_button(
                    "⬇ Pixel-faithful (.png)", data=png,
                    file_name=f"{Path(f.name).stem}_faithful.png",
                    mime="image/png", key=f"fpng_{f.name}")

    if micrographs:
        st.markdown("**Annotated micrographs** — legend / scale-bar regions boxed; "
                    "contrast and any burned-in thickness flagged.")
        mcols = st.columns(3)
        for i, (mname, mbytes, mcap) in enumerate(micrographs):
            mcols[i % 3].image(mbytes, caption=mcap, width="stretch")


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
                if n not in (None, 0) and a is not None:
                    dev_pct = (a - n) / abs(n) * 100
                    dev = f"{dev_pct:+.0f}%"
                    flag = ("🔴" if abs(dev_pct) >= COMP_CRIT_REL
                            else "🟠" if abs(dev_pct) >= COMP_WARN_REL else "")
                else:
                    dev, flag = "—", ""
                rows.append({
                    "": flag,
                    "Element": el,
                    "Nominal": f"{n:g}" if n is not None else "—",
                    "Actual":  f"{a:g}" if a is not None else "—",
                    "Δ":       dev,
                })
            st.dataframe(rows, width="stretch", hide_index=True,
                         column_config={"": st.column_config.TextColumn(width="small")})
            st.caption("Colour is an at-a-glance deviation hint (🟠 ≥ %g%%, 🔴 ≥ %g%%); "
                       "the **Findings** tab holds the authoritative result."
                       % (COMP_WARN_REL, COMP_CRIT_REL))

    elif rtype == 'coating':
        st.write(f"**Report No.:** {parsed.get('report_no') or '—'}")
        st.write(f"**Title:** {parsed.get('title') or '—'}")
        rows = parsed.get('rows', [])
        if rows:
            lo, hi = rows[0].get('min'), rows[0].get('max')
            if lo is not None and hi is not None:
                st.write(f"**Design limit:** {lo:g} – {hi:g} mm")
            st.dataframe([
                {"Row": e['row'], "Measurements (mm)": ", ".join(f"{v:g}" for v in e['values'])}
                for e in rows
            ], width="stretch", hide_index=True,
                column_config={"Measurements (mm)": st.column_config.TextColumn(width="large")})

    legends = parsed.get('legends') or []
    if legends:
        st.markdown("**Micrograph legends — read from the images**")
        st.dataframe([
            {"Image": l.get('image', '—'), "Magnification": l.get('mag', '—'),
             "Scale": l.get('scale', '—'), "Legend ID": l.get('id', '—')}
            for l in legends
        ], width="stretch", hide_index=True)

    _flagged_cells(parsed)


def _render_lab_detail(r, ocr):
    """One report's full detail card: header + actions, verdict bar, metrics, tabs."""
    with st.container(border=True):
        head, actions = st.columns([3, 2])
        with head:
            st.markdown(f"#### {_LAB_EMOJI[r['verdict']]} {r['name']}")
            color, tint, vlabel = _LAB_VERDICT[r['verdict']]
            type_lbl = r['rtype'].capitalize() if r['rtype'] != 'unknown' else 'Unknown type'
            st.markdown(
                _chips([(vlabel, color, tint, _LAB_EMOJI[r['verdict']]),
                        (type_lbl, '#475569', '#eef2f6', '')]),
                unsafe_allow_html=True)
            if r['facts']:
                st.caption(r['facts'])
        with actions:
            b1, b2 = st.columns(2)
            if r['rtype'] in ('metallurgical', 'coating'):
                if b1.button("📁 Add to library", key=f"add_{r['name']}", width="stretch"):
                    added = add_to_library(r['name'], r['f'].getvalue(), r['parsed'], r['rtype'])
                    if added:
                        _gallery_counts.clear()
                        _gallery_photos.clear()
                    st.toast(f"Added {added} micrograph(s) to the library." if added
                             else "No new micrographs (already in library).")
            b2.download_button(
                "⬇ Findings (.csv)", data=_lab_findings_csv(r['findings']),
                file_name=f"{Path(r['name']).stem}_findings.csv", mime="text/csv",
                key=f"labcsv_{r['name']}", width="stretch")

        if r['rtype'] == 'unknown':
            st.warning("This workbook didn't match a metallurgical or coating layout, so only "
                       "a limited review ran. Check it's an AEG lab report `.xlsx`.")

        c = r['counts']
        st.markdown(_sevbar([(c['critical'], '#d62d38'), (c['warning'], '#e07b16'),
                             (c['info'], '#1a6ed6'), (c['pass'], '#1f9e50')]),
                    unsafe_allow_html=True)
        m = st.columns(4)
        m[0].metric("🔴 Fail", c['critical'])
        m[1].metric("🟠 Warning", c['warning'])
        m[2].metric("🔵 Note", c['info'])
        m[3].metric("🟢 Pass", c['pass'])

        ftab, atab, dtab = st.tabs(["📋 Findings", "🖼 Annotated view", "🔬 Extracted data"])
        with ftab:
            _lab_findings_tab(r['findings'], r['name'])
        with atab:
            _render_annotated(r, ocr)
        with dtab:
            _render_parsed(r['rtype'], r['parsed'])


def render_reviewer():
    st.caption(
        "Automated QA of AEG lab reports (`.xlsx`) — **metallurgical** or **coating**. "
        "Each report is checked against its own stated spec (nominal composition / design "
        "limits), plus hardness, completeness and sign-off."
    )

    up, opt = st.columns([3, 1])
    with up:
        files = st.file_uploader(
            "Lab report(s) (.xlsx)", type=["xlsx"], accept_multiple_files=True,
            key="lab_files", label_visibility="collapsed")
    with opt:
        ocr = st.toggle(
            "Analyse micrographs", value=True,
            help="Reads each micrograph's burned-in legend (magnification / scale), gauges "
                 "etched-vs-low-contrast, and reads any burned-in thickness measurements. "
                 "Requires the Tesseract OCR engine.")

    ocr_ok = _ocr_available()
    lo_ok = report_render.libreoffice_available()
    _ok, _off = ('#1f9e50', '#e8f6ee', '✓'), ('#94a3b8', '#f1f5f9', '○')
    st.markdown(_chips([
        ("OCR ready" if ocr_ok else "OCR unavailable", *(_ok if ocr_ok else _off)),
        ("Pixel-faithful ready" if lo_ok else "Grid view only", *(_ok if lo_ok else _off)),
    ]), unsafe_allow_html=True)

    if not files:
        st.info("Upload one or more `.xlsx` lab reports above to review.")
        return

    # Review every file (cached), collecting results for a batch overview.
    reviewed = []
    for f in files:
        try:
            with st.spinner(f"Reviewing {f.name}…"):
                rtype, parsed, findings = _review(f.name, f.getvalue(), ocr)
        except Exception as e:
            reviewed.append({'name': f.name, 'error': str(e)})
            continue
        counts = summarize(findings)
        verdict = ('critical' if counts['critical'] else
                   'warning' if counts['warning'] else 'pass')
        reviewed.append({'f': f, 'name': f.name, 'rtype': rtype, 'parsed': parsed,
                         'findings': findings, 'counts': counts, 'verdict': verdict,
                         'facts': _key_facts(rtype, parsed)})

    for r in [x for x in reviewed if 'error' in x]:
        st.error(f"Could not read **{r['name']}** — {r['error']}")
    ok = [r for r in reviewed if 'error' not in r]
    if not ok:
        return

    # ── Batch overview (only worth showing for more than one report) ──
    if len(ok) > 1:
        with st.container(border=True):
            tot = {k: sum(r['counts'][k] for r in ok)
                   for k in ('critical', 'warning', 'info', 'pass')}
            need = sum(1 for r in ok if r['verdict'] != 'pass')
            st.markdown(f"**Batch overview — {len(ok)} report(s)**")
            m = st.columns(4)
            m[0].metric("Reports", len(ok))
            m[1].metric("🔴 Fail", tot['critical'])
            m[2].metric("🟠 Warning", tot['warning'])
            m[3].metric("Need attention", need)
            st.dataframe([{
                "Verdict": f"{_LAB_EMOJI[r['verdict']]} {_LAB_VERDICT[r['verdict']][2]}",
                "Report": r['name'],
                "Type": r['rtype'],
                "🔴": r['counts']['critical'], "🟠": r['counts']['warning'],
                "🔵": r['counts']['info'], "🟢": r['counts']['pass'],
            } for r in ok], width="stretch", hide_index=True)

            allrows = [{'Report': r['name'], 'Severity': _SEV_LABELS.get(s, s),
                        'Category': c, 'Finding': msg}
                       for r in ok for s, c, msg in r['findings']]
            if pd is not None and allrows:
                st.download_button(
                    "⬇ All findings (.csv)",
                    data=pd.DataFrame(allrows).to_csv(index=False).encode('utf-8'),
                    file_name="lab_review_all_findings.csv", mime="text/csv",
                    key="lab_batch_csv")

        labels = [f"{_LAB_EMOJI[r['verdict']]}  {r['name']}" for r in ok]
        idx = st.selectbox("View report", range(len(ok)),
                           format_func=lambda i: labels[i], key="lab_pick")
        _render_lab_detail(ok[idx], ocr)
    else:
        _render_lab_detail(ok[0], ocr)


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
# ════════════════════════════════════════════════════════════════════════
_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_IIR_STYLE = {
    iir_review.FAIL: ('#d62d38', '#fdecee', 'Fail'),
    iir_review.WARN: ('#e07b16', '#fdf2e3', 'Warn'),
    iir_review.INFO: ('#1a6ed6', '#e9f1fc', 'Info'),
    iir_review.PASS: ('#1f9e50', '#e8f6ee', 'Pass'),
}
_IIR_ORDER = [iir_review.FAIL, iir_review.WARN, iir_review.INFO, iir_review.PASS]
_IIR_CHOICE_LABEL = {iir_review.FAIL: "🔴 Fail", iir_review.WARN: "🟠 Warn",
                     iir_review.INFO: "🔵 Info", iir_review.OFF: "⚪ Off"}


def _iir_catalog_view():
    """Data-driven view of every check, grouped by category (from CHECK_CATALOG)."""
    from itertools import groupby
    st.caption("Every check the review runs. Tune each one's severity — or switch it "
               "off — under **⚙️ Check severities**.")
    for cat, items in groupby(iir_review.CHECK_CATALOG, key=lambda t: t[0]):
        titles = [t[1] for t in items]
        st.markdown(f"**{cat}**")
        st.caption(" · ".join(titles))


def _iir_check_settings():
    """Severity dropdown per check; returns {check_title: severity|OFF}."""
    from itertools import groupby
    with st.expander("⚙️ Check severities — set each check's level or turn it off"):
        st.caption("Changes apply instantly to the overview and verdicts below — no need "
                   "to re-upload.")
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
                    format_func=lambda s: _IIR_CHOICE_LABEL[s],
                    key=f"iir_sev::{title}")
    return chosen


def _iir_filter_table(findings, key, with_report=False):
    """Severity pills + category filter + text search over IIR finding dicts."""
    if not findings:
        st.caption("No findings.")
        return
    cats = sorted({f['category'] for f in findings})
    counts = {s: sum(1 for f in findings if f['severity'] == s) for s in _IIR_ORDER}
    sev_opts = [s for s in _IIR_ORDER if counts[s]]
    default_sev = [s for s in (iir_review.FAIL, iir_review.WARN, iir_review.INFO)
                   if counts[s]] or sev_opts

    c1, c2, c3 = st.columns([3, 2, 2])
    with c1:
        picked = st.pills(
            "Severity", sev_opts, selection_mode="multi", default=default_sev,
            format_func=lambda s: f"{iir_review.SEV_ICON[s]} {_IIR_STYLE[s][2]} ({counts[s]})",
            label_visibility="collapsed", key=f"iirsev_{key}")
    cat_pick = c2.multiselect("Category", cats, placeholder="All categories",
                              label_visibility="collapsed", key=f"iircat_{key}")
    query = c3.text_input("Search", placeholder="🔍 search detail…",
                          label_visibility="collapsed", key=f"iirq_{key}")

    picked = picked or sev_opts
    q = (query or "").lower().strip()
    shown = [f for f in findings
             if f['severity'] in picked
             and (not cat_pick or f['category'] in cat_pick)
             and (not q or q in f"{f['check']} {f['detail']} {f['sheet']} "
                                 f"{f.get('report', '')}".lower())]
    if not shown:
        st.caption("No findings match the filter.")
        return

    rows = []
    for f in shown:
        row = {"Sev": iir_review.SEV_ICON[f['severity']], "Category": f['category'],
               "Check": f['check'], "Sheet": f['sheet'], "Detail": f['detail']}
        if with_report:
            row = {"Report": f['report'], **row}
        rows.append(row)
    st.dataframe(rows, width="stretch", hide_index=True,
                 column_config={
                     "Sev": st.column_config.TextColumn(width="small"),
                     "Detail": st.column_config.TextColumn(width="large")})


def _iir_overview(results):
    """Batch overview card: aggregate metrics + interactive table + batch download."""
    with st.container(border=True):
        tot = {k: sum(r['counts'][k] for r in results) for k in _IIR_ORDER}
        need = sum(1 for r in results
                   if r['template'] == 'unknown'
                   or iir_review.verdict_of(r['counts'])[0] != iir_review.PASS)
        st.markdown(f"**Batch overview — {len(results)} report(s)**")
        m = st.columns(4)
        m[0].metric("Reports", len(results))
        m[1].metric("🔴 Fail", tot[iir_review.FAIL])
        m[2].metric("🟠 Warn", tot[iir_review.WARN])
        m[3].metric("Need attention", need)

        overview = []
        for r in results:
            c = r['counts']
            sev, label = iir_review.verdict_of(c)
            rp = r['rp']
            verdict = ("⚠️ Unknown layout" if r['template'] == 'unknown'
                       else f"{iir_review.SEV_ICON[sev]} {label.split(' — ')[0]}")
            overview.append({
                "Verdict": verdict,
                "Report": r['src'],
                "Doc No": r['ident'].get('doc_no', ''),
                "Component": r['ident'].get('component', ''),
                "Recv": rp.get('received') if rp.get('found') else None,
                "Scrap": rp.get('scrap') if rp.get('found') else None,
                "Recond": rp.get('reconditionable') if rp.get('found') else None,
                "🔴": c[iir_review.FAIL], "🟠": c[iir_review.WARN],
                "🔵": c[iir_review.INFO], "🟢": c[iir_review.PASS],
                "Top issue": iir_review.top_issue(r['findings']),
            })
        st.dataframe(overview, width="stretch", hide_index=True,
                     column_config={"Top issue": st.column_config.TextColumn(width="large")})

        if len(results) > 1:
            records = [iir_review.record_of(r['data'], r['findings']) for r in results]
            buf = io.BytesIO()
            iir_review.build_batch_summary(records, buf)
            st.download_button(
                "⬇ Batch summary (.xlsx)", data=buf.getvalue(),
                file_name="IIR_Batch_Summary.xlsx", mime=_XLSX_MIME,
                key="iir_batch_dl", type="primary")


def _iir_protocol_tab(data):
    """The per-position serial-number protocol grid the checks are derived from."""
    rows = data.get('sn_rows') or []
    if not rows:
        st.caption("No serial-number protocol rows were parsed for this report.")
        return
    st.caption(f"{len(rows)} position(s) — the per-position grid the checks are derived from.")
    st.dataframe([{
        "Pos": r.get('pos'), "Part No": r.get('pn', ''), "Serial": r.get('sn', ''),
        "Scope": r.get('scope', ''), "Scrap": "✓" if r.get('scrap') else "",
        "Defects": ", ".join(r.get('defects') or []),
    } for r in rows], width="stretch", hide_index=True,
        column_config={"Defects": st.column_config.TextColumn(width="large")})
    sums = data.get('sn_sumrow') or {}
    if sums:
        st.caption("Protocol sum row — " + " · ".join(f"{k}: **{v}**" for k, v in sums.items()))


def _iir_extracted_tab(data):
    """Surface the rich parsed context that used to be Excel-only."""
    shown = False
    rp = data.get('received_parts') or {}
    if rp.get('found'):
        shown = True
        st.markdown("**Received-parts table**")
        st.caption(
            f"Rows **{rp.get('rows', '—')}** · Required **{rp.get('required', '—')}** · "
            f"Received **{rp.get('received', '—')}** · Scrap **{rp.get('scrap', '—')}** · "
            f"Reconditionable **{rp.get('reconditionable', '—')}**")
    ft = data.get('findings_tbl') or {}
    if ft:
        shown = True
        st.markdown("**Summary of damages**")
        st.dataframe([{"Finding": k, "Count": v} for k, v in ft.items()],
                     width="stretch", hide_index=True)
    op = data.get('operating') or {}
    if op:
        shown = True
        st.markdown("**Operating data**")
        st.caption(" · ".join(f"{k}: **{v}**" for k, v in op.items()))
    sm = data.get('spares_matrix') or []
    if sm:
        shown = True
        st.markdown("**Expected replacement components**")
        st.dataframe([{
            "Pos": r.get('pos'), "Scrap": "✓" if r.get('scrap') else "",
            "Components": ", ".join(r.get('comps') or []),
        } for r in sm], width="stretch", hide_index=True,
            column_config={"Components": st.column_config.TextColumn(width="large")})
    sl = data.get('spares_list') or []
    if sl:
        shown = True
        st.markdown("**Spare parts list**")
        st.dataframe([{"Part": r.get('part'), "Qty": r.get('qty')} for r in sl],
                     width="stretch", hide_index=True)
    photos = data.get('photos') or []
    if photos:
        shown = True
        st.markdown("**Photo sheets**")
        st.dataframe([{
            "Sheet": p.get('sheet'), "Captions": len(p.get('captions') or []),
            "Images": p.get('images'),
        } for p in photos], width="stretch", hide_index=True)
    footers = data.get('footers') or []
    if footers:
        shown = True
        with st.expander(f"Page footers ({len(footers)})"):
            frows = []
            for fr in footers:
                fr = list(fr) + [None, None, None]
                frows.append({"Sheet": fr[0], "Page": fr[1], "Of": fr[2]})
            st.dataframe(frows, width="stretch", hide_index=True)
    if not shown:
        st.caption("No additional extracted data available for this report.")


def _iir_report_card(r):
    """One report's detail card: identity, template, metrics, tabs, checklist."""
    data, ident, c = r['data'], r['ident'], r['counts']
    with st.container(border=True):
        sev, _ = iir_review.verdict_of(c)
        if r['template'] == 'unknown':
            st.error(
                f"**Unrecognized IIR layout** — “{r['src']}” didn't match a known template, "
                "so the review couldn't be fully scored. Confirm it's a Detailed Assessment "
                "Customer Report `.xlsx`.")
        badge = "⚠️" if r['template'] == 'unknown' else iir_review.SEV_ICON[sev]
        st.markdown(f"#### {badge} {r['src']}")
        fam = "?" if r['template'] in ('unknown', '?') else r['template']
        st.caption(
            f"**{ident.get('doc_no', '?')}** · {ident.get('customer', '?')} · "
            f"{ident.get('component', '?')}  —  prepared by "
            f"{ident.get('preparer') or ident.get('author', '?')}, "
            f"reviewed by {ident.get('reviewer', '?')}  ·  Template {fam}")

        rp = r['rp']
        m = st.columns(4)
        m[0].metric("Received", rp.get('received') if rp.get('found') else "—")
        m[1].metric("Scrap", rp.get('scrap') if rp.get('found') else "—")
        m[2].metric("Reconditionable", rp.get('reconditionable') if rp.get('found') else "—")
        m[3].metric("Positions", r['npos'])

        st.markdown(_sevbar([(c[iir_review.FAIL], '#d62d38'), (c[iir_review.WARN], '#e07b16'),
                             (c[iir_review.INFO], '#1a6ed6'), (c[iir_review.PASS], '#1f9e50')]),
                    unsafe_allow_html=True)
        m2 = st.columns(4)
        m2[0].metric("🔴 Fail", c[iir_review.FAIL])
        m2[1].metric("🟠 Warn", c[iir_review.WARN])
        m2[2].metric("🔵 Info", c[iir_review.INFO])
        m2[3].metric("🟢 Pass", c[iir_review.PASS])

        ftab, ptab, xtab = st.tabs(["📋 Findings", "🔢 Protocol", "🗂 Extracted data"])
        with ftab:
            _iir_filter_table(r['findings'], key=f"rep_{r['src']}")
        with ptab:
            _iir_protocol_tab(data)
        with xtab:
            _iir_extracted_tab(data)

        buf = io.BytesIO()
        iir_review.build_checklist(data, r['findings'], buf)
        st.download_button(
            "⬇ Findings checklist (.xlsx)", data=buf.getvalue(),
            file_name=f"IIR_Review_{Path(r['src']).stem}.xlsx",
            mime=_XLSX_MIME, key=f"iir_dl_{r['src']}")


def render_iir_tool():
    st.caption(
        "Automated consistency & completeness QA of **Incoming Inspection Reports** "
        "(Detailed Assessment Customer Reports, `.xlsx`). Upload one or more workbooks for a "
        "severity-tagged findings checklist and a batch overview."
    )

    top = st.columns([3, 1])
    with top[0]:
        iir_files = st.file_uploader(
            "IIR workbook(s) (.xlsx)", type=["xlsx"], accept_multiple_files=True,
            key="iir_uploader", label_visibility="collapsed")
    with top[1]:
        with st.popover("ℹ️ What it checks", use_container_width=True):
            _iir_catalog_view()

    if not iir_files:
        st.info("Upload one or more IIR Excel files above to run the review.")
        return

    # Auto-run: parse every workbook on upload (cached on bytes) — no separate button.
    parsed_items, perrs = [], []
    with st.spinner(f"Reading {len(iir_files)} report(s)…"):
        for f in iir_files:
            try:
                parsed_items.append({'src': f.name, 'data': _parse_iir(f.name, f.getvalue())})
            except Exception as e:
                perrs.append(f"{f.name}: {e}")
    for e in perrs:
        st.error(f"Could not read — {e}")
    if not parsed_items:
        return

    overrides = _iir_check_settings()

    # Re-run the checks for every report with the current severity settings (cheap).
    results = []
    for item in parsed_items:
        data = item['data']
        findings = iir_review.run_checks(data, overrides)
        results.append({
            'src': item['src'], 'data': data, 'findings': findings,
            'counts': iir_review.count_severities(findings),
            'ident': data['ident'], 'rp': data['received_parts'],
            'npos': len(data['sn_rows']), 'template': data.get('template', '?'),
        })

    _iir_overview(results)

    # ── Pooled findings across the whole batch (was Excel-only before) ──
    pooled = [{**f, 'report': r['src']} for r in results for f in r['findings']]
    if pooled:
        with st.container(border=True):
            st.markdown("**All findings across the batch**")
            _iir_filter_table(pooled, key="pool", with_report=True)

    # ── Per-report drill-down ──
    st.markdown("### Report detail")
    if len(results) > 1:
        def _label(i):
            r = results[i]
            badge = ("⚠️" if r['template'] == 'unknown'
                     else iir_review.SEV_ICON[iir_review.verdict_of(r['counts'])[0]])
            return f"{badge}  {r['src']}"
        idx = st.selectbox("View report", range(len(results)),
                           format_func=_label, key="iir_pick")
        _iir_report_card(results[idx])
    else:
        _iir_report_card(results[0])


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

    labels = [f"{ic}  {name}" for ic, name, _, _ in _TOOLS]
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
