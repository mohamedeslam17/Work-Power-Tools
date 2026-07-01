#!/usr/bin/env python3
"""
Annotated review images for the Lab Report Reviewer.

Two products, both pure-Pillow (degrade to None / [] when Pillow is absent):

  * render_report_image(...) — draws the report's data region as a
    spreadsheet-like grid, boxes and numbers the cells that triggered findings,
    and bakes a numbered legend underneath. The "issue areas, highlighted and
    annotated" view.
  * annotate_micrographs(...) — boxes each embedded micrograph's burned-in
    legend / scale-bar regions, flags low contrast and surfaces any thickness
    measurements read from it.

Cell anchoring comes from lab_review.collect_highlights(); this module only
draws. It never raises into the caller — on any trouble it returns None / [].
"""
import io
import os
import shutil
import signal
import subprocess
import tempfile
import textwrap
import threading
import zipfile

# Serialize LibreOffice conversions: concurrent headless soffice on a small host
# collides on its profile and can OOM. One render at a time.
_LO_LOCK = threading.Lock()

try:
    from PIL import Image, ImageChops, ImageDraw, ImageFont
    _PIL = True
except Exception:                       # pragma: no cover - exercised by guard
    _PIL = False

try:
    import fitz                          # PyMuPDF — rasterises the LibreOffice PDF
    _FITZ = True
except Exception:                       # pragma: no cover - exercised by guard
    _FITZ = False

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, PatternFill, Side

from lab_review import collect_highlights, analyze_images

# ── Palette ───────────────────────────────────────────────────────────────
_SEV_RGB = {
    'critical': (214, 45, 56),     # red
    'warning':  (236, 134, 18),    # orange
    'info':     (24, 110, 214),    # blue
    'pass':     (32, 160, 80),     # green
}
_SEV_RANK = {'critical': 3, 'warning': 2, 'info': 1, 'pass': 0}
_WHITE   = (255, 255, 255)
_GRID    = (203, 207, 214)
_HDR_BG  = (240, 242, 246)
_TEXT    = (28, 32, 38)
_MUTED   = (118, 124, 132)
_TITLE_BG = (37, 47, 62)

_FONT_DIRS = ('/usr/share/fonts/truetype/dejavu/', '/usr/share/fonts/dejavu/')


def _font(size, bold=False):
    name = 'DejaVuSans-Bold.ttf' if bold else 'DejaVuSans.ttf'
    for d in _FONT_DIRS:
        try:
            return ImageFont.truetype(d + name, size)
        except Exception:
            continue
    try:
        return ImageFont.truetype(name, size)
    except Exception:
        return ImageFont.load_default()


def _textw(draw, text, font):
    try:
        return draw.textlength(text, font=font)
    except Exception:
        return len(text) * (font.size * 0.6 if hasattr(font, 'size') else 7)


def _fit(draw, text, font, maxw):
    """Truncate `text` with an ellipsis so it fits within maxw pixels."""
    if _textw(draw, text, font) <= maxw:
        return text
    ell = '…'
    while text and _textw(draw, text + ell, font) > maxw:
        text = text[:-1]
    return (text + ell) if text else ''


# ════════════════════════════════════════════════════════════════════════
# THE ANNOTATED REPORT GRID
# ════════════════════════════════════════════════════════════════════════
def render_report_image(data, parsed, findings=None, rtype=None, filename=None,
                        max_rows=48, max_cols=18, scale=2):
    """Return PNG bytes of the annotated report grid, or None.

    `data` is the workbook bytes; `parsed` is lab_review's parsed dict (it must
    carry the 'loc' map). `scale` super-samples the canvas for crisp text.
    """
    if not _PIL:
        return None
    try:
        return _render(data, parsed, findings, filename, max_rows, max_cols, scale)
    except Exception:
        return None


def _render(data, parsed, findings, filename, max_rows, max_cols, scale):
    loc = parsed.get('loc') or {}
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    sheet_name = loc.get('sheet')
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

    highlights = [h for h in collect_highlights(parsed) if h.get('cell')]

    # Merged ranges → covered cells + anchor span lookup.
    spans = {}            # (r, c) of every covered cell → (r0, c0, r1, c1)
    for rng in ws.merged_cells.ranges:
        box = (rng.min_row, rng.min_col, rng.max_row, rng.max_col)
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                spans[(r, c)] = box

    def cell_text(r, c):
        v = ws.cell(row=r, column=c).value
        return '' if v is None else str(v).strip()

    # Content bounding box (cells with text), unioned with every highlight cell.
    rmin = cmin = 10**9
    rmax = cmax = 0
    seen = False
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 200),
                            max_col=min(ws.max_column, 60)):
        for cell in row:
            if cell_text(cell.row, cell.column):
                seen = True
                rmin, cmin = min(rmin, cell.row), min(cmin, cell.column)
                rmax, cmax = max(rmax, cell.row), max(cmax, cell.column)
    for h in highlights:
        r, c = h['cell']
        rmin, cmin = min(rmin, r), min(cmin, c)
        rmax, cmax = max(rmax, r), max(cmax, c)
        seen = True
    if not seen:
        return None
    r1 = min(rmax, rmin + max_rows - 1)
    c1 = min(cmax, cmin + max_cols - 1)
    rows = list(range(rmin, r1 + 1))
    cols = list(range(cmin, c1 + 1))

    # Measure column widths from their content.
    probe = Image.new('RGB', (8, 8))
    pd = ImageDraw.Draw(probe)
    f_cell = _font(13)
    f_hdr  = _font(12, bold=True)
    PAD = 8
    cw = {}
    for c in cols:
        w = _textw(pd, get_column_letter(c), f_hdr) + 10
        for r in rows:
            if spans.get((r, c), (r, c, r, c))[:2] != (r, c) and (r, c) in spans:
                continue                       # measured at the anchor instead
            w = max(w, _textw(pd, cell_text(r, c), f_cell) + 2 * PAD + 6)
        cw[c] = int(max(52, min(240, w)))
    RH = 30
    GUT = 46                                   # row-number gutter
    TH = 24                                    # column-letter header band
    TITLE_H = 58

    # x/y geometry.
    col_x = {}
    x = GUT
    for c in cols:
        col_x[c] = x
        x += cw[c]
    grid_right = x
    row_y = {}
    y = TITLE_H + TH
    for r in rows:
        row_y[r] = y
        y += RH
    grid_bottom = y

    def x_right(c):
        return col_x[c] + cw[c] if c in col_x else grid_right

    def rect_for(r, c):
        """Pixel rect (x0, y0, x1, y1) for a cell, honouring merges + clipping."""
        r0, c0, r2, c2 = spans.get((r, c), (r, c, r, c))
        r0, r2 = max(r0, rmin), min(r2, r1)
        c0, c2 = max(c0, cmin), min(c2, c1)
        x0 = col_x.get(c0, GUT)
        x1 = x_right(c2)
        y0 = row_y.get(r0, TITLE_H + TH)
        y1 = row_y.get(r2, grid_bottom - RH) + RH
        return x0, y0, x1, y1

    # ── Legend: one entry per distinct finding note (a finding may box >1 cell) ──
    order, meta = [], {}
    for h in highlights:
        if h['note'] not in meta:
            meta[h['note']] = (h['severity'], h['category'])
            order.append(h['note'])
    num = {note: i for i, note in enumerate(order, 1)}

    content_w = max(grid_right, 760)
    wrap_chars = max(40, int((content_w - 70) / 8))
    leg_lines = []
    for note in order:
        sev, cat = meta[note]
        wrapped = textwrap.wrap(f"{cat} — {note}", wrap_chars) or ['']
        leg_lines.append((num[note], sev, wrapped))

    # Warning/critical findings not represented by a box (filename, photo
    # legends, caption numbering, …) — listed so nothing is silently missing.
    boxed_notes = {h['note'] for h in highlights}
    extra_lines = []
    for sev, cat, msg in (findings or []):
        if sev in ('critical', 'warning') and msg not in boxed_notes:
            extra_lines.append((sev, textwrap.wrap(f"{cat} — {msg}", wrap_chars) or ['']))

    def _block_h(lines):
        return 28 + sum(22 + 18 * (len(w) - 1) + 8 for *_, w in lines)

    if order or extra_lines:
        leg_h = 16
        if order:
            leg_h += _block_h(leg_lines)
        if extra_lines:
            leg_h += _block_h(extra_lines)
    else:
        leg_h = 44
    W = int(content_w + 16)
    H = int(grid_bottom + 18 + leg_h + 12)

    img = Image.new('RGB', (W * scale, H * scale), _WHITE)
    d = ImageDraw.Draw(img)
    f_title = _font(20 * scale, bold=True)
    f_sub   = _font(13 * scale)
    f_cellS = _font(13 * scale)
    f_hdrS  = _font(12 * scale, bold=True)
    f_legS  = _font(14 * scale)
    f_legbS = _font(14 * scale, bold=True)
    f_badge = _font(13 * scale, bold=True)

    def S(v):                                   # scale a coordinate
        return int(v * scale)

    # Title band.
    d.rectangle([0, 0, W * scale, S(TITLE_H)], fill=_TITLE_BG)
    title = filename or sheet_name or 'Lab report'
    d.text((S(14), S(11)), _fit(d, title, f_title, (W - 28) * scale),
           font=f_title, fill=_WHITE)
    counts = {}
    for h in highlights:
        counts[h['severity']] = counts.get(h['severity'], 0) + 1
    sub = "Annotated review — " + (
        ", ".join(f"{counts[s]} {s}" for s in ('critical', 'warning', 'info')
                  if counts.get(s)) or "no cell-level issues flagged")
    d.text((S(15), S(38)), sub, font=f_sub, fill=(200, 208, 220))

    # Column-letter header band.
    d.rectangle([S(GUT), S(TITLE_H), S(grid_right), S(TITLE_H + TH)], fill=_HDR_BG)
    for c in cols:
        cx = S(col_x[c] + cw[c] / 2)
        d.text((cx, S(TITLE_H + 5)), get_column_letter(c), font=f_hdrS,
               fill=_MUTED, anchor='ma')
    # Row-number gutter.
    d.rectangle([0, S(TITLE_H + TH), S(GUT), S(grid_bottom)], fill=_HDR_BG)
    for r in rows:
        d.text((S(GUT - 6), S(row_y[r] + RH / 2)), str(r), font=f_hdrS,
               fill=_MUTED, anchor='rm')

    # Cells (skip merge-covered non-anchors).
    drawn = set()
    for r in rows:
        for c in cols:
            box = spans.get((r, c))
            if box and box[:2] != (r, c):
                continue
            x0, y0, x1, y1 = rect_for(r, c)
            if (x0, y0, x1, y1) in drawn:
                continue
            drawn.add((x0, y0, x1, y1))
            d.rectangle([S(x0), S(y0), S(x1), S(y1)], outline=_GRID, width=1)
            txt = cell_text(r, c)
            if txt:
                d.text((S(x0 + PAD), S((y0 + y1) / 2)),
                       _fit(d, txt, f_cellS, (x1 - x0 - 2 * PAD) * scale),
                       font=f_cellS, fill=_TEXT, anchor='lm')

    # Highlight boxes + badges. A finding may box several cells (so the badge
    # number ties back to one legend line); a cell may carry several numbers.
    # The badge sits in the cell's top-right corner — usually clear of the value.
    by_cell = {}
    for h in highlights:
        by_cell.setdefault(h['cell'], set()).add((num[h['note']], h['severity']))
    for cell, items in by_cell.items():
        sev = max((s for _, s in items), key=lambda s: _SEV_RANK[s])
        color = _SEV_RGB[sev]
        x0, y0, x1, y1 = rect_for(*cell)
        d.rounded_rectangle([S(x0) + 1, S(y0) + 1, S(x1) - 1, S(y1) - 1],
                            radius=4 * scale, outline=color, width=max(2, scale + 1))
        label = ",".join(str(n) for n, _ in sorted(items))
        bh = 16 * scale
        bw = _textw(d, label, f_badge) + 9 * scale
        bx = max(S(x0) + 1, S(x1) - bw - 1)
        by = S(y0) + 1
        d.rounded_rectangle([bx, by, bx + bw, by + bh], radius=3 * scale, fill=color)
        d.text((bx + bw / 2, by + bh / 2), label, font=f_badge, fill=_WHITE, anchor='mm')

    # Legend panel.
    ly = grid_bottom + 16
    if order or extra_lines:
        if order:
            d.text((S(14), S(ly)), "Findings", font=f_legbS, fill=_TEXT)
            ly += 26
            for i, sev, wrapped in leg_lines:
                color = _SEV_RGB[sev]
                d.rounded_rectangle([S(14), S(ly), S(34), S(ly + 18)],
                                    radius=4 * scale, fill=color)
                d.text((S(24), S(ly + 9)), str(i), font=f_badge, fill=_WHITE, anchor='mm')
                d.text((S(44), S(ly)), wrapped[0], font=f_legS, fill=_TEXT)
                for ex in wrapped[1:]:
                    ly += 18
                    d.text((S(44), S(ly)), ex, font=f_legS, fill=_MUTED)
                ly += 26
        if extra_lines:
            d.text((S(14), S(ly)), "Also flagged (not tied to a cell)",
                   font=f_legbS, fill=_TEXT)
            ly += 26
            for sev, wrapped in extra_lines:
                d.ellipse([S(17), S(ly + 4), S(31), S(ly + 18)], fill=_SEV_RGB[sev])
                d.text((S(44), S(ly)), wrapped[0], font=f_legS, fill=_TEXT)
                for ex in wrapped[1:]:
                    ly += 18
                    d.text((S(44), S(ly)), ex, font=f_legS, fill=_MUTED)
                ly += 26
    else:
        d.text((S(14), S(ly)), "No cell-level issues to highlight on this sheet.",
               font=f_legS, fill=_MUTED)

    out = io.BytesIO()
    img.save(out, format='PNG')
    return out.getvalue()


# ════════════════════════════════════════════════════════════════════════
# ANNOTATED MICROGRAPHS
# ════════════════════════════════════════════════════════════════════════
def annotate_micrographs(data, parsed, max_images=12):
    """Return [(image_name, png_bytes, caption)] with issue regions boxed.

    Uses the per-image analysis already on parsed['images'] when present, else
    runs analyze_images() here. Boxes the burned-in legend and scale-bar regions,
    flags low contrast, and surfaces any thickness measurements read.
    """
    if not _PIL:
        return []
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except Exception:
        return []
    media = {n.split('/')[-1]: n for n in z.namelist() if n.startswith('xl/media')}

    images = parsed.get('images')
    if not images:
        images, _ = analyze_images(data, max_images=max_images)

    out = []
    for entry in images[:max_images]:
        name = entry.get('image')
        if name not in media:
            continue
        try:
            im = Image.open(io.BytesIO(z.read(media[name]))).convert('RGB')
        except Exception:
            continue
        out.append((name, _annotate_one(im, entry), _micro_caption(entry)))
    return out


def _micro_caption(entry):
    bits = []
    if entry.get('mag'):
        bits.append(entry['mag'])
    if entry.get('id'):
        bits.append(entry['id'])
    if entry.get('scale'):
        bits.append(f"scale {entry['scale']}")
    bits.append("etched" if entry.get('etched') else "low-contrast")
    meas = entry.get('measurements') or []
    if meas:
        bits.append("thickness " + ", ".join(f"{m} µm" for m in meas))
    return "  ·  ".join(bits)


def _annotate_one(im, entry):
    w, h = im.size
    big = max(1, 900 // max(1, w))                 # upsample small micrographs a touch
    if big > 1:
        im = im.resize((w * big, h * big))
        w, h = im.size
    d = ImageDraw.Draw(im, 'RGBA')
    f = _font(max(13, w // 42), bold=True)
    fs = _font(max(12, w // 48))

    def banner(xy, text, color):
        tw = _textw(d, text, fs) + 14
        d.rectangle([xy[0], xy[1], xy[0] + tw, xy[1] + fs.size + 10], fill=color + (235,))
        d.text((xy[0] + 7, xy[1] + 5), text, font=fs, fill=_WHITE)

    # Legend region (bottom-left) — where the ID / magnification is read.
    lr = (2, int(h * 0.90), int(w * 0.55), h - 2)
    d.rectangle(lr, outline=_SEV_RGB['warning'], width=max(2, w // 360))
    tag = "legend"
    if entry.get('mag') or entry.get('id'):
        tag += ": " + (entry.get('id') or entry.get('mag'))
    banner((lr[0], max(0, lr[1] - fs.size - 12)), tag, _SEV_RGB['warning'])

    # Scale-bar region (bottom-right).
    sr = (int(w * 0.72), int(h * 0.88), w - 2, h - 2)
    d.rectangle(sr, outline=_SEV_RGB['info'], width=max(2, w // 360))
    banner((sr[0], max(0, sr[1] - fs.size - 12)),
           "scale" + (f": {entry['scale']}" if entry.get('scale') else ""),
           _SEV_RGB['info'])

    # Contrast / thickness call-outs along the top.
    y = 4
    if entry.get('strong') is not None and not entry.get('etched'):
        banner((4, y), "low contrast — verify etch state", _SEV_RGB['critical'])
        y += fs.size + 14
    meas = entry.get('measurements') or []
    if meas:
        banner((4, y), "thickness read: " + ", ".join(f"{m} µm" for m in meas),
               _SEV_RGB['pass'])

    out = io.BytesIO()
    im.save(out, format='PNG')
    return out.getvalue()


# ════════════════════════════════════════════════════════════════════════
# PIXEL-FAITHFUL VIEW  (LibreOffice render of the real workbook + overlays)
# ════════════════════════════════════════════════════════════════════════
# How it works: the flagged cells are filled in the workbook with an
# exactly-known colour (a severity tint, perturbed per finding so each is
# unique) and given a medium border in the severity colour. LibreOffice then
# renders the *real* workbook — original fonts, column widths, borders, merged
# cells and embedded micrographs intact — to PDF, which PyMuPDF rasterises.
# Each finding's cell is then found back by its unique fill colour (no fragile
# cell→pixel maths), so a numbered badge lands exactly on it and a legend ties
# the numbers to the findings. Falls back (caller's job) to the drawn grid when
# LibreOffice or PyMuPDF is unavailable.

_TINT = {                       # light fills (kept readable over the cell text)
    'critical': (255, 214, 217),
    'warning':  (255, 232, 200),
    'info':     (210, 228, 250),
    'pass':     (212, 240, 222),
}
_BORDER_HEX = {'critical': 'C62D38', 'warning': 'D9821A',
               'info': '1A6ED6', 'pass': '1F9E50'}


def libreoffice_available():
    """True when a pixel-faithful render is possible here."""
    return bool(_PIL and _FITZ and _find_soffice())


def _find_soffice():
    for name in ('soffice', 'libreoffice'):
        p = shutil.which(name)
        if p:
            return p
    for p in ('/usr/bin/soffice', '/usr/bin/libreoffice',
              '/opt/libreoffice/program/soffice'):
        if os.path.exists(p):
            return p
    return None


def _unique_fill(i, sev):
    """A severity-tinted fill that is unique per finding index (so it can be
    found back in the raster by exact colour). Steps of 6 keep masks (±2) apart."""
    r, g, b = _TINT.get(sev, _TINT['warning'])
    b = max(140, b - (i % 8) * 6)
    g = max(140, g - (i // 8) * 6)
    return (r, g, b)


def render_report_faithful(data, parsed, findings=None, filename=None, dpi=130, timeout=90):
    """Return (png_bytes, status). status is 'ok' or a short reason on failure
    (so the caller can fall back to render_report_image). `findings` lets the
    legend also list warning/critical findings that aren't tied to a cell."""
    if not _PIL:
        return None, 'Pillow unavailable'
    if not _FITZ:
        return None, 'PyMuPDF unavailable'
    if not _find_soffice():
        return None, 'LibreOffice not installed'
    try:
        return _faithful(data, parsed, findings, filename, dpi, timeout)
    except subprocess.TimeoutExpired:
        return None, 'LibreOffice timed out'
    except Exception as e:
        return None, f'{type(e).__name__}: {e}'


def _faithful(data, parsed, findings, filename, dpi, timeout):
    loc = parsed.get('loc') or {}
    highlights = sorted((h for h in collect_highlights(parsed) if h.get('cell')),
                        key=lambda h: (h['cell'][0], h['cell'][1]))

    wb = openpyxl.load_workbook(io.BytesIO(data))      # keep styles + images
    sheet = loc.get('sheet')
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.active

    span = {}
    for rng in ws.merged_cells.ranges:
        for rr in range(rng.min_row, rng.max_row + 1):
            for cc in range(rng.min_col, rng.max_col + 1):
                span[(rr, cc)] = rng

    # Group highlights by cell: one fill + one badge per cell, even when a cell
    # carries several findings (e.g. unetched note + contrast mismatch).
    groups, by_cell = [], {}
    for h in highlights:
        cell = tuple(h['cell'])
        g = by_cell.get(cell)
        if g is None:
            g = {'cell': cell, 'severity': h['severity'], 'notes': [h['note']]}
            by_cell[cell] = g
            groups.append(g)
        else:
            g['notes'].append(h['note'])
            if _SEV_RANK[h['severity']] > _SEV_RANK[g['severity']]:
                g['severity'] = h['severity']

    keys = []
    for i, g in enumerate(groups):
        r, c = g['cell']
        sev = g['severity']
        rgb = _unique_fill(i, sev)
        fill = PatternFill('solid', fgColor='%02X%02X%02X' % rgb)
        side = Side(style='medium', color=_BORDER_HEX.get(sev, 'D9821A'))
        border = Border(left=side, right=side, top=side, bottom=side)
        rng = span.get((r, c))
        cells = ([(rr, cc) for rr in range(rng.min_row, rng.max_row + 1)
                  for cc in range(rng.min_col, rng.max_col + 1)]
                 if rng else [(r, c)])
        for rr, cc in cells:
            cell = ws.cell(row=rr, column=cc)
            cell.fill = fill
            cell.border = border
        keys.append({'num': i + 1, 'severity': sev, 'notes': g['notes'],
                     'rgb': rgb, 'ref': f'{get_column_letter(c)}{r}'})

    # Warning/critical findings not represented by a box (filename, photo
    # legends, caption numbering, …) — listed so nothing is silently missing.
    boxed_notes = {h['note'] for h in highlights}
    extras = [(sev, cat, msg) for (sev, cat, msg) in (findings or [])
              if sev in ('critical', 'warning') and msg not in boxed_notes]

    # Bound output to the used range and fit to one page wide.
    try:
        from openpyxl.worksheet.properties import PageSetupProperties
        ws.print_area = ws.dimensions
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    except Exception:
        pass

    with tempfile.TemporaryDirectory() as tmp:
        xpath = os.path.join(tmp, 'annotated.xlsx')
        wb.save(xpath)
        pdf = _xlsx_to_pdf(xpath, tmp, timeout)
        if not pdf:
            return None, 'LibreOffice conversion failed'
        pages = _raster_pdf(pdf, dpi)
    if not pages:
        return None, 'no pages rendered'
    return _compose_faithful(pages, keys, extras, filename, dpi), 'ok'


def _xlsx_to_pdf(xlsx_path, outdir, timeout):
    exe = _find_soffice()
    profile = 'file://' + os.path.join(outdir, 'lo_profile')
    cmd = [exe, '--headless', '--norestore', '--invisible', '--nologo',
           '-env:UserInstallation=' + profile,
           '--convert-to', 'pdf:calc_pdf_Export', '--outdir', outdir, xlsx_path]
    with _LO_LOCK:
        # start_new_session so a timeout can kill the whole soffice process group
        # (soffice forks soffice.bin — killing just the launcher leaves a zombie).
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                start_new_session=True)
        try:
            proc.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
            except Exception:
                proc.kill()
            proc.communicate()
            raise
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(xlsx_path))[0] + '.pdf')
    return pdf if os.path.exists(pdf) else None


def _raster_pdf(pdf, dpi):
    try:                       # quieten benign 'structure tree' notices on tagged PDFs
        fitz.TOOLS.mupdf_display_errors(False)
    except Exception:
        pass
    doc = fitz.open(pdf)
    mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pages = []
    for pg in doc:
        pix = pg.get_pixmap(matrix=mat, alpha=False)
        pages.append(Image.frombytes('RGB', (pix.width, pix.height), pix.samples).copy())
    doc.close()
    return pages


def _color_bbox(img, rgb, tol=2):
    """Bounding box of pixels matching rgb within tol on every channel, or None."""
    bands = []
    for ch, v in zip(img.split()[:3], rgb):
        lo, hi = v - tol, v + tol
        bands.append(ch.point(lambda p: 255 if lo <= p <= hi else 0))
    mask = ImageChops.multiply(ImageChops.multiply(bands[0], bands[1]), bands[2])
    return mask.getbbox()


def _compose_faithful(pages, keys, extras, filename, dpi):
    margin = 16
    gap = 12
    content_w = max(p.width for p in pages)
    page_w = content_w + 2 * margin

    # Locate every flagged cell on whichever page carries its unique fill.
    located = {}                                       # num -> (page_index, bbox)
    for k in keys:
        for pi, pg in enumerate(pages):
            bb = _color_bbox(pg, k['rgb'])
            if bb:
                located[k['num']] = (pi, bb)
                break

    fsz = max(15, int(content_w / 78))
    f_title = _font(int(fsz * 1.5), bold=True)
    f_sub   = _font(int(fsz * 0.95))
    f_leg   = _font(fsz)
    f_legb  = _font(fsz, bold=True)
    f_badge = _font(int(fsz * 0.95), bold=True)
    title_h = int(fsz * 3.2)
    line_h = int(fsz * 1.5)
    entry_gap = int(fsz * 0.6)
    head_h = int(fsz * 1.8)
    wrap_chars = max(48, int((page_w - 90) / (fsz * 0.56)))

    def wrap_notes(notes):
        lines = []
        for note in notes:
            lines += textwrap.wrap(note, wrap_chars) or ['']
        return lines or ['']

    # Boxed findings (numbered) then report-level extras (bulleted).
    key_entries = []
    for k in keys:
        notes = list(k['notes'])
        notes[0] = f"[{k['ref']}]  {notes[0]}"
        key_entries.append({'badge': str(k['num']), 'severity': k['severity'],
                            'lines': wrap_notes(notes)})
    extra_entries = [{'badge': None, 'severity': sev, 'lines': wrap_notes([f'{cat} — {msg}'])}
                     for sev, cat, msg in extras]

    def block_h(entries):
        return sum(line_h * len(e['lines']) + entry_gap for e in entries)

    leg_h = entry_gap
    if key_entries:
        leg_h += head_h + block_h(key_entries)
    if extra_entries:
        leg_h += head_h + block_h(extra_entries)
    if not key_entries and not extra_entries:
        leg_h += int(fsz * 2)

    # Page vertical offsets in the stacked canvas.
    y = title_h + margin
    offs = []
    for pg in pages:
        offs.append(y)
        y += pg.height + gap
    total_h = y - gap + margin + leg_h

    canvas = Image.new('RGB', (page_w, total_h), _WHITE)
    d = ImageDraw.Draw(canvas)

    # Title band.
    d.rectangle([0, 0, page_w, title_h], fill=_TITLE_BG)
    d.text((margin, int(fsz * 0.5)),
           _fit(d, filename or 'Lab report', f_title, page_w - 2 * margin),
           font=f_title, fill=_WHITE)
    counts = {}
    for k in keys:
        counts[k['severity']] = counts.get(k['severity'], 0) + 1
    sub = 'Pixel-faithful annotated review — ' + (
        ', '.join(f'{counts[s]} {s}' for s in ('critical', 'warning', 'info')
                  if counts.get(s)) or 'no cell-level issues flagged')
    d.text((margin + 1, int(fsz * 2.0)), sub, font=f_sub, fill=(200, 208, 220))

    # Paste the faithful pages, framed.
    for pg, oy in zip(pages, offs):
        canvas.paste(pg, (margin, oy))
        d.rectangle([margin, oy, margin + pg.width, oy + pg.height],
                    outline=(210, 214, 220), width=1)

    # Numbered badges on the located cells.
    br = int(dpi * 0.085)
    for k in keys:
        if k['num'] not in located:
            continue
        pi, bb = located[k['num']]
        cx, cy = margin + bb[2], offs[pi] + bb[1]
        color = _SEV_RGB[k['severity']]
        d.ellipse([cx - br, cy - br, cx + br, cy + br], fill=color, outline=_WHITE, width=2)
        d.text((cx, cy), str(k['num']), font=f_badge, fill=_WHITE, anchor='mm')

    # Legend.
    rr = int(fsz * 0.7)
    tx = margin + 2 * rr + 10

    def draw_block(entries, ly):
        for e in entries:
            color = _SEV_RGB[e['severity']]
            if e['badge'] is not None:
                d.ellipse([margin, ly, margin + 2 * rr, ly + 2 * rr], fill=color)
                d.text((margin + rr, ly + rr), e['badge'], font=f_badge,
                       fill=_WHITE, anchor='mm')
            else:
                dot = int(fsz * 0.3)
                d.ellipse([margin + rr - dot, ly + rr - dot,
                           margin + rr + dot, ly + rr + dot], fill=color)
            for j, line in enumerate(e['lines']):
                d.text((tx, ly + j * line_h - 1), line, font=f_leg,
                       fill=_TEXT if j == 0 else _MUTED)
            ly += line_h * len(e['lines']) + entry_gap
        return ly

    ly = (offs[-1] + pages[-1].height + gap + margin) if pages else title_h + margin
    if key_entries:
        d.text((margin, ly), 'Findings', font=f_legb, fill=_TEXT)
        ly = draw_block(key_entries, ly + head_h)
    if extra_entries:
        d.text((margin, ly), 'Also flagged (not tied to a cell)', font=f_legb, fill=_TEXT)
        ly = draw_block(extra_entries, ly + head_h)
    if not key_entries and not extra_entries:
        d.text((margin, ly), 'No cell-level issues to highlight on this sheet.',
               font=f_leg, fill=_MUTED)

    out = io.BytesIO()
    canvas.save(out, format='PNG')
    return out.getvalue()
