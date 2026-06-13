#!/usr/bin/env python3
"""
Lab Report Reviewer - Ansaldo Energia

Rule-based QA review of AEG materials-lab Excel reports. Two report families
are supported:

  * Metallurgical reports  (a "MET"-style sheet): header, sample/material,
    hardness, Nominal-vs-Actual chemical composition, comment, micrographs,
    sign-off.
  * Coating reports        (Cover + assessment sheet): coating-thickness
    measurements checked against the design MIN/MAX limits.

The reviewer is deterministic and runs fully offline. Every finding carries a
severity and a plain-English reason so an engineer can see the basis for it.

Public entry point:
    review_report(filename, data_bytes) -> (report_type, parsed, findings)
        report_type : 'metallurgical' | 'coating' | 'unknown'
        parsed      : dict of the extracted facts (for on-screen display)
        findings    : list of (severity, category, message)
                      severity in {'critical', 'warning', 'info', 'pass'}

Usage (CLI):  python3 lab_review.py report.xlsx
"""
import io
import re
import sys
import zipfile

import openpyxl

# Optional OCR stack, used only to read the burned-in legend on micrographs.
# The reviewer works fully without it — legend reading is skipped gracefully
# when Pillow / pytesseract / the Tesseract binary are not present.
try:
    import pytesseract
    from PIL import Image
    pytesseract.get_tesseract_version()
    _OCR_AVAILABLE = True
except Exception:
    _OCR_AVAILABLE = False

# ── Reference data ────────────────────────────────────────────────────────
# Chemical-element symbols expected in composition tables. Used to tell an
# element-header cell ("Cr", "Ni", ...) apart from an alloy name ("GTD-741").
ELEMENTS = {
    'Ni', 'Cr', 'Co', 'Mo', 'W', 'Al', 'Ti', 'Ta', 'C', 'B', 'Nb', 'V', 'Fe',
    'Zr', 'Cu', 'Mn', 'Si', 'Hf', 'Re', 'Y', 'Pt', 'Pd', 'S', 'P', 'N', 'O',
    'Mg', 'Ce', 'La', 'Sn', 'Ag',
}

# Composition tolerance bands (relative deviation of Actual vs Nominal). An
# absolute floor is applied as well so trace elements (e.g. C, B) don't trip
# the check on tiny absolute differences.
COMP_WARN_REL, COMP_WARN_ABS = 10.0, 0.10     # → warning
COMP_CRIT_REL, COMP_CRIT_ABS = 25.0, 0.20     # → critical

# ── Reference hardness ────────────────────────────────────────────────────
# Typical hardness of common Ni- and Co-based gas-turbine superalloys in the
# fully-heat-treated / aged condition, in HRC. ADVISORY — representative of
# published/typical data; actual values depend on the exact heat-treat cycle
# and the controlling specification, so verify before relying on them.
# Anchors from datasheets / open literature: IN738LC aged ~40-45 HRC;
# Rene 80 aged ~35 HRC (as-cast ~38); GTD-111 ~440→320 HV ≈ 44→32 HRC across
# aging; IN718 aged 36-44 HRC; Nimonic/C263 ~28 HRC.
#
# CONDITION NOTE: AEG reports record *pre-* and *post-solution* hardness. The
# solution-treated state is intentionally SOFTER than these aged ranges (re-
# aging follows), so post-solution readings below the range are EXPECTED and
# are reported as informational, never as failures.
def _hrc(lo, hi, base, note=''):
    return {'hrc': (lo, hi), 'base': base, 'note': note}

# Keys are normalised (uppercase, alphanumerics only) — see _alloy_key().
HARDNESS_REF = {
    # ── Nickel-based, precipitation (γ′) hardened ──
    'IN738':      _hrc(32, 44, 'Ni'),
    'IN738LC':    _hrc(32, 44, 'Ni'),
    'INCONEL738': _hrc(32, 44, 'Ni'),
    'IN792':      _hrc(32, 44, 'Ni'),
    'GTD111':     _hrc(32, 44, 'Ni'),
    'GTD444':     _hrc(32, 44, 'Ni'),
    'GTD741':     _hrc(30, 42, 'Ni', 'GE proprietary — typical Ni bucket range; verify.'),
    'RENE80':     _hrc(30, 40, 'Ni'),
    'RENE108':    _hrc(35, 45, 'Ni'),
    'RENE142':    _hrc(35, 45, 'Ni'),
    'RENEN5':     _hrc(35, 45, 'Ni'),
    'MARM247':    _hrc(38, 46, 'Ni'),
    'CM247LC':    _hrc(38, 46, 'Ni'),
    'IN100':      _hrc(36, 44, 'Ni'),
    'IN713':      _hrc(30, 42, 'Ni'),
    'IN713C':     _hrc(30, 42, 'Ni'),
    'WASPALOY':   _hrc(32, 42, 'Ni'),
    'UDIMET500':  _hrc(30, 40, 'Ni'),
    'UDIMET520':  _hrc(30, 40, 'Ni'),
    'UDIMET720':  _hrc(36, 46, 'Ni'),
    'IN718':      _hrc(36, 44, 'Ni'),
    'INCONEL718': _hrc(36, 44, 'Ni'),
    'NIMONIC263': _hrc(20, 32, 'Ni', 'Age-hardenable Ni-Co-Cr-Mo; aged ~28 HRC.'),
    'C263':       _hrc(20, 32, 'Ni'),
    'NI263':      _hrc(20, 32, 'Ni'),
    'HAYNES263':  _hrc(20, 32, 'Ni'),
    'NIMONIC90':  _hrc(30, 42, 'Ni'),
    'NIMONIC105': _hrc(32, 42, 'Ni'),
    'NIMONIC115': _hrc(32, 42, 'Ni'),
    'HAYNES282':  _hrc(28, 38, 'Ni'),
    # ── Nickel-based, solid-solution (not age-hardened; annealed, much softer) ──
    'IN625':      _hrc(8, 25, 'Ni', 'Solid-solution; annealed ~88-96 HRB.'),
    'INCONEL625': _hrc(8, 25, 'Ni', 'Solid-solution; annealed ~88-96 HRB.'),
    'HASTELLOYX': _hrc(8, 25, 'Ni', 'Solid-solution; annealed ~90 HRB.'),
    # ── Cobalt-based (carbide / solid-solution strengthened) ──
    'FSX414':     _hrc(25, 38, 'Co', 'Cast Co nozzle/vane alloy.'),
    'X40':        _hrc(30, 42, 'Co'),
    'X45':        _hrc(30, 42, 'Co'),
    'STELLITE31': _hrc(30, 42, 'Co'),
    'MARM509':    _hrc(30, 42, 'Co'),
    'ECY768':     _hrc(30, 42, 'Co'),
    'STELLITE6':  _hrc(36, 45, 'Co'),
    'HAYNES188':  _hrc(8, 25, 'Co', 'Solid-solution; annealed ~95 HRB.'),
    'L605':       _hrc(8, 25, 'Co', 'Solid-solution; annealed ~95-100 HRB.'),
    'HAYNES25':   _hrc(8, 25, 'Co', 'Solid-solution; annealed ~95-100 HRB.'),
}


def _alloy_key(material):
    """Normalise an alloy name for HARDNESS_REF lookup (e.g. 'GTD-741'→'GTD741')."""
    return re.sub(r'[^A-Z0-9]', '', (material or '').upper())

# Placeholder strings that mean "field not actually filled in".
_PLACEHOLDERS = {'', 'n/a', 'na', 'not provided', 'to follow', 'tbd', '-', '/'}


# ── Low-level cell helpers ────────────────────────────────────────────────
def _txt(v):
    return '' if v is None else str(v).strip()


def _find(ws, pattern, col=None, max_row=None):
    """Return (row, col) of the first cell whose text matches `pattern`."""
    rx = re.compile(pattern, re.I)
    for row in ws.iter_rows(max_row=max_row):
        for cell in row:
            if col is not None and cell.column != col:
                continue
            t = _txt(cell.value)
            if t and rx.search(t):
                return cell.row, cell.column
    return None


def _value_right(ws, row, col, max_scan=12):
    """First non-empty cell value to the right of (row, col), same row."""
    for c in range(col + 1, col + 1 + max_scan):
        t = _txt(ws.cell(row=row, column=c).value)
        if t:
            return t
    return None


def _value_below(ws, row, col, max_scan=6):
    """First non-empty cell value below (row, col), same column."""
    for r in range(row + 1, row + 1 + max_scan):
        t = _txt(ws.cell(row=r, column=col).value)
        if t:
            return t
    return None


def _is_placeholder(v):
    return _txt(v).lower() in _PLACEHOLDERS


def _num(v):
    """Parse a float from a cell that may carry units/symbols; else None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r'-?\d+(?:\.\d+)?', str(v))
    return float(m.group()) if m else None


# ── Report-type detection ─────────────────────────────────────────────────
def detect_type(wb):
    for ws in wb.worksheets:
        if _find(ws, r'Coating\s*Coverage\s*Assessment', max_row=10):
            return 'coating'
    for ws in wb.worksheets:
        if _find(ws, r'Design\s*limit') and _find(ws, r'Measurements'):
            return 'coating'
    for ws in wb.worksheets:
        if _find(ws, r'METALLURGICAL\s+EXAMINATION', max_row=6) or _find(ws, r'Sample\s*nr'):
            return 'metallurgical'
    return 'unknown'


# ════════════════════════════════════════════════════════════════════════
# METALLURGICAL REPORTS
# ════════════════════════════════════════════════════════════════════════
def _met_sheet(wb):
    for ws in wb.worksheets:
        if _find(ws, r'Sample\s*nr') or _find(ws, r'METALLURGICAL\s+EXAMINATION', max_row=6):
            return ws
    return wb.worksheets[0]


def _header(ws):
    out = {}
    labels = {
        'customer':     r'^Customer\s*:',
        'customer_ref': r'Customer\s*Ref',
        'aeg_ref':      r'AEG.*Ref',
        'job':          r'AEG.*Job',
        'machine':      r'Machine\s*Type',
        'qty':          r'Quantity',
        'eoh':          r'\bEOH\b',
    }
    for key, pat in labels.items():
        loc = _find(ws, pat)
        if loc:
            out[key] = _value_right(ws, *loc)
    return out


def _sample(ws):
    loc = _find(ws, r'Sample\s*nr')
    out = {}
    if not loc:
        return out
    hrow = loc[0]
    headers = {}
    for cell in ws[hrow]:
        t = _txt(cell.value).lower()
        if t:
            headers[t] = cell.column

    def below(substr):
        for h, c in headers.items():
            if substr in h:
                return _value_below(ws, hrow, c)
        return None

    out['description'] = below('description')
    out['serial']      = below('s/n')
    out['location']    = below('location')
    out['material']    = below('material')
    out['result']      = below('result')
    return out


def _hardness(ws):
    out = {}
    for key, pat in (('pre', r'Pre-?\s*Solution'), ('post', r'Post-?\s*Solution')):
        loc = _find(ws, pat)
        if loc:
            raw = _value_right(ws, *loc)
            out[key] = {'raw': raw, 'value': _num(raw)}
    return out


def _composition(ws, which):
    """Extract {element: value} for which='Nominal' or 'Actual'."""
    loc = _find(ws, r'\(\s*' + which + r'\s*\)')
    comp = {}
    if not loc:
        return comp
    hrow = loc[0]
    for cell in ws[hrow]:
        sym = _txt(cell.value)
        if sym.capitalize() in ELEMENTS:
            val = _num(ws.cell(row=hrow + 1, column=cell.column).value)
            if val is not None:
                comp[sym.capitalize()] = val
    return comp


def _comment(ws):
    loc = _find(ws, r'^Comment\s*:')
    return _value_below(ws, *loc) if loc else None


def _pictures(ws):
    rx = re.compile(r'Picture\s*\d+\s*:', re.I)
    pics = []
    for row in ws.iter_rows():
        for cell in row:
            if rx.search(_txt(cell.value)):
                pics.append((_txt(cell.value), _value_right(ws, cell.row, cell.column)))
    return pics


def _signoff(ws):
    out = {}
    for key, pat in (('met_lab', r'Met\.?\s*Lab'),
                     ('mat_eng', r'Mat\.?\s*Eng'),
                     ('date',    r'^Date\s*:')):
        loc = _find(ws, pat)
        if loc:
            out[key] = _value_right(ws, *loc)
    return out


def parse_metallurgical(wb, media=0):
    ws = _met_sheet(wb)
    return {
        'header':    _header(ws),
        'sample':    _sample(ws),
        'hardness':  _hardness(ws),
        'nominal':   _composition(ws, 'Nominal'),
        'actual':    _composition(ws, 'Actual'),
        'comment':   _comment(ws),
        'pictures':  _pictures(ws),
        'signoff':   _signoff(ws),
        'media':     media,
    }


def _review_composition(nominal, actual):
    findings = []
    if not nominal or not actual:
        findings.append(('warning', 'Composition',
                         'Could not read both Nominal and Actual composition tables.'))
        return findings

    common = sorted(set(nominal) & set(actual))
    flagged = False
    for el in common:
        nom, act = nominal[el], actual[el]
        if nom == 0:
            continue
        dev = act - nom
        rel = dev / abs(nom) * 100.0
        a, r = abs(dev), abs(rel)
        if r >= COMP_CRIT_REL and a >= COMP_CRIT_ABS:
            sev = 'critical'
        elif r >= COMP_WARN_REL and a >= COMP_WARN_ABS:
            sev = 'warning'
        else:
            continue
        flagged = True
        findings.append((sev, 'Composition',
                         f'{el}: actual {act:g} vs nominal {nom:g} wt% ({rel:+.0f}%).'))

    only_nom = sorted(set(nominal) - set(actual))
    only_act = sorted(set(actual) - set(nominal))
    if only_nom:
        findings.append(('info', 'Composition',
                         f'In spec but not reported in actual: {", ".join(only_nom)}.'))
    if only_act:
        findings.append(('info', 'Composition',
                         f'Reported but not in nominal spec: {", ".join(only_act)}.'))
    if not flagged:
        findings.append(('pass', 'Composition',
                         f'All {len(common)} matched elements within ±{COMP_WARN_REL:g}% tolerance.'))
    return findings


def _review_hardness(hardness, material):
    findings = []
    if not hardness:
        findings.append(('info', 'Hardness', 'No hardness-results section found.'))
        return findings

    pre = hardness.get('pre', {}).get('value')
    post = hardness.get('post', {}).get('value')
    if pre is None and post is None:
        findings.append(('warning', 'Hardness', 'Hardness section present but no values parsed.'))
        return findings

    # Real check: solution treatment should soften the material (post ≤ pre).
    if pre is not None and post is not None and post > pre + 0.5:
        findings.append(('warning', 'Hardness',
                         f'Post-solution hardness ({post:g} HRC) exceeds pre-solution '
                         f'({pre:g} HRC) — solution treatment normally softens the material.'))

    ref = HARDNESS_REF.get(_alloy_key(material))
    if ref:
        lo, hi = ref['hrc']
        note = (' ' + ref['note']) if ref['note'] else ''
        findings.append(('info', 'Hardness',
                         f'{material}: reference aged hardness {lo}–{hi} HRC '
                         f'({ref["base"]}-based, advisory).{note}'))
        for label, val in (('Pre-solution', pre), ('Post-solution', post)):
            if val is None:
                continue
            if val > hi + 2:
                findings.append(('info', 'Hardness',
                                 f'{label} {val:g} HRC is above the aged reference '
                                 f'{lo}–{hi} HRC — verify.'))
            elif val < lo and label == 'Post-solution':
                findings.append(('info', 'Hardness',
                                 f'{label} {val:g} HRC is below the aged reference '
                                 f'{lo}–{hi} HRC — expected for the solution-treated '
                                 f'(pre-aging) condition.'))
    else:
        findings.append(('info', 'Hardness',
                         f'No reference hardness on file for "{material}".'))

    if not any(s == 'warning' for s, _, _ in findings):
        parts = [f'{k}={v["value"]:g}' for k, v in hardness.items() if v.get('value') is not None]
        findings.append(('pass', 'Hardness', f'Hardness values recorded: {", ".join(parts)} HRC.'))
    return findings


def _review_completeness(parsed):
    findings = []
    hdr = parsed['header']
    for key, label in (('customer', 'Customer'), ('job', 'AEG Job No'),
                       ('machine', 'Machine type')):
        if _is_placeholder(hdr.get(key)):
            findings.append(('warning', 'Completeness', f'{label} is blank or a placeholder.'))
    for key, label in (('customer_ref', 'Customer Ref No'), ('eoh', 'EOH')):
        if _is_placeholder(hdr.get(key)):
            findings.append(('info', 'Completeness', f'{label} not provided.'))

    if _is_placeholder(parsed['sample'].get('material')):
        findings.append(('warning', 'Completeness', 'Sample material/alloy not stated.'))

    comment = parsed.get('comment') or ''
    if len(comment.strip()) < 40:
        findings.append(('warning', 'Completeness',
                         'Comment / discussion is missing or very short.'))

    pics = parsed.get('pictures', [])
    uncaptioned = [p for p, cap in pics if not cap]
    if not pics:
        findings.append(('warning', 'Micrographs', 'No micrograph captions found.'))
    elif uncaptioned:
        findings.append(('info', 'Micrographs',
                         f'{len(uncaptioned)} of {len(pics)} pictures have no caption.'))
    if parsed.get('media', 0) == 0:
        findings.append(('warning', 'Micrographs', 'No embedded images found in the workbook.'))

    so = parsed['signoff']
    missing = [lbl for key, lbl in (('met_lab', 'Met. Lab'), ('mat_eng', 'Mat. Eng'),
                                    ('date', 'Date')) if _is_placeholder(so.get(key))]
    if missing:
        findings.append(('warning', 'Sign-off', f'Missing sign-off field(s): {", ".join(missing)}.'))
    else:
        findings.append(('pass', 'Sign-off', 'Lab, engineer and date all present.'))
    return findings


def review_metallurgical(parsed):
    findings = []
    findings += _review_completeness(parsed)
    findings += _review_hardness(parsed['hardness'], parsed['sample'].get('material'))
    findings += _review_composition(parsed['nominal'], parsed['actual'])
    return findings


# ════════════════════════════════════════════════════════════════════════
# COATING REPORTS
# ════════════════════════════════════════════════════════════════════════
def _coating_signoff(wb):
    out = {}
    for ws in wb.worksheets:
        for key, pat in (('prepared', r'Prepared\s*by'),
                         ('approved', r'Approved\s*by'),
                         ('date',     r'^Date\s*:')):
            if key in out:
                continue
            loc = _find(ws, pat)
            if loc:
                out[key] = _value_right(ws, *loc)
    return out


def parse_coating(wb, media=0):
    # The assessment sheet is the one carrying the actual MIN/MAX design
    # limits — not the Cover sheet, whose table-of-contents also mentions
    # "Coating Coverage Assessment".
    aws = None
    for ws in wb.worksheets:
        if _find(ws, r'Design\s*limit') and _find(ws, r'Measurements'):
            aws = ws
            break

    data = {'title': None, 'report_no': None, 'rows': [],
            'signoff': _coating_signoff(wb), 'media': media}

    cover = wb.worksheets[0]
    t = _find(cover, r'Coating')
    if t:
        data['title'] = _txt(cover.cell(row=t[0], column=t[1]).value)
    rn = _find(cover, r'Report\s*No')
    if rn:
        data['report_no'] = _value_right(cover, *rn)

    if aws is None:
        return data

    meas_loc = _find(aws, r'Measurements')
    avg_loc  = _find(aws, r'Average\s*Values')
    min_loc  = _find(aws, r'^MIN$')
    max_loc  = _find(aws, r'^MAX$')
    if not (meas_loc and avg_loc and min_loc and max_loc):
        return data

    hrow = meas_loc[0]
    meas_cols = list(range(meas_loc[1], avg_loc[1]))     # measurement value columns
    min_col, max_col = min_loc[1], max_loc[1]

    cur_min = cur_max = None
    for r in range(hrow + 1, aws.max_row + 1):
        m = _num(aws.cell(row=r, column=min_col).value)
        x = _num(aws.cell(row=r, column=max_col).value)
        if m is not None:
            cur_min = m
        if x is not None:
            cur_max = x
        vals = [_num(aws.cell(row=r, column=c).value) for c in meas_cols]
        vals = [v for v in vals if v is not None]
        if not vals:
            continue
        data['rows'].append({'row': r, 'values': vals,
                             'min': cur_min, 'max': cur_max})
    return data


def review_coating(parsed):
    findings = []
    rows = parsed.get('rows', [])
    if not rows:
        findings.append(('warning', 'Coating', 'Could not read the coating-coverage assessment table.'))
        return findings

    out_of_range = 0
    total = 0
    limits_seen = False
    for entry in rows:
        lo, hi = entry['min'], entry['max']
        if lo is None or hi is None:
            continue
        limits_seen = True
        for v in entry['values']:
            total += 1
            if not (lo <= v <= hi):
                out_of_range += 1
                findings.append(('critical', 'Coating',
                                 f'Row {entry["row"]}: thickness {v:g} mm outside '
                                 f'design limit {lo:g}–{hi:g} mm.'))

    if not limits_seen:
        findings.append(('warning', 'Coating', 'No design MIN/MAX limits found to check against.'))
    elif out_of_range == 0:
        findings.append(('pass', 'Coating',
                         f'All {total} thickness measurements within design limits.'))

    so = parsed['signoff']
    missing = [lbl for key, lbl in (('prepared', 'Prepared by'), ('approved', 'Approved by'),
                                    ('date', 'Date')) if _is_placeholder(so.get(key))]
    if missing:
        findings.append(('warning', 'Sign-off', f'Missing sign-off field(s): {", ".join(missing)}.'))
    else:
        findings.append(('pass', 'Sign-off', 'Prepared-by, approved-by and date all present.'))

    if parsed.get('media', 0) == 0:
        findings.append(('warning', 'Micrographs', 'No embedded reference micrographs found.'))
    return findings


# ════════════════════════════════════════════════════════════════════════
# MICROGRAPH LEGEND OCR  (light "read the legend in the photo" support)
# ════════════════════════════════════════════════════════════════════════
# Burned-in legends follow the AEG convention "<job>_E_<mag>x-<n>" at the
# bottom-left and a scale bar ("10 µm") at the bottom-right. OCR of such small,
# speckle-surrounded text is best-effort: values are correct when read, but not
# every image yields one. Findings are therefore advisory.
_MAG_PATS = [
    re.compile(r'(\d{2,4})\s*[xX%]\s*[-_]\s*(\d)'),   # 500x-1  (magnification + index)
    re.compile(r'E\s*[_ €F]?\s*(\d{2,4})\s*[xX%]'),   # E_500x
    re.compile(r'(?<![\d.])(\d{2,4})\s*[xX%]'),       # 500x
]
_JOB_PAT   = re.compile(r'\b(\d{4})\b')
_SCALE_PAT = re.compile(r'(\d{1,3})\s*[µuμyptwb]+m', re.I)
_CAP_MAG   = re.compile(r'(\d{2,4})\s*[xX]\b')


def _safe_ocr(im, cfg='--psm 7'):
    try:
        return pytesseract.image_to_string(im, config=cfg) or ''
    except Exception:
        return ''


def _binarize(im, thr, scale=4):
    """Keep bright text (white-on-dark legend bar) and upscale small fonts."""
    return im.point(lambda p: 255 if p > thr else 0).resize(
        (max(1, im.width * scale), max(1, im.height * scale)))


def _read_one_legend(img_bytes):
    """OCR the burned-in legend of a single micrograph; None if unreadable."""
    try:
        im = Image.open(io.BytesIO(img_bytes)).convert('L')
    except Exception:
        return None
    w, h = im.size
    if w < 200 or h < 150:            # skip logos / thumbnails
        return None
    lc = im.crop((0, int(h * 0.90), int(w * 0.55), h))           # ID + magnification
    rc = im.crop((int(w * 0.72), int(h * 0.88), w, h))           # scale bar
    lblob = ' '.join(_safe_ocr(_binarize(lc, t)) for t in (110, 130, 150))
    rblob = ' '.join(_safe_ocr(_binarize(rc, t)) for t in (110, 140))

    out = {}
    mag_val, idx = None, None
    for pat in _MAG_PATS:                      # first plausible magnification wins
        for m in pat.finditer(lblob):
            n = int(m.group(1))
            if 25 <= n <= 20000:               # real micrograph mags; rejects OCR noise
                mag_val = n
                idx = m.group(2) if pat.groups == 2 else None
                break
        if mag_val is not None:
            break
    if mag_val is not None:
        job = _JOB_PAT.search(lblob)
        out['mag'] = f'{mag_val}x'
        out['id'] = (f'{job.group(1)}_' if job else '') + f'E_{mag_val}x' + \
                    (f'-{idx}' if idx else '')
    s = _SCALE_PAT.search(rblob) or _SCALE_PAT.search(lblob)
    if s:
        out['scale'] = f'{s.group(1)} µm'
    return out or None


def read_image_legends(data, max_images=24):
    """Extract embedded micrographs from an xlsx and OCR each legend.

    Returns (legends, ocr_used):
        legends  : list of {'image', 'mag', 'scale', 'id'} (only readable ones)
        ocr_used : False when the OCR stack is unavailable
    """
    if not _OCR_AVAILABLE:
        return [], False
    legends = []
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        names = [n for n in z.namelist() if n.startswith('xl/media')]
    except Exception:
        return [], True
    for n in sorted(names)[:max_images]:
        try:
            info = _read_one_legend(z.read(n))
        except Exception:
            info = None
        if info:
            info['image'] = n.split('/')[-1]
            legends.append(info)
    return legends, True


def _caption_mags(pictures):
    """Magnifications mentioned in the written picture captions, e.g. {'200x'}."""
    mags = set()
    for _, cap in pictures or []:
        for m in _CAP_MAG.finditer(cap or ''):
            mags.add(f'{m.group(1)}x')
    return mags


def _review_legends(legends, ocr_used, caption_mags):
    findings = []
    if not ocr_used:
        findings.append(('info', 'Photo legends',
                         'Legend OCR unavailable (Tesseract not installed) — skipped.'))
        return findings
    if not legends:
        findings.append(('info', 'Photo legends',
                         'Could not read a legend from any embedded micrograph.'))
        return findings

    img_mags = sorted({l['mag'] for l in legends if l.get('mag')},
                      key=lambda s: int(s[:-1]))
    findings.append(('info', 'Photo legends',
                     f'Read legends from {len(legends)} micrograph(s); '
                     f'magnifications: {", ".join(img_mags) if img_mags else "n/a"}.'))

    # Cross-check magnifications burned into the images against the captions.
    if img_mags and caption_mags:
        missing = [m for m in img_mags if m not in caption_mags]
        if missing:
            findings.append(('warning', 'Photo legends',
                             f'Magnification(s) {", ".join(missing)} appear in image legends '
                             f'but in no written caption — check the captions.'))
        else:
            findings.append(('pass', 'Photo legends',
                             'Image-legend magnifications all match the written captions.'))
    return findings


# ════════════════════════════════════════════════════════════════════════
# PUBLIC ENTRY POINT
# ════════════════════════════════════════════════════════════════════════
def _media_count(data):
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        return sum(1 for n in z.namelist() if n.startswith('xl/media'))
    except Exception:
        return 0


def review_report(filename, data, ocr=True):
    """Review one report. Returns (report_type, parsed, findings).

    ocr : when True (and the OCR stack is available) the burned-in legend of
          each embedded micrograph is read and cross-checked against captions.
    """
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    rtype = detect_type(wb)
    media = _media_count(data)

    if rtype == 'coating':
        parsed = parse_coating(wb, media)
        findings = review_coating(parsed)
    elif rtype == 'metallurgical':
        parsed = parse_metallurgical(wb, media)
        findings = review_metallurgical(parsed)
    else:
        parsed = {}
        findings = [('warning', 'Format',
                     'Unrecognised layout — not classified as a metallurgical or coating report.')]

    legends = []
    if ocr:
        legends, ocr_used = read_image_legends(data)
        cap_mags = _caption_mags(parsed.get('pictures', []))
        findings += _review_legends(legends, ocr_used, cap_mags)
    parsed['legends'] = legends
    return rtype, parsed, findings


def summarize(findings):
    """Return counts per severity."""
    out = {'critical': 0, 'warning': 0, 'info': 0, 'pass': 0}
    for sev, _, _ in findings:
        out[sev] = out.get(sev, 0) + 1
    return out


# ── CLI ───────────────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for path in sys.argv[1:]:
        with open(path, 'rb') as f:
            data = f.read()
        rtype, parsed, findings = review_report(path, data)
        counts = summarize(findings)
        print('=' * 78)
        print(f'{path}')
        print(f'  type: {rtype}   '
              f'critical={counts["critical"]} warning={counts["warning"]} '
              f'info={counts["info"]} pass={counts["pass"]}')
        for sev, cat, msg in findings:
            tag = {'critical': 'FAIL', 'warning': 'WARN', 'info': 'INFO', 'pass': 'OK  '}[sev]
            print(f'   [{tag}] {cat}: {msg}')
        for lg in parsed.get('legends', []):
            bits = [lg[k] for k in ('id', 'mag', 'scale') if lg.get(k)]
            print(f'     · {lg["image"]}: {"  ".join(bits)}')


if __name__ == '__main__':
    main()
