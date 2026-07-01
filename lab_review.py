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
import os
import re
import sys
import zipfile

import openpyxl

# Optional OCR / imaging stack. The reviewer works without it — legend, etch
# and thickness reading from micrographs are skipped gracefully when Pillow /
# pytesseract / the Tesseract binary are not present.
try:
    from PIL import Image, ImageFilter
    _PIL_AVAILABLE = True
except Exception:
    _PIL_AVAILABLE = False

try:
    import pytesseract
    pytesseract.get_tesseract_version()
    _OCR_AVAILABLE = _PIL_AVAILABLE
except Exception:
    _OCR_AVAILABLE = False

# Caption / etch / heat-treatment vocabulary lives in lab_vocab; import (and so
# re-export) it so external callers can keep doing `from lab_review import …`.
from lab_vocab import (  # noqa: E402,F401
    _PICNUM, _ETCH_PAT, _UNETCHED_PAT, _ALLOY_PAT, _norm_alloy,
    _ETCHANT_VOCAB, caption_etchant, report_etchants, image_etchant,
    _HT_VOCAB, HT_ORDER, caption_ht, report_ht, image_ht,
)

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
    'GTD222':     _hrc(25, 38, 'Ni', 'GE cast nozzle alloy (moderate γ′, weldable) — advisory; verify.'),
    'GTD241':     _hrc(25, 40, 'Ni', 'GE cast nozzle alloy — advisory; verify.'),
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
    'U500':       _hrc(30, 40, 'Ni'),     # Udimet 500 alias
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
    'HASTX':      _hrc(8, 25, 'Ni', 'Hastelloy X alias; solid-solution, ~90 HRB.'),
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


def _value_right_loc(ws, row, col, max_scan=12):
    """(value, (row, col)) of the first non-empty cell to the right, else (None, None)."""
    for c in range(col + 1, col + 1 + max_scan):
        t = _txt(ws.cell(row=row, column=c).value)
        if t:
            return t, (row, c)
    return None, None


def _value_below_loc(ws, row, col, max_scan=6):
    """(value, (row, col)) of the first non-empty cell below, else (None, None)."""
    for r in range(row + 1, row + 1 + max_scan):
        t = _txt(ws.cell(row=r, column=col).value)
        if t:
            return t, (r, col)
    return None, None


def _value_right(ws, row, col, max_scan=12):
    """First non-empty cell value to the right of (row, col), same row."""
    return _value_right_loc(ws, row, col, max_scan)[0]


def _value_below(ws, row, col, max_scan=6):
    """First non-empty cell value below (row, col), same column."""
    return _value_below_loc(ws, row, col, max_scan)[0]


def _is_placeholder(v):
    return _txt(v).lower() in _PLACEHOLDERS


def _num(v):
    """Parse a float from a cell that may carry units/symbols. Handles both the
    English '12.5' / '1,234.56' and the European '12,5' / '1.234,56' forms
    (common on Italian-lab reports); returns None when there's no number."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r'-?\d[\d.,  ]*\d|-?\d', str(v))
    if not m:
        return None
    s = re.sub(r'[ \s]', '', m.group())     # drop spaces / non-breaking spaces
    # The decimal separator is whichever of '.' / ',' appears LAST; the other is
    # a thousands grouper. (No separator ⇒ plain integer.)
    if s.rfind(',') > s.rfind('.'):
        s = s.replace('.', '').replace(',', '.')  # European: comma is the decimal
    else:
        s = s.replace(',', '')                    # English: comma groups thousands
    try:
        return float(s)
    except ValueError:
        return None


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
    out, loc = {}, {}
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
        lbl = _find(ws, pat)
        if lbl:
            out[key], vloc = _value_right_loc(ws, *lbl)
            loc[key] = {'label': lbl, 'value': vloc}
    return out, loc


def _sample(ws):
    hdr = _find(ws, r'Sample\s*nr')
    out, loc = {}, {}
    if not hdr:
        return out, loc
    hrow = hdr[0]
    headers = {}
    for cell in ws[hrow]:
        t = _txt(cell.value).lower()
        if t:
            headers[t] = cell.column

    def below(substr):
        for h, c in headers.items():
            if substr in h:
                val, vloc = _value_below_loc(ws, hrow, c)
                return val, {'label': (hrow, c), 'value': vloc}
        return None, None

    for key, substr in (('description', 'description'), ('serial', 's/n'),
                        ('location', 'location'), ('material', 'material'),
                        ('result', 'result')):
        out[key], lc = below(substr)
        if lc:
            loc[key] = lc
    return out, loc


def _hardness_unit(raw):
    """Hardness scale named in a cell (HRC / HV / HBW / HRB / …), or None."""
    m = re.search(r'\b(HRC|HRB|HRA|HV|HBW|HB|HK)\b', str(raw or ''), re.I)
    return m.group(1).upper() if m else None


def _hardness(ws):
    out, loc = {}, {}
    for key, pat in (('pre', r'Pre-?\s*Solution'), ('post', r'Post-?\s*Solution')):
        lbl = _find(ws, pat)
        if lbl:
            raw, vloc = _value_right_loc(ws, *lbl)
            out[key] = {'raw': raw, 'value': _num(raw), 'unit': _hardness_unit(raw)}
            loc[key] = {'label': lbl, 'value': vloc}
    return out, loc


def _coating(ws):
    """Coating presence / type as recorded in the structured cells."""
    out = {'present': None, 'type': None, 'received': None, 'outgoing': None}
    loc = {}
    for key, pat in (('present', r'^Coating\s*$'),
                     ('type', r'^Type of Coating'),
                     ('received', r'Received\s*Coating'),
                     ('outgoing', r'Outgoing\s*Coating')):
        lbl = _find(ws, pat)
        if lbl:
            out[key], vloc = _value_below_loc(ws, *lbl)
            loc[key] = {'label': lbl, 'value': vloc}
    return out, loc


def _composition(ws, which):
    """Extract {element: value} for which='Nominal' or 'Actual'.

    Handles the two layouts seen in practice: element headers in the
    '(Nominal/Actual)' label row with values below, OR element headers in the
    row above the label with values in the label row. Tolerates the 'Minimal'
    typo for 'Nominal'.
    """
    pat = r'\(\s*(?:Nominal|Minimal)\s*\)' if which == 'Nominal' else r'\(\s*Actual\s*\)'
    lbl = _find(ws, pat)
    comp, loc = {}, {}
    if not lbl:
        return comp, loc
    lrow = lbl[0]

    def elem_cells(r):
        return [c for c in ws[r] if r >= 1 and _txt(c.value).capitalize() in ELEMENTS]

    if len(elem_cells(lrow)) >= 2:           # elements in label row, values below
        ehdr, vrow = lrow, lrow + 1
    elif lrow > 1 and len(elem_cells(lrow - 1)) >= 2:   # elements above, values in label row
        ehdr, vrow = lrow - 1, lrow
    else:
        return comp, loc
    for cell in elem_cells(ehdr):
        val = _num(ws.cell(row=vrow, column=cell.column).value)
        if val is not None:
            el = _txt(cell.value).capitalize()
            comp[el] = val
            loc[el] = (vrow, cell.column)
    return comp, loc


def _comment(ws):
    lbl = _find(ws, r'^Comment\s*:')
    if not lbl:
        return None, {}
    val, vloc = _value_below_loc(ws, *lbl)
    return val, {'label': lbl, 'value': vloc}


def _pictures(ws):
    rx = re.compile(r'Picture\s*\d+\s*:', re.I)
    pics, loc = [], []
    for row in ws.iter_rows():
        for cell in row:
            if rx.search(_txt(cell.value)):
                cap, vloc = _value_right_loc(ws, cell.row, cell.column)
                pics.append((_txt(cell.value), cap))
                loc.append({'label': (cell.row, cell.column), 'value': vloc})
    return pics, loc


def _signoff(ws):
    out, loc = {}, {}
    for key, pat in (('met_lab', r'Met\.?\s*Lab'),
                     ('mat_eng', r'(?:Mat|Met)\.?\s*Eng'),
                     ('date',    r'^Date\s*:')):
        lbl = _find(ws, pat)
        if lbl:
            out[key], vloc = _value_right_loc(ws, *lbl)
            loc[key] = {'label': lbl, 'value': vloc}
    return out, loc


def parse_metallurgical(wb, media=0):
    ws = _met_sheet(wb)
    header, lh    = _header(ws)
    sample, ls    = _sample(ws)
    hardness, lhd = _hardness(ws)
    nominal, ln   = _composition(ws, 'Nominal')
    actual, la    = _composition(ws, 'Actual')
    coating, lc   = _coating(ws)
    comment, lcm  = _comment(ws)
    pictures, lp  = _pictures(ws)
    signoff, lso  = _signoff(ws)
    return {
        'header':    header,
        'sample':    sample,
        'hardness':  hardness,
        'nominal':   nominal,
        'actual':    actual,
        'coating':   coating,
        'comment':   comment,
        'pictures':  pictures,
        'signoff':   signoff,
        'media':     media,
        'loc': {
            'sheet':    ws.title,
            'header':   lh,
            'sample':   ls,
            'hardness': lhd,
            'nominal':  ln,
            'actual':   la,
            'coating':  lc,
            'comment':  lcm,
            'pictures': lp,
            'signoff':  lso,
        },
    }


def _composition_deviations(nominal, actual):
    """Elements whose Actual is out of tolerance vs Nominal.

    Returns (deviations, systemic) where deviations is a list of
    (element, nominal, actual, rel%) and `systemic` is True when so many
    elements are off that the actual material likely isn't the stated alloy.
    """
    deviations = []
    if not nominal or not actual:
        return deviations, False
    common = sorted(set(nominal) & set(actual))
    for el in common:
        nom, act = nominal[el], actual[el]
        if nom == 0:
            continue
        dev = act - nom
        rel = dev / abs(nom) * 100.0
        if abs(rel) >= COMP_WARN_REL and abs(dev) >= COMP_WARN_ABS:
            deviations.append((el, nom, act, rel))
    n_dev, n_common = len(deviations), len(common)
    systemic = n_dev >= 4 or (n_dev >= 3 and n_common and n_dev / n_common >= 0.5)
    return deviations, systemic


def _review_composition(nominal, actual):
    findings = []
    if not nominal or not actual:
        findings.append(('warning', 'Composition',
                         'Could not read both Nominal and Actual composition tables.'))
        return findings

    common = sorted(set(nominal) & set(actual))
    deviations, systemic = _composition_deviations(nominal, actual)
    n_dev, n_common = len(deviations), len(common)

    # A few elements off is normal service depletion / EDS scatter → warnings.
    # Many elements off together signals the actual material doesn't match the
    # stated alloy → one consolidated FAIL ("verify material/grade").
    if systemic:
        worst = sorted(deviations, key=lambda d: -abs(d[3]))[:4]
        detail = ", ".join(f"{el} {rel:+.0f}%" for el, _, _, rel in worst)
        findings.append(('critical', 'Composition',
                         f'{n_dev} of {n_common} elements out of tolerance ({detail} …) — actual '
                         f'composition does not match the stated alloy; verify material/grade.'))
    else:
        for el, nom, act, rel in deviations:
            findings.append(('warning', 'Composition',
                             f'{el}: actual {act:g} vs nominal {nom:g} wt% ({rel:+.0f}%).'))

    only_nom = sorted(set(nominal) - set(actual))
    only_act = sorted(set(actual) - set(nominal))
    if only_nom:
        findings.append(('info', 'Composition',
                         f'In spec but not reported in actual: {", ".join(only_nom)}.'))
    if only_act:
        findings.append(('info', 'Composition',
                         f'Reported but not in nominal spec: {", ".join(only_act)}.'))
    if not deviations:
        findings.append(('pass', 'Composition',
                         f'All {len(common)} matched elements within ±{COMP_WARN_REL:g}% tolerance.'))
    return findings


# Coating-type vocabulary (tolerant of the "MCrAlY"/"MCrAIY" spelling seen in
# the sheets). Each entry maps a canonical name to a detection pattern.
_COATING_TYPE_PATS = (
    ('TBC',       r'\bTBC\b|thermal\s*barrier'),
    ('MCrAlY',    r'MCR\w*Y'),
    ('aluminide', r'alumini[sz]|\baluminide\b|\bPt[-\s]?Al\b|platinum\s*alumin'),
    ('diffusion', r'diffusion\s*coat'),
    ('chromide',  r'chromi[sz]|\bchromide\b'),
)


def _coating_types_in(text):
    """Set of canonical coating types mentioned in a piece of text."""
    t = text or ''
    return {name for name, pat in _COATING_TYPE_PATS if re.search(pat, t, re.I)}


def _review_comment(parsed):
    """Flag where the free-text comment contradicts the coating cells."""
    findings = []
    comment = parsed.get('comment') or ''
    coat = parsed.get('coating') or {}
    if not comment:
        return findings
    cl = comment.lower()

    cell_types = set()
    for key in ('type', 'received', 'outgoing'):
        cell_types |= _coating_types_in(coat.get(key))
    comment_types = _coating_types_in(comment)

    present = (coat.get('present') or '').strip().lower()
    cell_has  = present == 'yes' or bool(cell_types)
    cell_none = present == 'no' or (not cell_types and _is_placeholder(coat.get('type')))

    comment_has = bool(comment_types) or bool(re.search(
        r'received with[^.]{0,30}coating|coated with|coating (?:was |is )?(?:applied|present|intact)', cl))
    comment_none = bool(re.search(
        r'\buncoated\b|no coating|without (?:any )?coating|not coated|'
        r'coating (?:is |was )?(?:fully )?removed', cl))

    # Coating type: comment names a type the cell disagrees with.
    if cell_types and comment_types and cell_types.isdisjoint(comment_types):
        findings.append(('warning', 'Comment',
                         f'Comment mentions {"/".join(sorted(comment_types))} coating but the '
                         f'coating cell says {"/".join(sorted(cell_types))}.'))
    elif cell_types and (cell_types & comment_types):
        findings.append(('pass', 'Comment',
                         f'Comment coating type matches the coating cell '
                         f'({"/".join(sorted(cell_types & comment_types))}).'))

    # Coating presence: cell vs comment.
    if cell_none and comment_has and not comment_none:
        what = "/".join(sorted(comment_types)) if comment_types else 'a coating'
        findings.append(('warning', 'Comment',
                         f'Coating cell indicates no coating, but the comment refers to {what}.'))
    elif cell_has and comment_none and not comment_has:
        label = "/".join(sorted(cell_types)) or present
        findings.append(('warning', 'Comment',
                         f'Coating cell indicates a coating ({label}), but the comment says '
                         f'it is uncoated.'))

    # Alloy named in the comment vs the material cell.
    material = (parsed.get('sample') or {}).get('material')
    if material:
        mkey = _norm_alloy(material)
        others = sorted({m.group(0) for m in _ALLOY_PAT.finditer(comment)
                         if _norm_alloy(m.group(0)) != mkey
                         and _norm_alloy(m.group(0)) not in mkey
                         and mkey not in _norm_alloy(m.group(0))})
        if others:
            findings.append(('warning', 'Comment',
                             f'Comment mentions alloy {", ".join(others)} but the material cell '
                             f'says "{material}".'))

    # Service verdict in the comment vs the Result cell.
    result = (parsed.get('sample') or {}).get('result') or ''
    rlow = result.lower()
    neg = re.search(r'not\s+suitable|unsuitable|not\s+recommend|\breject|\bscrap|'
                    r'beyond\s+repair|non[-\s]?conform|unacceptable', cl)
    pos = re.search(r'(?<!not )(?:\bsuitable for|\bacceptable|recommended for|'
                    r'reconditi|fit for service|return to service)', cl)
    result_pos = bool(re.search(r'accept|suitable|conform|\bpass\b', rlow)) and 'see comment' not in rlow
    result_neg = bool(re.search(r'reject|not\s+suitable|scrap|unacceptable', rlow))
    if result_pos and neg and not pos:
        findings.append(('warning', 'Comment',
                         f'Result cell says "{result}" but the comment indicates the part is NOT suitable.'))
    elif result_neg and pos and not neg:
        findings.append(('warning', 'Comment',
                         f'Result cell says "{result}" but the comment indicates the part IS suitable.'))
    elif 'see comment' in rlow and bool(neg) != bool(pos):
        findings.append(('info', 'Comment',
                         f'Result defers to the comment; the comment verdict reads '
                         f'{"not suitable / negative" if neg else "suitable / positive"}.'))
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

    # Determine the scale: explicit unit from a cell, else inferred (HRC tops out
    # near 70; a bigger number is HV/HBW). This stops HV microhardness (e.g. 420)
    # being compared to — and flagged against — the HRC reference.
    vals = [v for v in (pre, post) if v is not None]
    units = {hardness.get(k, {}).get('unit') for k in ('pre', 'post')} - {None}
    unit = next(iter(units)) if len(units) == 1 else None
    is_hrc = (unit == 'HRC') or (unit is None and all(v <= 72 for v in vals))
    ustr = unit or ('HRC' if is_hrc else 'HV')

    # Solution treatment should soften the material (post ≤ pre) — scale-agnostic.
    # ±2 guard band so measurement scatter doesn't trip a false warning.
    if pre is not None and post is not None and post > pre + 2:
        findings.append(('warning', 'Hardness',
                         f'Post-solution hardness ({post:g} {ustr}) exceeds pre-solution '
                         f'({pre:g} {ustr}) — solution treatment normally softens the material.'))

    ref = HARDNESS_REF.get(_alloy_key(material))
    if not is_hrc:
        parts = [f'{k}={v["value"]:g}' for k, v in hardness.items() if v.get('value') is not None]
        findings.append(('info', 'Hardness',
                         f'Hardness recorded in {ustr}: {", ".join(parts)} — not compared '
                         f'to the HRC reference.'))
    elif ref:
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
        findings.append(('pass', 'Hardness', f'Hardness values recorded: {", ".join(parts)} {ustr}.'))
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


# Caption / etch / heat-treatment vocabulary now lives in lab_vocab.py and is
# imported at the top of this module (and re-exported for external callers).


def _anchor_order(data):
    """Embedded micrographs in visual (drawing-anchor) order, top-to-bottom /
    left-to-right, excluding logos/thumbnails. Returns a list of image names."""
    if not _PIL_AVAILABLE:
        return []
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except Exception:
        return []
    real = set()
    for n in z.namelist():
        if n.startswith('xl/media'):
            try:
                im = Image.open(io.BytesIO(z.read(n)))
                if im.size[0] >= 200 and im.size[1] >= 150:
                    real.add(n.split('/')[-1])
            except Exception:
                pass
    placed = []
    for d in [n for n in z.namelist() if re.match(r'xl/drawings/drawing\d+\.xml$', n)]:
        try:
            rels = z.read(d.replace('drawings/', 'drawings/_rels/') + '.rels').decode('utf-8', 'ignore')
        except Exception:
            continue
        rid2media = {}                       # parse Id/Target independently of order
        for rel in re.findall(r'<Relationship\b[^>]*>', rels):
            rid = re.search(r'Id="([^"]+)"', rel)
            tgt = re.search(r'Target="([^"]+)"', rel)
            if rid and tgt:
                rid2media[rid.group(1)] = tgt.group(1).split('/')[-1]
        xml = z.read(d).decode('utf-8', 'ignore')
        # Tolerate both the xdr:-prefixed (Excel) and bare (openpyxl) namespaces.
        for anc in re.findall(r'<(?:xdr:)?(?:two|one)CellAnchor\b.*?'
                              r'</(?:xdr:)?(?:two|one)CellAnchor>', xml, re.DOTALL):
            fm = re.search(r'<(?:xdr:)?from>.*?<(?:xdr:)?col>(\d+)</(?:xdr:)?col>.*?'
                           r'<(?:xdr:)?row>(\d+)</(?:xdr:)?row>', anc, re.DOTALL)
            em = re.search(r'r:embed="([^"]+)"', anc)
            if fm and em and rid2media.get(em.group(1)) in real:
                placed.append((int(fm.group(2)), int(fm.group(1)), rid2media[em.group(1)]))
    placed.sort()
    return [p[2] for p in placed]


def image_captions(data, pictures):
    """Map each embedded micrograph to its picture caption via anchor order.

    Returns {image_name: caption}; empty when the image count doesn't match the
    caption count (caller falls back to magnification matching).
    """
    order = _anchor_order(data)
    numbered = [(int(m.group(1)), c) for l, c in (pictures or [])
                for m in [_PICNUM.search(l or '')] if m]
    # None-safe key: duplicate picture numbers with a missing caption would make
    # Python compare None to str and raise (crashing the whole review).
    caps = [c for _, c in sorted(numbered, key=lambda t: (t[0], t[1] or ''))]
    if not order or len(order) != len(caps):
        return {}
    return dict(zip(order, caps))


def _picture_image_pairs(data, pictures, images):
    """Align captions to embedded micrographs by anchor order.

    Returns [(label, caption, image_entry), …] when the picture count matches the
    embedded-micrograph count (so the pairing is trustworthy), else [].
    """
    order = _anchor_order(data)
    pics = sorted(((int(m.group(1)), l, c) for l, c in (pictures or [])
                   for m in [_PICNUM.search(l or '')] if m),
                  key=lambda t: (t[0], t[1] or '', t[2] or ''))   # None-safe
    if not order or len(order) != len(pics):
        return []
    by_name = {im.get('image'): im for im in (images or [])}
    pairs = []
    for (_, label, cap), name in zip(pics, order):
        im = by_name.get(name)
        if im is not None:
            pairs.append((label, cap, im))
    return pairs


def _review_captions(parsed):
    """Caption integrity: numbering, etch status, and comment picture references."""
    findings = []
    pics = parsed.get('pictures') or []
    if not pics:
        return findings
    comment = parsed.get('comment') or ''

    nums = []
    for label, _ in pics:
        m = _PICNUM.search(label or '')
        if m:
            nums.append(int(m.group(1)))

    dups = sorted({n for n in nums if nums.count(n) > 1})
    if dups:
        findings.append(('warning', 'Captions',
                         f'Caption picture number(s) repeated: {", ".join(map(str, dups))}.'))
    if nums:
        missing = sorted(set(range(1, max(nums) + 1)) - set(nums))
        if missing:
            findings.append(('info', 'Captions',
                             f'Picture numbering gap — missing {", ".join(map(str, missing))}.'))

    no_etch = [(label or '?').rstrip(':') for label, cap in pics
               if not _ETCH_PAT.search(f"{label} {cap or ''}")]
    if no_etch:
        findings.append(('warning', 'Captions',
                         f'No etch status in caption(s): {", ".join(no_etch)}.'))
    else:
        findings.append(('pass', 'Captions', 'Every caption states an etch status.'))

    # Surface captions that explicitly state unetched / as-polished — legitimate
    # for thickness / crack work, but worth confirming for a microstructure report.
    for label, cap in pics:
        if _UNETCHED_PAT.search(f"{label} {cap or ''}"):
            findings.append(('info', 'Captions',
                             f'{(label or "?").rstrip(":")} caption states unetched / '
                             f'as-polished — confirm intended (a microstructure '
                             f'assessment is normally etched).'))

    refs = [int(m.group(1)) for m in
            re.finditer(r'pic(?:ture)?\.?\s*(?:no\.?\s*)?(\d+)', comment, re.I)]
    if refs and max(refs) > len(pics):
        findings.append(('warning', 'Captions',
                         f'Comment refers to Picture {max(refs)} but only {len(pics)} '
                         f'picture(s) are present.'))
    return findings


def review_metallurgical(parsed):
    findings = []
    findings += _review_completeness(parsed)
    findings += _review_hardness(parsed['hardness'], parsed['sample'].get('material'))
    findings += _review_composition(parsed['nominal'], parsed['actual'])
    findings += _review_comment(parsed)
    findings += _review_captions(parsed)
    return findings


# ════════════════════════════════════════════════════════════════════════
# COATING REPORTS
# ════════════════════════════════════════════════════════════════════════
def _coating_signoff(wb):
    out, loc = {}, {}
    for ws in wb.worksheets:
        for key, pat in (('prepared', r'Prepared\s*by'),
                         ('approved', r'Approved\s*by'),
                         ('date',     r'^Date\s*:')):
            if key in out:
                continue
            lbl = _find(ws, pat)
            if lbl:
                out[key], vloc = _value_right_loc(ws, *lbl)
                loc[key] = {'sheet': ws.title, 'label': lbl, 'value': vloc}
    return out, loc


def parse_coating(wb, media=0):
    # The assessment sheet is the one carrying the actual MIN/MAX design
    # limits — not the Cover sheet, whose table-of-contents also mentions
    # "Coating Coverage Assessment".
    aws = None
    for ws in wb.worksheets:
        if _find(ws, r'Design\s*limit') and _find(ws, r'Measurements'):
            aws = ws
            break

    signoff, signoff_loc = _coating_signoff(wb)
    data = {'title': None, 'report_no': None, 'component': None, 'rows': [],
            'signoff': signoff, 'media': media,
            'loc': {'sheet': aws.title if aws is not None else None,
                    'signoff': signoff_loc}}

    cover = wb.worksheets[0]
    t = _find(cover, r'Coating')
    if t:
        data['title'] = _txt(cover.cell(row=t[0], column=t[1]).value)
    rn = _find(cover, r'Report\s*No')
    if rn:
        data['report_no'] = _value_right(cover, *rn)
    # Component (e.g. "2nd Stage Bucket") sits in the cover header text.
    for ws in wb.worksheets:
        for row in ws.iter_rows(max_row=25):
            for cell in row:
                comp = _canon_component(_txt(cell.value))
                if comp:
                    data['component'] = comp
                    break
            if data['component']:
                break
        if data['component']:
            break

    if aws is None:
        return data

    meas_loc = _find(aws, r'Measurements')
    avg_loc  = _find(aws, r'Average\s*Values')
    min_loc  = _find(aws, r'^MIN$')
    max_loc  = _find(aws, r'^MAX$')
    if not (meas_loc and avg_loc and min_loc and max_loc):
        return data

    hrow = meas_loc[0]
    min_col, max_col = min_loc[1], max_loc[1]
    # Measurement value columns run from 'Measurements' up to the first summary
    # column (Average / MIN / MAX), and never include MIN/MAX/Average themselves
    # — otherwise a MIN/MAX in that span gets range-checked as a measurement.
    rights = [c for c in (avg_loc[1], min_col, max_col) if c > meas_loc[1]]
    right_bound = min(rights) if rights else aws.max_column + 1
    meas_cols = [c for c in range(meas_loc[1], right_bound) if c not in (min_col, max_col)]

    cur_min = cur_max = None
    for r in range(hrow + 1, aws.max_row + 1):
        m = _num(aws.cell(row=r, column=min_col).value)
        x = _num(aws.cell(row=r, column=max_col).value)
        if m is not None:
            cur_min = m
        if x is not None:
            cur_max = x
        cells = [(c, _num(aws.cell(row=r, column=c).value)) for c in meas_cols]
        cells = [(c, v) for c, v in cells if v is not None]
        if not cells:
            continue
        data['rows'].append({'row': r, 'values': [v for _, v in cells],
                             'cells': cells, 'min': cur_min, 'max': cur_max})
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


_ETCH_THR = 0.05   # edge-density below this ⇒ image looks unetched / very low contrast


def _edge_density(im):
    """Fraction of strong edges in the image body — high ⇒ etched, low ⇒ unetched."""
    if not _PIL_AVAILABLE:
        return None
    w, h = im.size
    c = im.crop((int(w * 0.15), int(h * 0.15), int(w * 0.85), int(h * 0.80)))
    try:
        px = list(c.filter(ImageFilter.FIND_EDGES).get_flattened_data())
    except Exception:
        return None
    return sum(1 for p in px if p > 40) / len(px) if px else None


def _read_legend_im(im):
    """OCR the burned-in legend (ID / magnification / scale-bar) of one micrograph."""
    if not _OCR_AVAILABLE:
        return {}
    w, h = im.size
    lc = im.crop((0, int(h * 0.90), int(w * 0.55), h))           # ID + magnification
    rc = im.crop((int(w * 0.72), int(h * 0.88), w, h))           # scale bar
    lblob = ' '.join(_safe_ocr(_binarize(lc, t)) for t in (110, 130, 150))
    rblob = ' '.join(_safe_ocr(_binarize(rc, t)) for t in (110, 140))

    out = {}
    job_m = _JOB_PAT.search(lblob)
    mag_val, idx = None, None
    for pat in _MAG_PATS:
        for m in pat.finditer(lblob):
            n = int(m.group(1))
            if 25 <= n <= 20000:
                mag_val, idx = n, (m.group(2) if pat.groups == 2 else None)
                break
        if mag_val is not None:
            break
    if mag_val is not None:
        out['mag'] = f'{mag_val}x'
        out['id'] = (f'{job_m.group(1)}_' if job_m else '') + f'E_{mag_val}x' + \
                    (f'-{idx}' if idx else '')
    if job_m:
        out['job'] = job_m.group(1)
    s = _SCALE_PAT.search(rblob) or _SCALE_PAT.search(lblob)
    if s:
        out['scale'] = f'{s.group(1)} µm'
    return out


def _read_measurements_im(im):
    """Read thickness labels (e.g. '42 µm') burned into the image body."""
    if not _OCR_AVAILABLE:
        return []
    w, h = im.size
    body = im.crop((0, 0, w, int(h * 0.85)))        # exclude bottom legend + scale bar
    big = body.resize((body.width * 3, body.height * 3))
    bright = big.point(lambda p: 255 if p > 200 else 0)
    txt = _safe_ocr(bright, '--psm 11')
    return sorted({int(v) for v in re.findall(r'(\d{1,3})\s*[µuμ]m', txt, re.I)})


def analyze_images(data, want_bytes=False, max_images=40):
    """Single pass over embedded micrographs.

    Returns (images, ocr_used) where each image dict carries:
      'image', 'strong', 'etched', 'measurements', optional 'mag'/'scale'/'id'/'job',
      and 'bytes'/'ext' when want_bytes is set.
    """
    images = []
    if not _PIL_AVAILABLE:
        return images, False
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        names = sorted(n for n in z.namelist() if n.startswith('xl/media'))
    except Exception:
        return images, _OCR_AVAILABLE
    for n in names[:max_images]:
        raw = z.read(n)
        try:
            im = Image.open(io.BytesIO(raw)).convert('L')
        except Exception:
            continue
        w, h = im.size
        if w < 200 or h < 150:           # skip logos / thumbnails
            continue
        strong = _edge_density(im)
        entry = {'image': n.split('/')[-1],
                 'strong': strong,
                 'etched': (strong is None) or (strong >= _ETCH_THR),
                 'measurements': _read_measurements_im(im)}
        entry.update(_read_legend_im(im))
        if want_bytes:
            entry['bytes'] = raw
            entry['ext'] = n.rsplit('.', 1)[-1].lower()
        images.append(entry)
    return images, _OCR_AVAILABLE


def read_image_legends(data, max_images=40):
    """Back-compat: the legend subset of analyze_images()."""
    images, ocr_used = analyze_images(data, max_images=max_images)
    legends = [im for im in images if im.get('mag') or im.get('scale')]
    return legends, ocr_used


def _comment_thickness_um(comment):
    """Thickness values in the comment text, normalised to µm."""
    out = set()
    for m in re.finditer(r'(\d+(?:\.\d+)?)\s*(mm|µm|um|μm)\b', comment or '', re.I):
        v = float(m.group(1))
        out.add(round(v * 1000) if m.group(2).lower() == 'mm' else round(v))
    return out


def picture_etch_verdicts(images, pictures, data):
    """Per-picture caption↔contrast verdicts via 1:1 caption/micrograph pairing.

    Returns a list of {'index', 'label', 'severity', 'note'} (possibly empty) when
    captions pair to micrographs, else None (caller falls back to the aggregate
    count). `index` is the position in `pictures`, so a caller can find the caption
    cell. Contrast is advisory, so verdicts read "verify".
    """
    if not data:
        return None
    pairs = _picture_image_pairs(data, pictures, images)
    if not pairs:
        return None
    idx_of = {}
    for i, (label, _) in enumerate(pictures or []):
        idx_of.setdefault(label, i)
    out = []
    for label, cap, im in pairs:
        if im.get('strong') is None:
            continue
        pic = (label or 'Picture').rstrip(':')
        et = caption_etchant(f"{label} {cap or ''}")
        named = et and et not in ('Unetched', 'Etched (unspecified)')
        note = None
        if named and not im.get('etched'):
            note = (f'{pic}: caption names {et} but the micrograph reads low-contrast — '
                    f'the etch may not have developed; verify (contrast is advisory).')
        elif et == 'Unetched' and im.get('etched'):
            note = (f'{pic}: caption says unetched but the micrograph reads etched-type '
                    f'contrast — verify (contrast is advisory).')
        if note:
            out.append({'index': idx_of.get(label), 'label': label,
                        'severity': 'warning', 'note': note})
    return out


def _review_etch(images, pictures, verdicts):
    """Aggregate contrast summary plus the per-picture caption↔contrast verdicts.

    `verdicts` is picture_etch_verdicts(...): a list (per-picture findings) or
    None (couldn't pair 1:1 → fall back to the aggregate unetched-vs-low count).
    """
    findings = []
    scored = [im for im in images if im.get('strong') is not None]
    if not scored:
        return findings
    n_low = sum(1 for im in scored if not im.get('etched'))
    findings.append(('info', 'Photo etch',
                     f'{len(scored) - n_low} of {len(scored)} micrograph(s) show etched-type '
                     f'contrast; {n_low} low-contrast (unetched / faint post-HT).'))
    if verdicts is not None:
        for v in verdicts:
            findings.append((v['severity'], 'Photo etch', v['note']))
    else:
        n_cap = sum(1 for label, cap in (pictures or [])
                    if _UNETCHED_PAT.search(f"{label} {cap or ''}"))
        if n_low != n_cap:
            findings.append(('info', 'Photo etch',
                             f'{n_low} micrograph(s) read as low-contrast vs {n_cap} caption(s) '
                             f'marked "unetched" — worth a glance (faint post-HT etch reads low).'))
    return findings


def _review_thickness(parsed, images):
    """A1 — surface comment vs in-photo thickness measurements for comparison."""
    findings = []
    comment_um = _comment_thickness_um(parsed.get('comment'))
    photo_um = sorted({v for im in images for v in im.get('measurements', [])})
    if not (comment_um or photo_um):
        return findings
    parts = []
    if comment_um:
        parts.append('comment ' + ', '.join(f'{v} µm' for v in sorted(comment_um)))
    if photo_um:
        parts.append('photos ' + ', '.join(f'{v} µm' for v in photo_um))
    findings.append(('info', 'Thickness', 'Thickness values — ' + '; '.join(parts) + '.'))
    if comment_um and photo_um:
        lo, hi = min(photo_um), max(photo_um)
        outliers = [v for v in sorted(comment_um) if v < lo * 0.5 or v > hi * 2]
        if outliers:
            findings.append(('warning', 'Thickness',
                             f'Comment thickness {", ".join(f"{v} µm" for v in outliers)} is far '
                             f'from the photo measurements ({lo}–{hi} µm) — verify.'))
    return findings


def _caption_mags(pictures):
    """Magnifications mentioned in the written picture captions, e.g. {'200x'}."""
    mags = set()
    for _, cap in pictures or []:
        for m in _CAP_MAG.finditer(cap or ''):
            mags.add(f'{m.group(1)}x')
    return mags


def _digit_dist(a, b):
    """Positional digit difference between two same-length strings; len-gap otherwise."""
    if len(a) != len(b):
        return max(len(a), len(b))
    return sum(x != y for x, y in zip(a, b))


def _review_legends(legends, ocr_used, caption_mags, report_job=None):
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

    # Cross-check the job number burned into the legends against the report.
    # OCR misreads single digits, so all genuine photos share one job number:
    # pass if any legend matches exactly, and only warn when readings clearly
    # diverge (≥2 digits) — that suggests a micrograph from another report.
    legend_jobs = [l['job'] for l in legends if l.get('job')]
    if report_job and report_job.isdigit() and legend_jobs:
        if report_job in legend_jobs:
            findings.append(('pass', 'Photo legends',
                             f'Micrograph legends carry the report job number ({report_job}).'))
        else:
            best = min(legend_jobs, key=lambda j: _digit_dist(j, report_job))
            if _digit_dist(best, report_job) >= 2:
                seen = ", ".join(sorted(set(legend_jobs)))
                findings.append(('warning', 'Photo legends',
                                 f'Legend job number(s) [{seen}] do not match the report job '
                                 f'{report_job} — verify the micrographs belong to this report '
                                 f'(or an OCR misread).'))
            else:
                findings.append(('info', 'Photo legends',
                                 f'Legend job numbers are within one digit of the report job '
                                 f'({report_job}) — likely OCR variance.'))
    return findings


# ════════════════════════════════════════════════════════════════════════
# FILENAME vs CONTENT  (catch a mis-named workbook)
# ════════════════════════════════════════════════════════════════════════
# Component synonyms (GE terminology): bucket≡blade (rotating), vane≡nozzle
# (stationary). Order matters — multi-word parts first.
_PART_SYNONYMS = [
    (r'transition\s*piece',  'transition piece'),
    (r'combustion\s*liner',  'combustion liner'),
    (r'\bliner\b',           'combustion liner'),
    (r'\bbucket\b|\bblade\b', 'bucket'),
    (r'\bvane\b|\bnozzle\b',  'vane'),
    (r'\bshroud\b',          'shroud'),
    (r'\bdiaphragm\b',       'diaphragm'),
    (r'\bseal\b',            'seal'),
]


def _canon_component(text):
    """Canonical 'stage + part' from free text, e.g. '2nd Stage Bucket' → '2 bucket'."""
    t = (text or '').lower()
    part = next((name for pat, name in _PART_SYNONYMS if re.search(pat, t)), None)
    if part is None:
        return None
    m = re.search(r'(\d)\s*(?:st|nd|rd|th)?\s*stage', t)
    return (f'{m.group(1)} ' if m else '') + part


def _content_job(parsed, rtype):
    """4-digit AEG job number from the report content, for either report family."""
    if rtype == 'metallurgical':
        m = re.search(r'\d{4}', parsed.get('header', {}).get('job') or '')
    else:
        m = re.search(r'\d{4}', parsed.get('report_no') or '')
    return m.group() if m else ''


def review_filename(filename, parsed, rtype):
    """Check that the workbook's name agrees with its contents."""
    findings = []
    name = re.sub(r'\.xlsx?$', '', os.path.basename(filename or ''), flags=re.I)
    if not name:
        return findings
    low = name.lower()
    matched = []

    # Job number (filename vs content).
    fjob = re.search(r'\b(\d{4})\b', name)
    cjob = _content_job(parsed, rtype)
    if fjob and cjob:
        if fjob.group(1) == cjob:
            matched.append('job')
        else:
            findings.append(('warning', 'Filename',
                             f'Filename job number {fjob.group(1)} ≠ report job {cjob}.'))

    # Report type (filename keyword vs detected type).
    if 'coating' in low and rtype == 'metallurgical':
        findings.append(('warning', 'Filename',
                         'Filename says "Coating" but the content is a metallurgical report.'))
    elif re.search(r'metallurg', low) and rtype == 'coating':
        findings.append(('warning', 'Filename',
                         'Filename says "Metallurgical" but the content is a coating report.'))
    elif ('coating' in low and rtype == 'coating') or \
         (re.search(r'metallurg', low) and rtype == 'metallurgical'):
        matched.append('type')

    # Component / part.
    fcomp = _canon_component(name)
    ccomp = (_canon_component(parsed.get('sample', {}).get('description'))
             if rtype == 'metallurgical' else parsed.get('component'))
    if fcomp and ccomp:
        if fcomp == ccomp:
            matched.append('component')
        else:
            findings.append(('warning', 'Filename',
                             f'Filename component "{fcomp}" ≠ report description "{ccomp}".'))

    # Customer (advisory, lenient — pass on any shared word ≥3 chars).
    ccust = parsed.get('header', {}).get('customer') if rtype == 'metallurgical' else None
    if ccust:
        ctoks = set(re.findall(r'[a-z]{3,}', ccust.lower()))
        if ctoks and not (ctoks & set(re.findall(r'[a-z]{3,}', low))):
            findings.append(('info', 'Filename',
                             f'Filename customer doesn’t obviously match the report customer "{ccust}".'))

    if matched and not any(c == 'Filename' and s == 'warning' for s, c, _ in findings):
        findings.append(('pass', 'Filename',
                         f'Filename agrees with the report ({", ".join(matched)}).'))
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

    findings += review_filename(filename, parsed, rtype)

    images = []
    if ocr:
        images, ocr_used = analyze_images(data)
        legends = [im for im in images if im.get('mag') or im.get('scale')]
        cap_mags = _caption_mags(parsed.get('pictures', []))
        report_job = parsed.get('header', {}).get('job')
        findings += _review_legends(legends, ocr_used, cap_mags, report_job)
        etch_verdicts = picture_etch_verdicts(images, parsed.get('pictures', []), data)
        findings += _review_etch(images, parsed.get('pictures', []), etch_verdicts)
        findings += _review_thickness(parsed, images)
        parsed['photo_etch'] = etch_verdicts or []
    parsed['images'] = images
    parsed['legends'] = [im for im in images if im.get('mag') or im.get('scale')]
    return rtype, parsed, findings


def summarize(findings):
    """Return counts per severity."""
    out = {'critical': 0, 'warning': 0, 'info': 0, 'pass': 0}
    for sev, _, _ in findings:
        out[sev] = out.get(sev, 0) + 1
    return out


def collect_highlights(parsed):
    """Map the cell-anchored findings to worksheet cells, for the annotated view.

    Returns a list of {'cell': (row, col), 'severity', 'category', 'tag', 'note'}
    on the sheet named in parsed['loc']['sheet']. `tag` is a short label drawn on
    the image; `note` is the full sentence shown in the legend. Findings with no
    single cell to point at (e.g. "no embedded images") are intentionally omitted
    here — they still appear in the textual findings list.
    """
    loc = parsed.get('loc') or {}
    out = []

    def add(cell, severity, category, tag, note):
        if cell:
            out.append({'cell': tuple(cell), 'severity': severity,
                        'category': category, 'tag': tag, 'note': note})

    def anchor(entry):
        """Prefer the value cell; fall back to the label cell."""
        entry = entry or {}
        return entry.get('value') or entry.get('label')

    # ── Composition — Actual cells out of tolerance vs Nominal ──
    deviations, systemic = _composition_deviations(
        parsed.get('nominal') or {}, parsed.get('actual') or {})
    aloc = loc.get('actual') or {}
    for el, nom, act, rel in deviations:
        sev = 'critical' if systemic else 'warning'
        add(aloc.get(el), sev, 'Composition', f'{el} {rel:+.0f}%',
            f'{el}: actual {act:g} vs nominal {nom:g} wt% ({rel:+.0f}%).')

    # ── Hardness — post-solution should not exceed pre-solution ──
    hd = parsed.get('hardness') or {}
    pre = (hd.get('pre') or {}).get('value')
    post = (hd.get('post') or {}).get('value')
    hloc = loc.get('hardness') or {}
    if pre is not None and post is not None and post > pre + 0.5:
        note = (f'Post-solution hardness ({post:g} HRC) exceeds pre-solution '
                f'({pre:g} HRC) — solution treatment normally softens the material.')
        for key in ('pre', 'post'):
            add(anchor(hloc.get(key)), 'warning', 'Hardness', 'post > pre', note)

    # ── Completeness — blank header fields / material ──
    hdr = parsed.get('header') or {}
    hdr_loc = loc.get('header') or {}
    for key, label in (('customer', 'Customer'), ('job', 'AEG Job No'),
                       ('machine', 'Machine type')):
        if _is_placeholder(hdr.get(key)):
            add(anchor(hdr_loc.get(key)), 'warning', 'Completeness',
                f'{label} blank', f'{label} is blank or a placeholder.')
    smp = parsed.get('sample') or {}
    if _is_placeholder(smp.get('material')):
        add(anchor((loc.get('sample') or {}).get('material')), 'warning',
            'Completeness', 'Material blank', 'Sample material/alloy not stated.')

    # ── Completeness — missing or very short comment ──
    if len((parsed.get('comment') or '').strip()) < 40:
        add(anchor(loc.get('comment')), 'warning', 'Completeness',
            'Short comment', 'Comment / discussion is missing or very short.')

    # ── Sign-off — missing fields (point at the label) ──
    # One combined note (matching the textual finding) shared across the cells,
    # so the box and the findings list don't double-report it.
    so = parsed.get('signoff') or {}
    so_loc = loc.get('signoff') or {}
    so_fields = (('met_lab', 'Met. Lab'), ('mat_eng', 'Mat. Eng'), ('date', 'Date'))
    so_missing = [label for key, label in so_fields if _is_placeholder(so.get(key))]
    if so_missing:
        so_note = f'Missing sign-off field(s): {", ".join(so_missing)}.'
        for key, label in so_fields:
            if _is_placeholder(so.get(key)):
                entry = so_loc.get(key) or {}
                add(entry.get('label') or entry.get('value'), 'warning',
                    'Sign-off', f'{label} missing', so_note)

    # ── Captions — no etch status, or explicitly unetched ──
    pics = parsed.get('pictures') or []
    ploc = loc.get('pictures') or []
    no_etch = [(label or '?').rstrip(':') for label, cap in pics
               if not _ETCH_PAT.search(f"{label} {cap or ''}")]
    no_etch_note = f'No etch status in caption(s): {", ".join(no_etch)}.'
    for i, (label, cap) in enumerate(pics):
        text = f"{label} {cap or ''}"
        entry = ploc[i] if i < len(ploc) else {}
        if not _ETCH_PAT.search(text):
            add(anchor(entry), 'warning', 'Captions', 'No etch status', no_etch_note)
        elif _UNETCHED_PAT.search(text):
            add(anchor(entry), 'info', 'Captions', 'Unetched',
                f'{(label or "?").rstrip(":")} caption states unetched / as-polished — '
                f'confirm intended (a microstructure assessment is normally etched).')

    # ── Photo etch — per-picture caption↔contrast mismatch (anchor to caption) ──
    for v in parsed.get('photo_etch') or []:
        idx = v.get('index')
        entry = ploc[idx] if (idx is not None and 0 <= idx < len(ploc)) else {}
        add(anchor(entry), v.get('severity', 'warning'), 'Photo etch', 'etch?', v['note'])

    # ── Thickness — comment value far from the in-photo measurements ──
    comment_um = _comment_thickness_um(parsed.get('comment'))
    photo_um = sorted({u for im in (parsed.get('images') or [])
                       for u in im.get('measurements', [])})
    if comment_um and photo_um:
        lo_p, hi_p = min(photo_um), max(photo_um)
        outliers = [u for u in sorted(comment_um) if u < lo_p * 0.5 or u > hi_p * 2]
        if outliers:
            add(anchor(loc.get('comment')), 'warning', 'Thickness', 'thickness?',
                f'Comment thickness {", ".join(f"{u} µm" for u in outliers)} is far from '
                f'the photo measurements ({lo_p}–{hi_p} µm) — verify.')

    # ── Coating — thickness measurements outside design limits ──
    for entry in parsed.get('rows') or []:
        lo, hi = entry.get('min'), entry.get('max')
        if lo is None or hi is None:
            continue
        for col, v in entry.get('cells', []):
            if not (lo <= v <= hi):
                add((entry['row'], col), 'critical', 'Coating', f'{v:g} mm',
                    f'Row {entry["row"]}: thickness {v:g} mm outside design '
                    f'limit {lo:g}–{hi:g} mm.')

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
