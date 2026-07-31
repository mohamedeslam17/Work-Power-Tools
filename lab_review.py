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
from collections import Counter

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
    comment_picture_refs,
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
_EXCEL_ERRORS = {
    '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A',
    '#GETTING_DATA',
}
_NOT_QUANTIFIED = re.compile(
    r'^(?:<\s*(?:lod|loq|mdl)|(?:below|under)\s+(?:detection|quantification)'
    r'|not\s+detected|n\.?d\.?)$',
    re.I,
)


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
    text = _txt(v)
    return text.lower() in _PLACEHOLDERS or text.upper() in _EXCEL_ERRORS


def _excel_errors(wb):
    """Visible Excel error values, including cached formula failures."""
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = _txt(cell.value).upper()
                if value in _EXCEL_ERRORS:
                    out.append({
                        'sheet': ws.title,
                        'cell': cell.coordinate,
                        'row': cell.row,
                        'col': cell.column,
                        'value': value,
                    })
    return out


def _formula_issues(wb):
    """Formula constructs that make a released workbook non-self-contained."""
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                formula = _txt(cell.value)
                if cell.data_type == 'f' and re.search(r'\[[^\]]+\][^!]*!', formula):
                    out.append({
                        'sheet': ws.title,
                        'cell': cell.coordinate,
                        'row': cell.row,
                        'col': cell.column,
                        'formula': formula,
                        'kind': 'external-reference',
                    })
    return out


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

    for key, substr in (('sample_no', 'sample nr'), ('description', 'description'),
                        ('serial', 's/n'),
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
    """Extract numeric composition plus the table's raw/header metadata.

    Handles the two layouts seen in practice: element headers in the
    '(Nominal/Actual)' label row with values below, OR element headers in the
    row above the label with values in the label row. Tolerates the 'Minimal'
    typo for 'Nominal'.
    """
    pat = r'\(\s*(?:Nominal|Minimal)\s*\)' if which == 'Nominal' else r'\(\s*Actual\s*\)'
    lbl = _find(ws, pat)
    comp, loc = {}, {}
    meta = {'entries': [], 'duplicate_headers': []}
    if not lbl:
        return comp, loc, meta
    lrow = lbl[0]

    def elem_cells(r):
        return [c for c in ws[r] if r >= 1 and _txt(c.value).capitalize() in ELEMENTS]

    if len(elem_cells(lrow)) >= 2:           # elements in label row, values below
        ehdr, vrow = lrow, lrow + 1
    elif lrow > 1 and len(elem_cells(lrow - 1)) >= 2:   # elements above, values in label row
        ehdr, vrow = lrow - 1, lrow
    else:
        return comp, loc, meta

    header_cells = elem_cells(ehdr)
    counts = Counter(_txt(cell.value).capitalize() for cell in header_cells)
    meta['duplicate_headers'] = sorted(el for el, count in counts.items() if count > 1)
    for cell in header_cells:
        raw = ws.cell(row=vrow, column=cell.column).value
        val = _num(raw)
        el = _txt(cell.value).capitalize()
        entry = {
            'element': el,
            'raw': _txt(raw),
            'value': val,
            'header_cell': cell.coordinate,
            'value_cell': ws.cell(row=vrow, column=cell.column).coordinate,
            'row': vrow,
            'col': cell.column,
        }
        meta['entries'].append(entry)
        if val is not None:
            comp[el] = val
            loc[el] = (vrow, cell.column)
        elif el not in loc:
            loc[el] = (vrow, cell.column)
    return comp, loc, meta


def _comment(ws):
    lbl = _find(ws, r'^Comments?\s*:')
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


def parse_metallurgical(wb, media=0, micrograph_count=None):
    ws = _met_sheet(wb)
    header, lh    = _header(ws)
    sample, ls    = _sample(ws)
    hardness, lhd = _hardness(ws)
    nominal, ln, nominal_meta = _composition(ws, 'Nominal')
    actual, la, actual_meta    = _composition(ws, 'Actual')
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
        'composition_meta': {
            'nominal': nominal_meta,
            'actual': actual_meta,
        },
        'coating':   coating,
        'comment':   comment,
        'pictures':  pictures,
        'signoff':   signoff,
        'media':     media,
        'micrograph_count': micrograph_count,
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


def _deviation_severity(rel, dev):
    """Severity of one element's deviation, independent of how many other
    elements are also off — a single element gutted by 88% is critical on
    its own, not just when enough *other* elements are also out of spec."""
    if abs(rel) >= COMP_CRIT_REL and abs(dev) >= COMP_CRIT_ABS:
        return 'critical'
    return 'warning'


def _composition_deviations(nominal, actual):
    """Elements whose Actual is out of tolerance vs Nominal.

    Returns (deviations, systemic) where deviations is a list of
    (element, nominal, actual, rel%, severity) and `systemic` is True when so
    many elements are off that the actual material likely isn't the stated
    alloy (a separate, count-based signal from any single element's own
    severity — either can independently make the overall verdict critical).
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
            deviations.append((el, nom, act, rel, _deviation_severity(rel, dev)))
    n_dev, n_common = len(deviations), len(common)
    systemic = n_dev >= 4 or (n_dev >= 3 and n_common and n_dev / n_common >= 0.5)
    return deviations, systemic


def _review_composition(nominal, actual, composition_meta=None):
    findings = []
    if not nominal or not actual:
        findings.append(('warning', 'Composition',
                         'Could not read both Nominal and Actual composition tables.'))
        # Keep reviewing the table structure: a partially parsed table can still
        # reveal a duplicate header or an explicit <LOD result.

    common = sorted(set(nominal) & set(actual))
    deviations, systemic = _composition_deviations(nominal, actual)
    n_dev, n_common = len(deviations), len(common)
    composition_meta = composition_meta or {}
    actual_meta = composition_meta.get('actual') or {}
    entries = actual_meta.get('entries') or []
    actual_headers = {e.get('element') for e in entries if e.get('element')}

    duplicates = actual_meta.get('duplicate_headers') or []
    if duplicates:
        findings.append(('critical', 'Composition',
                         f'Actual composition table repeats element header(s): '
                         f'{", ".join(duplicates)} — at least one column is mislabeled, so the '
                         f'chemistry cannot be accepted as reported.'))

    for entry in entries:
        raw = entry.get('raw') or ''
        el = entry.get('element')
        if not el or entry.get('value') is not None or not _NOT_QUANTIFIED.match(raw):
            continue
        nom = nominal.get(el)
        if nom is not None and nom >= 0.5:
            findings.append(('critical', 'Composition',
                             f'{el}: actual result is "{raw}" while nominal is {nom:g} wt% — '
                             f'a major alloying element was not quantified.'))
        else:
            findings.append(('warning', 'Composition',
                             f'{el}: actual result is "{raw}" — verify the analytical method '
                             f'and detection limit.'))

    numeric_entries = [e['value'] for e in entries if e.get('value') is not None]
    if len(numeric_entries) >= 5:
        total = sum(numeric_entries)
        if total < 95.0 or total > 105.0:
            findings.append(('critical', 'Composition',
                             f'Actual composition totals {total:.2f} wt%, outside the '
                             f'95–105 wt% sanity band — check missing/misaligned columns.'))

    # _composition() computes duplicate_headers/entries for the Nominal table
    # too (nominal_meta below), but until now nothing downstream ever looked
    # at it — a mislabeled/duplicated column in the *spec* side of the table
    # was invisible even though the identical fault on the Actual side (right
    # above) is a critical finding. Mirror both of the Actual-side structural
    # checks for Nominal.
    nominal_meta = composition_meta.get('nominal') or {}
    nominal_entries = nominal_meta.get('entries') or []
    nominal_duplicates = nominal_meta.get('duplicate_headers') or []
    if nominal_duplicates:
        findings.append(('critical', 'Composition',
                         f'Nominal composition table repeats element header(s): '
                         f'{", ".join(nominal_duplicates)} — at least one column is mislabeled, '
                         f'so the spec cannot be accepted as reported.'))
    nominal_numeric = [e['value'] for e in nominal_entries if e.get('value') is not None]
    if len(nominal_numeric) >= 5:
        nominal_total = sum(nominal_numeric)
        if nominal_total < 95.0 or nominal_total > 105.0:
            findings.append(('critical', 'Composition',
                             f'Nominal composition totals {nominal_total:.2f} wt%, outside the '
                             f'95–105 wt% sanity band — the spec table may be for the wrong '
                             f'alloy or have missing/misaligned columns.'))

    # A few elements off is normal service depletion / EDS scatter → warnings,
    # unless a single element is itself far enough outside tolerance to be
    # critical regardless of how many others are affected. Many elements off
    # together signals the actual material doesn't match the stated alloy →
    # one consolidated FAIL ("verify material/grade").
    if systemic:
        worst = sorted(deviations, key=lambda d: -abs(d[3]))[:4]
        detail = ", ".join(f"{el} {rel:+.0f}%" for el, _, _, rel, _ in worst)
        findings.append(('critical', 'Composition',
                         f'{n_dev} of {n_common} elements out of tolerance ({detail} …) — actual '
                         f'composition does not match the stated alloy; verify material/grade.'))
    else:
        for el, nom, act, rel, sev in deviations:
            extra = ' — far outside tolerance, verify material/grade' if sev == 'critical' else ''
            findings.append((sev, 'Composition',
                             f'{el}: actual {act:g} vs nominal {nom:g} wt% ({rel:+.0f}%){extra}.'))

    only_nom = sorted(set(nominal) - (actual_headers or set(actual)))
    only_act = sorted((actual_headers or set(actual)) - set(nominal))
    if only_nom:
        findings.append(('info', 'Composition',
                         f'In spec but not reported in actual: {", ".join(only_nom)}.'))
    if only_act:
        major = []
        minor = []
        for el in only_act:
            values = [e.get('value') for e in entries
                      if e.get('element') == el and e.get('value') is not None]
            if values and max(abs(v) for v in values) >= 1.0:
                major.append(f'{el} {max(values, key=abs):g} wt%')
            else:
                minor.append(el)
        if major:
            findings.append(('critical', 'Composition',
                             f'Major element(s) reported but absent from the nominal table: '
                             f'{", ".join(major)} — verify the headers and stated alloy.'))
        if minor:
            findings.append(('info', 'Composition',
                             f'Reported but not in nominal spec: {", ".join(minor)}.'))
    if nominal and actual and not deviations:
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
        # collect_highlights() already treats an empty comment as "no verdict"
        # when Result defers to it (has_negative/has_positive both false on
        # ''), so the annotated view flags this — the plain findings list must
        # match instead of silently passing it via the early return below.
        result = (parsed.get('sample') or {}).get('result') or ''
        if 'see comment' in result.lower():
            findings.append(('critical', 'Disposition',
                             'Result says "See comment", but there is no comment at all '
                             '— the report gives no accept / reject / repair disposition.'))
        # A named coating type recorded in the cells (not a blank/'N/A' placeholder)
        # with zero comment text means the coating condition — thickness, coverage,
        # degradation — was never actually described anywhere in the report.
        present = (coat.get('present') or '').strip().lower()
        cell_types = _coating_types_in(coat.get('type'))
        if present != 'no' and (present == 'yes' or (cell_types and not _is_placeholder(coat.get('type')))):
            label = "/".join(sorted(cell_types)) or coat.get('type')
            findings.append(('warning', 'Comment',
                             f'Coating cell records a coating ({label}), but there is no '
                             f'comment describing its condition, thickness or coverage.'))
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
    elif 'see comment' in rlow and not neg and not pos:
        findings.append(('critical', 'Disposition',
                         'Result says "See comment", but the comment gives no clear '
                         'accept / reject / repair disposition.'))
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


def _review_evidence_claims(parsed):
    """Flag conclusions that claim more than the recorded evidence establishes."""
    comment = parsed.get('comment') or ''
    claims_mechanical_restoration = re.search(
        r'(?:restor\w*.{0,80}mechanical\s+propert|'
        r'mechanical\s+propert.{0,80}restor\w*)',
        comment,
        re.I | re.DOTALL,
    )
    if not claims_mechanical_restoration:
        return []

    # These templates only expose pre- and post-solution hardness. Neither is a
    # final aged-condition mechanical verification, and micrographs alone cannot
    # demonstrate restored mechanical properties.
    hardness = parsed.get('hardness') or {}
    if not any(key in hardness for key in ('post_age', 'final', 'aged')):
        return [('warning', 'Evidence',
                 'Comment claims that mechanical properties were restored, but the '
                 'report contains no final aged-condition mechanical result; '
                 'microstructure alone does not verify mechanical properties.')]
    return []


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
    micrographs = parsed.get('micrograph_count')
    if micrographs == 0:
        findings.append(('warning', 'Micrographs', 'No embedded images found in the workbook.'))
    elif micrographs is not None and pics:
        if micrographs < len(pics):
            findings.append(('critical', 'Micrographs',
                             f'{len(pics)} picture caption(s) but only {micrographs} embedded '
                             f'micrograph(s) — {len(pics) - micrographs} evidence image(s) '
                             f'are missing.'))
        elif micrographs > len(pics):
            findings.append(('warning', 'Micrographs',
                             f'{micrographs} embedded micrograph(s) but only {len(pics)} '
                             f'picture caption(s) — check for uncatalogued or stray images.'))

    so = parsed['signoff']
    missing = [lbl for key, lbl in (('met_lab', 'Met. Lab'), ('mat_eng', 'Mat. Eng'),
                                    ('date', 'Date')) if _is_placeholder(so.get(key))]
    if missing:
        findings.append(('warning', 'Sign-off', f'Missing sign-off field(s): {", ".join(missing)}.'))
    else:
        findings.append(('pass', 'Sign-off', 'Lab, engineer and date all present.'))
    return findings


def _identifier_tokens(text, kind):
    """Normalised identifiers from stacked/space-separated sample or S/N cells."""
    raw = _txt(text).upper()
    if not raw:
        return []
    if kind == 'sample':
        hits = re.findall(r'\bMS\s*[-_]?\s*[A-Z0-9]+\b', raw)
        return [re.sub(r'[\s_-]+', '', hit) for hit in hits]
    tokens = []
    for block in re.split(r'[\n;,]+', raw):
        for part in re.split(r'\s{2,}', block.strip()):
            token = re.sub(r'\s+', '', part)
            if len(token) >= 4 and re.fullmatch(r'[A-Z0-9-]+', token) and re.search(r'\d', token):
                tokens.append(token)
    return tokens


def _review_traceability(parsed):
    findings = []
    sample = parsed.get('sample') or {}
    sample_ids = _identifier_tokens(sample.get('sample_no'), 'sample')
    serial_ids = _identifier_tokens(sample.get('serial'), 'serial')

    for label, values in (('sample number', sample_ids), ('serial/part number', serial_ids)):
        duplicates = sorted(value for value, count in Counter(values).items() if count > 1)
        if duplicates:
            findings.append(('warning', 'Traceability',
                             f'Duplicate {label}(s): {", ".join(duplicates)}.'))

    # NOTE: a sample-number-count vs serial-number-count check was tried here
    # (PR #5) and deliberately removed (PR #7) after it false-positived on real
    # reports — a part can legitimately be sampled without a legible/recorded
    # serial. See test_sample_and_serial_count_mismatch_is_intentionally_skipped.
    # Do not re-add a hard count comparison without addressing that.

    result = sample.get('result')
    if _is_placeholder(result):
        findings.append(('warning', 'Disposition', 'Result field is blank or invalid.'))
    return findings


def _review_workbook_integrity(parsed):
    errors = parsed.get('excel_errors') or []
    formulas = parsed.get('formula_issues') or []
    findings = []
    if errors:
        shown = ', '.join(f'{e["sheet"]}!{e["cell"]}={e["value"]}' for e in errors[:6])
        extra = f' (+{len(errors) - 6} more)' if len(errors) > 6 else ''
        findings.append(('critical', 'Workbook integrity',
                         f'Excel error value(s) present: {shown}{extra}.'))
    if formulas:
        shown = ', '.join(f'{e["sheet"]}!{e["cell"]} {e["formula"]}' for e in formulas[:4])
        extra = f' (+{len(formulas) - 4} more)' if len(formulas) > 4 else ''
        findings.append(('critical', 'Workbook integrity',
                         f'External-workbook formula reference(s) make the report '
                         f'non-self-contained: {shown}{extra}.'))
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

    refs = comment_picture_refs(comment)
    if refs and max(refs) > len(pics):
        findings.append(('warning', 'Captions',
                         f'Comment refers to Picture {max(refs)} but only {len(pics)} '
                         f'picture(s) are present.'))

    return findings


def review_metallurgical(parsed):
    findings = []
    findings += _review_workbook_integrity(parsed)
    findings += _review_completeness(parsed)
    findings += _review_traceability(parsed)
    findings += _review_hardness(parsed['hardness'], parsed['sample'].get('material'))
    findings += _review_composition(
        parsed['nominal'], parsed['actual'], parsed.get('composition_meta'))
    findings += _review_comment(parsed)
    findings += _review_evidence_claims(parsed)
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
    findings = _review_workbook_integrity(parsed)
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

    micrographs = parsed.get('micrograph_count')
    if micrographs == 0:
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
    re.compile(r'(\d{1,5})\s*[xX%]\s*[-_]\s*(\d)'),   # 500x-1  (magnification + index)
    re.compile(r'E\s*[_ €F]?\s*(\d{1,5})\s*[xX%]'),   # E_500x
    re.compile(r'(?<![\d.])(\d{1,5})\s*[xX%]'),       # 500x
]
# Digit-bounded, not \b-bounded: "7420_E_500x-4" has no \w/\W boundary after
# the job number (underscore is a word char too), so \b(\d{4})\b never matches
# the documented convention above.
_JOB_PAT   = re.compile(r'(?<!\d)(\d{4})(?!\d)')
_SCALE_PAT = re.compile(r'(\d{1,3})\s*[µuμyptwb]+m', re.I)
_CAP_MAG   = re.compile(r'(\d{1,5})\s*[xX]\b')


def _magnification_from_ocr(text):
    """One plausible magnification from a single OCR preprocessing pass."""
    for pattern in _MAG_PATS:
        match = pattern.search(text or '')
        if match:
            value = int(match.group(1))
            if 1 <= value <= 50000:
                return value
    return None


def _select_magnification(ocr_reads):
    """Select a stable OCR magnification by voting across preprocessing reads.

    A number is accepted only when at least two independent threshold passes
    agree and there is no tie. One isolated read such as 600x is retained in
    the vote metadata but never promoted to a reported magnification.
    """
    candidates = [
        value for value in (_magnification_from_ocr(text) for text in ocr_reads)
        if value is not None
    ]
    votes = Counter(candidates)
    if not votes:
        return None, {}
    ranked = votes.most_common()
    value, count = ranked[0]
    tied = len(ranked) > 1 and ranked[1][1] == count
    return (value if count >= 2 and not tied else None), dict(votes)


def _caption_magnification(caption):
    """Single written caption magnification (the report's source of record)."""
    values = {int(match.group(1)) for match in _CAP_MAG.finditer(caption or '')}
    return f'{next(iter(values))}x' if len(values) == 1 else None


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
    lreads = [_safe_ocr(_binarize(lc, t)) for t in (110, 130, 150)]
    lblob = ' '.join(lreads)
    rblob = ' '.join(_safe_ocr(_binarize(rc, t)) for t in (110, 140))

    out = {}
    job_m = _JOB_PAT.search(lblob)
    mag_val, votes = _select_magnification(lreads)
    if votes:
        out['mag_votes'] = {f'{value}x': count for value, count in sorted(votes.items())}
        out['mag_read_count'] = len(lreads)
    if mag_val is not None:
        mag = f'{mag_val}x'
        out['ocr_mag'] = mag
        out['mag'] = mag
        out['mag_source'] = 'ocr'
        out['mag_confidence'] = votes[mag_val] / len(lreads)
        idx = None
        for read in lreads:
            match = _MAG_PATS[0].search(read or '')
            if match and int(match.group(1)) == mag_val:
                idx = match.group(2)
                break
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
    legends = [im for im in images
               if im.get('ocr_mag') or im.get('scale') or im.get('job') or im.get('id')]
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


def picture_magnification_verdicts(images, pictures, data):
    """Reconcile each written caption with its paired micrograph's stable OCR.

    The written caption is the report's source of record. OCR is accepted only
    after multi-pass consensus and is retained as secondary evidence. A
    disagreement is therefore a warning to inspect the burned-in legend, never
    a rejection of either number.
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
        caption_mag = _caption_magnification(cap)
        ocr_mag = im.get('ocr_mag') or (
            im.get('mag') if im.get('mag_source') == 'ocr' else None)
        if caption_mag:
            im['caption_mag'] = caption_mag
            im['mag'] = caption_mag
            im['mag_source'] = 'caption'
        if not (caption_mag and ocr_mag and caption_mag != ocr_mag):
            continue
        # An ID built from the disputed OCR magnitude is not safe to use for
        # duplicate-ID checks or downstream metadata.
        im.pop('id', None)
        votes = im.get('mag_votes') or {}
        count = votes.get(ocr_mag)
        total = im.get('mag_read_count')
        consensus = (
            f' in {count} of {total} preprocessing passes'
            if count is not None and total else ''
        )
        pic = (label or 'Picture').rstrip(':')
        note = (
            f'{pic}: written caption states {caption_mag}, while legend OCR '
            f'consistently read {ocr_mag}{consensus}. Treat the caption as the '
            f'recorded value and verify the burned-in legend; OCR may be wrong.'
        )
        out.append({
            'index': idx_of.get(label),
            'label': label,
            'severity': 'warning',
            'note': note,
        })
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
    return {mag for _, cap in pictures or []
            for mag in [_caption_magnification(cap)] if mag}


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

    img_mags = sorted({
        l.get('ocr_mag') or l.get('mag')
        for l in legends if l.get('ocr_mag') or l.get('mag')
    },
                      key=lambda s: int(s[:-1]))
    findings.append(('info', 'Photo legends',
                     f'Read legends from {len(legends)} micrograph(s); '
                     f'stable OCR magnifications: '
                     f'{", ".join(img_mags) if img_mags else "n/a"}.'))

    # Two visibly different micrographs sharing one burned-in ID means at
    # least one was captured/labelled wrong — the AEG "-n" suffix exists to
    # keep these unique, so a repeat is never expected.
    ids = [l['id'] for l in legends if l.get('id')]
    dupes = sorted({i for i in ids if ids.count(i) > 1})
    if dupes:
        findings.append(('warning', 'Photo legends',
                         f'Multiple micrographs share the identical burned-in ID '
                         f'{", ".join(dupes)} — verify they aren\'t the same image mislabeled.'))

    # Cross-check the job number burned into every legend against the report.
    # A previous implementation passed the whole set as soon as ONE image
    # matched, which hid a stray image from another job. Mixed exact/non-exact
    # readings are therefore always surfaced. A uniform one-digit mismatch is
    # still informational because it can be a repeatable OCR error.
    legend_jobs = [l['job'] for l in legends if l.get('job')]
    if report_job and report_job.isdigit() and legend_jobs:
        distinct = sorted(set(legend_jobs))
        mismatched = sorted({job for job in legend_jobs if job != report_job})
        if not mismatched:
            findings.append(('pass', 'Photo legends',
                             f'Micrograph legends carry the report job number ({report_job}).'))
        elif report_job in legend_jobs:
            findings.append(('warning', 'Photo legends',
                             f'Mixed micrograph job numbers: report {report_job}, but some '
                             f'images read {", ".join(mismatched)} — verify that no image was '
                             f'copied from another report.'))
        else:
            best = min(legend_jobs, key=lambda j: _digit_dist(j, report_job))
            if len(distinct) > 1 or _digit_dist(best, report_job) >= 2:
                seen = ", ".join(distinct)
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
# REPORT TITLE / FILENAME vs CONTENT  (catch a misidentified workbook)
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


def _component_identity(text):
    """Return (stage_number, canonical_part) from free report-title text."""
    raw = text or ''
    # Filenames use '_' as the word separator (e.g. "..._1st_Stage_Bucket_...").
    # Regex '_' is a \w character, so a \b-anchored pattern like \bbucket\b or
    # \b(\d{1,2})...stage\b never matches across it — every underscore-joined
    # filename would silently fail to match a component/stage it clearly names.
    # Normalising '_' to a space first makes the existing \b patterns work for
    # both filenames and the space-separated internal description text.
    lowered = raw.replace('_', ' ').lower()
    part = next((name for pat, name in _PART_SYNONYMS if re.search(pat, lowered)), None)
    stage_match = (
        re.search(r'\b(\d{1,2})\s*(?:st|nd|rd|th)?\s*(?:stage|stg)\b', lowered)
        or re.search(r'\b(?:stage|stg)\s*(\d{1,2})\b', lowered)
    )
    stage = int(stage_match.group(1)) if stage_match else None
    return stage, part


def _canon_component(text):
    """Canonical 'stage + part' from free text, e.g. '2nd Stage Bucket' → '2 bucket'."""
    stage, part = _component_identity(text)
    if part is None:
        return None
    return (f'{stage} ' if stage is not None else '') + part


def _canon_machine(text):
    """Canonical machine/set designation from a filename or report header.

    The aliases below reflect the forms used by the supplied reports:
    FS.7 ≡ MS7001 and 7FA ≡ MS7001FA. V-series names remain explicit.
    """
    raw = (text or '').upper()

    v_model = re.search(r'(?<![A-Z0-9])V\s*(\d{2})\s*[.\-]?\s*(\d)(?!\d)', raw)
    if v_model:
        return f'V{v_model.group(1)}.{v_model.group(2)}'

    ms_model = re.search(r'(?<![A-Z0-9])MS\s*([679])\s*001\s*([A-Z]{1,3})?', raw)
    if ms_model:
        frame, suffix = ms_model.group(1), ms_model.group(2) or ''
        return f'{frame}{suffix}' if suffix else f'MS{frame}001'

    fs_model = re.search(r'(?<![A-Z0-9])FS\s*[.\-]?\s*([679])(?!\d)', raw)
    if fs_model:
        return f'MS{fs_model.group(1)}001'

    short_f = re.search(r'(?<![A-Z0-9])([679])\s*F\s*([A-Z]?)(?![A-Z0-9])', raw)
    if short_f:
        return f'{short_f.group(1)}F{short_f.group(2)}'

    return None


def _content_job(parsed, rtype):
    """4-digit AEG job number from the report content, for either report family."""
    if rtype == 'metallurgical':
        m = re.search(r'\d{4}', parsed.get('header', {}).get('job') or '')
    else:
        m = re.search(r'\d{4}', parsed.get('report_no') or '')
    return m.group() if m else ''


def review_filename(filename, parsed, rtype):
    """Check that the report title/filename agrees with its contents."""
    findings = []
    name = re.sub(r'\.xlsx?$', '', os.path.basename(filename or ''), flags=re.I)
    if not name:
        return findings
    low = name.lower()
    matched = []
    category = 'Title identity'

    # Job number (title vs content). Digit-bounded, not \b-bounded — an
    # underscore right after the digits (e.g. "7504_AEN_Saudi...") is a word
    # char too, so \b(\d{4})\b would never match this very common convention.
    fjob = re.search(r'(?<!\d)(\d{4})(?!\d)', name)
    cjob = _content_job(parsed, rtype)
    if fjob and cjob:
        if fjob.group(1) == cjob:
            matched.append('job')
        else:
            findings.append(('critical', category,
                             f'Report title job number {fjob.group(1)} does not match '
                             f'the internal AEG job number {cjob}.'))
    elif cjob and not fjob:
        findings.append(('warning', category,
                         f'Report title does not state the internal AEG job number {cjob}.'))
    elif fjob and not cjob:
        findings.append(('warning', category,
                         f'Report title states job {fjob.group(1)}, but no internal '
                         f'AEG job number was found.'))

    # Report type.
    if 'coating' in low and rtype == 'metallurgical':
        findings.append(('critical', category,
                         'Report title says "Coating" but the content is metallurgical.'))
    elif re.search(r'metallurg', low) and rtype == 'coating':
        findings.append(('critical', category,
                         'Report title says "Metallurgical" but the content is a coating report.'))
    elif ('coating' in low and rtype == 'coating') or \
         (re.search(r'metallurg', low) and rtype == 'metallurgical'):
        matched.append('type')

    content_description = (
        parsed.get('sample', {}).get('description')
        if rtype == 'metallurgical'
        else parsed.get('component')
    ) or ''
    fstage, fpart = _component_identity(name)
    cstage, cpart = _component_identity(content_description)

    # Stage and component are checked independently so a title cannot pass by
    # getting one of the two right.
    if fstage is not None and cstage is not None:
        if fstage == cstage:
            matched.append('stage')
        else:
            findings.append(('critical', category,
                             f'Report title says Stage {fstage}, but the internal '
                             f'component description says Stage {cstage} '
                             f'("{content_description}").'))
    elif cstage is not None and fstage is None:
        findings.append(('warning', category,
                         f'Report title does not state Stage {cstage} from the internal '
                         f'component description "{content_description}".'))
    elif fstage is not None and cstage is None:
        findings.append(('warning', category,
                         f'Report title says Stage {fstage}, but the internal component '
                         f'description does not state a stage.'))

    if fpart and cpart:
        if fpart == cpart:
            matched.append('component')
        else:
            findings.append(('critical', category,
                             f'Report title component "{fpart}" does not match the '
                             f'internal component "{cpart}" ("{content_description}").'))
    elif cpart and not fpart:
        findings.append(('warning', category,
                         f'Report title does not state the internal component "{cpart}".'))
    elif fpart and not cpart:
        findings.append(('warning', category,
                         f'Report title states component "{fpart}", but no internal '
                         f'component name was found.'))

    # Machine/set designation: cover both the explicit V-series wording and GE
    # aliases used in these files (FS.7/MS7001 and 7FA/MS7001FA).
    content_machine = (
        parsed.get('header', {}).get('machine') if rtype == 'metallurgical' else None
    ) or ''
    fmachine = _canon_machine(name)
    cmachine = _canon_machine(content_machine)
    if rtype == 'metallurgical':
        if fmachine and cmachine:
            if fmachine == cmachine:
                matched.append('machine/set')
            else:
                findings.append(('critical', category,
                                 f'Report title machine/set "{fmachine}" does not match '
                                 f'the internal Machine Type "{content_machine}" '
                                 f'("{cmachine}").'))
        elif cmachine and not fmachine:
            findings.append(('warning', category,
                             f'Report title does not state the internal machine/set '
                             f'"{content_machine}".'))
        elif fmachine and not cmachine:
            findings.append(('warning', category,
                             f'Report title states machine/set "{fmachine}", but the '
                             f'internal Machine Type is blank or unrecognised.'))

    # Customer (advisory, lenient — pass on any shared word ≥3 chars).
    ccust = parsed.get('header', {}).get('customer') if rtype == 'metallurgical' else None
    if ccust:
        ctoks = set(re.findall(r'[a-z]{3,}', ccust.lower()))
        if ctoks and not (ctoks & set(re.findall(r'[a-z]{3,}', low))):
            findings.append(('info', category,
                             f'Report title customer does not obviously match the '
                             f'internal customer "{ccust}".'))

    if matched and not any(
            c == category and s in ('critical', 'warning') for s, c, _ in findings):
        findings.append(('pass', category,
                         f'Report title agrees with the content ({", ".join(matched)}).'))
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
    wb_formulas = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    rtype = detect_type(wb)
    media = _media_count(data)
    excel_errors = _excel_errors(wb)
    formula_issues = _formula_issues(wb_formulas)
    micrograph_names = _anchor_order(data) if _PIL_AVAILABLE else None
    micrograph_count = len(micrograph_names) if micrograph_names is not None else None

    if rtype == 'coating':
        parsed = parse_coating(wb, media)
        parsed['excel_errors'] = excel_errors
        parsed['formula_issues'] = formula_issues
        parsed['micrograph_count'] = micrograph_count
        findings = review_coating(parsed)
    elif rtype == 'metallurgical':
        parsed = parse_metallurgical(wb, media, micrograph_count)
        parsed['excel_errors'] = excel_errors
        parsed['formula_issues'] = formula_issues
        findings = review_metallurgical(parsed)
    else:
        parsed = {'excel_errors': excel_errors, 'formula_issues': formula_issues,
                  'micrograph_count': micrograph_count}
        findings = [('warning', 'Format',
                     'Unrecognised layout — not classified as a metallurgical or coating report.')]
        findings += _review_workbook_integrity(parsed)

    title_findings = review_filename(filename, parsed, rtype)
    parsed['title_findings'] = title_findings
    findings += title_findings

    images = []
    if ocr:
        images, ocr_used = analyze_images(data)
        mag_verdicts = picture_magnification_verdicts(
            images, parsed.get('pictures', []), data)
        for verdict in mag_verdicts or []:
            findings.append((verdict['severity'], 'Magnification', verdict['note']))
        parsed['photo_magnification'] = mag_verdicts or []
        legends = [im for im in images
                   if im.get('ocr_mag') or im.get('scale')
                   or im.get('job') or im.get('id')]
        cap_mags = _caption_mags(parsed.get('pictures', []))
        report_job = parsed.get('header', {}).get('job')
        findings += _review_legends(legends, ocr_used, cap_mags, report_job)
        etch_verdicts = picture_etch_verdicts(images, parsed.get('pictures', []), data)
        findings += _review_etch(images, parsed.get('pictures', []), etch_verdicts)
        findings += _review_thickness(parsed, images)
        parsed['photo_etch'] = etch_verdicts or []
    parsed['images'] = images
    parsed['legends'] = [im for im in images
                         if im.get('ocr_mag') or im.get('scale')
                         or im.get('job') or im.get('id')]
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

    # ── Workbook integrity — visible errors / external formula references ──
    for error in parsed.get('excel_errors') or []:
        add((error.get('row'), error.get('col')), 'critical', 'Workbook integrity',
            error.get('value') or 'Excel error',
            f'{error.get("sheet")}!{error.get("cell")} contains '
            f'{error.get("value")}.')
    for issue in parsed.get('formula_issues') or []:
        add((issue.get('row'), issue.get('col')), 'critical', 'Workbook integrity',
            'external formula',
            f'{issue.get("sheet")}!{issue.get("cell")} uses external-workbook '
            f'formula {issue.get("formula")}.')

    # ── Report title identity — point at the conflicting internal field ──
    title_header_loc = loc.get('header') or {}
    title_sample_loc = loc.get('sample') or {}
    for severity, category, note in parsed.get('title_findings') or []:
        if severity == 'pass':
            continue
        lowered = note.lower()
        cell, tag = None, 'title mismatch'
        if 'job number' in lowered or 'states job' in lowered:
            cell, tag = anchor(title_header_loc.get('job')), 'job mismatch'
        elif 'machine/set' in lowered or 'machine type' in lowered:
            cell, tag = anchor(title_header_loc.get('machine')), 'machine mismatch'
        elif 'stage' in lowered or 'component' in lowered:
            cell, tag = anchor(title_sample_loc.get('description')), 'component mismatch'
        elif 'customer' in lowered:
            cell, tag = anchor(title_header_loc.get('customer')), 'customer?'
        add(cell, severity, category, tag, note)

    # ── Composition — Actual cells out of tolerance vs Nominal ──
    deviations, systemic = _composition_deviations(
        parsed.get('nominal') or {}, parsed.get('actual') or {})
    aloc = loc.get('actual') or {}
    for el, nom, act, rel, sev in deviations:
        sev = 'critical' if systemic else sev
        add(aloc.get(el), sev, 'Composition', f'{el} {rel:+.0f}%',
            f'{el}: actual {act:g} vs nominal {nom:g} wt% ({rel:+.0f}%).')
    comp_meta = (parsed.get('composition_meta') or {}).get('actual') or {}
    duplicate_headers = set(comp_meta.get('duplicate_headers') or [])
    for entry in comp_meta.get('entries') or []:
        el, raw = entry.get('element'), entry.get('raw') or ''
        cell = (entry.get('row'), entry.get('col'))
        if el in duplicate_headers:
            add(cell, 'critical', 'Composition', f'duplicate {el}',
                f'Actual composition table repeats element header {el}.')
        if entry.get('value') is None and _NOT_QUANTIFIED.match(raw):
            nom = (parsed.get('nominal') or {}).get(el)
            sev = 'critical' if nom is not None and nom >= 0.5 else 'warning'
            note = (f'{el}: actual result is "{raw}"'
                    + (f' while nominal is {nom:g} wt%.' if nom is not None else '.'))
            add(cell, sev, 'Composition', f'{el} {raw}', note)
    nom_meta = (parsed.get('composition_meta') or {}).get('nominal') or {}
    nom_duplicate_headers = set(nom_meta.get('duplicate_headers') or [])
    for entry in nom_meta.get('entries') or []:
        el = entry.get('element')
        if el in nom_duplicate_headers:
            add((entry.get('row'), entry.get('col')), 'critical', 'Composition',
                f'duplicate {el}', f'Nominal composition table repeats element header {el}.')

    # ── Hardness — post-solution should not exceed pre-solution ──
    hd = parsed.get('hardness') or {}
    pre = (hd.get('pre') or {}).get('value')
    post = (hd.get('post') or {}).get('value')
    hloc = loc.get('hardness') or {}
    if pre is not None and post is not None and post > pre + 2:
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

    # ── Disposition — "See comment" must lead to an actual verdict ──
    result = smp.get('result') or ''
    comment = parsed.get('comment') or ''
    if 'see comment' in result.lower():
        has_negative = bool(re.search(
            r'not\s+suitable|unsuitable|not\s+recommend|\breject|\bscrap|'
            r'beyond\s+repair|non[-\s]?conform|unacceptable', comment, re.I))
        has_positive = bool(re.search(
            r'(?<!not )(?:\bsuitable for|\bacceptable|recommended for|'
            r'reconditi|fit for service|return to service)', comment, re.I))
        if not has_negative and not has_positive:
            note = ('Result says "See comment", but the comment gives no clear '
                    'accept / reject / repair disposition.')
            sloc = loc.get('sample') or {}
            add(anchor(sloc.get('result')), 'critical', 'Disposition', 'no verdict', note)
            add(anchor(loc.get('comment')), 'critical', 'Disposition', 'no verdict', note)

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
    micrographs = parsed.get('micrograph_count')
    if micrographs is not None and micrographs < len(pics):
        note = (f'{len(pics)} picture caption(s) but only {micrographs} embedded '
                f'micrograph(s) — {len(pics) - micrographs} evidence image(s) are missing.')
        for entry in ploc[micrographs:]:
            add(anchor(entry), 'critical', 'Micrographs', 'image missing', note)
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

    # ── Magnification — stable legend OCR disagrees with paired caption ──
    for verdict in parsed.get('photo_magnification') or []:
        idx = verdict.get('index')
        entry = ploc[idx] if (idx is not None and 0 <= idx < len(ploc)) else {}
        add(anchor(entry), verdict.get('severity', 'warning'), 'Magnification',
            'legend OCR?', verdict['note'])

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
