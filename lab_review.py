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

# Advisory hardness ranges (HRC) by alloy. ADVISORY ONLY — these are typical
# values; verify against the controlling specification and edit as needed.
# Alloys not listed here are simply reported without a range check.
HARDNESS_REF = {
    'GTD-111': (32, 42),
    'GTD-741': (25, 40),
    'RENE-80': (28, 42),
    'IN-738':  (30, 42),
    'IN738':   (30, 42),
    'NIMONIC-263': (20, 35),
    'NI-263':  (20, 35),
}

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

    if pre is not None and post is not None and post > pre + 0.5:
        findings.append(('warning', 'Hardness',
                         f'Post-solution hardness ({post:g}) exceeds pre-solution ({pre:g}) — '
                         'solution treatment normally softens the material.'))

    key = re.sub(r'\s+', '', (material or '').upper())
    rng = HARDNESS_REF.get(key)
    if rng:
        lo, hi = rng
        for label, val in (('Pre-solution', pre), ('Post-solution', post)):
            if val is not None and not (lo <= val <= hi):
                findings.append(('info', 'Hardness',
                                 f'{label} {val:g} HRC outside advisory range {lo}–{hi} HRC '
                                 f'for {material} (advisory — verify vs spec).'))
    if not findings:
        parts = [f'{k}={v["value"]:g}' for k, v in hardness.items() if v.get('value') is not None]
        findings.append(('pass', 'Hardness', f'Hardness recorded ({", ".join(parts)} HRC).'))
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
# PUBLIC ENTRY POINT
# ════════════════════════════════════════════════════════════════════════
def _media_count(data):
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        return sum(1 for n in z.namelist() if n.startswith('xl/media'))
    except Exception:
        return 0


def review_report(filename, data):
    """Review one report. Returns (report_type, parsed, findings)."""
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


if __name__ == '__main__':
    main()
