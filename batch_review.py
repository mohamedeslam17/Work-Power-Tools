#!/usr/bin/env python3
"""
Batch audit of a folder of AEG lab reports — drives tool tuning.

For each .xlsx it runs the reviewer and aggregates:
  * report-type detection + parse coverage (which fields were extracted)
  * the distribution of findings by severity and category
  * "discoveries" to feed back into the rules — alloys with no hardness
    reference, plus every etchant / coating-type / component / customer /
    machine / composition element seen
  * anomalies that signal a parser gap (unclassified layout, missing material,
    no composition table, empty comment, no sign-off, hard errors)

Usage:
    python3 batch_review.py <folder> [--ocr]     # --ocr also audits micrographs
"""
import os
import sys
import glob
import collections

from lab_review import (review_report, HARDNESS_REF, _alloy_key,
                        caption_etchant, _coating_types_in)


def _present(v):
    return bool(v) and str(v).strip().lower() not in (
        '', 'n/a', 'na', 'not provided', 'to follow', 'tbd', '-', '/')


def audit(folder, ocr=False):
    files = sorted(glob.glob(os.path.join(folder, '**', '*.xlsx'), recursive=True))
    if not files:
        print(f'No .xlsx files under {folder!r}')
        return

    agg = {
        'types': collections.Counter(),
        'findings': collections.Counter(),     # (severity, category)
        'unknown_alloys': collections.Counter(),
        'etchants': collections.Counter(),
        'coating_types': collections.Counter(),
        'components': collections.Counter(),
        'customers': collections.Counter(),
        'machines': collections.Counter(),
        'elements': collections.Counter(),
        'anomalies': collections.defaultdict(list),
        'errors': [],
    }

    for path in files:
        name = os.path.basename(path)
        try:
            with open(path, 'rb') as f:
                rtype, parsed, findings = review_report(path, f.read(), ocr=ocr)
        except Exception as e:
            agg['errors'].append((name, f'{type(e).__name__}: {e}'))
            print(f'  ✗ {name[:60]:62s} ERROR {type(e).__name__}')
            continue

        agg['types'][rtype] += 1
        for sev, cat, _ in findings:
            agg['findings'][(sev, cat)] += 1
        crit = sum(1 for s, _, _ in findings if s == 'critical')
        warn = sum(1 for s, _, _ in findings if s == 'warning')
        print(f'  {name[:60]:62s} {rtype:13s} 🔴{crit} 🟠{warn}')

        if rtype == 'unknown':
            agg['anomalies']['unclassified layout'].append(name)
            continue

        if rtype == 'metallurgical':
            hdr, smp = parsed.get('header', {}), parsed.get('sample', {})
            mat = smp.get('material')
            if _present(mat):
                if _alloy_key(mat) not in HARDNESS_REF:
                    agg['unknown_alloys'][mat] += 1
            else:
                agg['anomalies']['no material'].append(name)
            if not parsed.get('nominal') or not parsed.get('actual'):
                agg['anomalies']['no/partial composition'].append(name)
            for el in set(parsed.get('nominal', {})) | set(parsed.get('actual', {})):
                agg['elements'][el] += 1
            if len((parsed.get('comment') or '').strip()) < 40:
                agg['anomalies']['short/empty comment'].append(name)
            if not parsed.get('hardness'):
                agg['anomalies']['no hardness section'].append(name)
            so = parsed.get('signoff', {})
            if not all(_present(so.get(k)) for k in ('met_lab', 'mat_eng', 'date')):
                agg['anomalies']['incomplete sign-off'].append(name)
            for _, cap in parsed.get('pictures', []):
                et = caption_etchant(cap or '')
                if et:
                    agg['etchants'][et] += 1
            for k in ('type', 'received', 'outgoing'):
                for ct in _coating_types_in((parsed.get('coating') or {}).get(k)):
                    agg['coating_types'][ct] += 1
            if _present(smp.get('description')):
                agg['components'][smp['description'].strip()] += 1
            if _present(hdr.get('customer')):
                agg['customers'][hdr['customer'].strip()] += 1
            if _present(hdr.get('machine')):
                agg['machines'][hdr['machine'].strip()] += 1

    _report(len(files), agg)


def _section(title, counter, limit=40):
    if not counter:
        return
    print(f'\n{title} ({len(counter)}):')
    for k, n in counter.most_common(limit):
        print(f'  {n:3d}  {k}')


def _report(total, agg):
    print('\n' + '=' * 70)
    print(f'AUDITED {total} report(s):', dict(agg['types']))

    print('\nFindings by severity/category:')
    for (sev, cat), n in sorted(agg['findings'].items(), key=lambda x: (-x[1], x[0])):
        print(f'  {n:4d}  [{sev:8s}] {cat}')

    _section('⚠ Alloys with NO hardness reference (add to HARDNESS_REF)', agg['unknown_alloys'])
    _section('Etchants seen (extend _ETCHANT_VOCAB if any are missing)', agg['etchants'])
    _section('Coating types seen', agg['coating_types'])
    _section('Components seen', agg['components'])
    _section('Customers seen', agg['customers'])
    _section('Machine types seen', agg['machines'])
    _section('Composition elements seen', agg['elements'])

    if agg['anomalies']:
        print('\n⚠ PARSER GAPS / ANOMALIES (investigate these):')
        for kind, names in sorted(agg['anomalies'].items()):
            print(f'  {kind} — {len(names)} file(s):')
            for nm in names[:8]:
                print(f'      {nm}')
            if len(names) > 8:
                print(f'      … +{len(names) - 8} more')

    if agg['errors']:
        print(f'\n✗ HARD ERRORS — {len(agg["errors"])} file(s):')
        for nm, err in agg['errors']:
            print(f'  {nm}: {err}')


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    audit(args[0] if args else '.', ocr='--ocr' in sys.argv)
