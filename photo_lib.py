#!/usr/bin/env python3
"""
Per-alloy micrograph library for the AEG lab tools.

Extracts the embedded micrographs from a reviewed report and stores them, with
metadata, under a per-alloy structure plus an index. Storage is pluggable and
auto-selected by what's configured:

    GitHub   (gh_store)     — commits to the repo via a PAT  (recommended; no IT)
    Google   (drive_store) — uploads to your Drive via OAuth (alternative)
    local    (this module) — a folder on disk (default / fallback)

The index drives the in-app gallery: pick an alloy → its micrographs and the
data of the report/set they came from. On Streamlit Community Cloud the local
filesystem is wiped on reboot, so configure a cloud backend to persist runtime
additions (the committed seed library survives via git regardless).
"""
import io
import os
import re
import json
import zipfile
import datetime
from collections import Counter

import gh_store
import drive_store

LIBRARY_DIR = os.environ.get('PHOTO_LIBRARY_DIR', 'photo_library')
INDEX_NAME = 'index.json'


def _safe(s):
    return re.sub(r'[^A-Za-z0-9._-]+', '_', (s or 'Unknown')).strip('_') or 'Unknown'


# ── backend selection ─────────────────────────────────────────────────────
def use_github():
    try:
        return gh_store.is_configured()
    except Exception:
        return False


def use_drive():
    try:
        return drive_store.is_configured()
    except Exception:
        return False


def backend_name():
    if use_github():
        return f"GitHub ({gh_store.repo()})"
    if use_drive():
        return "Google Drive"
    return f"local ({LIBRARY_DIR}/)"


# ── shared extraction ─────────────────────────────────────────────────────
def _report_meta(filename, parsed, rtype):
    if rtype == 'metallurgical':
        hdr, smp = parsed.get('header', {}), parsed.get('sample', {})
        alloy = (smp.get('material') or 'Unknown').strip()
        job = re.sub(r'\D', '', hdr.get('job') or '')
        meta = {'customer': hdr.get('customer'), 'machine': hdr.get('machine'),
                'component': smp.get('description'), 'serial': smp.get('serial')}
    else:
        alloy = 'Coating'
        m = re.search(r'\d{4}', parsed.get('report_no') or '')
        job = m.group() if m else ''
        meta = {'component': parsed.get('component'), 'title': parsed.get('title')}
    meta['source'] = os.path.basename(filename or '')
    return alloy, job, meta


def _raw_image_bytes(data):
    out = {}
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except Exception:
        return out
    for n in z.namelist():
        if n.startswith('xl/media'):
            out[n.split('/')[-1]] = z.read(n)
    return out


def _records(filename, data, parsed, rtype):
    """Per-micrograph records (bytes + metadata), shared by every backend."""
    images = parsed.get('images')
    if not images:
        from lab_review import analyze_images
        images, _ = analyze_images(data)
    if not images:
        return []
    alloy, job, meta = _report_meta(filename, parsed, rtype)
    raw = _raw_image_bytes(data)
    from lab_review import report_etchants, image_etchant
    by_mag, primary = report_etchants(parsed.get('pictures'))
    recs = []
    for im in images:
        name = im.get('image')
        if not name or name not in raw:
            continue
        recs.append({
            'alloy': alloy, 'job': job, 'image': name, 'bytes': raw[name],
            'mag': im.get('mag'), 'scale': im.get('scale'),
            'etched': im.get('etched'), 'etchant': image_etchant(im.get('mag'), by_mag, primary),
            'measurements': im.get('measurements', []),
            'added': datetime.date.today().isoformat(), **meta,
        })
    return recs


# ── local backend ─────────────────────────────────────────────────────────
def _index_path(library_dir):
    return os.path.join(library_dir, INDEX_NAME)


def _load_local_index(library_dir=LIBRARY_DIR):
    p = _index_path(library_dir)
    if os.path.exists(p):
        try:
            with open(p, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return []
    return []


def _save_local_index(index, library_dir=LIBRARY_DIR):
    os.makedirs(library_dir, exist_ok=True)
    with open(_index_path(library_dir), 'w', encoding='utf-8') as f:
        json.dump(index, f, indent=2, ensure_ascii=False)


def _add_local(records, library_dir):
    from PIL import Image
    index = _load_local_index(library_dir)
    existing = {(r.get('job'), r.get('image'), r.get('source')) for r in index}
    added = 0
    for r in records:
        key = (r.get('job'), r.get('image'), r.get('source'))
        if key in existing:
            continue
        adir = os.path.join(library_dir, _safe(r['alloy']))
        os.makedirs(adir, exist_ok=True)
        name = f"{_safe(r.get('job', ''))}_{_safe(os.path.splitext(r['image'])[0])}.jpg"
        out_path = os.path.join(adir, name)
        try:
            Image.open(io.BytesIO(r['bytes'])).convert('RGB').save(out_path, 'JPEG', quality=85)
        except Exception:
            continue
        rec = {k: v for k, v in r.items() if k != 'bytes'}
        rec['path'] = os.path.relpath(out_path, library_dir).replace(os.sep, '/')
        index.append(rec)
        existing.add(key)
        added += 1
    if added:
        _save_local_index(index, library_dir)
    return added


# ── unified interface ─────────────────────────────────────────────────────
def add_to_library(filename, data, parsed, rtype, library_dir=LIBRARY_DIR):
    recs = _records(filename, data, parsed, rtype)
    if not recs:
        return 0
    if use_github():
        return gh_store.add_records(recs)
    if use_drive():
        return drive_store.add_records(recs)
    return _add_local(recs, library_dir)


def _index(library_dir=LIBRARY_DIR):
    if use_github():
        return gh_store.load_index()
    if use_drive():
        return drive_store.load_index()
    return _load_local_index(library_dir)


def alloy_counts(library_dir=LIBRARY_DIR):
    return Counter(r.get('alloy', 'Unknown') for r in _index(library_dir))


def photos_for(alloy, library_dir=LIBRARY_DIR):
    return [r for r in _index(library_dir) if r.get('alloy') == alloy]


def get_image_bytes(entry, library_dir=LIBRARY_DIR):
    if use_github():
        return gh_store.download(entry.get('path'))
    if use_drive():
        return drive_store.download(entry.get('drive_id'))
    p = os.path.join(library_dir, entry.get('path', ''))
    if os.path.exists(p):
        with open(p, 'rb') as f:
            return f.read()
    return None


# ── CLI: populate the (local) library from one or more reports ────────────
def _main():
    import sys
    from lab_review import review_report
    if len(sys.argv) < 2:
        print('usage: python3 photo_lib.py report.xlsx [more.xlsx ...]')
        sys.exit(1)
    total = 0
    for path in sys.argv[1:]:
        with open(path, 'rb') as f:
            data = f.read()
        rtype, parsed, _ = review_report(path, data, ocr=True)
        n = add_to_library(path, data, parsed, rtype)
        print(f'  +{n:2d}  {os.path.basename(path)}  ({rtype})')
        total += n
    print(f'\nAdded {total} micrograph(s) to {backend_name()}. Library holds:')
    for alloy, c in sorted(alloy_counts().items()):
        print(f'  {alloy}: {c}')


if __name__ == '__main__':
    _main()
