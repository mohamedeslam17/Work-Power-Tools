#!/usr/bin/env python3
"""
Per-alloy micrograph library for the AEG lab tools.

Extracts the embedded micrographs from a reviewed report and stores them, with
metadata, under a per-alloy folder structure plus a JSON index:

    <library>/index.json
    <library>/GTD-741/7227_image1.jpg
    <library>/Rene-80/6831_image2.jpg
    ...

The index drives the in-app gallery (pick an alloy → see its micrographs and the
data of the report/set they came from). The library lives on disk; in an
ephemeral environment, commit it (or point PHOTO_LIBRARY_DIR elsewhere) to make
it persistent.
"""
import io
import os
import re
import json
import zipfile
import datetime
from collections import Counter

LIBRARY_DIR = os.environ.get('PHOTO_LIBRARY_DIR', 'photo_library')
INDEX_NAME = 'index.json'


def _safe(s):
    return re.sub(r'[^A-Za-z0-9._-]+', '_', (s or 'Unknown')).strip('_') or 'Unknown'


def index_path(library_dir=LIBRARY_DIR):
    return os.path.join(library_dir, INDEX_NAME)


def load_index(library_dir=LIBRARY_DIR):
    p = index_path(library_dir)
    if os.path.exists(p):
        try:
            with open(p, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return []
    return []


def _save_index(index, library_dir=LIBRARY_DIR):
    os.makedirs(library_dir, exist_ok=True)
    with open(index_path(library_dir), 'w', encoding='utf-8') as f:
        json.dump(index, f, indent=2, ensure_ascii=False)


def _report_meta(filename, parsed, rtype):
    """(alloy, job, shared-metadata) for a reviewed report."""
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
    """{image-name: bytes} for embedded media, without any OCR."""
    out = {}
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except Exception:
        return out
    for n in z.namelist():
        if n.startswith('xl/media'):
            out[n.split('/')[-1]] = z.read(n)
    return out


def add_to_library(filename, data, parsed, rtype, library_dir=LIBRARY_DIR):
    """Persist a report's micrographs under <dir>/<alloy>/, updating the index.

    Reuses the per-image analysis already in parsed['images'] (no re-OCR) and
    pulls the raw bytes straight from the workbook. Returns the count added.
    """
    from PIL import Image  # local import keeps lab_review usable without Pillow

    images = parsed.get('images')
    if not images:
        from lab_review import analyze_images
        images, _ = analyze_images(data)
    if not images:
        return 0

    alloy, job, meta = _report_meta(filename, parsed, rtype)
    raw = _raw_image_bytes(data)
    index = load_index(library_dir)
    existing = {(r.get('job'), r.get('image'), r.get('source')) for r in index}

    adir = os.path.join(library_dir, _safe(alloy))
    os.makedirs(adir, exist_ok=True)
    added = 0
    for im in images:
        name = im.get('image')
        key = (job, name, meta['source'])
        if key in existing or name not in raw:
            continue
        out_name = f"{_safe(job)}_{_safe(os.path.splitext(name)[0])}.jpg"
        out_path = os.path.join(adir, out_name)
        try:
            Image.open(io.BytesIO(raw[name])).convert('RGB').save(out_path, 'JPEG', quality=85)
        except Exception:
            continue
        rec = {
            'alloy': alloy, 'job': job, 'image': name,
            'path': os.path.relpath(out_path, library_dir),
            'mag': im.get('mag'), 'scale': im.get('scale'),
            'etched': im.get('etched'), 'measurements': im.get('measurements', []),
            'added': datetime.date.today().isoformat(),
            **meta,
        }
        index.append(rec)
        existing.add(key)
        added += 1

    if added:
        _save_index(index, library_dir)
    return added


def alloy_counts(library_dir=LIBRARY_DIR):
    """{alloy: micrograph-count} across the library."""
    return Counter(r.get('alloy', 'Unknown') for r in load_index(library_dir))


def photos_for(alloy, library_dir=LIBRARY_DIR):
    return [r for r in load_index(library_dir) if r.get('alloy') == alloy]


# ── CLI: populate the library from one or more reports ────────────────────
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
    counts = alloy_counts()
    print(f'\nAdded {total} micrograph(s). Library now holds:')
    for alloy, c in sorted(counts.items()):
        print(f'  {alloy}: {c}')


if __name__ == '__main__':
    _main()
