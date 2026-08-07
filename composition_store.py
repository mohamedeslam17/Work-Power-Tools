#!/usr/bin/env python3
"""
Cross-SESSION duplicate-composition detection for the Lab Report Reviewer.

lab_review.find_duplicate_compositions() can only compare reports uploaded
together in the same call — it has nothing to check against once that call
ends. This module gives it memory: every reviewed metallurgical report's
"fingerprint" (job number, alloy, filename, Actual composition) is stored
persistently, so a report reviewed today can be caught copying one reviewed
weeks ago, without ever uploading them side by side.

Storage is pluggable and auto-selected exactly like the photo library
(`photo_lib.py`), reusing its two backends' underlying auth/transport code
but writing to their own path/folder so the two indexes never collide:

    GitHub  (gh_store)    — commits composition_index/fingerprints.json
    Google  (drive_store) — a separate "AEG Composition Index" Drive folder
    local   (this module) — composition_store/fingerprints.json (default)

Public entry points:
    check(filename, parsed)             -> findings vs. stored history
    remember(filename, parsed)          -> store this report's fingerprint
    check_and_record(filename, parsed)  -> both, in that order (best-effort)
"""
import os
import json

import gh_store
import drive_store
from lab_review import _txt

STORE_DIR = os.environ.get('COMPOSITION_STORE_DIR', 'composition_store')
INDEX_NAME = 'fingerprints.json'
MIN_COMMON_ELEMENTS = 5      # same floor as find_duplicate_compositions

_GH_BASE = 'composition_index'          # separate from gh_store.base() (photo library)
_DRIVE_ROOT = 'AEG Composition Index'   # separate from drive_store._ROOT_NAME (photos)


# ── backend selection (mirrors photo_lib.py) ───────────────────────────────
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
    return f"local ({STORE_DIR}/)"


# ── local backend ────────────────────────────────────────────────────────
def _local_path(store_dir):
    return os.path.join(store_dir, INDEX_NAME)


def _load_local(store_dir):
    p = _local_path(store_dir)
    if os.path.exists(p):
        try:
            with open(p, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return []
    return []


def _save_local(index, store_dir):
    os.makedirs(store_dir, exist_ok=True)
    with open(_local_path(store_dir), 'w', encoding='utf-8') as f:
        json.dump(index, f, indent=2, ensure_ascii=False)


# ── GitHub backend — reuses gh_store's generic PAT plumbing, own path ─────
def _gh_index_path():
    return f"{_GH_BASE}/{INDEX_NAME}"


def _gh_load():
    import base64
    r = gh_store._get(_gh_index_path())
    if r is None:
        return None, []
    j = r.json()
    try:
        return j.get('sha'), json.loads(base64.b64decode(j['content']).decode('utf-8'))
    except Exception as e:
        raise RuntimeError(
            f"composition index {_gh_index_path()} exists but is unreadable "
            f"({type(e).__name__}: {e}); refusing to overwrite it and lose the history.")


def _gh_save(index, sha):
    gh_store._put(_gh_index_path(),
                  json.dumps(index, indent=2, ensure_ascii=False).encode('utf-8'),
                  'composition index: update', sha=sha)


# ── Drive backend — reuses drive_store's generic OAuth plumbing, own folder ─
def _drive_load():
    svc = drive_store._service()
    root = drive_store._ensure_folder(svc, _DRIVE_ROOT, None)
    fid = drive_store._find_child(svc, INDEX_NAME, root)
    if not fid:
        return svc, root, None, []
    try:
        raw = svc.files().get_media(fileId=fid).execute()
        return svc, root, fid, json.loads(raw.decode('utf-8'))
    except Exception as e:
        raise RuntimeError(
            f"composition index in Drive exists but is unreadable "
            f"({type(e).__name__}: {e}); refusing to overwrite it and lose the history.")


def _drive_save(svc, root, index, fid):
    import io
    from googleapiclient.http import MediaIoBaseUpload
    buf = io.BytesIO(json.dumps(index, indent=2, ensure_ascii=False).encode('utf-8'))
    media = MediaIoBaseUpload(buf, mimetype='application/json', resumable=False)
    if fid:
        svc.files().update(fileId=fid, media_body=media).execute()
    else:
        svc.files().create(body={'name': INDEX_NAME, 'parents': [root]},
                           media_body=media, fields='id').execute()


# ── unified interface ───────────────────────────────────────────────────────
def load_index(store_dir=STORE_DIR):
    """Every stored fingerprint. Best-effort: an unreachable/misconfigured
    backend returns an empty history rather than breaking the review."""
    try:
        if use_github():
            return _gh_load()[1]
        if use_drive():
            return _drive_load()[3]
        return _load_local(store_dir)
    except Exception:
        return []


def _fingerprint(filename, parsed):
    hdr, smp = (parsed or {}).get('header') or {}, (parsed or {}).get('sample') or {}
    actual = (parsed or {}).get('actual') or {}
    job = _txt(hdr.get('job'))
    if not job or len(actual) < MIN_COMMON_ELEMENTS:
        return None
    return {
        'job': job,
        'alloy': _txt(smp.get('material')),
        'source': os.path.basename(filename or ''),
        'actual': actual,
    }


def _matches(a, b, min_common=MIN_COMMON_ELEMENTS):
    """Same rule as lab_review.find_duplicate_compositions: different job,
    enough matched elements to rule out coincidence, all exactly equal."""
    if a.get('job') and b.get('job') and a['job'] == b['job']:
        return False
    common = set(a.get('actual') or {}) & set(b.get('actual') or {})
    if len(common) < min_common:
        return False
    return all(a['actual'][el] == b['actual'][el] for el in common)


def check(filename, parsed, store_dir=STORE_DIR):
    """Findings for one report against every previously stored fingerprint."""
    fp = _fingerprint(filename, parsed)
    if not fp:
        return []
    findings = []
    for old in load_index(store_dir):
        if _matches(fp, old):
            common = set(fp['actual']) & set(old.get('actual') or {})
            findings.append(('critical', 'Composition',
                             f'Actual composition is identical to a previously reviewed '
                             f'report — job {old.get("job") or "?"} '
                             f'("{old.get("source") or "?"}") — across all {len(common)} '
                             f'matched element(s). Independent lab results are not '
                             f'expected to match exactly; verify this report wasn\'t '
                             f'copied from that older one.'))
    return findings


def remember(filename, parsed, store_dir=STORE_DIR):
    """Store this report's fingerprint. Idempotent per job number. Returns
    True if a new fingerprint was written."""
    fp = _fingerprint(filename, parsed)
    if not fp:
        return False
    if use_github():
        sha, index = _gh_load()
        if any(r.get('job') == fp['job'] for r in index):
            return False
        index.append(fp)
        _gh_save(index, sha)
        return True
    if use_drive():
        svc, root, fid, index = _drive_load()
        if any(r.get('job') == fp['job'] for r in index):
            return False
        index.append(fp)
        _drive_save(svc, root, index, fid)
        return True
    index = _load_local(store_dir)
    if any(r.get('job') == fp['job'] for r in index):
        return False
    index.append(fp)
    _save_local(index, store_dir)
    return True


def check_and_record(filename, parsed, store_dir=STORE_DIR):
    """check() then remember(), in that order, so a report is checked against
    history BEFORE it becomes part of that history itself. Recording is
    best-effort: a storage failure here must never take down the review."""
    findings = check(filename, parsed, store_dir)
    try:
        remember(filename, parsed, store_dir)
    except Exception:
        pass
    return findings


# ── CLI: audit one or more reports against (and into) the stored history ──
def _main():
    import sys
    from lab_review import review_report
    if len(sys.argv) < 2:
        print('usage: python3 composition_store.py report.xlsx [more.xlsx ...]')
        sys.exit(1)
    print(f'Backend: {backend_name()}\n')
    for path in sys.argv[1:]:
        with open(path, 'rb') as f:
            data = f.read()
        _, parsed, _ = review_report(path, data, ocr=False)
        findings = check_and_record(path, parsed)
        tag = '🔴' if findings else '🟢'
        print(f'{tag} {os.path.basename(path)}')
        for _sev, _cat, msg in findings:
            print(f'    {msg}')


if __name__ == '__main__':
    _main()
