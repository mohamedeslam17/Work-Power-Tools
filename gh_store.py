#!/usr/bin/env python3
"""
GitHub-backed persistent storage for the photo library.

Commits micrographs into the repo via the GitHub Contents API using a personal
access token. This is the simplest no-IT, browser-only option and survives
Streamlit Community Cloud reboots (the repo is the source of truth). Reads go
through the API, so additions appear immediately without waiting for a redeploy.

Configure via Streamlit secrets or environment variables:

    github_token   / GITHUB_TOKEN     fine-grained PAT (Contents: read & write)
    github_repo    / GITHUB_REPO      "owner/name"
    github_branch  / GITHUB_BRANCH    branch to store files on (default: main)
    github_base    / GITHUB_BASE      path prefix in the repo (default: photo_library)

Falls back to local storage when not configured.
"""
import os
import json
import base64

API = 'https://api.github.com'
_INDEX_NAME = 'index.json'


def _secret(key):
    try:
        import streamlit as st
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.environ.get(key.upper())


def repo():
    return _secret('github_repo')


def branch():
    return _secret('github_branch') or 'main'


def base():
    return (_secret('github_base') or 'photo_library').strip('/')


def is_configured():
    if not (_secret('github_token') and repo()):
        return False
    try:
        import requests  # noqa: F401
        return True
    except Exception:
        return False


def _headers(raw=False):
    return {
        'Authorization': f"Bearer {_secret('github_token')}",
        'X-GitHub-Api-Version': '2022-11-28',
        'Accept': 'application/vnd.github.raw+json' if raw else 'application/vnd.github+json',
    }


def _get(path, raw=False):
    import requests
    r = requests.get(f"{API}/repos/{repo()}/contents/{path}",
                     headers=_headers(raw=raw), params={'ref': branch()}, timeout=30)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r


def _put(path, content_bytes, message, sha=None):
    import requests
    body = {'message': message, 'branch': branch(),
            'content': base64.b64encode(content_bytes).decode()}
    if sha:
        body['sha'] = sha
    r = requests.put(f"{API}/repos/{repo()}/contents/{path}",
                     headers=_headers(), json=body, timeout=60)
    r.raise_for_status()
    return r.json()


def _read_index():
    r = _get(f"{base()}/{_INDEX_NAME}")
    if r is None:
        return None, []                     # index absent — safe to create a new one
    j = r.json()
    try:
        return j.get('sha'), json.loads(base64.b64decode(j['content']).decode('utf-8'))
    except Exception as e:
        # The index exists but couldn't be read/parsed. Do NOT return an empty
        # list — add_records would then commit it over the real index and wipe
        # the whole library. Refuse instead.
        raise RuntimeError(
            f"photo-library index {base()}/{_INDEX_NAME} exists but is unreadable "
            f"({type(e).__name__}: {e}); refusing to overwrite it and lose the library.")


def add_records(records):
    """Commit micrograph records (each with 'bytes') to the repo; return count added."""
    import requests
    from photo_lib import _safe
    sha, index = _read_index()
    existing = {(r.get('job'), r.get('image'), r.get('source')) for r in index}
    added = 0
    for r in records:
        key = (r.get('job'), r.get('image'), r.get('source'))
        if key in existing:
            continue
        name = f"{_safe(r.get('job', ''))}_{_safe(os.path.splitext(r['image'])[0])}.jpg"
        rel = f"{_safe(r['alloy'])}/{name}"
        try:
            _put(f"{base()}/{rel}", r['bytes'], f"library: add {rel}")
        except requests.HTTPError as e:
            code = getattr(e.response, 'status_code', None)
            if code == 422:                 # already in the repo — record it in the index
                pass
            elif code in (401, 403):
                raise RuntimeError('GitHub rejected the write (check the github_token '
                                   'has Contents: read & write on the repo).') from e
            else:                           # 5xx / rate-limit / network — surface, don't drop
                raise
        rec = {k: v for k, v in r.items() if k != 'bytes'}
        rec['path'] = rel
        index.append(rec)
        existing.add(key)
        added += 1
    if added:
        # Not wrapped: if this fails the images are committed but the index isn't,
        # and re-running recovers (existing files 422 → skipped, index re-committed).
        _put(f"{base()}/{_INDEX_NAME}",
             json.dumps(index, indent=2, ensure_ascii=False).encode('utf-8'),
             f"library: index +{added}", sha=sha)
    return added


def load_index():
    return _read_index()[1]


def download(rel_path):
    """Fetch one stored image's bytes by its index 'path' (relative to base())."""
    r = _get(f"{base()}/{rel_path}", raw=True)
    return r.content if r is not None else None
