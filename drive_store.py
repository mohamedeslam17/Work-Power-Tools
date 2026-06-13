#!/usr/bin/env python3
"""
Google Drive backend for the photo library — persistent storage that survives
Streamlit Community Cloud reboots and keeps image binaries out of the git repo.

Auth: OAuth acting **as you** (not a service account), so micrographs land in
your own Drive and use your 15 GB+ quota. Uses the least-privilege
`drive.file` scope and self-manages a folder named "AEG Photo Library" — no
folder ID to hunt down. Configure three values via Streamlit secrets / env:

    drive_client_id      / DRIVE_CLIENT_ID
    drive_client_secret  / DRIVE_CLIENT_SECRET
    drive_refresh_token  / DRIVE_REFRESH_TOKEN

Get the refresh token once (any machine/Colab with Python + a browser):
    python3 drive_store.py --auth
Push the existing local seed library up to Drive:
    python3 drive_store.py --migrate

Not configured (or google libs absent) ⇒ photo_lib falls back to local storage.
"""
import io
import os
import sys
import json

_SCOPES = ['https://www.googleapis.com/auth/drive.file']
_TOKEN_URI = 'https://oauth2.googleapis.com/token'
_ROOT_NAME = 'AEG Photo Library'
_INDEX_NAME = 'index.json'
_FOLDER_MIME = 'application/vnd.google-apps.folder'

_service_cache = None
_root_cache = None


# ── configuration ─────────────────────────────────────────────────────────
def _secret(key):
    try:
        import streamlit as st
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.environ.get(key.upper())


def _oauth_conf():
    cid = _secret('drive_client_id')
    secret = _secret('drive_client_secret')
    refresh = _secret('drive_refresh_token')
    return (cid, secret, refresh) if (cid and secret and refresh) else None


def is_configured():
    if not _oauth_conf():
        return False
    try:
        import google.oauth2.credentials      # noqa: F401
        import googleapiclient.discovery       # noqa: F401
        return True
    except Exception:
        return False


def _service():
    global _service_cache
    if _service_cache is None:
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
        cid, secret, refresh = _oauth_conf()
        creds = Credentials(None, refresh_token=refresh, client_id=cid,
                            client_secret=secret, token_uri=_TOKEN_URI, scopes=_SCOPES)
        _service_cache = build('drive', 'v3', credentials=creds, cache_discovery=False)
    return _service_cache


# ── Drive helpers ─────────────────────────────────────────────────────────
def _esc(s):
    return (s or '').replace('\\', '\\\\').replace("'", "\\'")


def _find_child(svc, name, parent, mime=None):
    q = f"name = '{_esc(name)}' and '{parent}' in parents and trashed = false"
    if mime:
        q += f" and mimeType = '{mime}'"
    files = svc.files().list(q=q, fields='files(id)', pageSize=1, spaces='drive').execute()
    files = files.get('files', [])
    return files[0]['id'] if files else None


def _ensure_folder(svc, name, parent):
    fid = _find_child(svc, name, parent, _FOLDER_MIME)
    if fid:
        return fid
    body = {'name': name, 'mimeType': _FOLDER_MIME}
    if parent:
        body['parents'] = [parent]
    return svc.files().create(body=body, fields='id').execute()['id']


def _root_id(svc):
    """Find-or-create the app-owned 'AEG Photo Library' folder in My Drive."""
    global _root_cache
    if _root_cache is None:
        q = (f"name = '{_ROOT_NAME}' and mimeType = '{_FOLDER_MIME}' and trashed = false")
        files = svc.files().list(q=q, fields='files(id)', pageSize=1, spaces='drive').execute()
        files = files.get('files', [])
        _root_cache = files[0]['id'] if files else _ensure_folder(svc, _ROOT_NAME, None)
    return _root_cache


def _read_index(svc, root):
    fid = _find_child(svc, _INDEX_NAME, root)
    if not fid:
        return None, []
    try:
        raw = svc.files().get_media(fileId=fid).execute()
        return fid, json.loads(raw.decode('utf-8'))
    except Exception:
        return fid, []


def _write_index(svc, root, index, fid=None):
    from googleapiclient.http import MediaIoBaseUpload
    buf = io.BytesIO(json.dumps(index, indent=2, ensure_ascii=False).encode('utf-8'))
    media = MediaIoBaseUpload(buf, mimetype='application/json', resumable=False)
    if fid:
        svc.files().update(fileId=fid, media_body=media).execute()
    else:
        svc.files().create(body={'name': _INDEX_NAME, 'parents': [root]},
                           media_body=media, fields='id').execute()


# ── public API (mirrors gh_store / the local backend) ─────────────────────
def add_records(records):
    """Upload micrograph records (each with 'bytes') to Drive; return count added."""
    from googleapiclient.http import MediaIoBaseUpload
    svc = _service()
    root = _root_id(svc)
    idx_fid, index = _read_index(svc, root)
    existing = {(r.get('job'), r.get('image'), r.get('source')) for r in index}
    folders, added = {}, 0
    for r in records:
        key = (r.get('job'), r.get('image'), r.get('source'))
        if key in existing:
            continue
        alloy = r['alloy']
        if alloy not in folders:
            folders[alloy] = _ensure_folder(svc, alloy, root)
        name = f"{r.get('job', '')}_{os.path.splitext(r['image'])[0]}.jpg"
        media = MediaIoBaseUpload(io.BytesIO(r['bytes']), mimetype='image/jpeg', resumable=False)
        file = svc.files().create(body={'name': name, 'parents': [folders[alloy]]},
                                  media_body=media, fields='id').execute()
        rec = {k: v for k, v in r.items() if k != 'bytes'}
        rec['drive_id'] = file['id']
        rec['name'] = name
        index.append(rec)
        existing.add(key)
        added += 1
    if added:
        _write_index(svc, root, index, idx_fid)
    return added


def load_index():
    svc = _service()
    return _read_index(svc, _root_id(svc))[1]


def download(drive_id):
    if not drive_id:
        return None
    return _service().files().get_media(fileId=drive_id).execute()


# ── CLI helpers ───────────────────────────────────────────────────────────
def _auth():
    """Portable browser auth → prints a refresh token for your secrets."""
    from urllib.parse import urlparse, parse_qs
    from google_auth_oauthlib.flow import InstalledAppFlow
    cid = _secret('drive_client_id') or input('OAuth client_id: ').strip()
    secret = _secret('drive_client_secret') or input('OAuth client_secret: ').strip()
    cfg = {'installed': {'client_id': cid, 'client_secret': secret,
                         'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
                         'token_uri': _TOKEN_URI, 'redirect_uris': ['http://localhost']}}
    flow = InstalledAppFlow.from_client_config(cfg, _SCOPES)
    flow.redirect_uri = 'http://localhost'
    url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    print('\n1) Open this URL, sign in, and approve:\n\n', url)
    print('\n2) Your browser will try to load http://localhost/?code=...  (a connection error is fine).')
    resp = input('\n3) Paste the FULL redirected URL (or just the code):\n> ').strip()
    code = parse_qs(urlparse(resp).query).get('code', [resp])[0] if 'code=' in resp else resp
    flow.fetch_token(code=code)
    print('\nAdd this to your Streamlit secrets:')
    print(f'  drive_refresh_token = "{flow.credentials.refresh_token}"')


def _migrate(library_dir='photo_library'):
    import photo_lib
    recs = []
    for r in photo_lib._load_local_index(library_dir):
        p = os.path.join(library_dir, r.get('path', ''))
        if os.path.exists(p):
            with open(p, 'rb') as f:
                recs.append({**r, 'bytes': f.read()})
    print(f'Uploading {len(recs)} micrograph(s) to Drive…')
    print(f'Done — added {add_records(recs)}.')


if __name__ == '__main__':
    if '--auth' in sys.argv:
        _auth()
    elif '--migrate' in sys.argv:
        if not is_configured():
            sys.exit('Drive not configured (need client id/secret + refresh token).')
        _migrate()
    else:
        print(__doc__)
