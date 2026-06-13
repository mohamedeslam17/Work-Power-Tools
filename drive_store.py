#!/usr/bin/env python3
"""
Google Drive backend for the photo library — persistent storage that survives
Streamlit Community Cloud reboots (which wipe the local filesystem).

Auth model: OAuth acting **as you** (not a service account), so micrographs
land in your own Drive and use your quota — the right choice for a personal
Gmail account. Configure via Streamlit secrets or environment variables:

    drive_client_id        / DRIVE_CLIENT_ID
    drive_client_secret    / DRIVE_CLIENT_SECRET
    drive_refresh_token    / DRIVE_REFRESH_TOKEN
    drive_folder_id        / DRIVE_FOLDER_ID   (the library's parent folder)

Get a refresh token once with:   python3 drive_store.py --auth
Push the existing local library:  python3 drive_store.py --migrate

When not configured (or google libraries absent) every entry point reports
"not configured" and photo_lib falls back to local storage.
"""
import io
import os
import sys
import json

_SCOPES = ['https://www.googleapis.com/auth/drive.file']
_TOKEN_URI = 'https://oauth2.googleapis.com/token'
_INDEX_NAME = 'index.json'
_FOLDER_MIME = 'application/vnd.google-apps.folder'

_service_cache = None


# ── configuration ─────────────────────────────────────────────────────────
def _secret(key):
    """Look up a value in Streamlit secrets, then the environment."""
    try:
        import streamlit as st
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.environ.get(key.upper())


def folder_id():
    return _secret('drive_folder_id') or None


def _oauth_conf():
    cid = _secret('drive_client_id')
    secret = _secret('drive_client_secret')
    refresh = _secret('drive_refresh_token')
    if cid and secret and refresh:
        return cid, secret, refresh
    return None


def is_configured():
    """True only when creds + folder id are present and google libs import."""
    if not (folder_id() and _oauth_conf()):
        return False
    try:
        import google.oauth2.credentials          # noqa: F401
        import googleapiclient.discovery           # noqa: F401
        return True
    except Exception:
        return False


def _service():
    global _service_cache
    if _service_cache is not None:
        return _service_cache
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
    files = svc.files().list(q=q, fields='files(id,name)', pageSize=1,
                             spaces='drive').execute().get('files', [])
    return files[0]['id'] if files else None


def _ensure_folder(svc, name, parent):
    fid = _find_child(svc, name, parent, _FOLDER_MIME)
    if fid:
        return fid
    meta = {'name': name, 'mimeType': _FOLDER_MIME, 'parents': [parent]}
    return svc.files().create(body=meta, fields='id').execute()['id']


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


# ── public API (mirrors the local backend in photo_lib) ───────────────────
def add_records(records):
    """Upload micrograph records (each with 'bytes') to Drive; return count added."""
    from googleapiclient.http import MediaIoBaseUpload
    svc, root = _service(), folder_id()
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
    svc, root = _service(), folder_id()
    return _read_index(svc, root)[1]


def download(drive_id):
    return _service().files().get_media(fileId=drive_id).execute()


# ── CLI helpers ───────────────────────────────────────────────────────────
def _auth():
    """One-time browser consent → prints a refresh token for your secrets."""
    from google_auth_oauthlib.flow import InstalledAppFlow
    cid = _secret('drive_client_id') or input('OAuth client_id: ').strip()
    secret = _secret('drive_client_secret') or input('OAuth client_secret: ').strip()
    cfg = {'installed': {'client_id': cid, 'client_secret': secret,
                         'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
                         'token_uri': _TOKEN_URI,
                         'redirect_uris': ['http://localhost']}}
    flow = InstalledAppFlow.from_client_config(cfg, _SCOPES)
    creds = flow.run_local_server(port=0)
    print('\nAdd this to your secrets:')
    print(f'  drive_refresh_token = "{creds.refresh_token}"')


def _migrate(library_dir='photo_library'):
    """Push the existing local library up to Drive."""
    import photo_lib
    index = photo_lib._load_local_index(library_dir)
    recs = []
    for r in index:
        p = os.path.join(library_dir, r.get('path', ''))
        if not os.path.exists(p):
            continue
        with open(p, 'rb') as f:
            recs.append({**r, 'bytes': f.read()})
    print(f'Uploading {len(recs)} micrograph(s) to Drive…')
    print(f'Done — added {add_records(recs)}.')


if __name__ == '__main__':
    if '--auth' in sys.argv:
        _auth()
    elif '--migrate' in sys.argv:
        if not is_configured():
            sys.exit('Drive not configured (need client id/secret/refresh token + folder id).')
        _migrate()
    else:
        print(__doc__)
