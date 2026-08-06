#!/usr/bin/env python3
"""
Sending a reviewed report to other people.

Two independent transports, both optional, both configured the way the rest of
the app is (Streamlit secrets first, environment variables second):

  * **E-mail** — stdlib `smtplib`, so this adds no dependency. The annotated
    PDF is attached and the review comments are also written into the message
    body, so the mail says something useful before anyone opens the attachment.

        smtp_host / SMTP_HOST                 e.g. smtp.office365.com
        smtp_port / SMTP_PORT                 default 465 with SSL, else 587
        smtp_user / SMTP_USER                 login (optional on an open relay)
        smtp_password / SMTP_PASSWORD
        smtp_from / SMTP_FROM                 defaults to smtp_user
        smtp_ssl / SMTP_SSL                   implicit TLS ("smtps"), default off
        smtp_starttls / SMTP_STARTTLS         default on when not using SSL

  * **Google Drive** — reuses `drive_store`'s existing OAuth credentials,
    uploads the PDF and grants each named recipient read access, returning a
    link. Needs no extra configuration beyond the Drive keys the photo library
    already uses.

Nothing here knows about Streamlit widgets: it takes addresses and bytes and
reports what happened, so the caller decides what to show. Neither transport is
required — with neither configured the reviewer downloads the PDF and sends it
themselves, which is why every function here is guarded by an `is_configured`
style check rather than raising at import.
"""
import os
import re
import smtplib
from email.message import EmailMessage

_ADDRESS = re.compile(r'^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]{2,}$')
_SPLIT = re.compile(r'[,;\s]+')

REVIEW_FOLDER = 'AEG Report Reviews'


def _secret(key):
    try:
        import streamlit as st
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.environ.get(key.upper())


def _flag(key, default=False):
    raw = _secret(key)
    if raw is None or str(raw).strip() == '':
        return default
    return str(raw).strip().lower() in ('1', 'true', 'yes', 'on')


# ── recipients ────────────────────────────────────────────────────────────
def parse_recipients(text):
    """Split a typed recipient list into (valid, rejected).

    Deliberately returns the rejects instead of dropping them: a mistyped
    address in a list of five must be visible, or the reviewer believes they
    sent the report to someone who never received it.
    """
    seen, valid, rejected = set(), [], []
    for token in _SPLIT.split(str(text or '')):
        token = token.strip().strip('<>')
        if not token:
            continue
        key = token.lower()
        if key in seen:
            continue
        seen.add(key)
        (valid if _ADDRESS.match(token) else rejected).append(token)
    return valid, rejected


# ── e-mail ────────────────────────────────────────────────────────────────
def email_config():
    """The SMTP settings, or None when e-mail isn't set up here."""
    host = _secret('smtp_host')
    if not host:
        return None
    ssl = _flag('smtp_ssl')
    sender = _secret('smtp_from') or _secret('smtp_user')
    if not sender:
        return None
    try:
        port = int(_secret('smtp_port') or (465 if ssl else 587))
    except (TypeError, ValueError):
        port = 465 if ssl else 587
    return {
        'host': str(host).strip(),
        'port': port,
        'user': _secret('smtp_user'),
        'password': _secret('smtp_password'),
        'sender': str(sender).strip(),
        'ssl': ssl,
        'starttls': _flag('smtp_starttls', not ssl),
        'timeout': 30,
    }


def email_configured():
    return email_config() is not None


def email_sender():
    config = email_config()
    return config['sender'] if config else None


def send_email(recipients, subject, body, attachments=(), reply_to=None, cc=()):
    """Send one message with attachments. Returns the addresses it went to.

    `attachments` is [(filename, bytes, mimetype)]. Raises on refusal or on a
    transport failure — the caller shows the reason rather than reporting a
    send that did not happen.
    """
    config = email_config()
    if not config:
        raise RuntimeError(
            'E-mail is not configured — set smtp_host and smtp_from '
            '(see share.py) or download the PDF and send it yourself.')
    recipients = [address for address in recipients if address]
    if not recipients:
        raise ValueError('No recipients given.')

    message = EmailMessage()
    message['From'] = config['sender']
    message['To'] = ', '.join(recipients)
    if cc:
        message['Cc'] = ', '.join(cc)
    message['Subject'] = subject or 'Lab report review'
    if reply_to:
        message['Reply-To'] = reply_to
    message.set_content(body or '')
    for name, data, mimetype in attachments or ():
        maintype, _, subtype = str(mimetype or 'application/octet-stream').partition('/')
        message.add_attachment(data, maintype=maintype,
                               subtype=subtype or 'octet-stream', filename=name)

    if config['ssl']:
        server = smtplib.SMTP_SSL(config['host'], config['port'],
                                  timeout=config['timeout'])
    else:
        server = smtplib.SMTP(config['host'], config['port'],
                              timeout=config['timeout'])
    with server:
        server.ehlo()
        if config['starttls'] and not config['ssl']:
            server.starttls()
            server.ehlo()
        if config['user'] and config['password']:
            server.login(config['user'], config['password'])
        server.send_message(message, to_addrs=recipients + list(cc or ()))
    return recipients + list(cc or ())


# ── Google Drive ──────────────────────────────────────────────────────────
def drive_configured():
    try:
        import drive_store
        return drive_store.is_configured()
    except Exception:
        return False


def share_via_drive(filename, data, recipients=(), mimetype='application/pdf',
                    notify=True, message=None):
    """Upload the annotated report to Drive and give recipients read access.

    Returns {'link', 'granted', 'failed'}. Access is granted per named person —
    never a public "anyone with the link" share, because these are customer
    documents. A recipient the API refuses lands in 'failed' with its reason
    instead of failing the whole send.
    """
    import drive_store
    if not drive_store.is_configured():
        raise RuntimeError('Google Drive is not configured for this deployment.')
    uploaded = drive_store.upload_file(filename, data, mimetype,
                                       folder=REVIEW_FOLDER)
    granted, failed = [], []
    for address in recipients or ():
        try:
            drive_store.grant_reader(uploaded['id'], address, notify=notify,
                                     message=message)
            granted.append(address)
        except Exception as error:
            failed.append((address, f'{type(error).__name__}: {error}'))
    return {'link': uploaded.get('link'), 'id': uploaded['id'],
            'granted': granted, 'failed': failed}


def transports():
    """Which ways of sending are available here, for the UI to offer."""
    return {'email': email_configured(), 'drive': drive_configured()}


if __name__ == '__main__':
    print(__doc__)
    print('e-mail configured:', email_configured(), '| from:', email_sender())
    print('drive configured :', drive_configured())
