import unittest
from email import message_from_bytes
from unittest.mock import patch

import share


class RecipientTests(unittest.TestCase):
    """A mistyped address must be visible, not silently dropped."""

    def test_addresses_are_split_deduplicated_and_checked(self):
        valid, rejected = share.parse_recipients(
            "a@aeg.com, b@aeg.com; A@AEG.COM\n c@aeg.com  notanaddress b@")

        self.assertEqual(valid, ["a@aeg.com", "b@aeg.com", "c@aeg.com"])
        self.assertEqual(rejected, ["notanaddress", "b@"])

    def test_nothing_typed_is_nothing_sent(self):
        self.assertEqual(share.parse_recipients(""), ([], []))
        self.assertEqual(share.parse_recipients(None), ([], []))


class _FakeSMTP:
    """Enough of smtplib.SMTP to see what would have gone out."""

    instances = []

    def __init__(self, host, port, timeout=None):
        self.host, self.port, self.timeout = host, port, timeout
        self.started_tls = False
        self.login_args = None
        self.sent = []
        _FakeSMTP.instances.append(self)

    def __enter__(self):
        return self

    def __exit__(self, *_exc):
        return False

    def ehlo(self):
        pass

    def starttls(self):
        self.started_tls = True

    def login(self, user, password):
        self.login_args = (user, password)

    def send_message(self, message, to_addrs=None):
        self.sent.append((message, to_addrs))


_CONFIG = {
    'smtp_host': 'smtp.aeg.test',
    'smtp_from': 'lab@aeg.test',
    'smtp_user': 'lab@aeg.test',
    'smtp_password': 'secret',
}


class EmailTests(unittest.TestCase):
    def setUp(self):
        _FakeSMTP.instances = []

    def test_not_configured_reports_it_instead_of_pretending_to_send(self):
        with patch.object(share, '_secret', return_value=None):
            self.assertFalse(share.email_configured())
            with self.assertRaises(RuntimeError):
                share.send_email(['a@aeg.test'], 'subject', 'body')

    def test_the_annotated_pdf_and_the_comments_both_travel(self):
        with patch.object(share, '_secret', side_effect=_CONFIG.get), \
             patch.object(share, 'smtplib') as smtplib_module:
            smtplib_module.SMTP = _FakeSMTP
            sent_to = share.send_email(
                ['first@aeg.test', 'second@aeg.test'],
                'Lab report review — 6943',
                'R1: confirm which bucket was sectioned.',
                attachments=[('6943_annotated.pdf', b'%PDF-1.4 fake',
                              'application/pdf')])

        self.assertEqual(sent_to, ['first@aeg.test', 'second@aeg.test'])
        server = _FakeSMTP.instances[0]
        self.assertEqual((server.host, server.port), ('smtp.aeg.test', 587))
        self.assertTrue(server.started_tls)
        self.assertEqual(server.login_args, ('lab@aeg.test', 'secret'))

        message, to_addrs = server.sent[0]
        self.assertEqual(to_addrs, ['first@aeg.test', 'second@aeg.test'])
        self.assertEqual(message['From'], 'lab@aeg.test')
        self.assertEqual(message['Subject'], 'Lab report review — 6943')
        # Re-parse it the way a mail client would.
        parsed = message_from_bytes(message.as_bytes())
        parts = list(parsed.walk())
        body = next(part for part in parts
                    if part.get_content_type() == 'text/plain')
        self.assertIn(b'confirm which bucket was sectioned',
                      body.get_payload(decode=True))
        pdf = next(part for part in parts
                   if part.get_content_type() == 'application/pdf')
        self.assertEqual(pdf.get_filename(), '6943_annotated.pdf')
        self.assertEqual(pdf.get_payload(decode=True), b'%PDF-1.4 fake')

    def test_sending_to_nobody_is_refused(self):
        with patch.object(share, '_secret', side_effect=_CONFIG.get):
            with self.assertRaises(ValueError):
                share.send_email([], 'subject', 'body')

    def test_implicit_tls_uses_the_smtps_port_and_no_starttls(self):
        config = dict(_CONFIG, smtp_ssl='true')
        with patch.object(share, '_secret', side_effect=config.get), \
             patch.object(share, 'smtplib') as smtplib_module:
            smtplib_module.SMTP_SSL = _FakeSMTP
            share.send_email(['first@aeg.test'], 'subject', 'body')

        server = _FakeSMTP.instances[0]
        self.assertEqual(server.port, 465)
        self.assertFalse(server.started_tls)


class DriveShareTests(unittest.TestCase):
    def test_each_recipient_is_granted_read_access_by_name(self):
        calls = []

        class FakeDriveStore:
            @staticmethod
            def is_configured():
                return True

            @staticmethod
            def upload_file(name, data, mimetype, folder=None):
                calls.append(('upload', name, mimetype, folder))
                return {'id': 'file-1', 'link': 'https://drive.test/file-1'}

            @staticmethod
            def grant_reader(file_id, email, notify=True, message=None):
                if email == 'outside@example.test':
                    raise RuntimeError('not a domain user')
                calls.append(('grant', file_id, email, notify))
                return 'perm-1'

        with patch.dict('sys.modules', {'drive_store': FakeDriveStore}):
            result = share.share_via_drive(
                '6943_annotated.pdf', b'%PDF fake',
                ['first@aeg.test', 'outside@example.test'],
                message='Please review.')

        self.assertEqual(result['link'], 'https://drive.test/file-1')
        self.assertEqual(result['granted'], ['first@aeg.test'])
        self.assertEqual(len(result['failed']), 1)
        self.assertEqual(result['failed'][0][0], 'outside@example.test')
        self.assertIn(('upload', '6943_annotated.pdf', 'application/pdf',
                       share.REVIEW_FOLDER), calls)


if __name__ == '__main__':
    unittest.main()
