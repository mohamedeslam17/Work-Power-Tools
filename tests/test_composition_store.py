import shutil
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import composition_store as cs
import drive_store


class CompositionStoreTests(unittest.TestCase):
    def setUp(self):
        self.store_dir = tempfile.mkdtemp(prefix='composition_store_test_')
        # Force the local backend regardless of the environment's secrets.
        self._gh = patch.object(cs, 'use_github', return_value=False).start()
        self._drive = patch.object(cs, 'use_drive', return_value=False).start()
        self.addCleanup(patch.stopall)
        self.addCleanup(shutil.rmtree, self.store_dir, ignore_errors=True)

    def _parsed(self, job, actual, material='GTD 111'):
        return {'header': {'job': job}, 'sample': {'material': material},
                'actual': dict(actual)}

    def test_second_session_catches_a_copy_from_the_first(self):
        actual = {'Ni': 59.53, 'Cr': 13.57, 'Co': 9.14, 'Mo': 1.46,
                  'W': 4.73, 'Al': 2.49, 'Ti': 4.81, 'Ta': 3.05, 'Fe': 0.08}

        first = cs.check_and_record('6630.xlsx', self._parsed('6630', actual),
                                     store_dir=self.store_dir)
        self.assertEqual(first, [])   # nothing in history yet

        second = cs.check_and_record(
            '6991.xlsx', self._parsed('6991', dict(actual, Cu=0.07)),
            store_dir=self.store_dir)

        self.assertEqual(len(second), 1)
        severity, category, message = second[0]
        self.assertEqual(severity, 'critical')
        self.assertEqual(category, 'Composition')
        self.assertIn('6630', message)
        self.assertIn('9 matched element', message)

    def test_reviewing_the_same_job_twice_does_not_duplicate_history(self):
        actual = {'Ni': 60.0, 'Cr': 14.0, 'Co': 9.5, 'Mo': 1.5, 'W': 3.8}
        parsed = self._parsed('7000', actual)

        cs.check_and_record('a.xlsx', parsed, store_dir=self.store_dir)
        cs.check_and_record('a.xlsx', parsed, store_dir=self.store_dir)

        self.assertEqual(len(cs.load_index(self.store_dir)), 1)

    def test_different_parts_with_different_chemistry_are_not_flagged(self):
        cs.check_and_record(
            'a.xlsx',
            self._parsed('7000', {'Ni': 60.0, 'Cr': 14.0, 'Co': 9.5, 'Mo': 1.5, 'W': 3.8}),
            store_dir=self.store_dir)

        findings = cs.check_and_record(
            'b.xlsx',
            self._parsed('7001', {'Ni': 61.2, 'Cr': 12.9, 'Co': 9.8, 'Mo': 1.4, 'W': 4.1}),
            store_dir=self.store_dir)

        self.assertEqual(findings, [])

    def test_sparse_composition_is_not_a_coincidence_match(self):
        # Two elements matching by chance shouldn't fire -- MIN_COMMON_ELEMENTS
        # guards against a coincidental match on a thin table.
        cs.check_and_record(
            'a.xlsx', self._parsed('7000', {'Ni': 60.0, 'Cr': 14.0}),
            store_dir=self.store_dir)

        findings = cs.check_and_record(
            'b.xlsx', self._parsed('7001', {'Ni': 60.0, 'Cr': 14.0}),
            store_dir=self.store_dir)

        self.assertEqual(findings, [])

    def test_reports_with_no_job_number_are_not_stored_or_matched(self):
        actual = {'Ni': 60.0, 'Cr': 14.0, 'Co': 9.5, 'Mo': 1.5, 'W': 3.8}
        findings = cs.check_and_record('a.xlsx', self._parsed('', actual),
                                        store_dir=self.store_dir)
        self.assertEqual(findings, [])
        self.assertEqual(cs.load_index(self.store_dir), [])


class _FakeDrive:
    """Minimal in-memory stand-in for the Drive v3 API surface composition_store
    touches (files().create/update/get_media), so the Drive backend can be
    exercised without real OAuth credentials."""

    def __init__(self):
        self._files = {}      # fid -> {'name', 'parent', 'is_folder', 'content'}
        self._next = 0

    def _new_id(self):
        self._next += 1
        return f'id{self._next}'

    def ensure_folder(self, svc, name, parent):
        for fid, f in self._files.items():
            if f['is_folder'] and f['name'] == name and f['parent'] == parent:
                return fid
        fid = self._new_id()
        self._files[fid] = {'name': name, 'parent': parent, 'is_folder': True, 'content': None}
        return fid

    def find_child(self, svc, name, parent, mime=None):
        for fid, f in self._files.items():
            if not f['is_folder'] and f['name'] == name and f['parent'] == parent:
                return fid
        return None

    def service(self):
        drive = self

        class FilesAPI:
            def get_media(self, fileId):
                return MagicMock(execute=lambda: drive._files[fileId]['content'])

            def create(self, body, media_body=None, fields=None):
                fid = drive._new_id()
                content = None
                if media_body is not None:
                    media_body._fd.seek(0)
                    content = media_body._fd.read()
                drive._files[fid] = {'name': body.get('name'),
                                     'parent': (body.get('parents') or [None])[0],
                                     'is_folder': False, 'content': content}
                return MagicMock(execute=lambda: {'id': fid})

            def update(self, fileId, media_body=None):
                media_body._fd.seek(0)
                drive._files[fileId]['content'] = media_body._fd.read()
                return MagicMock(execute=lambda: {})

        class Svc:
            def files(self):
                return FilesAPI()

        return Svc()


class CompositionStoreDriveBackendTests(unittest.TestCase):
    """The Drive backend needs its own coverage: it's structurally different
    code from the local backend (JSON-in-a-file) that the other tests never
    exercise, and it must land in its own Drive folder, never the photo
    library's "AEG Photo Library"."""

    def setUp(self):
        self.drive = _FakeDrive()
        patch.object(cs, 'use_github', return_value=False).start()
        patch.object(cs, 'use_drive', return_value=True).start()
        patch.object(drive_store, '_service', return_value=self.drive.service()).start()
        patch.object(drive_store, '_ensure_folder', side_effect=self.drive.ensure_folder).start()
        patch.object(drive_store, '_find_child', side_effect=self.drive.find_child).start()
        self.addCleanup(patch.stopall)

    def _parsed(self, job, actual, material='GTD 111'):
        return {'header': {'job': job}, 'sample': {'material': material},
                'actual': dict(actual)}

    def test_second_session_catches_a_copy_from_the_first_via_drive(self):
        actual = {'Ni': 59.53, 'Cr': 13.57, 'Co': 9.14, 'Mo': 1.46,
                  'W': 4.73, 'Al': 2.49, 'Ti': 4.81, 'Ta': 3.05, 'Fe': 0.08}

        first = cs.check_and_record('6630.xlsx', self._parsed('6630', actual))
        self.assertEqual(first, [])

        second = cs.check_and_record(
            '6991.xlsx', self._parsed('6991', dict(actual, Cu=0.07)))

        self.assertEqual(len(second), 1)
        self.assertIn('6630', second[0][2])

    def test_uses_its_own_drive_folder_not_the_photo_library_one(self):
        cs.check_and_record(
            'a.xlsx',
            self._parsed('7000', {'Ni': 60.0, 'Cr': 14.0, 'Co': 9.5, 'Mo': 1.5, 'W': 3.8}))

        folder_names = {f['name'] for f in self.drive._files.values() if f['is_folder']}
        self.assertIn('AEG Composition Index', folder_names)
        self.assertNotIn('AEG Photo Library', folder_names)

    def test_reviewing_the_same_job_twice_does_not_duplicate_history_on_drive(self):
        parsed = self._parsed('7000', {'Ni': 60.0, 'Cr': 14.0, 'Co': 9.5, 'Mo': 1.5, 'W': 3.8})

        cs.check_and_record('a.xlsx', parsed)
        cs.check_and_record('a.xlsx', parsed)

        self.assertEqual(len(cs.load_index()), 1)


if __name__ == '__main__':
    unittest.main()
