"""Storage-safety tests for the photo library and its two cloud backends.

Every case here pins a defect recovered from claude/tool-audit-improvements-v7bj10
(6 Jul 2026), a branch whose fixes were never merged. The originals shipped
without tests, which is most of why they were easy to lose.
"""
import json
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import gh_store
import photo_lib


class StoredNameTests(unittest.TestCase):
    def test_two_reports_sharing_a_job_number_do_not_overwrite_each_other(self):
        # The old name was job + image stem only, so the same job number in two
        # different reports mapped to one file and the second silently replaced
        # the first — a library that loses images it says it has.
        first = {'job': '6943', 'image': 'image1.jpeg', 'alloy': 'GTD 111',
                 'source': 'TBR 6943 … (Final).xlsx'}
        second = dict(first, source='TBR 6943 … (Rev A).xlsx')

        self.assertNotEqual(photo_lib._stored_filename(first),
                            photo_lib._stored_filename(second))
        # Stable for the same report, so re-running an add is idempotent.
        self.assertEqual(photo_lib._stored_filename(first),
                         photo_lib._stored_filename(dict(first)))

    def test_a_workbook_supplied_alloy_cannot_escape_the_library(self):
        # _safe kept leading dots, so an alloy of '..' became a path segment
        # that walked out of the library directory.
        self.assertNotIn('..', photo_lib._safe('..'))
        self.assertNotIn('/', photo_lib._safe('../../etc'))
        self.assertEqual(photo_lib._safe(''), 'Unknown')


class LocalIndexTests(unittest.TestCase):
    def test_a_corrupt_index_is_refused_rather_than_silently_replaced(self):
        # Returning [] for an unreadable index meant the next save wrote a fresh
        # index over the real one and orphaned every stored image.
        with tempfile.TemporaryDirectory() as folder:
            with open(photo_lib._index_path(folder), 'w', encoding='utf-8') as f:
                f.write('[{"job": "6943", trunca')

            with self.assertRaises(RuntimeError) as caught:
                photo_lib._load_local_index(folder)
            self.assertIn('refusing to overwrite', str(caught.exception))

    def test_a_missing_index_is_simply_empty(self):
        with tempfile.TemporaryDirectory() as folder:
            self.assertEqual(photo_lib._load_local_index(folder), [])

    def test_saving_the_index_leaves_no_half_written_file_behind(self):
        with tempfile.TemporaryDirectory() as folder:
            photo_lib._save_local_index([{'job': '6943'}], folder)

            self.assertEqual(photo_lib._load_local_index(folder),
                             [{'job': '6943'}])
            self.assertFalse(os.path.exists(photo_lib._index_path(folder) + '.tmp'))

    def test_an_unclassified_workbook_is_not_filed(self):
        # The UI gated on report type; the CLI did not, so an 'unknown' layout
        # dumped its images into the library with no usable metadata.
        with tempfile.TemporaryDirectory() as folder:
            added = photo_lib.add_to_library(
                'mystery.xlsx', b'not a workbook', {}, 'unknown', folder)

            self.assertEqual(added, 0)
            self.assertEqual(photo_lib._load_local_index(folder), [])


class _FakeResponse:
    def __init__(self, payload=None, content=b''):
        self._payload = payload or {}
        self.content = content

    def json(self):
        return self._payload


class GitHubIndexTests(unittest.TestCase):
    def test_an_index_over_one_megabyte_is_re_fetched_raw(self):
        # The Contents API stops inlining content past ~1 MB and returns
        # encoding 'none' with an empty body. Decoding that as base64 produced
        # an exception on every read, which the caller turned into a permanent
        # "index unreadable" outage once the library grew.
        index = [{'job': '6943', 'image': 'image1.jpeg'}]
        calls = []

        def fake_get(path, raw=False):
            calls.append((path, raw))
            if raw:
                return _FakeResponse(content=json.dumps(index).encode('utf-8'))
            return _FakeResponse({'sha': 'abc123', 'content': '', 'encoding': 'none'})

        with patch.object(gh_store, '_get', side_effect=fake_get):
            sha, loaded = gh_store._read_index()

        self.assertEqual(sha, 'abc123')
        self.assertEqual(loaded, index)
        self.assertTrue(any(raw for _path, raw in calls),
                        'large index was never re-fetched with the raw media type')

    def test_a_small_index_still_comes_from_the_inline_base64(self):
        import base64
        index = [{'job': '7227'}]
        payload = {
            'sha': 'def456',
            'encoding': 'base64',
            'content': base64.b64encode(json.dumps(index).encode('utf-8')).decode(),
        }

        with patch.object(gh_store, '_get',
                          side_effect=lambda path, raw=False: _FakeResponse(payload)):
            sha, loaded = gh_store._read_index()

        self.assertEqual((sha, loaded), ('def456', index))

    def test_an_unreadable_index_is_never_reported_as_empty(self):
        payload = {'sha': 'x', 'encoding': 'base64', 'content': 'not base64 json'}

        with patch.object(gh_store, '_get',
                          side_effect=lambda path, raw=False: _FakeResponse(payload)):
            with self.assertRaises(RuntimeError):
                gh_store._read_index()


if __name__ == '__main__':
    unittest.main()
