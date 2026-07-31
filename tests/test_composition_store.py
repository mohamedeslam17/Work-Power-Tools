import shutil
import tempfile
import unittest
from unittest.mock import patch

import composition_store as cs


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


if __name__ == '__main__':
    unittest.main()
