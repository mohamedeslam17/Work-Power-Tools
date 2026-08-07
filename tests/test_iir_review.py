"""First tests for the IIR reviewer.

Pins the fixes recovered from claude/tool-audit-improvements-v7bj10 (6 Jul 2026)
that can be checked without a real incoming-inspection workbook. The parser
changes on that branch cannot be verified here — no IIR report is available in
the repo — so they are deliberately NOT covered by these tests; see the branch
salvage notes in docs/FEEDBACK.md.
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import iir_review
from iir_review import (
    CHECK_CATALOG,
    DEFAULT_SEVERITY,
    FAIL,
    OFF,
    PASS,
    WARN,
    apply_overrides,
)


def finding(check, severity, detail='detail'):
    return {'check': check, 'severity': severity, 'category': 'Test',
            'sheet': 'Sheet1', 'detail': detail}


class SeverityOverrideTests(unittest.TestCase):
    def test_passing_every_default_does_not_erase_a_softened_severity(self):
        # The settings panel sends the whole catalog, defaults included. Applying
        # a default as an override overwrote severities the check logic had
        # deliberately softened for context (e.g. the WARN a final-repair report
        # gets for fewer positions), silently re-escalating them to FAIL.
        _category, check, _default = CHECK_CATALOG[0]
        default = DEFAULT_SEVERITY[check]
        softened = WARN if default == FAIL else FAIL

        kept = apply_overrides([finding(check, softened)],
                               {title: DEFAULT_SEVERITY[title]
                                for _cat, title, _sev in CHECK_CATALOG})

        self.assertEqual(len(kept), 1)
        self.assertEqual(kept[0]['severity'], softened)

    def test_an_explicit_override_still_applies(self):
        check = CHECK_CATALOG[0][1]
        target = WARN if DEFAULT_SEVERITY[check] != WARN else FAIL

        kept = apply_overrides([finding(check, DEFAULT_SEVERITY[check])],
                               {check: target})

        self.assertEqual(kept[0]['severity'], target)

    def test_off_drops_the_finding_but_never_a_pass_confirmation(self):
        check = CHECK_CATALOG[0][1]

        self.assertEqual(apply_overrides([finding(check, FAIL)], {check: OFF}), [])
        # A PASS is a confirmation the check ran and was satisfied; switching the
        # check off must not delete that evidence.
        self.assertEqual(len(apply_overrides([finding(check, PASS)], {check: OFF})), 1)


class CatalogCoverageTests(unittest.TestCase):
    def test_every_catalog_check_has_a_default_severity(self):
        for _category, title, severity in CHECK_CATALOG:
            self.assertIn(title, DEFAULT_SEVERITY,
                          f'{title} is offered in the settings with no default')
            self.assertEqual(DEFAULT_SEVERITY[title], severity)

    def test_the_catalog_covers_the_section_based_checks_too(self):
        # Checks emitted only by the section-based (Family B) path were missing
        # from the catalog, so the severity settings — including OFF — could not
        # reach them at all. These five are that path's own checks.
        titles = {title for _category, title, _severity in CHECK_CATALOG}
        for title in ('Material recorded', 'Serial-number list found',
                      'Serial count = stated Qty', 'Item numbers unique',
                      'Overall Assessment tally'):
            self.assertIn(title, titles)


class ScanBoundTests(unittest.TestCase):
    def test_row_scans_are_capped(self):
        # One stray cell at Excel's last row pushes ws.max_row to ~1,048,576 and
        # turned every per-row loop into a minutes-long, worker-freezing scan.
        class Sheet:
            max_row = 1_048_576

        self.assertLessEqual(iir_review._max_row(Sheet()), iir_review._SCAN_ROWS_CAP)

        class SmallSheet:
            max_row = 120

        self.assertEqual(iir_review._max_row(SmallSheet()), 120)


if __name__ == '__main__':
    unittest.main()
