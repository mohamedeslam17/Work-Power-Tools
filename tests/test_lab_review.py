import unittest

from openpyxl import Workbook

from lab_review import (
    _comment,
    _formula_issues,
    _review_comment,
    _review_completeness,
    _review_composition,
    _review_captions,
    _review_evidence_claims,
    _review_legends,
    _review_traceability,
    collect_highlights,
    review_filename,
)


def messages(findings, severity=None):
    return [
        message for sev, _category, message in findings
        if severity is None or sev == severity
    ]


class LabReviewRegressionTests(unittest.TestCase):
    def test_plural_comments_label_is_parsed(self):
        ws = Workbook().active
        ws["A1"] = "Comments:"
        ws["A2"] = "A complete metallurgical discussion."

        value, location = _comment(ws)

        self.assertEqual(value, "A complete metallurgical discussion.")
        self.assertEqual(location["value"], (2, 1))

    def test_missing_captioned_micrographs_is_release_blocking(self):
        parsed = {
            "header": {
                "customer": "AEN",
                "job": "7504",
                "machine": "MS7001FA",
                "customer_ref": "4400015000",
                "eoh": "13934",
            },
            "sample": {"material": "GTD-111"},
            "comment": "A sufficiently long metallurgical discussion for the report.",
            "pictures": [(f"Picture {number}:", "Ortho. Acid Etched 100x")
                         for number in range(1, 11)],
            "micrograph_count": 4,
            "signoff": {"met_lab": "Lab", "mat_eng": "Engineer", "date": "2026-07-28"},
        }

        findings = _review_completeness(parsed)

        self.assertTrue(any(
            "6 evidence image(s) are missing" in message
            for message in messages(findings, "critical")
        ))

    def test_duplicate_headers_and_major_lod_are_not_silently_dropped(self):
        nominal = {"Al": 3.0, "Cr": 14.0, "Nb": 2.8}
        actual = {"Cr": 1.7, "Nb": 1.43}
        meta = {
            "actual": {
                "duplicate_headers": ["Nb"],
                "entries": [
                    {"element": "Cr", "raw": "1.7", "value": 1.7},
                    {"element": "Al", "raw": "<LOD", "value": None},
                    {"element": "Nb", "raw": "1.43", "value": 1.43},
                    {"element": "Nb", "raw": "1.43", "value": 1.43},
                ],
            },
        }

        findings = _review_composition(nominal, actual, meta)
        critical = messages(findings, "critical")

        self.assertTrue(any("repeats element header" in message for message in critical))
        self.assertTrue(any("major alloying element was not quantified" in message
                            for message in critical))

    def test_see_comment_requires_a_real_disposition(self):
        parsed = {
            "comment": (
                "The microstructure contains primary carbides and gamma-prime. "
                "Solution heat treatment dissolved part of the secondary phases."
            ),
            "coating": {},
            "sample": {"material": "Rene80", "result": "See comment"},
        }

        findings = _review_comment(parsed)

        self.assertTrue(any(
            "no clear accept / reject / repair disposition" in message
            for message in messages(findings, "critical")
        ))

    def test_microstructure_does_not_prove_restored_mechanical_properties(self):
        parsed = {
            "comment": (
                "The heat-treatment cycle restores the material's microstructural "
                "stability and mechanical properties."
            ),
            "hardness": {
                "pre": {"value": 35.2, "unit": "HRC"},
                "post": {"value": 32.2, "unit": "HRC"},
            },
        }

        findings = _review_evidence_claims(parsed)

        self.assertTrue(any(
            "no final aged-condition mechanical result" in message
            for message in messages(findings, "warning")
        ))

    def test_sample_and_serial_count_mismatch_is_intentionally_skipped(self):
        parsed = {
            "sample": {
                "sample_no": "MS 7221C    MS 7217C\nMS 7375C\nMS 7441C",
                "serial": "K3WP045262\nK3WP045203\nK3WP045216",
                "result": "See comment",
            },
        }

        findings = _review_traceability(parsed)

        self.assertFalse(any(
            "sample-to-part mapping" in message
            for message in messages(findings)
        ))

    def test_spaced_serial_prefix_is_one_identifier(self):
        parsed = {
            "sample": {
                "sample_no": "MS 7242C\nMS 7306C\nMS 7431C\nMS 7439C",
                "serial": "H3AR 027331\nH3AR 027329\nH3AR 027537\nH3AR 016178",
                "result": "Acceptable",
            },
        }

        findings = _review_traceability(parsed)

        self.assertFalse(any(category == "Traceability" for _sev, category, _msg in findings))

    def test_external_formula_reference_is_detected(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "MET"
        ws["L4"] = "=[1]Cover!G44"

        issues = _formula_issues(wb)

        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0]["cell"], "L4")
        self.assertEqual(issues[0]["kind"], "external-reference")

    def test_structured_table_formula_is_not_an_external_reference(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=SUM(Table1[Thickness])"

        self.assertEqual(_formula_issues(wb), [])

    def test_one_correct_photo_does_not_hide_a_wrong_job_photo(self):
        legends = [
            {"image": "image1.png", "job": "7645", "mag": "100x"},
            {"image": "image2.png", "job": "7646", "mag": "25x"},
        ]

        findings = _review_legends(
            legends, ocr_used=True, caption_mags={"25x", "100x"}, report_job="7645")

        self.assertTrue(any(
            "Mixed micrograph job numbers" in message
            for message in messages(findings, "warning")
        ))

    def test_unsupported_600x_caption_is_release_blocking(self):
        parsed = {
            "pictures": [
                ("Picture 1:", "Middle Wall Micrograph\nOrtho. Acid Etched 600x"),
            ],
            "comment": "",
        }

        findings = _review_captions(parsed)

        self.assertTrue(any(
            "Unsupported magnification 600x" in message
            and "25x, 50x, 100x, 200x, 500x, 1000x" in message
            for message in messages(findings, "critical")
        ))

    def test_approved_caption_magnifications_are_accepted(self):
        parsed = {
            "pictures": [
                (f"Picture {index}:", f"Ortho. Acid Etched {mag}x")
                for index, mag in enumerate((25, 50, 100, 200, 500, 1000), 1)
            ],
            "comment": "",
        }

        findings = _review_captions(parsed)

        self.assertFalse(any(
            category == "Magnification"
            for _severity, category, _message in findings
        ))

    def test_title_identity_accepts_machine_aliases(self):
        cases = [
            (
                "7420 AEN Saudi FS.7 3rd Stage Bucket Metallurgical Report.xlsx",
                "7420", "MS7001", "3rd Stage Bucket",
            ),
            (
                "7504 AEN Saudi 7FA 3rd Stage Bucket Metallurgical Report.xlsx",
                "7504", "MS7001FA", "3rd Stage Bucket",
            ),
            (
                "7646 AEN Saudi V84.2 2nd Stage Vane Metallurgical Report.xlsx",
                "7646", "V84.2", "2nd Stage Vane",
            ),
        ]
        for filename, job, machine, description in cases:
            with self.subTest(filename=filename):
                parsed = {
                    "header": {
                        "job": job,
                        "machine": machine,
                        "customer": "AEN Saudi",
                    },
                    "sample": {"description": description},
                }

                findings = review_filename(filename, parsed, "metallurgical")

                self.assertFalse(any(
                    severity in ("critical", "warning")
                    for severity, category, _message in findings
                    if category == "Title identity"
                ))
                self.assertTrue(any(
                    "machine/set" in message
                    for message in messages(findings, "pass")
                ))

    def test_title_identity_rejects_stage_component_and_machine_mismatches(self):
        parsed = {
            "header": {
                "job": "7646",
                "machine": "MS7001FA",
                "customer": "AEN Saudi",
            },
            "sample": {"description": "3rd Stage Bucket"},
        }

        findings = review_filename(
            "7646 AEN Saudi V84.2 2nd Stage Vane Metallurgical Report.xlsx",
            parsed,
            "metallurgical",
        )
        critical = messages(findings, "critical")

        self.assertTrue(any("title says Stage 2" in message for message in critical))
        self.assertTrue(any(
            'title component "vane"' in message
            for message in critical
        ))
        self.assertTrue(any("title machine/set" in message for message in critical))

    def test_new_identity_and_magnification_failures_are_marked_on_report(self):
        title_note = (
            'Report title says Stage 1, but the internal component description '
            'says Stage 2 ("2nd Stage Vane").'
        )
        parsed = {
            "title_findings": [("critical", "Title identity", title_note)],
            "header": {"job": "7646", "machine": "V84.2", "customer": "AEN"},
            "sample": {
                "description": "2nd Stage Vane",
                "material": "Rene80",
                "result": "Acceptable",
            },
            "pictures": [
                ("Picture 1:", "Middle Wall Micrograph\nOrtho. Acid Etched 600x"),
            ],
            "comment": "A sufficiently long metallurgical discussion for this report.",
            "micrograph_count": 1,
            "signoff": {"met_lab": "Lab", "mat_eng": "Engineer", "date": "2026-07-30"},
            "loc": {
                "sheet": "MET",
                "header": {
                    "job": {"value": (4, 4)},
                    "machine": {"value": (3, 12)},
                    "customer": {"value": (3, 4)},
                },
                "sample": {
                    "description": {"value": (9, 4)},
                    "material": {"value": (9, 10)},
                    "result": {"value": (9, 12)},
                },
                "pictures": [{"value": (20, 4)}],
                "signoff": {},
            },
        }

        highlights = collect_highlights(parsed)

        self.assertTrue(any(
            item["category"] == "Title identity"
            and item["cell"] == (9, 4)
            for item in highlights
        ))
        self.assertTrue(any(
            item["category"] == "Magnification"
            and item["cell"] == (20, 4)
            and "600x" in item["note"]
            for item in highlights
        ))


if __name__ == "__main__":
    unittest.main()
