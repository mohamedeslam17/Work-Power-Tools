import unittest
from unittest.mock import patch

from openpyxl import Workbook

from lab_review import (
    _canon_machine,
    _comment,
    _formula_issues,
    _review_comment,
    _review_completeness,
    _review_composition,
    _review_captions,
    _review_evidence_claims,
    _review_legends,
    _review_traceability,
    _select_magnification,
    collect_highlights,
    find_duplicate_compositions,
    picture_magnification_verdicts,
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

    def test_see_comment_with_no_comment_at_all_is_critical(self):
        # _review_comment used to bail out silently on an empty comment, so a
        # report with Result="See comment" and literally no comment text (as
        # in the real 7398 report) only ever got a generic Completeness
        # warning — never the stronger Disposition failure that an *ambiguous*
        # comment gets. collect_highlights() already treated it as "no
        # verdict"; the plain findings list must agree.
        parsed = {
            "comment": None,
            "coating": {},
            "sample": {"material": "GTD 111", "result": "See comment"},
        }

        findings = _review_comment(parsed)

        self.assertTrue(any(
            "no comment at all" in message
            for message in messages(findings, "critical")
        ))

    def test_restored_microstructure_counts_as_a_positive_disposition(self):
        # Auditing 27 real AEG reports found the plain accept/reject
        # vocabulary ("acceptable", "suitable for", ...) essentially never
        # appears — the house style states the outcome as the microstructure
        # or properties having been *restored*/*recovered* by the repair heat
        # treatment, with no explicit accept/reject phrase anywhere. Without
        # recognising that as a positive signal, nearly every real report hit
        # the critical "no clear disposition" finding.
        parsed = {
            "comment": (
                "Stress relief was followed by aging to recover the "
                "microstructure with the reformation of the strengthening "
                "phase (gamma prime), as shown in Pics 3 & 4."
            ),
            "coating": {},
            "sample": {"material": "GTD-111", "result": "See comment"},
        }

        findings = _review_comment(parsed)

        self.assertFalse(any(
            "no clear accept / reject / repair disposition" in message
            for message in messages(findings, "critical")
        ))
        self.assertTrue(any(
            "suitable / positive" in message for message in messages(findings, "info")
        ))

    def test_negated_restore_is_not_a_positive_disposition(self):
        parsed = {
            "comment": (
                "Embrittlement was severe and the microstructure could not "
                "be restored by the standard solution heat treatment."
            ),
            "coating": {},
            "sample": {"material": "GTD-111", "result": "See comment"},
        }

        findings = _review_comment(parsed)

        self.assertTrue(any(
            "no clear accept / reject / repair disposition" in message
            for message in messages(findings, "critical")
        ))

    def test_coating_cell_without_any_comment_is_flagged(self):
        parsed = {
            "comment": None,
            "coating": {"present": None, "type": "MCrAIY", "received": None, "outgoing": None},
            "sample": {"material": "GTD 111", "result": "Acceptable"},
        }

        findings = _review_comment(parsed)

        self.assertTrue(any(
            "no comment describing its condition" in message
            for message in messages(findings, "warning")
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

    def test_nominal_duplicate_header_is_flagged(self):
        # _composition() already computed duplicate_headers/entries for the
        # NOMINAL table (mirroring the Actual side), but _review_composition
        # only ever read composition_meta['actual'] — a mislabeled/duplicated
        # column in the spec itself was silently invisible.
        nominal = {"Al": 3.0, "Cr": 14.0}
        actual = {"Al": 2.9, "Cr": 13.5}
        meta = {
            "nominal": {
                "duplicate_headers": ["Cr"],
                "entries": [
                    {"element": "Al", "raw": "3.0", "value": 3.0, "row": 13, "col": 6},
                    {"element": "Cr", "raw": "14.0", "value": 14.0, "row": 13, "col": 7},
                    {"element": "Cr", "raw": "14.0", "value": 14.0, "row": 13, "col": 8},
                ],
            },
            "actual": {"duplicate_headers": [], "entries": []},
        }

        findings = _review_composition(nominal, actual, meta)

        self.assertTrue(any(
            "Nominal composition table repeats element header" in message
            for message in messages(findings, "critical")
        ))

    def test_nominal_total_outside_sanity_band_is_flagged(self):
        meta = {
            "nominal": {
                "duplicate_headers": [],
                "entries": [
                    {"element": el, "raw": str(v), "value": v}
                    for el, v in {"Ni": 20.0, "Cr": 5.0, "Co": 4.0, "Mo": 1.0, "W": 1.0}.items()
                ],
            },
            "actual": {"duplicate_headers": [], "entries": []},
        }

        findings = _review_composition({}, {}, meta)

        self.assertTrue(any(
            "Nominal composition totals" in message and "sanity band" in message
            for message in messages(findings, "critical")
        ))

    def test_hyphenated_unetched_caption_is_recognised(self):
        # _UNETCHED_PAT required the literal contiguous word "unetched", so the
        # common hyphenated spelling "Un-etched" (used in real AEG captions,
        # e.g. report 7504's Picture 1) was never surfaced as an explicit
        # unetched/as-polished caption.
        parsed = {
            "pictures": [
                ("Picture 1:", "Micrograph- Shroud Tip Overview\nUn-etched 25x"),
            ],
            "comment": "",
        }

        findings = _review_captions(parsed)

        self.assertTrue(any(
            "states unetched / as-polished" in message
            for message in messages(findings, "info")
        ))

    def test_plural_pics_range_reference_is_parsed(self):
        # The old regex only matched singular "Pic"/"Picture" — real AEG
        # comments write the plural "Pics." with an en-dash range ("Ref. pics.
        # 9-10"), which never matched at all, so an out-of-range reference
        # written that way could slip through unflagged.
        parsed = {
            "pictures": [(f"Picture {n}:", "Etched 500x") for n in range(1, 5)],
            "comment": "Solution HT was performed (Ref. pics. 9–10).",
        }

        findings = _review_captions(parsed)

        self.assertTrue(any(
            "Comment refers to Picture 10 but only 4 picture(s)" in message
            for message in messages(findings, "warning")
        ))

    def test_underscore_separated_filename_matches_stage_and_component(self):
        # Real AEG filenames use '_' as the word separator. Regex '_' counts
        # as a \w character, so \b-anchored patterns (component/stage) never
        # matched across it — every underscore-joined filename spuriously
        # looked like it disagreed with its own report content.
        parsed = {
            "header": {"job": "7398", "machine": "MS7001", "customer": "AEN SAUDI"},
            "sample": {"description": "1st Stage Bucket"},
        }

        findings = review_filename(
            "7398__AEN_Saudi_FS.7_1st_Stage_Bucket_Metallurgical_Report__AEG_final.xlsx",
            parsed,
            "metallurgical",
        )

        self.assertFalse(any(
            severity in ("critical", "warning") and category == "Title identity"
            for severity, category, _message in findings
        ))
        self.assertTrue(any(
            "stage" in message and "component" in message
            for message in messages(findings, "pass")
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

    def test_written_600x_is_not_rejected(self):
        parsed = {
            "pictures": [
                ("Picture 1:", "Middle Wall Micrograph\nOrtho. Acid Etched 600x"),
            ],
            "comment": "",
        }

        findings = _review_captions(parsed)

        self.assertFalse(any(
            category == "Magnification"
            for _severity, category, _message in findings
        ))

    def test_one_ocr_number_cannot_be_promoted_to_a_fact(self):
        selected, votes = _select_magnification([
            "7646_E_500x-1",
            "7646_E_600x-1",
            "unreadable",
        ])

        self.assertIsNone(selected)
        self.assertEqual(votes, {500: 1, 600: 1})

    def test_consensus_accepts_600_without_an_allow_list(self):
        selected, votes = _select_magnification([
            "7646_E_600x-1",
            "7646 E 600x",
            "7646_E_500x-1",
        ])

        self.assertEqual(selected, 600)
        self.assertEqual(votes, {600: 2, 500: 1})

    def test_caption_wins_when_stable_ocr_disagrees(self):
        image = {
            "image": "image1.png",
            "mag": "600x",
            "ocr_mag": "600x",
            "mag_source": "ocr",
            "mag_votes": {"600x": 2, "500x": 1},
            "mag_read_count": 3,
            "id": "7646_E_600x-1",
        }
        pair = [(
            "Picture 1:",
            "Middle Wall Micrograph\nOrtho. Acid Etched 500x",
            image,
        )]

        with patch("lab_review._picture_image_pairs", return_value=pair):
            verdicts = picture_magnification_verdicts(
                [image],
                [("Picture 1:", "Ortho. Acid Etched 500x")],
                b"workbook",
            )

        self.assertEqual(image["mag"], "500x")
        self.assertEqual(image["ocr_mag"], "600x")
        self.assertEqual(image["mag_source"], "caption")
        self.assertNotIn("id", image)
        self.assertEqual(verdicts[0]["severity"], "warning")
        self.assertIn("caption states 500x", verdicts[0]["note"])
        self.assertIn("OCR consistently read 600x", verdicts[0]["note"])

    def test_matching_written_600x_is_preserved(self):
        image = {
            "image": "image1.png",
            "mag": "600x",
            "ocr_mag": "600x",
            "mag_source": "ocr",
            "mag_votes": {"600x": 3},
            "mag_read_count": 3,
        }
        pair = [("Picture 1:", "Ortho. Acid Etched 600x", image)]

        with patch("lab_review._picture_image_pairs", return_value=pair):
            verdicts = picture_magnification_verdicts(
                [image],
                [("Picture 1:", "Ortho. Acid Etched 600x")],
                b"workbook",
            )

        self.assertEqual(verdicts, [])
        self.assertEqual(image["mag"], "600x")
        self.assertEqual(image["mag_source"], "caption")

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

    def test_frame_5_machine_alias_is_recognised(self):
        # _canon_machine restricted the GE frame digit to [679]; a real report
        # in a 27-report batch used "FS.5" (GE Frame 5 exists — MS5001), which
        # silently failed to resolve at all, dropping the machine/set
        # cross-check for that report instead of matching or flagging it.
        self.assertEqual(_canon_machine('AEN Saudi FS.5 1st Stage Bucket'), 'MS5001')
        self.assertEqual(_canon_machine('MS 5001'), 'MS5001')

    def test_duplicate_actual_composition_across_different_jobs_is_flagged(self):
        # Found on a real 27-report batch: two different AEG job numbers
        # (different customer refs, serials, engineers, dates) carried a
        # byte-identical Actual composition on every matched element — almost
        # certainly a copy-paste, since independent EDS/ICP results are not
        # expected to match exactly. A single-report review can never see
        # this; it only becomes visible across a batch.
        actual = {
            'Ni': 59.53, 'Cr': 13.57, 'Co': 9.14, 'Mo': 1.46,
            'W': 4.73, 'Al': 2.49, 'Ti': 4.81, 'Ta': 3.05, 'Fe': 0.08,
        }
        reports = [
            ('6630.xlsx', {'header': {'job': '6630'}, 'actual': dict(actual)}),
            ('6991.xlsx', {'header': {'job': '6991'},
                           'actual': dict(actual, Cu=0.07)}),
            ('unrelated.xlsx', {'header': {'job': '7000'},
                                 'actual': {'Ni': 61.0, 'Cr': 12.0, 'Co': 9.0,
                                            'Mo': 1.5, 'W': 4.0}}),
        ]

        findings = find_duplicate_compositions(reports)

        self.assertEqual(len(findings), 1)
        severity, category, message = findings[0]
        self.assertEqual(severity, 'critical')
        self.assertEqual(category, 'Composition')
        self.assertIn('6630.xlsx', message)
        self.assertIn('6991.xlsx', message)
        self.assertIn('9 matched element', message)

    def test_same_job_number_is_not_a_duplicate_composition_signal(self):
        actual = {'Ni': 60.0, 'Cr': 14.0, 'Co': 9.5, 'Mo': 1.5, 'W': 3.8}
        reports = [
            ('a.xlsx', {'header': {'job': '7000'}, 'actual': dict(actual)}),
            ('b.xlsx', {'header': {'job': '7000'}, 'actual': dict(actual)}),
        ]

        self.assertEqual(find_duplicate_compositions(reports), [])

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

    def test_identity_and_magnification_warnings_are_marked_on_report(self):
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
                ("Picture 1:", "Middle Wall Micrograph\nOrtho. Acid Etched 500x"),
            ],
            "photo_magnification": [{
                "index": 0,
                "label": "Picture 1:",
                "severity": "warning",
                "note": (
                    "Picture 1: written caption states 500x, while legend OCR "
                    "consistently read 600x in 2 of 3 preprocessing passes. "
                    "Treat the caption as the recorded value and verify the "
                    "burned-in legend; OCR may be wrong."
                ),
            }],
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
            and item["severity"] == "warning"
            and "600x" in item["note"]
            for item in highlights
        ))


if __name__ == "__main__":
    unittest.main()
