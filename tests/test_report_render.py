import unittest
from unittest.mock import patch

from PIL import Image, ImageDraw

from report_render import _annotate_faithful_pages, build_issue_index


class ReportRenderTests(unittest.TestCase):
    def test_one_finding_with_two_cells_is_counted_once(self):
        highlights = [
            {
                "cell": (9, 13),
                "severity": "critical",
                "category": "Disposition",
                "note": "Result says See comment, but no disposition is provided.",
            },
            {
                "cell": (21, 2),
                "severity": "critical",
                "category": "Disposition",
                "note": "Result says See comment, but no disposition is provided.",
            },
            {
                "cell": (65, 4),
                "severity": "warning",
                "category": "Captions",
                "note": "No etch status in Picture 5 and Picture 6.",
            },
            {
                "cell": (65, 11),
                "severity": "warning",
                "category": "Captions",
                "note": "No etch status in Picture 5 and Picture 6.",
            },
            {
                "cell": (80, 8),
                "severity": "critical",
                "category": "Composition",
                "note": 'Al: actual result is "<LOD" while nominal is 3 wt%.',
            },
        ]
        findings = [
            ("critical", "Disposition",
             "Result says See comment, but no disposition is provided."),
            ("warning", "Captions", "No etch status in Picture 5 and Picture 6."),
            ("critical", "Composition",
             'Al: actual result is "<LOD" while nominal is 3 wt%.'),
            ("critical", "Composition",
             'Al: actual result is "<LOD" while nominal is 3 wt% — '
             "a major alloying element was not quantified."),
            ("warning", "Title identity", "Filename and report job do not match."),
            ("warning", "Title identity", "Filename and report job do not match."),
        ]

        with patch("report_render.collect_highlights", return_value=highlights):
            issues, extras = build_issue_index({}, findings)

        self.assertEqual(len(issues), 3)
        self.assertEqual(issues[0]["num"], 1)
        self.assertEqual(issues[0]["refs"], ["M9", "B21"])
        self.assertEqual(issues[1]["num"], 2)
        self.assertEqual(issues[1]["refs"], ["D65", "K65"])
        self.assertEqual(issues[2]["refs"], ["H80"])
        self.assertEqual(len(extras), 1)
        self.assertEqual(extras[0]["category"], "Title identity")

    def test_same_issue_number_can_mark_multiple_locations(self):
        page = Image.new("RGB", (500, 300), "white")
        draw = ImageDraw.Draw(page)
        first, second = (255, 214, 217), (255, 214, 211)
        draw.rectangle((40, 40, 160, 90), fill=first)
        draw.rectangle((250, 150, 420, 220), fill=second)
        anchors = [
            {
                "severity": "critical",
                "issue_nums": [1],
                "rgb": first,
                "ref": "M9",
            },
            {
                "severity": "critical",
                "issue_nums": [1],
                "rgb": second,
                "ref": "B21",
            },
        ]
        issues = [{
            "num": 1,
            "severity": "critical",
            "category": "Disposition",
            "note": "One issue, two affected fields.",
            "cells": [(9, 13), (21, 2)],
            "refs": ["M9", "B21"],
            "pages": [],
        }]

        entries, annotated = _annotate_faithful_pages(
            [page], anchors, issues, dpi=150)

        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0]["issue_nums"], [1])
        self.assertEqual(issues[0]["pages"], [1])
        self.assertGreater(len(entries[0]["png"]), 1000)
        self.assertEqual(annotated[0].size, page.size)


if __name__ == "__main__":
    unittest.main()
