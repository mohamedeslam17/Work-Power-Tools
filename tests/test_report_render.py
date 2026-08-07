import io
import unittest
from unittest.mock import patch

import fitz
import openpyxl
from PIL import Image, ImageChops, ImageDraw
from openpyxl.worksheet.pagebreak import Break

from report_render import (
    _annotate_faithful_pages,
    _compose_download_pages,
    _download_issue_cards,
    _faithful_view,
    _pages_to_pdf,
    _place_callout_cards,
    _trim_blank_trailing_pages,
    build_issue_index,
    compose_shared_report,
    crop_issue_detail,
    normalize_comments,
)


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

    def test_faithful_render_preserves_workbook_pagination(self):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "MET"
        sheet["A1"] = "Report"
        sheet["C20"] = "Last printed cell"
        sheet.print_area = "A1:C20"
        sheet.page_setup.orientation = "portrait"
        sheet.page_setup.scale = 76
        sheet.page_margins.left = 0.42
        sheet.row_breaks.append(Break(id=10))
        source = io.BytesIO()
        workbook.save(source)

        captured = {}

        def inspect_saved_workbook(path, _outdir, _timeout):
            annotated = openpyxl.load_workbook(path)
            saved = annotated["MET"]
            captured.update({
                "print_area": str(saved.print_area),
                "orientation": saved.page_setup.orientation,
                "scale": saved.page_setup.scale,
                "left_margin": saved.page_margins.left,
                "row_breaks": [item.id for item in saved.row_breaks.brk],
            })
            return "fake.pdf"

        with (
            patch("report_render.collect_highlights", return_value=[]),
            patch("report_render._xlsx_to_pdf", side_effect=inspect_saved_workbook),
            patch(
                "report_render._raster_pdf",
                return_value=[Image.new("RGB", (300, 400), "white")],
            ),
            patch("report_render._compose_faithful_view", return_value={"ok": True}),
        ):
            result = _faithful_view(
                source.getvalue(),
                {"loc": {"sheet": "MET"}},
                findings=[],
                filename="report.xlsx",
                dpi=150,
                timeout=30,
            )

        self.assertEqual(result, {
            "ok": True,
            "source_page_count": 1,
            "effective_page_count": 1,
            "omitted_blank_pages": [],
        })
        self.assertEqual(captured["print_area"], "'MET'!$A$1:$C$20")
        self.assertEqual(captured["orientation"], "portrait")
        self.assertEqual(captured["scale"], 76)
        self.assertEqual(captured["left_margin"], 0.42)
        self.assertEqual(captured["row_breaks"], [10])

    def test_blank_trailing_pages_are_not_counted_as_report_pages(self):
        content = Image.new("RGB", (600, 800), "white")
        ImageDraw.Draw(content).rectangle((60, 80, 540, 650), fill="black")
        blank = Image.new("RGB", (600, 800), "white")
        # Footer-only ink must not make this an effective report page.
        ImageDraw.Draw(blank).text((260, 760), "Page 10", fill="black")

        pages, omitted = _trim_blank_trailing_pages([content, blank, blank.copy()])

        self.assertEqual(len(pages), 1)
        self.assertEqual(omitted, [2, 3])

    def test_download_pdf_embeds_all_issue_comments_beside_report(self):
        report = Image.new("RGB", (600, 800), "white")
        entry = {"number": 1, "issue_nums": [1]}
        note = " ".join(
            f"comment{index}" for index in range(180)
        )
        issues = [{
            "num": 1,
            "severity": "critical",
            "category": "Disposition",
            "note": note,
            "refs": ["M9", "B21"],
            "pages": [1],
        }]

        cards, _pad, _height = _download_issue_cards(
            issues, panel_width=500, panel_height=800, dpi=150)
        embedded_words = " ".join(
            line for card in cards for line in card["note_lines"]
        ).split()
        self.assertEqual(embedded_words, note.split())

        download_pages = _compose_download_pages(
            [report],
            [entry],
            issues,
            extras=[],
            filename="report.xlsx",
            dpi=150,
        )
        self.assertEqual(len(download_pages), 1)
        self.assertGreater(download_pages[0].width, report.width + 500)
        self.assertEqual(download_pages[0].height, report.height)

        pdf = _pages_to_pdf(download_pages, dpi=150)
        document = fitz.open(stream=pdf, filetype="pdf")
        self.assertEqual(document.page_count, 1)
        document.close()


def _one_page_view(issue_note="Duplicate serial/part number(s): C1ZP093046."):
    """A rendered view of one page carrying one marked cell."""
    page = Image.new("RGB", (700, 900), "white")
    ImageDraw.Draw(page).rectangle((80, 300, 240, 340), outline="black")
    buffer = io.BytesIO()
    page.save(buffer, format="PNG")
    return {
        "filename": "report.xlsx",
        "pages": [{
            "number": 1,
            "png": buffer.getvalue(),
            "issue_nums": [1],
            "size": [700, 900],
            "placements": [{
                "issue_nums": [1],
                "severity": "critical",
                "cell_bbox": [80, 300, 240, 340],
                "badge_bbox": [225, 285, 265, 315],
            }],
        }],
        "issues": [{
            "num": 1,
            "severity": "critical",
            "category": "Traceability",
            "note": issue_note,
            "cells": [(9, 7)],
            "refs": ["G9"],
            "pages": [1],
        }],
        "extras": [],
        "dpi": 150,
    }


class CalloutTests(unittest.TestCase):
    """The sent report has to explain itself on the page."""

    def test_reviewer_comments_reach_the_page_in_full(self):
        view = _one_page_view()
        comment = (
            "Confirm with the workshop which bucket was actually sectioned. "
            + " ".join(f"word{index}" for index in range(140))
        )
        package = compose_shared_report(
            view,
            comments=[{"text": comment, "author": "M. Eslam", "issue": 1},
                      {"text": "  ", "author": "M. Eslam"}],
            filename="report.xlsx",
        )

        self.assertEqual(len(package["pages"]), 1)
        self.assertTrue(package["pdf"])

        records = list(view["issues"]) + normalize_comments(
            [{"text": comment, "author": "M. Eslam", "issue": 1}])
        cards, _pad, _height = _download_issue_cards(
            records, panel_width=500, panel_height=900, dpi=150)
        embedded = " ".join(
            line for card in cards for line in card["note_lines"]).split()
        # Whole comment present, and continuation cards did not duplicate the
        # leader line back to the same cell.
        for word in comment.split():
            self.assertIn(word, embedded)
        leaders = [card for card in cards if card["anchor_nums"] == [1]]
        self.assertEqual(len(leaders), 2)     # the finding, and its comment

    def test_an_empty_comment_is_not_drawn(self):
        self.assertEqual(normalize_comments([{"text": "   "}, ""]), [])
        numbered = normalize_comments([
            {"text": "first"}, {"text": " "}, {"text": "second", "issue": "3"}])
        self.assertEqual([record["label"] for record in numbered], ["R1", "R2"])
        self.assertEqual(numbered[1]["issue"], 3)

    def test_a_dismissed_finding_keeps_its_callout_with_the_reason(self):
        view = _one_page_view()
        issues = [dict(view["issues"][0],
                       state="dismissed",
                       reason="Confirmed against the workshop record.")]
        cards, _pad, _height = _download_issue_cards(
            issues, panel_width=500, panel_height=900, dpi=150)

        # A marker is drawn on the page for every finding, so every marker needs
        # a callout — a dismissed one says so instead of vanishing.
        self.assertTrue(cards[0]["title"].startswith("1. DISMISSED"))
        body = " ".join(cards[0]["note_lines"])
        self.assertIn("Confirmed against the workshop record.", body)

        package = compose_shared_report(view, issues=issues, filename="report.xlsx")
        self.assertEqual(len(package["pages"]), 1)

    def test_a_callout_is_placed_beside_its_own_row(self):
        tall = {"height": 300, "anchor_nums": [1]}
        short = {"height": 60, "anchor_nums": [2]}
        unanchored = {"height": 60, "anchor_nums": []}
        columns = _place_callout_cards(
            [dict(tall), dict(short), dict(unanchored)],
            tops={1: 400, 2: 700},
            card_top=100, card_bottom=900, gap=10)

        self.assertEqual(len(columns), 1)
        first, second, third = columns[0]
        self.assertEqual(first[1], 400)          # level with its cell
        self.assertEqual(second[1], 710)         # pushed past the card above it
        self.assertEqual(third[1], 780)          # unanchored: next free slot

    def test_callouts_that_do_not_fit_continue_in_another_column(self):
        cards = [{"height": 300, "anchor_nums": [number]} for number in (1, 2, 3)]
        columns = _place_callout_cards(
            cards, tops={1: 100, 2: 400, 3: 700},
            card_top=100, card_bottom=800, gap=10)

        self.assertEqual([len(column) for column in columns], [2, 1])

    def test_a_leader_line_runs_from_the_callout_onto_the_marked_cell(self):
        view = _one_page_view()
        page = view["pages"][0]
        source = Image.open(io.BytesIO(page["png"])).convert("RGB")

        composed = _compose_download_pages(
            [source], [page], view["issues"], [], "report.xlsx", 150)[0]

        # The band across the marked cell's row, right of the cell and left of
        # the comment panel, is blank on the report and drawn on afterwards.
        band = (260, 280, source.width, 330)
        self.assertIsNone(source.crop(band).getbbox() and
                          ImageChops.invert(source.crop(band)).getbbox())
        drawn = ImageChops.invert(composed.crop(band)).getbbox()
        self.assertIsNotNone(drawn, "no leader line drawn between cell and callout")

    def test_locate_zooms_the_marked_cell_of_an_exact_page(self):
        view = _one_page_view()
        page = view["pages"][0]

        crop = crop_issue_detail(page["png"], page["placements"], 1)
        magnified = Image.open(io.BytesIO(crop))

        self.assertGreater(magnified.width, 240 - 80)     # wider than the cell
        self.assertIsNone(crop_issue_detail(page["png"], page["placements"], 9))
        self.assertIsNone(crop_issue_detail(page["png"], [], 1))


if __name__ == "__main__":
    unittest.main()
