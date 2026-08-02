"""Regression tests against the five REAL AEG reports.

These are the only tests in the project that run against genuine customer
workbooks rather than synthetic fixtures written to match the parser's own
assumptions. Everything asserted here is ground truth read out of the real
files by hand.

    python3 scripts/recover_corpus.py     # materialise corpus/ (gitignored)
    python3 -m unittest tests.test_corpus_regression -v

The whole module skips if the corpus is absent, so the suite still runs in a
fresh clone. Recover it before doing extraction work — synthetic fixtures did
not catch a single one of the defects in docs/REBUILD.md §2.1.

Do not rename the files: the title-identity checks compare the filename
against report content, so the names are load-bearing test input.
"""
from __future__ import annotations

import re
import unittest
from collections import Counter
from pathlib import Path

import lab_review as L

CORPUS = Path(__file__).resolve().parent.parent / "corpus"

MET = {
    "6831": "R 6831 AEN Saudi V84.2 2nd Stage Vane  Metallurgical Report - AEG (final).xlsx",
    "6943": "TBR 6943 AEN Saudi FS.7 1st Stage Bucket  Metallurgical Report - AEG (Final).xlsx",
    "7227": "R 7227 AEN Saudi FS.7 2nd Stage Bucket Metallurgical Report - AEG.xlsx",
    "7253": "R 7253 PSM Thomassen FS.6 Transition Piece  Metallurgical Report - AEG.xlsx",
}
COATING = "6983 TSME FS.6 2nd Stage Bucket Aluminizing Coating Report - AEG.xlsx"

# Ground truth, read out of the real workbooks.
#   job: (machine, customer, material, n_samples, n_serials, n_nominal, n_actual, n_pictures)
TRUTH = {
    "6831": ("V84.2",   "AEN-SEC",            "Rene-80",     4, 4,  9, 11, 4),
    "6943": ("MS7001",  "AEN SAUDI",          "GTD 111",     4, 4, 10, 10, 4),
    "7227": ("MS7001",  "AEN-Saudi",          "GTD-741",     2, 2, 12, 11, 8),
    "7253": ("MS 6001", "PSM Thomassen Gulf", "Nimonic-263", 1, 1,  7,  6, 4),
}

_cache: dict[str, tuple] = {}


def setUpModule():
    if not CORPUS.is_dir() or not list(CORPUS.glob("*.xlsx")):
        raise unittest.SkipTest(
            "corpus/ is empty — run: python3 scripts/recover_corpus.py")


def review(filename):
    if filename not in _cache:
        data = (CORPUS / filename).read_bytes()
        _cache[filename] = L.review_report(filename, data, ocr=False)
    return _cache[filename]


def stems(findings):
    """(severity, category, leading clause) per finding, values stripped."""
    out = set()
    for sev, cat, msg in findings:
        if sev == "pass":
            continue
        out.add((sev, cat, re.split(r"[:.] ", msg)[0][:58]))
    return out


# ── extraction must stay correct on real files ────────────────────────────
class RealReportExtractionTests(unittest.TestCase):
    """Locks in what the parser gets RIGHT today, so a rebuild cannot regress it."""

    def test_report_families_are_detected(self):
        for job, name in MET.items():
            with self.subTest(job=job):
                self.assertEqual(review(name)[0], "metallurgical")
        self.assertEqual(review(COATING)[0], "coating")

    def test_header_identity_is_extracted(self):
        for job, name in MET.items():
            machine, customer, *_ = TRUTH[job]
            with self.subTest(job=job):
                header = review(name)[1]["header"]
                self.assertEqual(header.get("job"), job)
                self.assertEqual(header.get("machine"), machine)
                self.assertEqual(header.get("customer"), customer)

    def test_material_and_composition_tables_are_extracted(self):
        for job, name in MET.items():
            _m, _c, material, _s, _sn, n_nom, n_act, _p = TRUTH[job]
            with self.subTest(job=job):
                parsed = review(name)[1]
                self.assertEqual(parsed["sample"].get("material"), material)
                self.assertEqual(len(parsed["nominal"]), n_nom)
                self.assertEqual(len(parsed["actual"]), n_act)

    def test_captions_are_extracted(self):
        for job, name in MET.items():
            n_pictures = TRUTH[job][7]
            with self.subTest(job=job):
                self.assertEqual(len(review(name)[1]["pictures"]), n_pictures)


# ── D3 · the real multi-sample shape ──────────────────────────────────────
class RealMultiSampleTests(unittest.TestCase):
    """On real reports every sample lives in ONE cell, whitespace/newline
    separated — NOT one row per sample as tests/test_parser_contract.py first
    assumed. B9 on report 6831 reads:

        'MS 6369C        MS 6411C\\nMS 6889C         MS 6931C'

    with the matching serials packed the same way into G9, one shared material
    in K9 and Result = 'See comment' — the per-sample verdict is deferred to
    the free-text comment. Three of the four real reports carry more than one
    sample, so D3 is active on 75% of this corpus.
    """

    def test_identifier_tokenizer_sees_every_sample(self):
        """The tokenizer already gets this right; it is only used for counting."""
        for job, name in MET.items():
            n_samples, n_serials = TRUTH[job][3], TRUTH[job][4]
            with self.subTest(job=job):
                sample = review(name)[1]["sample"]
                self.assertEqual(
                    len(L._identifier_tokens(sample.get("sample_no") or "", "sample")),
                    n_samples)
                self.assertEqual(
                    len(L._identifier_tokens(sample.get("serial") or "", "serial")),
                    n_serials)

    @unittest.expectedFailure
    def test_model_exposes_one_record_per_sample(self):
        """D3 against real files. The tokens exist but never become samples, so
        every downstream check reviews a 4-sample report as though it were one."""
        for job, name in MET.items():
            n_samples = TRUTH[job][3]
            with self.subTest(job=job):
                self.assertEqual(len(review(name)[1]["samples"]), n_samples)

    @unittest.expectedFailure
    def test_shared_material_is_carried_onto_every_sample(self):
        """One material cell covers all samples in the row; each sample record
        must carry it so the hardness and composition checks can run per sample."""
        parsed = review(MET["6831"])[1]
        self.assertEqual([s.get("material") for s in parsed["samples"]],
                         ["Rene-80"] * 4)


# ── D11 · findings have no scope ──────────────────────────────────────────
class FindingScopeTests(unittest.TestCase):
    """Measured on the real corpus: 8 of 24 distinct findings fire on all four
    reports, so a majority of any report's output is identical to every other
    report's.

    Was 9 of 25 with two constant criticals. The "See comment" disposition
    check was dropped on Mohamed's instruction (FEEDBACK.md entry 003), which
    left one constant critical and made report 6943 visibly differ from the
    other three — its duplicated serial is now the only second critical in the
    corpus, which is exactly the signal the constancy was burying.

    These are not wrong — the template genuinely never states an acceptance
    specification — but they are criticisms of the TEMPLATE, re-emitted per
    report. They need a scope dimension so they are said once, not every time.
    """

    def _constant_findings(self):
        per = {job: stems(review(name)[2]) for job, name in MET.items()}
        counts = Counter(k for keys in per.values() for k in keys)
        return {k for k, c in counts.items() if c == len(MET)}, per

    def test_the_constant_set_is_measured_and_has_not_grown(self):
        """Guard: documents the measured baseline so drift is visible."""
        constant, _ = self._constant_findings()
        self.assertLessEqual(
            len(constant), 8,
            "more findings became report-invariant; the noise floor is rising")

    @unittest.expectedFailure
    def test_no_critical_fires_on_every_report(self):
        """D11. One critical still fires on all four reports — "no governing
        acceptance specification". It is true of the TEMPLATE, so every report
        starts at one critical before anything about it has been examined.
        (The second, the "See comment" disposition check, was dropped in
        FEEDBACK.md entry 003.)

        The cost is concrete: report 6943 has a genuine third critical, a
        duplicated serial (C1ZP 093046 appears twice in one cell). That is a
        real find, and it arrives looking exactly like the two that mean
        nothing. A critical that cannot be absent carries no information, and it
        camouflages the ones that can.
        """
        constant, _ = self._constant_findings()
        self.assertEqual(
            [k for k in constant if k[0] == "critical"], [],
            "these criticals fire on 100% of reports and hide the real ones")

    @unittest.expectedFailure
    def test_report_specific_findings_outnumber_constant_ones(self):
        """D11. A review's output should be mostly about the report in hand."""
        constant, per = self._constant_findings()
        for job, keys in per.items():
            with self.subTest(job=job):
                specific = len(keys - constant)
                self.assertGreater(
                    specific, len(keys & constant),
                    f"{job}: {len(keys & constant)} constant vs {specific} specific")


if __name__ == "__main__":
    unittest.main(verbosity=2)
