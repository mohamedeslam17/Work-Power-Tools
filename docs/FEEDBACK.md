# Feedback log

Append-only conversation between agents working on the rebuild. Newest at the
bottom. Entry format and rules: [HANDOFF.md](HANDOFF.md#handing-back).

Never edit an earlier entry. Record deviations rather than absorbing them.

---

## 001 · Opus → Sonnet · 2 Aug 2026

**Status:** Audit complete. Phase 1 specified and not started.

**Done**
- Audited all 9,029 lines across 8 modules. Findings in `REBUILD.md`, narrative
  version linked at the bottom of that file.
- Reproduced 10 defects by running the shipped parser against synthetic
  workbooks — none of this is inferred from reading code.
- Encoded D1–D5 and D9–D10 as 12 `expectedFailure` tests plus 3 guard tests in
  `tests/test_parser_contract.py`. Those are your acceptance criteria.
- Confirmed baseline: 33 existing tests pass, 12 expected failures, 3 guards
  pass. `test_report_render.py` fails to import without PyMuPDF — pre-existing.

**Found**
- The single root cause is that the parser cannot distinguish "this field is
  absent from the report" from "I failed to find this field". Both are `None`.
  Fix that and D2, D3, D4, D9 and D10 become tractable; leave it and they stay
  a game of whack-a-mole, which is what the last 19 commits were.
- **D9 deserves your attention first even though D3 is the most dangerous.** A
  blank field adopts whatever sits next on its row — including another field's
  *value*. `machine` silently becomes `'J-1001'`. Because the result is
  non-empty, no completeness check fires and the wrong value flows into
  title-identity and traceability looking perfectly legitimate. Errors that
  announce themselves are cheaper than errors that don't.
- I found D9 while writing the test for D2: my first version passed because the
  field absorbed a neighbour instead of going blank. I had written a negative
  assertion. Assert positively — noted in HANDOFF.md, repeating it here because
  it will catch you too.
- `iir_review.py` independently reimplements the same label-scanning approach
  (`_find_label`, `_value_right_of`, `_num`, `_canon`). Same defects, second
  copy. Out of scope for Phase 1, but do not add a third copy.

**Questions**
- None blocking. Everything in Phase 1 is specified well enough to start.

**Next**
- Start with 1a (merge-aware grid reader with proper label↔value association).
  It closes D2, D9 and D10 together and everything else builds on it.
- Then 1b (field records with `status`) — that is the D1 fix and it is what
  makes the rest honest.
- Then 1c/1d (multi-sample, all composition tables), then 1e (`_num`).
- Keep the back-compat views (1f) working the whole way through so nothing in
  `ui/` or the 33 rule tests breaks while you go.

One thing worth flagging to Mohamed rather than to me: **the golden corpus.**
Everything above was verified against synthetic fixtures I wrote, which means
it proves the defects exist but not that fixing them is sufficient. Real
reports, with an engineer recording the findings that *should* be raised, is
the only thing that will tell us whether the rebuild actually worked. It cannot
be done by an agent and it is the highest-value action available.

---

## 002 · Opus → Sonnet · 2 Aug 2026

**Status:** Corpus recovered. Spec corrected against real files. Phase 1 still
not started — but read this before entry 001's plan, it changes the order.

**Done**
- Found five REAL AEG reports in this repo's own git history — committed early,
  removed when `*.xlsx` was gitignored, blobs still reachable. Four
  metallurgical, one coating. `scripts/recover_corpus.py` materialises them
  into `corpus/` (gitignored; never commit them, never rename them).
- Ran the live reviewer over all five. Added
  `tests/test_corpus_regression.py` — 10 tests, 4 expected failures.

**Found**

*I got D3 wrong.* Real reports do not put one sample per row. Every sample is
packed into a **single cell**, whitespace/newline separated — report 6831 cell
B9 is `'MS 6369C        MS 6411C\nMS 6889C         MS 6931C'`, with matching
serials in G9, one shared material, and `Result = 'See comment'` deferring the
verdict to prose. So the job is tokenize-and-fan-out, not read-more-rows.
`_identifier_tokens()` already tokenizes these correctly and is only ever used
for counting. **The row-based fixtures in `test_parser_contract.py` are
synthetic and unverified — I have marked them so. Where the two files
disagree, the corpus wins.**

*New defect, D11, and I think it outranks most of the parser work.* Across the
four real reports: 9 of 25 distinct findings fire on **all four**, so 60–75% of
any report's output is identical to every other report's. Both criticals are in
that set. The findings are not wrong — the template really never states an
acceptance spec — they are criticisms of the *template*, re-emitted per report.

The cost is sharp: report 6943 has a genuine third critical, serial
`C1ZP 093046` duplicated in one cell. A real find, arriving indistinguishable
from two that always fire. A critical that cannot be absent carries no
information and camouflages the ones that can.

*The parser is better than I judged it.* On all four real reports the job,
machine, customer, material, composition tables and captions extract correctly.
D2/D9/D10 are latent, not routine — they need a blank field or colliding prose.
That lowers their urgency. D9 is still easy to trigger on this template because
header fields sit in label/value pairs on one row, so a blank Customer scans
straight into `Machine Type:`.

**Revised order — this supersedes entry 001's "start with D9"**

1. **1g (D11, finding scope).** Cheapest, largest perceived win, independent of
   everything else. Roughly: tag each check `report` or `template`, stop
   emitting template-scoped ones per report.
2. **1c (D3), to the corpus shape.** The only release-critical defect that is
   active on real files today.
3. **1a/1b (D1, D2, D9, D10).** Still the right foundation, but latent — do it
   after the two above rather than before.
4. 1d (D4) and 1e (D5) last; neither is exercised by the current corpus.

**Questions**
- None blocking.
- For Mohamed, not for you: the two constant criticals ("no governing
  acceptance specification", "Result says See comment") are *true* of every
  report AEG issues. Is that a template problem to fix once at source, or
  should the tool keep saying it every time? That answer decides whether 1g is
  a display change or a rules change. Proceed with the display change and flag
  it — that is reversible either way.

**Next**
- Corpus caveat worth keeping in view: five reports, one template family, one
  coating file. It covers the happy path well and says nothing about template
  drift or revisions. It is enough to stop guessing; it is not enough to
  declare the rebuild finished.

---

## 003 · Mohamed's decision, recorded by Opus · 2 Aug 2026

**Status:** Disposition check removed. Entry 002's question is answered.

**Decision**
Mohamed: *"if you mean that the comment is not conclusive you can drop it."*

Dropped. The check raised a **critical** whenever `Result = "See comment"` and
the comment stated no accept / reject / repair verdict.

**Verified before removing.** I read all four real comments first, in case the
rule was firing because the *detector* missed a verdict that was present. It
was not — every comment is a descriptive record of the metallurgical condition
(heat-treatment sequence, oxidation and coating thicknesses, carbide
morphology) and none states a disposition. The rule was factually correct. It
was applying a QA expectation that does not match how AEG reports work: the
comment records condition, and the release decision does not live in it.

**Changed**
- `lab_review._review_comment()` — the final `elif` branch removed. The
  contradiction branches above it are untouched and remain the useful part of
  that rule.
- `lab_review.collect_highlights()` — the paired cell highlight removed.
- `tests/test_lab_review.py::test_see_comment_requires_a_real_disposition` →
  **inverted**, not deleted, and renamed
  `test_see_comment_without_a_verdict_is_not_a_finding`, so the decision is
  explicit and the check cannot come back by accident. Added
  `test_comment_contradicting_the_result_is_still_flagged` to guard the
  branches that survive.
- Corpus baseline tightened: constant set 9 → 8.

**Effect on the real corpus**

| | before | after |
|---|---|---|
| Distinct findings | 25 | 24 |
| Constant across all four | 9 | 8 |
| Constant criticals | 2 | 1 |
| Reports whose critical count differs | 0 | **1** |

That last row is the point. Report 6943 now stands alone at two criticals
because of its duplicated serial `C1ZP 093046`. One always-true check was
enough to camouflage it. D11 is real and this is the cheapest possible
demonstration.

**Question for Mohamed — same shape, still open**
One constant critical remains: **"No governing acceptance specification, limits
or decision rule is stated, so the report cannot support acceptance as issued."**
It fires on all four real reports for the same reason the dropped one did — the
template does not carry it.

If the acceptance spec genuinely lives outside the report (in the work order,
the customer contract, or the Mat. Eng's sign-off), then this is the same call
and it should go too, and the corpus would be left with **zero** constant
criticals — meaning a critical would finally mean something. Say the word and
I will drop it the same way.

If instead the report *should* cite its controlling spec and simply never does,
then it is a genuine finding about a template gap, and the right fix is D11's
`scope` field — say it once, not on every report.

I have not touched it either way.
