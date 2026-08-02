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

## 003 · Sonnet → Opus · 2 Aug 2026

**Status:** 1a–1f done, plus 1c corrected to the real (corpus) shape. 1g
(D11) deliberately not started — see Found. `tests/test_parser_contract.py`:
15/15 pass. `tests/test_lab_review.py`: 33/33 pass unmodified.
`tests/test_corpus_regression.py`: 8/8 pass, 2 open (both D11). Full local
suite (`python3 -m unittest discover -s tests`, corpus recovered): 63 tests,
OK, 2 expected failures.

**Sequencing note, so the two commits on this branch make sense together:**
I started this session on entry 001's plan (Phase 1, 1a→1f, "start with
D9") and had 1a–1f fully implemented and passing — including the
row-per-sample shape of D3 — *before* fetching and discovering entry 002 had
landed concurrently on this same branch with the corpus-corrected spec. I
rebased on top of it rather than overwrite it. Net effect: everything in
entry 002's revised order got done except 1g, just not in that order, since
1a/1b/1d/1e were already finished by the time I saw the reordering.

**Done**
- 1a/1b/1d/1e/1f as scoped: merge-aware label→value resolution with a
  boundary stop-set (D2, D9, D10), `parsed['fields']` with
  value/raw/cell/status/confidence (D1), every composition table not just
  the first (D4), `_num()` hardening against cross-reference prose and
  left-censored numeric forms (D5). Detail in the first commit on this
  branch and unchanged since.
- **1c corrected to the real shape.** Real AEG reports don't put one sample
  per row — every sample is packed into a single cell, whitespace/newline
  separated, exactly as entry 002 found. `_sample_rows()` still reads every
  row of the table (closing the row-based D3 shape, which the corpus doesn't
  use but the synthetic fixtures still check and now pass), and a new
  `_fan_out_row()`/`_split_packed()` step turns a packed row into N
  per-sample records: whichever field packs into the same token count as
  `sample_no` is distributed index-paired with it, anything else (a shared
  material, a verdict deferred to the comment) is broadcast onto every
  fanned sample. `parsed['samples']` is the fanned canonical list.
  `parsed['sample']` (back-compat) stays the **un-fanned row**, not
  `samples[0]` — `_review_traceability`'s existing `_identifier_tokens()`
  counting already tokenizes a packed cell correctly by itself, and would be
  shortchanged by seeing only the first fanned sample. Verified end to end:
  this is exactly what still produces the corpus's "genuine third critical"
  (the duplicated serial on report 6943) — recovered the corpus, ran
  `review_report()` against all four real MET files, and confirmed the
  duplicate-serial critical, the two constant criticals, and nothing
  spurious from the new `_review_samples()` rule (which stayed silent on all
  four, since every real `Result` is "See comment", not a reject/scrap
  token).
- **The corpus also caught label patterns I'd guessed wrong.** My first
  commit's `_HEADER_LABELS` patterns for `job`/`aeg_ref`/`customer_ref`/`qty`
  didn't tolerate this template's period-abbreviated labels (`'AEG. Job.
  No:'`, `'AEG. Ref. No:'`, `'Customer Ref. No.:'`, `'Quantity p. Set:'`) —
  the strict label matcher requires the pattern to consume the whole cell,
  and mine didn't account for the periods, so all four real reports came
  back `not_located` for those fields. Recovering the corpus and running
  `test_corpus_regression.py` caught this immediately; fixed in the second
  commit. This is exactly the failure mode flagged as a risk in the first
  commit's message ("no real report to check it against") — worth noting
  because it argues for recovering the corpus *before* writing any label
  pattern, not just before/after the whole change.
- Also needed for the corpus shape: report 7227 has zero blank rows between
  the sample table and the next section (`'Hardness Results:'` sits directly
  under the last sample row), so the label-boundary index alone didn't stop
  the row scan there. Every label in this report family ends in `':'` and no
  observed sample-number value does, so `_sample_rows()` now also stops at a
  colon-terminated primary-column cell.

**Found**
- **1g (D11) is not a display change, it's a rules change, and I didn't
  implement it.** I worked out why while checking whether it was reachable
  without touching rule bodies: `test_no_critical_fires_on_every_report` and
  `test_report_specific_findings_outnumber_constant_ones` call
  `review_report()` **independently per file** — no batch context, no shared
  state — and assert that a specific critical (the ones that are always true
  of this template) doesn't come back in a single file's own findings. There
  is no way to satisfy that without either (a) `review_report()` gaining
  cross-report memory it currently has no access to, which isn't what these
  tests exercise, or (b) actually changing what
  `_review_acceptance_and_methods()`/`_review_comment()` emit for those two
  specific checks — a rule-body change. Entry 002's own question to Mohamed
  ("is that a template problem to fix once at source, or should the tool
  keep saying it every time... that answer decides whether 1g is a display
  change or a rules change") is, on inspection, not actually open-ended: it
  *is* a rules change, full stop, given how the acceptance tests are
  written. I did not want to guess the answer to a question the log itself
  flagged as Mohamed's, so I left 1g alone rather than either weaken the
  rule bodies or bend the tests to fit a display-only implementation that
  wouldn't satisfy them. The two D11 tests are still `@unittest.
  expectedFailure`; `test_the_constant_set_is_measured_and_has_not_grown`
  (the guard) still passes unmodified.
- Minor spec/test inconsistency, not blocking: REBUILD.md §6's acceptance
  criterion 1 still reads "every `@unittest.expectedFailure` in
  `test_parser_contract.py` has been removed and all 15 tests pass," but
  entry 002 re-decorated two of those 15 (the row-based D3 tests) as
  intentionally still-open/unverified. Both now pass anyway (see Done), so
  it isn't live, but the criterion's wording and the file's own docstring
  disagree about whether that was ever meant to be permanent.

**Questions**
- None blocking. 1g/D11 is the one open item, and it's genuinely Mohamed's
  call, not a blocked-on-me item — nothing else in Phase 1 depends on it.

**Next**
- 1g (D11), once Mohamed answers entry 002's question. Concretely, once it's
  a "fix at source" decision: downgrade or restructure whichever specific
  checks in `_review_acceptance_and_methods()`/`_review_comment()` produce
  the two constant criticals (which is a rule-body change, so flag it as
  such rather than folding it in quietly), or, if the answer is "keep saying
  it," close D11 by adding real scope information instead (e.g. a
  batch-level function alongside the existing `add_version_findings()`
  pattern that marks — not deletes — the template-constant ones so the UI
  can eventually display them separately; that's a Phase-2/UI-adjacent
  follow-up either way, not a Phase 1 one).
- Phase 2: port the rule bodies onto the canonical `fields`/`samples`/
  `composition` model, retiring the back-compat views as each rule moves
  over. `_review_samples()` generalizes naturally once rules read `samples`
  directly instead of the singular back-compat `sample`.
- The corpus is thin (five reports, one template family) — broadening it is
  still valuable but no longer blocking, per entry 002.

---
## 004 · Mohamed's decision, recorded by Opus · 2 Aug 2026

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
**Note on ordering.** Sonnet's entry 003 landed on this branch concurrently
with this change and was merged, not overwritten — hence this entry is 004.
Sonnet's Phase 1 work is unaffected by the removal; the two touch different
code. Worth connecting explicitly: Sonnet **declined** to implement 1g,
reasoning that removing an always-true critical is a rule-body change and
Mohamed's call rather than something to guess. That reasoning was right, and
this entry is that call arriving for one of the two. The other is still open
above. Measured after both changes: constant findings 9 → 7, constant
criticals 2 → 1.

---

## 005 · Mohamed's decision, recorded by Opus · 2 Aug 2026

**Status:** Second always-true critical removed. D11 closed for criticals.

**Decision**
Asked where the controlling specification actually lives, Mohamed answered:
it lives elsewhere — work order, customer contract — and the report is not
meant to restate it. So a report that does not cite it is not defective.

**Changed**
- `_review_acceptance_and_methods()` — the `Acceptance criteria` critical
  removed. The `Chemistry method` and `Hardness evidence` warnings are
  deliberately **kept**: how the chemistry and hardness work was performed and
  recorded does belong in the report.
- **Extended slightly beyond the literal instruction, flagged here so it can be
  reversed in one line:** `'governing material specification'` was also dropped
  from the `Chemistry method` marker list. It is the same complaint as the
  removed critical, worded differently, and leaving it would have kept the
  phrase in every report's findings. The other three markers (test method /
  instrument, calibration / traceability, measurement uncertainty) are
  untouched and still fire. Reinstate by re-adding the `_CONTROL_REFERENCE`
  entry to that dict.
- `tests/test_lab_review.py::test_missing_release_basis_is_explicit` →
  **inverted** and renamed
  `test_missing_release_basis_is_not_a_finding_but_methods_still_are`. Added
  `test_a_report_stating_its_specification_is_still_not_penalised` to guard
  that the surviving method checks recognise real evidence.
- `test_no_critical_fires_on_every_report` **passed** — decorator removed.
- Corpus guard tightened 7 → 6.

**Effect on the real corpus**

| | as audited | now |
|---|---|---|
| Constant findings | 9 | 6 |
| Constant criticals | 2 | **0** |
| Reports with no critical | 0 | **3** |
| Reports with a critical | 4 | **1** |

Report 6943 now carries exactly one critical in the whole corpus: its
duplicated serial `C1ZP 093046`. Three reports are clean. A critical means
something for the first time.

**Found**
- Both removals were verified before being made, not after. For entry 004 I
  read all four real comments to confirm the rule was not simply failing to
  detect a verdict that was present; it was not. For this one the rule was
  likewise correct — none of the four reports cites a controlling spec. In
  both cases the rule was factually right and the *expectation* was wrong,
  which is not something a test suite can tell you. It needed Mohamed.

**Next**
- D11 is closed for criticals, open for warnings: six constant warnings remain,
  and on the shortest report (7253) they still outnumber the report-specific
  ones. `test_report_specific_findings_outnumber_constant_ones` is the one
  remaining `expectedFailure` in the corpus file.
- That residue is the real `scope` work Sonnet described in entry 003 — mark
  template-constant findings rather than delete them, so the UI can show them
  once. It is UI-adjacent, so it belongs with the Phase 4 work, not here.
- Phase 2 (porting rule bodies onto the canonical model) is otherwise the next
  substantial item.
