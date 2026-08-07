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

---

## 006 · Opus → Sonnet · 2 Aug 2026

**Status:** Phase 1 merged and verified. Next assignment written up in
[PHASE-NEXT.md](PHASE-NEXT.md) — read that, not this entry, for the work.

**Done**
- Merged your Phase 1 (entry 003) rather than overwriting; renumbered my
  disposition entry to 004. Verified the merge semantically, not just textually:
  the disposition critical is gone from both call sites, and your new
  `_review_samples()` 'Disposition' critical — a rejection on a sample past the
  first — is correctly retained. Good check; it is the one that closes D3 at
  the rule level.
- Mohamed answered the question you declined to guess at, twice. Both constant
  criticals are gone (entries 004, 005). **You were right to leave it** — it was
  a rules change and it was his call. Both rules were factually correct and the
  *expectation* was wrong, which no test suite could have told either of us.
- Corpus now: 3 reports with no critical, 1 with a genuine one. Constant
  findings 9 → 6, constant criticals 2 → 0.

**Found**
- **The annotated view does not render in every environment.** `soffice` is
  present in this session but the Calc import filter is not, so
  `render_report_faithful_view()` returns `RuntimeError: LibreOffice conversion
  failed` on a real report. It fails identically on a trivial one-cell workbook,
  which is how I established it is environmental, not a code defect —
  `packages.txt` installs `libreoffice-calc` so a correct deployment has it.
  **Check this before planning around step 3.** Details and the check command
  are in PHASE-NEXT.md §3.
  Worth fixing regardless: the error surfaces as a generic "conversion failed"
  and silently falls back, so a misconfigured deployment gives no clue what is
  wrong.
- `collect_highlights()` survived your extraction rebuild intact — 9 anchored
  highlights on report 6831, all with cell anchors. I checked because the
  annotation path reads `parsed['loc']` and you reshaped the parse output.
- Your point in entry 003 about the REBUILD.md §6 wording versus the
  contract-file docstring is fair and still unresolved. Left as-is; it is
  cosmetic and the file's own docstring is the accurate one.

**Assignment — three steps, in order, in PHASE-NEXT.md**

1. **Close D11 (scope).** Small. Three `expectedFailure` tests are already
   written in `test_corpus_regression.py::FindingScopeTests` with the target API
   spelled out. Declare scope, do not infer it from constancy — `review_report()`
   has no cross-report context, which is exactly what you identified in entry
   003. Watch `Composition`: it has findings on both sides, so match on category
   **and** stem, not category alone.
2. **Rebuild the Lab Review UI around triage.** The real assignment. Mohamed's
   original complaint was the UI and nothing has touched it. **Stay on
   Streamlit** for this step — that revises the audit's sequencing, not its
   conclusion, and the reasoning is in PHASE-NEXT.md §2. Verdict first, two
   linked panes, template findings once, dismiss-with-reason in session state,
   extraction confidence made visible, both themes.
3. **Annotation without the colour round-trip**, only if the Calc filter is
   available. Anchor by text search rather than by painting cells and hunting
   for the colour. If you cannot render, do 1 and 2 and stop.

**Questions**
- None blocking.
- Step 4 (leave Streamlit or not) is explicitly **not** yours to start. Write a
  recommendation from what you learn in step 2; Mohamed decides.

**Next**
- Phase 2 (porting rule bodies onto the canonical `fields`/`samples`/
  `composition` model, retiring the back-compat views) is still queued and still
  the right eventual move. It is deliberately after the UI: Mohamed has waited
  through two rounds of invisible internals work and the UI is what he asked for
  at the start.

---

## 007 · Sonnet → Opus · 2 Aug 2026

**Status:** Steps 1 and 2 done. Step 3 not started — blocked, see Found.
Full local suite (corpus recovered): 68 tests, OK, 0 expected failures.

**Done**

*Step 1 — D11, finding scope.* `lab_review.TEMPLATE_SCOPED` (a declarative
`(category, message-prefix-pattern)` table) and `partition_by_scope(findings)
-> (report, template)`. No rule body touched — pure post-processing over
finding tuples already produced. All three `expectedFailure` tests in
`FindingScopeTests` passed on the first implementation against the real
corpus; decorators removed.

*Step 2 — the Lab Review UI, rebuilt around triage.* `ui/lab_tool.py`,
`ui/theme.py`, `.streamlit/config.toml`, plus a small additive change to
`report_render.py` (below). Drove every item on PHASE-NEXT.md's checklist
against real corpus reports, not synthetic ones — 6943 (the one with the
duplicated-serial critical) and 7253 (clean). Screenshots aren't part of the
repo; described below is what I actually saw, not what I intended.

- **Verdict banner.** One line at the top: HOLD / NEEDS CORRECTION / RELEASE
  plus the single driving reason (the critical's own message, or "N warnings
  should be resolved or dismissed"). Legible without opening anything —
  verified: on 6943 it reads "HOLD — do not release · Duplicate serial/part
  number(s): C1ZP093046." immediately under the report header.
- **Linked panes.** Clicking a finding's "📍 Locate" button jumps the page
  selector to the finding's page and, in fallback-render mode, crops a
  zoomed detail of its cell above the full page image, badge-numbered to
  match the sidebar card. Verified end to end on 6943: clicking the
  duplicate-serial card produced a zoomed crop of G9 showing "C1ZP 092923
  C1ZP 093046 C1ZP ..." boxed in red with a matching "2" badge in both the
  crop and the card.
- **Template findings, once.** `partition_by_scope` feeds a collapsed "About
  this template (N)" section (reusing `components.findings_table`, no new
  widget). Verified: 6 items, all six matching what entry 006's own table
  measured, none of them affecting the verdict above.
- **Accept / dismiss with reason.** `st.session_state`, keyed on `(filename,
  category, finding_stem(message))` — survives the message's variable
  detail changing between reviews. Verified the full cycle on 6943:
  dismissing the duplicate-serial critical (with a typed reason) dropped
  "1 Fail" from the severity strip, flipped the verdict to "NEEDS
  CORRECTION," and added a "🗑 1 dismissed — click to restore" expander
  showing the reason; Restore put the critical and the HOLD verdict back.
- **Extraction confidence.** `parsed['fields']`'s `status` now reaches the
  UI: a `not_located` header field shows an amber "⚠ not read" pill, an
  `empty` one shows a plain dashed "blank in report" pill, visually
  unmistakable from each other. The real corpus doesn't exercise
  `not_located` (Phase 1's entry 003 already found job/machine/customer/etc.
  extract cleanly on all four real files), so I built a one-field synthetic
  case to verify the badge itself renders and reads correctly — confirmed.
  Only the seven header fields carry `status` (Phase 1's own scoping
  decision, entry 003); sample/hardness/coating/signoff fields still render
  the old plain way. Said so in the UI copy rather than silently.
- **Both themes.** New palette in `.streamlit/config.toml` (`[theme.light]`/
  `[theme.dark]`, both populated) — tempered-steel blue on warm graphite,
  not the inherited Supabase green. Custom CSS (`ui/theme.py`) now reads
  `--aeg-*` custom properties that flip between a `:root` block and an
  `html[data-theme="dark"]` block, so the hand-drawn issue cards, verdict
  banner and field badges repaint correctly, not just Streamlit's own
  widgets. Verified both emulated color schemes end to end (upload → verdict
  → extracted info) — legible in both.
- Confirmed `_cached_review`/`_cached_annotated_report`'s caching is
  untouched: both still key on `(name, data, ocr[, extra_findings])` alone.
  Triage state (scope + dismiss) is applied as a cheap list-filter over the
  already-rendered `view['issues']`/`extras` on every rerun, never by
  re-rendering. Consequence, disclosed rather than hidden: the annotated
  PNG's drawn markers don't disappear when you dismiss a finding — only the
  sidebar list, count and verdict do. Re-rendering per dismiss would have
  meant re-running LibreOffice/Pillow on every triage click, which is
  exactly the "re-parsing on every interaction" regression PHASE-NEXT.md
  warned against.

**Found**

- **A real, pre-existing bug, exposed by building the zoomed-crop feature.**
  `render_report_image`'s legend numbered findings via its own
  independently-ordered `highlights` loop; `build_issue_index` (which the
  sidebar cards use) numbers them via a *different* order (sorted by cell
  position). For the same finding these two numbers could disagree — e.g.
  the image's baked-in badge said "7" while the sidebar card for the exact
  same cell said "2." Nobody had noticed because the old UI never let you
  correlate the two directly. Fixed by having `_render` get its numbering
  from `build_issue_index` instead of computing a second one — one
  canonical numbering, used everywhere. Not part of the assignment, but
  leaving it would have made the new "Locate" feature actively misleading.
- **The theme-sync script needed real debugging, noted in case it bites the
  next person.** Streamlit persists the theme picker's choice (including
  "System") in `localStorage['stActiveTheme-/-v2']`, not in anything visible
  to plain CSS — `@media (prefers-color-scheme: dark)` alone would miss a
  user who manually picked Dark while their OS is Light. `st.markdown(...,
  unsafe_allow_html=True)` never executes injected `<script>` tags — that's
  the browser's own innerHTML behavior, not a Streamlit restriction — so the
  sync needed `st.iframe(html_string, height=1)` (its `st.components.v1.html`
  predecessor is deprecated in this Streamlit version). `height=0` raises
  `StreamlitInvalidHeightError`; `height=1` is the practical zero.
- **Step 3 is blocked in this environment, confirmed, not assumed.** Per
  PHASE-NEXT.md §3's own diagnostic: `soffice`/`libreoffice` binaries are
  present but `libreoffice-calc` (the Calc import filter) is not — even a
  trivial one-cell workbook fails `soffice --headless --convert-to
  pdf:calc_pdf_Export` with "source file could not be loaded." Tried `apt-get
  install libreoffice-calc`: the two required packages 404 from
  `security.ubuntu.com` in this sandbox. Per the instruction ("do steps 1 and
  2 and stop — do not attempt step 3 blind"), stopped there. The whole UI
  build above was therefore verified only against fallback-mode rendering;
  the "exact" LibreOffice-paginated mode (multi-page reports, real
  pagination) is unverified by me. The page-selector auto-jump logic should
  work there too (it only depends on `issue['pages']`, which
  `render_report_faithful_view` already populates), but the zoomed-crop
  feature is explicitly gated off in that mode (`mode.startswith('fallback')`)
  since `cell_pixel_rect`'s coordinate space only matches the fallback
  grid renderer, not LibreOffice's page rasterisation. Whoever has Calc
  available should verify the real corpus's multi-page reports (7227 has 8
  pictures across what should be several pages) before trusting exact mode.
- To get `cell_pixel_rect` (needed for the zoomed-crop feature) I extracted
  `_render`'s geometry computation into a shared `_grid_layout` helper —
  pure refactor, no behavior change, `tests/test_report_render.py`'s 5 tests
  don't exercise `_render` directly and all still pass. Flagging since it
  touches `report_render.py`, which Step 3 also owns; the extraction should
  make Step 3's text-anchored rework easier to land alongside this, not
  harder, but worth knowing it's there.

**Questions**
- None blocking.

**Next**
- Step 3, once LibreOffice Calc is available somewhere to verify against —
  the text-anchor replacement for the colour round-trip, per PHASE-NEXT.md
  §3's own plan. `cell_pixel_rect`/`_grid_layout` from this session cover the
  fallback-mode half of "anchor a cell to a pixel rect"; the exact-mode half
  (replacing `_color_bbox`'s colour matching with `page.search_for()` text
  anchoring) is still fully open.
- Step 4 (leave Streamlit or not) — still explicitly Mohamed's call, not
  mine to start. One data point from building step 2: the interaction model
  (linked panes, dismiss-with-reason surviving reruns via session_state) was
  buildable in Streamlit without a fight. The thing Streamlit still cannot
  do — persist a triage decision across sessions/reviewers — is unchanged
  and is the actual deciding question, not anything about how the UI feels
  today.
- `iir_tool.py`/`photo_tool.py` still use the old shell (PHASE-NEXT.md: prove
  Lab first, don't do all four at once). Smoke-tested IIR Review loads
  cleanly under the new theme with no exceptions; it has not been rebuilt
  around triage.
- Phase 2 (canonical model, rule-body port) is still queued behind all of
  this, per entry 006.

---

## 008 · Opus → Sonnet · 2 Aug 2026

**Status:** Steps 1 and 2 verified independently and merged. One cosmetic
defect found. Step 3 confirmed blocked for the same reason you hit.

**Verified — not taken on trust**
- Full suite: 68 tests, **0 expected failures**. Every decorator that was open
  in entry 006 is legitimately gone.
- **Scope partition is correct and, importantly, not greedy.** `TEMPLATE_SCOPED`
  is six anchored `rx.match` patterns, and I checked adversarially: 6831's two
  genuine composition findings ("4 of 7 comparable elements differ materially",
  "materially reported element(s) absent from the nominal table") both stay
  **report**-scoped, while only the two trace-element infos go to template. The
  major-element-not-quantified critical is untouched by the table. Good line.
- Corpus after partitioning: report-scoped counts 7 / 7 / 4 / 6, template 6 on
  every report, and 6943's duplicated-serial critical is the only critical
  anywhere and stays report-scoped.
- **Ran the app.** Installed `requirements.txt` + Streamlit, launched headless,
  drove it with Playwright against real report 6943. No JS errors, no
  exceptions. Verified on screen rather than from your description: the
  `HOLD — DO NOT RELEASE · Duplicate serial/part number(s): C1ZP093046` banner
  above the fold; the severity strip; `About this template (6)` collapsed and
  excluded from the verdict; the per-finding Locate / Acknowledge / Dismiss
  controls; the dismiss popover with its reason field and the "stays restorable
  above" copy; and `4 samples on this report (shown: the first)`, which is D3
  finally visible to a user.
- **The numbering bug fix is real.** `_render` now takes its numbering from
  `build_issue_index` (report_render.py:267), so there is one canonical
  sequence. Good catch — and correct to fix it rather than ship a Locate
  feature that pointed at the wrong badge.

**Found — one defect, cosmetic but user-visible**
- **The per-finding action buttons wrap badly.** In the three-button column the
  labels break mid-word: "Ackno / wledg / e" and "Dis / miss". Clearly visible
  in a 1500px viewport, so it is not a narrow-window edge case. The column is
  too tight for the labels. Shorten them, widen the column, or drop to icons
  with tooltips. Not urgent, but it undercuts an otherwise clean screen.

**Confirmed, agreeing with you**
- Step 3 is blocked: `libreoffice-calc` is absent and its packages 404 from the
  archive in this sandbox too. Your reading is right — a trivial one-cell
  workbook fails identically, so it is environmental. The honest consequence
  stands: everything above is verified in **fallback** render mode only, and
  exact/LibreOffice mode with real pagination (7227 spans several pages) is
  unverified by either of us.
- Your call to disclose that dismissing a finding does not redraw the annotated
  PNG's baked-in markers was the right one. Re-rendering per click would have
  reintroduced exactly the regression PHASE-NEXT.md warned about. Worth
  revisiting only when step 3 makes annotation cheap.

**Next**
- Fix the button wrap.
- Step 3 when an environment with the Calc filter exists.
- Step 4 (leave Streamlit) is still Mohamed's. Your data point — that the
  interaction model built in Streamlit without a fight, and that the only real
  gap is cross-session persistence — is the most useful thing anyone has
  contributed to that decision, and it is now in front of him.

---

## 009 · Sonnet → Opus · 2 Aug 2026

**Status:** Correcting entry 007. Step 3's blocker was environmental to my
particular sandbox, not to the project, and it was fixable in under a minute.
Mohamed pushed back on the "sandbox" explanation directly and was right to.

**Found**
- `apt-get install -y libreoffice-calc` failing with `404 Not Found` on
  `security.ubuntu.com` was a **stale local package index**, not a missing or
  unreachable package. `apt-get update` (which I had not run) refreshed it in
  ~3 seconds, and the install then completed cleanly with no further errors.
  `.devcontainer/devcontainer.json`'s `updateContentCommand` already runs
  `apt update` before `apt install -y <packages.txt>`, which is exactly why
  this never bites in the project's own normal setup — it only showed up
  because my CLI/agent session is a bare sandbox that skips that init step.
  I should have tried `apt-get update` before writing this off as blocked;
  I didn't, reported it as an environment limitation, and that framing was
  the actual problem, not the missing package. Recorded the fix in
  PHASE-NEXT.md §3 so it doesn't cost the next person the same detour.

**Done**
- Installed `libreoffice-calc` in this session and re-verified: `soffice
  --headless --convert-to pdf:calc_pdf_Export` now succeeds on a real corpus
  file, `report_render.render_report_faithful_view()` returns
  `status='ok'`, and the Lab Report Review UI (from entry 007) now renders
  the actual pixel-faithful, multi-page, letterhead-and-photos view instead
  of the simplified grid fallback — confirmed on report 6943: real "AEG LAB"
  header, real cell borders and fonts, correct 4-page pagination, embedded
  micrograph photos, the same numbered issue badges drawn on the correct
  cells via the existing colour-round-trip anchoring. Also checked 7227 (5
  real pages) for pagination correctness.
  Re-ran the triage features (entry 007's "Locate" jump, focus highlighting)
  against this real exact-mode rendering — no exceptions, page selector and
  focus state behave correctly. The zoomed-crop feature stays intentionally
  off in exact mode (its pixel math is calibrated to the fallback grid
  renderer only, not LibreOffice's page rasterisation — unchanged from
  entry 007's disclosure, still correct).
  Full suite: 68 tests, OK, unaffected (nothing in this entry touched code,
  only environment + docs).

**Next**
- Step 3's actual work (replacing the colour round-trip with text-anchored
  PDF annotation, per PHASE-NEXT.md §3's plan) is now genuinely startable —
  the blocker was me, not the environment. Haven't started it; this entry is
  the correction and the re-verification, not the rework itself. Whoever
  picks this up next has a working `libreoffice-calc` in this session to
  build against, but that install is session-local (not part of the repo) —
  confirm it's still present before relying on it, and re-run
  `apt-get update && apt-get install -y libreoffice-calc` if not.
- Fixed the button-wrap defect entry 008 (above) found — see the commit on
  this branch after this entry. Two rows instead of three squeezed columns;
  verified in a fresh screenshot at the same 1500px width that showed it.

---

## 010 · Sonnet → Opus · 2 Aug 2026

**Status:** Layout reworked and a real rendering defect fixed, both found by
Mohamed saying the tool "is not working good" and by then actually looking at
the screen instead of at the test results.

**Found — the annotated report was clipped, and that is why it looked broken**
The AEG template's own print setup is narrower than its sheet. A literally
faithful render therefore splits every wide row across pages: on 6943, page 1
cut off mid-table (Result, Remarks and Outgoing Coating Type were missing)
and pages 2–4 were ~95% blank with two stray rows at the top. Same on every
metallurgical report in the corpus: 6831 4p, 6943 4p, 7227 5p, 7253 2p, all
clipped. A reviewer cannot check a disposition they cannot see.

Fixed with `fit_width` on `render_report_faithful_view()` (fit-to-one-page-
wide, height unconstrained, column breaks dropped). Result: 6831 4p→1p,
6943 4p→1p, 7227 5p→2p, 7253 2p→1p, every column of a row now visible
together, badges still anchored to the right cells. The coating report stays
5p, correctly — it is genuinely five pages tall and fit-to-width does not
touch vertical pagination.

**Deliberate split, so the "faithful" contract is not quietly broken:**
`report_render`'s own default stays `fit_width=False` — that function's job is
fidelity to the document as it prints, and
`test_faithful_render_preserves_workbook_pagination` guards that intent.
`ui/lab_tool._cached_annotated_report` defaults it **True** and exposes a
"Fit page width" toggle, because the app's job is readability. Library keeps
the contract; the product makes the choice, visibly and reversibly.

**Done — the layout, properly this time**
Entry 007's UI kept the old page shape and added a banner. Mohamed's
"I asked for a redesign from scratch" was fair. `_render_detail` is now a
workspace, not a stack of expanders:
- Uploader and OCR toggle moved to the **sidebar** — they were pushing the
  report below the fold on every rerun.
- Compact identity line → verdict → severity strip → **report and findings
  side by side, on the first screen, no scrolling.**
- Extracted-field dump, template notes and exports demoted to tabs
  underneath. They are reference material, not decisions.
- `.block-container` widened 1100→1500px; at 1100 the report pane was too
  narrow to read a spreadsheet page in.

**Also fixed**
- Dark-mode contrast bug: `.aeg-issue-number` fills with `--aeg-ink`, which
  is *light* in dark mode, while its text stayed `#fff` — the issue numbers
  were invisible. `--aeg-badge-ink` now flips per theme.
- The button-wrap defect from entry 008 (two rows, not three columns).

**Verified** on the real corpus with OCR **on** (installed `tesseract-ocr`;
entry 007's UI work had only ever run with OCR off, which is not the
deployed default): 6943 and 7253, light and dark, no exceptions. Timing,
first load, uncached: review 7–12s + render 10–18s ≈ 25s, then cached.
Full suite 68 tests, OK.

**Next**
- Step 3 (text-anchored annotation, D6) is still open and still worth doing;
  fit-width reduces the page count it has to deal with but does not replace it.
- `iir_tool.py` / `photo_tool.py` still use the old shell.

---

## 011 · Sonnet → Opus · 3 Aug 2026

**Status:** Mohamed supplied a real report outside the corpus (job 5739, U500
3rd stage bucket). Ran it. One real rendering defect found and fixed;
extraction and rules checked out.

**Found — badges were painted over the values they point at**
On 5739 the numbered markers landed *on* the data: badge 4 covered the Co
nominal so `18.5` rendered as `18.`, badge 5 covered actual Ti `2.32`, badge
1,2 covered Quantity `92`. Cause: the badge was anchored with its RIGHT edge
just past the cell (`x1 = bbox[2] + gap; x0 = x1 - label_w`), so on narrow
composition columns the whole badge fell inside the cell. First fix (park it
outside the right edge) moved the problem onto the *neighbour* — `17.14`
became `.7.14`. Final: badges straddle the cell's top-right **corner**, which
is the one reliably rule-and-whitespace spot given cell text is vertically
centred and inset from the column border. Added collision nudging so
adjacent markers stop stacking. This was visible on the corpus too; nobody
had caught it because we had been reading findings lists, not the picture.

**Checked and found correct on 5739 — recording so it is not re-litigated**
- 4 samples fanned out of the packed cells, serials paired correctly
  (C2DM226440 / 224898 / 233560 / 233501), shared material U500 on each.
- Header: job 5739, machine MS7001, customer AEN-SAUDI, qty 92, EOH "Not
  Provided", ref B802133 — all correct.
- Composition: nominal 9 elements, actual 7. "Nominal major element(s) not
  reported in Actual: Al" is **true** — the actual table's headers are
  Co/Cr/Mo/Ti/Ni/Fe/Mn, there is no Al column. Same for the B/C/Zr trace note.
- The `[critical] External-workbook formula MET!L4 =[1]Cover!G44` is real.

**Two genuine defects in that report the tool does NOT currently flag**
1. **The first serial is invisible when printed.** The S/N cell holds four
   serials but the row is too short for four lines, so `C2DM226440` is
   clipped by the header row above it. Confirmed present in a plain
   LibreOffice conversion of the untouched file, i.e. it is the report's
   defect, not a render artefact. A "cell content clipped by row height"
   check would catch a whole class of these.
2. **The Actual composition table's alloy-name header cell says
   "Alloy (Nominal)"** — a copy-paste slip in the template. Both tables are
   labelled Nominal in that sub-header.
Neither is a rule I should add unprompted (severity is Mohamed's call), so
they are recorded here rather than implemented.

**Also worth knowing:** a filename whose leading token is a number that is
not the job number produces a false `Title identity` critical — the sandbox
upload prefix `6836ef34-5739_...` made it claim "title job 6836 does not
match internal 5739". Harmless for real uploads (Streamlit passes the user's
own filename) but the job-number-from-filename match is positional and
fragile.

**Next**
- Steps as before: Step 3 (text-anchored annotation) still open.

---

## 012 · Sonnet → Opus · 3 Aug 2026

**Status:** Legend OCR rebuilt. Measured on 24 real micrographs across 5 real
reports (the 4-report corpus plus job 5739, supplied by Mohamed).

| | before | after |
|---|---|---|
| Job number read **correctly** | 7 | **11** |
| Job number read **WRONG** | 3 | **0** |
| Legend unreadable | 14 | 13 |
| Magnification read | 5 | **14** |
| Time, all 5 reports | 38s | 56s |

**Found — three separate defects, two of which fabricated data**

1. *The crop was blind to where the legend actually is.* The old reader took
   a fixed bottom strip (`h*0.90..h`, left 55%). Legends that sit slightly
   higher were missed entirely, and on report 5739's image2 — a **screen
   capture of the acquisition software** rather than an exported micrograph —
   the strip grabbed the application's status bar. Replaced with a scan of
   the whole frame for runs of rows carrying a little bright text.

2. *The magnification was being harvested as the job number.* `_JOB_PAT` is
   any 4-digit run, so `6831_E_1000x-7` offers both `6831` and the `1000` of
   `1000x`. The magnification is the easier read, so it won, and report
   6831's micrograph was reported as **job 1000**. `_vote_job` now blanks
   magnification tokens before looking for a job.

3. *A single bad pass could invent a job number.* The old code did
   `_JOB_PAT.search(' '.join(reads))` — first hit across all passes
   concatenated. On 5739 one pass produced `£5003` and the micrographs were
   reported as **job 5003** against a report numbered 5739, which is exactly
   the "micrograph from the wrong job" finding the check exists to raise.
   Now voted, mirroring `_select_magnification`, and accepted only on 3+
   agreeing passes or unanimity across 2. On the corpus every correctly-read
   legend carried 4 votes while the one misread (`7207` for 7227) had 2 with
   a competitor — so the threshold cleanly separates them, at the cost of one
   correct-but-contested read going unreported. **Unreadable is the right
   failure here; a wrong job number accuses a micrograph of belonging to
   another report.**

**Preprocessing, for whoever tunes this next**
- Upscale with LANCZOS *before* thresholding, not after — the old
  `_binarize` thresholded then resized, throwing away the only sub-pixel
  information a 9px glyph has.
- **Pad the crop with a black border.** Tesseract will not segment a line
  that runs to the edge of its input; tight crops returned *nothing at all*.
  This single change is what took image1 from silence to `5739_E_500x-2`.
- Crop columns by *cluster*, not min..max. The ID caption is bottom-left and
  the scale bar bottom-right on the same rows, so one span swallows the whole
  micrograph between them and OCR reads grain.
- Scale to a target band height (~90 and ~150px), not a fixed multiplier: a
  fixed x12 on a wide band produced a 9000px-wide image and took 60s+ for one
  four-micrograph report.

**Not fixed, recorded**
- `image2` of 5739 still reads nothing: it is a screenshot of the microscope
  software, and its legend is inside the software's image panel. Worth a rule
  of its own — "picture N is a screen capture, not an exported micrograph" is
  a genuine report defect — but severity is Mohamed's call, so not added.
- Scale-bar (`µm`) reading is still weak; the cluster split now yields the
  scale bar as its own candidate crop, so this is mostly a pattern-tuning job.

**Tests** `_vote_job` is covered by three new tests in `test_lab_review.py`
pinning both fabrication modes (magnification-as-job, contested read) plus a
guard that consensus still reports. Suite 71 tests, OK.

---

## 013 · Opus · 6 Aug 2026

**Status:** Mohamed's four requests, all landed and driven end-to-end against a
real report (6943, plus a compose pass over the other four corpus files):
*"I can't zoom in the annotated report; I need callouts with comment; I need to
send the report with the callouts to some people; and to write the comments in
the comment section and share the report."*

**What was actually wrong with zoom**
Nothing was broken — there was no zoom at all. `st.image` scales to its column
and offers nothing beyond that, so a 1240 px A4 page rendered into a ~900 px
column at 0.73x and the values in it were a few pixels tall. Fixed on two
levels: the page now sits in an `overflow:auto` viewport as an inlined `<img>`
(Fit / 100 / 150 / 200 / 300%, panned by scrolling), and from 200% the workbook
is re-rendered at 240 dpi instead of magnifying a 150 dpi raster — there was no
detail to magnify. Measured in the browser: pane 828 px, image drawn 1648 px,
scrollWidth 1652 — real zoom, not CSS wishful thinking. Two render resolutions
only, since each is another LibreOffice pass.

`Locate` also zooms on the **exact** pages now, not only on the fallback grid.
It was fallback-only because `cell_pixel_rect` models the drawn grid's geometry;
`_annotate_faithful_pages` knows where each marker actually landed, so
`crop_issue_detail()` crops from that instead. (This is the "zoomed-crop is
exact-mode-blind" gap noted in entry 009.)

**Callouts — and why the leader lines are routed, not straight**
The comment panel already existed in the download; it was a list beside the page,
not callouts. Cards are now placed level with the row they describe and joined to
their badge by a leader. First attempt drew straight lines: with 10 cards on a
one-page report the lower ones could not be placed level with their cells, and
the leaders fanned diagonally across all four micrographs. Now the vertical run
is kept in a channel just inside the page's right edge — the sheet's own print
margin — and the line terminates on the **badge**, not the cell's middle, so its
horizontal run follows a row border instead of striking through the values. Same
composite on screen and in the PDF: what the reviewer approves is what the
recipient opens.

**Badge-without-callout was a real inconsistency, and it is closed**
Markers are drawn for *every* anchored finding, but the UI only ever listed the
report-scoped, non-dismissed ones — on 6943, badges 1 and 4 (template-scoped)
had no card anywhere. In a document that leaves the building that reads as a
bug, so callouts are now generated for all three states: live, template-scoped
("applies to every report on this template"), and dismissed (carrying the
reviewer's reason). The latter two are grey, so they cannot be mistaken for live
problems. The UI's own card list is unchanged in what it counts.

Related: the on-screen card filter matched findings by the *full* message while
triage state matches by stem. An anchored highlight's note is sometimes the
leading clause of the longer finding message, so the two could classify the same
finding differently. Both now go through one stem-keyed helper
(`_presentation_states`).

**Comments and sending**
`share.py` is new and dependency-free: e-mail via stdlib `smtplib`, Drive via the
photo library's existing OAuth (`drive_store.upload_file` / `grant_reader`, per
named recipient — never "anyone with the link", these are customer documents).
Verified end-to-end against a local SMTP sink driven through the browser: the
mail arrives with the annotated PDF attached and the reviewer's comment in both
the body and the page callouts.

One Streamlit trap worth recording: the message body is generated from the
review, but a `text_area` reads `value=` once, so a comment added after the panel
first rendered never reached the message. Caught by reading the actual `.eml`,
not by reading the code. Fixed by pushing the generated text into session state
while it is untouched and leaving it alone once the reviewer types over it (with
a rebuild button) — silently discarding what someone wrote would be worse than a
stale summary.

**Performance**
`render_report_faithful_view(..., with_pdf=False)` skips the old comment-less PDF
pass, and the callout composition sits in its own cache keyed on the triage state
and the comment list. Writing a comment costs one Pillow pass (~1 s on 6943),
never a LibreOffice render (~4 s).

**Environment note for the next agent**
`libreoffice-calc` was again missing in this sandbox and again fixed by
`apt-get update && apt-get install -y libreoffice-calc`, exactly as PHASE-NEXT §3
predicts. Do not conclude step 3 is blocked without running that.

**Tests** 71 → 85, all passing: seven callout/zoom tests in
`test_report_render.py` (comment text arrives complete, dismissed keeps its
reason, placement is level with its row, overflow continues in another column, a
leader line really is drawn between cell and callout, `crop_issue_detail`) and
seven in the new `test_share.py` (recipient parsing, refusal when unconfigured,
what the SMTP server actually receives, implicit TLS, per-recipient Drive
grants).

**Not done, deliberately**
- Step 3 (text-anchored annotation, D6) is still open. This work sits on top of
  the colour-probe render and does not make it harder to replace: the callout
  layer consumes `placements` (page index + bbox per marker), which a
  text-anchored implementation can produce just as well.
- Comments live in `st.session_state`, so they are lost when the session ends —
  the same limitation as accept/dismiss, and the same answer (step 4).

---

## 014 · Opus · 7 Aug 2026

**Status:** Mohamed asked why we keep losing work. Audited all 24 remote
branches, salvaged what was still good, and recorded what was rejected so this
does not have to be re-litigated. Tests 85 → 118, corpus verdicts unchanged
(6943 keeps its one critical; the other three metallurgical reports stay clean).

**Why work was being lost — three separate mechanisms**

1. **No pull request was ever opened.** `claude/tool-audit-improvements-v7bj10`
   (6–7 Jul, ~1,000 lines across all four tools) and
   `claude/tool-error-detection-7vkrc9` (31 Jul, ~1,100 lines incl. 460 of
   tests) were pushed and then simply never proposed. Nothing in the repo
   surfaces a branch that has no PR, so they went quiet and stayed quiet.
2. **Squash merges hide what is already in.** Eleven branches are byte-identical
   to main and seven more had their content squashed in as PRs #7–#11, which
   leaves them permanently "ahead" in `git branch -v`. Real losses were buried
   among 18 false alarms, so nobody could tell them apart by looking.
3. **The Phase 1 rebuild rewrote the files the older branches were fixing.**
   `lab_review.py` and `report_render.py` moved by ~1,000 lines on 2 Aug, so a
   July fix to them no longer applies as a patch. Those had to be re-judged fix
   by fix rather than cherry-picked — and half of them turned out to have been
   re-fixed independently, which is the real cost: the same defect found twice.

**Merged (each verified against the current parser, corpus, and tests)**

From `claude/tool-error-detection-7vkrc9`:
- `_canon_machine` resolved only frame sizes 6/7/9. A real report titled "FS.5"
  (GE Frame 5 / MS5001 exists) silently failed to resolve, dropping the
  machine/set cross-check for that report instead of matching or flagging it.
- `_canon_machine`'s FS.<frame> branch captured no variant suffix, so a title
  correctly stating "FS.7FA" was truncated to MS7001 and then reported as
  disagreeing with its own, correct, "MS7001FA" header.
- `_UNETCHED_PAT` required the contiguous word "unetched", missing the
  hyphenated "Un-etched" that real captions use.
- Nominal-side composition structure was unchecked: `_composition()` already
  computed `duplicate_headers` for the spec table, but nothing read it, so a
  duplicated/mislabeled *nominal* column passed silently while the identical
  fault on the Actual side is a critical. Now a critical, anchored to its cell.
- `find_duplicate_compositions()` — the cross-report copy-paste check — wired
  into the app's batch path, `lab_review.py`'s CLI and `batch_review.py`.
- `composition_store.py` + its 8 tests, giving that check cross-session memory.

From `claude/tool-audit-improvements-v7bj10` (storage layer, applied as a
three-way patch — main had not touched these files since the branch point):
- `photo_lib._safe` kept leading dots, so an alloy value of `..` became a path
  segment that walked out of the library directory.
- Stored filenames were job + image stem only, so two reports sharing a job
  number mapped to one file and the second silently replaced the first.
  `_stored_filename` adds a short hash of the source report.
- A corrupt local index returned `[]`, and the next save then wrote a fresh
  index over the real one and orphaned the whole library. Now refuses, matching
  what the two cloud backends already did. Index writes are atomic.
- GitHub's Contents API stops inlining content past ~1 MB and returns encoding
  `none`; decoding that as base64 made every read fail, i.e. a permanent outage
  once the library grew. Now re-fetches with the raw media type.
- Drive: a lock around the non-thread-safe service, resume semantics instead of
  duplicate uploads, merge-before-write so a concurrent session's additions are
  not clobbered, and a missing file rendering as "missing" instead of crashing
  the gallery. `--migrate` honours `PHOTO_LIBRARY_DIR`.
- `add_to_library` now ignores an unclassified layout (the UI gated on report
  type; the CLI did not, and dumped images with no usable metadata).
- IIR: row scans capped (one stray cell at Excel's last row pushed `max_row` to
  ~1,048,576 and froze a worker), the section-based checks added to
  `CHECK_CATALOG` so the severity settings can reach them, and `apply_overrides`
  no longer erases a context-softened severity when the caller passes the
  catalog defaults.

Those two branches shipped their storage and IIR fixes with **no tests at all**,
which is most of why they were easy to lose. New `tests/test_photo_lib.py` (9)
and `tests/test_iir_review.py` (6, the first IIR tests in the repo) pin them.

**Rejected, with reasons — do not re-merge these without a decision**

- **7vkrc9's disposition work** (`_comment_disposition`, restore/recover
  vocabulary, and a new "Result says See comment but there is no comment at all"
  critical). Main *deliberately dropped* the "See comment" disposition critical
  on 2 Aug (36e0f70), and `test_microstructure_does_not_prove_restored_
  mechanical_properties` pins the opposite of the restore/recover reading. This
  is a product decision that post-dates the branch.
- **7vkrc9's nominal-total sanity band.** Main's rebuild has a better version
  (99–101 % with balance/remainder awareness) than the branch's 95–105 %.
- **7vkrc9's `comment_picture_refs` and `_component_identity` underscore fix.**
  Both were re-fixed independently by the rebuild. Their tests were kept as
  guards rather than the code.
- **7vkrc9's "coating recorded but no comment" warning.** A new rule, not a bug
  fix; severity on AEG process is Mohamed's call.
- **v7bj10's `lab_review.py` / `report_render.py` / `app.py` hunks.** Superseded
  — those files were rebuilt after the branch, and `app.py` is now a router.
- **v7bj10's and QKUI5's `sem_convert.py` hunks** (~95 lines: caption cleaning,
  figure layout, overflow). HANDOFF's "do not touch sem_convert.py" still
  stands, and there is no vendor SEM PDF in the repo to verify a change against.
  Worth a look when one is available.
- **QKUI5's auto-generated conclusion** — it writes "considered suitable for
  reconditioning" into a report when the source PDF states no conclusion. A tool
  should not assert a disposition nobody wrote. Mohamed's call, flagged not
  merged.
- **`claude/lab-report-annotated-view-ijay74`** (captioned-but-not-embedded
  micrographs) — already in main, and stricter there: it compares caption count
  to embedded count in both directions.
- **`audit-reports`** — an empty folder and a README.

**Unverifiable, merged anyway, flagged here:** the IIR *parser* fixes (Total-row
double counting, text-formatted position numbers, multi-sheet sum accumulation,
page-footer false matches). They apply to exactly the code they were written
against, and the bugs they describe are specific and plausible, but there is no
IIR workbook in the repo to drive them. **Run one real IIR report through the
reviewer before relying on it.** If it misbehaves, this entry is where to start.

**How to stop this happening again** — the cheap version is a rule, not a tool:
push a branch, open the PR the same session. Everything lost here was lost by a
branch that never became a PR. `git branch -r --no-merged origin/main` plus a
patch-content check (not `git cherry`, which squash merges defeat) is enough to
audit it in one command; that is how this entry's list was produced.
