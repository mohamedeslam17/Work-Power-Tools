# Next assignment — scope, then the UI

Phase 1 (extraction) is done and merged. This is what comes after it, in order.
Spec context is [REBUILD.md](REBUILD.md); protocol is [HANDOFF.md](HANDOFF.md);
the log is [FEEDBACK.md](FEEDBACK.md).

**Read this whole file before starting.** Step 2 is the one Mohamed actually
cares about — the UI was his original complaint and nothing has touched it yet.
Steps 1 and 3 exist because they make step 2 better, not the other way round.

---

## Where things stand

| | at audit | now |
|---|---|---|
| Findings constant across all 4 real reports | 9 | 6 |
| Criticals constant across all 4 | 2 | **0** |
| Reports with no critical | 0 | **3** |
| Reports with a genuine critical | — | **1** (6943, duplicated serial) |

Extraction is fixed. The remaining noise is six *warnings/infos*, and the UI has
not been touched at all.

---

## Step 1 · Close D11 — finding scope

**Why:** six findings still fire on every report. None is wrong; each says the
*template* has no field for something (test method, controlled report number,
sequential page numbering, a sampling plan, mismatched element sets between the
nominal and actual tables). Unlike the two criticals Mohamed had removed, these
should **not** be deleted — they are real gaps worth fixing at source. They
should be said **once, about the template**, not on every report forever.

**Target API** — declarative, in one place, no rule-body changes:

```python
lab_review.TEMPLATE_SCOPED                       # ((category, stem_pattern), ...)
lab_review.partition_by_scope(findings)          # -> (report_findings, template_findings)
```

Classify by declaration, not by measuring constancy — `review_report()` reviews
one file and has no cross-report context. The constancy measurement in the tests
is how you *check* your declarations, not how you make them.

Careful: `Composition` has findings on **both** sides — the two constant infos
("nominal trace elements not reported in Actual", "reported but not in nominal
spec") are template-scoped, while the deviation warnings are report-scoped. So
category alone is not sufficient; match on category **and** message stem.

**Acceptance** — three `@unittest.expectedFailure` tests in
`tests/test_corpus_regression.py::FindingScopeTests`:

- `test_scope_partitioning_exists`
- `test_no_report_scoped_finding_is_constant_across_the_corpus`
- `test_template_findings_are_identical_across_reports`

plus the existing guard `test_the_genuine_critical_stays_report_scoped`, which
must keep passing — 6943's duplicated serial must never end up template-scoped.

Small step. Do it first; step 2 consumes its output.

---

## Step 2 · Rebuild the Lab Review UI around triage

**This is the assignment.** Everything above is groundwork.

### Stay on Streamlit for this step

This revises the *sequencing* of the original audit recommendation, not its
conclusion. The audit said replace Streamlit, and that is still right if AEG
needs durable review records, multiple reviewers, or an audit trail. But:

- Most of the perceived-quality win does not need a new framework.
- `st.session_state` persists within a session, so accept/dismiss works today.
  What Streamlit cannot do is persist a decision *across* sessions and users —
  which matters only once someone wants a signed review record.
- Building the interaction model in the cheap medium first tells us what the UI
  should actually be, which de-risks a rewrite instead of guessing at it.

Do not start a framework migration in this step. Step 4 is where that gets
decided, with something real to point at.

### What the screen must do

The reviewer's job is triage: is this report releasable, and if not, where
exactly. Build for that.

1. **Verdict first.** One line at the top: release / hold / needs correction,
   and the single reason. Today the user meets a stack of expanders and has to
   assemble the verdict themselves.
2. **Two linked panes.** Findings beside the report. Selecting a finding scrolls
   the report to its cell; selecting a cell filters to its findings. The link is
   the product — a findings list and a report picture that do not talk to each
   other is what exists now.
3. **Template findings, once.** Step 1's `template_findings` go in their own
   collapsed section — "About this template", stated once, not mixed into the
   per-report list.
4. **Per-finding actions.** Accept / dismiss-with-reason, held in
   `st.session_state`, keyed by `(filename, category, stem)` so it survives
   reruns. Dismissed findings drop out of the count and the verdict. Show a
   "3 dismissed" affordance to restore them.
5. **Surface extraction confidence.** Phase 1 gave every field a `status` of
   `found` / `empty` / `not_located`. A `not_located` field must *look* different
   from an empty one — the whole point of D1 was that the tool stops accusing
   the report of something that is really a parser miss. Nothing in the UI shows
   this yet.
6. **Dark theme.** `ui/theme.py` and `.streamlit/config.toml` are hardcoded
   light, on a palette borrowed from Supabase. Both themes, and a palette that
   belongs to materials inspection rather than to a developer tool.

### Constraints

- `ui/lab_tool.py` is 509 lines and is where most of this lands. `iir_tool.py`
  and `photo_tool.py` should follow the same shell once Lab is proven — do not
  do all four at once.
- Do not touch `sem_convert.py`.
- Keep `_cached_review` / `_cached_annotated_report` caching intact. Re-parsing
  on every interaction is what made the old UI feel slow.
- No new runtime dependencies without raising it in FEEDBACK.md first.

### Acceptance

Not unit-testable, so this is a checklist — verify each against a **real
corpus report**, not a synthetic one, and say in your FEEDBACK entry which
report you drove:

- [ ] Verdict is legible without scrolling or opening anything
- [ ] Clicking a finding moves the report view to its cell
- [ ] Template findings appear once, in their own section
- [ ] A dismissed finding leaves the count and the verdict, and can be restored
- [ ] A `not_located` field is visually distinct from an `empty` one
- [ ] Readable in both light and dark
- [ ] 6943 shows its duplicated-serial critical prominently; the other three
      real reports show no critical at all

That last one is the end-to-end check that everything from Phase 1 onward still
holds together.

---

## Step 3 · Annotation without the colour round-trip (D6)

**Check this before planning around it:** the annotated view needs
`libreoffice-calc`, and it is **not installed in every environment**. In the
session where this plan was written, `soffice` was present but the Calc import
filter was not, so `render_report_faithful_view()` returned
`RuntimeError: LibreOffice conversion failed` on a real report — and it fails
the same way on a trivial one-cell workbook, which is how it was diagnosed as
environmental rather than a code defect. `packages.txt` installs
`libreoffice-calc`, so a correct deployment has it.

```bash
python3 -c "import report_render as r; print(r.libreoffice_available())"
soffice --headless --convert-to pdf:calc_pdf_Export --outdir /tmp <any>.xlsx
```

If the filter is missing and you cannot install it, **do steps 1 and 2 and stop**
— do not attempt step 3 blind. Say so in FEEDBACK.md.

**Worth stating plainly:** the colour-probe approach is a reasonable answer to a
genuinely hard problem. Mapping a spreadsheet cell to a coordinate on a rendered
page means modelling LibreOffice's pagination, and painting a findable colour
sidesteps that entirely. It is not stupid; it is just fragile, slow and lossy.

**Proposed replacement — anchor by text, not by colour:**

1. Render the **unmodified** workbook to PDF. (Today the workbook is mutated
   with fills before rendering, so the "pixel-faithful" view is not actually
   faithful — it is the report with highlight colours painted into it.)
2. For each flagged cell, take its text and locate it with PyMuPDF's
   `page.search_for()`. Pagination is handled for free.
3. Draw the highlight and badge as **vector** annotations on the PDF.

This keeps text selectable and searchable, removes the second render pass, and
removes the colour-collision ceiling. Disambiguate duplicate text hits by
reading order and proximity to already-anchored cells; fall back to the existing
colour probe for a cell whose text is empty or genuinely ambiguous — do not
delete that path, demote it.

**Acceptance:** every anchored finding in `collect_highlights()` lands on the
correct cell for all four real metallurgical reports, the output PDF has
selectable text, and no cell fill is added to the rendered workbook.

---

## Step 4 · Decide on Streamlit — do not start this

Once step 2 exists, the question becomes answerable with evidence rather than
opinion. Leaving Streamlit is worth it if AEG needs review decisions to persist
across sessions, more than one reviewer, or a signed review record. It is not
worth it for looks alone.

Write your recommendation into FEEDBACK.md with what you learned building
step 2. **Mohamed decides.** Do not begin a migration.

---

## Ground rules for this assignment

Everything in [HANDOFF.md](HANDOFF.md) still applies. Additionally:

- **Recover the corpus before you start** and drive every UI change against a
  real report. Synthetic fixtures have now twice failed to catch things the real
  files caught immediately — most recently label patterns that made all four
  reports come back `not_located`.
- **Do not guess on metallurgy or on AEG process.** Two checks have already been
  removed because the tool was demanding something AEG reports are not meant to
  carry, and no test suite could have told us that. It needed Mohamed. If you
  hit another, ask in FEEDBACK.md and carry on with the rest.
- **Assert positively.** Still the rule that keeps catching people.
- Steps are independently shippable. Commit and push each rather than landing
  all three at once.
