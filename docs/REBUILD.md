# Extraction rebuild — specification

Durable spec for the Work-Power-Tools rebuild. Read this before touching
`lab_review.py`. The working protocol is in [HANDOFF.md](HANDOFF.md); the
running conversation between agents is in [FEEDBACK.md](FEEDBACK.md).

---

## 1. Why

The tool's review rules are good. Its extraction layer is not, and every
"it's not parsing reports correctly" complaint traces to one decision: fields
are located by scanning the entire worksheet for a regex label, then walking
right until a non-empty cell appears.

```python
def _find(ws, pattern, col=None, max_row=None):
    """Return (row, col) of the FIRST cell whose text matches `pattern`."""
```

That strategy has no idea what a label is, cannot tell a label from prose that
happens to contain the same words, stops at the first hit, only ever looks
rightward, and never records whether it succeeded. Ten reproducible defects
follow from it. 19 commits of fixes to `lab_review.py` have not converged,
because there was no measurement telling anyone whether a change helped.

**The rules are the asset — roughly 1,400 lines encoding real metallurgical
judgement gathered over many rounds of feedback on real reports, plus 33
regression tests each marking a defect somebody actually found. Do not rewrite
them. Move them onto a model that is worth trusting.**

## 2. Defect register

All eleven were reproduced by running the shipped parser — D1–D5 and D9–D10
against synthetic workbooks, D3 and D11 against the five real reports in
`corpus/`. Acceptance criteria are executable in
[`tests/test_parser_contract.py`](../tests/test_parser_contract.py) (synthetic)
and [`tests/test_corpus_regression.py`](../tests/test_corpus_regression.py)
(real files — **the authoritative one where the two disagree**). Read §2.1
before building: the real reports corrected the shape of D3 and added D11.

| ID | Defect | Consequence |
|----|--------|-------------|
| **D1** | No parse-confidence concept anywhere. A field that was not located and a field that is genuinely blank both return `None`. | Root cause. Every "field is blank" finding is unfalsifiable — it may be a report defect or a parser failure, and nothing distinguishes them. |
| **D2** | `_find()` returns the first regex match on the sheet, so one line of prose containing a field's words captures the label. | False positive. Adding `"Machine Type per customer drawing pack"` to a clean report makes it report the plainly-present `V94.3A` as missing. |
| **D3** | Only one sample is ever built. Real reports pack every sample into a *single cell*, whitespace/newline separated (see §2.1); `_sample()` reads that cell as one opaque value. | **Release-critical false negative.** Active on 3 of the 4 real reports. Samples 2…N never reach any of the 20 checks, so a rejection or a different alloy on a later sample is invisible. |
| **D4** | `_composition()` stops at the first `(Actual)` table. | False negative. A second sample's grossly off-spec chemistry (Ni 41% against a 61% nominal) is never read. |
| **D5** | `_num()` extracts the first digit-run from any string with no signal that the cell was not numeric. | Fabricated data. `'N/A (rev 2)'` → `2.0`, entering the review as a hardness measurement. `'see note 3'` → `3.0`. `'<0.01'` → `0.01`, losing below-LOD meaning. |
| **D6** | Cell anchoring paints each flagged cell a unique RGB, converts via LibreOffice, rasterises, then colour-matches pixels at ±2 tolerance to recover position. | Fragile and slow. 90s subprocess in the interactive path; breaks on pre-filled cells; colours collide past ~100 findings in one severity. Geometry is already available directly. |
| **D7** | All 33 tests either inject `parsed` dicts directly or build workbooks matching the parser's own assumptions. | No safety net. Zero coverage of the real failure surface; no test runs against a real report file. |
| **D8** | `detect_type()` is three regex probes; no template registry, fingerprinting or versioning. | Template drift shows up as mysterious wrong findings rather than an actionable "unrecognised template". |
| **D9** | The rightward scan does not stop at the next label, so a blank field adopts whatever is next on the row — an adjacent label, or worse, an adjacent field's **value**. | Most insidious shape of D1. `machine` silently becomes `'J-1001'`. Non-empty, so no completeness check fires, and the wrong value flows into title-identity and traceability as though correctly read. |
| **D10** | Header fields only ever scan rightward (`_value_right`). | A template variant that stacks the value beneath its label loses every header field at once. |
| **D11** | Findings have no scope. Template-level criticisms are re-emitted per report. | Measured on real files: a majority of a report's findings are identical on every report. Partially addressed — one always-true critical was dropped in FEEDBACK.md entry 004, which surfaced a genuine one. One constant critical remains ("no governing acceptance specification"). |

> D9 was found while writing the test for D2: the first version of that test
> *passed* because the field absorbed a neighbouring value instead of going
> blank. Negative assertions are not sufficient here — **always assert the
> expected value positively.**

### 2.1 What the real reports changed

Five real AEG reports were recovered from this repository's own git history
(committed early, removed when `*.xlsx` was gitignored). Recover them with
`python3 scripts/recover_corpus.py`; assertions live in
[`tests/test_corpus_regression.py`](../tests/test_corpus_regression.py).

Running against them corrected two things and found one new defect. **Every
correction below came from real files and none of it was visible in synthetic
fixtures** — which is the whole argument for keeping the corpus in the loop.

**D3 has a different shape than originally specified.** Real reports do not put
one sample per row. Every sample is packed into a *single cell*, whitespace and
newline separated. From report 6831, cell `B9`:

```
'MS 6369C        MS 6411C\nMS 6889C         MS 6931C'     ← 4 samples
'CD70356     CD70386\nCD70374  CD70385'                   ← 4 serials, cell G9
'Rene-80'                                                  ← 1 shared material
'See comment'                                              ← verdict deferred to prose
```

So D3 is not "read more rows", it is "tokenize one cell into N samples, carry
the shared fields onto each, and resolve the per-sample verdict out of the
comment text". `_identifier_tokens()` already tokenizes these correctly — it is
only ever used for *counting* in `_review_traceability`, never for building
sample records. Three of the four real metallurgical reports carry more than
one sample, so **D3 is active on 75% of the corpus.**

**The happy path parses better than expected.** On all four real reports the
job, machine, customer, material, composition tables and captions all extract
correctly. D2/D9/D10 are *latent* — they fire when a field is blank or when
prose collides with a label, not routinely. That lowers their urgency relative
to D3 and D11, though the template makes D9 easy to trigger: header fields sit
in pairs on one row (`B3:D3` label, `E3:H3` value, `I3:K3` label, `L3:O3`
value), so a blank Customer scans straight into `Machine Type:`.

**New — D11: findings have no scope.** Measured across the four real reports:

| | as audited | after entry 004 |
|---|---|---|
| Distinct findings across the corpus | 25 | 24 |
| Findings that fire on **all four** reports | **9** | **8** |
| Criticals that fire on all four reports | **2 of 2** | **1** |
| Reports whose critical count differs from the others | 0 | 1 |

The last row is the one that matters: dropping a single always-true check made
report 6943 visibly different from its peers for the first time.

The constant findings are not wrong — the template genuinely never states an
acceptance specification, and `Result` genuinely says "See comment" every time.
They are criticisms of the **template**, re-emitted per report.

The cost is concrete. Report 6943 has a real third critical: serial
`C1ZP 093046` appears twice in one cell. That is exactly the kind of find the
tool exists for, and it arrives looking identical to the two criticals that
fire unconditionally. **A critical that cannot be absent carries no
information, and it camouflages the ones that can.**

Findings need a `scope` dimension — `report` vs `template` — with
template-scoped observations stated once, outside the per-report finding list.
This is likely the single biggest perceived-quality win available, and it is
independent of the extraction rebuild.

**Also observed, not yet a defect:** column `T` of the metallurgical sheet holds
data-validation dropdown sources — machine types (`MS 3002`…`MS 9001`),
sign-off names, material lists — inside the same worksheet. `_find()` scans
every column, and `_workbook_text()` folds column T into the document text used
by the evidence and document-control checks. No wrong finding has been traced
to this yet, but any whole-sheet scan should exclude the dropdown region.

## 3. Target architecture

```
ingest → identify template → extract → VALIDATE ⛔ → review → present
```

The load-bearing addition is stage 4. Extraction incompleteness is reported as
extraction incompleteness, never as a report defect. No rule runs against a
field the extractor could not resolve.

## 4. Document model

The contract tests assert this shape. **If you adopt a better one, change the
tests in the same commit and record the deviation in FEEDBACK.md — do not
silently reinterpret them.**

```python
parsed = {
    # ── canonical ──────────────────────────────────────────────────────
    'samples': [                      # ALWAYS a list, even for one sample
        {'sample_no', 'description', 'serial', 'location', 'material',
         'result', 'nominal', 'actual', 'hardness', 'cells': {...}},
    ],
    'fields': {                       # every scalar field, with provenance
        'machine': {
            'value':      'V94.3A',
            'raw':        'V94.3A',
            'cell':       'C5',
            'status':     'found',    # 'found' | 'empty' | 'not_located'
            'confidence': 1.0,        # 0.0–1.0
        },
    },
    'composition': {
        'nominal_tables': [Table, ...],
        'actual_tables':  [Table, ...],
    },

    # ── back-compat views, derived from the canonical data ─────────────
    # Keep these through Phases 1 and 2 so the 33 existing rule tests and
    # everything in ui/ keep working. Retire them in Phase 2 as each rule
    # moves onto the canonical model.
    'header':  {...},   # header[k]  == fields[k]['value']
    'sample':  {...},   # sample     == samples[0]
    'nominal': {...},   # nominal    == composition['nominal_tables'][0]['values']
    'actual':  {...},
    'loc':     {...},
}

Table = {
    'label_cell', 'header_row', 'value_row',
    'values':  {'Ni': 60.8, ...},
    'entries': [{'element', 'raw', 'value', 'header_cell', 'value_cell'}, ...],
    'duplicate_headers': [...],
}
```

`status` semantics — this is the D1 fix and the whole point of the exercise:

- `found` — label located, value cell holds a usable value
- `empty` — label located, value cell genuinely blank → *a report defect*
- `not_located` — label not found → *an extraction failure, never a finding
  about the report*

## 5. Phase 1 — scope

Ship behind the existing Streamlit UI. **No interface changes.** The point is
to prove false positives collapse before anything visual moves.

| # | Work | Closes |
|---|------|--------|
| 1a | Merge-aware geometric grid reader. Label↔value association that stops at the next label and searches right **and** below. | D2, D9, D10 |
| 1b | Field records carrying `value` / `raw` / `cell` / `status` / `confidence`. | D1 |
| 1c | Multi-sample extraction — `samples` as a list. | D3 |
| 1d | All composition tables, not just the first. | D4 |
| 1e | `_num()` hardening, plus a distinct not-quantified signal for `<LOD` forms. | D5 |
| 1f | Back-compat views so the 33 existing tests and `ui/` keep passing untouched. | — |
| 1g | Add a `scope` field (`report` \| `template`) to findings and stop emitting template-scoped ones per report. Independent of the extraction work — do it first if you want the fastest visible win. | D11 |

**Not in Phase 1:** template registry (D8 — needs real report files, see §7),
the annotation rebuild (D6), the UI, FastAPI, storage. Phases 2–5 are in the
audit; do not start them without checking in via FEEDBACK.md.

## 6. Acceptance criteria

Phase 1 is done when all of these hold:

1. Every `@unittest.expectedFailure` in `tests/test_parser_contract.py` has
   been removed and all 15 tests pass.
2. All 33 tests in `tests/test_lab_review.py` pass **unmodified**. If a rule
   test must change, that is a signal you altered rule behaviour — stop and
   raise it in FEEDBACK.md first.
3. The three guard tests in the contract file still pass (they cover the common
   single-sample / single-table / decimal-parsing paths).
4. `python3 lab_review.py <report.xlsx>` still runs and produces findings.
5. The Streamlit app still runs with no behavioural change beyond fewer
   spurious findings.
6. `python3 scripts/recover_corpus.py && python3 -m unittest
   tests.test_corpus_regression` passes, with the four real-file extraction
   tests still green. **Run this before and after every change** — it is the
   only real-file signal in the project.

## 7. Boundaries

- **Do not touch `sem_convert.py`.** Different job (PDF→Word), not implicated
  in any complaint, reviewed structurally only.
- **Do not rewrite the rule bodies.** That is Phase 2 and it is a careful,
  one-rule-at-a-time port with tests carried across.
- **Do not delete existing tests.** They each mark a real defect found on a
  real report.
- **A real corpus exists — use it.** Five real reports (four metallurgical,
  one coating) are recoverable from git history via
  `python3 scripts/recover_corpus.py`. They are customer documents: they land
  in `corpus/`, which is gitignored, and they must **never** be committed. Do
  not rename them — the title-identity checks read the filename.
  Five reports is thin: it covers one template family well and says nothing
  about template drift, revisions, or the coating family beyond a single file.
  Broadening it is still worth doing, but it is no longer a blocker.
- **Do not guess on metallurgy.** Tolerances, hardness ranges and severity
  calls are domain decisions. Raise them in FEEDBACK.md.

## 8. Full audit

Narrative version, with the probe transcripts and the module-by-module
disposition table:
<https://claude.ai/code/artifact/8c1ad4da-1c8b-4848-b4a4-cbab3ccb21be>
