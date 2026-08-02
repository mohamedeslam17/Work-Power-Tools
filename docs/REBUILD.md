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

All ten were reproduced against the shipped parser. D1–D5 and D9–D10 are
encoded as executable acceptance criteria in
[`tests/test_parser_contract.py`](../tests/test_parser_contract.py).

| ID | Defect | Consequence |
|----|--------|-------------|
| **D1** | No parse-confidence concept anywhere. A field that was not located and a field that is genuinely blank both return `None`. | Root cause. Every "field is blank" finding is unfalsifiable — it may be a report defect or a parser failure, and nothing distinguishes them. |
| **D2** | `_find()` returns the first regex match on the sheet, so one line of prose containing a field's words captures the label. | False positive. Adding `"Machine Type per customer drawing pack"` to a clean report makes it report the plainly-present `V94.3A` as missing. |
| **D3** | `_sample()` reads exactly one row beneath the `Sample nr` header. | **Release-critical false negative.** A 3-sample report where sample 2 is a `REJECT` reviews clean; samples 2 and 3 never reach any of the 20 checks. |
| **D4** | `_composition()` stops at the first `(Actual)` table. | False negative. A second sample's grossly off-spec chemistry (Ni 41% against a 61% nominal) is never read. |
| **D5** | `_num()` extracts the first digit-run from any string with no signal that the cell was not numeric. | Fabricated data. `'N/A (rev 2)'` → `2.0`, entering the review as a hardness measurement. `'see note 3'` → `3.0`. `'<0.01'` → `0.01`, losing below-LOD meaning. |
| **D6** | Cell anchoring paints each flagged cell a unique RGB, converts via LibreOffice, rasterises, then colour-matches pixels at ±2 tolerance to recover position. | Fragile and slow. 90s subprocess in the interactive path; breaks on pre-filled cells; colours collide past ~100 findings in one severity. Geometry is already available directly. |
| **D7** | All 33 tests either inject `parsed` dicts directly or build workbooks matching the parser's own assumptions. | No safety net. Zero coverage of the real failure surface; no test runs against a real report file. |
| **D8** | `detect_type()` is three regex probes; no template registry, fingerprinting or versioning. | Template drift shows up as mysterious wrong findings rather than an actionable "unrecognised template". |
| **D9** | The rightward scan does not stop at the next label, so a blank field adopts whatever is next on the row — an adjacent label, or worse, an adjacent field's **value**. | Most insidious shape of D1. `machine` silently becomes `'J-1001'`. Non-empty, so no completeness check fires, and the wrong value flows into title-identity and traceability as though correctly read. |
| **D10** | Header fields only ever scan rightward (`_value_right`). | A template variant that stacks the value beneath its label loses every header field at once. |

> D9 was found while writing the test for D2: the first version of that test
> *passed* because the field absorbed a neighbouring value instead of going
> blank. Negative assertions are not sufficient here — **always assert the
> expected value positively.**

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

## 7. Boundaries

- **Do not touch `sem_convert.py`.** Different job (PDF→Word), not implicated
  in any complaint, reviewed structurally only.
- **Do not rewrite the rule bodies.** That is Phase 2 and it is a careful,
  one-rule-at-a-time port with tests carried across.
- **Do not delete existing tests.** They each mark a real defect found on a
  real report.
- **The golden corpus cannot be built in this repo.** `*.xlsx` is gitignored
  and no real reports are present, by design — they are customer documents.
  Collecting real reports and having an engineer adjudicate the expected
  findings is the highest-value action available, and it is a **human task for
  Mohamed**, not an agent task. Work against synthetic fixtures until it
  exists, and treat that as a known limitation rather than a solved problem.
- **Do not guess on metallurgy.** Tolerances, hardness ranges and severity
  calls are domain decisions. Raise them in FEEDBACK.md.

## 8. Full audit

Narrative version, with the probe transcripts and the module-by-module
disposition table:
<https://claude.ai/code/artifact/8c1ad4da-1c8b-4848-b4a4-cbab3ccb21be>
