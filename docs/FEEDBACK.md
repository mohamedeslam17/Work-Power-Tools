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
