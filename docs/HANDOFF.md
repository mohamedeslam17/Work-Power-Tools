# Working protocol

How agents pick up this work and hand it back. Spec is in
[REBUILD.md](REBUILD.md); the running log is [FEEDBACK.md](FEEDBACK.md).

---

## Start here

```bash
pip install -r requirements.txt
python3 -m unittest discover -s tests -v      # baseline BEFORE you change anything
```

Expected baseline: **33 passed, 12 expected failures, 3 guards passed.**
`tests/test_report_render.py` needs PyMuPDF (`fitz`); if it is not installed
that error is pre-existing and not yours.

Then read, in order:

1. `docs/REBUILD.md` §2 (what is broken) and §4 (the model you are building)
2. `tests/test_parser_contract.py` — your acceptance criteria, executable
3. `docs/FEEDBACK.md` — the last entry tells you where things stand

## The expected-failure mechanic

Open defects are marked `@unittest.expectedFailure`, so the suite stays green
while work is in progress.

**When you fix a defect its test starts passing, and unittest reports an
UNEXPECTED SUCCESS, which fails the run.** That is the signal, not a problem.
Delete the decorator for that test in the same commit as the fix.

A red suite saying `unexpected successes=1` means you finished something.

## Ground rules

**Scope**
- Phase 1 only, as scoped in REBUILD.md §5. Do not start Phase 2+ without
  raising it in FEEDBACK.md first.
- Do not touch `sem_convert.py`, the rule bodies, or `.streamlit/`.
- No UI changes in Phase 1. The point is to prove the extraction fix in
  isolation.

**Tests**
- Never delete or weaken an existing test. Each marks a real defect found on a
  real report.
- If a test in `test_lab_review.py` starts failing, you changed rule behaviour.
  Stop and raise it — do not adjust the test to match your code.
- Add a contract test for any *new* defect you find, following the existing
  pattern, and add it to the register in REBUILD.md §2.

**Assertions**
- Always assert expected values positively. A negative assertion ("no finding
  was emitted") can pass because the parser produced something worse. That
  exact trap produced D9 — see the note at the end of REBUILD.md §2.

**Domain**
- Do not guess on metallurgy. Tolerances, hardness ranges, severity calls and
  anything touching `HARDNESS_REF` or `lab_vocab.py` are decisions for Mohamed.
  Ask in FEEDBACK.md and continue with the parts that do not depend on the
  answer.

**Commits**
- One defect per commit where practical, message naming the ID: `Fix D3:
  parse every sample row, not just the first`.
- Branch: `claude/tool-audit-rebuild-2kqxt2`. Push with
  `git push -u origin claude/tool-audit-rebuild-2kqxt2`.
- Do not open a PR unless Mohamed asks.

## Handing back

Append an entry to `docs/FEEDBACK.md` before you finish. Newest at the bottom.

```markdown
## NNN · <author> → <recipient> · <date>

**Status:** <what phase/item, and done | partial | blocked>

**Done**
- …

**Found**
- Anything the spec got wrong, or new defects. Include a reproduction.

**Questions**
- Blocking ones marked BLOCKING. Domain questions go to Mohamed, not to the
  other agent.

**Next**
- What you would pick up next, and why.
```

Rules for the log: append only, never edit someone else's entry, and always
record a deviation from the spec rather than quietly absorbing it. If the spec
is wrong, say so — it was written from probes against synthetic fixtures, not
from real reports, and it will have gaps.
