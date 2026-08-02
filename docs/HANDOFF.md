# Working protocol

How agents pick up this work and hand it back. Spec is in
[REBUILD.md](REBUILD.md); the running log is [FEEDBACK.md](FEEDBACK.md).

---

## Start here

```bash
pip install -r requirements.txt
python3 scripts/recover_corpus.py             # real reports -> corpus/ (gitignored)
python3 -m unittest discover -s tests -v      # baseline BEFORE you change anything
```

Expected baseline, with `requirements.txt` installed and the corpus
recovered: **68 tests, 3 expected failures, rest passing** — the three are the
D11 scope tests, which are step 1 of your assignment.

**Current assignment: [PHASE-NEXT.md](PHASE-NEXT.md).** Phase 1 (extraction) is
done and merged; that file supersedes REBUILD.md §5 for what to do now.

**Recover the corpus first.** `tests/test_corpus_regression.py` runs against
five real AEG reports and skips silently without them — and synthetic fixtures
did not catch one single defect that the real files caught. Never commit or
rename those files: they are customer documents, and the title-identity checks
read the filename.
`tests/test_report_render.py` needs PyMuPDF (`fitz`) — it is in
`requirements.txt`, so install those first or you will see an import error that
is not yours. The annotated *view* additionally needs the LibreOffice Calc
import filter, which is a separate thing and is missing in some environments —
see PHASE-NEXT.md §3 before relying on it.

Then read, in order:

1. `docs/PHASE-NEXT.md` — the current assignment, in order
2. `docs/REBUILD.md` §2 (the defect register) and §4 (the document model)
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
- Whatever `PHASE-NEXT.md` currently scopes, in the order it gives. Do not
  start anything past it without raising it in FEEDBACK.md first.
- Do not touch `sem_convert.py`.
- Phase 1 is finished; its "no UI changes" rule no longer applies — the UI is
  now the assignment.

**Tests**
- Never delete or weaken an existing test. Each marks a real defect found on a
  real report.
- If a test in `test_lab_review.py` starts failing, you changed rule behaviour.
  Stop and raise it — do not adjust the test to match your code.
- Add a contract test for any *new* defect you find, following the existing
  pattern, and add it to the register in REBUILD.md §2.
- Where `test_parser_contract.py` (synthetic) and `test_corpus_regression.py`
  (real files) disagree about a report's shape, **the corpus wins.** The
  synthetic fixtures were written before the real reports were recovered and
  one of them already encoded a layout that does not exist.

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
