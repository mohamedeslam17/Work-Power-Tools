#!/usr/bin/env python3
"""Recover the real-report corpus from this repository's git history.

Five real AEG reports were committed early in the project's life and later
removed when ``*.xlsx`` was gitignored. The blobs are still reachable from
history, so the corpus can be reconstituted on demand without ever putting
customer documents back into the working tree.

    python3 scripts/recover_corpus.py

Writes to ``corpus/`` (gitignored). Safe to re-run; existing files are skipped.

Why this exists
---------------
Every test in ``tests/test_lab_review.py`` and
``tests/test_parser_contract.py`` runs against synthetic workbooks that were
written to match what the parser expects. That is how the extraction defects
went unnoticed for 19 commits. These five files are the only real evidence in
the project of what an AEG report actually looks like, and running against
them found things no synthetic fixture did — see docs/REBUILD.md §2.1.

Do NOT commit the recovered files. They are customer documents.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
DEST = REPO / "corpus"

# blob sha -> filename. Reachable from history; they survive a fresh clone.
BLOBS = {
    "57f4542e9c94f5126787221eaf73f4e98d209209":
        "R 6831 AEN Saudi V84.2 2nd Stage Vane  Metallurgical Report - AEG (final).xlsx",
    "1817612e2859881ba4216c212768e8f91d95d00c":
        "TBR 6943 AEN Saudi FS.7 1st Stage Bucket  Metallurgical Report - AEG (Final).xlsx",
    "f6abd865dd1fc9a5c387111a9806b4257182ba88":
        "R 7227 AEN Saudi FS.7 2nd Stage Bucket Metallurgical Report - AEG.xlsx",
    "49205b0e543eca9dad87e334e47794e64cc5b069":
        "R 7253 PSM Thomassen FS.6 Transition Piece  Metallurgical Report - AEG.xlsx",
    "64936d68bb443499ab089673cf93661d64ff9c95":
        "6983 TSME FS.6 2nd Stage Bucket Aluminizing Coating Report - AEG.xlsx",
}


def recover(dest: Path = DEST) -> int:
    """Materialise the corpus. Returns the number of files written."""
    dest.mkdir(parents=True, exist_ok=True)
    written = 0
    for sha, name in BLOBS.items():
        target = dest / name
        if target.exists():
            print(f"  skip   {name}")
            continue
        try:
            blob = subprocess.run(
                ["git", "cat-file", "blob", sha],
                cwd=REPO, capture_output=True, check=True).stdout
        except subprocess.CalledProcessError:
            print(f"  MISSING {sha[:10]} — {name}", file=sys.stderr)
            continue
        target.write_bytes(blob)
        written += 1
        print(f"  write  {name}  ({len(blob):,} bytes)")
    return written


def main() -> int:
    print(f"Recovering real-report corpus into {DEST}/")
    written = recover()
    have = len(list(DEST.glob("*.xlsx")))
    print(f"\n{written} written, {have} present of {len(BLOBS)} expected.")
    if have < len(BLOBS):
        print("Some blobs were unreachable — the history may have been rewritten.",
              file=sys.stderr)
        return 1
    print("\nFilenames are load-bearing: the title-identity checks compare the "
          "filename against report content, so do not rename these.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
