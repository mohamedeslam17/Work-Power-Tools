# Work-Power-Tools

Materials-engineering tools for Ansaldo Energia (AEG), packaged as a single
Streamlit app with four tabs.

## Tools

### 🔬 SEM Report Converter
Ingests vendor **SEM PDFs** and generates formatted Ansaldo Word reports
(cover, table of contents, micrograph pages, γ′ summary, conclusion).
Implemented in [`sem_convert.py`](sem_convert.py).

### 🧪 Lab Report Review *(new)*
Rule-based QA review of AEG materials-lab **Excel reports**. Runs fully
offline and flags issues with a plain-English reason for each finding.
Implemented in [`lab_review.py`](lab_review.py).

Two report families are recognised automatically:

| Family            | Checks performed |
|-------------------|------------------|
| **Metallurgical** | Filename ↔ content (job / type / component / customer); Actual-vs-Nominal composition (element-by-element, ±10% → warning, ±25% → fail); hardness (pre/post-solution sanity + built-in alloy reference); completeness; sign-off; coating cell ↔ comment (presence & type); caption integrity (duplicate/gap numbers, etch status required, comment over-references); comment ↔ material/Result verdict; micrograph legends, etch contrast and burned-in thickness. |
| **Coating**       | Filename ↔ content; coating-thickness measurements vs design MIN/MAX limits; sign-off; reference-micrograph presence. |

Findings are graded **🔴 Fail / 🟠 Warning / 🔵 Note / 🟢 Pass** and shown on screen.

### 🖼️ Photo Library *(new)*
Extracts the embedded micrographs from a reviewed report into a **per-alloy**
folder structure with a JSON index, and serves an in-app gallery: pick an alloy
→ see its micrographs with the data of the report they came from (job, magnification,
etch state, source, any thickness measurements). Implemented in
[`photo_lib.py`](photo_lib.py). Add to it via the button in the review tab, or the CLI:

```bash
python3 photo_lib.py "path/to/report.xlsx" [more.xlsx ...]
```

**Persistent storage.** Streamlit Community Cloud wipes the local filesystem on
reboot, so runtime additions need a cloud backend. The library is pluggable and
auto-selects by what's configured:

* **Google Drive** *(recommended — 15 GB+, no repo bloat)* — OAuth acting as you
  ([`drive_store.py`](drive_store.py)), least-privilege `drive.file` scope, into a
  self-managed **"AEG Photo Library"** folder in your Drive. Setup:
  1. Google Cloud console → enable the **Drive API** → create an **OAuth client
     (Desktop)**; add yourself as a test user on the consent screen.
  2. `python3 drive_store.py --auth` (any machine/Colab with a browser) → prints a
     refresh token.
  3. Add to Streamlit secrets:
     ```toml
     drive_client_id     = "....apps.googleusercontent.com"
     drive_client_secret = "...."
     drive_refresh_token = "1//...."
     ```
  4. `python3 drive_store.py --migrate` pushes the seeded local library up to Drive.
* **GitHub** *(alternative — no IT, browser-only)* — commits micrographs into the
  repo via a fine-grained PAT (Contents: read & write). Set `github_token`,
  `github_repo`, `github_branch`, `github_base`. Note: repos suit only small
  libraries (≈1–5 GB practical limit).
* **Local** *(default / fallback)* — a folder (`PHOTO_LIBRARY_DIR`, default
  `photo_library/`); **not tracked in git** (customer micrographs are kept out of
  the repo), and runtime additions don't persist across reboots — use the Drive
  backend for persistence.

**Built-in hardness reference.** `HARDNESS_REF` in [`lab_review.py`](lab_review.py)
holds typical *aged-condition* hardness (HRC) for common Ni- and Co-based
gas-turbine superalloys (IN738, GTD-111/741, René 80, Nimonic/C263, FSX-414,
X-40, IN718, …). The reviewer surfaces the alloy's reference range and notes
that **post-solution** readings are expected to run *below* it (the solution-
treated state precedes re-aging), so those are informational, not failures.
Values are advisory — verify against the controlling spec.

**Micrograph analysis (OCR).** With the *Analyse micrographs* option on, the
reviewer reads each embedded micrograph's burned-in legend (`<job>_E_<mag>x-<n>`
+ scale bar) and cross-checks the magnification and job number against the
captions; gauges etched-vs-low-contrast via edge density (advisory — faint
post-HT etching reads as low-contrast); and reads burned-in thickness labels
(e.g. `42 µm`) to surface alongside the comment's thickness values. Best-effort;
needs the **Tesseract** engine (`packages.txt` installs `tesseract-ocr`; the app
degrades gracefully without it).

### 🛠️ IIR Review *(new)*
Automated consistency/completeness QA of **Incoming Inspection Reports** (Detailed
Assessment Customer Reports) delivered as `.xlsx`. Upload one or more workbooks for a
severity-tagged findings checklist plus an on-screen batch overview. Implemented in
[`iir_review.py`](iir_review.py).

**Two report layouts are recognised automatically:** the classic *Contents /
Summary of Received (or Reconditioned) Parts / Serial Number* template, and the
section-based *Introduction / CONFIGURATION / SN registration / Incoming Photos*
template used by most reports. Identity, the serial registration and photos are
read from whichever layout a workbook uses; an unrecognised workbook is flagged
("unrecognized layout") instead of being silently mis-scored.

Checks span **Identity/metadata**, **Quantities** (Received = Scrap + Reconditionable,
positions = received, serial-scope totals reconcile, sum-row vs marked scopes,
Received-Parts table vs Serial-Number protocol), **Integrity** (unique/contiguous
positions, serial numbers, valid repair-scope L/M/H/S, scrap ↔ scope 'S'),
**Consistency** (Summary-of-Damages counts vs protocol marks, executive-summary
cross-checks), **Completeness** (a photo per caption, page numbering) and **Spares** —
the damage-driven *Expected Replacement Components* matrix tallied per component and
reconciled to the serial protocol (position coverage + scrap), plus the consumables
*Spare Parts List*. Each check's
severity is tunable (🔴 Fail / 🟠 Warn / 🔵 Info / ⚪ Off) live in the UI; defaults live in
`iir_review.CHECK_CATALOG`. Review several at once for a combined **Batch Summary** workbook.

```bash
python3 iir_review.py "report.xlsx"     # one report  → findings checklist
python3 iir_review.py *.xlsx             # many reports → checklists + batch summary
```

## Running

```bash
pip install -r requirements.txt
streamlit run app.py
```

The Lab Report Reviewer can also be run from the command line:

```bash
python3 lab_review.py "path/to/report.xlsx" [more.xlsx ...]
```

## Notes

* Composition tolerances and the advisory hardness ranges are constants at the
  top of [`lab_review.py`](lab_review.py) (`COMP_WARN_REL`, `HARDNESS_REF`, …) —
  adjust them to match your controlling specification.
* Raw customer report `.xlsx` files are **not tracked in git** (`*.xlsx` is
  git-ignored); supply your own workbooks at runtime / on the command line.
