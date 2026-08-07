# Work-Power-Tools

Materials-engineering tools for Ansaldo Energia (AEG), packaged as a single
Streamlit app with four tools. `app.py` is a thin router; the presentation
layer lives in `ui/` (`ui/theme.py`, `ui/components.py` for the shared
severity model / findings list / batch table, and one module per tool). The
domain logic each tool calls into — `lab_review.py`, `iir_review.py`,
`sem_convert.py`, `report_render.py`, `photo_lib.py` — is unchanged.

> **⚠ Extraction rebuild in progress.** A 2026 audit reproduced ten defects in
> the metallurgical extraction layer, including two release-critical false
> negatives (multi-sample reports collapse to sample 1; only the first
> composition table is read). The review *rules* are sound; the *parser* under
> them is being rebuilt. Before changing `lab_review.py`, read
> [`docs/REBUILD.md`](docs/REBUILD.md) — spec and defect register,
> [`docs/HANDOFF.md`](docs/HANDOFF.md) — working protocol, and
> [`docs/FEEDBACK.md`](docs/FEEDBACK.md) — current state. Acceptance criteria
> are executable in [`tests/test_parser_contract.py`](tests/test_parser_contract.py).

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
| **Metallurgical** | Report title ↔ content identity (job / type / stage / component / machine-set / customer); workbook formula/error integrity; Actual-vs-Nominal composition (element-by-element, duplicate/misaligned headers, `<LOD` major elements, chemistry-total sanity); hardness (pre/post-solution sanity + built-in alloy reference); duplicate sample/serial identifiers; explicit report disposition; sign-off; coating cell ↔ comment (presence & type); caption ↔ embedded-image completeness; caption integrity (duplicate/gap numbers, etch status required, caption ↔ image-legend magnification, comment over-references); comment ↔ material/Result verdict; micrograph job identity, legends, etch contrast and burned-in thickness. |
| **Coating**       | Filename ↔ content; coating-thickness measurements vs design MIN/MAX limits; sign-off; reference-micrograph presence. |

Findings are graded **🔴 Fail / 🟠 Warning / 🔵 Note / 🟢 Pass**.

**Report-centric annotated review.** The annotated report is the default and
automatic result after upload — there is no separate findings table or evidence
mode to open. The reviewer presents a **pixel-faithful, page-at-a-time** render
of the report: issue areas are highlighted and numbered on the page, while
concise matching issue cards remain beside that page. One issue can mark several
affected cells without being counted or repeated several times. The real
workbook is converted with **LibreOffice**
(original fonts, column widths, borders and embedded micrographs intact) and each
flagged cell is filled and badged. Page controls replace the former extremely
tall stacked image, and the downloadable annotated output is a multi-page PDF.
The review covers composition deviations, hardness, blank header/sign-off fields,
captions without an etch status, and coating thickness outside the design limits.
Implemented in
[`report_render.py`](report_render.py); cell anchoring comes from
`lab_review.collect_highlights`.

**Comment callouts on the page.** Every finding is written out beside the page it
marks, in a comment margin, joined to its own cell by a leader line — the same
composite on screen and in the PDF, so what the reviewer approves is what the
recipient opens. The vertical run of each leader is kept in the sheet's own right
print margin, so only the short run along the marked cell's row crosses any
content. Each callout is placed level with the row it is about; when the comments
outgrow one page height they continue in a further margin column rather than
being cut. Every numbered badge has a matching callout: findings the reviewer
dismissed carry their dismissal reason, and template-scoped ones say so, both in
grey so they do not read as live problems to whoever receives the report.

**Zoom.** The page sits in a scrollable viewport rather than an `st.image` that
can only fit its column: **Fit / 100% / 150% / 200% / 300%**, panned by
scrolling, and from 200% the workbook is re-rendered at a higher resolution so
there is real detail to magnify. **Locate** on any finding shows a magnified crop
of its cell, on the exact render as well as the fallback. The **Callouts / Page**
switch toggles between the sendable composite and the report at full width.

**Review comments, and sending the report.** Comments are written in the app, can
be attached to a finding (inheriting its cell, so the callout points at the value
it is about), to one page, or to the whole report, and travel with the document:
into the callout margin, into the PDF, and into the message body. From the same
panel the annotated PDF can be downloaded, e-mailed, or shared from Google Drive
— see [Sending a reviewed report](#sending-a-reviewed-report).

LibreOffice takes a few seconds and is installed by `packages.txt` as
`libreoffice-calc`; the PDF is rasterised with **PyMuPDF**. If exact rendering is
unavailable, the reviewer automatically shows a simplified annotated sheet instead
of an empty evidence area. Cell highlights are placed by filling each flagged cell
with a uniquely detectable colour in the workbook, rendering it, then locating that
colour in the raster — so every numbered badge lands on the actual affected cell.
When several reports are uploaded, a compact report selector replaces the former
batch table and keeps the review focused on one annotated report at a time.

**Cross-report checks.** Two things are invisible to a single-report review and
are therefore run across a batch: a revised upload that still carries the
earlier version's blockers (`add_version_findings`), and an **Actual composition
that is byte-identical under two different job numbers**
(`find_duplicate_compositions`) — independent EDS/ICP results are not expected to
agree to the reported decimal, so an exact match across five or more elements
says the table was copied rather than measured. Found on a real 27-report batch.
Same-job pairs are excluded: those are revisions, which the first check covers.

[`composition_store.py`](composition_store.py) gives that check **memory across
sessions** — it fingerprints each reviewed report (job, alloy, source, Actual
composition) so a report reviewed today can be caught copying one reviewed weeks
ago without uploading them side by side. Storage auto-selects like the photo
library (GitHub → Drive → local `composition_store/`, gitignored) and it is
opt-in: only the CLI records fingerprints.

```bash
python3 composition_store.py "path/to/report.xlsx" [more.xlsx ...]
python3 lab_review.py report_a.xlsx report_b.xlsx     # cross-report section
python3 batch_review.py folder/                       # whole-folder audit
```

#### Sending a reviewed report

The annotated PDF can always be downloaded. Sending it from the app is optional
and appears only once one of the two transports is configured — implemented in
[`share.py`](share.py), which adds no dependency (e-mail is stdlib `smtplib`, and
Drive reuses the photo library's OAuth credentials).

* **E-mail** — the PDF is attached and the review is written into the body
  (verdict, findings, and the reviewer's comments in full) so the mail says
  something useful before the attachment is opened. Streamlit secrets or
  environment variables:
  ```toml
  smtp_host     = "smtp.office365.com"
  smtp_from     = "materials.lab@example.com"   # defaults to smtp_user
  smtp_user     = "materials.lab@example.com"   # optional on an open relay
  smtp_password = "...."
  # smtp_port   = 587      # default 587, or 465 with smtp_ssl
  # smtp_ssl    = false    # implicit TLS
  # smtp_starttls = true   # default on unless smtp_ssl
  ```
* **Google Drive** — uploads the PDF to an app-owned **"AEG Report Reviews"**
  folder and grants each named recipient read access, returning a link. Uses the
  same `drive_client_id` / `drive_client_secret` / `drive_refresh_token` as the
  photo library. Access is granted **per person, never "anyone with the link"** —
  these are customer documents.

Recipients are typed as a comma- or space-separated list; anything that isn't an
address is reported back rather than silently dropped, and sending is only
possible with at least one valid recipient.

### 🖼️ Photo Library *(new)*
Extracts the embedded micrographs from a reviewed report into a **per-alloy**
folder structure with a JSON index, and serves an in-app gallery: pick an alloy
→ see its micrographs with the data of the report they came from (job, magnification,
etch state, source, any thickness measurements). Implemented in
[`photo_lib.py`](photo_lib.py). Add to it via **Actions → Add to photo library**
in Lab Report Review, or the CLI:

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

**Micrograph analysis (OCR).** Enabled by default when Tesseract is available,
because mixed-job evidence is release-critical; it can still be switched off for
a faster text-only pass. The reviewer reads each embedded micrograph's burned-in
legend (`<job>_E_<mag>x-<n>`
+ scale bar) and cross-checks the magnification and job number against the
captions; gauges etched-vs-low-contrast via edge density (advisory — faint
post-HT etching reads as low-contrast); and reads burned-in thickness labels
(e.g. `42 µm`) to surface alongside the comment's thickness values. Best-effort;
needs the **Tesseract** engine (`packages.txt` installs `tesseract-ocr`; the app
degrades gracefully without it).

Magnification has no hard-coded allow-list: **600x and any other positive written
value are valid report data**. The written caption is the source of record. Each
image legend is OCR-read through three different preprocessing passes, and a
number is accepted as stable OCR only when at least two passes agree without a
tie. An isolated or ambiguous number is suppressed. If stable OCR disagrees with
the paired caption, the reviewer marks that caption with a warning to inspect the
burned-in legend; it does not reject either value. Downstream photo-library
metadata also prefers the paired written caption over OCR.

**Etch handling.** A caption that explicitly states *unetched / as-polished* is
surfaced on its own (legitimate for thickness / crack work, but worth confirming
for a microstructure assessment) — separate from the *no etch status* warning for
captions that say nothing. When captions can be mapped one-to-one to the embedded
micrographs (by drawing-anchor order), each picture is checked for a caption↔image
mismatch: a caption naming an etchant whose micrograph reads low-contrast (the etch
may not have developed), or a caption saying unetched whose micrograph reads
etched. Contrast is advisory, so mismatches read "verify", not fail.

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
severity is tunable (🔴 Fail / 🟠 Warn / 🔵 Info / ⚪ Off) live in the UI, via a
**Check severities** popover with Strict / Default / Lenient presets or
per-check tuning; defaults live in `iir_review.CHECK_CATALOG`. Review several
at once for a combined **Batch Summary** workbook.

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

Run the deterministic reviewer regression suite with:

```bash
python3 -m unittest discover -s tests -v
```

## Notes

* Composition tolerances and the advisory hardness ranges are constants at the
  top of [`lab_review.py`](lab_review.py) (`COMP_WARN_REL`, `HARDNESS_REF`, …) —
  adjust them to match your controlling specification.
* Raw customer report `.xlsx` files are **not tracked in git** (`*.xlsx` is
  git-ignored); supply your own workbooks at runtime / on the command line.
