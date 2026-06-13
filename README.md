# Work-Power-Tools

Materials-engineering tools for Ansaldo Energia (AEG), packaged as a single
Streamlit app with three tabs.

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

* **GitHub** *(recommended — no IT, browser-only)* — commits micrographs into the
  repo via a fine-grained PAT (Contents: read & write). Set in
  `.streamlit/secrets.toml` / Streamlit Cloud secrets:
  ```toml
  github_token  = "github_pat_..."
  github_repo   = "owner/Work-Power-Tools"
  github_branch = "main"          # branch to store files on
  github_base   = "photo_library" # path prefix in the repo
  ```
* **Google Drive** *(alternative)* — OAuth acting as you (see [`drive_store.py`](drive_store.py)).
  Run `python3 drive_store.py --auth` once for a refresh token, then set
  `drive_client_id`, `drive_client_secret`, `drive_refresh_token`, `drive_folder_id`.
  Needs `google-api-python-client google-auth google-auth-oauthlib`.
* **Local** *(default / fallback)* — a folder (`PHOTO_LIBRARY_DIR`, default
  `photo_library/`); the committed seed library survives via git, but runtime
  additions don't persist across reboots.

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
* Sample reports used during development live in `Samples 1/`.
