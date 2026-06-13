# Work-Power-Tools

Materials-engineering tools for Ansaldo Energia (AEG), packaged as a single
Streamlit app with two tabs.

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
| **Metallurgical** | Actual-vs-Nominal chemical composition (element-by-element, beyond ±10% → warning, ±25% → fail); hardness (pre/post-solution sanity + built-in alloy reference, see below); completeness of header/sample fields; micrograph presence & captions; sign-off (lab / engineer / date). |
| **Coating**       | Coating-thickness measurements vs the design MIN/MAX limits; sign-off; reference-micrograph presence. |

Findings are graded **🔴 Fail / 🟠 Warning / 🔵 Note / 🟢 Pass** and shown on screen.

**Built-in hardness reference.** `HARDNESS_REF` in [`lab_review.py`](lab_review.py)
holds typical *aged-condition* hardness (HRC) for common Ni- and Co-based
gas-turbine superalloys (IN738, GTD-111/741, René 80, Nimonic/C263, FSX-414,
X-40, IN718, …). The reviewer surfaces the alloy's reference range and notes
that **post-solution** readings are expected to run *below* it (the solution-
treated state precedes re-aging), so those are informational, not failures.
Values are advisory — verify against the controlling spec.

**Micrograph legends (OCR).** With the *Read micrograph legends* option on, the
reviewer reads the burned-in legend (`<job>_E_<mag>x-<n>` and the scale bar)
from each embedded micrograph and cross-checks the magnification against the
written captions. This is best-effort and needs the **Tesseract** engine
(`packages.txt` installs `tesseract-ocr`; the app degrades gracefully without it).

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
