# Work Power Tools

A small Streamlit workspace of engineering report utilities for Ansaldo Energia
Materials Engineering. Pick a tool from the sidebar.

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 🔬 SEM Report Converter
Converts vendor SEM metallurgical PDFs into formatted Ansaldo Word reports —
parses job/serial/material/γ′ data, extracts and lays out the SEM figures, and
builds the cover, TOC, microstructure and conclusion sections.

- Module: `sem_convert.py` (also runnable as a CLI: `python sem_convert.py vendor.pdf`)

## 🛠️ IIR Review
Automated quality/consistency review of **Incoming Inspection Reports** (Detailed
Assessment Customer Reports) delivered as `.xlsx`. Upload one or more workbooks
and get a severity-tagged findings checklist (`.xlsx`) plus an on-screen summary.

Checks include:
- **Identity / metadata** — doc number (format + cover/contents match), customer,
  component, preparer, reviewer, approver, PO# present and well-formed
- **Quantities** — Received = Scrap + Reconditionable; positions listed = received;
  serial-number scope totals reconcile; the stated sum row matches the scopes
  actually marked on each row (catches stale totals); Received-Parts table vs
  Serial-Number protocol agree (catches off-by-one counts)
- **Integrity** — unique & contiguous position numbers, serial numbers present and
  unique, valid repair-scope values (L/M/H/S), scrap mark ↔ scope 'S', every part scoped
- **Consistency** — Summary-of-Damages finding counts reconcile with the actual
  defect marks in the protocol; executive-summary received count and scrap positions agree
- **Completeness** — an embedded photo for every incoming-photo caption; consistent
  page numbering

Output workbook tabs: **Review Summary** (identity + verdict), **Findings**
(colour-coded by severity), **Extracted Data** (per-position traceability table).

**Batch mode** — review several reports at once for a combined `IIR_Batch_Summary.xlsx`
(one row per report with verdict + counts, plus a pooled **All Findings** tab).

- Module: `iir_review.py`. CLI: `python iir_review.py report.xlsx` (one report),
  `python iir_review.py report.xlsx out.xlsx` (explicit output), or
  `python iir_review.py *.xlsx` (many reports → individual checklists + batch summary).
