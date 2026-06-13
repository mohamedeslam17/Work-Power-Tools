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
  serial-number scope totals reconcile; Received-Parts table vs Serial-Number
  protocol agree (catches off-by-one counts)
- **Integrity** — unique & contiguous position numbers, serial numbers present and
  unique, every part has a repair scope or a scrap mark
- **Consistency** — executive-summary received count and named scrap positions
  match the protocol; finding counts ≤ received quantity
- **Completeness** — an embedded photo for every incoming-photo caption; consistent
  page numbering

Output workbook tabs: **Review Summary** (identity + verdict), **Findings**
(colour-coded by severity), **Extracted Data** (per-position traceability table).

- Module: `iir_review.py` (also runnable as a CLI: `python iir_review.py report.xlsx [findings.xlsx]`)
