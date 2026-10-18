# GST Audit Reconciliation Tool

**Created by: Purva Doshi**

Browser-based GST audit reconciliation (Streamlit). Upload the portal downloads and the client's Tally / Excel registers; the app reconciles and produces audit-ready Excel with live formulas.

## Tabs
1. **ITC Reco** - GSTR-2B vs Purchase Register, invoice-wise. Reads 2B (B2B, B2BA, B2B-CDNR/DNR, ISD, IMPG) and ties its detail sheets to the portal's own ITC summary. Matching: GSTIN + invoice no -> numeric part -> amount -> GSTIN typo (only when the two GSTINs could be the same supplier: same PAN or <=2 character difference). Flags tax-head errors, ITC not available, reverse charge, amendments, credit/debit notes. Indicative ITC for 3B Table 4(A).
2. **Sales Reco** - E-invoice / GSTR-1 vs Sales Register. Header-driven (any Tally layout), editable ledger classes, cancelled-but-live, credit/debit notes, HSN tie-out, e-invoice auto-population errors.
3. **3B Check** - upload the **filed GSTR-3B PDF** (Tables 3.1, 4, 6.1 read automatically). Checks: 3.1(a) / 3.1(b) vs e-invoice regular vs SEZ/export, vs GSTR-1 and books (if supplied), RCM vs 2B (with invoice-date timing), ITC vs 2B and vs the ITC reco, reversals, 4(C) arithmetic, Table 6.1 payment (ITC + cash, cross-utilisation, opening-ledger use), 3.1(e).

## Safety checks built in
- Files from different GSTINs or tax periods are flagged.
- Sales / ITC results are reused in the 3B tab only while the same files are still uploaded.
- A check whose reference was not supplied is skipped, never compared with zero.
- Results show a warning when settings changed after the last run.
- ITC reco warns when the same supplier invoice number is booked more than once (amounts are added together for matching).
- Invoice-number matching ignores financial-year parts (`1299/2025-26` matches on 1299).

## Amount formats understood
`1,234.00`, `(1,234.00)`, `1,234.00 Dr / Cr`, `500/-`, `Rs 1,000`, `INR 750`, `1234.00-` (SAP trailing minus); dates as real dates, text, or Excel serial numbers.

## Files accepted
`.xlsx` and legacy `.xls` (needs `xlrd`, included). 3B as PDF (needs `pdfplumber`, included). HTML/XML saved as `.xls` must be re-saved as `.xlsx` in Excel.

## Run
```
pip install -r requirements.txt
streamlit run app.py
```
**Confidentiality:** real client data (Tally exports, GSTINs) should only be processed on your own machine or firm-approved infrastructure. A public Streamlit Cloud deployment is fine for the synthetic `sample_test_files/` demo only.

## Project layout
- `gst_core.py` - parsing, matching, Excel output (no Streamlit dependency, unit-testable)
- `app.py` - Streamlit UI only
- `tests/` - `pip install -r requirements-dev.txt && pytest`

## Try it
`sample_test_files/` contains synthetic data only (no client data).

## Not computed (reminders only)
Sec 16(4) time limit, Rule 37 (180-day), Rule 42/43 working, 17(5) blocked credits, line-by-line GSTR-1 (B2CS, amendments).
