# GST Audit Reconciliation Tool

**Created by: Purva Doshi**

Free web app for GST audit reconciliation. Runs in any browser (Streamlit).

## Tabs
1. **ITC Reco — GSTR-2B vs Purchase Register (invoice-wise)**
   - 2B read by header name from every sheet: B2B, B2BA (amendments), B2B-CDNR (credit/debit notes), ISD, IMPG.
   - Books read by header name (any Tally / SAP / Excel layout); older Tally sub-row and pivot layouts still accepted.
   - Matching tiers: GSTIN + invoice no -> GSTIN + numeric part of invoice no -> GSTIN + tax amount -> invoice no/amount with GSTIN typo.
   - Flags: tax-head mismatch (IGST vs CGST/SGST), ITC-not-available on 2B, reverse charge (booked as regular ITC / not booked), only-in-2B, only-in-books, GSTIN inferred from party name, credit/debit notes.
   - Output: indicative ITC for 3B Table 4(A) (lower of books and 2B), tie-out with live Excel formulas, Auditor Actions, Party Summary, Exceptions, integrity checks.
2. **Sales Reco — E-Invoice / GSTR-1 vs Sales Register** (header-driven; ledger columns auto-classified and editable; cancelled-but-live detection; credit/debit notes; HSN tie-out).
3. **3B Check** — 3B vs GSTR-1 vs Books (output) and 3B vs 2B vs ITC reco (input), reverse charge, 4(C) arithmetic, indicative tax utilisation (Rule 88A). 2B / books / ITC references auto-fill from the other tabs.

## Deploy on Streamlit Cloud (free)
1. Push these files to GitHub  2. share.streamlit.io -> select repo -> `app.py` -> Deploy

## Try it
`sample_test_files/` has synthetic data (no real client) that exercises every status.
