# Export Sales Automation — Prototype

This is a working prototype, already tested against your real BG047 and
BG048 invoices/shipping bills. It reads invoices + shipping bills and
appends correctly formatted rows to the `Export Sales` sheet of your
checklist, matching the existing layout and formulas exactly.

## Files

- `parser.py` — extracts data from the invoice (.xls/.xlsx) and the
  shipping bill (PDF).
- `categorize.py` — maps an invoice line-item description (e.g.
  `"Textile - Full Shoes for Adults (Gents)"`) to your fixed category
  vocabulary (e.g. `"Full Shoes (Gents)"`). Material (Leather/Textile/
  Leather+Textile) is ignored, matching how your existing sheet
  categorizes items — only item type + gender + grade matter.
- `excel_writer.py` — appends the new rows to `Export Sales`, continuing
  the SLNO sequence, writing `VALUE(INR)` as a live formula
  (`=K{row}*J{row}`) and adding the same SUM-formula subtotal row every
  existing invoice group has.
- `app.py` — a Streamlit front-end: upload invoices + shipping bills +
  the checklist, preview the rows that will be added, download the
  updated file.

## Running it locally

```bash
pip install -r requirements.txt
# The .xls invoices are old Crystal Reports "BIFF" files. openpyxl can't
# read them directly, so we shell out to LibreOffice to convert them to
# .xlsx first. Install LibreOffice (or the "soffice" CLI) once:
#   Windows/Mac: install LibreOffice normally
#   Ubuntu/Debian: sudo apt install libreoffice
streamlit run app.py
```

## Deploying it so anyone on your team can use it

- **Streamlit Community Cloud** (free, easiest): push this folder to a
  private GitHub repo, connect it at streamlit.io/cloud, done. Note:
  Community Cloud's servers don't have LibreOffice pre-installed — add
  a `packages.txt` file with the single line `libreoffice` and Streamlit
  Cloud will install it automatically (it uses apt under the hood).
- **Your own server / a colleague's PC**: just run
  `streamlit run app.py` — it opens a local web page, no coding needed
  to use it day-to-day.
- **Data privacy**: these are customs/GST/financial documents. Don't
  route them through a public/shared AI API. Everything in this
  prototype runs locally with plain Python (regex + openpyxl +
  pdfplumber) — no data ever leaves the machine it runs on, and no LLM
  call is needed for the extraction itself.

## Important finding from testing against your real file

`Export Sales Summary` currently has month-column blocks only up to
**August 2026** (columns ...CG). September 2026 (needed for BG047/048)
doesn't have a block yet. That sheet is 100% formula-driven off
`Export Sales`, so:

1. This tool only ever writes to `Export Sales` (the raw data). It
   never touches `Export Sales Summary`.
2. Once September's 5-column block exists in `Export Sales Summary`
   (however you currently add each new month — usually copy the
   previous month's block sideways and it drags the relative
   references along; you only fix the two *absolute* references to the
   month header), the SUMIFS formulas will already pick up the BG047/
   BG048 rows correctly, because they reference full columns
   (`'Export Sales'!I:I` etc.), not fixed ranges.
3. If you want, this can be automated too (a script that clones the
   previous month's formula block and shifts the right references) —
   just say the word and I'll add it.

## Known limitations to be aware of

- **Category mapping** is regex-based on the item description. It
  currently recognises: Full Shoes, Half Boots, Sandals, Chappals, Shoe
  Uppers, Shoe Material, Finished Leather, Good Soles — each with
  (Gents)/(Ladies) and an optional "Grade II" suffix. Anything it can't
  recognise is written as `UNMAPPED: <description>` so it's obvious and
  safe rather than silently wrong — you'd add one line to
  `categorize.py` to teach it a new item type.
- **Invoice date vs SB date vs LEO date**: the parser takes the invoice
  date printed in the invoice header, and the SB No./SB Date from the
  shipping bill's summary block. For BG048 note the shipping bill shows
  *Submission* 11-Sep-26 but *LEO* (Let Export Order — the actual
  clearance) 15-Sep-26; the tool currently uses the header "SB Date"
  field (11-Sep-26), matching what your existing rows do for the SB
  Date column — double check which date your team actually wants there
  before relying on it for a shipment with a gap between submission and
  LEO.
- **"Inland Handling Charges" and similar freight-type lines** aren't
  in the invoice item table (they're added by your team as a separate
  entry per shipment) — the tool won't invent one; add it manually if
  needed.
- Only tested against this one invoice format (Crystal Reports IGST
  invoice) and this one shipping bill format (ICEGATE Part I/II). A
  differently formatted invoice (e.g. from a different customer/
  template) would need the regex patterns in `parser.py` adjusted.
