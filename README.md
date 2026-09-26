# Export Sales Automation — Prototype

This is a working prototype, already tested against your real BG047 and
BG048 invoices/shipping bills. It reads invoices + shipping bills and
appends correctly formatted rows to the `Export Sales` sheet of your
checklist, matching the existing layout and formulas exactly.

## Files

**Export Sales tab:**
- `parser.py` — extracts data from the invoice (.xls/.xlsx) and the
  shipping bill (PDF).
- `categorize.py` — maps an invoice line-item description (e.g.
  `"Textile - Full Shoes for Adults (Gents)"`) to your fixed category
  vocabulary (e.g. `"Full Shoes (Gents)"`).
- `excel_writer.py` — appends the new rows to `Export Sales`, continuing
  the SLNO sequence, writing `VALUE(INR)` as a live formula, copying
  the exact font/fill/border styling from your existing rows (including
  the coloured subtotal-row background), and adding the same
  SUM-formula subtotal row every existing invoice group has.
- `summary_writer.py` — if the invoice's month doesn't have a 5-column
  formula block yet in `Export Sales Summary`, clones the most recent
  block forward (shifting only its own internal cell references,
  leaving `Export Sales`! references and the category-label column
  untouched) so every row's quirks are preserved exactly, then fills in
  as many months as needed to catch up.

**Local Sales tab:**
- `local_parser.py` — extracts invoice no./date, customer, goods
  description, pairs and taxable value (INR, excl. GST) from domestic
  "TAX-Invoice" GST invoices, and a second parser for the small
  LUT/courier export invoices (e.g. Fontana via DHL) that get logged
  here rather than through Export Sales.
- `local_writer.py` — finds the matching customer's block in
  `Local Sales working`, works out which of its Pairs/INR/Euro columns
  the item belongs to (Shoes vs Material vs Uppers — checking the
  product type first, so a boot described as "...Leather & Material"
  still files under Shoes, not Material), and appends the invoice's
  total to that month's running `+` formula, exactly the way it's
  always been maintained by hand. It never restructures the sheet and
  raises an error (rather than guessing) if it can't confidently find
  the right block.

**Front-end:**
- `app.py` — Streamlit app with two tabs (Export Sales, Local Sales).
  The Local Sales tab makes you confirm the customer name and category
  for every invoice before anything is written — nothing is silently
  auto-filed.

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

- **Category mapping (Export)** is regex-based on the item description.
  Anything it can't recognise is written as `UNMAPPED: <description>` so
  it's obvious and safe rather than silently wrong.
- **Local Sales is intentionally conservative.** `Local Sales working`
  is a hand-built sheet (per-customer blocks, running `+` formulas, some
  cells even reference *other* Excel files not in this workbook, and
  Baer GmbH/Stroeber's numbers are a back-calculated residual off
  `Export Sales Summary`, not invoice-driven at all). The tool:
  - never touches the Baer GmbH/Stroeber block — that stays fully manual.
  - requires you to confirm the customer name and category for every
    invoice in the UI before writing anything.
  - raises a clear error rather than guessing if it can't find a
    matching customer block or category column — check spelling of the
    customer name against the block label in the sheet if this happens.
  - only ever adds one `+<value>` term to the Pairs/INR/Euro cells for
    one month/customer/category. It never edits "Rate" columns (those
    recompute themselves) or any other cell.
- Only tested against the invoice formats you've shared so far (BG-series
  export IGST invoices, C-series domestic GST invoices, one Fontana
  LUT-courier invoice). A new customer's invoice template, or an invoice
  from a new courier, may need `parser.py`/`local_parser.py` adjusted.
- **Invoice date vs SB date vs LEO date**: the parser takes the invoice
  date printed in the invoice header, and the SB No./SB Date from the
  shipping bill's summary block (not the LEO/clearance date, which can
  be a few days later) — worth double-checking which date you actually
  want recorded for shipments with a gap between submission and LEO.
- "Inland Handling Charges" and similar freight-type lines aren't in the
  invoice item table — add those manually if needed.
