"""
Baer Shoes - Export Sales Automation
=====================================
Streamlit app: upload invoices (.xls/.xlsx) + shipping bill PDFs +
the master checklist workbook -> get back the checklist with new rows
appended to 'Export Sales', in the exact existing format.

Run with:  streamlit run app.py
"""
import io
import tempfile
import os
import streamlit as st
import openpyxl

from parser import parse_invoice, parse_shipping_bill
from excel_writer import append_invoice
from categorize import map_category

st.set_page_config(page_title="Export Sales Automation", layout="wide")
st.title("📦 Export Sales — Invoice & Shipping Bill Automation")

st.markdown(
    "Upload one or more **invoices** (.xls/.xlsx) and their matching "
    "**shipping bill PDFs**, plus the master **checklist** workbook. "
    "The tool extracts invoice date, invoice no., quantity, value, "
    "customer, SB no./date and exchange rate, maps each line item to "
    "the correct category, and appends it to the `Export Sales` sheet "
    "in the exact same layout — formulas included. Every other sheet "
    "(including `Export Sales Summary`) is left untouched; its SUMIFS "
    "formulas already reference full columns, so they will pick up the "
    "new rows automatically the next time Excel recalculates."
)

col1, col2, col3 = st.columns(3)
with col1:
    invoice_files = st.file_uploader(
        "Invoices (.xls / .xlsx)", type=["xls", "xlsx"], accept_multiple_files=True
    )
with col2:
    sb_files = st.file_uploader(
        "Shipping Bill PDFs", type=["pdf"], accept_multiple_files=True
    )
with col3:
    checklist_file = st.file_uploader("Master checklist (.xlsx)", type=["xlsx"])

if st.button("Run extraction & preview", type="primary"):
    if not (invoice_files and sb_files and checklist_file):
        st.error("Please upload at least one invoice, one shipping bill, and the checklist.")
        st.stop()

    tmpdir = tempfile.mkdtemp()

    # --- parse invoices ---
    invoices = {}
    for f in invoice_files:
        path = os.path.join(tmpdir, f.name)
        with open(path, "wb") as out:
            out.write(f.getbuffer())
        data = parse_invoice(path)
        invoices[data["invoice_no"]] = data

    # --- parse shipping bills ---
    sbs = {}
    for f in sb_files:
        path = os.path.join(tmpdir, f.name)
        with open(path, "wb") as out:
            out.write(f.getbuffer())
        data = parse_shipping_bill(path)
        sbs[data["invoice_no"]] = data

    # --- match & flag mismatches ---
    unmatched = set(invoices) ^ set(sbs)
    if unmatched:
        st.warning(f"These invoice numbers only appear in one of the two uploads: {unmatched}")

    # --- load checklist ---
    checklist_path = os.path.join(tmpdir, checklist_file.name)
    with open(checklist_path, "wb") as out:
        out.write(checklist_file.getbuffer())
    wb = openpyxl.load_workbook(checklist_path)

    st.subheader("Preview of rows to be added")
    any_unmapped = False
    for inv_no in sorted(set(invoices) & set(sbs)):
        inv, sb = invoices[inv_no], sbs[inv_no]
        st.markdown(f"**{inv_no}**  (SB {sb['sb_no']}, {sb['sb_date']}, ex.rate {inv['exchange_rate']})")
        rows = []
        for item in inv["line_items"]:
            cat = map_category(item["description"], item["grade"])
            if cat.startswith("UNMAPPED"):
                any_unmapped = True
            rows.append({
                "Category": cat,
                "Pairs": item["qty"],
                "Value (EUR)": item["value_eur"],
                "Original description": item["description"],
            })
        st.table(rows)

    if any_unmapped:
        st.error(
            "⚠️ Some line items could not be mapped to an existing category "
            "(shown as 'UNMAPPED: ...'). Add the new item type/keyword to "
            "categorize.py before trusting the output, or fix it manually "
            "in Excel afterwards."
        )

    # --- write ---
    for inv_no in sorted(set(invoices) & set(sbs)):
        append_invoice(wb, invoices[inv_no], sbs[inv_no])

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)

    st.success("Done. Download the updated checklist below and open it in Excel "
               "(press Ctrl+Alt+F9 to force a full recalculation of all formulas).")
    st.download_button(
        "⬇️ Download updated checklist",
        data=out_buf,
        file_name="Checklist_updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
