"""
Baer Shoes - Sales Automation
=============================
Two tabs:
  1. Export Sales  -- invoices + shipping bills -> append to 'Export Sales'
     (+ auto-extend 'Export Sales Summary' with the month's formula block)
  2. Local Sales    -- domestic invoices (and LUT-courier invoices the
     company logs here by convention) -> append into the matching
     customer's running '+' formula in 'Local Sales working'

Run with:  streamlit run app.py
"""
import io
import os
import tempfile
import datetime as dt

import streamlit as st
import openpyxl

from parser import parse_invoice, parse_shipping_bill
from excel_writer import append_invoice
from summary_writer import ensure_month_block
from categorize import map_category

from local_parser import parse_domestic_invoice, parse_lut_courier_invoice
from local_writer import append_local_entry, categorize_local

st.set_page_config(page_title="Baer Shoes — Sales Automation", layout="wide")
st.title("📦 Baer Shoes — Sales Automation")

tab_export, tab_local = st.tabs(["Export Sales", "Local Sales"])


def _save_upload(f, tmpdir):
    path = os.path.join(tmpdir, f.name)
    with open(path, "wb") as out:
        out.write(f.getbuffer())
    return path


def _parse_month(ddmmyyyy: str) -> dt.datetime:
    d = dt.datetime.strptime(ddmmyyyy, "%d-%m-%Y")
    return dt.datetime(d.year, d.month, 1)


# ============================================================
# TAB 1 — EXPORT SALES  (unchanged logic from before)
# ============================================================
with tab_export:
    st.markdown(
        "Upload one or more **invoices** (.xls/.xlsx) and their matching "
        "**shipping bill PDFs**, plus the master **checklist** workbook. "
        "Appends to `Export Sales` in the exact existing format (formulas "
        "included), and auto-adds the month's formula block to "
        "`Export Sales Summary` if it doesn't exist yet."
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        invoice_files = st.file_uploader(
            "Invoices (.xls / .xlsx)", type=["xls", "xlsx"],
            accept_multiple_files=True, key="exp_inv",
        )
    with col2:
        sb_files = st.file_uploader(
            "Shipping Bill PDFs", type=["pdf"], accept_multiple_files=True, key="exp_sb",
        )
    with col3:
        checklist_file = st.file_uploader("Master checklist (.xlsx)", type=["xlsx"], key="exp_checklist")

    if st.button("Run extraction & preview", type="primary", key="exp_run"):
        if not (invoice_files and sb_files and checklist_file):
            st.error("Please upload at least one invoice, one shipping bill, and the checklist.")
            st.stop()

        tmpdir = tempfile.mkdtemp()

        invoices = {}
        for f in invoice_files:
            data = parse_invoice(_save_upload(f, tmpdir))
            invoices[data["invoice_no"]] = data

        sbs = {}
        for f in sb_files:
            data = parse_shipping_bill(_save_upload(f, tmpdir))
            sbs[data["invoice_no"]] = data

        unmatched = set(invoices) ^ set(sbs)
        if unmatched:
            st.warning(f"These invoice numbers only appear in one of the two uploads: {unmatched}")

        checklist_path = _save_upload(checklist_file, tmpdir)
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
                    "Category": cat, "Pairs": item["qty"],
                    "Value (EUR)": item["value_eur"], "Original description": item["description"],
                })
            st.table(rows)

        if any_unmapped:
            st.error(
                "⚠️ Some line items could not be mapped to an existing category "
                "(shown as 'UNMAPPED: ...'). Add the new item type/keyword to "
                "categorize.py before trusting the output."
            )

        for inv_no in sorted(set(invoices) & set(sbs)):
            inv = invoices[inv_no]
            month = _parse_month(inv["invoice_date"])
            ensure_month_block(wb, month)
            append_invoice(wb, inv, sbs[inv_no])

        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)

        st.success("Done. Download below and open in Excel (Ctrl+Alt+F9 to force recalc).")
        st.download_button(
            "⬇️ Download updated checklist", data=out_buf,
            file_name="Checklist_updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="exp_download",
        )


# ============================================================
# TAB 2 — LOCAL SALES  (new)
# ============================================================
with tab_local:
    st.markdown(
        "Upload domestic (GST) invoices, plus any small LUT/courier export "
        "invoices your company logs here by convention (e.g. Fontana via "
        "DHL). For each invoice, pick which type it is, then confirm the "
        "customer this money should be attributed to before writing "
        "anything — **the tool will only extend the existing running "
        "formula for that customer/month/category; it never restructures "
        "the sheet or touches other customers' blocks.**"
    )

    local_files = st.file_uploader(
        "Local sale / LUT-courier invoices (.xls / .xlsx)",
        type=["xls", "xlsx"], accept_multiple_files=True, key="loc_inv",
    )
    local_checklist_file = st.file_uploader(
        "Master checklist (.xlsx)", type=["xlsx"], key="loc_checklist"
    )

    if local_files and local_checklist_file:
        tmpdir = tempfile.mkdtemp()
        parsed = []
        for f in local_files:
            path = _save_upload(f, tmpdir)
            try:
                data = parse_domestic_invoice(path)
                data["_kind"] = "domestic"
                if data["invoice_no"] is None or data["pairs"] is None:
                    raise ValueError("looks incomplete")
            except Exception:
                data = parse_lut_courier_invoice(path)
                data["_kind"] = "lut_courier"
            data["_filename"] = f.name
            parsed.append(data)

        st.subheader("Review before writing — confirm each invoice")
        confirmed = []
        for i, d in enumerate(parsed):
            with st.expander(f"{d['_filename']}  —  {d.get('invoice_no') or 'invoice no. not found'}", expanded=True):
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.write("**Detected type:**", d["_kind"])
                    st.write("**Invoice date:**", d.get("invoice_date"))
                    st.write("**Goods description:**", d.get("goods_description"))
                with c2:
                    st.write("**Pairs:**", d.get("pairs"))
                    st.write("**Value INR (excl. GST):**", d.get("value_inr"))
                    if d.get("value_eur") is not None:
                        st.write("**Value EUR:**", d.get("value_eur"))
                with c3:
                    customer = st.text_input(
                        "Customer block to credit this to (must match an "
                        "existing block name in 'Local Sales working')",
                        value=d.get("customer") or "", key=f"cust_{i}",
                    )
                    default_cat = categorize_local(d.get("goods_description"))
                    category = st.selectbox(
                        "Category column", ["shoes", "material", "uppers"],
                        index=["shoes", "material", "uppers"].index(default_cat),
                        key=f"cat_{i}",
                    )
                    include = st.checkbox("Include this invoice", value=True, key=f"inc_{i}")
                d["_customer_override"] = customer
                d["_category_override"] = category
                d["_include"] = include
                confirmed.append(d)

        if st.button("Append confirmed invoices", type="primary", key="loc_run"):
            checklist_path = _save_upload(local_checklist_file, tmpdir)
            wb = openpyxl.load_workbook(checklist_path)

            results, errors = [], []
            for d in confirmed:
                if not d["_include"]:
                    continue
                try:
                    month = _parse_month(d["invoice_date"])
                    res = append_local_entry(
                        wb, d["_customer_override"], month, d["_category_override"],
                        pairs=d["pairs"], inr_value=d.get("value_inr"), eur_value=d.get("value_eur"),
                    )
                    results.append((d["invoice_no"], res))
                except Exception as e:
                    errors.append((d.get("invoice_no", d["_filename"]), str(e)))

            if errors:
                st.error("Some invoices could NOT be written (nothing was guessed — fix and retry):")
                for name, msg in errors:
                    st.write(f"- **{name}**: {msg}")

            if results:
                st.success(f"Written {len(results)} invoice(s):")
                for name, res in results:
                    st.write(f"- **{name}** → {res}")

                out_buf = io.BytesIO()
                wb.save(out_buf)
                out_buf.seek(0)
                st.download_button(
                    "⬇️ Download updated checklist", data=out_buf,
                    file_name="Checklist_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="loc_download",
                )
