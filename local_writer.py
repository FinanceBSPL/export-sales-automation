"""
Appends a local-sale (or Fontana-style LUT-courier) invoice's totals into
the correct customer / month / category cell of 'Local Sales working'.

This sheet is NOT a clean table -- each customer has its own hand-built
block, and each month's cell is a running '+' formula that gets a new
term added every time an invoice comes in (e.g. '=216+224' becomes
'=216+224+50'). This module replicates exactly that convention rather
than restructuring the sheet, so it stays visually/behaviourally
identical to how the accountant has always maintained it.

Safety rules baked in:
  - Only ever WRITES into a Pairs cell and its matching currency cell for
    one category, for one customer block, for one month row. Never
    touches 'Rate' columns (those are formulas like '=D77/B77' that
    recompute themselves), never touches unrelated ad-hoc columns.
  - Refuses (raises) rather than guesses if the customer block or the
    category can't be found -- silence would be worse than an error here.
"""
import re
import datetime as dt

SHEET = "Local Sales working"

CATEGORY_KEYWORDS = [
    (r'shoe|boot|sandal|chappal', 'shoes'),   # checked first: the PRODUCT type
    (r'upper', 'uppers'),
    (r'\bmaterial\b', 'material'),            # only a fallback: composition
]                                              # wording ("...& Material") must
                                               # not override an actual shoe/
                                               # boot/sandal product match.


def categorize_local(goods_description: str) -> str:
    desc = (goods_description or "").lower()
    for pat, label in CATEGORY_KEYWORDS:
        if re.search(pat, desc):
            return label
    return "shoes"  # sensible default: most local invoices are finished footwear


def _find_customer_block(ws, customer_query: str) -> int:
    """Return the row of the customer's block label, matched
    case-insensitively and ignoring extra whitespace/punctuation."""
    def norm(s):
        return re.sub(r'[^a-z0-9]', '', s.lower())

    target = norm(customer_query)
    best = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip() and not v.strip().lower().startswith(("month", "total")):
            label = norm(v)
            if label and (label in target or target in label):
                best = r
    if best is None:
        raise ValueError(
            f"Could not find a customer block matching '{customer_query}' in "
            f"'{SHEET}'. Add it manually first, or check the spelling."
        )
    return best


def _parse_block_columns(ws, label_row: int) -> dict:
    """Read the two header rows under a customer label and build
    {category_name_lower: {'pairs': col, 'inr': col, 'euro': col}}."""
    cat_row = label_row + 1
    sub_row = label_row + 2

    cat_starts = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=cat_row, column=c).value
        if v and str(v).strip().lower() != "month":
            cat_starts.append((c, str(v).strip()))
    cat_starts.append((ws.max_column + 1, None))

    colmap = {}
    for i in range(len(cat_starts) - 1):
        start_c, name = cat_starts[i]
        end_c = cat_starts[i + 1][0]
        entry = {"pairs": None, "inr": None, "euro": None}
        for c in range(start_c, end_c):
            sub = ws.cell(row=sub_row, column=c).value
            if not sub:
                continue
            sub_l = str(sub).strip().lower()
            if sub_l == "pairs":
                entry["pairs"] = c
            elif sub_l == "inr":
                entry["inr"] = c
            elif sub_l == "euro":
                entry["euro"] = c
        colmap[name.lower()] = entry

    return colmap


def _find_category_key(colmap: dict, category: str) -> str:
    for key in colmap:
        if category in key:
            return key
    raise ValueError(
        f"No '{category}' column found for this customer block "
        f"(found categories: {list(colmap.keys())}). Add it manually, or "
        f"check goods_description keyword matching."
    )


def _month_row(ws, label_row: int, target_month: dt.datetime) -> int:
    first_month_row = label_row + 3
    for r in range(first_month_row, first_month_row + 14):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, dt.datetime) and v.year == target_month.year and v.month == target_month.month:
            return r
    raise ValueError(f"Could not find a row for {target_month:%b %Y} under this customer block.")


def _append_or_set(cell, new_value):
    old = cell.value
    if old in (None, 0):
        cell.value = new_value
    elif isinstance(old, str) and old.startswith("="):
        cell.value = f"{old}+{new_value}"
    else:
        # existing literal number -> turn into a running formula, matching
        # the sheet's own convention, without losing the prior total
        cell.value = f"={old}+{new_value}"


def append_local_entry(wb, customer: str, invoice_month: dt.datetime, category: str,
                        pairs: float, inr_value: float = None, eur_value: float = None) -> dict:
    ws = wb[SHEET]
    label_row = _find_customer_block(ws, customer)
    colmap = _parse_block_columns(ws, label_row)
    cat_key = _find_category_key(colmap, category)
    cols = colmap[cat_key]
    row = _month_row(ws, label_row, invoice_month)

    touched = {}
    if cols["pairs"]:
        _append_or_set(ws.cell(row=row, column=cols["pairs"]), pairs)
        touched["pairs_cell"] = ws.cell(row=row, column=cols["pairs"]).coordinate
    if cols["inr"] and inr_value is not None:
        _append_or_set(ws.cell(row=row, column=cols["inr"]), inr_value)
        touched["inr_cell"] = ws.cell(row=row, column=cols["inr"]).coordinate
    if cols["euro"] and eur_value is not None:
        _append_or_set(ws.cell(row=row, column=cols["euro"]), eur_value)
        touched["euro_cell"] = ws.cell(row=row, column=cols["euro"]).coordinate

    touched["customer_block_row"] = label_row
    touched["month_row"] = row
    touched["category_matched"] = cat_key
    return touched
