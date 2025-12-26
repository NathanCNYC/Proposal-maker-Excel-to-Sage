"""
extract_proposal.py  –  v3
  • Contact / Phone / Email read from column E (two cells right of the label).
  • Total Price read from column C, one row below the label.
"""

import xlwings as xw


def extract_proposal_data() -> dict:
    # ── 1. grab active workbook ───────────────────────────────────────────────
    app = xw.apps.active
    if not app:
        raise RuntimeError("Excel isn’t running or no window is active.")

    wb = app.books.active
    if not wb or "M Breakdown" not in wb.name:
        raise RuntimeError(
            "Activate the correct workbook (its name must include 'M Breakdown') and rerun."
        )
    
    wb_path = wb.fullname 

    ws = wb.sheets["Proposal"]

    # ── 2. find last used row (any column) ────────────────────────────────────
    used_rows = ws.api.UsedRange.Rows.Count + ws.api.UsedRange.Row - 1

    # holders
    data = {
        "JOB": None,
        "CONTACT": None,
        "PHONE": None,
        "EMAIL": None,
        "SCOPE_OF_WORK": "",
        "TOTAL_PRICE": None,
    }
    row_price_based_on = row_total_price = None

    # ── 3. scan column C ──────────────────────────────────────────────────────
    for row in range(1, used_rows + 1):
        label = str(ws.range(f"C{row}").value or "").strip()

        if not label:
            continue

        label_lower = label.lower().rstrip(":")  # normalize

        if label_lower == "job":
            data["JOB"] = ws.range(f"D{row}").value  # still one cell right
        elif label_lower == "contact":
            data["CONTACT"] = ws.range(f"E{row}").value  # two cells right
        elif label_lower == "phone":
            data["PHONE"] = ws.range(f"E{row}").value
        elif label_lower == "email":
            data["EMAIL"] = ws.range(f"E{row}").value
        elif label.startswith("Price based on") and row_price_based_on is None:
            row_price_based_on = row
        elif label_lower == "total price":
            row_total_price = row
            break  # reached the end of the block

    # ── 4. scope of work ──────────────────────────────────────────────────────
    if row_price_based_on and row_total_price:
        lines = []
        for r in range(row_price_based_on, row_total_price):
            c_text = str(ws.range(f"C{r}").value or "").strip()
            d_text = str(ws.range(f"D{r}").value or "").strip()
            lines.append(f"{c_text} {d_text}".strip())
        data["SCOPE_OF_WORK"] = "\n".join(lines).strip()

    # ── 5. total price (same column, one row below) ───────────────────────────
    if row_total_price:
        below_val = ws.range(f"C{row_total_price + 1}").value
        if below_val is not None:
            try:
                data["TOTAL_PRICE"] = f"{float(below_val):,.2f}"
            except (TypeError, ValueError):
                data["TOTAL_PRICE"] = below_val  # fallback as-is

    data["_SOURCE_FILE"] = wb_path

    return data


if __name__ == "__main__":
    extracted = extract_proposal_data()
    print("\n=== EXTRACTED PROPOSAL DATA ===")
    for key, val in extracted.items():
        print(f"{key}:\n{val}\n")
