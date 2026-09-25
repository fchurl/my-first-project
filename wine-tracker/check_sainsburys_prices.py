"""
Script: For every wine in wines.xlsx, check its current Sainsbury's price
(via its Trolley.co.uk product page) and flag whether it looks discounted.

Why via Trolley.co.uk and not sainsburys.co.uk directly: sainsburys.co.uk
sits behind Akamai bot-management that blocks automated browser traffic
outright, independent of login - there's no reliable way to script around
that without deliberately evading it, which this tool won't do. Trolley is
a legitimate, independent UK price-comparison site that already aggregates
Sainsbury's prices; its robots.txt disallows automating its /search/ pages,
but says nothing against revisiting a specific product page you already
found yourself - so that's what this script does.

See check_tesco_prices.py for the Tesco equivalent - both share the same
wines.xlsx and the same "Trolley Product URL" column, just write to their
own retailer-specific columns.

Setup (one-time, per wine)
---------------------------
You maintain wines.xlsx yourself. It needs "Wine Name" and "Trolley Product
URL" columns. For each wine:
  1. Go to https://www.trolley.co.uk and search for it yourself (normal
     browsing - the script never automates the search).
  2. Open the matching product page and copy its URL
     (looks like https://www.trolley.co.uk/product/<slug>/<id>).
  3. Paste it into the "Trolley Product URL" column for that wine.

Run this script any time after that to refresh prices:
    python check_sainsburys_prices.py
"""
import sys
import time
from datetime import date

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from playwright.sync_api import sync_playwright

from config import EXCEL_PATH, TROLLEY_DELAY_SECONDS
from price_checker_common import get_workbook, find_store_row, PRICE_RE

STORE_KEYWORD = "sainsbury"
STORE_LABEL = "Sainsbury's"

HEADERS = [
    "Wine Name",
    "Trolley Product URL",
    f"Current Price ({STORE_LABEL})",
    f"Discount Available ({STORE_LABEL})",
    f"Discount Details ({STORE_LABEL})",
    f"Last Checked ({STORE_LABEL})",
    f"Notes ({STORE_LABEL})",
]
EXAMPLE_ROW = [
    "(example) Villa Maria Sauvignon Blanc",
    "https://www.trolley.co.uk/product/villa-maria-sauvignon-blanc/YXV486",
    "", "", "", "", "delete this row once you've added your own wines",
]


def main():
    wb, ws, col = get_workbook(EXCEL_PATH, HEADERS, EXAMPLE_ROW)

    rows = []
    skipped = []
    for r in range(2, ws.max_row + 1):
        wine_name = ws.cell(row=r, column=col["Wine Name"]).value
        if not wine_name:
            continue
        trolley_url = ws.cell(row=r, column=col["Trolley Product URL"]).value
        if trolley_url:
            rows.append((r, str(wine_name).strip(), str(trolley_url).strip()))
        else:
            skipped.append(str(wine_name).strip())
            ws.cell(row=r, column=col[f"Notes ({STORE_LABEL})"]).value = (
                "No Trolley Product URL set yet - add one to check this wine."
            )

    if skipped:
        print(f"Skipping {len(skipped)} wine(s) with no Trolley Product URL yet:")
        for name in skipped:
            print(f"  - {name}")

    if not rows:
        wb.save(EXCEL_PATH)
        print(f"\nNo rows with both a Wine Name and a Trolley Product URL found in {EXCEL_PATH.name}.")
        print("Fill those in (see the instructions at the top of this script) and re-run.")
        return

    print(f"\nChecking {len(rows)} wine(s) at {STORE_LABEL}...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        for r, wine_name, trolley_url in rows:
            print(f"- {wine_name}")

            if "/search/" in trolley_url:
                ws.cell(row=r, column=col[f"Notes ({STORE_LABEL})"]).value = (
                    "This looks like a Trolley search-results URL, not a product page. "
                    "Search on trolley.co.uk yourself, open the specific wine, and paste that URL instead."
                )
                continue

            try:
                item = find_store_row(page, trolley_url, STORE_KEYWORD)
            except Exception as e:
                ws.cell(row=r, column=col[f"Notes ({STORE_LABEL})"]).value = f"Error loading page: {e}"
                time.sleep(TROLLEY_DELAY_SECONDS)
                continue

            if item is None:
                ws.cell(row=r, column=col[f"Notes ({STORE_LABEL})"]).value = f"Not currently listed at {STORE_LABEL} on Trolley."
                ws.cell(row=r, column=col[f"Last Checked ({STORE_LABEL})"]).value = date.today().isoformat()
                print(f"    not listed at {STORE_LABEL} on this Trolley page")
                time.sleep(TROLLEY_DELAY_SECONDS)
                continue

            price_match = PRICE_RE.search(item["price"] or "")
            new_price = float(price_match.group(1)) if price_match else None

            prev_cell = ws.cell(row=r, column=col[f"Current Price ({STORE_LABEL})"])
            prev_price = None
            if prev_cell.value:
                m = PRICE_RE.search(str(prev_cell.value))
                if m:
                    prev_price = float(m.group(1))

            discount_bits = []
            if item.get("offer"):
                discount_bits.append(f"Offer: {item['offer']}")
            if new_price is not None and prev_price is not None and new_price < prev_price:
                discount_bits.append(f"Price dropped from £{prev_price:.2f} to £{new_price:.2f}")

            discount_available = bool(discount_bits)

            ws.cell(row=r, column=col[f"Current Price ({STORE_LABEL})"]).value = (
                f"£{new_price:.2f}" if new_price is not None else (item["price"] or "")
            )
            ws.cell(row=r, column=col[f"Discount Available ({STORE_LABEL})"]).value = "Yes" if discount_available else "No"
            ws.cell(row=r, column=col[f"Discount Details ({STORE_LABEL})"]).value = "; ".join(discount_bits)
            ws.cell(row=r, column=col[f"Last Checked ({STORE_LABEL})"]).value = date.today().isoformat()
            ws.cell(row=r, column=col[f"Notes ({STORE_LABEL})"]).value = ""

            if discount_available:
                print(f"    DISCOUNT: {'; '.join(discount_bits)}")
            elif new_price is not None:
                print(f"    price £{new_price:.2f}")
            else:
                print("    price not found on page")

            time.sleep(TROLLEY_DELAY_SECONDS)

        browser.close()

    wb.save(EXCEL_PATH)
    print(f"Done. {EXCEL_PATH.name} updated.")


if __name__ == "__main__":
    main()
