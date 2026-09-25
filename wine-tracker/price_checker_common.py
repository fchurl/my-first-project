"""
Shared logic for the per-retailer price checkers (check_sainsburys_prices.py,
check_tesco_prices.py, ...). Each of those scripts owns its own columns in
wines.xlsx (e.g. "Current Price (Tesco)") and just tells this module which
store name to look for on a given Trolley.co.uk product page.
"""
import re
from openpyxl import Workbook, load_workbook
from playwright.sync_api import TimeoutError as PWTimeout

PRICE_RE = re.compile(r"£\s?(\d+(?:\.\d{1,2})?)")


def get_workbook(excel_path, base_headers, example_row=None):
    """
    Load wines.xlsx, adding any of base_headers that aren't already present
    as columns (appended at the end) without disturbing existing columns/
    data owned by other scripts. Creates the file with base_headers (plus
    an optional example row) if it doesn't exist yet.
    """
    if excel_path.exists():
        wb = load_workbook(excel_path)
        ws = wb.active
        header = [c.value for c in ws[1]] if ws.max_row >= 1 else []
        for h in base_headers:
            if h not in header:
                ws.cell(row=1, column=len(header) + 1).value = h
                header.append(h)
        return wb, ws, {name: i + 1 for i, name in enumerate(header)}

    wb = Workbook()
    ws = wb.active
    ws.title = "Wines"
    ws.append(base_headers)
    if example_row:
        ws.append(example_row)
    return wb, ws, {name: i + 1 for i, name in enumerate(base_headers)}


def find_store_row(page, url, store_keyword):
    """
    Visit a Trolley.co.uk product page and return {store, price, offer} for
    the first "Where To Buy" row whose store name contains store_keyword
    (case-insensitive), or None if that store isn't listed for this product.
    """
    page.goto(url, wait_until="domcontentloaded")
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except PWTimeout:
        pass

    items = page.eval_on_selector_all(
        ".comparison-table ._item",
        """els => els.map(el => {
            const svg = el.querySelector('svg[title]');
            const priceEl = el.querySelector('._price b');
            const offerEl = el.querySelector('._product-offer');
            return {
                store: svg ? svg.getAttribute('title') : null,
                price: priceEl ? priceEl.textContent.trim() : null,
                offer: offerEl ? offerEl.textContent.trim() : null,
            };
        })"""
    )

    for item in items:
        store = (item.get("store") or "").lower()
        if store_keyword in store:
            return item
    return None
