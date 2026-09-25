# Telegraph Wine Tracker

## Why this isn't quite the "two scrapers" originally planned

Both telegraph.co.uk and sainsburys.co.uk sit behind Akamai bot-management,
which blocks automated browser traffic outright - independent of being
logged in. Getting past that would mean deliberately engineering evasion
(fingerprint spoofing, etc.), which isn't something this tool does, even
for legitimate personal/subscriber use. So:

- **Telegraph articles**: no automated fetching. `parse_telegraph_articles.py`
  parses article text you've saved yourself (see below) - useful if you
  ever want to bulk-extract a column's worth of wines, but for a single
  wine here and there it's often just as fast to add the row to
  `wines.xlsx` by hand.
- **Retailer prices**: `check_sainsburys_prices.py` and
  `check_tesco_prices.py` don't touch sainsburys.co.uk or tesco.com at all.
  Both go via [Trolley.co.uk](https://www.trolley.co.uk), an independent UK
  grocery price-comparison site that already aggregates prices across
  retailers and isn't behind the same bot-wall. Its `robots.txt` disallows
  automating its `/search/` pages, but doesn't restrict revisiting a
  specific product page you've already found - so that one manual step
  (finding the product on Trolley) stays yours, and the repeated part
  (checking prices) is scripted.

## Setup

```
pip install -r requirements.txt
python -m playwright install chromium
```

(`playwright install` is only needed for the price checkers -
`parse_telegraph_articles.py` doesn't use a browser at all.)

## `wines.xlsx` - your wine list

This is the file all the scripts read and write. If it doesn't exist yet,
running any price-checker script creates it with the right headers (and
one example row - delete it once you've added your own wines). Running
`check_tesco_prices.py` after `check_sainsburys_prices.py` (or vice versa)
just adds its own columns alongside the other's - neither touches the
other's data.

Columns:

| Column | Who fills it in |
|---|---|
| Wine Name | You |
| Trolley Product URL | You (see below) - shared by both retailer checkers |
| Current Price (Sainsbury's) / (Tesco) | `check_sainsburys_prices.py` / `check_tesco_prices.py` |
| Discount Available (Sainsbury's) / (Tesco) | same |
| Discount Details (Sainsbury's) / (Tesco) | same |
| Last Checked (Sainsbury's) / (Tesco) | same |
| Notes (Sainsbury's) / (Tesco) | same (e.g. if it's not stocked there, or the URL looks wrong) |

**To fill in "Trolley Product URL" for a wine:** go to
[trolley.co.uk](https://www.trolley.co.uk), search for it yourself
(normal browsing - not scripted), open the matching product page, and
paste its URL (looks like `https://www.trolley.co.uk/product/<slug>/<id>`)
into that column. This is the one manual step per wine; after that, both
scripts re-check it automatically on every run - the same URL is used for
both, since a Trolley product page lists every retailer that stocks it.

## Running the price checkers

```
python check_sainsburys_prices.py
python check_tesco_prices.py
```

For every row with both a Wine Name and a Trolley Product URL, each script:
- finds its retailer in that product's "Where To Buy" list on Trolley,
- records the current price,
- flags **Discount Available = Yes** if either:
  - Trolley shows a promo badge on that retailer's listing (e.g. "NECTAR"
    or "CLUBCARD" pricing), or
  - the price has dropped since the last time you ran the script (it
    compares against whatever was already recorded for that retailer).

Rows where the wine isn't listed at that retailer on Trolley get a note
instead of a price - it might be out of stock, delisted, or just not one
Trolley tracks for that store. It's normal for a wine to show up at one
retailer and not the other.

Run either as often as you like; `TROLLEY_DELAY_SECONDS` in `config.py`
(default 4s) paces the requests between wines so each behaves like a slow,
normal visitor rather than hammering the site. Keep runs infrequent (e.g.
once a day) - this is for your own personal price-watching, not bulk data
collection, and Trolley is a free service kept running by affiliate
commissions, not built to serve high-volume scraping.

Adding another retailer later (e.g. Waitrose) just means copying
`check_tesco_prices.py` and changing `STORE_KEYWORD` / `STORE_LABEL` at the
top - the shared logic lives in `price_checker_common.py`.

## `parse_telegraph_articles.py` (optional)

If you ever want to bulk-import wines from a Telegraph column instead of
typing them into `wines.xlsx` by hand:

1. Open the article on telegraph.co.uk in your normal browser (your
   subscription login works as usual).
2. Select and copy the article's text.
3. Create a `.txt` file inside `to_parse/` (created on first run) and paste
   the text in. Optionally add these as the first lines:
   ```
   URL: https://www.telegraph.co.uk/wine/the-article-url
   Date: 2026-09-15
   ```
4. Run `python parse_telegraph_articles.py`. It parses every `.txt` file in
   `to_parse/`, adds new wine rows to `wines.xlsx`, and moves processed
   files into `to_parse/processed/`.

The parsing is heuristic (regex/keyword scanning, not real understanding of
the text), so every row gets a `Needs Review` flag when it's unsure - check
those in Excel. This script writes different columns (Rating, Retailer,
Review Date, etc.) than the price checker uses; both scripts happily share
the same `wines.xlsx`, they just care about different columns.

If you don't plan to use this, you can ignore/delete
`parse_telegraph_articles.py`, `common.py`, and `to_parse/`.

## Files

| File | Purpose |
|---|---|
| `config.py` | Delay/pacing settings, retailer list |
| `price_checker_common.py` | Shared logic for the retailer price checkers |
| `check_sainsburys_prices.py` | Checks Sainsbury's prices via Trolley.co.uk |
| `check_tesco_prices.py` | Checks Tesco prices via Trolley.co.uk |
| `parse_telegraph_articles.py` | Optional - bulk-parses saved Telegraph article text |
| `common.py` | Excel schema + parsing heuristics used only by `parse_telegraph_articles.py` |
| `wines.xlsx` | Your wine list / the output - created on first run |
| `to_parse/` | Drop saved article `.txt` files here (optional workflow) |
