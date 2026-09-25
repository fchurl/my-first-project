"""
Script 1: Parse wine recommendations out of Telegraph articles you've saved
yourself, and record them in wines.xlsx.

Why this is manual-input rather than fetching the article automatically:
telegraph.co.uk sits behind Akamai bot-management, which blocks automated
browser traffic outright (independent of whether you're logged in) - so
there's no reliable way to script the fetch without deliberately evading
that protection, which this tool won't do. Reading the article yourself in
your normal browser and saving the text is a 10-second manual step; parsing
it into structured Excel rows is the part worth automating, and that's what
this script does.

How to use it
--------------
1. Open the wine article on telegraph.co.uk in your normal browser (your
   subscription login works as usual - you're just reading the page, not
   scripting it).
2. Select the article text (Ctrl+A on the article body, or just the
   relevant paragraphs) and copy it (Ctrl+C).
3. Create a new .txt file inside the `to_parse` folder next to this script
   (any filename). Paste the article text in. Optionally add these as the
   very first lines, each on its own line:
       URL: https://www.telegraph.co.uk/wine/...
       Date: 2026-09-15
   (Both optional - URL helps de-dup if you ever re-save the same article;
   Date should be the article's publish/review date. If omitted, today's
   date is recorded instead.)
4. Run:
       python parse_telegraph_articles.py
   It parses every .txt file in `to_parse`, adds new wine rows to
   wines.xlsx, and moves processed files into `to_parse/processed/`.

You can also parse a single specific file without moving anything:
       python parse_telegraph_articles.py path\\to\\article.txt
"""
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from config import TO_PARSE_DIR, PROCESSED_DIR
from common import get_workbook, save_workbook, find_existing_row, append_wine_row, parse_article, today_iso

URL_LINE_RE = re.compile(r"^URL:\s*(\S+)", re.I)
DATE_LINE_RE = re.compile(r"^Date:\s*(\S+)", re.I)


def parse_file(path):
    raw = path.read_text(encoding="utf-8", errors="replace")
    lines = raw.splitlines()

    source_url = f"local:{path.name}"
    review_date = today_iso()
    body_start = 0

    for i, line in enumerate(lines[:5]):
        m = URL_LINE_RE.match(line.strip())
        if m:
            source_url = m.group(1)
            body_start = i + 1
            continue
        m = DATE_LINE_RE.match(line.strip())
        if m:
            review_date = m.group(1)
            body_start = i + 1

    body = "\n".join(lines[body_start:])
    return source_url, review_date, body


def process_file(ws, path):
    source_url, review_date, body = parse_file(path)
    entries = parse_article(body, source_url)

    added = 0
    for e in entries:
        if find_existing_row(ws, e["wine_name"], source_url):
            continue
        append_wine_row(
            ws,
            wine_name=e["wine_name"],
            review_date=review_date,
            rating=e["rating"],
            price=e["price"],
            retailer=e["retailer"],
            source_url=source_url,
            needs_review=e["needs_review"],
            note=e["note"],
        )
        added += 1
    return added


def main():
    if len(sys.argv) > 1:
        # Single-file mode: parse without touching the to_parse folder.
        from pathlib import Path
        path = Path(sys.argv[1])
        if not path.exists():
            print(f"File not found: {path}")
            sys.exit(1)
        wb, ws = get_workbook()
        added = process_file(ws, path)
        if added:
            save_workbook(wb)
        print(f"Added {added} wine row(s) from {path.name}.")
        return

    TO_PARSE_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    files = sorted(p for p in TO_PARSE_DIR.glob("*.txt") if p.is_file())
    if not files:
        print(f"No .txt files found in {TO_PARSE_DIR}")
        print("Save an article's text there (see the instructions at the top of this file) and re-run.")
        return

    wb, ws = get_workbook()
    total_added = 0
    for path in files:
        print(f"Parsing {path.name}...")
        added = process_file(ws, path)
        total_added += added
        print(f"  {added} wine row(s) added.")
        path.rename(PROCESSED_DIR / path.name)

    if total_added:
        save_workbook(wb)
    print(f"Done. {total_added} new wine row(s) added to wines.xlsx across {len(files)} file(s).")


if __name__ == "__main__":
    main()
