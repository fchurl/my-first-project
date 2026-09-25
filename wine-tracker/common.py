"""
Shared helpers used by both telegraph_wine_scraper.py and
sainsburys_price_checker.py: the Excel schema, and best-effort text
parsing for pulling wine/price/rating/retailer out of free-form article text.

The parsing here is heuristic (regex + keyword matching), not a proper NLP
model, because Telegraph articles aren't structured data - wording varies
article to article. Anything it isn't confident about gets written with
Needs Review = "Yes" so you can eyeball and fix it in Excel rather than the
script silently guessing wrong.
"""
import re
from datetime import date
from difflib import SequenceMatcher

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from config import EXCEL_PATH, KNOWN_RETAILERS

COLUMNS = [
    "Wine Name",
    "Review Date",
    "Rating (Buy/Avoid)",
    "Telegraph Price",
    "Retailer",
    "Source URL",
    "Needs Review",
    "Extraction Note",
    "Sainsbury's Product URL",
    "Current Price",
    "Discount Available",
    "Discount Details",
    "Price Last Checked",
]

PRICE_RE = re.compile(r"£\s?(\d+(?:\.\d{1,2})?)")
RATING_KEYWORDS = {
    "buy": ["buy now", "buy it", "well worth buying", "highly recommend", "buy:", " buy "],
    "avoid": ["avoid", "skip this", "steer clear", "don't bother", "disappointing"],
}
STAR_RE = re.compile(r"[★☆]{2,5}")
SCORE_RE = re.compile(r"\b(\d{1,2})\s?/\s?(10|20)\b")
YEAR_RE = re.compile(r"\b(19|20)\d{2}\b")


def _name_score(line):
    """Heuristic: does this line look like a wine name/heading rather than
    a prose sentence? Higher is more name-like."""
    score = 0
    if YEAR_RE.search(line):
        score += 2  # vintage year is a strong signal
    words = line.split()
    if words:
        cap_ratio = sum(1 for w in words if w[:1].isupper()) / len(words)
        if cap_ratio > 0.6:
            score += 2  # Title Case heading
    if line.endswith((".", "!", "?")):
        score -= 2  # reads like a finished sentence, not a heading
    if len(line) > 60:
        score -= 1
    return score


def get_workbook():
    if EXCEL_PATH.exists():
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Wines"
        ws.append(COLUMNS)
        for i, col in enumerate(COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(i)].width = max(14, len(col) + 2)
    return wb, ws


def save_workbook(wb):
    EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(EXCEL_PATH)


def row_dicts(ws):
    """Yield (row_index, dict) for every data row, 1-indexed to match openpyxl."""
    header = [c.value for c in ws[1]]
    for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        values = [c.value for c in row]
        yield idx, dict(zip(header, values))


def find_existing_row(ws, wine_name, source_url):
    """Dedup key: same source article + same wine name already recorded."""
    for idx, d in row_dicts(ws):
        if d.get("Source URL") == source_url and (d.get("Wine Name") or "").strip().lower() == wine_name.strip().lower():
            return idx
    return None


def append_wine_row(ws, *, wine_name, review_date, rating, price, retailer,
                     source_url, needs_review, note):
    row = {
        "Wine Name": wine_name,
        "Review Date": review_date,
        "Rating (Buy/Avoid)": rating,
        "Telegraph Price": price,
        "Retailer": retailer,
        "Source URL": source_url,
        "Needs Review": "Yes" if needs_review else "No",
        "Extraction Note": note,
        "Sainsbury's Product URL": None,
        "Current Price": None,
        "Discount Available": None,
        "Discount Details": None,
        "Price Last Checked": None,
    }
    ws.append([row[c] for c in COLUMNS])


def similarity(a, b):
    return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()


def detect_retailer(text):
    for retailer in KNOWN_RETAILERS:
        if retailer.lower() in text.lower():
            return retailer
    return None


def detect_rating(text):
    lower = text.lower()
    for word in RATING_KEYWORDS["avoid"]:
        if word in lower:
            return "Avoid"
    for word in RATING_KEYWORDS["buy"]:
        if word in lower:
            return "Buy"
    stars = STAR_RE.search(text)
    if stars:
        filled = stars.group(0).count("★")
        return "Buy" if filled >= 3 else "Avoid"
    score = SCORE_RE.search(text)
    if score:
        num, denom = int(score.group(1)), int(score.group(2))
        return "Buy" if (num / denom) >= 0.6 else "Avoid"
    return None


def parse_article(text, url):
    """
    Best-effort split of an article into one or more wine entries.

    Strategy: scan paragraph-by-paragraph for a price mention. A paragraph
    (or a couple of lines around it) containing a price is treated as one
    wine's write-up; the preceding non-empty line(s) are assumed to hold the
    wine's name (Telegraph wine columns conventionally lead each entry with
    the wine name/vintage as a short heading line).

    Falls back to a single "Needs Review" row with the whole article if no
    price is found at all, so nothing silently gets dropped.
    """
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    entries = []
    used_name_indexes = set()

    for i, line in enumerate(lines):
        price_match = PRICE_RE.search(line)
        if not price_match:
            continue

        # context window: a couple of lines before/after for rating+retailer
        window = " ".join(lines[max(0, i - 2): i + 3])

        # candidate name = best-scoring preceding line (heading-like beats
        # prose-like) within a small window, not already used and not a
        # price line itself
        best_j, best_score = None, None
        for j in range(i - 1, max(-1, i - 6), -1):
            if j in used_name_indexes:
                continue
            candidate = lines[j]
            if PRICE_RE.search(candidate):
                continue
            if not (3 <= len(candidate) <= 80):
                continue
            score = _name_score(candidate)
            if best_score is None or score > best_score:
                best_j, best_score = j, score

        if best_j is None:
            continue  # can't anchor this price to a wine name, skip
        name = lines[best_j]
        used_name_indexes.add(best_j)

        retailer = detect_retailer(window)
        rating = detect_rating(window)
        price = f"£{price_match.group(1)}"

        needs_review = not (retailer and rating)
        note_bits = []
        if not retailer:
            note_bits.append("retailer not detected")
        if not rating:
            note_bits.append("rating not detected")
        note = "; ".join(note_bits) if note_bits else ""

        entries.append(dict(
            wine_name=name,
            price=price,
            retailer=retailer or "",
            rating=rating or "",
            needs_review=needs_review,
            note=note,
        ))

    if not entries:
        entries.append(dict(
            wine_name="(unparsed article - see note)",
            price="",
            retailer="",
            rating="",
            needs_review=True,
            note=f"No price found by automated parsing; review manually at {url}",
        ))

    return entries


def today_iso():
    return date.today().isoformat()
