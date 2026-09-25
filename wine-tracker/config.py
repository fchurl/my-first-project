"""
Shared, user-editable settings.
"""
from pathlib import Path

BASE_DIR = Path(__file__).parent

# Excel file the parser writes to. Safe to open in Excel between runs, just
# close it before running the script since it needs to save to the file.
EXCEL_PATH = BASE_DIR / "wines.xlsx"

# Drop saved Telegraph article text files (.txt) here to have them parsed.
# See README.md for the exact workflow.
TO_PARSE_DIR = BASE_DIR / "to_parse"

# Processed files get moved here so re-running the script doesn't add
# duplicate rows.
PROCESSED_DIR = TO_PARSE_DIR / "processed"

# UK retailers the parser looks for when scanning article text. Add to this
# list if a Telegraph review mentions a retailer that isn't being picked up.
KNOWN_RETAILERS = [
    "Sainsbury's", "Sainsburys", "Tesco", "Waitrose", "Majestic",
    "Marks & Spencer", "M&S", "Aldi", "Lidl", "Asda", "Morrisons",
    "Co-op", "Naked Wines", "The Wine Society", "Laithwaites",
    "Virgin Wines", "Ocado", "Amazon",
]

# --- Sainsbury's price checker (via Trolley.co.uk) --------------------------

# Seconds to wait between checking each wine, to be a reasonable, low-volume
# visitor rather than hammering the site.
TROLLEY_DELAY_SECONDS = 4
