#!/usr/bin/env python3
"""
Sync Hampstead & Westminster Masters hockey fixtures (England Hockey website)
into a Google Calendar.

See eh_hockey_lib.py for how the England Hockey fixtures API is discovered
and called, and how Google Calendar events are created/updated.

Setup
-----
1. pip install -r requirements.txt

2. Create a Google OAuth client so this script can write to your calendar:
   - Go to https://console.cloud.google.com/apis/credentials
   - Create a project (or pick an existing one), enable the "Google Calendar API"
   - Create OAuth client credentials of type "Desktop app"
   - Download the JSON and save it next to this script as credentials.json

3. Run it:
   python hockey_calendar_sync.py --dry-run     # preview what would be added
   python hockey_calendar_sync.py                # actually write to your calendar

   The first real run opens a browser to authorize access; a token.json is
   then cached so you won't be prompted again.

4. Optional: schedule it (e.g. Windows Task Scheduler, cron) to run weekly so
   newly-confirmed kick-off times/venues stay in sync.
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, date

import requests

from eh_hockey_lib import (
    FixtureFetchError,
    build_event,
    fetch_fixtures,
    get_calendar_service,
    is_real_match,
    upsert_event,
)

DEFAULT_TEAM_PAGE_URL = (
    "https://www.englandhockey.co.uk/teams/hampstead---westminster-masters-mens"
)
DEFAULT_TEAM_ID = "d04f8a35-6542-4e50-b606-eaab342373f6"  # Hampstead & Westminster Masters
TEAM_LABEL = "Hampstead & Westminster Masters"
SOURCE_TAG = "hockey_calendar_sync"


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--team-url", default=DEFAULT_TEAM_PAGE_URL, help="England Hockey team page URL")
    parser.add_argument("--team-id", default=DEFAULT_TEAM_ID, help="England Hockey team GUID")
    parser.add_argument("--calendar-id", default="primary", help="Google Calendar ID to write to")
    parser.add_argument("--credentials", default="credentials.json", help="Path to Google OAuth client credentials JSON")
    parser.add_argument("--token", default="token.json", help="Path to cached OAuth token")
    parser.add_argument("--duration-minutes", type=int, default=120, help="Event length for fixtures with a confirmed kick-off time")
    parser.add_argument("--include-past", action="store_true", help="Also add/update fixtures already played (default: only today onward)")
    parser.add_argument("--dry-run", action="store_true", help="Print what would happen without touching Google Calendar")
    args = parser.parse_args()

    try:
        fixtures = fetch_fixtures(args.team_url)
    except (requests.RequestException, FixtureFetchError) as exc:
        print(f"Failed to fetch fixtures: {exc}", file=sys.stderr)
        sys.exit(1)

    fixtures = [f for f in fixtures if is_real_match(f, args.team_id)]

    if not args.include_past:
        today = date.today()
        fixtures = [
            f for f in fixtures
            if datetime.fromisoformat(f["fixtureDate"]).date() >= today
        ]

    fixtures.sort(key=lambda f: f["fixtureDate"])

    if not fixtures:
        print("No matching fixtures found.")
        return

    print(f"Found {len(fixtures)} fixture(s) for the current season.")

    def make_event(f):
        return build_event(
            f, args.team_id, args.duration_minutes, args.team_url,
            team_label=TEAM_LABEL, source_tag=SOURCE_TAG,
            summary_template="Hockey ({ha}) vs {opponent}",
        )

    if args.dry_run:
        for f in fixtures:
            event = make_event(f)
            when = event["start"].get("dateTime") or event["start"].get("date")
            print(f"  [DRY RUN] {when}  {event['summary']}  @ {event['location']}")
        return

    service = get_calendar_service(args.credentials, args.token)

    created = updated = 0
    for f in fixtures:
        event = make_event(f)
        result = upsert_event(service, args.calendar_id, event)
        when = event["start"].get("dateTime") or event["start"].get("date")
        print(f"  {result:7s} {when}  {event['summary']}")
        if result == "created":
            created += 1
        else:
            updated += 1

    print(f"\nDone. {created} event(s) created, {updated} updated.")


if __name__ == "__main__":
    main()
