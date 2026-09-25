#!/usr/bin/env python3
"""
Add Hampstead & Westminster M1 HOME fixtures (England Hockey website) to a
Google Calendar, so you can plan to go and watch them play.

Only fixtures where M1 are the home team are added - away fixtures are
skipped. See eh_hockey_lib.py for how the England Hockey fixtures API is
discovered/called, and how Google Calendar events are created/updated.

Setup: same as hockey_calendar_sync.py (pip install -r requirements.txt,
Google OAuth credentials.json). If you've already run hockey_calendar_sync.py
and authorized once, this script reuses the same token.json/credentials.json.

Usage:
   python hockey_m1_home_watch_sync.py --dry-run     # preview
   python hockey_m1_home_watch_sync.py                # write to your calendar
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

DEFAULT_TEAM_PAGE_URL = "https://www.englandhockey.co.uk/teams/hampstead---westminster-m1-mens"
DEFAULT_TEAM_ID = "e6c5e1f7-bd25-4489-94bd-637c27c4dea8"  # Hampstead & Westminster M1
TEAM_LABEL = "H&W M1"
SOURCE_TAG = "hockey_m1_home_watch_sync"


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
    fixtures = [f for f in fixtures if f["homeTeamId"] == args.team_id]  # home only

    if not args.include_past:
        today = date.today()
        fixtures = [
            f for f in fixtures
            if datetime.fromisoformat(f["fixtureDate"]).date() >= today
        ]

    fixtures.sort(key=lambda f: f["fixtureDate"])

    if not fixtures:
        print("No upcoming home fixtures found.")
        return

    print(f"Found {len(fixtures)} home fixture(s) for the current season.")

    def make_event(f):
        return build_event(
            f, args.team_id, args.duration_minutes, args.team_url,
            team_label=TEAM_LABEL, source_tag=SOURCE_TAG,
            summary_template="Watch: {team_label} vs {opponent}",
            extra_description_lines=["Home match - go and support the team!"],
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
