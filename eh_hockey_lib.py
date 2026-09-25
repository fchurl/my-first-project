"""
Shared helpers for syncing England Hockey (englandhockey.co.uk) team fixtures
into Google Calendar. Used by hockey_calendar_sync.py and
hockey_m1_home_watch_sync.py - see those files for usage.
"""

from __future__ import annotations

import re
from datetime import datetime, timedelta

import truststore

truststore.inject_into_ssl()  # englandhockey.co.uk sends an incomplete cert
# chain; Windows/browsers tolerate it via AIA fetching but Python's bundled
# certifi store doesn't, so delegate verification to the OS trust store.

import requests

LONDON_TZ = "Europe/London"

BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "application/json",
    "Origin": "https://www.englandhockey.co.uk",
}

DATA_URL_RE = re.compile(
    r'data-module="competitions-team-fixtures"[^>]*'
    r'data-url="([^"]+)"[^>]*data-url-key="([^"]+)"'
)


class FixtureFetchError(RuntimeError):
    pass


def get_fixtures_api_details(team_page_url: str) -> tuple[str, str]:
    """Scrape the team page for the current fixtures API URL + public API key."""
    resp = requests.get(team_page_url, headers=BROWSER_HEADERS, timeout=20)
    resp.raise_for_status()
    match = DATA_URL_RE.search(resp.text)
    if not match:
        raise FixtureFetchError(
            "Could not find the fixtures API URL on the team page. "
            "England Hockey may have changed their site structure."
        )
    return match.group(1), match.group(2)


def fetch_fixtures(team_page_url: str) -> list[dict]:
    """Return the flat list of fixture dicts for the team, across all competitions."""
    api_url, api_key = get_fixtures_api_details(team_page_url)
    headers = {**BROWSER_HEADERS, "x-api-key": api_key}
    resp = requests.get(api_url, headers=headers, timeout=20)
    resp.raise_for_status()
    groups = resp.json()

    fixtures = []
    for group in groups:
        fixtures.extend(group.get("fixtures", []))
    return fixtures


def is_real_match(fixture: dict, team_id: str) -> bool:
    """Filter out byes / placeholder 'HOLDING Club' entries."""
    if fixture.get("isBye"):
        return False
    home = fixture.get("homeTeam", {})
    away = fixture.get("awayTeam", {})
    if home.get("clubName") == "HOLDING Club" or away.get("clubName") == "HOLDING Club":
        return False
    if fixture.get("homeTeamId") != team_id and fixture.get("awayTeamId") != team_id:
        return False
    return True


def build_event(
    fixture: dict,
    team_id: str,
    duration_minutes: int,
    team_page_url: str,
    team_label: str,
    source_tag: str,
    summary_template: str = "{team_label} ({ha}) vs {opponent}",
    extra_description_lines: list[str] | None = None,
) -> dict:
    home = fixture["homeTeamId"] == team_id
    opponent = (fixture["awayTeam"] if home else fixture["homeTeam"])["teamName"].strip()
    venue = fixture.get("venue") or "Venue TBC"
    fixture_time = fixture.get("fixtureTime") or ""
    time_confirmed = fixture_time not in ("", "00:00")

    start_dt = datetime.fromisoformat(fixture["fixtureDate"])

    summary = summary_template.format(
        team_label=team_label, ha="H" if home else "A", opponent=opponent
    )

    description_lines = [
        f"Competition: {fixture.get('competitionName', 'Unknown')}",
        f"{'Home' if home else 'Away'} fixture for {team_label}",
    ]
    if extra_description_lines:
        description_lines.extend(extra_description_lines)
    if not time_confirmed:
        description_lines.append(
            "Kick-off time not yet confirmed on England Hockey - re-run this "
            "script later to pick up the confirmed time."
        )
    if fixture.get("isResult"):
        description_lines.append(
            f"Result: {fixture.get('homeTeamScoreAsInt')} - {fixture.get('awayTeamScoreAsInt')}"
        )
    description_lines.append(f"Fixture list: {team_page_url}")

    event: dict = {
        "summary": summary,
        "location": venue,
        "description": "\n".join(description_lines),
        "extendedProperties": {
            "private": {
                "ehFixtureId": fixture["id"],
                "ehSource": source_tag,
            }
        },
    }

    if time_confirmed:
        end_dt = start_dt + timedelta(minutes=duration_minutes)
        # Explicit `date: None` clears a stale all-day date left over from an
        # earlier patch (Google Calendar merges start/end fields rather than
        # replacing them wholesale, so date+dateTime can otherwise coexist
        # and get rejected as "Invalid start time").
        event["start"] = {"dateTime": start_dt.isoformat(), "timeZone": LONDON_TZ, "date": None}
        event["end"] = {"dateTime": end_dt.isoformat(), "timeZone": LONDON_TZ, "date": None}
    else:
        day = start_dt.date()
        event["start"] = {"date": day.isoformat(), "dateTime": None, "timeZone": None}
        event["end"] = {"date": (day + timedelta(days=1)).isoformat(), "dateTime": None, "timeZone": None}

    return event


def get_calendar_service(credentials_path: str, token_path: str):
    from google.auth.exceptions import RefreshError
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    scopes = ["https://www.googleapis.com/auth/calendar.events"]
    creds = None
    try:
        creds = Credentials.from_authorized_user_file(token_path, scopes)
    except FileNotFoundError:
        pass

    if not creds or not creds.valid:
        refreshed = False
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                refreshed = True
            except RefreshError:
                pass  # refresh token expired/revoked - fall back to interactive login
        if not refreshed:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, scopes)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as f:
            f.write(creds.to_json())

    return build("calendar", "v3", credentials=creds)


def find_existing_event(service, calendar_id: str, fixture_id: str) -> dict | None:
    resp = (
        service.events()
        .list(
            calendarId=calendar_id,
            privateExtendedProperty=f"ehFixtureId={fixture_id}",
            singleEvents=True,
        )
        .execute()
    )
    items = resp.get("items", [])
    return items[0] if items else None


def upsert_event(service, calendar_id: str, event: dict) -> str:
    fixture_id = event["extendedProperties"]["private"]["ehFixtureId"]
    existing = find_existing_event(service, calendar_id, fixture_id)
    if existing:
        service.events().patch(
            calendarId=calendar_id, eventId=existing["id"], body=event
        ).execute()
        return "updated"
    else:
        service.events().insert(calendarId=calendar_id, body=event).execute()
        return "created"
