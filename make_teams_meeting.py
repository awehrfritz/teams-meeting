import os
import re
import sys
import json
import html
import argparse
import requests
from datetime import date, datetime, time, timedelta, timezone
import msal
from urllib.parse import quote, unquote_plus

# Configuration
# Client ID of "Microsoft Graph Command Line Tools", a first-party Microsoft public client app
# (the one used by Connect-MgGraph). Your tenant may require admin consent before it can be used.
# Reference: https://learn.microsoft.com/en-us/troubleshoot/entra/entra-id/governance/verify-first-party-apps-sign-in
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
# "organizations" accepts any work or school account (online meetings require one)
AUTHORITY = "https://login.microsoftonline.com/organizations"
# Delegated permissions required by POST /me/onlineMeetings and, for the default subject, GET /me
# Reference: https://learn.microsoft.com/en-us/graph/api/application-post-onlinemeetings
SCOPES = ["OnlineMeetings.ReadWrite", "User.Read"]
GRAPH_ME_URL = "https://graph.microsoft.com/v1.0/me"
GRAPH_ONLINE_MEETINGS_URL = "https://graph.microsoft.com/v1.0/me/onlineMeetings"
# File locations follow the XDG Base Directory spec: the token cache can be deleted at any time
# (it only forces a new sign-in), so it goes to the cache directory; the meeting log is user data.
# Reference: https://specifications.freedesktop.org/basedir-spec/latest/
TOKEN_CACHE_FILE = os.path.join(os.environ.get("XDG_CACHE_HOME") or os.path.expanduser("~/.cache"),
                                "teams-meeting", "token_cache.bin")
# Graph can't list a user's meetings, so every created meeting is recorded here (one JSON object per line)
MEETINGS_LOG_FILE = os.path.join(os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share"),
                                 "teams-meeting", "meetings.jsonl")
# Used when no subject is given and the user's name isn't known; with the name, the default is
# "Meeting with <name>", like Teams' own "Meet now"
FALLBACK_SUBJECT = "Teams meeting"

def get_authenticated_session():
    """Handles token generation and caching using Device Code Flow."""
    # 1. Set up a persistent local token cache so you don't have to log in every time
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache
    )

    # 2. Try to get a valid token quietly from the local cache
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    # 3. If cache is empty or expired, trigger the Device Code login process
    if not result:
        print("No valid cached token found. Starting interactive login...", file=sys.stderr)
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise Exception(f"Could not initiate device flow: {json.dumps(flow, indent=2)}")

        # Print authentication instructions directly to the user's terminal
        print("\n" + "="*70, file=sys.stderr)
        print(flow["message"], file=sys.stderr)
        print("="*70 + "\n", file=sys.stderr)

        # Block and wait for the user to complete the authentication in their browser
        result = app.acquire_token_by_device_flow(flow)

    # 4. Save the updated token to the cache directory
    if cache.has_state_changed:
        # The cache holds a refresh token, so keep it private (O_CREAT's mode only applies to new
        # files, hence the chmod for caches created by earlier versions)
        os.makedirs(os.path.dirname(TOKEN_CACHE_FILE), mode=0o700, exist_ok=True)
        fd = os.open(TOKEN_CACHE_FILE, os.O_WRONLY | os.O_TRUNC | os.O_CREAT, 0o600)
        os.fchmod(fd, 0o600)
        with os.fdopen(fd, "w") as f:
            f.write(cache.serialize())

    if "access_token" in result:
        # Create a pre-configured requests session with the Auth headers
        session = requests.Session()
        session.headers.update({
            "Authorization": f"Bearer {result['access_token']}",
            "Content-Type": "application/json"
        })
        return session
    else:
        error_msg = result.get("error_description", result.get("error", "Unknown Authentication Error"))
        raise Exception(f"Authentication failed: {error_msg}")

def default_subject(session):
    """Returns "Meeting with <your name>", or FALLBACK_SUBJECT if the name can't be fetched."""
    response = session.get(GRAPH_ME_URL, params={"$select": "displayName"})
    name = response.json().get("displayName") if response.ok else None
    return f"Meeting with {name}" if name else FALLBACK_SUBJECT

def _to_graph_time(dt):
    """Formats a datetime as a UTC timestamp for Graph (naive datetimes are taken as local time)."""
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

def join_information_html(meeting):
    """Returns the meeting's invitation block as HTML (as Teams puts it into invitations), or None.

    Graph delivers it as a URL-encoded data URL ("data:text/html,...") in joinInformation.
    """
    content = (meeting.get("joinInformation") or {}).get("content") or ""
    if not content.startswith("data:"):
        return content or None
    return unquote_plus(content.split(",", 1)[1])

def short_join_url(meeting):
    """Returns the short join link (https://teams.microsoft.com/meet/...), or None.

    It's only contained in the invitation HTML; its 'p' parameter is not the passcode,
    so the link can't be built from the other fields.
    """
    match = re.search(r'href="(https://teams\.microsoft\.com/meet/[^"]+)"', join_information_html(meeting) or "")
    return html.unescape(match.group(1)) if match else None

def log_meeting(meeting):
    """Appends the essentials of a created meeting to MEETINGS_LOG_FILE."""
    join_id = meeting.get("joinMeetingIdSettings") or {}
    entry = {
        "id": meeting.get("id"),
        "subject": meeting.get("subject"),
        "startDateTime": meeting.get("startDateTime"),
        "endDateTime": meeting.get("endDateTime"),
        "joinWebUrl": meeting.get("joinWebUrl"),
        "joinShortUrl": short_join_url(meeting),
        "joinMeetingId": join_id.get("joinMeetingId"),
        "passcode": join_id.get("passcode"),
        "meetingOptionsWebUrl": meeting.get("meetingOptionsWebUrl"),
    }
    # The log contains join links and passcodes, so keep it private
    os.makedirs(os.path.dirname(MEETINGS_LOG_FILE), mode=0o700, exist_ok=True)
    fd = os.open(MEETINGS_LOG_FILE, os.O_WRONLY | os.O_APPEND | os.O_CREAT, 0o600)
    with os.fdopen(fd, "a") as f:
        f.write(json.dumps(entry) + "\n")

def create_online_meeting(subject=None, start=None, end=None, settings=None):
    """Creates a Teams meeting and returns the onlineMeeting resource from Graph as a dict.

    subject defaults to "Meeting with <your name>", start to now and end to one hour after start. settings is a dict of further
    onlineMeeting properties (e.g. lobbyBypassSettings); anything left out gets the
    defaults from your organisation's Teams meeting policy.
    Reference: https://learn.microsoft.com/en-us/graph/api/resources/onlinemeeting
    Raises an Exception if authentication or the Graph API request fails.
    """
    session = get_authenticated_session()

    if subject is None:
        subject = default_subject(session)
    if start is None:
        start = datetime.now(timezone.utc)
    if end is None:
        end = start + timedelta(hours=1)

    payload = {
        "startDateTime": _to_graph_time(start),
        "endDateTime": _to_graph_time(end),
        "subject": subject,
        **(settings or {})
    }

    print("Requesting Teams link from Microsoft Graph API...", file=sys.stderr)
    response = session.post(
        GRAPH_ONLINE_MEETINGS_URL,
        json=payload,
        # Needed for newer enum values such as lobby scope "invited" or "organizationExcludingGuests"
        headers={"Prefer": "include-unknown-enum-members"}
    )

    if response.status_code != 201:
        raise Exception(f"API Error ({response.status_code}):\n{json.dumps(response.json(), indent=2)}")

    meeting = response.json()
    if "joinInformation" not in meeting:
        # Not documented as part of the create response, so fetch the meeting if it's missing
        response = session.get(f"{GRAPH_ONLINE_MEETINGS_URL}/{quote(meeting['id'], safe='')}")
        if response.ok:
            meeting = response.json()
    try:
        log_meeting(meeting)
    except OSError as e:
        # The meeting exists either way, so a failed log write is only a warning
        print(f"Warning: could not record meeting in {MEETINGS_LOG_FILE}: {e}", file=sys.stderr)
    return meeting

def parse_start(value):
    """argparse type for --start: a date and time, or just a time meaning today.

    A time alone must still be in the future, which catches typos like '9:00' for '19:00'.
    """
    try:
        date.fromisoformat(value)
    except ValueError:
        pass
    else:
        # A date alone would silently mean midnight
        raise argparse.ArgumentTypeError(f"'{value}' has no time, expected e.g. '{value} 14:00'")
    try:
        start_time = time.fromisoformat(value)
    except ValueError:
        try:
            return datetime.fromisoformat(value)
        except ValueError:
            raise argparse.ArgumentTypeError(f"invalid date/time '{value}', expected e.g. '14:00' or '2026-10-02 14:00'")
    start = datetime.combine(date.today(), start_time)
    if start.astimezone() <= datetime.now().astimezone():
        raise argparse.ArgumentTypeError(f"{value} today is already past; give a date for another day, "
                                         f"e.g. '{date.today() + timedelta(days=1)} {value}'")
    return start

def parse_args():
    parser = argparse.ArgumentParser(
        description="Create a Microsoft Teams meeting and print its join link. "
                    "Settings that are not given use your organisation's defaults.")
    parser.add_argument("-s", "--subject", help="meeting subject (default: 'Meeting with <your name>')")
    parser.add_argument("-t", "--start", type=parse_start,
                        help="start in local time, e.g. '2026-10-02 14:00', or just '14:00' for today (default: now)")
    parser.add_argument("-d", "--duration", type=int, default=60, help="duration in minutes (default: 60)")
    parser.add_argument("-l", "--lobby-bypass",
                        choices=["everyone", "organization", "organizationAndFederated",
                                 "organizationExcludingGuests", "invited", "organizer"],
                        help="who joins directly without waiting in the lobby; 'everyone' disables the lobby, "
                             "'organization' includes guests, 'organizationExcludingGuests' does not")
    parser.add_argument("-p", "--presenters", choices=["everyone", "organization", "organizer"],
                        help="who can present")
    parser.add_argument("-c", "--chat", choices=["enabled", "disabled", "limited"],
                        help="meeting chat ('limited': only during the meeting)")
    parser.add_argument("-r", "--record", action="store_true", help="start recording automatically")
    output = parser.add_mutually_exclusive_group()
    output.add_argument("--short", action="store_true",
                        help="print the short join link (teams.microsoft.com/meet/...) instead of the long one")
    output.add_argument("--html", action="store_true",
                        help="print the invitation block as HTML (link, meeting ID, passcode) instead of the link")
    return parser.parse_args()

def settings_from_args(args):
    """Translates command-line options into onlineMeeting properties."""
    settings = {}
    if args.lobby_bypass:
        settings["lobbyBypassSettings"] = {"scope": args.lobby_bypass}
    if args.presenters:
        settings["allowedPresenters"] = args.presenters
    if args.chat:
        settings["allowMeetingChat"] = args.chat
    if args.record:
        settings["recordAutomatically"] = True
    return settings

if __name__ == "__main__":
    args = parse_args()

    # Naive times (no UTC offset given) are taken as local time
    start = (args.start or datetime.now()).astimezone()
    end = start + timedelta(minutes=args.duration)

    try:
        meeting = create_online_meeting(args.subject, start, end, settings_from_args(args))
    except Exception as e:
        print(e, file=sys.stderr)
        sys.exit(1)

    # Confirmation of what Graph actually stored (UTC timestamps), shown in local time
    meeting_start = datetime.fromisoformat(meeting["startDateTime"]).astimezone()
    meeting_end = datetime.fromisoformat(meeting["endDateTime"]).astimezone()
    end_format = "%H:%M" if meeting_end.date() == meeting_start.date() else "%a %Y-%m-%d %H:%M"
    print(f"Created: {meeting['subject']}, {meeting_start:%a %Y-%m-%d %H:%M}–{meeting_end.strftime(end_format)}",
          file=sys.stderr)

    # Details for joining without the link and for changing settings later go to stderr
    join_id = meeting.get("joinMeetingIdSettings") or {}
    if join_id.get("joinMeetingId"):
        print(f"Meeting ID: {join_id['joinMeetingId']}  Passcode: {join_id.get('passcode') or '-'}", file=sys.stderr)
    if meeting.get("meetingOptionsWebUrl"):
        print(f"Meeting options (organizer only): {meeting['meetingOptionsWebUrl']}", file=sys.stderr)

    # Clean output to stdout so you can safely pipe or copy the link directly
    if args.html or args.short:
        output = join_information_html(meeting) if args.html else short_join_url(meeting)
        if not output:
            print("Graph returned no invitation block; printing the long link instead", file=sys.stderr)
        print(output or meeting["joinWebUrl"])
    else:
        print(meeting["joinWebUrl"])
