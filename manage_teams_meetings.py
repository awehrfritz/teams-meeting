"""Lists and deletes Teams meetings created by make_teams_meeting.py.

Microsoft Graph can't list a user's online meetings; it can only look up a meeting
by its ID, join link or meeting ID. make_teams_meeting.py therefore records every
meeting it creates in a local log, which this script reads.
"""
import os
import sys
import json
import argparse
from datetime import datetime
from urllib.parse import quote

from make_teams_meeting import GRAPH_ONLINE_MEETINGS_URL, MEETINGS_LOG_FILE, get_authenticated_session


def read_log():
    """Returns the logged meetings, oldest first."""
    if not os.path.exists(MEETINGS_LOG_FILE):
        return []
    with open(MEETINGS_LOG_FILE) as f:
        return [json.loads(line) for line in f if line.strip()]


def write_log(meetings):
    os.makedirs(os.path.dirname(MEETINGS_LOG_FILE), mode=0o700, exist_ok=True)
    fd = os.open(MEETINGS_LOG_FILE, os.O_WRONLY | os.O_TRUNC | os.O_CREAT, 0o600)
    with os.fdopen(fd, "w") as f:
        for meeting in meetings:
            f.write(json.dumps(meeting) + "\n")


def local_time(timestamp):
    """Formats a Graph UTC timestamp in local time."""
    if not timestamp:
        return "?"
    return datetime.fromisoformat(timestamp).astimezone().strftime("%Y-%m-%d %H:%M")


def meeting_url(meeting_id):
    return f"{GRAPH_ONLINE_MEETINGS_URL}/{quote(meeting_id, safe='')}"


def get_meeting(session, meeting_id):
    """Returns the meeting from Graph, or None if it no longer exists."""
    response = session.get(meeting_url(meeting_id))
    if response.status_code == 404:
        return None
    if response.status_code != 200:
        raise Exception(f"API Error ({response.status_code}):\n{json.dumps(response.json(), indent=2)}")
    return response.json()


def find_meeting_by_join_url(session, join_url):
    """Looks up a meeting by its join link; returns None if not found."""
    response = session.get(GRAPH_ONLINE_MEETINGS_URL, params={"$filter": f"JoinWebUrl eq '{join_url}'"})
    # Graph answers an unknown join link with 404 rather than an empty list
    if response.status_code == 404:
        return None
    if response.status_code != 200:
        raise Exception(f"API Error ({response.status_code}):\n{json.dumps(response.json(), indent=2)}")
    meetings = response.json().get("value", [])
    return meetings[0] if meetings else None


def delete_meeting(session, meeting_id):
    """Deletes a meeting; returns False if it didn't exist (anymore)."""
    response = session.delete(meeting_url(meeting_id))
    if response.status_code == 404:
        return False
    if response.status_code != 204:
        raise Exception(f"API Error ({response.status_code}):\n{json.dumps(response.json(), indent=2)}")
    return True


def cmd_list(args):
    meetings = read_log()
    if not meetings:
        print(f"No meetings recorded in {MEETINGS_LOG_FILE}", file=sys.stderr)
        return
    session = get_authenticated_session() if args.check else None
    for number, meeting in enumerate(meetings, 1):
        line = f"{number:3}  {local_time(meeting.get('startDateTime'))}  {meeting.get('subject')}"
        if session:
            current = get_meeting(session, meeting["id"])
            line += f"  [expires {local_time(current.get('expiryDateTime'))}]" if current else "  [DELETED/EXPIRED]"
        print(line)
        if args.verbose:
            print(f"     Meeting ID: {meeting.get('joinMeetingId')}  Passcode: {meeting.get('passcode') or '-'}")
            print(f"     Join: {meeting.get('joinShortUrl') or meeting.get('joinWebUrl')}")
            print(f"     Options: {meeting.get('meetingOptionsWebUrl')}")


def cmd_delete(args):
    meetings = read_log()
    session = get_authenticated_session()
    remove = set()  # log indices to drop

    for target in args.meetings:
        if target.startswith("https://"):
            # Join link, which also works for meetings that aren't in the log
            meeting = find_meeting_by_join_url(session, target)
            if meeting is None:
                print(f"Not found: {target}", file=sys.stderr)
                continue
            meeting_id, subject = meeting["id"], meeting.get("subject")
            remove.update(i for i, m in enumerate(meetings) if m.get("id") == meeting_id)
        else:
            try:
                index = int(target) - 1
                if not 0 <= index < len(meetings):
                    raise ValueError
            except ValueError:
                print(f"Invalid meeting number or join link: {target}", file=sys.stderr)
                continue
            meeting_id, subject = meetings[index]["id"], meetings[index].get("subject")
            remove.add(index)

        if delete_meeting(session, meeting_id):
            print(f"Deleted: {subject}", file=sys.stderr)
        else:
            print(f"Already deleted or expired: {subject}", file=sys.stderr)

    if remove:
        write_log([m for i, m in enumerate(meetings) if i not in remove])


def cmd_prune(args):
    meetings = read_log()
    session = get_authenticated_session()
    kept = []
    for meeting in meetings:
        if get_meeting(session, meeting["id"]):
            kept.append(meeting)
        else:
            print(f"Removing from log (deleted/expired): {meeting.get('subject')}", file=sys.stderr)
    write_log(kept)


def parse_args():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)

    list_parser = commands.add_parser("list", help="list recorded meetings")
    list_parser.add_argument("-c", "--check", action="store_true",
                             help="ask Graph whether each meeting still exists and when it expires")
    list_parser.add_argument("-v", "--verbose", action="store_true",
                             help="also show meeting ID, passcode and links")
    list_parser.set_defaults(func=cmd_list)

    delete_parser = commands.add_parser("delete", help="delete meetings")
    delete_parser.add_argument("meetings", nargs="+", metavar="NUMBER_OR_LINK",
                               help="number from 'list', or a join link (also for meetings not in the log)")
    delete_parser.set_defaults(func=cmd_delete)

    prune_parser = commands.add_parser("prune", help="remove deleted/expired meetings from the log")
    prune_parser.set_defaults(func=cmd_prune)

    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    try:
        args.func(args)
    except Exception as e:
        print(e, file=sys.stderr)
        sys.exit(1)
