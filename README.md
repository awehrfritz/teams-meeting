# Teams meeting

Small command-line tools for creating Microsoft Teams meeting links via the
[Microsoft Graph API](https://learn.microsoft.com/en-us/graph/api/application-post-onlinemeetings).

| Script | What it does |
| --- | --- |
| `make_teams_meeting.py` | Creates a Teams meeting and prints its join link (**primary script**) |
| `manage_teams_meetings.py` | Lists and deletes meetings created by `make_teams_meeting.py` |

All scripts print only the join link to stdout (status messages go to stderr),
so the output can be piped, e.g. `pixi run meeting | wl-copy`.

## Authentication

The scripts sign in to Microsoft 365 with the
[device code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code)
using the first-party "Microsoft Graph Command Line Tools" app, so no app
registration of your own is needed. A work or school account is required, and
your tenant may require admin consent for this app.

On first use, the script prints a URL and a code; open the URL in a browser,
enter the code and sign in. The tokens are then cached, so later runs work
without signing in again until the refresh token expires. The cache is
`~/.cache/teams-meeting/token_cache.bin` (MSAL cache); delete it to force a new
sign-in. The locations follow the
[XDG Base Directory](https://specifications.freedesktop.org/basedir-spec/latest/)
convention (`$XDG_CACHE_HOME` and `$XDG_DATA_HOME` are respected): the token
cache is disposable, so it lives in the cache directory, while the meeting log
(see [`manage_teams_meetings.py`](#manage_teams_meetingspy)) is user data.

## Installation

The scripts use [pixi](https://pixi.sh) to manage dependencies
(`python`, `msal`, `requests`):

```sh
pixi install
```

### Dev container

`.devcontainer/` contains a rootless podman dev container for
[Zed](https://zed.dev/docs/dev-containers). It needs these Zed settings:

```json
"use_podman": true,
"dev_container_use_buildkit": false,
```

The container runs as your user (`--userns=keep-id`), without any capabilities
and confined by SELinux. The project folder is mounted with podman's `:z`
option, which relabels it to the SELinux type `container_file_t` so the
container can access it (undo with `restorecon -R .`). Zed can't pass mount
options, so its own project mount is redirected to an unused tmpfs
(`workspaceMount`) and the project is mounted via `runArgs` instead.

Two podman volumes keep the container's state separate from the host. Rootless
podman stores them under `~/.local/share/containers/storage/volumes/`:

- `teams-meeting-pixi`: the container's pixi environment, mounted over `.pixi`
  (the host's environment contains host paths and must not be rebuilt)
- `teams-meeting-home`: the container's home directory, i.e. its own pixi
  cache, Teams token cache and meeting log, so you sign in to Microsoft
  separately inside the container

Zed doesn't rebuild the container when the configuration changes; remove it
(and the volumes, if needed) with `podman rm -f` / `podman volume rm` and reopen
the project. The image uses the latest pixi release (`pixi:trixie-slim`), but
podman keeps using its local copy of the base image; to update pixi in the
container, run `podman pull ghcr.io/prefix-dev/pixi:trixie-slim` before
removing the container.

krun microVMs (`--runtime=krun`) can't be used: Zed connects to dev containers
with `podman exec`, which krun doesn't support.

## Usage

### `make_teams_meeting.py`

Without options, creates a one-hour meeting starting now, named
"Meeting with <your name>" (like Teams' own "Meet now"):

```sh
pixi run meeting
```

Options:

| Option | Description |
| --- | --- |
| `-s`, `--subject` | Meeting subject (default: "Meeting with <your name>") |
| `-t`, `--start` | Start in local time, e.g. `"2026-10-09 14:00"`, or just `14:00` for today if still in the future (default: now) |
| `-d`, `--duration` | Duration in minutes (default: 60) |
| `-l`, `--lobby-bypass` | Who joins directly without waiting in the lobby (see below) |
| `-p`, `--presenters` | Who can present: `everyone`, `organization` or `organizer` |
| `-c`, `--chat` | Meeting chat: `enabled`, `disabled` or `limited` (only during the meeting) |
| `-r`, `--record` | Start recording automatically |
| `--short` | Print the short join link (`https://teams.microsoft.com/meet/...`) instead of the long one |
| `--html` | Print the invitation block as HTML instead of the link (see below) |

Settings that are not given are left to your organisation's Teams meeting
policy, and the policy may also override settings that are given.

The subject is shown to everyone who joins (meeting window and chat) and names
recordings and attendance reports, so set `-s` for anything but quick ad-hoc
meetings. The name in the default is fetched from your Microsoft 365 profile;
if that fails, the default is "Teams meeting".

Lobby bypass values for `-l`:

| Value | Joins directly |
| --- | --- |
| `everyone` | Everyone, i.e. no lobby |
| `organization` | People from your organisation **and guests** |
| `organizationExcludingGuests` | People from your organisation only |
| `organizationAndFederated` | Your organisation, trusted organisations and guests |
| `invited` | Only people the organizer invited |
| `organizer` | Only the organizer; everyone else waits in the lobby |

Example:

```sh
pixi run meeting -s "Group meeting" -t "2026-10-09 14:00" -d 90 -l everyone -p organizer
pixi run meeting -s "Quick sync" -t 14:30 -d 30     # today at 14:30
```

`--html` prints the same invitation block that Teams puts into meeting
invitations: heading, short join link, meeting ID, passcode, help links and the
meeting options link. To paste it as formatted text, e.g. into an email:

```sh
pixi run meeting -s "Group meeting" -t 14:00 --html | wl-copy --type text/html
```

The short link is only part of this block (Graph has no separate field for
it), so `--short` takes it from there. Both links lead to the same meeting.

Besides the link on stdout, the script prints to stderr:

- a confirmation with the subject and time as stored by Teams, in local time,
  e.g. `Created: Group meeting, Wed 2026-10-21 14:00–15:30`
- the meeting ID and passcode, for joining by typing them into Teams
- the meeting options link, where the organizer can change the settings later
  in the browser

The start and end times don't restrict when people can join, but they set when
the link expires: 60 days after the meeting's start or end time, extended by
another 60 days whenever the meeting is used or updated
([Microsoft docs](https://learn.microsoft.com/en-us/graph/api/resources/onlinemeeting)).
For a meeting further in the future, set `--start` accordingly.

The function `create_online_meeting(subject, start, end, settings)` can also be
imported from other scripts; all arguments are optional, with the same defaults
as the command line. It returns the full
[onlineMeeting](https://learn.microsoft.com/en-us/graph/api/resources/onlinemeeting)
response as a dict; the join link is in its `"joinWebUrl"` key.

`settings` is a dict of further onlineMeeting properties, e.g.
`{"lobbyBypassSettings": {"scope": "everyone"}}`. The function raises an
exception on failure. Naive datetimes are taken as local time.

### `manage_teams_meetings.py`

Microsoft Graph can't list a user's Teams meetings; it can only look up a single
meeting by its ID, join link or meeting ID. `make_teams_meeting.py` therefore
records every meeting it creates in `~/.local/share/teams-meeting/meetings.jsonl`
(subject, times, long and short join links, meeting ID, passcode, options link;
readable only by you), and this script works with that log:

```sh
pixi run meetings list             # numbered list of recorded meetings
pixi run meetings list -v          # ... with meeting ID, passcode and (short) links
pixi run meetings list --check     # ... and ask Graph whether each still exists and when it expires
pixi run meetings delete 2 5       # delete meetings by their number from 'list'
pixi run meetings delete 'https://teams.microsoft.com/l/meetup-join/...'   # by join link
pixi run meetings prune            # drop deleted/expired meetings from the log
```

Deleting by join link also works for meetings that aren't in the log. Meetings
created before the log existed or in the Teams app aren't listed.

## License

[MIT](LICENSE)
