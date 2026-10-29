# Live testing

The unit suite runs against a fake Graph. The live suite (`tests/live`, marker `live`) runs the same
tools against a real Microsoft 365 test tenant with two real people, and measures what Microsoft's
documentation leaves unknown. It never runs in CI and never runs by default: `pytest -q` deselects
it, and `pytest -m live` without credentials reports every test as skipped, not failed.

## What you need

- A Microsoft 365 test tenant with **two users** (A and B), both with Exchange, OneDrive and Teams
  licences, each having opened Outlook, OneDrive and Teams once so the mailbox, drive and chat
  service exist. Planner and SharePoint-list probes skip when the tenant has no plan or list.
- A tenant administrator for the one-time consent below.

## Register the app (once)

1. Entra admin centre, App registrations, New registration. Name it for the test (`m365-mcp-live`),
   single tenant, **no redirect URI**.
2. Authentication, Advanced settings: **Allow public client flows = Yes**. This is what permits
   the device code flow. No client secret is created or used.
3. API permissions, Add, Microsoft Graph, **Delegated**: every scope in
   [requirements.md](requirements.md) "Permissions", plus `People.Read` (list_people ranks through
   `/me/people`). Then "Grant admin consent for <tenant>": five of the scopes
   (`ChannelMessage.Read.All`, `ChannelMessage.ReadWrite`, `OnlineMeetingRecording.Read.All`,
   `OnlineMeetingTranscript.Read.All`, `Place.Read.All`) cannot be consented by a user alone.
4. Copy the Application (client) ID.

## Sign the two users in

```sh
export M365_LIVE_CLIENT_ID=<application id>
export M365_LIVE_TENANT=<test tenant id or domain>   # only that tenant's accounts can sign in
python scripts/live_login.py            # prints a code and a URL for user A, then for user B
```

Open the URL in a private window and sign in as the account asked for: first A, then B. The script
writes `.live-tokens.json` (mode 0600, git-ignored) with `access_token`, `refresh_token` and
`expires_at` per user, and never prints a token. When the two people sign in at different times,
`--user a` or `--user b` signs in only that one and keeps the other already in the file. Access
tokens last about an hour; renew the users in the file with

```sh
python scripts/live_login.py --refresh
```

`--no-admin-scopes` leaves out the five admin scopes (their tests then fail on a missing permission,
which is itself the correct result); `--scopes "Chat.ReadWrite ChannelMessage.Send"` replaces the
whole list, which is how the ChatMessage.Send / ChannelMessage.Send finding is repeated with a token
lacking one of them (write it to another file with `--out` and point `M365_LIVE_TOKENS` at it).

## Run

```sh
pytest -m live -q                       # everything
pytest -m live -q tests/live/test_findings.py     # only the probes
```

The suite starts only when both tokens exist and have at least ten minutes left; otherwise every test
is skipped with the reason and the path it looked at.

| Variable | Meaning |
|---|---|
| `M365_LIVE_TOKENS` | token file, default `.live-tokens.json` in the repository root |
| `M365_LIVE_FINDINGS` | findings report, default `live-findings.json` in the repository root |
| `M365_LIVE_RECORDED_EVENT` | an event id in A's calendar with a Teams recording; enables the recording test |

What a run does: A and B exchange a 10 MB letter, a reply, a meeting, files, a workbook, a Teams
message, tasks, a OneNote page and a contact, and A's automatic reply is switched on and off. Real
mail and invitations are sent between the two test users only. Everything carries a run id
`m365-live-<8 hex>`; at the end the session removes it (mail to Deleted Items, events, the files
folder to the recycle bin, tasks, pages, contacts, test chat messages soft-deleted). A cleanup step
that fails is printed and raised as a warning, never as a test failure. Two things stay behind by
design: the chat between A and B, and a notebook if the tenant had none (Graph cannot delete
notebooks; the warning names it).

The run is bounded by the tenant, not by the suite: letters are polled for up to 60 s, and the
people-search probe waits up to two minutes for a new contact to be indexed.

## Turn the measurements into the findings document

```sh
python scripts/live_findings_to_md.py   # live-findings.json -> docs/graph-live-findings.md
```

The table lists every unknown with what was observed, the date, and what the code does about it.
Rows the run did not re-measure keep their earlier values. When an observation contradicts the
code, change the code and its test in the same commit, and add a note to
[graph-reference.md](graph-reference.md) pointing at the finding.

## Adding a probe

Write the test in `tests/live/test_findings.py`, record with `live.finding("name", observed)`, and add
the name with its question and the code's current behaviour to `CATALOGUE` in
`scripts/live_findings_to_md.py`. A probe fails only when it cannot measure, never because Graph
answered something unexpected: the answer is the result.
