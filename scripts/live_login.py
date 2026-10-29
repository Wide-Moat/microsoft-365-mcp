#!/usr/bin/env python3
"""Sign two test users in with the OAuth 2.0 device code flow and save their Graph tokens.

For the live suite (`pytest -m live`), never for production. The app registration must allow
public client flows ("Allow public client flows" = Yes under Authentication); no redirect URI
and no client secret are needed. Its client id goes in M365_LIVE_CLIENT_ID.

    M365_LIVE_CLIENT_ID=<app id> python scripts/live_login.py            # sign in user a, then user b
    python scripts/live_login.py --user b                                # sign in only user b, keep a
    python scripts/live_login.py --refresh                               # renew the users in the file
    python scripts/live_login.py --tenant <tenant id or domain>          # only that tenant's accounts sign in

The scopes are the delegated ones in docs/requirements.md "Permissions" plus offline_access. The
five that need administrator consent are included unless --no-admin-scopes is given; --scopes
replaces the whole list (short names such as Mail.ReadWrite get the Graph prefix).

The file written by --out holds {"a": {...}, "b": {...}} with access_token, refresh_token and
expires_at, and is created with mode 0600. A token is never printed.
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import re
import tempfile
import sys
import time
from pathlib import Path

import httpx

# --tenant (or M365_LIVE_TENANT) replaces "organizations". Without it the browser's current account
# answers, and one signed in to another organisation gets AADSTS5000228 for a tenant nobody meant.
AUTHORITY = "https://login.microsoftonline.com/organizations/oauth2/v2.0"
TENANT_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9.-]{0,252}$")
GRAPH = "https://graph.microsoft.com/"

# docs/requirements.md "Permissions". People.Read is added because list_people ranks through /me/people.
USER_SCOPES = [
    "User.Read", "User.ReadBasic.All", "People.Read", "MailboxSettings.ReadWrite",
    "Mail.ReadWrite", "Mail.Send",
    "Calendars.ReadWrite", "Calendars.Read.Shared", "Presence.Read.All", "OnlineMeetings.Read",
    "Files.ReadWrite.All", "Sites.ReadWrite.All", "Files.ReadWrite",
    "Chat.ReadWrite", "ChatMessage.Send", "ChannelMessage.Send", "Team.ReadBasic.All", "Channel.ReadBasic.All",
    "Contacts.ReadWrite", "Tasks.ReadWrite", "Notes.ReadWrite",
]
ADMIN_SCOPES = [
    "ChannelMessage.Read.All", "ChannelMessage.ReadWrite",
    "OnlineMeetingRecording.Read.All", "OnlineMeetingTranscript.Read.All", "Place.Read.All",
]


def scope_string(names: list[str]) -> str:
    full = [n if n.startswith("https://") else GRAPH + n for n in names]
    return " ".join(full + ["offline_access"])


def fail(message: str) -> None:
    print(f"live_login: {message}", file=sys.stderr)
    raise SystemExit(1)


def oauth_error(resp: httpx.Response) -> str:
    """The error code and description of a failed token request; never the whole body."""
    try:
        body = resp.json()
        return f"{body.get('error', resp.status_code)}: {str(body.get('error_description', '')).splitlines()[0]}"
    except (ValueError, IndexError, AttributeError):
        return f"HTTP {resp.status_code}"


def claims(access_token: str) -> dict:
    try:
        payload = access_token.split(".")[1]
        return json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
    except (IndexError, ValueError):
        return {}


def stored(tokens: dict, now: float | None = None) -> dict:
    now = time.time() if now is None else now
    return {
        "access_token": tokens["access_token"],
        "refresh_token": tokens.get("refresh_token", ""),
        "expires_at": int(now + int(tokens.get("expires_in", 3600))),
        "scope": tokens.get("scope", ""),
    }


def device_login(http: httpx.Client, client_id: str, scope: str, label: str) -> dict:
    resp = http.post(f"{AUTHORITY}/devicecode", data={"client_id": client_id, "scope": scope})
    if resp.status_code != 200:
        fail(f"device code request refused ({oauth_error(resp)}). Is public client flow enabled on the app?")
    dc = resp.json()
    print(f"\nUser {label}: open {dc['verification_uri']} and enter the code {dc['user_code']}")
    print(f"Sign in as the {label.upper()} test account. The code expires in {int(dc['expires_in']) // 60} minutes.")
    interval = int(dc.get("interval", 5))
    deadline = time.time() + int(dc["expires_in"])
    while time.time() < deadline:
        time.sleep(interval)
        r = http.post(f"{AUTHORITY}/token", data={
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id, "device_code": dc["device_code"]})
        if r.status_code == 200:
            return stored(r.json())
        error = (r.json() if r.content else {}).get("error")
        if error == "authorization_pending":
            continue
        if error == "slow_down":
            interval += 5
            continue
        fail(f"sign-in of user {label} failed ({oauth_error(r)})")
    fail(f"the code for user {label} expired before sign-in")
    raise AssertionError("unreachable")


def refresh(http: httpx.Client, client_id: str, entry: dict, label: str) -> dict:
    if not entry.get("refresh_token"):
        fail(f"user {label} has no refresh token; sign in again without --refresh")
    # No scope: the token endpoint then returns the scopes consented at sign-in.
    r = http.post(f"{AUTHORITY}/token", data={
        "grant_type": "refresh_token", "client_id": client_id, "refresh_token": entry["refresh_token"]})
    if r.status_code != 200:
        fail(f"refresh of user {label} failed ({oauth_error(r)}); sign in again")
    new = stored(r.json())
    new["refresh_token"] = new["refresh_token"] or entry["refresh_token"]
    return new


def write(path: Path, data: dict) -> None:
    # A fresh file of our own (O_EXCL through mkstemp, mode 0600 from the start), never a
    # predictable name that could already exist with wider permissions or as a symlink.
    fd, tmp = tempfile.mkstemp(prefix=".live-tokens-", dir=path.parent or ".")
    try:
        with os.fdopen(fd, "w") as f:
            json.dump(data, f, indent=2)
        os.replace(tmp, path)
    except BaseException:
        os.unlink(tmp)
        raise


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--out", default=".live-tokens.json", help="token file (default .live-tokens.json)")
    p.add_argument("--refresh", action="store_true", help="renew the users in the file from their refresh tokens")
    p.add_argument("--user", choices=("a", "b"), help="sign in only this user and keep the other one in the file")
    p.add_argument("--no-admin-scopes", action="store_true", help="leave out the scopes that need administrator consent")
    p.add_argument("--scopes", help="space or comma separated list replacing the default scopes")
    p.add_argument("--client-id", default=os.environ.get("M365_LIVE_CLIENT_ID", ""), help="default: $M365_LIVE_CLIENT_ID")
    p.add_argument("--tenant", default=os.environ.get("M365_LIVE_TENANT", "organizations"),
                   help="tenant id or domain the test users belong to; default $M365_LIVE_TENANT or organizations")
    args = p.parse_args(argv)
    if not TENANT_RE.match(args.tenant) or ".." in args.tenant:
        fail(f"--tenant must be a tenant id or domain, got {args.tenant!r}")
    global AUTHORITY
    AUTHORITY = f"https://login.microsoftonline.com/{args.tenant}/oauth2/v2.0"
    if not args.client_id:
        fail("set M365_LIVE_CLIENT_ID (or --client-id) to the application id of the test app registration")
    out = Path(args.out)
    with httpx.Client(timeout=30) as http:
        if args.refresh:
            if not out.exists():
                fail(f"{out} does not exist; sign in first")
            data = json.loads(out.read_text())
            if not data:
                fail(f"{out} has no users")
            for label in sorted(data):
                data[label] = refresh(http, args.client_id, data[label], label)
                print(f"user {label}: refreshed, valid for {(data[label]['expires_at'] - int(time.time())) // 60} minutes")
            write(out, data)
            return 0
        if args.scopes:
            names = args.scopes.replace(",", " ").split()
        else:
            names = USER_SCOPES + ([] if args.no_admin_scopes else ADMIN_SCOPES)
        scope = scope_string(names)
        # --user keeps the other user's tokens, so the two people can sign in at different times.
        data = json.loads(out.read_text()) if args.user and out.exists() else {}
        for label in (args.user,) if args.user else ("a", "b"):
            data[label] = device_login(http, args.client_id, scope, label)
            who = claims(data[label]["access_token"])
            print(f"user {label}: signed in as {who.get('upn') or who.get('unique_name') or who.get('preferred_username') or 'unknown'}")
        if {"a", "b"} <= data.keys():
            oa, ob = (claims(data[k]["access_token"]).get("oid") for k in ("a", "b"))
            if oa and oa == ob:
                fail("both sign-ins are the same account; the suite needs two different users")
        write(out, data)
        missing = {"a", "b"} - data.keys()
        print(f"\nwrote {out} (mode 0600). " + (f"Sign in user {missing.pop()} with --user, then run pytest -m live" if missing else "Run: pytest -m live"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
