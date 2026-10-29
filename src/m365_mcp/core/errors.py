"""The one error type a tool raises, and the sentences a person reads.

`str()` of an error is its code alone. Tracebacks and log lines print `str()`, and the
parameters — candidate names, addresses, counts — are the person's data, so they reach the
model through the result and never the log.
"""

from __future__ import annotations

from typing import Any

from mcp.server.mcpserver.exceptions import ToolError

# key -> sentence. Every sentence names the next step for the person.
MESSAGES: dict[str, str] = {
    "reconnect": "Microsoft 365 is not connected or the sign-in expired. Reconnect Microsoft 365 in the portal.",
    "missing_permission": (
        "This needs the Microsoft permission {permission}, which was not granted. "
        "A Microsoft 365 administrator must approve the app for the organisation."
    ),
    "area_off": "The {area} tools are switched off for this connection.",
    "forbidden": "Microsoft 365 does not let you do this here: you lack access to this item, or the organisation's policy blocks it.",
    "not_found": "It does not exist, or it is not shared with you. Check the id, or find it again with the tool that lists it.",
    "throttled": "Microsoft is limiting requests right now. Try again in {seconds}.",
    "failed": ("Microsoft 365 refused the request, and nothing was changed. Check the values given, or do it in "
               "the Microsoft 365 app. Reference: {request_id}."),
    "uncertain": ("Microsoft 365 failed while handling this, so it may or may not have happened. Check before "
                  "trying again. Reference: {request_id}."),
    "failed_read": ("Microsoft 365 failed while reading this; nothing was changed. Try again in a moment. "
                    "Reference: {request_id}."),
    "no_onedrive": ("Your OneDrive is not set up yet, or your licence does not include it. Open OneDrive once in "
                    "the browser, or ask your Microsoft 365 administrator to assign a licence; reconnecting will "
                    "not help. Nothing was changed."),
    "service_unavailable": ("Microsoft 365 turned this service down for your account although the sign-in is valid: "
                            "your OneDrive or OneNote may not be set up yet, or your licence does not include it. Open "
                            "it once in the browser, or ask your administrator; reconnecting will not help. Nothing "
                            "was changed."),
    "unreachable": ("Microsoft Teams cannot reach that person: they are outside the organisation or not in its "
                    "directory. Write to them by mail instead. Nothing was sent."),
    "conflict": "It changed since it was read. Read it again, then retry.",
    "too_large": "It is too large: {size} against a limit of {limit}.",
    "exists": "An item with that name already exists: {name}.",
    "ambiguous": "More than one {kind} matches \"{query}\". Ask which one is meant: {candidates}.",
    "unresolved": "No {kind} matches \"{query}\". {where}",
    "confirm": ("This is seen by other people at once. Ask the person to confirm{names}, then call again with "
                "user_confirmed=true. {outcome}"),
    "confirm_outside": ("This is seen by other people at once, and some recipients are outside the organisation "
                        "({domains}). Tell the person, ask them to confirm, then call again with user_confirmed=true. "
                        "Nothing was sent."),
    "too_many_recipients": "At most {limit} recipients per action; this has {count}.",
    "distribution_list": ("{address} is not a person in the organisation's directory: it is a group or distribution "
                          "list, an alias, or an address that does not exist. Name the people, or use a person's "
                          "main address; bulk sending is not supported. Nothing was sent."),
    "outside_domains": ("This server only reaches people at the organisation's allowed domains, and these are "
                        "outside them: {domains}. Remove those recipients. Nothing was sent."),
    "rate": ("Too many sends in a short time. The allowance is the person's: 10 at once, then one every 6 seconds, "
             "shared by mail, Teams, invitations and sharing in every chat they have open. Wait {seconds}. If it "
             "is refused again, another chat is sending too, so wait longer. {then}"),
    "invalid": "{reason}",
    "not_owner": "Only the organiser can change this meeting. Respond to it instead.",
    "not_own_draft": "Only drafts created by the assistant can be changed or sent here. This draft was written by the person.",
    "not_own_message": "Only your own messages can be changed.",
    "cancelled_event": "This meeting was cancelled, so there is nothing to respond to.",
    "system_folder": "{name} is a system folder and is not changed.",
    "links_unavailable": "File transfer is not configured on this server, so files cannot be moved.",
    "links_full": "Too many file links are open on this server right now. Try again in a few minutes.",
    "unsupported": "{reason}",
    "unexpected": "Something went wrong on our side; check whether the change happened before trying again. Reference: {request_id}.",
}


# Where to look when a name matched nothing, by the kind of thing named.
WHERE = {
    "person": "Check the spelling, or give the address; list_people finds people by part of a name.",
    "folder": "list_messages kind=folders lists the folders.",
    "team": "list_conversations kind=channel lists the teams and their channels.",
    "channel": "list_conversations kind=channel lists the teams and their channels.",
}


def _sentence_params(key: str, params: dict[str, Any]) -> dict[str, Any]:
    out = dict(params)
    if "seconds" in out:
        out["seconds"] = f"{out['seconds']} second" + ("" if out["seconds"] == 1 else "s")
    if key == "confirm":
        out.setdefault("outcome", "Nothing was sent.")
        out["names"] = f" (it goes to: {out['names']})" if out.get("names") else ""
    if key == "rate":
        out.setdefault("then", "Nothing was sent.")
    if key == "unresolved":
        out["where"] = WHERE.get(str(out.get("kind")), "Check the spelling, or list them to see the names.")
    return out


class M365Error(ToolError):
    """A refusal or failure the person should read. `isError` is set by the SDK."""

    def __init__(self, key: str, **params: Any) -> None:
        if key not in MESSAGES:
            raise KeyError(f"unknown error key {key!r}")
        self.key = key
        self.params = params
        super().__init__(key)
        # `args[0]` is what the SDK sends to the model. It is built here, after
        # `super().__init__`, so `str(self)` stays the bare key.
        self._text = MESSAGES[key].format(**_sentence_params(key, params))

    def __str__(self) -> str:
        return self.key

    @property
    def text(self) -> str:
        return self._text


def invalid(reason: str) -> M365Error:
    return M365Error("invalid", reason=reason)
