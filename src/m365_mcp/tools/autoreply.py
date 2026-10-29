"""Automatic replies: the out-of-office text Outlook sends to everyone who writes to the person."""

from __future__ import annotations

import re
from datetime import UTC, datetime
from typing import Annotated, Any, Literal

from pydantic import Field

from m365_mcp.core.context import Call, call
from m365_mcp.core.errors import invalid
from m365_mcp.core.guard import shown
from m365_mcp.core.html import prepare_outgoing, text_to_html
from m365_mcp.core.time import graph_dt, parse
from m365_mcp.registry import tool

_STATUS = {"off": "disabled", "on": "alwaysEnabled", "scheduled": "scheduled"}
_AUDIENCE = {"none": "none", "contacts": "contactsOnly", "all": "all"}


# Only real HTML tag names switch to HTML, so "<alex@example.com>" or "<name>" stay plain text.
_TAG = re.compile(r"</?(?:p|br|b|i|u|em|strong|div|span|a|ul|ol|li|table|tr|td|th|h[1-6]|hr|img|font|"
                  r"blockquote|pre|code|small|center|script|style|iframe|object|embed|form|input|svg)(?=[\s/>])[^>]*>", re.I)


def _html(text: str, report: dict[str, list[str]]) -> str:
    """Plain text becomes paragraphs; text holding HTML goes through the outgoing pipeline like a letter
    (rule 29), so its markup is cleaned rather than escaped into visible tags."""
    prepared = prepare_outgoing(text if _TAG.search(text) else text_to_html(text))
    if prepared.errors:
        raise invalid(" ".join(prepared.errors))
    if prepared.inline_images:
        raise invalid("Images embedded in an automatic reply are not supported; link to them instead.")
    report["sanitized"] += prepared.sanitized
    report["warnings"] += prepared.warnings
    return prepared.html


async def _has_message(c: Call) -> bool:
    current = await c.graph.get("/me/mailboxSettings/automaticRepliesSetting")
    return bool((current.get("internalReplyMessage") or "").strip() or (current.get("externalReplyMessage") or "").strip())


@tool(area="mail", title="Set the automatic reply", scopes=["MailboxSettings.ReadWrite"], notifies_others=True)
async def save_automatic_reply(
    status: Literal["off", "on", "scheduled"],
    internal_message: Annotated[str | None, Field(description="Text or simple HTML for colleagues in the organisation.")] = None,
    external_message: Annotated[str | None, Field(description="Text or simple HTML for people outside it.")] = None,
    audience: Annotated[Literal["none", "contacts", "all"] | None, Field(
        description="Who outside the organisation gets it. Default contacts; none where this installation "
                    "limits recipient domains, which refuses contacts and all.")] = None,
    start: Annotated[str | None, Field(description="Scheduled only: 2026-10-23 08:00.")] = None,
    end: Annotated[str | None, Field(description="Scheduled only: when it stops.")] = None,
    user_confirmed: Annotated[bool, Field(description="True only when the person asked to turn it on.")] = False,
) -> dict[str, Any]:
    """Turn the out-of-office reply off, on, or on for a period, and set its text. Everyone who writes
    to the person while it is on receives that text, so on and scheduled need user_confirmed; off
    does not. get_connection shows whether one is set."""
    if status != "scheduled" and (start or end):
        raise invalid("start and end belong to status scheduled; use it for a reply that runs for a period.")
    if status == "scheduled" and not (start and end):
        raise invalid("A scheduled reply needs both start and end.")
    async with call("save_automatic_reply") as c:
        c.need("MailboxSettings.ReadWrite")
        setting: dict[str, Any] = {"status": _STATUS[status]}
        receipt: dict[str, Any] = {}
        report: dict[str, list[str]] = {"sanitized": [], "warnings": []}
        allowed = c.rt.settings.allowed_recipient_domains
        if audience is None:
            # Unasked, the reply stays inside where the installation limits domains; asked for outside,
            # it is refused below.
            audience = "none" if allowed else "contacts"
        if status != "off":
            if status == "scheduled":
                s, e = parse(start or "", c.tz, field="start"), parse(end or "", c.tz, field="end")
                if e <= s:
                    raise invalid("end must be after start.")
                if e <= datetime.now(UTC):
                    raise invalid("end is in the past, so the reply would never run.")
                setting["scheduledStartDateTime"], setting["scheduledEndDateTime"] = graph_dt(s), graph_dt(e)
            if internal_message is not None:
                setting["internalReplyMessage"] = _html(internal_message, report)
            if external_message is not None:
                setting["externalReplyMessage"] = _html(external_message, report)
            setting["externalAudience"] = _AUDIENCE[audience]
            if audience != "none" and allowed:
                # Whoever writes from outside gets the reply, from any domain: the allowed list cannot
                # be checked against people not known in advance, so an outside audience is refused.
                raise invalid("This installation limits whom the assistant reaches to "
                              f"{', '.join(sorted(allowed))}, and an automatic reply to people outside the "
                              "organisation goes to any domain. Use audience=none, or set it in Outlook. "
                              "Nothing was changed.")
            if internal_message is None and external_message is None and not await _has_message(c):
                raise invalid("There is no reply text yet. Give internal_message, and external_message if people outside should get one.")
            receipt = c.guard(kind="automatic_reply", user_confirmed=user_confirmed)
        await c.graph.patch("/me/mailboxSettings", {"automaticRepliesSetting": setting})
        out: dict[str, Any] = {"automatic_reply": status}
        if status != "off":
            out["audience"] = audience
            if status == "scheduled":
                out["from"], out["until"] = start, end
            if audience != "none" and external_message is None:
                out["note"] = "No external_message was given; people outside get whatever external text was already set."
            out.update(shown(receipt))
            html = {k: list(dict.fromkeys(v)) for k, v in report.items() if v}
            if html:
                out["html"] = html
        return out
