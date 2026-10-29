"""The one gate in front of everything other people see."""

from __future__ import annotations

import math
import time
from collections.abc import Callable
from dataclasses import dataclass

from m365_mcp.core.errors import M365Error

MAX_RECIPIENTS = 20
CAPACITY = 10
REFILL_PER_SECOND = 10 / 60


# What a refusal by the rate says happens next, by what was asked (a refused send left its draft,
# and an agent repeating save_draft -> send_draft made a new draft each time).
_RATE_THEN = {"mail": "Nothing was sent and the draft is kept: call send_draft again with the same id later, and do "
                      "not save a new draft.",
              "share": "Nothing was shared.", "event": "Nothing was sent; nothing was changed."}

# What the confirmation refusal says did not happen, by what was asked; the default is a message.
_UNDONE = {"assignment": "Nothing was saved.", "chat_members": "Nothing was changed.", "share": "Nothing was shared.",
           "automatic_reply": "Nothing was changed.", "booking": "Nothing was booked or changed.",
           "event_response": "Nothing was answered."}


@dataclass
class _Bucket:
    tokens: float
    at: float


class Guard:
    def __init__(self, clock: Callable[[], float] = time.monotonic) -> None:
        self._clock = clock
        self._buckets: dict[str, _Bucket] = {}

    def _take(self, person: str, kind: str = "") -> int:
        now = self._clock()
        b = self._buckets.get(person)
        if b is None:
            b = self._buckets[person] = _Bucket(CAPACITY, now)
        b.tokens = min(CAPACITY, b.tokens + (now - b.at) * REFILL_PER_SECOND)
        b.at = now
        if b.tokens < 1:
            wait = math.ceil((1 - b.tokens) / REFILL_PER_SECOND)
            raise M365Error("rate", seconds=wait, then=_RATE_THEN.get(kind, "Nothing was sent."))
        b.tokens -= 1
        return int(b.tokens)

    def check(self, person: str, *, kind: str, user_confirmed: bool | None, recipients: int = 0,
              group_addresses: list[str] | None = None, needs_confirmation: bool = True,
              addresses: list[str] | None = None, allowed_domains: frozenset[str] = frozenset()) -> dict:
        """Call before the Graph request. Returns the receipt the tool includes in its result."""
        if needs_confirmation and user_confirmed is not True:
            raise M365Error("confirm", outcome=_UNDONE.get(kind, "Nothing was sent."))
        if recipients > MAX_RECIPIENTS:
            raise M365Error("too_many_recipients", limit=MAX_RECIPIENTS, count=recipients)
        if group_addresses:
            raise M365Error("distribution_list", address=group_addresses[0])
        outside = outside_domains(addresses or [], allowed_domains)
        if outside:
            raise M365Error("outside_domains", domains=", ".join(outside))
        left = self._take(person, kind)
        return {"kind": kind, "recipients": recipients, "sends_left_now": left}


def shown(receipt: dict) -> dict:
    """The part of a receipt the person's conversation needs: what is left of the allowance, always, so
    an agent can pace a run of sends rather than meet the refusal (shown at 0-2 only, it came
    too late to plan by)."""
    left = receipt.get("sends_left_now")
    return {"sends_left": left} if left is not None else {}


def outside_domains(addresses: list[str], allowed: frozenset[str]) -> list[str]:
    """The domains among `addresses` that are neither an allowed domain nor a subdomain of one.
    Domains only: the refusal is read by the model, and the addresses are the person's data."""
    if not allowed:
        return []
    bad: set[str] = set()
    for a in addresses:
        domain = a.rpartition("@")[2].strip().rstrip(".").lower() if "@" in a else ""
        if not any(domain == d or domain.endswith("." + d) for d in allowed):
            bad.add(domain or "(no domain)")
    return sorted(bad)


GUARD = Guard()
