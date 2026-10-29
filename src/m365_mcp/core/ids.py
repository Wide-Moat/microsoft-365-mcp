"""Ids from the model, checked before they become one segment of a Graph path.

The Graph client refuses a path that an id could rewrite (`graph/client.py`), but only once the
path is built, and the person then reads that "an id or path is not valid". An id is one
segment: one holding a separator was never returned by a tool, so to the person it is an id that
does not exist (rule 11), and nothing reaches Graph.
"""

from __future__ import annotations

import re

from m365_mcp.core.errors import M365Error

_EMPTY = "An id or path in this request is empty or not valid. Use the ids and paths the tools returned."
_NOT_ONE_SEGMENT = re.compile(r"[/\\?#\s]")


def segment(value: str | None) -> str:
    """The id, or a refusal: empty is "not valid", anything that is not one path segment is not found.
    `None` is not accepted here; callers decide what an absent id means before calling."""
    v = (value or "").strip()
    if not v:
        raise M365Error("invalid", reason=_EMPTY)
    if _NOT_ONE_SEGMENT.search(v) or v in (".", ".."):
        raise M365Error("not_found")
    return v
