"""Graph failures, carried without their body: Graph error texts can hold addresses and subjects."""

from __future__ import annotations

import re

from m365_mcp.core.errors import M365Error


class GraphError(Exception):
    def __init__(self, status: int, code: str, request_id: str, retry_after: int | None = None,
                 missing_scope: str | None = None, inner_code: str = "", method: str = "") -> None:
        self.status = status
        self.code = code
        self.request_id = request_id
        self.retry_after = retry_after
        self.missing_scope = missing_scope
        self.inner_code = inner_code
        self.method = method
        super().__init__(f"graph {status} {code}")


_SCOPE_IN_TEXT = re.compile(r"\b([A-Z][A-Za-z]+(?:\.[A-Z][A-Za-z]+)+(?:\.All)?)\b")


def scope_from_message(message: str) -> str | None:
    """Graph sometimes names the missing scope in its message; only a scope-shaped token is kept."""
    m = _SCOPE_IN_TEXT.search(message or "")
    return m.group(1) if m else None


# Ids Graph could not parse. Each service says it its own way; all of them mean the same to the person
# as an id that parses and is not there (rule 11): it does not exist, or it is not shared with you.
INVALID_ID_CODES = frozenset({
    "ErrorInvalidIdMalformed", "ErrorInvalidId", "ErrorInvalidIdEmpty", "ErrorInvalidIdNotAnItemAttachmentId",
    "20112",  # OneNote: invalid entity id
})
# A 401 that is the token's own fault. Others come from one service with a valid token: OneNote
# answers 401 "40001" for a person who has no OneNote; Teams answers 401 TenantNotFound for a chat
# member who is no user here. Telling those people to reconnect sends them round in a circle.
TOKEN_CODES = frozenset({"", "invalidauthenticationtoken", "tokenexpired", "unauthenticated"})  # compared casefolded


def to_person_error(e: GraphError, *, read_only: bool = False) -> M365Error:
    code = e.code or ""
    inner = e.inner_code or ""
    if code == "invalidRequestPath":
        return M365Error("invalid", reason="An id or path in this request is empty or not valid. Use the ids and paths the tools returned.")
    if e.status == 401 and inner == "TenantNotFound":
        return M365Error("unreachable")
    if e.status == 401 and code.casefold() not in TOKEN_CODES:
        return M365Error("service_unavailable")
    if e.status == 401 or code == "InvalidAuthenticationToken":
        return M365Error("reconnect")
    # Graph has sent it as the code itself, and (live, 2026-10-23) as innerError under "Forbidden".
    if "GraphAccessToTranscriptsDisabled" in (code, inner):
        return M365Error("unsupported", reason="The organisation has turned off access to meeting transcripts.")
    if e.status == 403 and inner == "provisioningNotAllowed":
        return M365Error("no_onedrive")
    if e.status == 403:
        if e.missing_scope:
            return M365Error("missing_permission", permission=e.missing_scope)
        return M365Error("forbidden")
    if e.status == 404 or code in {"ErrorItemNotFound", "itemNotFound", "ResourceNotFound"}:
        return M365Error("not_found")
    if e.status == 400 and (code in INVALID_ID_CODES or inner in INVALID_ID_CODES
                            or (inner == "ParsingFailed" and e.method == "GET")):
        return M365Error("not_found")
    if e.status == 409 and code in {"nameAlreadyExists", "ErrorFolderExists"}:
        return M365Error("invalid", reason="An item with that name already exists there.")
    if e.status in (409, 412):
        return M365Error("conflict")
    if e.status == 413 or code in {"ErrorMessageSizeExceeded", "RequestBodyTooLarge"}:
        return M365Error("invalid", reason="It is larger than Microsoft 365 accepts for this item.")
    if e.status in (429, 503):
        return M365Error("throttled", seconds=e.retry_after or 30)
    if e.status >= 500:
        if read_only:
            return M365Error("failed_read", request_id=e.request_id or "none")
        # A write may have completed before the server failed: never promise it did not.
        return M365Error("uncertain", request_id=e.request_id or "none")
    return M365Error("failed", request_id=e.request_id or "none")
