"""The person's word for what cannot be taken back, when no other person is told.

Messages that other people see go through the guard. Deleting a task, a page or a contact
notifies nobody, so it needs only the word itself, not a draw from the bucket.
"""

from __future__ import annotations

from m365_mcp.core.errors import invalid


def require(user_confirmed: bool | None, what: str) -> None:
    if user_confirmed is not True:
        raise invalid(f"{what} cannot be undone from here. Ask the person to confirm, then call again with "
                      "user_confirmed=true. Nothing was changed.")

