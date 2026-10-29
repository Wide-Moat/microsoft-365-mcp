"""Fake People search for tests that resolve a person by name."""

from __future__ import annotations

import httpx


def person(graph, name: str, addr: str, group: bool = False) -> None:
    def route(seen):
        q = seen.params.get("$search", "").strip('"').lower()
        if q and (q in name.lower() or q in addr.lower()):
            return httpx.Response(200, json={"value": [{
                "displayName": name, "scoredEmailAddresses": [{"address": addr}],
                "personType": {"class": "Group" if group else "Person", "subclass": "OrganizationUser"}}]})
        return httpx.Response(200, json={"value": []})
    graph.on("GET", "/me/people", route)


def me(graph, uid: str = "me-id", mail: str = "me@x.test") -> None:
    graph.on("GET", "/me", {"id": uid, "displayName": "Me", "mail": mail, "userPrincipalName": mail})
