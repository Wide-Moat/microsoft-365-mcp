"""Short links: the signed value stays on the server and the link carries an id."""

import pytest

from m365_mcp.links.store import LinkStore, LinksFull


def test_an_id_returns_the_value_it_was_given_until_it_expires():
    store = LinkStore(clock=lambda: 1000.0)
    key = store.put("signed-value", expires=1300)
    assert store.get(key) == "signed-value"
    assert store.get(key, now=1299.9) == "signed-value"
    assert store.get(key, now=1300.0) is None


def test_an_id_is_short_unguessable_and_url_safe():
    store = LinkStore(per_owner=1000, clock=lambda: 0.0)
    keys = {store.put("v", expires=10) for _ in range(1000)}
    assert len(keys) == 1000
    for k in keys:
        assert 22 <= len(k) <= 32 and k.replace("-", "").replace("_", "").isalnum()


def test_an_unknown_or_malformed_id_is_none():
    store = LinkStore(clock=lambda: 0.0)
    assert store.get("nope") is None
    assert store.get("") is None
    assert store.get("x" * 5000) is None


def test_expired_entries_make_room_and_live_ones_are_never_dropped():
    now = [0.0]
    store = LinkStore(max_entries=3, clock=lambda: now[0])
    old = store.put("a", expires=5)
    store.put("b", expires=100)
    store.put("c", expires=100)
    now[0] = 10.0
    fresh = store.put("d", expires=100)  # "a" has expired and gives its place
    assert store.get(old) is None and store.get(fresh) == "d"
    with pytest.raises(LinksFull):
        store.put("e", expires=100)  # three live links: refused, not silently evicting one in use


def test_one_person_cannot_fill_the_store_for_everyone():
    store = LinkStore(max_entries=10, per_owner=3, clock=lambda: 0.0)
    for _ in range(3):
        store.put("v", expires=100, owner="p1")
    with pytest.raises(LinksFull):
        store.put("v", expires=100, owner="p1")
    assert store.get(store.put("v", expires=100, owner="p2")) == "v"
