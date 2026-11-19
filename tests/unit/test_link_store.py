"""Short links: the signed value stays in the shared store and the link carries an id."""

import pytest

from m365_mcp.links.store import LinksFull, LinkStore
from m365_mcp.state.memory import MemoryStore


def store_at(now: list[float], **kw) -> LinkStore:
    return LinkStore(MemoryStore(clock=lambda: now[0]), **kw)


async def test_an_id_returns_the_value_it_was_given_until_it_expires():
    now = [1000.0]
    store = store_at(now)
    key = await store.put("signed-value", expires=1300)
    assert await store.get(key) == "signed-value"
    now[0] = 1299.9
    assert await store.get(key) == "signed-value"
    now[0] = 1300.0
    assert await store.get(key) is None


async def test_an_id_is_short_unguessable_and_url_safe():
    store = store_at([0.0], per_owner=1000)
    keys = {await store.put("v", expires=10) for _ in range(1000)}
    assert len(keys) == 1000
    for k in keys:
        assert 22 <= len(k) <= 32 and k.replace("-", "").replace("_", "").isalnum()


async def test_an_unknown_or_malformed_id_is_none():
    store = store_at([0.0])
    assert await store.get("nope") is None
    assert await store.get("") is None
    assert await store.get("x" * 5000) is None


async def test_expired_entries_make_room_and_live_ones_are_never_dropped():
    now = [0.0]
    store = store_at(now, max_entries=3)
    old = await store.put("a", expires=5)
    await store.put("b", expires=100)
    await store.put("c", expires=100)
    now[0] = 10.0
    fresh = await store.put("d", expires=100)  # "a" has expired and gives its place
    assert await store.get(old) is None and await store.get(fresh) == "d"
    with pytest.raises(LinksFull):
        await store.put("e", expires=100)  # three live links: refused, not silently evicting one in use


async def test_one_person_cannot_fill_the_store_for_everyone():
    store = store_at([0.0], max_entries=10, per_owner=3)
    for _ in range(3):
        await store.put("v", expires=100, owner="p1")
    with pytest.raises(LinksFull):
        await store.put("v", expires=100, owner="p1")
    assert await store.get(await store.put("v", expires=100, owner="p2")) == "v"


async def test_the_store_never_holds_a_usable_link():
    # Someone who reads the store has neither the id (stored under a keyed hash) nor the value in clear.
    mem = MemoryStore(clock=lambda: 0.0)
    store = LinkStore(mem)
    key = await store.put("SIGNED.VALUE", expires=100, owner="person-key", tag="draft:person-key:D1")
    held = repr(mem._links) + repr(mem._buckets)
    for needle in (key, "SIGNED.VALUE", "person-key", "D1"):
        assert needle not in held, needle


async def test_a_rotation_keeps_finding_what_the_previous_key_issued():
    # Pods roll one by one after a key rotation: a link an old pod issued is opened by a new one.
    mem = MemoryStore(clock=lambda: 0.0)
    old_key, new_key = b"o" * 32, b"n" * 32
    before = LinkStore(mem, key=old_key)
    after = LinkStore(mem, key=new_key, previous=(old_key,))
    link = await before.put("SIGNED.VALUE", expires=100, owner="p", tag="draft:p:D1")
    assert await after.get(link) == "SIGNED.VALUE"
    assert await after.claim(link) and not await before.claim(link), "one upload holds it under both keys"
    assert await after.busy_under("draft:p:D1")
    await after.consume(link, until=100)
    assert await before.spent(link) and await after.spent(link)
    assert await LinkStore(mem, key=new_key).get(await before.put("V", expires=100)) is None, (
        "without the previous key, its links are not found"
    )


async def test_a_link_delivered_under_one_key_cannot_be_claimed_under_another():
    # During a rotation: an upload verified on one pod, then delivered through another, must not be
    # claimed again by falling back to the other key's marks.
    mem = MemoryStore(clock=lambda: 0.0)
    old_key, new_key = b"o" * 32, b"n" * 32
    before = LinkStore(mem, key=old_key)
    after = LinkStore(mem, key=new_key, previous=(old_key,))
    link = await before.put("V", expires=100, owner="p")
    assert await before.claim(link)
    await before.consume(link, until=100)
    assert not await after.claim(link)


async def test_memory_forgets_a_bucket_once_it_is_full_again_and_never_sooner():
    now = [0.0]
    mem = MemoryStore(clock=lambda: now[0])
    await mem.take("a", capacity=2, refill_per_second=1)
    await mem.take("b", capacity=2, refill_per_second=1)
    now[0] = 0.5
    await mem.take("c", capacity=2, refill_per_second=1)
    assert mem.bucket_count() == 3, "a refilling bucket was forgotten, which would refill it"
    now[0] = 5.0
    await mem.take("c", capacity=2, refill_per_second=1)
    assert mem.bucket_count() == 1


async def test_memory_forgets_a_bucket_by_its_own_policy_not_the_callers():
    mem = MemoryStore(clock=lambda: 0.0)
    await mem.take("big", capacity=10, refill_per_second=0.1)  # 9 left, full in 10 s
    await mem.take("small", capacity=1, refill_per_second=1)  # 9 >= 1 must not forget "big"
    taken, left = await mem.take("big", capacity=10, refill_per_second=0.1)
    assert taken and round(left) == 8, "another policy's take refilled this bucket"


async def test_an_upload_finishing_after_its_link_expired_frees_its_draft_during_a_rotation():
    # Claimed under the previous key, its link then expires while the body arrives: the release (or the
    # delivery) must still clear the claim, or send_draft waits for the whole hold.
    now = [0.0]
    mem = MemoryStore(clock=lambda: now[0])
    old_key, new_key = b"o" * 32, b"n" * 32
    before = LinkStore(mem, key=old_key)
    after = LinkStore(mem, key=new_key, previous=(old_key,))
    failed = await before.put("V", expires=10, owner="p", tag="draft:p:D1")
    delivered = await before.put("W", expires=10, owner="p", tag="draft:p:D2")
    assert await after.claim(failed) and await after.claim(delivered)
    now[0] = 20.0
    assert await after.busy_under("draft:p:D1") and await after.busy_under("draft:p:D2")
    await after.release(failed)
    await after.consume(delivered, until=30)
    assert not await after.busy_under("draft:p:D1") and not await after.busy_under("draft:p:D2")
    assert await after.spent(delivered)
