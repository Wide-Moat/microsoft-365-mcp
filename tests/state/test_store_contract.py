"""One contract for every StateStore: what replicas rely on, run against memory and against real Redis,
Valkey and Valkey behind Sentinel (see conftest.py). A behaviour that differs between them fails here."""

from __future__ import annotations

import asyncio
import time

import pytest

from m365_mcp.state.store import Namer, StoreUnavailable

T = {"owner_cap": 300, "total_cap": 5000, "tag": ""}


async def test_ping_answers(store):
    await store.ping()


async def test_a_bucket_bursts_to_capacity_then_refuses(store):
    got = [await store.take("person-a", capacity=3, refill_per_second=0.5) for _ in range(4)]
    assert [t for t, _ in got] == [True, True, True, False]
    assert [round(left) for _, left in got[:3]] == [2, 1, 0]
    assert got[3][1] < 1, "a refusal reports the tokens there are, below one"
    assert (await store.take("person-b", capacity=3, refill_per_second=0.5))[0], "buckets are per person"


async def test_a_bucket_refills_with_time(store):
    for _ in range(2):
        await store.take("p", capacity=2, refill_per_second=4)
    assert not (await store.take("p", capacity=2, refill_per_second=4))[0]
    await asyncio.sleep(0.4)  # 1.6 tokens back
    assert (await store.take("p", capacity=2, refill_per_second=4))[0]


async def test_two_replicas_draw_from_one_bucket(replicas):
    a, b = replicas
    for _ in range(3):
        assert (await a.take("shared", capacity=3, refill_per_second=0.01))[0]
    taken, _ = await b.take("shared", capacity=3, refill_per_second=0.01)
    assert not taken, "the second replica granted an allowance of its own"


async def test_concurrent_takes_never_exceed_the_capacity(replicas):
    a, b = replicas
    results = await asyncio.gather(*(s.take("race", capacity=5, refill_per_second=0.001) for s in (a, b) * 10))
    assert sum(1 for t, _ in results if t) == 5


async def test_a_link_stored_by_one_replica_is_read_by_another(replicas):
    a, b = replicas
    assert await a.link_put("id-1", b"sealed", expires=time.time() + 60, owner="o", **T)
    assert await b.link_get("id-1") == b"sealed"
    assert await b.link_get("id-unknown") is None


async def test_a_link_expires(store):
    assert await store.link_put("short", b"v", expires=time.time() + 1.0, owner="o", **T)
    assert await store.link_get("short") == b"v"
    await asyncio.sleep(1.2)
    assert await store.link_get("short") is None


async def test_an_expiry_already_past_stores_nothing(store):
    assert not await store.link_put("late", b"v", expires=time.time() - 1, owner="o", **T)
    assert await store.link_get("late") is None


async def test_one_owner_is_capped_and_a_live_link_is_never_evicted(replicas):
    a, b = replicas
    exp = time.time() + 60
    for i in range(2):
        assert await a.link_put(f"c{i}", b"v", expires=exp, owner="busy-person", owner_cap=2, total_cap=100, tag="")
    assert not await b.link_put("c2", b"v", expires=exp, owner="busy-person", owner_cap=2, total_cap=100, tag="")
    assert await b.link_get("c0") == b"v" and await b.link_get("c2") is None
    assert await b.link_put("c3", b"v", expires=exp, owner="someone-else", owner_cap=2, total_cap=100, tag="")


async def test_an_expired_link_gives_its_place_back(store):
    assert await store.link_put("e0", b"v", expires=time.time() + 0.6, owner="o", owner_cap=1, total_cap=100, tag="")
    assert not await store.link_put("e1", b"v", expires=time.time() + 60, owner="o", owner_cap=1, total_cap=100, tag="")
    await asyncio.sleep(0.8)
    assert await store.link_put("e1", b"v", expires=time.time() + 60, owner="o", owner_cap=1, total_cap=100, tag="")


async def test_the_store_as_a_whole_is_capped(store):
    exp = time.time() + 60
    assert await store.link_put("t0", b"v", expires=exp, owner="o1", owner_cap=10, total_cap=2, tag="")
    assert await store.link_put("t1", b"v", expires=exp, owner="o2", owner_cap=10, total_cap=2, tag="")
    assert not await store.link_put("t2", b"v", expires=exp, owner="o3", owner_cap=10, total_cap=2, tag="")


async def test_one_upload_holds_a_link_on_every_replica(replicas):
    a, b = replicas
    await a.link_put("up", b"v", expires=time.time() + 60, owner="o", **T)
    assert await a.link_claim("up", hold=60)
    assert not await b.link_claim("up", hold=60), "two replicas accepted a body for one link"
    await a.link_release("up")
    assert await b.link_claim("up", hold=60), "a released link takes the corrected body"


async def test_concurrent_claims_admit_exactly_one(replicas):
    a, b = replicas
    results = await asyncio.gather(*(s.link_claim("contested", hold=60) for s in (a, b) * 10))
    assert results.count(True) == 1


async def test_a_delivered_upload_spends_its_link_everywhere(replicas):
    a, b = replicas
    exp = time.time() + 60
    await a.link_put("once", b"v", expires=exp, owner="o", owner_cap=1, total_cap=100, tag="")
    assert await a.link_claim("once", hold=60)
    await a.link_consume("once", until=exp)
    assert await b.link_spent("once") and await a.link_spent("once")
    assert not await b.link_claim("once", hold=60)
    assert await b.link_get("once") is None, "the delivered link's value is gone"
    assert await b.link_put("next", b"v", expires=exp, owner="o", owner_cap=1, total_cap=100, tag=""), (
        "a spent link still counted against its owner"
    )


async def test_a_spent_mark_ends_with_the_link(store):
    await store.link_consume("brief", until=time.time() + 0.5)
    assert await store.link_spent("brief")
    await asyncio.sleep(0.7)
    assert not await store.link_spent("brief")


async def test_a_busy_mark_ends_with_its_hold(store):
    assert await store.link_claim("held", hold=0.5)
    await asyncio.sleep(0.7)
    assert await store.link_claim("held", hold=0.5), "a replica that died mid-upload held the link forever"


async def test_a_draft_is_busy_while_an_upload_into_it_runs(replicas):
    a, b = replicas
    exp = time.time() + 60
    await a.link_put("body", b"v", expires=exp, owner="o", owner_cap=10, total_cap=100, tag="draft:o:D1")
    await a.link_put("file", b"v", expires=exp, owner="o", owner_cap=10, total_cap=100, tag="draft:o:D1")
    assert not await b.link_busy_under("draft:o:D1")
    assert await a.link_claim("file", hold=60)
    assert await b.link_busy_under("draft:o:D1")
    assert not await b.link_busy_under("draft:o:D2")
    await a.link_release("file")
    assert not await b.link_busy_under("draft:o:D1")


async def test_sending_a_draft_spends_its_links_and_no_others(replicas):
    a, b = replicas
    exp = time.time() + 60
    for link_id, tag in (("x1", "draft:o:D1"), ("x2", "draft:o:D1"), ("y1", "draft:o:D2")):
        await a.link_put(link_id, b"v", expires=exp, owner="o", owner_cap=10, total_cap=100, tag=tag)
    assert await a.link_claim("x2", hold=60)
    assert await b.link_revoke("draft:o:D1") == 2
    for link_id in ("x1", "x2"):
        assert await a.link_spent(link_id) and await a.link_get(link_id) is None
        assert not await a.link_claim(link_id, hold=60)
    assert not await a.link_busy_under("draft:o:D1")
    assert await a.link_get("y1") == b"v" and not await a.link_spent("y1")
    assert await b.link_revoke("draft:o:D1") == 0
    assert await b.link_revoke("") == 0 and not await b.link_busy_under("")


async def test_no_name_or_value_in_the_store_carries_an_identifier(replicas):
    a, _ = replicas
    if a.kind != "redis":
        pytest.skip("memory holds no names anyone else reads")
    exp = time.time() + 60
    await a.take("person-key-123", capacity=10, refill_per_second=1)
    await a.link_put(
        "link-id-xyz",
        b"sealed",
        expires=exp,
        owner="person-key-123",
        owner_cap=10,
        total_cap=100,
        tag="draft:person-key-123:DRAFT-9",
    )
    await a.link_claim("link-id-xyz", hold=60)
    seen = []
    async for key in a._r.scan_iter(match=f"{a._p}*"):
        seen.append(key)
        kind = await a._r.type(key)
        if kind == b"hash":
            seen += list((await a._r.hgetall(key)).values())
        elif kind == b"zset":
            seen += [m for m, _ in await a._r.zrange(key, 0, -1, withscores=True)]
    blob = b"\n".join(seen)
    assert len(seen) >= 5, "the scan found nothing: it proves nothing"
    for needle in (b"person-key-123", b"link-id-xyz", b"DRAFT-9"):
        assert needle not in blob, needle


async def test_an_unreachable_store_fails_fast_and_says_so():
    from prometheus_client import REGISTRY

    from m365_mcp.state.remote import RedisStore
    from m365_mcp.state.store import Namer

    dead = RedisStore.standalone("redis://127.0.0.1:9/0", Namer(b"k" * 32), "dead:")
    before = (
        REGISTRY.get_sample_value("mcp_state_store_errors_total", {"server": "microsoft-365-mcp", "op": "take"}) or 0
    )
    t0 = time.monotonic()
    with pytest.raises(StoreUnavailable):
        await dead.take("p", capacity=1, refill_per_second=1)
    assert time.monotonic() - t0 < 3.5, "an unreachable store held the call"
    assert (
        REGISTRY.get_sample_value("mcp_state_store_errors_total", {"server": "microsoft-365-mcp", "op": "take"})
        == before + 1
    )
    assert REGISTRY.get_sample_value("mcp_state_store_up", {"server": "microsoft-365-mcp"}) == 0
    await dead.aclose()


async def test_a_sentinel_failover_is_followed(replicas):
    a, b = replicas
    if getattr(a, "_sentinel", None) is None:
        pytest.skip("Sentinel only")
    exp = time.time() + 120
    assert await a.link_put("before-failover", b"kept", expires=exp, owner="o", **T)
    await a._r.execute_command("WAIT", 1, 2000)  # on the replica before the master changes
    old = await a._sentinel.discover_master("mymaster")
    await a._sentinel.sentinels[0].execute_command("SENTINEL", "FAILOVER", "mymaster")
    deadline = time.monotonic() + 30
    now_master, last = old, None
    while True:
        try:
            now_master = await a._sentinel.discover_master("mymaster")
            if now_master != old:
                await b.ping()
                break
        except StoreUnavailable as e:
            last = repr(e.__cause__)
        assert time.monotonic() < deadline, f"no new master within 30 s: {old} -> {now_master}, {last}"
        await asyncio.sleep(0.5)
    deadline = time.monotonic() + 30
    while True:
        try:
            assert await b.link_get("before-failover") == b"kept"
            assert (await a.take("after", capacity=1, refill_per_second=1))[0]
            break
        except StoreUnavailable:
            assert time.monotonic() < deadline, "the store did not come back within 30 s of the failover"
            await asyncio.sleep(0.5)


async def test_a_full_store_refuses_and_the_operation_fails_closed(replicas):
    # With noeviction a full store refuses writes. Every script must be refused whole, before its first
    # write: a link stored past the limit, or a token taken, would be the store failing open.
    a, _ = replicas
    if a.kind != "redis":
        pytest.skip("memory has no limit to reach")
    master = a._r
    await master.config_set("maxmemory-policy", "noeviction")
    await master.config_set("maxmemory", "1")
    try:
        with pytest.raises(StoreUnavailable):
            await a.link_put("oom", b"v", expires=time.time() + 60, owner="o", **T)
        with pytest.raises(StoreUnavailable):
            await a.take("oom-person", capacity=3, refill_per_second=1)
    finally:
        await master.config_set("maxmemory", "0")
    assert await a.link_get("oom") is None, "a write went through a full store"


async def test_a_live_link_is_never_overwritten(replicas):
    # Ids are 128 random bits, so a second put under one id is a bug or an attack: it must not replace
    # the value a link already stands for.
    a, b = replicas
    exp = time.time() + 60
    assert await a.link_put("same-id", b"first", expires=exp, owner="o1", **T)
    assert not await b.link_put("same-id", b"second", expires=exp, owner="o2", **T)
    assert await a.link_get("same-id") == b"first"


async def test_an_upload_admitted_before_its_link_expired_still_holds_the_draft(replicas):
    # The body may take minutes after the link's own expiry: sending the draft meanwhile would send
    # it without what is still arriving.
    a, b = replicas
    await a.link_put("late", b"v", expires=time.time() + 0.6, owner="o", owner_cap=10, total_cap=100, tag="draft:o:D7")
    assert await a.link_claim("late", hold=60)
    await asyncio.sleep(0.8)
    assert await b.link_get("late") is None, "the link itself has expired"
    assert await b.link_busy_under("draft:o:D7"), "the running upload no longer holds its draft"
    await b.link_revoke("draft:o:D7")
    assert await a.link_spent("late") and not await a.link_busy_under("draft:o:D7")


async def test_expired_links_do_not_pile_up_under_one_draft(replicas):
    a, _ = replicas
    if a.kind != "redis":
        pytest.skip("memory prunes on every operation")
    for i in range(5):
        await a.link_put(
            f"gone-{i}", b"v", expires=time.time() + 0.3, owner="o", owner_cap=10, total_cap=100, tag="draft:o:D8"
        )
    await asyncio.sleep(0.5)
    await a.link_put("live", b"v", expires=time.time() + 60, owner="o", owner_cap=10, total_cap=100, tag="draft:o:D8")
    assert await a._r.zcard(a._k("tag", "draft:o:D8")) == 1


async def test_during_a_rotation_pods_holding_the_keys_in_either_order_share_one_allowance(replicas):
    # Phase 2 of a rotation: some pods name things under [old, new], some under [new, old]. A send on
    # either draws from both buckets, so ten sends in all, wherever they land.
    a, b = replicas
    old, new = Namer(b"o" * 32), Namer(b"n" * 32)
    before = a.view(old)  # names under the old key, opens with the new one too
    after = b.view(new)
    results = []
    for i in range(12):
        store, also = (before, (new,)) if i % 2 else (after, (old,))
        results.append((await store.take("rotating", capacity=10, refill_per_second=0.001, also=also))[0])
    assert results.count(True) == 10 and results[-2:] == [False, False]
