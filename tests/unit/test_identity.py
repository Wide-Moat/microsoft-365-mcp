"""Which tokens are accepted: Graph's audience always, and, when the installation names them, only
tokens issued to its own gateway's client."""

from dataclasses import replace

import pytest
from helpers import make_token

from m365_mcp.config import Settings
from m365_mcp.core.context import configure
from m365_mcp.core.errors import M365Error
from m365_mcp.identity import current_token, person_from_token

GATEWAY = "11111111-2222-3333-4444-555555555555"
OTHER = "99999999-8888-7777-6666-555555555555"


@pytest.fixture
def only_gateway(rt):
    rt.settings = replace(rt.settings, expected_client_ids=frozenset({GATEWAY}))
    configure(rt)
    return rt


async def _connection(graph, invoke):
    graph.on("GET", "/me", {"displayName": "Ann"})
    graph.on("GET", "/me/mailboxSettings", {})
    return await invoke("get_connection")


async def test_a_token_for_another_audience_is_refused_before_graph(graph, invoke):
    current_token.set(make_token(aud="api://something-else"))
    r = await invoke("get_connection")
    assert r.error and "Reconnect" in r.text
    assert graph.seen == []


@pytest.mark.parametrize("token", ["garbage", "a.b", "a.b.c.d", "x.bm90IGpzb24.y"])
async def test_a_token_that_is_not_a_jwt_is_refused_before_graph(graph, invoke, token):
    current_token.set(token)
    r = await invoke("get_connection")
    assert r.error and "Reconnect" in r.text
    assert graph.seen == []


@pytest.mark.parametrize("missing", ["tid", "oid"])
async def test_a_token_without_a_tenant_or_object_id_is_refused(graph, invoke, missing):
    current_token.set(make_token(**{missing: ""}))
    r = await invoke("get_connection")
    assert r.error and "Reconnect" in r.text
    assert graph.seen == []


async def test_a_token_issued_to_another_client_is_refused_before_graph(graph, invoke, only_gateway):
    current_token.set(make_token(azp=OTHER))
    r = await invoke("get_connection")
    assert r.error and "Reconnect" in r.text
    assert graph.seen == []


async def test_a_token_naming_no_client_is_refused_when_clients_are_configured(graph, invoke, only_gateway):
    current_token.set(make_token())
    r = await invoke("get_connection")
    assert r.error and "Reconnect" in r.text
    assert graph.seen == []


async def test_a_token_issued_to_the_gateway_is_accepted(graph, invoke, only_gateway):
    current_token.set(make_token(azp=GATEWAY))
    r = await _connection(graph, invoke)
    assert not r.error, r.text


async def test_a_v1_token_names_its_client_in_appid(graph, invoke, only_gateway):
    current_token.set(make_token(appid=GATEWAY))
    r = await _connection(graph, invoke)
    assert not r.error, r.text


async def test_a_v1_token_from_another_client_is_refused(graph, invoke, only_gateway):
    current_token.set(make_token(appid=OTHER))
    r = await invoke("get_connection")
    assert r.error and "Reconnect" in r.text
    assert graph.seen == []


async def test_without_expected_clients_any_client_is_accepted(graph, invoke):
    current_token.set(make_token(azp=OTHER))
    r = await _connection(graph, invoke)
    assert not r.error, r.text


def test_client_ids_compare_without_case():
    p = person_from_token(make_token(azp=GATEWAY.upper()), clients=frozenset({GATEWAY}))
    assert p.object_id == "oid-a"
    with pytest.raises(M365Error):
        person_from_token(make_token(azp=OTHER), clients=frozenset({GATEWAY}))


def test_expected_client_ids_are_read_as_a_comma_list(monkeypatch):
    monkeypatch.setenv("M365_AREAS", "account")
    monkeypatch.setenv("M365_EXPECTED_CLIENT_IDS", f" {GATEWAY.upper()} , {OTHER},")
    assert Settings.from_env().expected_client_ids == frozenset({GATEWAY, OTHER})


def test_expected_client_ids_unset_means_none(monkeypatch):
    monkeypatch.setenv("M365_AREAS", "account")
    monkeypatch.delenv("M365_EXPECTED_CLIENT_IDS", raising=False)
    assert Settings.from_env().expected_client_ids == frozenset()
