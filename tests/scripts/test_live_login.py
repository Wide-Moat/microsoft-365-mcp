"""scripts/live_login.py: one user at a time, merged into the token file, never two copies of one account."""

from __future__ import annotations

import base64
import importlib.util
import json
from pathlib import Path

import httpx
import pytest

SPEC = importlib.util.spec_from_file_location("live_login", Path(__file__).parents[2] / "scripts" / "live_login.py")
live_login = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(live_login)


def jwt(oid: str, upn: str) -> str:
    body = base64.urlsafe_b64encode(json.dumps({"oid": oid, "upn": upn}).encode()).decode().rstrip("=")
    return f"h.{body}.s"


def fake_signin(oid: str, upn: str):
    def device_login(http, client_id, scope, label):
        return {"access_token": jwt(oid, upn), "refresh_token": "r-" + oid, "expires_at": 9999999999, "scope": scope}
    return device_login


@pytest.fixture
def out(tmp_path: Path) -> Path:
    return tmp_path / "tokens.json"


def test_user_option_signs_in_only_that_user_and_keeps_the_other(out, monkeypatch):
    out.write_text(json.dumps({"a": {"access_token": jwt("oid-a", "a@x.test"), "refresh_token": "ra", "expires_at": 1}}))
    monkeypatch.setattr(live_login, "device_login", fake_signin("oid-b", "b@x.test"))
    assert live_login.main(["--client-id", "c", "--out", str(out), "--user", "b"]) == 0
    data = json.loads(out.read_text())
    assert set(data) == {"a", "b"}
    assert data["a"]["refresh_token"] == "ra"
    assert live_login.claims(data["b"]["access_token"])["oid"] == "oid-b"


def test_user_option_on_a_missing_file_writes_just_that_user(out, monkeypatch):
    monkeypatch.setattr(live_login, "device_login", fake_signin("oid-a", "a@x.test"))
    assert live_login.main(["--client-id", "c", "--out", str(out), "--user", "a"]) == 0
    assert set(json.loads(out.read_text())) == {"a"}


def test_user_option_refuses_the_same_account_as_the_other_user(out, monkeypatch):
    out.write_text(json.dumps({"a": {"access_token": jwt("same", "a@x.test"), "refresh_token": "ra", "expires_at": 1}}))
    before = out.read_text()
    monkeypatch.setattr(live_login, "device_login", fake_signin("same", "a@x.test"))
    with pytest.raises(SystemExit) as e:
        live_login.main(["--client-id", "c", "--out", str(out), "--user", "b"])
    assert e.value.code == 1
    assert out.read_text() == before


def test_refresh_renews_only_the_users_present(out, monkeypatch):
    out.write_text(json.dumps({"a": {"access_token": jwt("oid-a", "a@x.test"), "refresh_token": "ra", "expires_at": 1}}))

    def refresh(http, client_id, entry, label):
        return {**entry, "expires_at": 9999999999}
    monkeypatch.setattr(live_login, "refresh", refresh)
    assert live_login.main(["--client-id", "c", "--out", str(out), "--refresh"]) == 0
    assert json.loads(out.read_text())["a"]["expires_at"] == 9999999999


def test_unknown_user_label_is_refused(out):
    with pytest.raises(SystemExit) as e:
        live_login.main(["--client-id", "c", "--out", str(out), "--user", "c"])
    assert e.value.code == 2


def _sign_in_capturing_urls(monkeypatch, out, argv):
    """Run main() against a fake token endpoint and return the URLs it posted to."""
    seen: list[str] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(str(request.url))
        if request.url.path.endswith("/devicecode"):
            return httpx.Response(200, json={"device_code": "d", "user_code": "U", "verification_uri": "https://v",
                                             "expires_in": 900, "interval": 0})
        return httpx.Response(200, json={"access_token": jwt("oid-" + str(len(seen)), "u@x.test"),
                                         "refresh_token": "r", "expires_in": 3600})

    real_client = httpx.Client
    monkeypatch.setattr(live_login.httpx, "Client", lambda **kw: real_client(transport=httpx.MockTransport(handler)))
    monkeypatch.setattr(live_login.time, "sleep", lambda s: None)
    assert live_login.main(argv) == 0
    return seen


def test_tenant_option_pins_sign_in_to_that_tenant(out, monkeypatch):
    # /organizations lets the browser's current account answer, and a person signed in to another
    # organisation gets AADSTS5000228 for a tenant they never meant (seen 2026-10-24).
    monkeypatch.delenv("M365_LIVE_TENANT", raising=False)
    seen = _sign_in_capturing_urls(monkeypatch, out, ["--client-id", "c", "--out", str(out), "--user", "a",
                                                      "--tenant", "contoso.onmicrosoft.com"])
    assert seen and all(u.startswith("https://login.microsoftonline.com/contoso.onmicrosoft.com/oauth2/v2.0/") for u in seen)


def test_tenant_comes_from_the_environment_and_defaults_to_organizations(out, monkeypatch):
    monkeypatch.setenv("M365_LIVE_TENANT", "04ee5f53-8c4d-42c3-bbf1-85a951bfddfd")
    seen = _sign_in_capturing_urls(monkeypatch, out, ["--client-id", "c", "--out", str(out), "--user", "a"])
    assert all("/04ee5f53-8c4d-42c3-bbf1-85a951bfddfd/oauth2/v2.0/" in u for u in seen)
    monkeypatch.delenv("M365_LIVE_TENANT")
    seen = _sign_in_capturing_urls(monkeypatch, out, ["--client-id", "c", "--out", str(out), "--user", "b"])
    assert all("/organizations/oauth2/v2.0/" in u for u in seen)


@pytest.mark.parametrize("bad", ["", "a/b", "x?y", "../common", "a b"])
def test_tenant_that_is_not_a_tenant_name_is_refused(out, monkeypatch, capsys, bad):
    with pytest.raises(SystemExit):
        live_login.main(["--client-id", "c", "--out", str(out), "--tenant", bad])
    assert "--tenant must be a tenant id or domain" in capsys.readouterr().err
    assert not out.exists()
