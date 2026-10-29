import pytest

from m365_mcp.config import ConfigError, Settings


def env(monkeypatch, **kw):
    for k in ("M365_AREAS", "M365_SWITCHES", "M365_LINK_SECRET", "M365_LINK_SIGNING_KEYS",
              "M365_LINK_ENCRYPTION_KEYS", "M365_LINKS_BASE_URL"):
        monkeypatch.delenv(k, raising=False)
    for k, v in kw.items():
        monkeypatch.setenv(k, v)


def test_one_secret_gives_two_different_keys(monkeypatch):
    env(monkeypatch, M365_LINK_SECRET="x" * 64, M365_LINKS_BASE_URL="http://m365-mcp.prod.svc.cluster.local:8001")
    s = Settings.from_env()
    assert s.links_ready and s.sign_keys[0] != s.enc_keys[0] and len(s.sign_keys[0]) == 32


def test_byte_areas_without_keys_refuse_to_start(monkeypatch):
    env(monkeypatch)
    with pytest.raises(ConfigError):
        Settings.from_env()


@pytest.mark.parametrize("url", ["http://links.example.com", "http://10.0.0.5:8001", "ftp://x.svc.cluster.local"])
def test_plain_http_is_only_for_cluster_internal_names(monkeypatch, url):
    env(monkeypatch, M365_LINK_SECRET="x" * 64, M365_LINKS_BASE_URL=url)
    with pytest.raises(ConfigError):
        Settings.from_env()


def test_short_secret_is_refused(monkeypatch):
    env(monkeypatch, M365_LINK_SECRET="short", M365_LINKS_BASE_URL="https://l.example.com")
    with pytest.raises(ConfigError):
        Settings.from_env()


def test_account_alone_needs_no_keys(monkeypatch):
    env(monkeypatch, M365_AREAS="account")
    assert Settings.from_env().areas == frozenset({"account"})


def test_unknown_area_is_refused(monkeypatch):
    env(monkeypatch, M365_AREAS="mail,fax")
    with pytest.raises(ConfigError):
        Settings.from_env()


@pytest.mark.parametrize("url", ["https:///links", "https://links.example.com/?x=1", "https://links.example.com/#f"])
def test_malformed_links_base_is_refused(monkeypatch, url):
    env(monkeypatch, M365_LINK_SECRET="x" * 64, M365_LINKS_BASE_URL=url)
    with pytest.raises(ConfigError):
        Settings.from_env()
