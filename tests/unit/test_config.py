import os

import pytest

from m365_mcp.config import ConfigError, Settings


def env(monkeypatch, **kw):
    import os

    for k in [k for k in os.environ if k.startswith("M365_")]:
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


# -- the key file --------------------------------------------------------------------------------

KEY_LINE = "k1:" + "A" * 43
LINKS = "https://links.example.com"


def key_file(tmp_path, text=KEY_LINE + "\n", mode=0o400, name="key"):
    f = tmp_path / name
    f.write_text(text)
    f.chmod(mode)
    return str(f)


def test_a_key_file_gives_the_link_keys_and_the_stores_key(monkeypatch, tmp_path):
    env(monkeypatch, M365_KEY_FILE=key_file(tmp_path), M365_LINKS_BASE_URL=LINKS)
    s = Settings.from_env()
    assert s.links_ready and s.key_source == "file" and len(s.state_key) == 32
    assert len({s.sign_keys[0], s.enc_keys[0], s.state_key}) == 3, "one purpose, one key"


def test_a_rotated_key_file_opens_with_every_line_and_seals_with_the_first(monkeypatch, tmp_path):
    two = "k2:" + "B" * 42 + "A\n" + KEY_LINE + "\n"
    env(monkeypatch, M365_KEY_FILE=key_file(tmp_path, two), M365_LINKS_BASE_URL=LINKS)
    s = Settings.from_env()
    env(monkeypatch, M365_KEY_FILE=key_file(tmp_path, name="old"), M365_LINKS_BASE_URL=LINKS)
    old = Settings.from_env()
    assert len(s.sign_keys) == 2 and s.sign_keys[1] == old.sign_keys[0] and s.sign_keys[0] != old.sign_keys[0]


@pytest.mark.parametrize("mode", [0o440, 0o404, 0o600 | 0o004])
def test_a_key_file_others_may_read_is_refused(monkeypatch, tmp_path, mode):
    env(monkeypatch, M365_KEY_FILE=key_file(tmp_path, mode=mode), M365_LINKS_BASE_URL=LINKS)
    with pytest.raises(ConfigError, match="M365_KEY_FILE: the file is readable by a group or by others"):
        Settings.from_env()


def test_a_key_file_of_another_user_is_refused(monkeypatch, tmp_path):
    path = key_file(tmp_path)
    real = os.geteuid()
    monkeypatch.setattr(os, "geteuid", lambda: real + 1)
    env(monkeypatch, M365_KEY_FILE=path, M365_LINKS_BASE_URL=LINKS)
    with pytest.raises(ConfigError, match="belongs to another user"):
        Settings.from_env()


def test_a_key_file_that_is_a_link_is_refused(monkeypatch, tmp_path):
    # A Secret mounted straight into the container is a symbolic link: the init container's copy is not.
    real = key_file(tmp_path)
    link = tmp_path / "linked"
    link.symlink_to(real)
    env(monkeypatch, M365_KEY_FILE=str(link), M365_LINKS_BASE_URL=LINKS)
    with pytest.raises(ConfigError, match="not a link"):
        Settings.from_env()


@pytest.mark.parametrize(
    "text", ["", "k1:short\n", "k1\n", KEY_LINE + "\n" + KEY_LINE + "\n", "k1:" + "A" * 42 + "B\n"]
)
def test_a_key_file_that_is_not_keys_is_refused(monkeypatch, tmp_path, text):
    env(monkeypatch, M365_KEY_FILE=key_file(tmp_path, text), M365_LINKS_BASE_URL=LINKS)
    with pytest.raises(ConfigError, match="M365_KEY_FILE"):
        Settings.from_env()


def test_a_key_file_larger_than_any_key_is_refused(monkeypatch, tmp_path):
    env(monkeypatch, M365_KEY_FILE=key_file(tmp_path, "x" * 70_000), M365_LINKS_BASE_URL=LINKS)
    with pytest.raises(ConfigError, match="larger than"):
        Settings.from_env()


def test_a_key_in_the_environment_still_works_with_a_warning(monkeypatch):
    env(monkeypatch, M365_LINK_SECRET="x" * 64, M365_LINKS_BASE_URL=LINKS)
    s = Settings.from_env()
    assert s.key_source == "env" and any("refused from 0.3.0" in w for w in s.warnings)


def test_a_key_file_and_a_key_in_the_environment_together_are_refused(monkeypatch, tmp_path):
    env(monkeypatch, M365_KEY_FILE=key_file(tmp_path), M365_LINK_SECRET="x" * 64, M365_LINKS_BASE_URL=LINKS)
    with pytest.raises(ConfigError, match="remove the environment variables"):
        Settings.from_env()


# -- the shared state store ----------------------------------------------------------------------


def test_no_store_is_memory_with_a_warning(monkeypatch, tmp_path):
    env(monkeypatch, M365_KEY_FILE=key_file(tmp_path), M365_LINKS_BASE_URL=LINKS)
    s = Settings.from_env()
    assert s.state is None and any("one replica only" in w for w in s.warnings)


def test_a_standalone_store(monkeypatch, tmp_path):
    env(
        monkeypatch,
        M365_KEY_FILE=key_file(tmp_path),
        M365_LINKS_BASE_URL=LINKS,
        M365_STATE_URL="rediss://valkey.store.svc:6380/2",
        M365_STATE_USERNAME="mcp",
    )
    st = Settings.from_env().state
    assert st.url.startswith("rediss://") and st.tls and st.username == "mcp" and st.prefix == "microsoft-365-mcp:"


def test_sentinels_one_host_is_enough(monkeypatch, tmp_path):
    env(
        monkeypatch,
        M365_KEY_FILE=key_file(tmp_path),
        M365_LINKS_BASE_URL=LINKS,
        M365_STATE_SENTINELS="app-redis-sentinel:26379",
    )
    st = Settings.from_env().state
    assert st.sentinels == (("app-redis-sentinel", 26379),) and st.master == "mymaster" and st.password is None


def test_a_url_and_sentinels_together_are_refused(monkeypatch, tmp_path):
    env(
        monkeypatch,
        M365_KEY_FILE=key_file(tmp_path),
        M365_LINKS_BASE_URL=LINKS,
        M365_STATE_URL="redis://r:6379/0",
        M365_STATE_SENTINELS="s:26379",
    )
    with pytest.raises(ConfigError, match="set one"):
        Settings.from_env()


@pytest.mark.parametrize("url", ["redis://:pw@r:6379/0", "redis://user:pw@r:6379/0", "redis://user@r:6379/0"])
def test_credentials_in_the_url_are_refused(monkeypatch, tmp_path, url):
    env(monkeypatch, M365_KEY_FILE=key_file(tmp_path), M365_LINKS_BASE_URL=LINKS, M365_STATE_URL=url)
    with pytest.raises(ConfigError, match="M365_STATE_PASSWORD_FILE") as e:
        Settings.from_env()
    assert "pw" not in str(e.value)


@pytest.mark.parametrize("name", ["M365_STATE_PASSWORD", "M365_STATE_SENTINEL_PASSWORD"])
def test_a_password_in_the_environment_is_refused(monkeypatch, tmp_path, name):
    env(
        monkeypatch,
        M365_KEY_FILE=key_file(tmp_path),
        M365_LINKS_BASE_URL=LINKS,
        M365_STATE_URL="redis://r:6379/0",
        **{name: "hunter2"},
    )
    with pytest.raises(ConfigError, match="never read from the environment"):
        Settings.from_env()


def test_passwords_come_from_files_the_server_alone_reads(monkeypatch, tmp_path):
    pw, spw = tmp_path / "password", tmp_path / "sentinelPassword"
    pw.write_text("store-pass\n")
    spw.write_text("sentinel-pass")
    pw.chmod(0o400)
    spw.chmod(0o400)
    env(
        monkeypatch,
        M365_KEY_FILE=key_file(tmp_path),
        M365_LINKS_BASE_URL=LINKS,
        M365_STATE_SENTINELS="s:26379",
        M365_STATE_PASSWORD_FILE=str(pw),
        M365_STATE_SENTINEL_PASSWORD_FILE=str(spw),
    )
    st = Settings.from_env().state
    assert st.password == "store-pass" and st.sentinel_password == "sentinel-pass"
    assert "store-pass" not in repr(st) and "sentinel-pass" not in repr(st)
    pw.chmod(0o644)
    with pytest.raises(ConfigError, match="M365_STATE_PASSWORD_FILE: the file is readable"):
        Settings.from_env()


@pytest.mark.parametrize("value", ["s", "s:", "s:0", "s:99999", "s:2x"])
def test_a_sentinel_that_is_not_host_port_is_refused(monkeypatch, tmp_path, value):
    env(monkeypatch, M365_KEY_FILE=key_file(tmp_path), M365_LINKS_BASE_URL=LINKS, M365_STATE_SENTINELS=value)
    with pytest.raises(ConfigError, match="M365_STATE_SENTINELS"):
        Settings.from_env()


def test_a_store_without_a_key_to_name_things_is_refused(monkeypatch):
    env(monkeypatch, M365_AREAS="account", M365_STATE_URL="redis://r:6379/0")
    with pytest.raises(ConfigError, match="needs the sealing key"):
        Settings.from_env()


def test_a_ca_bundle_must_be_certificates(monkeypatch, tmp_path):
    bad = tmp_path / "ca.pem"
    bad.write_text("not a certificate")
    env(monkeypatch, M365_KEY_FILE=key_file(tmp_path), M365_LINKS_BASE_URL=LINKS, M365_CA_BUNDLE=str(bad))
    with pytest.raises(ConfigError, match="M365_CA_BUNDLE"):
        Settings.from_env()


@pytest.mark.parametrize("query", ["?password=secret", "?ssl_cert_reqs=none", "#x"])
def test_options_in_the_store_url_are_refused(monkeypatch, tmp_path, query):
    # redis-py reads them from the URL and they win: a password past the file, certificate checks off.
    env(
        monkeypatch,
        M365_KEY_FILE=key_file(tmp_path),
        M365_LINKS_BASE_URL=LINKS,
        M365_STATE_URL="rediss://store:6380/0" + query,
    )
    with pytest.raises(ConfigError, match="no query or fragment"):
        Settings.from_env()
