"""Every line every process of the image writes is one JSON object: the server from its first line to its
last, a start that fails on its settings, and the init container's staging step, both ways. The real
processes are started, not their functions called, so a line written outside logging (a print, a library's
own handler, an uncaught exception) fails here too."""

from __future__ import annotations

import json
import os
import signal
import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
KEY = "k1:" + "A" * 43  # 32 zero bytes, base64url


def _env(**extra: str) -> dict[str, str]:
    env = {k: v for k, v in os.environ.items() if not k.startswith(("M365_", "LOG_LEVEL"))}
    env.update({"PYTHONPATH": str(ROOT / "src"), "PYTHONUNBUFFERED": "1", **extra})
    return env


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _lines(text: str) -> list[dict]:
    lines = [ln for ln in text.splitlines() if ln.strip()]
    parsed = []
    for ln in lines:
        rec = json.loads(ln)  # a line that is not JSON fails the test here, with the line in the report
        assert isinstance(rec, dict) and {"ts", "level", "logger", "event", "msg"} <= set(rec), ln
        assert rec["ts"].endswith("Z") and "T" in rec["ts"], ln
        parsed.append(rec)
    return parsed


def _key_file(tmp_path: Path, mode: int = 0o400) -> Path:
    f = tmp_path / "key"
    f.write_text(KEY + "\n")
    f.chmod(mode)
    return f


def test_a_running_server_writes_only_json_lines(tmp_path):
    ports = [_free_port() for _ in range(3)]
    proc = subprocess.Popen(
        [sys.executable, "-m", "m365_mcp.main"],
        cwd=tmp_path,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        env=_env(
            M365_KEY_FILE=str(_key_file(tmp_path)),
            M365_LINKS_BASE_URL="https://links.example.com",
            PORT=str(ports[0]),
            LINKS_PORT=str(ports[1]),
            M365_METRICS_PORT=str(ports[2]),
            LOG_LEVEL="DEBUG",
        ),
    )
    try:
        deadline = time.monotonic() + 30
        while True:
            try:
                urllib.request.urlopen(f"http://127.0.0.1:{ports[2]}/healthz", timeout=2).read()
                break
            except OSError:
                assert proc.poll() is None, "the server stopped"
                assert time.monotonic() < deadline, "the server did not answer within 30 s"
                time.sleep(0.2)
        for url in (
            f"http://127.0.0.1:{ports[0]}/readyz",
            f"http://127.0.0.1:{ports[1]}/d/not-a-link",
            f"http://127.0.0.1:{ports[2]}/metrics",
        ):
            try:
                urllib.request.urlopen(url, timeout=5).read()
            except urllib.error.HTTPError:
                pass
    finally:
        proc.send_signal(signal.SIGTERM)
        out, _ = proc.communicate(timeout=30)
    records = _lines(out.decode())
    events = {r["event"] for r in records}
    assert len(records) >= 2 and {"start", "config_warning"} <= events, events


def test_a_start_refused_for_its_settings_is_one_json_line_and_exit_2(tmp_path):
    key = _key_file(tmp_path, mode=0o644)  # readable by others: refused
    done = subprocess.run(
        [sys.executable, "-m", "m365_mcp.main"],
        cwd=tmp_path,
        capture_output=True,
        timeout=60,
        env=_env(M365_KEY_FILE=str(key), M365_LINKS_BASE_URL="https://links.example.com"),
    )
    assert done.returncode == 2
    records = _lines(done.stdout.decode() + done.stderr.decode())
    assert [r["event"] for r in records] == ["config_error"]
    assert "M365_KEY_FILE" in records[0]["msg"] and KEY not in done.stderr.decode()


def test_a_bad_log_level_is_a_json_refusal_too(tmp_path):
    done = subprocess.run(
        [sys.executable, "-m", "m365_mcp.main"],
        cwd=tmp_path,
        capture_output=True,
        timeout=60,
        env=_env(LOG_LEVEL="LOUD", M365_AREAS="account"),
    )
    assert done.returncode == 2
    assert [r["event"] for r in _lines(done.stdout.decode() + done.stderr.decode())] == ["config_error"]


def test_an_uncaught_exception_is_one_json_line_without_its_message(tmp_path):
    code = "from m365_mcp import logjson; logjson.install(); raise RuntimeError('secret.person@contoso.test')"
    done = subprocess.run([sys.executable, "-c", code], capture_output=True, timeout=60, env=_env())
    assert done.returncode != 0
    out = done.stdout.decode() + done.stderr.decode()
    records = _lines(out)
    assert [r["event"] for r in records] == ["uncaught"] and "RuntimeError" in records[0]["exc"]
    assert "secret.person" not in out


def test_the_staging_step_writes_json_both_ways(tmp_path):
    src, dst = tmp_path / "secret", tmp_path / "run"
    src.mkdir()
    dst.mkdir()
    (src / "key").write_text(KEY + "\n")
    (src / "password").write_text("store-password\n")
    ok = subprocess.run(
        [sys.executable, "-m", "m365_mcp.keys.stage", str(src), str(dst)], capture_output=True, timeout=60, env=_env()
    )
    assert ok.returncode == 0
    staged = _lines(ok.stdout.decode() + ok.stderr.decode())
    assert [r["event"] for r in staged] == ["staged"] and staged[0]["names"] == ["key", "password"]
    assert "store-password" not in ok.stdout.decode() and KEY not in ok.stdout.decode()
    assert oct((dst / "key").stat().st_mode & 0o777) == "0o400"
    empty = tmp_path / "empty"
    empty.mkdir()
    bad = subprocess.run(
        [sys.executable, "-m", "m365_mcp.keys.stage", str(empty), str(dst)], capture_output=True, timeout=60, env=_env()
    )
    assert bad.returncode == 2
    assert [r["event"] for r in _lines(bad.stdout.decode() + bad.stderr.decode())] == ["stage_empty"]
