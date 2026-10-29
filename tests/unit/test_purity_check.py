"""The purity check (manifesto rule 41) refuses real data and passes placeholders."""
import importlib.util
import subprocess
from pathlib import Path

import pytest

# Assembled from parts, so the fixtures that must fail do not trip a scan of this file.
REAL_ADDRESS = "someone" + "@" + "acme.io"
REAL_IP = ".".join(["8", "8", "4", "4"])

_spec = importlib.util.spec_from_file_location("purity_check", Path(__file__).parents[2] / "scripts" / "purity_check.py")
purity = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(purity)


@pytest.mark.parametrize("line", [
    f'owner = "{REAL_ADDRESS}"',
    "https://acmecorp-my.sharepoint.com/personal/x",
    "login.microsoftonline.com/acmecorp.onmicrosoft.com/oauth2",
    "both contoso.sharepoint.com and acmecorp.sharepoint.com",
    f"server {REAL_IP} answered",
])
def test_real_data_is_a_finding(line):
    assert purity.findings(line)


@pytest.mark.parametrize("line", [
    "19:abc" + "@thread.v2 and 19:x" + "@thread.tacv2",
    "me@x.test, a@example.com, admin@contoso.onmicrosoft.com, megan@contoso.com",
    "https://contoso.sharepoint.com/sites/x and x.sharepoint.com,aaaa and *.sharepoint.com",
    "M365x123456.onmicrosoft.com and https://<tenant>.sharepoint.com",
    "bind 0.0.0.0, 127.0.0.1, 10.0.0.1, 192.168.1.1, 169.254.169.254, 203.0.113.7, 10.0.0.0/8",
    "version 1.2.3.4.5 and 999.1.1.1",
])
def test_placeholders_pass(line):
    assert purity.findings(line) == []


def test_a_repository_with_a_finding_fails_and_an_empty_one_is_not_clean(tmp_path):
    def repo(name, content):
        d = tmp_path / name
        d.mkdir()
        subprocess.run(["git", "init", "-q", str(d)], check=True)
        if content is not None:
            (d / "f.txt").write_text(content)
            subprocess.run(["git", "-C", str(d), "add", "-A"], check=True)
        return str(d)
    assert purity.main(repo("bad", REAL_ADDRESS + "\n")) == 1
    assert purity.main(repo("good", "me@x.test\n")) == 0
    assert purity.main(repo("empty", None)) == 2
