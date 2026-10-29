"""Fail when tracked files name a real organisation's data (manifesto rule 41).

    python scripts/purity_check.py [root]

Looks at every file git tracks and reports, with file and line:

- an e-mail address outside the domains reserved for examples (example.com/.org/.net, *.test,
  *.invalid, localhost) and Microsoft's documentation tenants;
- a Microsoft 365 tenant name: <tenant>.onmicrosoft.com or <tenant>[-my].sharepoint.com, unless the
  tenant is one of Microsoft's documentation placeholders (contoso, fabrikam, ...);
- a public IPv4 address (private, loopback, link-local and the RFC 5737 documentation ranges pass).

Exit codes: 0 clean, 1 findings, 2 nothing was scanned (a check that read no file cannot pass).
"""
from __future__ import annotations

import ipaddress
import re
import subprocess
import sys
from pathlib import Path

PLACEHOLDER_TENANTS = ("contoso", "fabrikam", "northwind", "northwindtraders", "adatum", "tailspintoys",
                       "wingtiptoys", "woodgrove", "woodgrovebank", "litware", "x", "tenant", "yourtenant",
                       "your-tenant", "example")
EMAIL = re.compile(r"(?<![\w.%+-])[\w.%+-]+@([A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+)")
TENANT = re.compile(r"(?<![\w.*<{-])([A-Za-z0-9][A-Za-z0-9-]*)\.(?:sharepoint|onmicrosoft)\.com", re.I)
IPV4 = re.compile(r"(?<![\d.])(\d{1,3}(?:\.\d{1,3}){3})(?![\d.]|/\d)")
DOC_NETS = [ipaddress.ip_network(n) for n in ("192.0.2.0/24", "198.51.100.0/24", "203.0.113.0/24")]


def _tenant_ok(name: str) -> bool:
    n = name.lower().removesuffix("-my").removesuffix("-admin").removesuffix("-files")
    return n in PLACEHOLDER_TENANTS or re.fullmatch(r"m365x\d+", n) is not None


def _email_ok(domain: str) -> bool:
    d = domain.lower()
    if d in ("thread.v2", "thread.tacv2", "unq.gbl.spaces"):
        return True  # a Teams conversation id (19:...@thread.v2), not an address
    if d.endswith((".test", ".invalid", ".localhost", ".example")) or d in ("localhost", "example.com", "example.org", "example.net"):
        return True
    if d.endswith((".example.com", ".example.org", ".example.net")):
        return True
    m = re.fullmatch(r"([a-z0-9-]+)\.(?:onmicrosoft\.com|com)", d)
    return bool(m and _tenant_ok(m.group(1)))


def _ip_ok(text: str) -> bool:
    try:
        ip = ipaddress.ip_address(text)
    except ValueError:
        return True  # 999.1.1.1 is a version string, not an address
    return (ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_unspecified or ip.is_multicast
            or str(ip) == "255.255.255.255" or any(ip in n for n in DOC_NETS))


def findings(line: str) -> list[str]:
    out = [f"e-mail {m.group(0)}" for m in EMAIL.finditer(line) if not _email_ok(m.group(1))]
    out += [f"tenant {m.group(0)}" for m in TENANT.finditer(line) if not _tenant_ok(m.group(1))]
    out += [f"address {m.group(1)}" for m in IPV4.finditer(line) if not _ip_ok(m.group(1))]
    return out


def main(root: str = ".") -> int:
    files = subprocess.run(["git", "-C", root, "ls-files", "-z"], capture_output=True, check=True).stdout.split(b"\0")
    files = [f.decode() for f in files if f]
    scanned, report = 0, []
    for name in files:
        path = Path(root, name)
        # This script and its test hold the patterns and the fixtures that must fail.
        if name in ("scripts/purity_check.py", "tests/unit/test_purity_check.py") or not path.is_file():
            continue
        data = path.read_bytes()
        if b"\0" in data:
            continue
        scanned += 1
        for no, line in enumerate(data.decode("utf-8", "replace").splitlines(), 1):
            report += [f"{name}:{no}: {f}" for f in findings(line)]
    print(f"purity: {scanned} of {len(files)} tracked files scanned")
    if scanned == 0:
        print("purity: nothing was scanned")
        return 2
    for r in report:
        print(r)
    print("purity: clean" if not report else f"purity: {len(report)} finding(s)")
    return 1 if report else 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
