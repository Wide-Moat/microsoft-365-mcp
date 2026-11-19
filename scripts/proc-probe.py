"""Run inside the server container as its user: can this shell read the server process (PID 1)?

Prints three words: its own /proc/self/environ, then PID 1's environ and maps. A healthy result is
`READ DENIED DENIED`; anything else, including an error other than a refusal, is a failure the caller
sees, so a dead container or a missing file cannot pass for a closed one.
"""

import sys


def probe(path: str) -> str:
    try:
        with open(path, "rb") as f:
            f.read(1)
        return "READ"
    except PermissionError:
        return "DENIED"
    except Exception as e:  # noqa: BLE001 - every other outcome is reported, never taken as denial
        return f"OTHER:{type(e).__name__}"


print(probe("/proc/self/environ"), probe("/proc/1/environ"), probe("/proc/1/maps"))
sys.exit(0)
