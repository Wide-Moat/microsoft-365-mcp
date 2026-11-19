"""Run by the chart's init container: copy the mounted Secrets into a memory volume.

The server container then mounts only that volume, never a Secret: each key of it is a file readable by
the server's user alone (0400), on a tmpfs that disappears with the pod. A container restart finds the
files still there; a new pod copies the Secrets again.

    python -m m365_mcp.keys.stage <secret dir> <memory dir>

It writes one JSON line naming how many files it copied and their names, never a value.
"""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

from m365_mcp.keys.files import MAX_SECRET_FILE
from m365_mcp.logjson import emit


def copy(src: Path, dst: Path) -> int:
    """Copy each key of a Secret volume to `dst`, mode 0400, replacing atomically. Returns how many.
    Names starting with `..` (`..data`, the timestamped dir) are Kubernetes' own; a key may start with
    one dot."""
    names = sorted(n for n in os.listdir(src) if not n.startswith("..") and (src / n).is_file())
    copied = 0
    for name in names:
        with (src / name).open("rb") as f:
            body = f.read(MAX_SECRET_FILE + 1)
        if len(body) > MAX_SECRET_FILE:
            raise ValueError(f"{name} is larger than {MAX_SECRET_FILE} bytes")
        tmp = dst / f"..{name}.tmp"
        tmp.unlink(missing_ok=True)  # left by an attempt killed mid-copy; 0400, so it cannot be reopened
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o400)
        try:
            os.fchmod(fd, 0o400)  # the umask must not widen or narrow it
            view = memoryview(body)
            while view:
                n = os.write(fd, view)  # may write less, e.g. on a tmpfs near its size limit
                if n <= 0:
                    raise OSError("the key could not be written to the memory volume")
                view = view[n:]
            os.fsync(fd)
        finally:
            os.close(fd)
        os.replace(tmp, dst / name)
        copied += 1
    return copied


def main(argv: list[str] | None = None) -> None:
    args = sys.argv[1:] if argv is None else argv
    if len(args) != 2:
        emit("stage_usage", "usage: python -m m365_mcp.keys.stage <secret dir> <memory dir>", level=logging.ERROR)
        raise SystemExit(2)
    src, dst = Path(args[0]), Path(args[1])
    try:
        n = copy(src, dst)
    except (OSError, ValueError) as e:
        for left in dst.glob("..*.tmp"):
            left.unlink(missing_ok=True)
        # The exception's type and the file name only; a message from the file system never carries a value.
        emit(
            "stage_failed",
            f"staging the secrets failed: {e.__class__.__name__}",
            level=logging.ERROR,
            reason=str(e) if isinstance(e, ValueError) else e.__class__.__name__,
        )
        raise SystemExit(1) from None
    if n == 0:
        # A Secret with no keys would start a server that cannot sign a link; stop the pod here instead.
        emit("stage_empty", f"no secret files in {src}: the Secret is empty or not mounted", level=logging.ERROR)
        raise SystemExit(2)
    names = sorted(p for p in os.listdir(dst) if not p.startswith(".."))
    emit("staged", f"staged {n} files into {dst}", files=n, names=names)


if __name__ == "__main__":
    main()
