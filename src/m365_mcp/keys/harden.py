"""Keep the sealing key inside this process.

On Linux the process marks itself not dumpable before it reads any setting. Then another process of the
same user (a shell opened with `kubectl exec` runs as that user) cannot read this process's memory,
environment or open files through /proc, cannot attach a debugger, and no core dump is written. Root in
the pod still can; cluster access control decides who gets that.
"""

from __future__ import annotations

import ctypes
import sys

PR_GET_DUMPABLE = 3
PR_SET_DUMPABLE = 4


def non_dumpable() -> bool:
    """True when the process is now not dumpable; False where the system has no such switch."""
    if not sys.platform.startswith("linux"):
        return False
    libc = ctypes.CDLL(None, use_errno=True)
    if libc.prctl(PR_SET_DUMPABLE, 0, 0, 0, 0) != 0 or libc.prctl(PR_GET_DUMPABLE, 0, 0, 0, 0) != 0:
        raise OSError(ctypes.get_errno(), "could not mark the process not dumpable")
    return True
