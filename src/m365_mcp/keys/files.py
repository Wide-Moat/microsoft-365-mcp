"""Reading a secret from a file, and only from a file the server's own user alone may read.

The chart's init container copies each key of the Secret into a memory volume as a regular file, mode
0400, owned by the server's user; the server container never mounts the Secret itself. Anything else
is refused before a byte is read: a symbolic link (a Secret mounted straight into the container is one),
a directory or a device, a file a group or anyone else may read, a file owned by another user, and a
file larger than any real key, so a wrong mount cannot exhaust memory.
"""

from __future__ import annotations

import base64
import os
import stat
from pathlib import Path

MAX_SECRET_FILE = 64 * 1024
KEY_BYTES = 32


class SecretFileError(ValueError):
    """Why a secret file was refused. The message names the setting, never a value."""


def read_secret_file(path: str | Path, setting: str) -> bytes:
    try:
        fd = os.open(path, os.O_RDONLY | os.O_NOFOLLOW | os.O_NONBLOCK)
    except OSError as e:
        raise SecretFileError(
            f"{setting}: the file cannot be opened ({e.__class__.__name__}); it must be a regular file, not a link"
        ) from None
    try:
        st = os.fstat(fd)
        if not stat.S_ISREG(st.st_mode):
            raise SecretFileError(f"{setting}: not a regular file")
        if stat.S_IMODE(st.st_mode) & 0o077:
            raise SecretFileError(
                f"{setting}: the file is readable by a group or by others; it must be 0400 "
                "(the chart's init container stages it so)"
            )
        if st.st_uid != os.geteuid():
            raise SecretFileError(f"{setting}: the file belongs to another user; it must belong to the server's user")
        chunks: list[bytes] = []
        size = 0
        while size <= MAX_SECRET_FILE:
            chunk = os.read(fd, MAX_SECRET_FILE + 1 - size)
            if not chunk:
                break
            chunks.append(chunk)
            size += len(chunk)
        if size > MAX_SECRET_FILE:
            raise SecretFileError(f"{setting}: larger than {MAX_SECRET_FILE} bytes; it is not a secret file")
        return b"".join(chunks)
    finally:
        os.close(fd)


def read_password_file(path: str | Path, setting: str) -> str:
    """A password: the file's text without its final newline. An empty file is refused."""
    try:
        text = read_secret_file(path, setting).decode("utf-8")
    except UnicodeDecodeError:
        raise SecretFileError(f"{setting}: the file is not UTF-8 text") from None
    text = text[:-1] if text.endswith("\n") else text
    if not text or "\n" in text:
        raise SecretFileError(f"{setting}: the file must hold one password on one line")
    return text


def _unb64(s: str) -> bytes:
    raw = base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))
    if base64.urlsafe_b64encode(raw).rstrip(b"=").decode() != s:  # one spelling per key
        raise ValueError("non-canonical base64")
    return raw


def read_key_file(path: str | Path, setting: str) -> list[tuple[str, bytes]]:
    """One key per line, `<key id>:<base64url of 32 bytes>`; the first line seals, every line opens, so a
    rotation adds a line on top and later removes the old one. Blank lines are ignored."""
    try:
        lines = read_secret_file(path, setting).decode("ascii").splitlines()
    except UnicodeDecodeError:
        raise SecretFileError(f"{setting}: the file is not ASCII text") from None
    keys: list[tuple[str, bytes]] = []
    for n, line in enumerate(lines, 1):
        line = line.strip()
        if not line:
            continue
        kid, sep, material = line.partition(":")
        try:
            key = _unb64(material) if sep else b""
        except ValueError:
            key = b""
        if not kid or len(kid) > 32 or len(key) != KEY_BYTES:
            raise SecretFileError(f"{setting}: line {n} is not <key id>:<base64url of 32 bytes>")
        if any(kid == k for k, _ in keys):
            raise SecretFileError(f"{setting}: key id {kid} appears twice; a rotation adds a new id on top")
        keys.append((kid, key))
    if not keys:
        raise SecretFileError(f"{setting}: the file holds no key")
    return keys
