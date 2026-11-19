"""The container's entry point: JSON logging before anything else is imported, so even an import that
fails is one JSON line."""

from __future__ import annotations


def main() -> None:
    from m365_mcp import logjson

    logjson.install()
    from m365_mcp.app import run

    run()


if __name__ == "__main__":
    main()
