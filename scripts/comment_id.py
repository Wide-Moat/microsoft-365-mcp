"""Print the id of this PR's budget comment, or nothing, from a GitHub comments listing on stdin."""

import json
import sys

marker = "<!-- m365-budget -->"
found = [c for c in json.load(sys.stdin) if c.get("body", "").startswith(marker)]
print(found[0]["id"] if found else "")
