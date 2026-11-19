"""Read the container logs chart-e2e.sh kept (one file per pod and container) and count their lines.

Prints `<JSON lines> <other lines> <JSON lines from the init container> <files>`; every line that is not
a JSON object with ts, level, logger, event and msg is printed to stderr.
"""

import json
import pathlib
import sys

good = bad = init = files = 0
for f in sorted(pathlib.Path(sys.argv[1]).iterdir()):
    files += 1
    for line in f.read_text(errors="replace").splitlines():
        if not line.strip():
            continue
        try:
            rec = json.loads(line)
            ok = isinstance(rec, dict) and {"ts", "level", "logger", "event", "msg"} <= set(rec)
        except ValueError:
            ok = False
        if ok:
            good += 1
            init += ".stage-keys" in f.name
        else:
            bad += 1
            print(f"NOT JSON in {f.name}: {line[:200]}", file=sys.stderr)
print(good, bad, init, files)
