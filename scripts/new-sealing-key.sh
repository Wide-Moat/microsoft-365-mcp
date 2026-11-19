#!/usr/bin/env bash
# The sealing key of microsoft-365-mcp, in the Secret the chart's init container stages.
#
#   rev=$(scripts/new-sealing-key.sh <namespace> <secret> [field])                  # first key, or phase 1
#   rev=$(scripts/new-sealing-key.sh --promote <namespace> <secret> [field])        # phase 2
#   rev=$(scripts/new-sealing-key.sh --drop-previous <namespace> <secret> [field])  # phase 3
#   helm upgrade … --set key.existingSecret=<secret> --set key.revision="$rev"      # after each one
#
# The key file holds one key per line: the first issues, every line opens. A rotation goes in three
# deploys, so every pod serving during a rollout opens everything any other pod issues:
#   1. a new key is added as the second line: it opens, it does not issue yet;
#   2. --promote swaps the two lines: the new key issues, the old one still opens;
#   3. --drop-previous, once the old key's links have expired (minutes): only the new key is left.
# After a leak, run 1, 2 and 3 at once, accepting that links issued in between are refused.
#
# --revision changes nothing and prints the revision of what the Secret holds now: after a run whose
# output was lost, it says what to deploy.
#
# Run against a namespace without the Secret, the first form creates it with one key. It prints the
# revision for key.revision: the key ids in order, never a key. The key goes from openssl to a 0600 file
# in a private directory and kubectl reads that file; nothing secret is in an argument or the output.
# Other fields of the Secret are kept.
set -euo pipefail
mode=add
case "${1:-}" in
  --promote) mode=promote; shift;;
  --drop-previous) mode=drop; shift;;
  --revision) mode=revision; shift;;
esac
usage="usage: new-sealing-key.sh [--promote|--drop-previous|--revision] <namespace> <secret> [field]"
ns=${1:?$usage}
name=${2:?$usage}
field=${3:-key}

umask 077
work=$(mktemp -d)
cleanup() { rm -rf "$work"; }
trap cleanup EXIT
k() { kubectl --request-timeout=30s -n "$ns" "$@"; }

exists=0
if k get secret "$name" -o json >"$work/secret.json" 2>"$work/err"; then
  exists=1
  # A literal key lookup: a field such as `sealing.key` is not a JSONPath.
  jq -r --arg f "$field" '.data[$f] // empty' "$work/secret.json" >"$work/b64"
  # The version read: the write below is refused if anyone changed the Secret since (two rotations at once).
  version=$(jq -r '.metadata.resourceVersion' "$work/secret.json")
  rm -f "$work/secret.json"
  if [ -s "$work/b64" ]; then
    openssl base64 -d -A <"$work/b64" >"$work/current" || { echo "the Secret's $field is not base64" >&2; exit 2; }
  else
    : >"$work/current"
  fi
  # The key lines as the server reads them: blank lines are not keys.
  grep -v '^[[:space:]]*$' "$work/current" | sed 's/^[[:space:]]*//;s/[[:space:]]*$//' >"$work/lines" || true
elif grep -q NotFound "$work/err"; then
  : >"$work/lines"
else
  echo "the Secret could not be read: $(head -1 "$work/err")" >&2
  exit 2
fi
n=$(wc -l <"$work/lines" | tr -d ' ')

newest=$(cut -d: -f1 "$work/lines" | sort | tail -1)
new_key() {
  # Ids sort in the order the keys were made, always after every key the Secret holds: --promote relies
  # on it to tell a promoted pair from one waiting.
  if [[ -n "$newest" && ! "$newest" =~ ^k[0-9]{14}[0-9a-f]{4}$ ]]; then
    echo "the Secret holds key id '$newest', not one this script made: rotate it by hand" >&2
    exit 2
  fi
  kid="k$(date -u +%Y%m%d%H%M%S)$(openssl rand -hex 2)"
  for _ in 1 2 3; do
    [[ -z "$newest" || "$kid" > "$newest" ]] && break
    sleep 1
    kid="k$(date -u +%Y%m%d%H%M%S)$(openssl rand -hex 2)"
  done
  [[ -z "$newest" || "$kid" > "$newest" ]] || {
    echo "the newest key id $newest is from the future by this clock: check the clock" >&2
    exit 2
  }
  material=$(openssl rand 32 | openssl base64 -A | tr '+/' '-_' | tr -d '=')
  test "${#material}" -eq 43 || { echo "the generated key is not 32 bytes" >&2; exit 2; }
  printf '%s:%s\n' "$kid" "$material"
  unset material
}
case "$mode" in
  add)
    if [ "$n" -eq 0 ]; then
      new_key >"$work/key"
    else
      if [ "$n" -ne 1 ]; then
        rev=$(cut -d: -f1 "$work/lines" | paste -sd_ -)
        echo "the Secret holds $n keys (revision $rev): a rotation is in progress; deploy that revision, then --promote" >&2
        exit 2
      fi
      { cat "$work/lines"; new_key; } >"$work/key"
    fi;;
  promote)
    [ "$n" -eq 2 ] || { echo "--promote needs two keys (run phase 1 first); the Secret holds $n" >&2; exit 2; }
    first=$(sed -n 1p "$work/lines" | cut -d: -f1) second=$(sed -n 2p "$work/lines" | cut -d: -f1)
    # Key ids carry their creation time, so a second run (a lost reply, a retried job) changes nothing
    # instead of swapping the keys back.
    if [[ "$first" > "$second" ]]; then
      cp "$work/lines" "$work/key"
    else
      { sed -n 2p "$work/lines"; sed -n 1p "$work/lines"; } >"$work/key"
    fi;;
  drop)
    [ "$n" -ge 1 ] || { echo "the Secret holds no key" >&2; exit 2; }
    sed -n 1p "$work/lines" >"$work/key";;
  revision)
    [ "$n" -ge 1 ] || { echo "the Secret holds no key" >&2; exit 2; }
    cut -d: -f1 "$work/lines" | paste -sd_ -
    exit 0;;
esac
rm -f "$work/current" "$work/lines" "$work/b64"

value=$(openssl base64 -A <"$work/key")
if [ "$exists" = 1 ]; then
  printf '{"metadata":{"resourceVersion":"%s"},"data":{"%s":"%s"}}' "$version" "$field" "$value" >"$work/patch.json"
else
  printf '{"data":{"%s":"%s"}}' "$field" "$value" >"$work/patch.json"
fi
unset value
if [ "$exists" = 1 ]; then
  k patch secret "$name" --type merge --patch-file "$work/patch.json" >/dev/null || {
    echo "the Secret changed while this ran (or could not be written): nothing was changed; run it again" >&2
    exit 2
  }
else
  printf '{"apiVersion":"v1","kind":"Secret","type":"Opaque","metadata":{"name":"%s","namespace":"%s"},' \
    "$name" "$ns" >"$work/new.json"
  tail -c +2 "$work/patch.json" >>"$work/new.json"
  k create -f "$work/new.json" >/dev/null
fi
cut -d: -f1 "$work/key" | paste -sd_ -
