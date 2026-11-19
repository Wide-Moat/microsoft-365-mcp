#!/usr/bin/env bash
# The state store contract (tests/state) against real servers, never a fake:
#   Redis 7.4 and Valkey 8.1 standalone, with a password;
#   Valkey 8.1 master + replica behind three Sentinels, with passwords, the Sentinels with one of their own;
#   Redis 7.4 master + replica behind one Sentinel, no password anywhere (how a cluster operator often runs it).
#
#   scripts/state-contract.sh                       # the contract suite
#   scripts/state-contract.sh scripts/mutation_check.py --only state   # anything else, beside the servers
#
# Everything runs in one Docker network, the tests too, so the hostnames the Sentinels announce are the ones
# the client reaches, on Linux and on a laptop alike. The tests run as the caller's user on a copy of the
# tree; M365_TEST_STORES_REQUIRED=1 makes a missing server a failure rather than a skip.
set -euo pipefail
cd "$(dirname "$0")/.."
VALKEY=valkey/valkey:8.1.12-alpine@sha256:9971099af41c36e887aabdfe22b80c0fe45df949163dfc9cb434febc57cfc3fc
REDIS=redis:7.4.6-alpine@sha256:3b73847e72874be07e6657b129a94761662b79bc0f679273757d4218573b2a98
PYTHON=python:3.12-slim-bookworm@sha256:6a0e06f5ba0b373d051aabd3fffb9b873f27f9a2a3a14bcbe37dee6f0e3bbc70
net=m365-state-$$
pw=contract-store-pass
spw=contract-sentinel-pass
names=()

cleanup() {
  for n in "${names[@]}"; do docker rm -f "$n" >/dev/null 2>&1 || true; done
  docker network rm "$net" >/dev/null 2>&1 || true
}
trap cleanup EXIT

start() {  # name image args...
  local name=$1 image=$2
  shift 2
  names+=("$name")
  timeout 300 docker run -d --name "$name" --hostname "$name" --network "$net" "$image" "$@" >/dev/null
}

docker network create "$net" >/dev/null
start "$net-redis" "$REDIS" redis-server --requirepass "$pw" --save "" --appendonly no
start "$net-valkey" "$VALKEY" valkey-server --requirepass "$pw" --save "" --appendonly no
m="$net-vk-a" r="$net-vk-b"
start "$m" "$VALKEY" valkey-server --requirepass "$pw" --masterauth "$pw" --replica-announce-ip "$m" \
  --save "" --appendonly no
start "$r" "$VALKEY" valkey-server --requirepass "$pw" --masterauth "$pw" --replica-announce-ip "$r" \
  --replicaof "$m" 6379 --save "" --appendonly no
sentinels=()
for i in 1 2 3; do
  s="$net-sentinel-$i"
  # Sentinel rewrites its configuration file, so each one writes its own inside its container.
  start "$s" "$VALKEY" sh -c "cat > /tmp/s.conf <<CONF
port 26379
sentinel resolve-hostnames yes
sentinel announce-hostnames yes
sentinel announce-ip $s
requirepass $spw
sentinel sentinel-pass $spw
sentinel monitor mymaster $m 6379 2
sentinel auth-pass mymaster $pw
sentinel down-after-milliseconds mymaster 2000
sentinel failover-timeout mymaster 10000
sentinel parallel-syncs mymaster 1
CONF
exec valkey-sentinel /tmp/s.conf"
  sentinels+=("$s:26379")
done

rm_="$net-rd-a" rr="$net-rd-b" rs="$net-rd-sentinel"
start "$rm_" "$REDIS" redis-server --replica-announce-ip "$rm_" --save "" --appendonly no
start "$rr" "$REDIS" redis-server --replica-announce-ip "$rr" --replicaof "$rm_" 6379 --save "" --appendonly no
start "$rs" "$REDIS" sh -c "cat > /tmp/s.conf <<CONF
port 26379
sentinel resolve-hostnames yes
sentinel announce-hostnames yes
sentinel announce-ip $rs
sentinel monitor mymaster $rm_ 6379 1
sentinel down-after-milliseconds mymaster 2000
sentinel failover-timeout mymaster 10000
CONF
exec redis-sentinel /tmp/s.conf"

# Ready: the replica is in sync and every Sentinel knows the master and the other two.
for _ in $(seq 1 60); do
  # A server or sentinel not listening yet is "not ready", never a reason to stop: `|| true` in each probe.
  link=$(docker exec "$r" valkey-cli -a "$pw" --no-auth-warning info replication 2>/dev/null | tr -d '\r' \
    | sed -n 's/^master_link_status://p' || true)
  known=0
  for i in 1 2 3; do
    n=$(docker exec "$net-sentinel-$i" valkey-cli -p 26379 -a "$spw" --no-auth-warning \
      sentinel master mymaster 2>/dev/null | tr -d '\r' | awk 'p{print; exit} /^num-other-sentinels$/{p=1}' || true)
    [ "${n:-0}" = 2 ] && known=$((known + 1))
  done
  [ "$link" = up ] && [ "$known" = 3 ] && break
  sleep 1
done
[ "$link" = up ] && [ "$known" = 3 ] || { echo "the Sentinel set did not come up (link=$link, sentinels=$known)"; exit 1; }
for _ in $(seq 1 60); do
  rlink=$(docker exec "$rr" redis-cli info replication 2>/dev/null | tr -d '\r' \
    | sed -n 's/^master_link_status://p' || true)
  rmaster=$(docker exec "$rs" redis-cli -p 26379 sentinel get-master-addr-by-name mymaster 2>/dev/null \
    | head -1 | tr -d '\r' || true)
  [ "$rlink" = up ] && [ -n "$rmaster" ] && break
  sleep 1
done
[ "$rlink" = up ] && [ -n "$rmaster" ] || { echo "the Redis Sentinel did not come up (link=$rlink)"; exit 1; }
echo "stores up: redis, valkey, valkey $m + $r behind ${#sentinels[@]} sentinels, redis $rm_ + $rr behind one sentinel"

cmd=("$@")
[ "${#cmd[@]}" -gt 0 ] || cmd=(-m pytest -q -p no:cacheprovider tests/state)
names+=("$net-tests")
# The tree is mounted read-only and copied inside: a mutation run edits the copy, never the checkout.
timeout 1800 docker run --rm --name "$net-tests" --network "$net" --user "$(id -u):$(id -g)" \
  -v "$PWD:/repo:ro" -w /tmp -e HOME=/tmp -e PYTHONDONTWRITEBYTECODE=1 -e PIP_DISABLE_PIP_VERSION_CHECK=1 \
  -e M365_TEST_STORES_REQUIRED=1 -e M365_TEST_STATE_PASSWORD="$pw" -e M365_TEST_SENTINEL_PASSWORD="$spw" \
  -e M365_TEST_REDIS_URL="redis://$net-redis:6379/0" -e M365_TEST_VALKEY_URL="redis://$net-valkey:6379/0" \
  -e M365_TEST_SENTINELS="$(IFS=,; echo "${sentinels[*]}")" -e M365_TEST_SENTINEL_MASTER=mymaster \
  -e M365_TEST_REDIS_SENTINELS="$rs:26379" \
  "$PYTHON" sh -c 'mkdir -p /tmp/src && cd /repo && tar -cf - pyproject.toml requirements-test.lock src tests scripts \
    | tar -xf - -C /tmp/src && cd /tmp/src \
    && pip install -q --target /tmp/deps --require-hashes -r requirements-test.lock >/dev/null \
    && PYTHONPATH=/tmp/src/src:/tmp/src/tests:/tmp/deps exec python "$@"' sh "${cmd[@]}"
