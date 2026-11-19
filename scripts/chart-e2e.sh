#!/usr/bin/env bash
# The chart in a real cluster (kind in CI): three replicas sharing Valkey behind Sentinel, the sealing key
# delivered the way the chart promises, JSON logs, metrics, autoscaling. No Microsoft 365 sign-in: a link is
# issued by the product's own code inside a pod (scripts/e2e-probe.py) for a token Graph will reject, so a
# replica that answers with Graph's refusal (404 "reconnect", or 502 when Graph cannot be reached) found,
# opened and verified the link, and one that does not know it answers 403.
#
#   IMAGE=microsoft-365-mcp:e2e scripts/chart-e2e.sh            # the image must already be in the cluster
#   IMAGE=… KNOWN_BAD=1 scripts/chart-e2e.sh                    # an image from before 0.2.0: every check must fail
#
# Every check prints PASS or FAIL and the run goes on, so one run shows them all; the script ends with the
# count and fails unless every check passed (or, with KNOWN_BAD=1, unless every check failed).
set -uo pipefail
cd "$(dirname "$0")/.."
image=${IMAGE:?IMAGE=<repository:tag> already loaded into the cluster}
bad=${KNOWN_BAD:-0}
ns=m365-e2e
rel=e2e
full=$rel-microsoft-365-mcp
secret=m365-key
results=()
pass() { echo "PASS $*"; results+=("PASS|$*"); }
fail() { echo "FAIL $*"; results+=("FAIL|$*"); }
na() { echo "N/A  $*"; }
die() { echo "SETUP FAILED: $*"; exit 2; }
k() { kubectl --request-timeout=30s -n "$ns" "$@"; }

# -- setup ------------------------------------------------------------------------------------------
# Only ever a throwaway kind cluster: the run deletes and recreates its namespace and kills pods.
context=$(kubectl config current-context 2>/dev/null)
case "$context" in kind-*) ;; *) die "the current context '$context' is not a kind cluster; this script only runs in one";; esac
# A namespace of its own, fresh every run: a store left from an earlier run keeps its old passwords. An
# existing one is deleted only when an earlier run of this script made it (its label says so).
if kubectl --request-timeout=30s get namespace "$ns" >/dev/null 2>&1; then
  owner=$(kubectl --request-timeout=30s get namespace "$ns" -o jsonpath='{.metadata.labels.m365-e2e/owner}')
  [ "$owner" = chart-e2e ] || die "namespace $ns exists and this script did not make it; not touching it"
  timeout 300 kubectl delete namespace "$ns" --wait=true >/dev/null || die "the old namespace did not go"
fi
printf '{"apiVersion":"v1","kind":"Namespace","metadata":{"name":"%s","labels":{"m365-e2e/owner":"chart-e2e"}}}' "$ns" \
  | kubectl --request-timeout=30s create -f - >/dev/null || die "namespace"
pw=$(openssl rand -hex 16) spw=$(openssl rand -hex 16)
k create secret generic valkey-auth --from-literal=password="$pw" --from-literal=sentinelPassword="$spw" \
  >/dev/null || die "store Secret"
k create secret generic m365-state --from-literal=password="$pw" --from-literal=sentinelPassword="$spw" \
  >/dev/null || die "state Secret"
k apply -f tests/e2e/valkey-sentinel.yaml >/dev/null || die "the store manifest"
k wait --for=condition=Ready pod/valkey-a pod/valkey-b --timeout=180s >/dev/null || die "valkey did not start"
k rollout status statefulset/valkey-sentinel --timeout=180s >/dev/null || die "the sentinels did not start"
sentinel() { k exec valkey-sentinel-0 -- sh -c 'valkey-cli -p 26379 -a "$SPW" --no-auth-warning '"$*"; }
for _ in $(seq 1 60); do
  others=$(sentinel sentinel master mymaster | awk 'p{print; exit} /^num-other-sentinels$/{p=1}')
  replicas=$(sentinel sentinel master mymaster | awk 'p{print; exit} /^num-slaves$/{p=1}')
  [ "${others:-0}" = 2 ] && [ "${replicas:-0}" = 1 ] && break
  sleep 2
done
[ "${others:-0}" = 2 ] && [ "${replicas:-0}" = 1 ] || die "Sentinel does not see its peers and the replica"
master_before=$(sentinel sentinel get-master-addr-by-name mymaster | head -1)
echo "store: master $master_before, 1 replica, 3 sentinels"

kid1=$(scripts/new-sealing-key.sh "$ns" "$secret") || die "no key"
[ -n "$kid1" ] && [ "${kid1#*_}" = "$kid1" ] || die "the first key is not one key: '$kid1'"
install() {
  timeout 300 helm upgrade --install "$rel" charts/microsoft-365-mcp -n "$ns" --timeout 240s \
    --set image.repository="${image%:*}" --set image.tag="${image##*:}" --set image.pullPolicy=Never \
    --set replicaCount=3 --set "state.sentinel.hosts={valkey-sentinel:26379}" --set state.existingSecret=m365-state \
    --set key.existingSecret="$secret" --set linksBaseUrl="http://$full.$ns.svc:8001" "$@" >/dev/null
}
serving() {  # the server pods that run and are not on their way out, oldest first
  k get pod -l app.kubernetes.io/instance="$rel" -o json | jq -r '[.items[] | select(.status.phase == "Running"
    and .metadata.deletionTimestamp == null)] | sort_by(.metadata.creationTimestamp) | .[].metadata.name'
}
patch_old() {
  k patch deployment "$full" --type json -p '[
    {"op":"remove","path":"/spec/template/spec/initContainers"},
    {"op":"add","path":"/spec/template/spec/containers/0/env/-","value":{"name":"M365_LINK_SECRET","value":"known-bad-run-0123456789abcdefghijklmnopqrstuvwxyz"}},
    {"op":"replace","path":"/spec/template/spec/containers/0/readinessProbe/httpGet/port","value":"http"},
    {"op":"replace","path":"/spec/template/spec/containers/0/livenessProbe/httpGet/port","value":"http"}]' >/dev/null
  k rollout status deploy/"$full" --timeout=240s >/dev/null
}
if [ "$bad" = 1 ]; then
  # 0.1.x has no init container step, no key file, no metrics port: run it the way it ran, with the link
  # secret in the environment and its probes on port 8000, and the rest of the chart as it is.
  install --set key.revision="$kid1" || die "helm install"
  patch_old || die "the old image did not roll out"
else
  install --wait --set key.revision="$kid1" || die "helm install"
fi
mapfile -t pods < <(serving)
[ "${#pods[@]}" = 3 ] || die "${#pods[@]} running pods, not 3"
ip() { k get pod "$1" -o jsonpath='{.status.podIP}'; }
A=${pods[0]} B=${pods[1]} C=${pods[2]}
IPA=$(ip "$A") IPB=$(ip "$B") IPC=$(ip "$C")
echo "pods: A=$A ($IPA) B=$B ($IPB) C=$C ($IPC)"
probe() { local p=$1; shift; timeout 90 kubectl -n "$ns" exec -i "$p" -c server -- python - "$@" <scripts/e2e-probe.py 2>&1 | tail -1; }
resolved() {  # a link one replica found and verified: Graph's refusal of the forged token, never 403
  case "$1" in 404*Reconnect*|502*reached*) return 0;; *) return 1;; esac
}
restarts() { k get pod "$1" -o jsonpath='{.status.containerStatuses[0].restartCount}'; }
metrics() {  # the metrics page of the pod at $1, fetched from inside the cluster
  timeout 30 kubectl -n "$ns" exec "$C" -c server -- python -c \
    "import urllib.request;print(urllib.request.urlopen('http://$1:9090/metrics',timeout=5).read().decode())" 2>&1
}

na "1 an authorization code redeemed on one pod: this server issues none (the gateway signs people in)"
na "2 a shared sign-in attempt limit: this server has no sign-in"

# -- 3. the send allowance is one per person across replicas ------------------------------------------
a=$(probe "$A" guard e2e-guard 10)
b=$(probe "$B" guard e2e-guard 1)
if [ "$a" = "$(printf 'ok %.0s' {1..10} | sed 's/ $//')" ] && [ "$b" = rate ]; then
  pass "3 ten sends drawn on pod A, the eleventh refused on pod B"
else
  fail "3 send allowance: pod A '$a', pod B '$b' (want ten ok, then rate)"
fi

# -- 4. a short link made on pod A works through pod B ----------------------------------------------
id=$(probe "$A" link)
got=$(probe "$C" get "http://$IPB:8001/d/$id")
control=$(probe "$C" get "http://$IPB:8001/d/AAAAAAAAAAAAAAAAAAAAAA")
if resolved "$got" && [ "${control%% *}" = 403 ]; then
  pass "4 a link made on pod A downloads through pod B (Graph's answer: ${got%% *}); an unknown id is 403"
else
  fail "4 link made on pod A, through pod B: '$got'; unknown id: '$control'"
fi
up=$(probe "$A" upload-link)
claimed=$(probe "$A" claim "$up")
put=$(probe "$C" put "http://$IPB:8001/u/$up")
if [ "$claimed" = True ] && [ "${put%% *}" = 409 ]; then
  pass "4b an upload running on pod A holds its link on pod B (409)"
else
  fail "4b upload claimed on pod A ('$claimed'), PUT through pod B: '$put' (want 409)"
fi

# -- key delivery -----------------------------------------------------------------------------------
spec=$(k get pod "$A" -o json)
if echo "$spec" | jq -e '[.spec.containers[] | select(.name=="server") | .volumeMounts[].name] | index("secret") == null' >/dev/null \
  && echo "$spec" | jq -e '[.spec.initContainers[]? | select(.name=="stage-keys") | .volumeMounts[].name] | index("secret") != null' >/dev/null; then
  pass "k1 the Secrets are mounted into the init container only"
else
  fail "k1 the Secrets are not mounted into the init container alone"
fi
material=$(k get secret "$secret" -o json | jq -r '.data.key' | base64 -d | head -1 | cut -d: -f2)
envs=$(timeout 60 kubectl -n "$ns" exec "$A" -c server -- env 2>&1)
if [ -n "$envs" ] && [ "${#material}" = 43 ] && ! grep -qF "$material" <<<"$envs" && ! grep -q "M365_LINK_SECRET\|$pw" <<<"$envs"; then
  pass "k2 no key and no store password in the environment a shell gets"
else
  fail "k2 a key or a password is in the server container's environment (or the probe saw nothing)"
fi
mode=$(timeout 60 kubectl -n "$ns" exec "$A" -c server -- stat -c '%a %u' /run/microsoft-365-mcp/key /run/microsoft-365-mcp/password /run/microsoft-365-mcp/sentinelPassword 2>&1 | tr '\n' ' ')
if [ "$mode" = "400 10001 400 10001 400 10001 " ]; then
  pass "k3 the key and both passwords are staged 0400, owned by the server's user"
else
  fail "k3 staged files: '$mode'"
fi
got=$(timeout 60 kubectl -n "$ns" exec -i "$A" -c server -- python - <scripts/proc-probe.py 2>&1)
if [ "$got" = "READ DENIED DENIED" ]; then pass "k4 a shell of the server's user cannot read its /proc"
else fail "k4 own environ, PID 1 environ, PID 1 maps: $got"; fi

# -- 5. the master dies: fail closed, then Sentinel fails over --------------------------------------
before=$(for p in "${pods[@]}"; do restarts "$p"; done | tr '\n' ' ')
t0=$(date +%s)
k delete pod valkey-a --wait=false >/dev/null
during=$(probe "$A" guard e2e-outage 1)
ready_during=$(for p in "${pods[@]}"; do k get pod "$p" -o jsonpath='{.status.conditions[?(@.type=="Ready")].status}'; done)
readyz=$(probe "$C" get "http://$IPA:9090/readyz")
down_seen=no
for _ in $(seq 1 10); do
  m=$(metrics "$IPB" | grep '^mcp_state_store_up')
  case "$m" in *" 0.0") down_seen=yes; break;; esac
  sleep 1
done
back=""
while [ $(( $(date +%s) - t0 )) -le 30 ]; do
  back=$(probe "$B" guard e2e-after 1)
  [ "$back" = ok ] && break
  sleep 1
done
took=$(( $(date +%s) - t0 ))
master_after=$(sentinel sentinel get-master-addr-by-name mymaster | head -1)
after=$(for p in "${pods[@]}"; do restarts "$p"; done | tr '\n' ' ')
if [ "$during" = state_unavailable ]; then pass "5a with the master gone a send fails closed (state_unavailable)"
else fail "5a with the master gone a send answered '$during' (want state_unavailable)"; fi
if [ "$ready_during" = "TrueTrueTrue" ] && [ "${readyz%% *}" = 200 ] && [ "$before" = "$after" ]; then
  pass "5b readiness stays 200 through the outage, no pod restarted"
else
  fail "5b ready '$ready_during', readyz '${readyz:0:60}', restarts before '$before' after '$after'"
fi
if [ "$down_seen" = yes ]; then pass "5c mcp_state_store_up went to 0 during the outage"
else fail "5c mcp_state_store_up never showed 0 during the outage"; fi
# Recovery means something: the same sends failed closed a moment before (5a).
if [ "$during" = state_unavailable ] && [ "$back" = ok ] && [ "$took" -le 30 ] && [ "$master_after" != "$master_before" ]; then
  pass "5d Sentinel promoted $master_after; sends work again ${took}s after the kill"
else
  fail "5d after ${took}s: send '$back', master '$master_after' (was '$master_before')"
fi
id=$(probe "$B" link)
got=$(probe "$A" get "http://$IPC:8001/d/$id")
if resolved "$got"; then pass "5e after the failover a link made on pod B downloads through pod C"
else fail "5e after the failover, link from B through C: '$got'"; fi

# -- 7. metrics on their own port, from every pod -----------------------------------------------------
sleep 6  # the store watch pings every 5 s
ok7=1
for p in "${pods[@]}"; do
  pip=$(ip "$p")
  m=$(metrics "$pip")
  series=$(grep -vc '^#' <<<"$m")
  main=$(probe "$C" get "http://$pip:8000/metrics")
  if grep -q '^mcp_http_requests_in_flight{' <<<"$m" && grep -qx 'mcp_state_store_up{server="microsoft-365-mcp"} 1.0' <<<"$m" \
    && grep -q '^mcp_build_info{server="microsoft-365-mcp",version=' <<<"$m" && [ "${main%% *}" = 404 ] && [ "$series" -lt 5000 ]; then
    echo "  $p: $series series, store up, MCP port /metrics ${main%% *}"
  else
    ok7=0; echo "  $p: series $series, MCP port /metrics '${main:0:40}', store '$(grep '^mcp_state_store_up' <<<"$m")'"
  fi
done
if [ "$ok7" = 1 ]; then pass "7 every pod serves its metrics on 9090 only, store up, under 5000 series"
else fail "7 metrics"; fi

# -- key lifetime: a container restart keeps it, a rotation keeps the previous one for the rollout ------
id=$(probe "$A" link)
r0=$(restarts "$A")
timeout 30 kubectl -n "$ns" exec "$A" -c server -- sh -c 'kill 1' >/dev/null 2>&1
for _ in $(seq 1 60); do
  [ "$(restarts "$A")" -gt "$r0" ] && [ "$(k get pod "$A" -o jsonpath='{.status.containerStatuses[0].ready}')" = true ] && break
  sleep 2
done
got=$(probe "$C" get "http://$IPA:8001/d/$id")
if [ "$(restarts "$A")" -gt "$r0" ] && resolved "$got"; then pass "k5 a container restart keeps the key: a link from before opens after"
else fail "k5 after a container restart: restarts $(restarts "$A") (was $r0), link '$got'"; fi
# A rotation in three deploys (scripts/new-sealing-key.sh): add the new key (opens only), promote it
# (issues; the old one still opens), drop the old one. Every pod serving during a rollout must open what
# any other pod issues, so the middle rollout is paused with both kinds of pod running and checked both ways.
redeploy() {  # $1 revision; $2 "wait" or "nowait"
  if [ "$bad" = 1 ]; then
    timeout 300 helm upgrade "$rel" charts/microsoft-365-mcp -n "$ns" --reuse-values --set key.revision="$1" \
      >/dev/null || die "helm upgrade"
    patch_old
  elif [ "$2" = wait ]; then
    install --wait --set key.revision="$1" || die "helm upgrade to $1"
  else
    install --set key.revision="$1" || die "helm upgrade to $1"
  fi
}
revision_of() { k get pod "$1" -o jsonpath='{.metadata.annotations.microsoft-365-mcp/key-revision}'; }
before_link=$(probe "$A" link)
rev2=$(scripts/new-sealing-key.sh "$ns" "$secret") || die "rotation phase 1"
redeploy "$rev2" wait
k rollout status deploy/"$full" --timeout=240s >/dev/null || die "rollout after phase 1"
mapfile -t pods < <(serving)
A=${pods[0]} B=${pods[1]}; IPB=$(ip "$B")
got=$(probe "$A" get "http://$IPB:8001/d/$before_link")
if [ "$(tr '_' '\n' <<<"$rev2" | wc -l | tr -d ' ')" = 2 ] && [ "${rev2%%_*}" = "$kid1" ] && resolved "$got"; then
  pass "k6 phase 1 adds the new key behind the current one: links from before still open"
else
  fail "k6 phase 1: revision '$rev2' (key 1 '$kid1'), link from before '$got'"
fi
rev3=$(scripts/new-sealing-key.sh --promote "$ns" "$secret") || die "rotation phase 2"
# A new pod becomes available only after 40 s Ready, so old pods keep serving beside it: a window wide
# enough to pause in, whatever the speed of the cluster. Helm leaves a field the chart never sets.
[ "$bad" = 1 ] || k patch deploy "$full" --type merge -p '{"spec":{"minReadySeconds":40}}' >/dev/null \
  || die "minReadySeconds"
redeploy "$rev3" nowait
mixed=""
if [ "$bad" = 0 ]; then
  for _ in $(seq 1 90); do
    mapfile -t now < <(serving)
    promoted="" waiting=""
    for p in "${now[@]}"; do
      ready=$(k get pod "$p" -o jsonpath='{.status.conditions[?(@.type=="Ready")].status}')
      [ "$ready" = True ] || continue
      case "$(revision_of "$p")" in "$rev3") promoted=${promoted:-$p};; "$rev2") waiting=${waiting:-$p};; esac
    done
    if [ -n "$promoted" ] && [ -n "$waiting" ]; then
      k rollout pause deploy/"$full" >/dev/null || die "rollout pause"
      break
    fi
    sleep 1
  done
  # Both still serving after the pause: neither on its way out.
  for p in $promoted $waiting; do
    [ "$(k get pod "$p" -o jsonpath='{.metadata.deletionTimestamp}')" = "" ] || { promoted=""; waiting=""; }
  done
  if [ -n "$promoted" ] && [ -n "$waiting" ]; then
    new_link=$(probe "$promoted" link)
    a=$(probe "$promoted" get "http://$(ip "$waiting"):8001/d/$new_link")
    w_link=$(probe "$waiting" link)
    b=$(probe "$waiting" get "http://$(ip "$promoted"):8001/d/$w_link")
    resolved "$a" && resolved "$b" && mixed=yes
    echo "  promoted $promoted issued, $waiting opened: ${a%% *}; $waiting issued, $promoted opened: ${b%% *}"
  fi
  k rollout resume deploy/"$full" >/dev/null 2>&1
  k patch deploy "$full" --type json -p '[{"op":"remove","path":"/spec/minReadySeconds"}]' >/dev/null 2>&1
fi
k rollout status deploy/"$full" --timeout=240s >/dev/null || die "rollout after phase 2"
mapfile -t pods < <(serving)
A=${pods[0]} B=${pods[1]}; IPB=$(ip "$B")
promoted_link=$(probe "$A" link)
got_before=$(probe "$A" get "http://$IPB:8001/d/$before_link")
got_new=$(probe "$A" get "http://$IPB:8001/d/$promoted_link")
if [ "$mixed" = yes ] && resolved "$got_before" && resolved "$got_new"; then
  pass "k7 phase 2 promotes the new key: mid-rollout old and new pods open each other's links, and the oldest still open"
else
  fail "k7 phase 2: mixed pods '${mixed:-not both}', link from before '$got_before', new link '$got_new'"
fi
rev4=$(scripts/new-sealing-key.sh --drop-previous "$ns" "$secret") || die "rotation phase 3"
redeploy "$rev4" wait
k rollout status deploy/"$full" --timeout=240s >/dev/null || die "rollout after phase 3"
mapfile -t pods < <(serving)
A=${pods[0]} B=${pods[1]}; IPB=$(ip "$B")
got=$(probe "$A" get "http://$IPB:8001/d/$before_link")
fresh_got=$(probe "$A" get "http://$IPB:8001/d/$promoted_link")
if [ "${got%% *}" = 403 ] && resolved "$fresh_got"; then
  pass "k8 phase 3 drops the old key: what it sealed is refused (403), the new key's links open"
else
  fail "k8 phase 3: link under the dropped key '$got', under the new key '$fresh_got'"
fi

# -- 6. every line of every container, the init container too, is JSON ------------------------------
total=0 bad_lines=0 init_lines=0
for p in "${pods[@]}"; do
  for c in stage-keys server; do
    out=$(k logs "$p" -c "$c" 2>/dev/null)
    prev=$(k logs "$p" -c "$c" --previous 2>/dev/null)
    n=$(printf '%s\n%s\n' "$out" "$prev" | python3 -c '
import json, sys
good = bad = 0
for line in sys.stdin:
    if not line.strip():
        continue
    try:
        rec = json.loads(line)
        ok = isinstance(rec, dict) and {"ts", "level", "logger", "event", "msg"} <= set(rec)
    except ValueError:
        ok = False
    if ok:
        good += 1
    else:
        bad += 1
        print("NOT JSON:", line.rstrip()[:200], file=sys.stderr)
print(good, bad)')
    total=$((total + ${n% *})); bad_lines=$((bad_lines + ${n#* }))
    [ "$c" = stage-keys ] && init_lines=$((init_lines + ${n% *}))
  done
done
if [ "$bad_lines" = 0 ] && [ "$total" -gt 0 ] && [ "$init_lines" -ge "${#pods[@]}" ]; then
  pass "6 all $total log lines of ${#pods[@]} pods are JSON, $init_lines of them from the init container"
else
  fail "6 logs: $total JSON lines, $bad_lines not JSON, $init_lines from the init container"
fi

# -- 8. the autoscaler renders and reads a CPU target -------------------------------------------------
timeout 300 helm upgrade "$rel" charts/microsoft-365-mcp -n "$ns" --reuse-values --set autoscaling.enabled=true \
  >/dev/null || die "helm upgrade with the autoscaler"
[ "$bad" = 1 ] && patch_old
cpu=""
for _ in $(seq 1 60); do
  cpu=$(k get hpa "$full" -o jsonpath='{.status.currentMetrics[?(@.type=="Resource")].resource.current.averageUtilization}' 2>/dev/null)
  [ -n "$cpu" ] && break
  sleep 5
done
if [ -n "$cpu" ]; then pass "8 the HorizontalPodAutoscaler reads CPU (${cpu}% of requests)"
else fail "8 the HorizontalPodAutoscaler shows no CPU target: $(k get hpa 2>&1 | tail -1)"; fi

# -- the count ------------------------------------------------------------------------------------
passed=0 failed=0
for r in "${results[@]}"; do case "$r" in PASS*) passed=$((passed + 1));; FAIL*) failed=$((failed + 1));; esac; done
want=19
echo
echo "chart e2e: $passed passed, $failed failed, of $want checks (image $image)"
[ $((passed + failed)) -eq "$want" ] || { echo "$((passed + failed)) of $want checks ran"; exit 1; }
if [ "$bad" = 1 ]; then
  # Two checks hold on the old image too, and are expected to: its /readyz answered 200 whatever happened
  # (5b), and the autoscaler is the chart's alone (8). Every other check must fail there.
  unexpected=$(for r in "${results[@]}"; do case "$r" in "PASS|5b "*|"PASS|8 "*) ;; PASS*) echo "${r#PASS|}";; esac; done)
  [ -z "$unexpected" ] || { echo "KNOWN_BAD: checks passed on the old image: $unexpected"; exit 1; }
  echo "KNOWN_BAD: every check but 5b and 8 failed on the old image, as it must"
  exit 0
fi
[ "$failed" -eq 0 ]
