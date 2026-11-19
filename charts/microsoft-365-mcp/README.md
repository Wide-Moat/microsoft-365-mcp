# microsoft-365-mcp chart

```sh
helm upgrade --install microsoft-365-mcp oci://ghcr.io/wide-moat/charts/microsoft-365-mcp --version <x.y.z> \
  -n <namespace> -f values.yaml --set key.revision=<key id>
```

## What it deploys

| Object | Always | Notes |
|---|---|---|
| Deployment | yes | init container `stage-keys` (the only one that mounts a Secret) + `server` |
| Service | yes | ports `http` (80 → 8000, MCP, for the gateway), `links` (8001, for the sandbox), `metrics` (9090) |
| PodDisruptionBudget | `podDisruptionBudget.enabled` (on) | `maxUnavailable: 1` |
| HorizontalPodAutoscaler | `autoscaling.enabled` (off) | |
| ServiceMonitor | `metrics.serviceMonitor.enabled` (off) | for the Prometheus Operator |
| Ingress | `ingress.enabled` (off) | the links port only, paths `/d/` and `/u/` |

No NetworkPolicy: who reaches the MCP port is the cluster's network decision.

## Values

| Key | Default | Meaning |
|---|---|---|
| `image.repository`, `image.tag`, `image.digest` | `ghcr.io/wide-moat/microsoft-365-mcp`, appVersion, none | a digest wins over the tag |
| `imagePullSecrets` | `[]` | |
| `linksBaseUrl` | none | https address of the links port, or `http://<service>.<namespace>.svc[.cluster.local][:port]`; required while an area that moves files is on |
| `areas` | `[]` (defaults) | subset of mail, calendar, contacts, files, lists, excel, teams, recordings, tasks, notes, bookings, account |
| `switches` | `null` (defaults) | a list replaces them: teams_channels, transcripts, rooms, chat_admin |
| `allowedRecipientDomains` | `[]` | |
| `expectedClientIds` | `[]` | set it: your Entra app's client id |
| `logLevel` | `INFO` | |
| `caBundle.configMap`, `caBundle.key` | none, `ca.crt` | extra CAs for Graph and the store |
| `replicaCount` | 2 | more than 1 needs `state` |
| `autoscaling.enabled` | false | |
| `autoscaling.minReplicas`, `maxReplicas` | 2, 10 | |
| `autoscaling.targetCPUUtilizationPercentage` | 70 | |
| `autoscaling.inFlightPerPod` | null | also scale on `mcp_http_requests_in_flight` per pod; needs a custom metrics API (prometheus-adapter) |
| `podDisruptionBudget.enabled`, `maxUnavailable` | true, 1 | |
| `metrics.port` | 9090 | metrics and the probes |
| `metrics.serviceMonitor.enabled`, `interval`, `labels` | false, 30s, {} | |
| `state.url` | none | `redis://` or `rediss://`, no credentials in it |
| `state.sentinel.hosts`, `state.sentinel.master` | `[]`, mymaster | one host (a Service in front of the sentinels) is enough |
| `state.db`, `state.tls` | 0, false | Sentinel mode; with a URL the URL decides |
| `state.username` | none | ACL user |
| `state.existingSecret` | none | keys `password`, and `sentinelPassword` in Sentinel mode |
| `key.existingSecret` | none, required | the sealing key file |
| `key.secretKey` | `key` | its key in the Secret |
| `key.revision` | none | set to the key id after a rotation: the pods restart and read it |
| `service.type`, `service.port`, `service.linksPort` | ClusterIP, 80, 8001 | |
| `ingress.enabled`, `className`, `annotations`, `tlsSecretName` | false | host from `linksBaseUrl` |
| `tmp.medium`, `tmp.sizeLimit` | disk, 2Gi | the upload spool, up to 1 GiB for all uploads at once |
| `resources` | 50m / 192Mi request, 512Mi limit | a CPU limit is refused |
| `podAnnotations`, `podLabels`, `nodeSelector`, `tolerations`, `affinity` | | |
| `topologySpreadConstraints` | `[]` | empty: replicas spread across nodes, softly |

The chart refuses to render: more than one replica (or the autoscaler) without a store; a URL and
Sentinels together; credentials in `state.url`; a missing key Secret; a missing or plain-http
`linksBaseUrl` outside the cluster; a CPU limit; `metrics.port` 8000 or 8001.

## The key

`scripts/new-sealing-key.sh <namespace> <secret>` writes the key into the Secret and prints the
revision for `key.revision` (key ids, never a key). The key stays the same across deploys. A rotation
is three deploys, each with the revision the script prints: add the new key (it opens, the old one
still issues), `--promote` (the new one issues, the old one still opens), `--drop-previous` once the
old key's links have expired. Every pod serving during each rollout opens what any other issues
([state](../../docs/components/state.md#key-rotation)). The Secret is mounted into
the init container only, which copies it into a memory volume (0400, the server's user); the server
process marks itself non-dumpable, so a shell in the container cannot read its memory or environment.

## The store

Redis or Valkey, standalone or behind Sentinel, with `maxmemory-policy noeviction`. It holds ciphertext
and keyed hashes only ([state](../../docs/components/state.md)). If it goes down, reads keep working,
sends and new links fail closed with "try again in a few seconds", and readiness stays green.

## Publishing the links port

The links port moves files up to 250 MB (OneDrive) and 150 MB (mail attachments). With ingress-nginx,
set on `ingress.annotations`:

| Annotation | Value |
|---|---|
| `nginx.ingress.kubernetes.io/proxy-body-size` | `260m` or more |
| `nginx.ingress.kubernetes.io/proxy-request-buffering` | `"off"` |
| `nginx.ingress.kubernetes.io/proxy-buffering` | `"off"` |
| `nginx.ingress.kubernetes.io/proxy-read-timeout`, `proxy-send-timeout` | `"900"` |

The chart sets none of them by default.
