# 3. Run the server

One image, three ports:

| Port | Path | Who reaches it |
|---|---|---|
| 8000 | `/mcp` | your MCP gateway only. Never publish it. |
| 8001 | `/d/…`, `/u/…` | the agent's sandbox, to download and upload files through signed links. Publish it on an https address, those two paths only. |
| 9090 | `/metrics`, `/healthz`, `/readyz` | Prometheus and the kubelet. Never publish it. |

## On Kubernetes: the chart

```sh
kubectl create namespace m365
scripts/new-sealing-key.sh m365 m365-key        # prints the revision (the key id)
helm upgrade --install microsoft-365-mcp oci://ghcr.io/wide-moat/charts/microsoft-365-mcp \
  --version <version> -n m365 -f values.yaml --set key.revision=<the key id>
```

The [chart README](../../charts/microsoft-365-mcp/README.md) lists every value. What you decide:

| Value | What it is |
|---|---|
| `key.existingSecret` | The Secret holding the sealing key. The script above writes one; any other way (a secrets manager, encrypted values) works if the Secret holds `<id>:<base64url of 32 bytes>` per line. The key does not change on a deploy. |
| `linksBaseUrl` | The https address of port 8001 (or a cluster-internal `http://<service>.<namespace>.svc:8001` when the sandbox runs in the cluster). |
| `state.url` or `state.sentinel.hosts` | The Redis or Valkey every replica shares. Required for more than one replica. Set `maxmemory-policy noeviction` on it. |
| `state.existingSecret` | `password` (and `sentinelPassword` with Sentinel), when the store has them. |
| `expectedClientIds` | Your Entra app's Application (client) ID. A token issued to any other app is refused. Set it. |
| `replicaCount` | 2 by default; one without a store. |

The chart runs non-root with a read-only root filesystem, no service-account token, a rolling update
that never lowers the number of serving pods, replicas spread across nodes, a PodDisruptionBudget, and
no CPU limit. It ships no NetworkPolicy: whether the gateway alone reaches port 8000 is your cluster's
network decision.

Egress the server needs: 443 to `graph.microsoft.com` and to the hosts Microsoft serves files from
(`*.sharepoint.com`, `*.onedrive.com`, `*.1drv.com`, `*.svc.ms`); the server follows file redirects
only to Microsoft's own domains. Behind a TLS-inspecting proxy, give its CA in `caBundle`.

## Without Kubernetes

```sh
mkdir -p keys && printf 'k1:%s\n' "$(openssl rand 32 | basenc --base64url | tr -d '=')" > keys/key
chmod 0400 keys/key && sudo chown 10001 keys/key
docker run -d --name m365-mcp -p 8000:8000 -p 8001:8001 -p 9090:9090 \
  -v "$PWD/keys:/run/keys:ro" -e M365_KEY_FILE=/run/keys/key \
  -e M365_LINKS_BASE_URL='https://m365-files.example.com' \
  ghcr.io/wide-moat/microsoft-365-mcp:<version>
```

The key file must be a regular file of the server's user (uid 10001), readable by nobody else.
Without `M365_STATE_URL` or `M365_STATE_SENTINELS` the shared state is this process's memory: run one.

## Settings

| Variable | Default | What it does |
|---|---|---|
| `M365_KEY_FILE` | none | The sealing key file. Required for file links. |
| `M365_LINKS_BASE_URL` | none | The public https address of port 8001. |
| `M365_AREAS` | all but `bookings` | Which products the agent gets, comma-separated: `mail`, `calendar`, `contacts`, `files`, `lists`, `excel`, `teams`, `recordings`, `tasks`, `notes`, `bookings`. |
| `M365_SWITCHES` | `teams_channels,transcripts,rooms` | Features that need an admin-only permission, plus `chat_admin` (creating group chats and changing members), off by default. Empty turns them all off. |
| `M365_ALLOWED_RECIPIENT_DOMAINS` | none (no limit) | Domains that mail, invitations, file shares and Teams messages may reach, comma-separated. Subdomains are included. |
| `M365_EXPECTED_CLIENT_IDS` | none (any) | Your Entra app's Application (client) ID. A token issued to any other app is refused. Set it. |
| `M365_STATE_*` | none (memory) | The shared store: [state](../components/state.md). |
| `M365_METRICS_PORT` | 9090 | Metrics and probes. |
| `M365_CA_BUNDLE` | none | Extra CAs (PEM) for Graph and the store. |
| `LOG_LEVEL` | `INFO` | Every line is JSON whatever the level. |
| `M365_LINK_SECRET`, `M365_LINK_SIGNING_KEYS`, `M365_LINK_ENCRYPTION_KEYS` | none | **Deprecated.** Still read in 0.2.x with a warning at every start; refused from 0.3.0. Use `M365_KEY_FILE`. |

One person can get fewer products than the installation: the gateway may send the header
`X-M365-Areas: mail,calendar` with their calls.

## Check it

```sh
curl -s http://localhost:9090/readyz
```

`{"ok": true, …, "store": "up"}`. A call to `/mcp` without a person's Microsoft token is refused, which
is the point: tools only work through the gateway.

Next: [Connect a gateway](4-connect-a-gateway.md).
