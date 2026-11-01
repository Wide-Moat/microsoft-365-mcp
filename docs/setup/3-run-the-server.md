# 3. Run the server

One container, two ports:

| Port | Path | Who reaches it |
|---|---|---|
| 8000 | `/mcp` | your MCP gateway only. Never publish it. |
| 8001 | `/d/…`, `/u/…` | the agent's sandbox, to download and upload files through signed links. Publish it on an https address. |

## Start it

Generate one secret for the file links, at least 32 characters, and keep it in your secret store:

```sh
openssl rand -base64 48
```

Then run the image from a release (or build it with `docker build -t m365-mcp .`):

```sh
docker run -d --name m365-mcp \
  -p 8000:8000 -p 8001:8001 \
  -e M365_LINK_SECRET='<the secret>' \
  -e M365_LINKS_BASE_URL='https://m365-files.example.com' \
  ghcr.io/wide-moat/microsoft-365-mcp:<version>
```

`M365_LINKS_BASE_URL` is the https address that reaches port 8001. The server writes it into every
file link. Without it, every tool works except moving files.

Run **one replica**. A file link lives in the memory of the process that issued it, so a second replica
would refuse half the links.

## Settings

| Variable | Default | What it does |
|---|---|---|
| `M365_LINK_SECRET` | none | The secret both link keys are derived from. Required for file links. |
| `M365_LINKS_BASE_URL` | none | The public https address of port 8001. |
| `M365_AREAS` | all but `bookings` | Which products the agent gets, comma-separated: `mail`, `calendar`, `contacts`, `files`, `lists`, `excel`, `teams`, `recordings`, `tasks`, `notes`, `bookings`. |
| `M365_SWITCHES` | `teams_channels,transcripts,rooms` | Features that need an admin-only permission, plus `chat_admin` (creating group chats and changing members), off by default. Empty turns them all off. |
| `M365_ALLOWED_RECIPIENT_DOMAINS` | none (no limit) | Domains that mail, invitations, file shares and Teams messages may reach, comma-separated. Subdomains are included. |
| `M365_EXPECTED_CLIENT_IDS` | none (any) | Your Entra app's Application (client) ID. A token issued to any other app is refused. Set it. |

One person can get fewer products than the installation: the gateway may send the header
`X-M365-Areas: mail,calendar` with their calls.

## On Kubernetes

The same container, with:

- a Deployment of one replica and a rolling update that never runs two pods at once (`maxSurge: 0`);
- a Service for port 8000 that only the gateway's pods may reach (a NetworkPolicy);
- an Ingress or HTTPRoute that publishes only `/d/` and `/u/` of port 8001;
- egress on 443 to `graph.microsoft.com` and to the hosts Microsoft serves files from
  (`*.sharepoint.com`, `*.onedrive.com`, `*.1drv.com`, `*.svc.ms`); the server follows file redirects
  only to Microsoft's own domains;
- `M365_LINK_SECRET` from a Secret.

## Check it

```sh
curl -s http://localhost:8000/mcp -o /dev/null -w '%{http_code}\n'
```

Any answer from the server means it is up. A call without a person's Microsoft token is refused, which
is the point: tools only work through the gateway.

Next: [Connect a gateway](4-connect-a-gateway.md).
