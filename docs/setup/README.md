# Setting it up

Four people may be involved. In a small company they are one person.

| Step | Who | Once per | Guide |
|---|---|---|---|
| 1. Register an app in Microsoft Entra | Microsoft 365 administrator | organisation | [Register the Entra app](1-entra-app.md) |
| 2. Grant consent for the organisation | Microsoft 365 administrator | organisation | [Grant admin consent](2-admin-consent.md) |
| 3. Run the server | whoever runs your infrastructure | installation | [Run the server](3-run-the-server.md) |
| 4. Connect it to your MCP gateway | whoever runs your infrastructure | installation | [Connect a gateway](4-connect-a-gateway.md) |
| 5. Connect your account and try it | each person | person | [Check it works](5-check-it-works.md) |

## How the pieces fit

```
person ── signs in to Microsoft once ──▶ MCP gateway ── keeps the token
AI agent ── calls a tool ──▶ MCP gateway ── adds the person's token ──▶ m365-mcp ──▶ Microsoft Graph
agent's sandbox ── curl a signed link ──▶ m365-mcp ── streams the file ──▶ Microsoft Graph
```

- **MCP** (Model Context Protocol) is how an AI agent calls tools.
- **Microsoft Graph** is the API behind Microsoft 365.
- **The MCP gateway** sits between the agent and its tools. It signs each person in to Microsoft
  (OAuth) and adds that person's token to every call. The server never stores a token.
- **Delegated sign-in** means every call runs as the person, with their rights. The agent sees what the
  person sees in Outlook or OneDrive, and nothing else.

## What you need

- A Microsoft 365 tenant, and an administrator who can register apps and grant consent.
- A place to run one container that your gateway can reach, and an https address for its file links.
- An MCP gateway that does OAuth per person and forwards the token, such as LiteLLM.
