# 4. Connect a gateway

The server does not sign anyone in. An MCP gateway does: it sends each person through Microsoft's
sign-in once, keeps their token, and adds `Authorization: Bearer <token>` to every call it forwards to
`http://<server>:8000/mcp`.

Any gateway that can do OAuth 2.0 authorization code flow per person and forward the token works. The
example below is [LiteLLM](https://docs.litellm.ai/docs/mcp)'s proxy configuration.

## LiteLLM

Put the app's ID and secret from step 1 in the gateway's environment as `M365_CLIENT_ID` and
`M365_CLIENT_SECRET`, then add the server to `proxy_config`:

```yaml
mcp_servers:
  microsoft365:
    url: http://m365-mcp:8000/mcp
    transport: http
    description: "Your Microsoft 365: mail, calendar, contacts, files, lists, Excel, Teams, To Do, Planner and OneNote."
    auth_type: oauth2
    oauth2_flow: authorization_code
    client_id: os.environ/M365_CLIENT_ID
    client_secret: os.environ/M365_CLIENT_SECRET
    authorization_url: https://login.microsoftonline.com/<tenant ID>/oauth2/v2.0/authorize
    token_url: https://login.microsoftonline.com/<tenant ID>/oauth2/v2.0/token
    token_endpoint_auth_method: client_secret_post
    scopes:
      - offline_access
      - https://graph.microsoft.com/User.Read
      - https://graph.microsoft.com/User.ReadBasic.All
      - https://graph.microsoft.com/People.Read
      - https://graph.microsoft.com/MailboxSettings.ReadWrite
      - https://graph.microsoft.com/Mail.ReadWrite
      - https://graph.microsoft.com/Mail.Send
      - https://graph.microsoft.com/Calendars.ReadWrite
      - https://graph.microsoft.com/Calendars.Read.Shared
      - https://graph.microsoft.com/Presence.Read.All
      - https://graph.microsoft.com/OnlineMeetings.Read
      - https://graph.microsoft.com/Contacts.ReadWrite
      - https://graph.microsoft.com/Files.ReadWrite.All
      - https://graph.microsoft.com/Sites.ReadWrite.All
      - https://graph.microsoft.com/Chat.ReadWrite
      - https://graph.microsoft.com/ChatMessage.Send
      - https://graph.microsoft.com/ChannelMessage.Send
      - https://graph.microsoft.com/Team.ReadBasic.All
      - https://graph.microsoft.com/Channel.ReadBasic.All
      - https://graph.microsoft.com/ChannelMessage.Read.All
      - https://graph.microsoft.com/ChannelMessage.ReadWrite
      - https://graph.microsoft.com/OnlineMeetingTranscript.Read.All
      - https://graph.microsoft.com/OnlineMeetingRecording.Read.All
      - https://graph.microsoft.com/Place.Read.All
      - https://graph.microsoft.com/Tasks.ReadWrite
      - https://graph.microsoft.com/Notes.ReadWrite
```

- For a multitenant app, use `organizations` in place of `<tenant ID>` in both URLs.
- The scopes must match the permissions on the app (step 1). Remove the ones you left out there.
- The gateway needs egress to `login.microsoftonline.com` for the token exchange.
- Set `M365_EXPECTED_CLIENT_IDS` on the server to the same Application (client) ID, so the server
  accepts only tokens issued to your gateway.

## Other gateways

What the server expects on every request to `/mcp`:

- `Authorization: Bearer <the person's Microsoft Graph access token>`, issued for
  `https://graph.microsoft.com` with the scopes above;
- optionally `X-M365-Areas: mail,calendar` to give this person fewer products than the installation.

The server is stateless: each request stands alone, and a refreshed token is simply the next one it
receives.

Next: [Check it works](5-check-it-works.md).
