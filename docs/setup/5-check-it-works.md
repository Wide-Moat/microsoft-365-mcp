# 5. Check it works

As a person with a Microsoft 365 account:

1. In your AI client, connect Microsoft 365. The gateway sends you to Microsoft's sign-in; sign in
   with your work account and accept.
2. Ask: **"What is in my inbox today?"** Compare the answer with Outlook.
3. Ask: **"When are Anna and I both free tomorrow afternoon?"** The server finds Anna by name and
   reads both calendars.
4. Ask: **"Draft a reply to the last message from Anna."** The draft appears in Outlook's Drafts.
   Nothing is sent until you say so.
5. Ask: **"Send it."** The agent asks you to confirm, then sends.

## When sign-in fails

Microsoft's error page shows a code:

| Code | Cause | Fix |
|---|---|---|
| `AADSTS50011` | The redirect URI does not match | Make the app's redirect URI exactly `<gateway address>/callback` (step 1) |
| `AADSTS65001` | Consent is missing | Grant admin consent (step 2) |
| `AADSTS7000215` | The client secret is wrong | Copy the secret's **Value**, not its ID, into the gateway (step 4) |
| "Need admin approval" | Admin-only permissions are not consented | Grant admin consent (step 2), or leave those permissions out |

## When a tool fails

The server answers with a sentence that names the cause and what to do: a missing permission, a
person not found, an item that does not exist, Microsoft throttling. Ask the agent to show you the
exact message.

## Rotating the client secret

Create a new secret in the app (step 1), put it in the gateway, restart the gateway, then delete the
old secret. People stay connected: their refresh tokens do not depend on the secret.
