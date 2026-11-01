# 2. Grant admin consent

Five permissions are ones only an administrator can grant: Teams channel reading and editing, meeting
recordings and transcripts, and rooms. The gateway asks for all permissions in one sign-in, so until
an administrator consents, Microsoft stops every person at **"Need admin approval"**.

Consent once, for the whole organisation, before anyone connects.

## Open the consent link

A Global Administrator, Privileged Role Administrator or Cloud Application Administrator opens:

```
https://login.microsoftonline.com/<tenant ID>/adminconsent?client_id=<Application (client) ID>&redirect_uri=<your gateway's public address>/callback
```

For a multitenant app used by another organisation, replace `<tenant ID>` with `organizations` and
send the link to that organisation's administrator.

1. Sign in.
2. Check that the list shows every permission you added in step 1.
3. Accept **on behalf of your organisation**.

Microsoft then sends you to the gateway's callback, which shows an error because nobody is signing in
at that moment. That is expected: the consent is recorded.

## Check it

**Entra admin center → Identity → Applications → Enterprise applications →** your app **→
Permissions.** Each permission shows **Granted for** your organisation.

## Without the five admin-only permissions

You can leave them out. Remove them from the app (step 1) and from the gateway's scopes (step 4), and
switch the matching features off in the server:

```sh
M365_SWITCHES=            # no Teams channel reading, transcripts or rooms
M365_AREAS=mail,calendar,contacts,files,lists,excel,teams,tasks,notes   # every default area but recordings
```

People can then consent for themselves, but only where the organisation's consent policy allows it
(**Enterprise applications → Consent and permissions**). Many organisations allow it only for verified
publishers, or not at all.

Next: [Run the server](3-run-the-server.md).
