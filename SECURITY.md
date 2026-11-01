# Security

This server holds a person's delegated Microsoft Graph token for the length of one request and acts
as that person. A defect here can read or send someone's mail, so reports are welcome and taken
seriously.

## Reporting a vulnerability

Report it privately through GitHub: **Security → Report a vulnerability** on this repository. Please
do not open a public issue for a vulnerability.

Include what you found, how to reproduce it, and what an attacker gains. We acknowledge a report
within three working days and keep you informed until it is fixed.

## Supported versions

The latest release on `main` receives security fixes.

## What is in scope

- Token handling: the token logged, echoed, stored or sent anywhere but Microsoft Graph.
- A way past a refusal in [`docs/requirements.md`](docs/requirements.md): sending without the second
  step, deleting for good, reaching another person's mailbox, more recipients than allowed.
- Signed links: forging, replaying past their expiry, using one for another operation.
- Content from other people (mail, chats, documents) that the server passes to the model unmarked.
- Server-side request forgery through URLs that Graph returns.

The threat model is in [`docs/architecture.md`](docs/architecture.md).
