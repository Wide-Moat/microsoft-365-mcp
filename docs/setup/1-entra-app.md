# 1. Register the Entra app

The gateway signs people in to Microsoft through an app registration in Microsoft Entra. You register
it once. It holds no data; it only says which permissions the gateway may ask for.

Open the [Microsoft Entra admin center](https://entra.microsoft.com) as an administrator.

## Create the registration

1. **Identity → Applications → App registrations → New registration.**
2. **Name:** what people will see on Microsoft's sign-in screen, for example "AI assistant".
3. **Supported account types:**
   - **Single tenant** if only your organisation will sign in. Use your tenant ID in the gateway's
     URLs (step 4).
   - **Multitenant** if you run the server for several organisations. Each one then grants consent
     separately (step 2).
4. **Redirect URI:** platform **Web**, value `<your gateway's public address>/callback`. It must match
   exactly, including the path.
5. **Register.** Copy the **Application (client) ID** and the **Directory (tenant) ID**.

## Add the permissions

**API permissions → Add a permission → Microsoft Graph → Delegated permissions.** Add these. All are
delegated: every call runs as the person who signed in.

| Area | Permissions |
|---|---|
| Always | `User.Read`, `User.ReadBasic.All`, `People.Read`, `MailboxSettings.ReadWrite`, `offline_access` |
| Mail | `Mail.ReadWrite`, `Mail.Send` |
| Calendar | `Calendars.ReadWrite`, `Calendars.Read.Shared`, `Presence.Read.All`, `OnlineMeetings.Read` |
| Contacts | `Contacts.ReadWrite` |
| Files, lists, Excel | `Files.ReadWrite.All`, `Sites.ReadWrite.All` |
| Teams chats | `Chat.ReadWrite`, `ChatMessage.Send`, `ChannelMessage.Send`, `Team.ReadBasic.All`, `Channel.ReadBasic.All` |
| To Do and Planner | `Tasks.ReadWrite` |
| OneNote | `Notes.ReadWrite` |
| Teams channels (admin only) | `ChannelMessage.Read.All`, `ChannelMessage.ReadWrite` |
| Meeting recordings and transcripts (admin only) | `OnlineMeetingRecording.Read.All`, `OnlineMeetingTranscript.Read.All` |
| Rooms (admin only) | `Place.Read.All` |

Leave out the areas you will switch off in the server (`M365_AREAS`, step 3). The consent in step 2
grants exactly what is listed here, so a missing permission means a feature that fails later.

## Create a secret

**Certificates & secrets → Client secrets → New client secret.** Copy the **Value** now: Microsoft
shows it once. A secret lasts at most 24 months; note the date and create the next one before it.

## Optional: a verified publisher

Without one, Microsoft's consent screen calls the app "unverified". Administrators can still consent,
but some organisations allow people to consent only to verified apps. **Branding & properties →
Publisher domain**, then verify through the Microsoft AI Cloud Partner Program.

## You now have

- Application (client) ID
- Directory (tenant) ID
- Client secret value

Next: [Grant admin consent](2-admin-consent.md).
