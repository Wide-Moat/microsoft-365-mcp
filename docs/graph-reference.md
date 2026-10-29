# Microsoft Graph: what the documentation says

These facts come from the Microsoft Learn pages on Graph v1.0 and cover delegated permissions only.
They were read on 2026-10-22 and are what the [requirements](requirements.md) are built on. Where a
fact here and the behaviour of a real tenant disagree, the tenant wins. That result goes into
`graph-live-findings.md` with a recorded fixture, and this page gets a note pointing to it.

Short forms: `A` = `https://learn.microsoft.com/en-us/graph/api`, `G` = `https://learn.microsoft.com/en-us/graph`.
**Unknown** means the documentation does not say. Each unknown is settled on the test tenant
before the code relies on it.

## Teams chats

The chat APIs work only with work or school accounts.

### Listing chats

`GET /me/chats` ([A/chat-list](https://learn.microsoft.com/en-us/graph/api/chat-list)):
- `$expand` accepts only `members` and `lastMessagePreview`.
- `$top` is at most 50.
- `$orderby` accepts only `lastMessagePreview/createdDateTime desc`.
- An expanded chat returns at most 25 members.
- A chat is unread when `lastMessagePreview.createdDateTime` is later than
  `viewpoint.lastMessageReadDateTime`.

### Reading messages

`GET /me/chats/{id}/messages` ([A/chat-list-messages](https://learn.microsoft.com/en-us/graph/api/chat-list-messages)):
- `$top` is at most 50.
- Order is descending only, by `lastModifiedDateTime` (the default) or `createdDateTime`.
- `$filter` accepts `gt` and `lt` on `lastModifiedDateTime`, and only `lt` on `createdDateTime`.
- `$filter` is **silently ignored** unless `$orderby` names the same property.
- System messages come back with the body `<systemEventMessage/>`.

### Sending

`POST /chats/{id}/messages` ([A/chatmessage-post](https://learn.microsoft.com/en-us/graph/api/chatmessage-post)):
- The page lists `ChannelMessage.Send` for chats. `ChatMessage.Send` exists in the permissions
  reference. Which one is enforced is **unknown**. Settled for this app (live finding `chat_send_scope`): consent always covers both, and even a refresh asking for neither returns both, so a token lacking one cannot be had.
- `contentType` is `html` or `text`. Teams accepts only part of HTML, and Microsoft does not
  publish the allowed tags (**unknown**). Settled by live finding `teams_html_tags`: `pre` is dropped, the rest are kept.
- **Mentions:** the body carries `<at id="0">Name</at>`, plus `mentions[{id, mentionText,
  mentioned.user{id, displayName, userIdentityType: aadUser}}]`. The body must be HTML.
- **Files** must already be in OneDrive or SharePoint. The message carries
  `attachments[{id: <eTag GUID>, contentType: "reference", contentUrl: <webDavUrl>, name}]`, and the
  body carries `<attachment id="…"></attachment>`.
- **Inline images** are sent as `hostedContents` with a temporary id, at most 4 MB each.
- There is no threaded reply in chats. Only channel messages have `/replies`.

### Creating chats and adding members

- `POST /chats` ([A/chat-post](https://learn.microsoft.com/en-us/graph/api/chat-post)): `chatType` is
  `oneOnOne` or `group`, and the caller must be among the members. If a one-to-one chat already
  exists, Graph returns it.
- Adding a member: `POST /chats/{id}/members` ([A/chat-post-members](https://learn.microsoft.com/en-us/graph/api/chat-post-members)).
  `Chat.ReadWrite` is enough, so `ChatMember.ReadWrite` and its administrator consent are not needed.

### Editing and soft delete

These are not used in v1 and remain candidates.
- Edit: `PATCH .../messages/{id}` with `Chat.ReadWrite`.
- Soft delete and its undo: `POST /users/{id}/chats/{id}/messages/{id}/softDelete` and
  `undoSoftDelete`.

## Teams channels

- **Joined teams:** `GET /me/joinedTeams` with `Team.ReadBasic.All` ([A/user-list-joinedteams](https://learn.microsoft.com/en-us/graph/api/user-list-joinedteams)).
  Host teams of shared channels are not included.
- **Channels:** `GET /teams/{id}/channels` with `Channel.ReadBasic.All`. Use `$select` to leave out
  `email`, which is slow.
- **Messages:** `GET /teams/{t}/channels/{c}/messages` needs **`ChannelMessage.Read.All`, which an
  administrator must consent to** ([A/channel-list-messages](https://learn.microsoft.com/en-us/graph/api/channel-list-messages)).
  - `$top` is at most 50, and `$expand=replies` is supported.
  - There is no `$filter` and no `$orderby`. Messages are sorted by the last activity in their
    thread.
- **Posting and replying:** `ChannelMessage.Send`, which needs no administrator consent.
- **Files tab:** `GET .../filesFolder` returns a driveItem.

## Search

`POST /search/query` ([A/search-query](https://learn.microsoft.com/en-us/graph/api/search-query)):
- One `searchRequest` per call.
- `chatMessage` cannot be combined with other entity types and has no custom sort.
- `driveItem`, `drive`, `site`, `list` and `listItem` can be mixed with each other.
- `person` replaces the People API, which is in maintenance mode.

## Meetings

- **Finding the meeting:** `GET /me/onlineMeetings?$filter=JoinWebUrl eq '…'` with
  `OnlineMeetings.Read`.
- **Transcripts:** `/me/onlineMeetings/{id}/transcripts/{tid}/content` needs
  **`OnlineMeetingTranscript.Read.All` (administrator consent)**.
  - The default format is VTT with speakers.
  - `Accept: application/vnd.microsoft.graph.transcript+text` returns the text without speakers.
  - Tenant policy produces two errors: `GraphAccessToTranscriptsDisabled`, and
    `SpeakerAttributionNotAllowed` (retry without speakers).
  - Transcripts are not available for meetings that are not tied to a calendar event.
- **Recordings:** `OnlineMeetingRecording.Read.All` (administrator consent). Only the organiser can
  download them. Out of scope for v1.

## Presence

`POST /communications/getPresencesByUserId` with `Presence.Read.All`, which needs no administrator
consent:
- up to 650 ids per call;
- 1,500 requests per 30 s per app per tenant.

## OneDrive and SharePoint

- **Shared with me:** `GET /me/drive/sharedWithMe` is **deprecated**. It returns degraded data until
  November 2026 and then stops. Microsoft names no replacement; we use search instead.
- **Search inside one drive:** `GET /me/drive/root/search(q='…')` supports `$top`, `$skipToken` and
  `$orderby`.
- **Search across sites:** `/search/query` with `driveItem` (`Files.Read.All` or `Sites.Read.All`).
- **Sites:** `GET /sites?search=` (`Sites.Read.All`) and `GET /me/followedSites`. Microsoft notes a
  known sync issue on `followedSites`. A team's site is `GET /groups/{id}/sites/root` for each of
  `/me/joinedTeams`; search can leave it out (live finding `site_search`).
- **Download:** `GET .../content` answers 302 with a pre-authenticated URL. `Range` goes on that
  URL, not on `/content`.
- **Conversion:** `?format=pdf` covers doc, docx, xls, xlsx, ppt, pptx, rtf, odt, md, eml, msg and
  others. The page states a 2 MB source limit, which has to be verified live.
- **Upload:**
  - `PUT .../content` accepts up to 250 MB.
  - An upload session (`createUploadSession`) sends fragments in order. Each fragment is a multiple
    of 320 KiB and under 60 MiB.
  - Conflict behaviour is `fail`, `replace` or `rename`.
- **Copy:** returns 202 with a monitor URL and is asynchronous. **Move and rename** are a `PATCH` of
  `parentReference` / `name`. **Create a folder:** `POST .../children` with `folder: {}`.
- **Delete:** `DELETE` moves the item to the recycle bin. `permanentDelete` is irreversible and
  forbidden here. **Restore from the recycle bin works only for OneDrive Personal**, not for work
  or school accounts.
- **Versions:** list them, and roll back with `restoreVersion` (`Files.ReadWrite`).
- **Sharing:**
  - `createLink` takes `type` `view` or `edit`, and `scope` `anonymous`, `organization` or `users`.
    We never send `anonymous`. Without a `scope`, the tenant's default link type is created, so we
    always send one.
  - `invite` takes `recipients`, `roles` (`read` or `write`), `sendInvitation` and `message` (at
    most 2,000 characters). A partial failure returns 207.
  - A permission can be deleted only when it is not inherited.
- **SharePoint lists:** items need `Sites.Read.All` to read and `Sites.ReadWrite.All` to write.
  A candidate, not in v1.

## Mail

- **Drafts and sending:** `POST /me/messages` creates a draft, and `POST /me/messages/{id}/send`
  sends it (`Mail.Send`).
- **Replies:** `createReply` (`Mail.ReadWrite`) takes **either** `comment` **or** `message.body`.
  Both together return 400.
- **Size:** the default message size limit is 35 MB, and administrators can change it.
  - Attachments under 3 MB go in one POST.
  - Attachments from 3 to 150 MB use an upload session, in chunks of up to 4 MB.
- **Listing:**
  - `$top` goes up to 1,000, but large pages that include bodies risk a 504.
  - With `$filter` plus `$orderby`, every `$orderby` property must also appear in `$filter`,
    first and in the same order.
- **`$search`:** returns up to 1,000 results, sorted by send date. KQL accepts from, to, cc,
  subject, body, attachment, hasAttachments, received and others. Whether it combines with
  `$filter` or `$orderby` is **unknown**. Settled by live finding `search_with_filter`: neither combines with `$search` (400).
- **`uniqueBody`** comes back only when asked for with `$select`.
- **Time zone:** `GET /me/mailboxSettings` (`MailboxSettings.Read`).
- **Immutable ids:** with `Prefer: IdType="ImmutableId"` on **every** request, an id survives a
  move within the mailbox. It does change on archiving or export. Calendar and folder ids are
  constant anyway. [G/outlook-immutable-id](https://learn.microsoft.com/en-us/graph/outlook-immutable-id)

## Scheduling

- `POST /me/calendar/getSchedule` (`Calendars.ReadBasic`) returns an error when a slot holds more
  than 1,000 entries.
- `POST /me/findMeetingTimes` (`Calendars.Read.Shared`).
- Neither needs administrator consent.

## Throttling

From [G/throttling-limits](https://learn.microsoft.com/en-us/graph/throttling-limits):
- **Outlook:** 10,000 requests per 10 min per app and mailbox, 4 concurrent requests, 150 MB of
  uploads per 5 min.
- **Teams:** about 1 rps per person and 1 rps per chat or channel. Members: 4 per minute per chat.
- **SharePoint and OneDrive:** 3,000 requests per 5 min per person.
- Always honour `Retry-After`. Without it, back off exponentially.

## Administrator consent (delegated)

From [G/permissions-reference](https://learn.microsoft.com/en-us/graph/permissions-reference).

**Needs consent:**
- `ChannelMessage.Read.All`, `ChannelMessage.ReadWrite`
- `ChatMember.ReadWrite`
- `OnlineMeetingTranscript.Read.All`, `OnlineMeetingRecording.Read.All`
- `People.Read.All`, `Place.Read.All`
- `Sites.FullControl.All`, `Group.Read.All`

**Needs no consent:**
- `Chat.*`, `ChatMessage.Send`, `ChannelMessage.Send`
- `Team.ReadBasic.All`, `Channel.ReadBasic.All`
- `OnlineMeetings.Read`, `Presence.Read.All`
- `Files.*`, `Sites.Read(Write).All`
- `Mail.*`, `MailboxSettings.*`, `Calendars.*`
- `People.Read`, `User.Read`, `User.ReadBasic.All`

A tenant's own consent policy can still require an administrator for any of them.

## Contacts, To Do, Planner, lists, OneNote, Excel, automatic replies, recordings, Bookings

Read on 2026-10-22, from the same sources.

### Contacts

- **Scope:** `Contacts.ReadWrite`, no administrator consent.
- **Endpoints:** `/me/contacts` and `/me/contactFolders`.
- **Filtering:** only `emailAddresses/any(a:a/address eq '…')` works. `$search` is not documented.
  Names are found through `/search/query` (person) or the directory instead.
- **Update:** a `PATCH` of a structured property (an address) must send the whole property.
- **Delete:** where a deleted contact goes is unknown. Settled by live finding `deleted_contact_location`: Deleted Items, under the same id.

### Directory

- `GET /users?$search="displayName:…"` needs `ConsistencyLevel: eventual`.
- `User.ReadBasic.All`, no administrator consent.
- `$skip` is not supported.

### To Do

- **Scope:** `Tasks.ReadWrite`, no administrator consent.
- **Endpoints:** `/me/todo/lists/{id}/tasks`, plus `checklistItems` and `linkedResources`.
- **Ids:** a task's id changes when it moves to another list.
- **Unknown:** paging limits, and whether deletion is permanent. Settled by live findings `todo_paging` (50 a page by default, `$top` honoured to at least 60, `@odata.nextLink`) and `deleted_todo_task` (gone: 400 on GET, not listed).

### Planner

- **Scope:** `Tasks.Read` / `Tasks.ReadWrite` cover group plans. `Group.Read.All` is not needed.
- **Endpoints:** `/me/planner/tasks`, `/me/planner/plans` (without plans made in a team: those are
  `/groups/{id}/planner/plans`), `/planner/plans/{id}/buckets|tasks`,
  `/planner/tasks/{id}/details`.
- **Writes:** `PATCH` and `DELETE` require `If-Match` with the item's etag. A task and its details
  have separate etags. Conflicts return 409 or 412.
- **Assignments:** keyed by user id, with `@odata.type: #microsoft.graph.plannerAssignment` and
  `orderHint`.
- **Not exposed:** premium plans.
- **Throttling:** no published limits.

### SharePoint lists

- **Scopes:** `Sites.Read.All` / `Sites.ReadWrite.All`, no administrator consent.
- **Endpoints:**
  - read: `GET /sites/{id}/lists/{id}/items?expand=fields(select=…)`;
  - create: `POST …/items {fields}`;
  - update: `PATCH …/items/{id}/fields` (optional `If-Match`, 412 on mismatch);
  - delete: `DELETE`.
- **Filtering:** `$filter` on `fields/X` supports eq, ne, lt, gt, le, ge and startswith. It works
  best on indexed columns, one indexed field at a time.
- **Delete:** whether it goes to the recycle bin is unknown. Settled by live finding `deleted_list_item`: it goes to the site recycle bin (read through beta; v1.0 has none).

### OneNote

- **Scopes:** `Notes.ReadWrite`, no administrator consent.
- **Page content:** `GET /me/onenote/pages/{id}/content` returns HTML.
- **Create:** `POST /sections/{id}/pages` with HTML, or multipart for images and files.
- **Update:** `PATCH …/content` with `{target, action: append|insert|prepend|replace, content}`.
  Tables are replaced whole.
- **Size limits:** 4 MB per request; about 70 MB per POST; 25 MB per part; 150 images per page.
- **Other limits:**
  - Libraries with more than 5,000 OneNote items cannot be queried (error 10008).
  - Pages: 20 per page by default, at most 100.
- **Throttling:** 120 requests per minute and 400 per hour per person, 5 concurrent. **No
  `Retry-After` on 429.**

### Excel

- **Base path:** `…/items/{id}/workbook/…`, `Files.ReadWrite`.
- **Sessions:** `createSession {persistChanges}`, header `workbook-session-id`. A persistent
  session expires after about 5 minutes idle, a non-persistent one after about 7; either answers
  404 once expired.
- **Formats:** `.xls` and OneDrive consumer files are unsupported.
- **Ranges:**
  - Endpoints: `range(address=…)` for read and `PATCH`, `usedRange`, `tables/{id}/rows`.
  - A `null` value leaves the cell unchanged.
  - Ranges over 5 million cells fail.
- **Throttling:** 1,500 requests per 10 s per app per tenant. Avoid concurrent writes to one
  workbook.

### Automatic replies

- **Endpoints:** read `GET /me/mailboxSettings/automaticRepliesSetting`; write
  `PATCH /me/mailboxSettings` with `{automaticRepliesSetting}`.
- **Scope:** `MailboxSettings.ReadWrite`, no administrator consent.
- **Fields:**
  - `status`: `disabled`, `alwaysEnabled` or `scheduled`;
  - `externalAudience`: `none`, `contactsOnly` or `all`;
  - internal and external messages;
  - scheduled start and end.

### Recordings and transcripts

- **Endpoints:** `/me/onlineMeetings/{id}/recordings[/{id}/content]` (MP4) and
  `/transcripts/{id}/content` (VTT).
- **Scopes:** `OnlineMeetingRecording.Read.All` and `OnlineMeetingTranscript.Read.All`, both need
  **administrator consent**.
- **Download rights:** under delegated permissions, **only the organiser can download recording
  content**. Participants cannot, unless an administrator unblocks it. Transcripts are available
  to invitees.
- **Finding the meeting:** take the event's `onlineMeeting.joinUrl`, then
  `/me/onlineMeetings?$filter=JoinWebUrl eq '…'`.
- **Not available:** expired meetings, meetings with no calendar event, live events.
- **Storage:**
  - a private meeting's recording is in the organiser's OneDrive, folder `Recordings`;
  - a channel meeting's recording is in the channel's SharePoint, folder `Recordings`;
  - invitees get a view link, without download for scheduled meetings.
- **Size:** about 400 MB per hour of video.
- **Downloading:** `Range` goes on the pre-authenticated download URL.

### Bookings

- **Licence:** requires Microsoft 365 Business Premium for the business. Shared booking pages only.
- **Scopes:** `Bookings.Read.All` to read; `BookingsAppointment.ReadWrite.All` to create, update,
  cancel and delete. No administrator consent.
- **Endpoints:**
  - `/solutions/bookingBusinesses`: up to 500 results, no paging;
  - `…/{id}/appointments|services|staffMembers`;
  - `POST …/appointments/{id}/cancel {cancellationMessage}`, which notifies the customer and staff.
- **Throttling:** 4 concurrent requests per mailbox.

### Teams: editing your own messages

- **Edit:**
  - chat message: `PATCH /chats/{id}/messages/{id}`, `Chat.ReadWrite`;
  - channel message: `ChannelMessage.ReadWrite`, needs administrator consent.
- **Delete and restore:** `softDelete` and `undoSoftDelete` on
  `/users/{id}/chats/{id}/messages/{id}`. A `/me` form is not documented.
- **Throttling:** 1 request per second per chat.

### Insights

`/me/insights/used` and `/shared` stop returning data in November 2026; `/trending` remains. Not
used.
