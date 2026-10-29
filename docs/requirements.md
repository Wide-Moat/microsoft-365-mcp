# Requirements

The tools, what each must do, and what each must refuse. Every row in a **Must refuse** table is a
test, and its **Test** column names it: pytest ids relative to `tests/unit/`, one or more per row,
each test named by one row only. `scripts/refusal_map.py` fails the suite on a row without a test or
a test that does not exist, and `scripts/refusal_evidence.py` removes each row's guard and requires
its test to fail (manifesto rule 36). The rules behind them are in the [manifesto](../MANIFESTO.md);
numbers in brackets point there.

There are 43 tools; 39 are published by default (Bookings is off, and `save_chat` sits behind the
`chat_admin` switch, which is off). One question is one tool: where two questions differ only in
the form or depth of the answer, they are one tool with a parameter, never two tools (rule 1).

## Conventions shared by every tool

| Input | Accepted form | Server does |
|---|---|---|
| Person | display name, part of a name, or an address | resolves through the person's contacts and People, then the directory; a name that several people match is an error listing up to 5 candidates with name, address, job title — even when one display name equals it exactly; an address at the organisation's own domains must be a user in its directory ([resolve](components/resolve.md)) |
| Mail folder | well-known name (`inbox`, `sent`, `drafts`, `deleted`, `archive`, `junk`), a display name, or a path `Inbox/Projects` | resolves; two folders with one name is an error listing their paths |
| Time | `YYYY-MM-DD HH:MM` (local to the person's mailbox time zone) or with an explicit offset | converts; output is local with offset (`2026-10-23T15:00:00+01:00`) |
| Duration | minutes | — |
| Cursor | opaque string from the previous page | wraps Graph's `@odata.nextLink`; a cursor from another tool or person is refused |
| Limit | 1–25, default 10 | values above 25 are clamped and the result says so; below 1 is refused |

Every result that can be cut carries `truncated: true` and the size it was cut from. Every list
carries `has_more` and `cursor`.

A parameter that needs an area or a switch of its own (channel reading in `get_conversation`,
`include_recording` and `include_transcript` in `get_event`, `rooms` in `find_meeting_times`) is
removed from the tool's schema while that area or switch is off — for the installation, or for one
person through `X-M365-Areas` — exactly as a switched-off area removes whole tools (rule 13). Sent
anyway, it is refused before anything reaches Graph.

Every tool, on every call, must refuse with a clear `isError`:

| Must refuse | The person is told | Test |
|---|---|---|
| No token / expired and not refreshable | reconnect Microsoft 365 in the portal | `test_account.py::test_no_token_means_reconnect_and_no_graph_call`, `test_account.py::test_expired_token_means_reconnect` |
| A token for another audience, or, when the installation names its gateway's clients, issued to another client | reconnect Microsoft 365 in the portal | `test_identity.py::test_a_token_for_another_audience_is_refused_before_graph`, `test_identity.py::test_a_token_issued_to_another_client_is_refused_before_graph` |
| Graph `403` for a missing delegated scope | which permission is missing; that a Microsoft 365 administrator must grant it | `test_refusals.py::test_a_missing_permission_is_named_with_the_administrator_step` |
| Graph `404` on an id the person cannot see | the item does not exist or is not shared with you — never which | `test_mail.py::test_not_found_does_not_say_which` |
| An id Graph cannot parse (`ErrorInvalidIdMalformed`, To Do's `invalidRequest` around it, OneNote's `20112`, Teams' `ParsingFailed` on a read, and Teams' bare `400` — `BadRequest` "UnknownError" or `InvalidParameter` — for a message id that is not a number, on reads and writes, measured; a To Do delete reads the task first, because Graph answers DELETE of such an id with 500, measured) | the same: it does not exist or is not shared with you | `test_refusals.py::test_a_malformed_id_is_not_found`, `test_refusals.py::test_a_todo_delete_of_an_id_to_do_cannot_read_is_not_found`, `test_teams.py::test_a_message_id_that_is_not_one_is_not_found` |
| An id holding a path separator (`/`, `\`), `?`, `#` or a space, which no tool returns (`a/b` became two path segments) | the same: it does not exist or is not shared with you; nothing reaches Graph | `test_refusals.py::test_an_id_holding_a_separator_is_not_found_before_graph` |
| An empty id | that an id or path is empty or not valid; nothing reaches Graph, which would answer `/me/events/` with the whole collection | `test_refusals.py::test_an_empty_id_is_refused_before_graph` |
| An empty id given to a `save_` tool, which would read it as "create" (QA: `save_event` and `save_task` created new items) | that an id or path is empty or not valid; nothing is created | `test_refusals.py::test_an_empty_id_on_a_save_is_refused_not_read_as_new` |
| Arguments outside the schema: a value outside an enum, a wrong type, a missing argument | the field and what it accepts — never the validator's report or the value given | `test_refusals.py::test_a_value_outside_an_enum_names_the_allowed_values` |
| Graph `403` because the person's OneDrive was never provisioned (`provisioningNotAllowed`) | that OneDrive is not set up or not licensed; open it once or ask the administrator — never "reconnect" | `test_refusals.py::test_an_unprovisioned_onedrive_is_not_a_sign_in_problem` |
| Graph `401` from one service while the token is valid (OneNote's `40001` for a person without OneNote) | that the service is not set up for the account — never "reconnect"; only `InvalidAuthenticationToken` means the sign-in | `test_refusals.py::test_a_service_401_with_a_valid_token_does_not_say_reconnect` |
| Graph `429` / `503` after retries honouring `Retry-After` | Microsoft is throttling; try again in N seconds | `test_refusals.py::test_throttling_after_retries_says_how_long_to_wait` |
| Any other Graph failure | that the action did not happen, with a correlation id; after a server error in a tool that writes, that it may or may not have happened — never a success; in a tool that only reads, that nothing was changed | `test_refusals.py::test_a_refused_write_says_nothing_changed_with_a_reference`, `test_refusals.py::test_a_server_failure_on_a_write_never_claims_nothing_changed`, `test_refusals.py::test_a_server_failure_on_a_read_says_nothing_changed` |
| A parameter whose area or switch is off (it is absent from the schema too) | that this part is switched off | `test_contract.py::test_a_parameter_behind_an_off_switch_is_hidden_and_refused`, `test_contract.py::test_a_parameter_of_an_area_off_for_the_person_is_hidden_and_refused` |
| An argument the tool does not have (validation would drop it, and the call would do less than asked: a series becomes one event, a file is not attached) | its name; nothing was done | `test_contract.py::test_an_argument_the_tool_does_not_have_is_refused_before_graph` |
| A `limit` below 1 | the range | `test_contract.py::test_a_limit_below_one_is_refused` |
| A `user_confirmed` that is not the JSON boolean `true` or `false`: a word (`"yes"`, `"true"`), a number, null (validation would read `"yes"` and `1` as true, and mail was sent so) | that only the boolean true confirms; nothing was done | `test_confirmation.py::test_a_confirmation_that_is_not_the_boolean_true_is_refused`, `test_confirmation.py::test_send_draft_with_the_word_yes_sends_nothing` |

## Account and people

### `get_connection`
Who the person is (name, address), mailbox time zone and working hours (`not set in Outlook` when there are none), which areas are available
(mail, calendar, files, teams) and which permissions are missing for the rest. Called first when a
tool fails for rights, or when the person asks what the assistant can do.

### `list_people(query, source=all, include_org=false, limit, cursor)`
Find colleagues by name or address. Returns name, address, job title, department; with
`include_org`, office, manager and phone. Ranked by relevance to the person through Microsoft Search
(`/search/query`, entity `person`) — the People API is in maintenance mode — then the directory
(`User.ReadBasic.All`). `source`: `all`, `contacts` (the person's own Outlook contacts), `directory`.
Each result says where it came from, and a personal contact carries its `contact_id` for `save_contact` and `delete_contact`, so "write to my accountant" finds a personal contact that the
directory does not know. Contacts are read from `/me/contacts` itself and matched on name, company,
job title and address: the People index learns a new contact late or never (a live finding: not
within 35 minutes).

| Must refuse | The person is told | Test |
|---|---|---|
| Empty query | ask for a name or address | `test_account.py::test_list_people_refuses_an_empty_query` |

## Contacts

### `save_contact(id?, name, emails[], phones[], company?, job_title?, notes?, folder?)`
Create or update a personal Outlook contact. Structured fields are sent whole (Graph replaces
them on update), so the tool reads the contact first and merges. `folder` names a contact folder
for a new contact. The id is the `contact_id` that `list_people` returns for each personal contact.

### `delete_contact(id)`
Deletes one personal contact. Where Graph puts it (Deleted Items or nowhere) is a live finding;
until it is settled the description says the deletion may not be undoable and the tool requires
`user_confirmed`.

| Must refuse | The person is told | Test |
|---|---|---|
| A directory person (not a personal contact), by directory id or by address | only personal contacts are changed; where to find a contact's `contact_id` | `test_contacts.py::test_a_directory_person_is_not_a_contact` |
| A contact already deleted: it keeps its id in Deleted Items, where PATCH and a second DELETE succeed and the second DELETE takes it beyond the person's reach (measured) | it sits in Deleted Items; restore it in Outlook first. A contact in none of the contact folders read, when not all could be read, is not changed either: it cannot be told from a deleted one | `test_contacts.py::test_a_contact_in_deleted_items_is_not_changed_or_deleted_again`, `test_contacts.py::test_a_contact_whose_folder_cannot_be_placed_is_not_changed` |
| An email that is not an address (Graph saves `not-an-email`, measured) | that it is not an email address | `test_contacts.py::test_an_email_that_is_not_an_address_is_refused` |
| A fourth email address (Outlook keeps three) | the limit of three | `test_contacts.py::test_a_fourth_email_names_outlooks_limit_of_three` |
| `delete_contact` without `user_confirmed` (until the live finding says it is reversible) | ask the person | `test_contacts.py::test_delete_needs_confirmation` |

## Mail

### `save_mail_folder(id?, name, parent?)`
Create a folder, rename it, or move it under another. No deletion of folders exists. A rename and a
move in one call are checked against the folder tree before either is sent; when Graph still refuses
the move after the rename went through, the error says the folder was renamed and not moved.

| Must refuse | The person is told | Test |
|---|---|---|
| Renaming or moving a well-known folder (Inbox, Sent, Drafts, Deleted, Archive, Junk, Outbox) | system folders are not changed | `test_mail.py::test_system_folder_is_not_renamed` |
| A name that already exists under that parent, on create, rename or move | the existing folder's path | `test_mail.py::test_a_folder_name_that_exists_names_the_existing_folder`, `test_mail.py::test_renaming_or_moving_onto_an_existing_name_names_that_folder` |
| A name containing `/`, which separates folders in a path | choose a name without it; nothing was changed | `test_mail.py::test_a_folder_name_with_a_slash_is_refused` |
| Moving a folder inside itself or one of its subfolders | that it cannot; nothing was changed — checked before the rename, so a refused move never follows a rename that went through | `test_mail.py::test_a_folder_is_not_moved_under_itself` |

### `list_messages(kind=messages, folder?, query?, from?, to?, since?, until?, unread_only?, has_attachments?, flagged?, importance?, limit, cursor)`
Envelopes only: id, from, to/cc count, subject, received time, folder, read, flag, categories,
has attachments, a 200-character preview. No folder searches all mail folders. Newest first.
`kind=folders` returns the folder tree instead, as paths with total and unread counts — called when
the person names a folder the server did not find, or asks about folders.
`$search` ignores `Prefer: IdType="ImmutableId"`, so search results are translated to immutable ids
(`translateExchangeIds`) before they are returned (rule 34). A backslash in `query` is dropped: inside
the quoted `$search` Graph reads it as an escape and answers 400 (measured); quotes and parentheses
pass. Each envelope names its folder; `unread` is present when a message is unread, and absent
means read, as every false flag is absent.

| Must refuse | The person is told | Test |
|---|---|---|
| `query` combined with ordering Graph cannot do with `$search` | handled inside: search results are ranked by Graph; the result says so | `test_mail.py::test_search_uses_kql_and_no_orderby` |
| `since` after `until` | the window is empty | `test_mail.py::test_empty_window_is_refused` |
| A `from`/`to` that resolves to several people | candidates | `test_mail.py::test_an_ambiguous_sender_lists_candidates`, `test_mail.py::test_an_ambiguous_recipient_filter_lists_candidates` |
| A cursor beside filters other than the ones it was made with (given alone, it keeps its own) | start without the cursor, or give it alone | `test_mail.py::test_a_message_cursor_beside_other_filters_is_refused` |

### `get_message(id, include_body=true, include_attachments=false, include_thread=false, max_chars=20000)`
One message. The body arrives as text inside a block marked as message content written by someone
else (rule 25), reduced from HTML (rule 27), `uniqueBody` when part of a thread. Exchange rebuilds
`uniqueBody` and drops styles that hid text in the full body (`font-size:0`, the `hidden` attribute),
and its text nodes do not line up with the full body's: when the full body hides any text, the full
body's own text is shown instead (what it hides dropped, the thread included, `thread_cut` absent).
An empty unique part falls back to the full body too. With
`include_attachments`: name, type, size and a download link per attachment (300 s); attached
messages download as `.eml`; cloud attachments return their file link. An attachment's `size` is
Graph's, which counts the attachment's own properties too (measured: a 10-byte file listed as 180,
10,485,760 bytes as 10,485,891); the download is the exact file. With `include_thread`: the
envelopes of the rest of the conversation. `max_chars` below 1 is refused; any value from 1 is honoured.

| Must refuse | The person is told | Test |
|---|---|---|
| `max_chars` below 1 | the range | `test_mail.py::test_max_chars_below_one_is_refused` |

### `update_messages(ids[≤50], action, folder?, categories?)`
`action`: `move`, `mark_read`, `mark_unread`, `flag`, `unflag`, `complete_flag`, `categorize`,
`delete`. `delete` moves to Deleted Items. Every action is reversible with another call.

| Must refuse | The person is told | Test |
|---|---|---|
| More than 50 ids | split the request | `test_mail.py::test_more_than_50_ids_is_refused` |
| `move` without a folder, or to a folder that does not resolve | candidates | `test_mail.py::test_move_without_folder_is_refused`, `test_mail.py::test_move_to_a_folder_that_does_not_resolve_is_refused` |
| An id that is not a message the person owns | not found (rule 16), in that id's `failed` entry | `test_mail.py::test_a_message_the_person_cannot_see_is_not_found` |

### `save_draft(id?, mode=new, source_id?, to, cc, bcc, subject, body, body_format=text, body_upload=false, attachments_upload=0, remove_attachments[], importance?)`
`mode`: `new`, `reply`, `reply_all`, `forward`. Replies and forwards are created by Graph
(`createReply`, `createReplyAll`, `createForward`) from `source_id`, which produces a draft holding
the quoted history; the new text is then placed above that history and the history is kept byte for
byte (rule 31). Each draft's new text sits in a part whose id is drawn for that draft and kept in its
mark: a letter written here and now quoted carries a part of its own, and taking that for ours
replaced the quoted letter with the reply (live finding). Graph refuses `comment` and `message.body` together, so only one is ever sent. With `id`, updates a draft this server created. Never sends.

Returns: draft id, resolved recipients, subject, a 500-character preview, attachments, the HTML
report `{sanitized[], warnings[], size}`, a body upload link when `body_upload`, and N attachment
upload links when `attachments_upload=N` (180 s each).

| Must refuse | The person is told | Test |
|---|---|---|
| `id` of a draft the person wrote themselves | only drafts made here are changed | `test_mail.py::test_changing_a_draft_the_person_wrote_is_refused` |
| More than 20 recipients in total | the count | `test_mail.py::test_more_than_20_recipients_is_refused` |
| A distribution list or group address among recipients | which one | `test_mail.py::test_distribution_list_is_refused` |
| A body uploaded to a draft already sent | it was already sent; the upload link answers 422 | `test_link_routes_mail.py::test_a_body_cannot_be_uploaded_to_a_draft_already_sent` |
| An attachment uploaded to a draft already sent, being sent, or no longer a draft (it would go out unseen, or land on the copy in Sent Items) | it was already sent, or is not a draft; the upload link answers 422 | `test_link_routes_mail.py::test_an_attachment_cannot_be_uploaded_to_a_draft_already_sent` |
| `reply`/`forward` without `source_id` | — | `test_mail.py::test_reply_without_source_is_refused` |
| `body_format=html` inline over 20,000 characters | use `body_upload` | `test_mail.py::test_long_inline_html_must_be_uploaded` |
| Uploaded HTML not UTF-8, not `text/html`, or over 2 MB | the limit | `test_link_routes_mail.py::test_draft_body_not_utf8_is_rejected`, `test_link_routes_mail.py::test_draft_body_wrong_content_type_is_415`, `test_link_routes_mail.py::test_draft_body_over_the_limit_is_413` |
| HTML with relative `src`/`href` | which URLs | `test_mail.py::test_relative_urls_refuse_the_draft` |
| An attachment that would take the message past the tenant's size limit | that it is larger than Microsoft 365 accepts; the numbers are not yet reported (Graph does not return the limit) | `test_link_routes_mail.py::test_an_attachment_past_the_message_size_limit_is_refused` |
| An attachment over 150 MB | the upload answers 413 | `test_link_routes_mail.py::test_an_attachment_over_the_link_limit_is_413`, `test_link_routes_mail.py::test_draft_attachment_links_carry_the_150_mb_limit` |

Must transform, not refuse (and report in `sanitized`): `script`, `iframe`, `object`, `embed`,
`form`, `on*` attributes, `javascript:` URLs removed; `data:` images turned into inline attachments
referenced by `cid:`; `<style>` inlined. Must warn: flex, grid, `position`, background images, web
fonts (`@font-face`, `@import`, a stylesheet `<link>`), images without `alt`, width over 800 px.
Attachments are listed whatever `hasAttachments` says: Graph sets it false when every attachment is
inline.

### `send_draft(id, user_confirmed)`
Sends one draft this server created. Re-reads it and re-checks recipients first. Irreversible.

| Must refuse | The person is told | Test |
|---|---|---|
| `user_confirmed` not true | ask the person | `test_mail.py::test_send_without_confirmation_sends_nothing` |
| Draft not created here, or no longer a draft | — | `test_mail.py::test_send_refuses_a_draft_the_person_wrote`, `test_mail.py::test_send_refuses_a_message_that_is_no_longer_a_draft` |
| A draft already sent: Exchange keeps it in Drafts, still `isDraft`, for a few seconds, and a second call sent a second copy (live finding) | it was already sent; nothing is sent again. The mark is set to `sent` before the send, against the draft's etag, and restored when Graph refuses the send. Two calls at once (measured): the loser's mark answers 412, or both marks go through and Exchange answers the second send 403 `ErrorAccessDenied` while the draft still says `isDraft` for a few seconds; both are this case, not a conflict or a policy; a 403 that does not settle within a few seconds keeps the mark and says to check Sent Items | `test_mail.py::test_send_refuses_a_draft_already_sent`, `test_mail.py::test_a_second_send_of_one_draft_sends_nothing`, `test_mail.py::test_the_loser_of_two_marks_at_once_is_told_it_was_already_sent`, `test_mail.py::test_the_loser_whose_send_exchange_refuses_is_told_it_was_already_sent`, `test_mail.py::test_a_send_refused_while_the_draft_never_settles_keeps_it_marked_sent` |
| Recipients changed to over 20 or to a list since the draft was saved | — | `test_mail.py::test_send_refuses_a_list_added_after_the_draft`, `test_mail.py::test_send_refuses_recipients_grown_past_20` |
| Bucket empty | the wait; mass mailing is out of scope | `test_mail.py::test_bucket_stops_the_eleventh_send` |
| A recipient outside `M365_ALLOWED_RECIPIENT_DOMAINS`, when it is set | the domains outside it, never the addresses | `test_recipient_domains.py::test_send_draft_refuses_a_recipient_outside_the_allowed_domains` |
| An upload link `save_draft` handed out for this draft, used once the draft was sent (a body or file the person never saw). Between the sent mark and the send the route refuses on the mark; the links are spent only after a send that went through, so a failed send leaves them usable | the link is spent: 403, ask for a new one | `test_link_routes_mail.py::test_sending_a_draft_revokes_its_upload_links` |
| An upload into this draft (body or file) still running: it passed its check of the draft before the send and would add what the person never saw. Checked before the sent mark, and again after it, when the mark is taken back | an upload to this draft is still running; wait, then call `send_draft` again; nothing was sent | `test_link_routes_mail.py::test_send_draft_refuses_while_an_upload_to_the_draft_is_running`, `test_link_routes_mail.py::test_an_upload_starting_while_the_draft_is_marked_sent_stops_the_send` |

### `save_automatic_reply(status, internal_message?, external_message?, audience?, start?, end?, user_confirmed)`
`status`: `off`, `on`, `scheduled`. `audience` for people outside the organisation: `none`,
`contacts`, `all`. Left out, it is `contacts`, or `none` where `M365_ALLOWED_RECIPIENT_DOMAINS` is
set: a reply nobody asked to send outside stays inside, and an outside audience asked for is refused. The current setting is shown by `get_connection`. Turning it on or changing its
text needs `user_confirmed`, because everyone who writes receives it; turning it off does not.
Messages pass the outgoing HTML pipeline: plain text becomes paragraphs, text holding HTML is
sanitised like a letter and the result reports what was removed.

| Must refuse | The person is told | Test |
|---|---|---|
| `scheduled` without both dates, or an end in the past | — | `test_autoreply.py::test_scheduled_without_both_dates_is_refused`, `test_autoreply.py::test_an_end_in_the_past_is_refused` |
| `on`/`scheduled` without `user_confirmed` | ask the person | `test_autoreply.py::test_turning_it_on_without_confirmation_writes_nothing`, `test_autoreply.py::test_scheduling_without_confirmation_writes_nothing` |
| `on`/`scheduled` with an outside audience (`contacts` or `all`) while `M365_ALLOWED_RECIPIENT_DOMAINS` is set: whoever writes from any domain gets it | the allowed domains; use `audience=none`; nothing was changed | `test_recipient_domains.py::test_an_automatic_reply_to_people_outside_is_refused_when_domains_are_limited` |

## Calendar and scheduling

Scheduling is a first-class area: finding when people are free is the most common calendar
question after "what do I have today".

### `list_events(start, end, query?, calendar?, include_attendees=false, limit, cursor)`
Events in a window of at most 62 days, recurring series expanded into occurrences (`calendarView`).
Each: id, subject, start, end, duration, location, online meeting link, organiser, my response,
whether I organise it, attendee count; attendees with their responses on request. A colleague's
calendar the person has been given access to can be named in `calendar`.

### `get_event(id, include_attendees=true, include_body=false, include_transcript=false, include_recording=false)`
One event. `include_recording` adds the Teams meeting's recordings and transcript files as download
links (see [Meeting recordings and transcripts](#meeting-recordings-and-transcripts)); it belongs to
the `recordings` area, so with that area off the parameter does not exist. `include_transcript` returns the Teams meeting transcript as text (speaker-attributed
VTT reduced to `speaker: text` lines, capped) when one exists and the person may read it; otherwise
says why not. It needs `OnlineMeetingTranscript.Read.All` (administrator consent) and is behind a
switch. Tenant policy can disable it (`GraphAccessToTranscriptsDisabled`) or speaker attribution
(`SpeakerAttributionNotAllowed` → retried without speakers); each is reported in plain words.

### `find_meeting_times(people[≤20], start?, end?, duration_minutes?, show_busy=false, interval_minutes=30, include_presence=false, working_hours_only=true, required_all=true, rooms=false, limit=5)`
One question — when are these people free — answered at the depth asked. The window defaults to the
next 7 days and is at most 62.

- **Without `duration_minutes`:** free/busy of each person (and room): busy blocks with
  free/tentative/busy/out-of-office status, a one-character-per-slot availability string, their
  working hours and time zone. Only what each person's calendar sharing exposes — usually status and
  times, not subjects. `include_presence` adds each person's current Teams presence (available,
  busy, in a meeting, away). Graph `getSchedule` (`Calendars.ReadBasic`) and `getPresencesByUserId`
  (`Presence.Read.All`).
- **With `duration_minutes`:** suggested slots where everyone (or, with `required_all=false`, at least half) is free,
  ranked, with who is unavailable in each (`not_free`) and who has only accepted tentatively
  (`tentative`; Graph counts them as available, so its "all attendees are available" reason is
  dropped then), and a room when `rooms` (`rooms_note` when none was suggested). The scheduler's
  answer to "find an hour with Anna and Boris next week". `show_busy` (or `include_presence`) adds
  the free/busy view beside the slots. Graph `findMeetingTimes` (`Calendars.Read.Shared`); rooms need
  `Place.Read.All`, which an administrator grants, so `rooms` is behind a switch. Measured:
  `findMeetingTimes` never offers a time when a named person is busy, whatever
  `minimumAttendeePercentage` says (50 and 0 alike, attendees required or optional); with
  `required_all=false` the times when at least half are free come from `getSchedule` at 15-minute
  resolution, on the hour and the half hour, after Graph's own, with `confidence` the share free;
  like Graph's, none leaves out the person themselves. A calendar that cannot be read is listed in
  `unknown`, never as `not_free`; rooms are not looked up for these times, and `rooms_note` says so.
- **`working_hours_only`** (the default) cuts the window to the person's working hours from Outlook
  (`mailboxSettings/workingHours`), or to Monday–Friday 09:00–17:00 in the mailbox time zone when none
  are set, and the result names them in `working_hours`. Exchange treats every hour as working when
  none are set, and suggested nights and Saturdays (measured); attendees' own working hours still apply
  through `activityDomain=work`. A window with no working time long enough asks Graph nothing and says
  so. Free/busy rows say `working_hours: not set in Outlook` when a person has none.

Nothing is booked or sent: both Graph calls are POSTs that only read.

| Must refuse | The person is told | Test |
|---|---|---|
| More than 20 people (nobody is sent anything, so not "recipients"), or a distribution list | — | `test_calendar.py::test_find_meeting_times_more_than_20_people_or_a_group_is_refused` |
| A window over 62 days, or an end before the start | narrow it | `test_calendar.py::test_find_meeting_times_window_over_62_days_is_refused` |
| Outlook working hours in a time zone the server cannot place (they would be read as UTC) | call with `working_hours_only=false` | `test_calendar.py::test_working_hours_in_a_zone_the_server_cannot_place_are_not_shifted_to_utc` |

### `save_event(id?, subject, start, duration_minutes, attendees[], location?, room?, online_meeting=true, body?, reminder_minutes?, show_as?, user_confirmed)`
Create or change an event. Without attendees it stays in the person's calendar and needs no
confirmation. With attendees, or when changing one that has them, Graph notifies them at once:
`user_confirmed` is required and the bucket is drawn from.

`online_meeting` asks Graph for `isOnlineMeeting` with `onlineMeetingProvider=teamsForBusiness`.
Measured: a calendar whose `allowedOnlineMeetingProviders` is empty saves a plain event (201,
`isOnlineMeeting: false`), and where Teams meetings are offered the join link appears after the create,
from 1 s to over a minute later. The tool re-reads the event for up to 10 s; without a link,
`online_meeting_note` says which of the two it is. `online_meeting=false` on an event that has a link
is saved and the link stays; `online_meeting_note` says so. Attendee counts leave the organiser out on every
copy (an attendee's copy lists the organiser among the attendees; the organiser's does not).

| Must refuse | The person is told | Test |
|---|---|---|
| Changing an event the person does not organise | you are not the organiser; respond instead | `test_calendar.py::test_changing_an_event_someone_else_organises_is_refused` |
| Notifying attendees without `user_confirmed` | ask the person | `test_calendar.py::test_inviting_without_confirmation_writes_nothing` |
| More than 20 attendees, or a distribution list | — | `test_calendar.py::test_more_than_20_attendees_are_refused`, `test_calendar.py::test_a_distribution_list_is_refused` |
| End before start, duration 0 | — | `test_calendar.py::test_zero_or_negative_duration_is_refused` |
| A negative `reminder_minutes` | minutes before the start, so 0 or more | `test_calendar.py::test_a_negative_reminder_is_refused` |
| Recurring series rules | not supported in this version; one occurrence can be changed | `test_calendar.py::test_a_recurring_series_cannot_be_changed` |
| An attendee outside `M365_ALLOWED_RECIPIENT_DOMAINS`, when it is set | the domains outside it, never the addresses | `test_recipient_domains.py::test_inviting_an_attendee_outside_the_allowed_domains_is_refused` |

### `respond_to_event(id, response, comment?, send_response=true)`
`response`: `accept`, `tentative`, `decline`. The answer alone is one's own record: no confirmation,
but the bucket is drawn from when a response is sent. A `comment` is text the person writes to the
organiser, so it is a message like any other: it needs `user_confirmed`, is at most 1,000
characters, and the organiser must be inside `M365_ALLOWED_RECIPIENT_DOMAINS` when that is set. With `send_response=false` no comment is sent: Graph answers 400
("'SendResponse' must be true when 'Comment' is not null") even for an empty one (measured). No message
reaches the organiser then, but inside one organisation their copy of the meeting shows the answer
within seconds; the result says `organizer_told: false` and that the calendar may still
show it.

| Must refuse | The person is told | Test |
|---|---|---|
| A cancelled event | the event was cancelled — never "done" | `test_calendar.py::test_responding_to_a_cancelled_event_is_refused` |
| An event the person organises | organisers do not respond | `test_calendar.py::test_the_organiser_does_not_respond` |
| A `comment` with `send_response=false` | the comment goes to the organiser; leave it out to answer quietly | `test_calendar.py::test_a_comment_without_sending_the_response_is_refused` |
| A `comment` without `user_confirmed` | ask the person; nothing was answered | `test_calendar.py::test_a_comment_to_the_organiser_needs_confirmation` |
| A `comment` over 1,000 characters | the limit and the length; nothing was answered | `test_calendar.py::test_a_comment_longer_than_the_cap_is_refused` |
| A `comment` to an organiser outside `M365_ALLOWED_RECIPIENT_DOMAINS`, when it is set | the domains outside it, never the addresses | `test_recipient_domains.py::test_a_comment_to_an_organiser_outside_the_allowed_domains_is_refused` |

### `cancel_event(id, comment?, user_confirmed)`
Cancels an event the person organises; attendees get the cancellation at once. Irreversible.

| Must refuse | The person is told | Test |
|---|---|---|
| A `comment` while an attendee is outside `M365_ALLOWED_RECIPIENT_DOMAINS`, when it is set (the comment goes to every attendee; the cancellation without one is not refused, or the meeting would stay standing) | the domains outside it, never the addresses | `test_recipient_domains.py::test_a_cancellation_comment_to_an_attendee_outside_the_allowed_domains_is_refused` |

## Meeting recordings and transcripts

Recordings are what people reuse most after a meeting: watch again, cut a clip, take screenshots
for slides, summarise. The video is downloaded into the agent's sandbox, where it can be cut and
processed; it never passes through the model.

### `get_event(id, include_recording=true)`
For one Teams meeting: its recordings (start, duration, size) and transcripts, each with a
download link — video as MP4, transcript as VTT with speakers (or plain text where the
organisation does not allow speaker attribution). The meeting is found from the calendar event's
join URL. There is no separate recording tool: the recording is a depth of the event, as its
transcript text is.

Graph allows downloading a recording only to the **organiser** under delegated permissions;
participants may watch the copy shared with them in Teams, but not download it, unless the
organisation's administrator has allowed it. The tool tries, in order, the meeting's recording API
and the recording file in the organiser's OneDrive or the channel's SharePoint (when shared with
the person), and when neither allows download it says exactly that, with the link to watch it in
the browser. Transcripts are readable by participants. Needs `OnlineMeetingRecording.Read.All` and
`OnlineMeetingTranscript.Read.All` (administrator consent).

| Must refuse | The person is told | Test |
|---|---|---|
| `include_recording` on an event that is not a Teams meeting | — | `test_recordings.py::test_an_event_that_is_not_a_teams_meeting_is_refused` |
| A meeting with no recording or transcript yet | says it is not available yet, not "not found" | `test_recordings.py::test_no_recording_yet_is_a_sentence_not_an_error` |
| A participant without download rights | explains organiser-only download and returns the watch link | `test_recordings.py::test_a_participant_without_download_rights_is_told_and_given_the_meeting_link` |
| A meeting whose recording has expired | says it expired | `test_recordings.py::test_an_expired_recording_says_it_expired` |

Videos are large (about 400 MB per hour): the link streams with `Range` support, and the
response says the size so the agent can plan.

## Files: OneDrive and SharePoint

### `list_files(location=my, site?, folder?, query?, kind?, include_lists=false, modified_since?, limit, cursor)`
`location`: `my` (own OneDrive), `shared` (shared with me), `recent`, `site` (a SharePoint site
library). `shared` must not use `/me/drive/sharedWithMe`, which stops returning data in November
2026; it goes through Microsoft Search (`driveItem`). With `query` and no folder, searches across everything the person can reach. Each item:
id, name, kind, size, modified time and by whom, path, web link. `kind`: `file`, `folder`,
`document`, `spreadsheet`, `presentation`, `pdf`, `image` keep only that kind of item; `site` lists
SharePoint sites instead of files.

`kind=site`: the SharePoint sites the person follows and the sites of their teams (`/me/joinedTeams` →
`/groups/{id}/sites/root`), or those whose name matches `query`, with their document libraries and,
with `include_lists`, their lists. Measured: `/sites?search=` left a team's own site out, and answered
500 on a tenant whose search was not set up, so names are also matched against the followed and team
sites, and a site named in `site` is looked up there first.

When Microsoft Search answers 5xx (measured on the test tenant for every `driveItem` query and for the
drive's own `search(q=)`), a query on `my` or `site` compares the names in the one folder asked
(up to 200 items) and says so in `note`; those matches continue by cursor like any other page. SharePoint is the organisation's intranet:
team and project sites, each with document libraries, lists and pages. A library's id is the drive id
the other file tools accept.

| Must refuse | The person is told | Test |
|---|---|---|
| `location=site` without a site | find one with `kind=site` | `test_files.py::test_site_location_without_a_site_is_refused` |
| A `folder` that is a file (Graph: 422 `notSupported` for its children, 500 for its search) | that it is a file; give its folder, or read it with `get_file` | `test_files.py::test_a_folder_that_is_a_file_is_said_to_be_a_file` |
| `location=shared` while Microsoft Search does not answer | that shared files are found only through it | `test_files.py::test_shared_files_while_microsoft_search_is_down_is_said_plainly` |
| A site name that matches no followed or team site while Microsoft Search does not answer (not "no site matches": that is not known) | give the site's address | `test_files.py::test_a_site_name_while_search_is_down_is_not_called_missing` |

### `get_file(id | path, include_versions=false, include_permissions=false)`
Metadata, path, web link to open in the browser; versions; who has access and how.

| Must refuse | The person is told | Test |
|---|---|---|
| A path holding `#` (Graph cannot take it in a path safely; the id works) | use the id from `list_files` | `test_files.py::test_a_path_with_a_hash_is_refused_with_the_way_to_reach_it` |

### `download_file(id, format=original)`
A 300 s download link and a ready `curl` command. Graph answers `/content` with a 302 to a
pre-authenticated URL; our link streams through the server so that URL never reaches the model. The
link asks for `Accept-Encoding: identity` and decodes an encoded answer, so the saved file is the file:
OneDrive gzips text for a client that accepts it, and plain `curl` saved the gzip bytes (measured).
`format=pdf` converts the source types Graph documents (doc, docx, dot, dotm, dotx, eml, epub, htm,
html, markdown, md, msg, odp, ods, odt, pps, ppsx, ppt, pptx, rtf, tif, tiff, xls, xlsm, xlsx).
Measured: csv, txt, json and xml answer 406; md, html, rtf and eml convert. The source size limit is
a live finding to verify.

| Must refuse | The person is told | Test |
|---|---|---|
| A folder | list it instead | `test_files.py::test_download_of_a_folder_is_refused` |
| `format=pdf` for a type Graph cannot convert | the supported types | `test_files.py::test_pdf_of_an_unsupported_type_is_refused_with_the_supported_ones` |

### `upload_file(folder, name, on_conflict=fail)`
A 180 s upload link. `on_conflict`: `fail`, `replace` (the old content stays in version history),
`rename`. Up to 4 MB in one `PUT`; larger through an upload session in fragments that are multiples
of 320 KiB and under 60 MiB each, sent in order. Links stay usable until they expire (180 s), not for
one transfer only; rule 32 asks short life and fail-closed checks, which they have.

| Must refuse | The person is told | Test |
|---|---|---|
| A name holding a character OneDrive refuses (backslash, slash, colon, star, question mark, quote, angle brackets, pipe) or `#`, or only dots, in `upload_file` and in `update_file` (`rename`, `create_folder`) | the characters | `test_files.py::test_a_name_with_a_hash_is_refused_everywhere` |
| A name OneDrive reserves: `CON`, `PRN`, `AUX`, `NUL`, `COM0`–`COM9`, `LPT0`–`LPT9` with any extension, a trailing dot, `_vti_` (measured: 400 or 403 at the `PUT`, after the link was handed out) | the rule; nothing is created | `test_files.py::test_a_name_onedrive_reserves_is_refused_before_a_link_is_made` |

### `update_file(ids[≤20], action, target_folder?, new_name?, version?, user_confirmed?)`
`action`: `move`, `copy`, `rename`, `create_folder`, `delete` (to the recycle bin),
`restore_version`. Deleting a folder needs `user_confirmed`.

Graph cannot restore from the recycle bin for work and school accounts (`restore` is OneDrive
Personal only), so `delete` says where the person restores it — the recycle bin in OneDrive or the
site — and is still reversible, by the person. Copy is asynchronous: the tool waits on the monitor
URL for up to 30 s and otherwise returns "in progress" with the target path. A copy the monitor reports
`failed` (measured: a name that exists in the target answers 202, then `failed` with
`nameAlreadyExists`) is a failed item with its reason. Some items failing is a result with `failed`;
all of them failing is an error.

| Must refuse | The person is told | Test |
|---|---|---|
| `create_folder` with a name that exists in that folder (Graph would make "name 1") | the existing path | `test_files.py::test_a_folder_that_exists_is_refused_with_its_path` |
| Every item failing, a copy failing at its monitor included | nothing was changed, with each item's reason | `test_files.py::test_a_change_that_failed_for_every_item_is_an_error`, `test_files.py::test_a_copy_that_failed_at_the_monitor_is_an_error_with_its_reason` |
| A folder moved or copied into itself or a folder inside it (Graph: a bare 400) | why, for that item | `test_files.py::test_a_folder_into_its_own_subfolder_is_refused_with_why` |

### `share_file(id, action, people[]?, role=view, message?, user_confirmed)`
`action`: `link` (organisation-wide link), `invite` (specific people, notified), `revoke`.
`role`: `view`, `edit`. Anonymous links do not exist (rule 24). `revoke` removes named people's direct
access; an organisation link, once made, is removed in OneDrive or SharePoint ("Manage access"), and
access inherited from a parent folder is revoked there.

| Must refuse | The person is told | Test |
|---|---|---|
| `invite` or `link` without `user_confirmed` | ask the person | `test_files.py::test_share_without_confirmation_writes_nothing` |
| More than 20 people | — | `test_files.py::test_invite_more_than_20_people_is_refused` |
| `revoke` of the owner's own access | — | `test_files.py::test_revoke_of_the_owner_is_refused` |
| Inviting someone outside `M365_ALLOWED_RECIPIENT_DOMAINS`, when it is set | the domains outside it, never the addresses | `test_recipient_domains.py::test_sharing_with_someone_outside_the_allowed_domains_is_refused` |

## SharePoint lists

Lists are the registers and trackers that live on SharePoint sites: requests, assets, contracts,
issue logs. `list_files(kind=site, include_lists=true)` shows the lists of a site.

### `list_list_items(site, list, where?, fields[]?, order_by?, include_columns=false, limit, cursor)`
Rows of one list with the chosen fields. `where` is a small filter language
(`Status = 'Open' and Amount > 600`) translated to Graph `$filter` on `fields/…`; filtering works
reliably on indexed columns only, so a filter on a non-indexed column in a large list returns a
clear error suggesting a narrower filter. `include_columns` returns the column names, types and
choices — what the model needs before writing. Measured: the fields are expanded as
`fields($select=…)`; Graph refuses `fields(select=…)` with 400, and so no row of any list could be read.
A filter or sort Graph refuses with 400 names both causes it shares that code for: a value that does
not fit its column (`Created gt 'yesterday'`), and a non-indexed column on a large list.

| Must refuse | The person is told | Test |
|---|---|---|
| A text longer than its column's `maxLength` (QA: 300 characters in `Title` came back as a bare refusal) | the limit and the length | `test_lists.py::test_a_text_longer_than_its_column_is_refused_with_the_limit` |

### `save_list_item(site, list, id?, fields)`
Create or update one row. Fields are checked against the list's columns (names, types, choice
values, a text's `maxLength`) before the call; unknown fields are refused with the column list. Updates send `If-Match`
when the row was read in the same conversation; a conflict says the row changed.

### `delete_list_item(site, list, id, user_confirmed)`
Deletes one row. Whether it goes to the site's recycle bin is a live finding; until settled it needs
`user_confirmed`.

## Excel

Read and write workbooks that live in OneDrive or SharePoint, in place, without downloading them.
Only `.xlsx`/`.xlsm` (Graph does not support `.xls`).

### `get_range(file, sheet?, address? | table?, include_ranges=false, max_cells=2000)`
Without `sheet`, `address` or `table`: the workbook's structure — sheets with their used ranges,
tables with columns and row counts (from the table's `dataBodyRange`: `rows/$count` answers 400,
measured), named ranges; with `include_ranges`, the first 20 rows of each table. With them: values (and formulas on request) of a range, a sheet's used range or a table, as
rows. Bigger ranges are refused with their size; the agent downloads the file for bulk work instead.
Rows, formulas and table previews come inside a marked content block, one JSON array per row
(rule 25: cell text is someone else's).

### `save_range(file, sheet, address? | table?, values[][], action=write)`
`action`: `write` (values into the range; `null` leaves a cell unchanged), `append_rows` (to a
table). Uses a persistent workbook session for multi-call edits and closes it. Formulas are
written as given. `written_cells` counts the cells given a value, not the `null` ones.

| Must refuse | The person is told | Test |
|---|---|---|
| `.xls` or a non-workbook | — | `test_excel.py::test_non_workbooks_and_xls_are_refused` |
| `values` shape not matching the address | the expected shape | `test_excel.py::test_shape_not_matching_the_address_is_refused_with_the_expected_shape` |
| More than 5,000 cells in one write | split it | `test_excel.py::test_more_than_5000_cells_is_refused` |
| A text over 32,767 characters in one cell (Excel's limit; Graph cuts it and answers success, measured) | the limit | `test_excel.py::test_a_cell_longer_than_excel_holds_is_refused` |
| A read over `max_cells` | the range's size; download the file instead | `test_excel.py::test_get_range_over_max_cells_is_refused_with_its_size` |
| An address outside A1:XFD1048576 (Graph: a bare 400 `InvalidArgument`) | the sheet's bounds | `test_excel.py::test_an_address_outside_a_sheet_is_refused_with_the_bounds` |
| A sheet or table the workbook does not have (Graph: 404, which read as if the file were missing) | the sheets or tables it has | `test_excel.py::test_a_missing_sheet_or_table_is_named_with_the_real_ones` |
| `write` with a table, or `append_rows` with an address (QA: the table was ignored and A1 overwritten) | which argument belongs to which action | `test_excel.py::test_an_address_and_a_table_together_are_refused` |
| A formula Excel refuses (`=SUM((`: 400 `InvalidArgument`, measured) | that a formula is not valid; nothing was written | `test_excel.py::test_a_formula_excel_refuses_is_named` |

## Teams

### `list_conversations(kind=chat, team?, query?, unread_only=false, limit, cursor)`
`kind`: `chat`, `channel`, `all`.

- **Chats**, latest first: id, topic or the other person's name, kind (one-to-one, group, meeting),
  last message preview and time, unread (from `viewpoint.lastMessageReadDateTime`), members count.
  With `query`, searches messages across chats through Microsoft Search (`chatMessage`) and returns
  hits: chat, author, time, snippet.
- **Channels**: teams the person has joined and their channels (`Team.ReadBasic.All`,
  `Channel.ReadBasic.All`), narrowed by `team` and by `query` on team and channel names.
- **All**: the chats page, and on the first page the teams with their channels beside it.

| Must refuse | The person is told | Test |
|---|---|---|
| `unread_only` with `kind=channel`, or `team` with `kind=chat` | which kind takes it | `test_teams.py::test_list_conversations_refuses_options_of_the_other_kind` |
| A `query` beside a cursor from the plain list or from a search for other words | start the search without the cursor | `test_teams.py::test_a_cursor_from_another_search_or_from_the_plain_list_is_refused` |
| A cursor beside arguments it was not made with: the chat list's with `kind=channel`, a search's with `unread_only`, the channel list's for another team or name (given alone, a cursor continues its own list, as in `list_messages`) | the cursor continues another list; start without it | `test_teams.py::test_a_cursor_is_refused_beside_arguments_it_was_not_made_with` |

Graph hands out a next link on the last page of chats, and behind it come empty pages that each link
on, without end (measured 2026-10-23): an empty page ends the list, and the tool reads ahead before
saying `has_more`. A search's `moreResultsAvailable` stays true after its last hit, with `total`
far above the hits it returns (measured: 21 hits, `total` 65): the tool reads one hit ahead instead.
Snippets come inside the content block, as message text does. The channel list continues by
`cursor` too; with `kind=all` the rest of the teams continue from `teams_cursor` with `kind=channel`.

### `get_conversation(chat | person | team + channel, messages=20, thread?, before?, cursor?, include_members=false, attachments_only=false)`
One chat or one channel with the depth asked. Messages: author, time, text (reduced from Teams HTML,
inside the content block), mentions, attachments, reactions count, reply-to. A channel returns its
posts, or one thread's replies with `thread`. Graph cannot filter or sort channel messages and returns
them by the last activity in each thread. Reading channels needs `ChannelMessage.Read.All`, which an
administrator grants, so `team`, `channel` and `thread` are behind the `teams_channels` switch;
reading chats and posting do not need it.

Every answer carries `has_more`. A chat continues with `before=next_before` (one message more than
asked is read, so the end is known); a channel or a thread continues with `cursor`, which re-reads the page it stopped on and skips the posts already returned (posts move as activity reorders them). `messages` above
50 is clamped and `messages_clamped_to` says so. A post's `replies` counts the replies still there.

| Must refuse | The person is told | Test |
|---|---|---|
| None of chat, person, or team and channel; or more than one of them | name one conversation | `test_teams.py::test_get_conversation_needs_exactly_one_conversation` |
| A channel named without its team, or `thread` without a channel | — | `test_teams.py::test_get_conversation_thread_and_channel_need_their_team` |
| `include_members` on a channel | members are listed for chats only | `test_teams.py::test_get_conversation_members_are_for_chats_only` |
| `messages` below 1 | the range | `test_teams.py::test_messages_out_of_range_are_refused_below_and_clamped_above_with_a_note` |
| A cursor of another conversation, or a cursor beside `before` | — | `test_teams.py::test_a_conversation_cursor_is_refused_elsewhere_and_beside_before` |
| `person` naming the person themselves | that a one-to-one chat is with another person | `test_teams.py::test_a_one_to_one_chat_with_oneself_is_refused_with_the_reason` |

### `send_message(to, body, body_format=text, reply_to?, mentions[]?, file_ids[]?, user_confirmed)`
`to`: a chat id, one or more people (the one-to-one or group chat is found or created), or
`team/channel`. The 20-recipient limit counts people named; a channel post reaches every member of
the channel and counts as one send, which the description says. `reply_to` replies in a channel thread; Graph has no threaded reply in chats.
Mentions become `<at>` tags with the resolved users. Files are uploaded to the person's OneDrive
first, with `upload_file`, and attached by reference through `file_ids`. There is no
`attachments_upload`: a Teams message has no draft, so an upload link would have to post the message
when the last file arrives — a send outside the confirmed call, of a message the person did not see
whole (rules 20, 21). Images embedded in `html` as `data:` URIs go as the message's hosted contents
(`@microsoft.graph.temporaryId`, at most 3 MB each, the outgoing pipeline's limit; Teams takes 4 MB).
Teams HTML is a restricted subset: the body passes the same sanitiser with Teams' allowlist.
Irreversible from here: the person can edit or delete it in Teams.

| Must refuse | The person is told | Test |
|---|---|---|
| `user_confirmed` not true | ask the person | `test_teams.py::test_send_without_confirmation_writes_nothing` |
| A channel the person cannot post to | Microsoft 365 does not let you do this here | `test_teams.py::test_a_channel_the_person_cannot_post_to_is_refused` |
| A person Teams cannot find in the tenant (Graph `401` `TenantNotFound` on creating the chat) | that Teams cannot reach them and mail can — never "reconnect" | `test_refusals.py::test_a_chat_with_someone_teams_cannot_reach_is_not_a_sign_in_problem` |
| Bucket empty | the wait | `test_teams.py::test_an_empty_bucket_stops_the_message` |
| People, or members of the chat named by id or of the channel, outside `M365_ALLOWED_RECIPIENT_DOMAINS`, when it is set | the domains outside it, never the addresses | `test_recipient_domains.py::test_a_teams_message_to_someone_outside_the_allowed_domains_is_refused`, `test_recipient_domains.py::test_a_message_to_an_existing_chat_with_an_outside_member_is_refused`, `test_recipient_domains.py::test_a_channel_post_with_a_member_outside_the_allowed_domains_is_refused` |
| A channel post while `M365_ALLOWED_RECIPIENT_DOMAINS` is set and `ChannelMember.Read.All` is not granted, so the channel's members cannot be checked | the permission an administrator must grant; nothing was sent | `test_recipient_domains.py::test_a_channel_post_whose_members_cannot_be_read_is_refused_when_domains_are_limited` |

Where `M365_ALLOWED_RECIPIENT_DOMAINS` is set, a message to a chat by id or to a channel first reads
its members, every page up to ten (`MEMBER_PAGES`), and checks each address. Two consequences the
description states: a channel post then needs `ChannelMember.Read.All`, an administrator-consented
permission asked for only in that case; and a chat or channel with members past the tenth page is
refused, because who is in it cannot be checked. `update_message` makes the same check for `edit`
and `restore`.

### `update_message(id, chat | team/channel, action, body?)`
`action`: `edit`, `delete`, `restore`. Only the person's own messages. Delete is Teams' soft delete
and `restore` undoes it. Each of the three changes what everyone in the conversation sees at once, so
each needs `user_confirmed` and draws from the same bucket as a new message. Editing or deleting in channels needs `ChannelMessage.ReadWrite`
(administrator consent); in chats it does not. A soft-deleted channel post or reply answers 404 to a
GET (measured), so `restore` finds it, and checks its author, in the channel's or thread's listing.

| Must refuse | The person is told | Test |
|---|---|---|
| A message someone else wrote, deleted ones included | only your own messages | `test_teams.py::test_someone_elses_message_is_never_changed`, `test_teams.py::test_someone_elses_deleted_channel_post_is_not_restored` |
| `edit` without `body` | — | `test_teams.py::test_edit_needs_a_body_and_one_place` |
| `user_confirmed` not true, for any action | ask the person; nothing was changed | `test_teams.py::test_changing_a_message_without_confirmation_changes_nothing` |
| Bucket empty | the wait | `test_teams.py::test_changing_a_message_draws_from_the_send_allowance` |
| A message id that was never there: Graph answers `403` in a chat or channel the person reads (measured), which is no right they lack | it does not exist, or it is not shared with you; a restore looks in the listing first | `test_teams.py::test_an_unknown_message_in_a_chat_is_not_found_not_forbidden`, `test_teams.py::test_restoring_a_post_never_in_the_channel_is_not_found` |
| An edit longer than Teams accepts (the same limit as `send_message`) | shorten it | `test_teams.py::test_an_edit_longer_than_teams_accepts_is_refused` |
| `restore` of a post older than the latest 200 the scan reads (not "does not exist": that is not known) | restore it in Teams | `test_teams.py::test_looking_for_a_deleted_post_reads_four_pages_and_no_more` |
| `edit` or `restore` while a member of the chat or channel is outside `M365_ALLOWED_RECIPIENT_DOMAINS`, when it is set (the text reaches whoever is in the conversation now, a guest added since included; `delete` delivers no text and is not checked) | the domains outside it, never the addresses; nothing was changed | `test_recipient_domains.py::test_changing_a_chat_message_with_an_outside_member_is_refused`, `test_recipient_domains.py::test_changing_a_channel_post_with_an_outside_member_is_refused` |

### `save_chat(id?, topic?, add_members[]?, remove_members[]?)`
Create a group chat, rename it, change its members. Switched off by default (switch `chat_admin`): members see every change at once, a rename too ("renamed the group"). Adding someone already in the chat adds nothing and says `already_members`.

| Must refuse | The person is told | Test |
|---|---|---|
| Creating a chat or changing its members without `user_confirmed` | ask the person | `test_teams.py::test_creating_a_group_chat_needs_confirmation_and_two_people` |
| Renaming a chat without `user_confirmed` | ask the person | `test_teams.py::test_renaming_a_group_chat_needs_confirmation` |
| Removing oneself (that is leaving the chat) | leave it in Teams | `test_teams.py::test_removing_oneself_from_a_chat_is_refused` |
| A new group chat without two different other people (a name twice, or oneself, which Graph refuses with a bare 400) | two other people, each named once | `test_teams.py::test_a_group_chat_needs_two_different_other_people` |
| A chat name over 250 characters, or blank | the limit, or that it cannot be blank | `test_teams.py::test_a_chat_name_too_long_or_blank_is_refused` |
| Renaming a one-to-one or meeting chat (Graph answers a bare 400) | only group chats have a name | `test_teams.py::test_a_one_to_one_chat_is_not_renamed` |
| Changing members of a chat whose member list cannot be read whole (over ten pages) | change them in Teams | `test_teams.py::test_a_chat_whose_members_cannot_all_be_read_is_not_changed` |
| Adding a member outside `M365_ALLOWED_RECIPIENT_DOMAINS`, when it is set | the domains outside it, never the addresses | `test_recipient_domains.py::test_adding_a_chat_member_outside_the_allowed_domains_is_refused` |

## Tasks: To Do and Planner

One set of tools for both: To Do holds the person's own tasks, Planner the team's. A task's
`source` says which.

### `list_tasks(source=all, list?, plan?, bucket?, status?, assigned_to_me=true, due_before?, query?, limit, cursor)`
With no `list`, `plan`, `status`, `due_before` or `query`: the To Do lists, and the Planner plans the
person is a member of with their buckets — as `list_notes` answers with notebooks. Plans come from
`/me/planner/plans` and from the groups of the person's teams (`/me/joinedTeams` →
`/groups/{id}/planner/plans`): a plan made in a team was missing from the first (measured). At most 20
teams are read; more is said with `truncated`. A team that refuses is passed over; throttling or an
outage fails the call rather than shortening the list. With any of them:
tasks with title, status, due date, importance or priority, assignees (Planner), list or plan and
bucket, checklist progress, overdue first, then by due date; across every list and plan unless one is
named. `status` is `open` (the default once tasks are listed), `completed` or `all`; each task's
status is in `save_task`'s words (`open`, `in_progress`, `completed`).

### `save_task(id?, source, list | plan, title, due?, reminder?, importance?, bucket?, assignees[]?, notes?, checklist[]?, status?)`
Create or update. Planner updates send the task's `If-Match` etag; a conflict says the task changed
and returns the current one. Assigning Planner tasks to other people notifies them: assignees other
than the person need `user_confirmed`. To Do keeps only the date of `due` (measured: 17:00 became
00:00); a due time other than midnight is saved as its date and `due_note` says so.

### `delete_task(id, source, user_confirmed)`
Graph documents no recycle bin for To Do or Planner tasks, so deletion is treated as irreversible.
Completing a task (`save_task(status=completed)`) is the reversible alternative, and the description
says so.

| Must refuse | The person is told | Test |
|---|---|---|
| Premium Planner plans | Graph does not expose them | `test_tasks.py::test_premium_or_foreign_plans_are_not_found` |
| A Planner plan the person is not a member of | no plan of yours matches | `test_tasks.py::test_a_plan_the_person_is_not_a_member_of_is_not_found` |
| `delete_task` without `user_confirmed` | ask the person | `test_tasks.py::test_delete_task_needs_confirmation` |
| A title over 255 characters (To Do cuts it and answers 201, measured), counted as Microsoft counts them: in UTF-16 units, so an emoji is two (QA: an emoji title was cut) | the limit and the count | `test_tasks.py::test_a_task_title_over_255_characters_is_refused`, `test_tasks.py::test_a_title_is_counted_as_microsoft_counts_it` |

## Notes: OneNote

### `list_notes(notebook?, section?, query?, limit, cursor)`
Notebooks and sections (with no arguments), or pages of a section, newest first; `query` searches
page titles. Measured: OneNote sends no `@odata.nextLink` once `$top` is given, and honours `$skip`;
the tool asks for one item more than the page and continues by `$skip` in its own cursor.

### `get_note_page(id, max_chars=20000)`
One page as text (reduced from OneNote HTML), with its images and attachments as download links.
`max_chars` is 1 to 50,000 and is honoured as given.

| Must refuse | The person is told | Test |
|---|---|---|
| `max_chars` below 1 or over 50,000 | the range | `test_notes.py::test_get_note_page_refuses_max_chars_outside_its_range` |
| A page OneNote lists, and answers GET of with its title, while its content, a change and a delete answer 404 (measured for 26 minutes and more, through Graph's own `contentUrl` too) — in `get_note_page`, `save_note_page` and `delete_note_page` | that OneNote lists the page but does not open it; open it in OneNote — never "does not exist" | `test_notes.py::test_a_page_onenote_lists_but_does_not_open_is_said_plainly` |

### `save_note_page(id?, section, title, body, action=create, body_upload=false, attachments_upload=0)`
`action`: `create`, `append`, `prepend` (OneNote does not replace a whole page body; delete and create instead); `id` with `title` alone renames the page.
A body change to a page made in the last minute, or to one whose listed title is empty, sends the
page's own title too, read from its content: measured, a body `PATCH` within seconds of creation left
the listed title empty for good in two tries of three, and one carrying the title kept it three of
three. Older pages are left alone, so a rename made meanwhile elsewhere is not undone. Body through the outgoing HTML pipeline; images
from the sandbox are uploaded and embedded. OneNote limits: 4 MB per request without files, 25 MB
per file, 150 images per page.

### `delete_note_page(id, user_confirmed)`
Whether deleted pages are recoverable is a live finding; until settled, `user_confirmed`.

OneNote throttles at 120 requests per minute per person and sends no `Retry-After`; the client
paces itself and backs off on its own.

## Bookings (switched off by default)

For organisations that take appointments from customers. Graph requires the business to have a
Microsoft 365 Business Premium subscription; it works with shared booking pages only.

### `list_bookings(business?, start?, end?, include_services=false, include_staff=false, limit, cursor)`
Booking businesses the person can manage (without `business`), or appointments in a window with
customer, service, staff and time.

### `save_booking(id?, business, service, start, customer_name, customer_email, staff[]?, notes?, user_confirmed)`
Create or change an appointment. The customer and staff are notified: `user_confirmed`.

### `cancel_booking(id, business, message, user_confirmed)`
Cancels with a message to the customer and staff. Irreversible.

## Rate limits the client enforces itself

Graph throttles per app and person; the client paces below the limits instead of collecting 429s.

| Area | Limit Graph applies | Client |
|---|---|---|
| Outlook (mail, calendar) | 10,000 requests / 10 min per mailbox, 4 concurrent, 150 MB upload / 5 min | at most 4 in flight per person |
| Teams chats and channels | about 1 request per second per person and per chat or channel | 1 rps per person per area |
| Chat members | 4 per minute per chat | serialised |
| OneDrive and SharePoint | 3,000 requests / 5 min per person | at most 4 in flight per person |
| OneNote | 120 / min and 400 / hour per person, 5 concurrent, no `Retry-After` | 1 rps per person, own backoff |
| Excel | 1,500 / 10 s per app per tenant; concurrent writes to one workbook conflict | writes to one workbook serialised |
| Bookings | 4 concurrent per mailbox | at most 4 in flight |

## Permissions

| Area | Delegated scopes | Administrator consent |
|---|---|---|
| Always | `User.Read`, `User.ReadBasic.All`, `People.Read`, `MailboxSettings.Read`, `offline_access` | no |
| Mail | `Mail.ReadWrite`, `Mail.Send` | no |
| Calendar and scheduling | `Calendars.ReadWrite`, `Calendars.Read.Shared`, `Presence.Read.All`, `OnlineMeetings.Read` | no |
| Files | `Files.ReadWrite.All`, `Sites.ReadWrite.All` | no |
| Teams chats | `Chat.ReadWrite`, `ChatMessage.Send`, `ChannelMessage.Send`, `Team.ReadBasic.All`, `Channel.ReadBasic.All` | no |
| Mail settings | `MailboxSettings.ReadWrite` (automatic replies) | no |
| Contacts | `Contacts.ReadWrite` | no |
| Lists | `Sites.ReadWrite.All` | no |
| Excel | `Files.ReadWrite` | no |
| Tasks | `Tasks.ReadWrite` | no |
| Notes | `Notes.ReadWrite` | no |
| Teams: channel reading, editing own channel posts | `ChannelMessage.Read.All`, `ChannelMessage.ReadWrite` | **yes** |
| Teams: posting to a channel while `M365_ALLOWED_RECIPIENT_DOMAINS` is set (its members are checked) | `ChannelMember.Read.All` | **yes** |
| Recordings and transcripts | `OnlineMeetingRecording.Read.All`, `OnlineMeetingTranscript.Read.All` | **yes** |
| Rooms | `Place.Read.All` | **yes** |
| Bookings (off by default) | `Bookings.Read.All`, `BookingsAppointment.ReadWrite.All` | no, but needs Business Premium |

**Administrator consent is given once, at installation.** The customer's administrator opens one
consent link and accepts every scope of the enabled areas for the whole organisation; after that
each person only signs in. The scopes marked **yes** are delegated — the server still sees only
what the person can see — Microsoft simply does not let a person grant them alone. Without that
consent Microsoft stops every person's sign-in at "Need admin approval", because the gateway asks for
all scopes in one sign-in: the consent is a prerequisite of the connector, not of five features. An
organisation that refuses those five scopes gets a connector entry without them and the matching
switches off; people can then consent alone where the tenant's consent policy allows it.

A tenant's own consent policy may still require an administrator for scopes Microsoft marks as
user-consentable; `get_connection` reports what was actually granted. Which of `ChatMessage.Send`
and `ChannelMessage.Send` Graph enforces for chat posts is a live finding to settle on the test
tenant.

## Out of scope

| What | Why |
|---|---|
| Teams Shifts | not needed |
| Microsoft 365 group conversations | niche; needs administrator consent for reading |
| Insights ("used", "shared") | Graph stops returning data in November 2026 |
| Forms, Viva, Stream (classic), Loop | no usable Graph API for this purpose |
| Tenant, group and mailbox administration | not the assistant's job; manifesto rule 24 |

## Live findings to settle on the test tenant before code relies on them

- Which of `ChatMessage.Send` / `ChannelMessage.Send` is enforced for chat posts.
- Whether `$search` combines with `$filter` on messages.
- Where deleted contacts, list items, OneNote pages and To Do tasks go.
- The source size limit of `?format=pdf` conversion (the source types are settled above).
- Whether a participant can download a shared recording through the drive item.
- Whether threaded replies exist for chat messages.
