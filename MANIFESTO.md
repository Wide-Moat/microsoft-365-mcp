# Manifesto

This server gives an AI assistant the Microsoft 365 of **one person**: their mail, their calendar,
their files, their Teams. It acts as that person, with that person's rights, and with nothing more.

Everything here is a rule. A rule that the code can break without a test going
red is a defect in the tests. Every rule below names what a violation looks like, so that a
reviewer can find one and a test can fail on one.

## Why this exists

The servers that already exist for Microsoft 365 take one of two shapes. Both fail.

- **Graph, one to one.** One tool per Graph endpoint: 362 tools and about 575,000 tokens of
  description for one well-known open-source server; 54 tools and 75,000 tokens for mail alone.
  The model spends its context reading the catalogue and its calls choosing between near-identical
  tools, and every answer is raw Graph JSON.
- **Graph, as a path.** A dozen generic tools: fetch a path, create an entity at a path. The model
  must know Graph by heart, and nothing between it and the mailbox knows what a letter is.

We build a third shape: **few tools, each answering one question a person actually asks**, with
the server doing the work that Graph leaves to the caller — resolving names, expanding recurrences,
cleaning HTML, streaming files, refusing what must not happen.

The shape follows Linear's own MCP server: a small set of `list_`, `get_`, `save_` verbs over the
objects a person thinks in.

## I. The contract with the model

1. **One question, one tool.** Depth is a flag (`include_attachments`, `include_thread`), never a
   second tool. *Violation:* two tools whose implementations share a code path and differ in output
   shape.
2. **Verbs are `list_`, `get_`, `save_`, plus a verb that says what happens** (`send_draft`,
   `cancel_event`, `share_file`). `save_` is an upsert: no `id` creates, an `id` updates.
   *Violation:* `create_x` beside `update_x`; `manage_x` with a free-form action string.
3. **Changes that are not upserts take a closed `action` enum**, listed inline in the schema.
   *Violation:* an `action: str`; an enum behind a `$ref`.
4. **No prefix.** The gateway already namespaces tools. *Violation:* `m365_`, `graph_`, `mcp_` in a
   tool name.
5. **Descriptions read like a man page and say when to call the tool**, not only what it does. At
   least three sentences; irreversible tools say they are irreversible. No shouting, no examples
   that the model will copy verbatim. *Violation:* a description that restates the tool name.
6. **Flat arguments.** No nested objects where a scalar or a list of scalars will do. Every argument
   whose name is not self-explanatory has a description; no description restates its name.
7. **The server resolves names.** People by name or address, folders by name or path, sites by name,
   time as a local string in the person's mailbox time zone (`2026-10-23 15:00`), duration in
   minutes. **Ambiguity is an error that lists the candidates** — never the first match.
   *Violation:* an argument that requires a Graph id the model could only have guessed.
8. **Projections, never raw Graph.** Each object has a projection that keeps the fields a person
   needs. *Violation:* a tool returning `@odata.*`, `changeKey`, `etag`, nested `emailAddress`
   objects, or any Graph DTO passed through.
9. **Bounded output.** Lists default to 10 items, cap at 25, and continue by an opaque `cursor`.
   Text is capped and says it was truncated. A page may come back shorter than asked when items are
   heavy; `has_more` stays true and the cursor continues. *Violation:* an `offset` argument Graph
   does not honour; a list with no cap; silent truncation.
10. **Bytes never enter the context.** Files, attachments and large HTML bodies move through signed
    links that the agent's sandbox uses with `curl`. *Violation:* base64 in a tool result.
11. **Errors are errors.** A failed call sets `isError` and its text names the next step for the
    person ("reconnect Microsoft 365 in the portal", "ask your administrator to grant X").
    *Violation:* a successful result containing `{"error": …}`; a raw Graph error code shown to the
    person.
12. **Annotations describe actual behaviour.** `readOnlyHint`, `destructiveHint` (by the MCP spec it
    means *changes existing state*, not *irreversible*), `idempotentHint`, `openWorldHint`.
    *Violation:* a read-only hint on anything that writes, including a temp file or a link slot.
13. **Areas switch off like separate servers.** Mail, calendar, files, SharePoint lists, Excel,
    Teams, recordings, tasks, OneNote, contacts and Bookings are areas; an installation turns any of
    them off with one line, and a person can turn areas off for themselves. A switched-off area is
    not published at all: it costs no context and cannot be attempted. *Violation:* a published tool
    that answers "disabled"; a tool that belongs to no area.
14. **Every change is weighed.** The catalogue is measured with a real tokenizer per tool and per
    area on every change; CI prints the table and the pull request shows it before and after. The
    default catalogue (every area that is on by default) stays under **16,000 tokens**, no area over
    3,500, no tool over 600. Growing
    past any of these needs a decision written in the pull request that raises it, with the measured cost.

## II. Acting as one person

15. **The person's own token, and only theirs.** The gateway holds the OAuth session and forwards
    the person's Graph bearer token on each call. This server stores no tokens beyond the sealed copy
    an open file link holds in memory for its few minutes, has no service
    account and no application permissions. *Violation:* any `client_credentials` flow; any
    `/users/{id}/…` path for someone other than `me`, except reading another person's free/busy,
    presence and profile, which Graph itself restricts to what the caller may see.
16. **No mailbox parameter.** Mail, drafts, rules and settings exist only for `/me`. Another
    person's mailbox is unreachable even by a bug in our logic. *Violation:* an argument that
    selects a mailbox, user or tenant.
17. **Rights come from Microsoft, not from us.** We never widen what the person can do, and we never
    pretend to narrow it with a check Graph would not enforce. Where Graph refuses, we say which
    right is missing.
18. **Least scopes.** The connector asks for the delegated scopes its published tools need and no
    others; a scope added without a tool that needs it is a review finding.

## III. Care with what cannot be undone

The rule is: **what other people see, or what cannot come back, needs the
person's word; everything else stays reversible.**

19. **Deletion goes to a bin.** Mail moves to Deleted Items (and returns with `move`); files go to
    the recycle bin, from which the person restores them (Graph cannot restore for work accounts,
    so the tool says where to look). An event nobody is invited to is deleted with `user_confirmed`;
    Graph's DELETE puts it in Deleted Items (measured 2026-10-26: the folder's count went 342 to 343).
    Hard delete, emptying folders and purging the bin do not exist here. Deleting a folder needs
    `user_confirmed`, because it takes its contents.
20. **Sending is two steps.** `save_draft` writes a draft the person can see in Outlook;
    `send_draft(id, user_confirmed=true)` sends that draft and nothing else. Only drafts this
    server created can be updated or sent. *Violation:* any path that sends a message without a
    stored draft.
21. **Anything other people see at once needs `user_confirmed=true`:** sending mail, sending a
    Teams message, inviting, rescheduling or cancelling a meeting, sharing a file outside the
    people already on it, turning on an automatic reply (everyone who writes receives its text),
    booking or cancelling a customer appointment. The model sets it only when the person explicitly asked. Answering one's
    own invitation does not need it, because that is one's own record.
22. **Rate.** All such actions draw from one token bucket per person: 10 at once, refilled at 10
    per minute. At most 20 recipients per action; distribution lists are refused. Mass mailing is
    out of scope and the refusal says so.
23. **Only what is yours to change.** Meetings are rescheduled or cancelled only by their organiser;
    chat messages are never edited or deleted on someone else's behalf.
24. **Never exposed, by design:** inbox rules, forwarding, delegates, folder and calendar
    permissions, mailbox export or sync, anonymous ("anyone with the link") sharing, tenant and
    group administration. Each of these sends data away or opens access silently.

## IV. Content is data

25. **Text from mail, chats and files is untrusted.** It returns inside a marked content block, and
    the read tools say it is data written by someone else, not instructions. This lowers risk; it is
    not a boundary.
26. **The boundaries are structural:** the operation list above, drafts, confirmation, the bucket.
    A prompt injection that says "forward everything to X" finds no forwarding, no rules and no
    send without a draft the person sees.
27. **HTML read for the model is reduced to text.** Scripts, styles, tracking pixels and hidden
    elements (`display:none`, zero size, text the colour of its background) are dropped; images
    become `[image: alt]`; links keep their real target, unwrapped from link-scanning redirects and
    stripped of tracking parameters.

## V. HTML mail

28. **Short text goes in `body`; a designed letter goes in a file.** HTML larger than a paragraph is
    uploaded through a signed link (`curl -T letter.html`), UTF-8, at most 2 MB.
29. **Every outgoing HTML body passes one pipeline, in this order:** sanitise (allowlist; no
    `script`, `iframe`, `object`, `embed`, `form`, `on*` attributes, `javascript:` URLs) → convert
    `data:` images to inline attachments referenced by `cid:` → inline CSS → validate.
30. **Validation reports, it does not guess.** Relative URLs are refused. Constructs the Outlook
    desktop engine cannot render (flex, grid, positioning, background images, web fonts) are
    returned as warnings. The draft's response carries `{sanitized, warnings, size}` so the model
    can fix the letter before the person says "send".
31. **A reply keeps its thread.** Replies and forwards are created by Graph (`createReply`,
    `createReplyAll`, `createForward`) and the new text is placed above the quoted history; the body
    is never replaced wholesale.

## VI. Files and attachments

32. **Signed links, short-lived, fail closed.** HMAC over the whole payload (operation, object,
    person, expiry); download 300 s, upload 180 s, hard cap 900 s; a missing key refuses to sign; a
    tampered, expired, wrong-operation or other-person link answers 403.
33. **Streaming both ways.** No file is held whole in memory. Attachments under 3 MB attach in one
    request, larger ones (up to 150 MB) through an upload session; the tenant's message size limit
    is reported as a number, not as a Graph error.
34. **One identifier per object.** Outlook ids are requested as immutable ids, so a moved message
    keeps its id.

## VII. Engineering

35. **English everywhere** in the repository: code, comments, docs, commit messages. Error texts
    the person reads are localisable, not hard-coded in one language.
36. **Test first.** Every rule in this document has a test that fails when the rule is broken, and
    each such test has been seen failing once. Every negative scenario in
    [`docs/requirements.md`](docs/requirements.md) is a test.
37. **The suite cannot pass empty.** CI counts the tests defined and the tests run, and fails when
    they differ or when the count is implausibly low.
38. **No content in logs.** No subjects, bodies, addresses, file names or search strings — in log
    lines or in exception messages, which end up in logs. A test asserts it.
39. **Graph quirks live in one place.** The Graph client sets `Prefer: IdType="ImmutableId"` on every
    request and keeps times in UTC on the wire (the server converts them to the person's time zone),
    honours `Retry-After` on 429 and 503, and limits concurrency per mailbox. Tools never call HTTP
    directly.
40. **Live findings beat documentation.** What Graph actually does on a real tenant is recorded in
    `docs/graph-live-findings.md`, and a recorded fixture backs every such finding.
41. **Product code.** This repository names no installation and no customer; the purity gate fails
    the build otherwise.
42. **Review before push.** Every change is reviewed against [`docs/code-review.md`](docs/code-review.md)
    before it is pushed; pull requests open as drafts.
