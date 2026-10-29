# Code review

What a reviewer — a person, or `ocr review` driven by [`.opencodereview/rule.json`](../.opencodereview/rule.json)
— looks for in every change. Each item points at the [manifesto](../MANIFESTO.md) rule it protects
and says what a violation looks like. A finding is fixed, or rejected with its reason written in
the review; it is never silently dropped.

## Blocking

A change with any of these does not merge.

| # | Look for | Rule |
|---|---|---|
| B1 | A Graph call outside `graph/` (raw `httpx`, `requests`, or SDK use in a tool module) | 39 |
| B2 | A path under `/users/{x}/` for mail, drafts, events, files or chats; any argument selecting a mailbox, user or tenant | 15, 16 |
| B3 | `client_credentials`, an application permission, a stored token, a token in a log line | 15 |
| B4 | Any code path that sends mail, posts a Teams message, invites, reschedules, cancels or shares outside existing members without checking `user_confirmed` and drawing from the bucket | 20–22 |
| B5 | `DELETE` on a message or chat message, or on an event that has attendees (those are cancelled, rule 21); `permanentDelete`; emptying a folder or the recycle bin | 19 |
| B6 | Inbox rules, forwarding, delegates, permissions, export, anonymous links — in a tool, a client method or a route (automatic replies are allowed only through `save_automatic_reply`, confirmed) | 24 |
| B7 | A tool result built from a Graph object without a projection; any `@odata.` key reaching the model | 8 |
| B8 | A list without a cap, an `offset` Graph does not honour, a truncation that does not say so | 9 |
| B9 | Bytes or base64 of a file, attachment or large HTML in a tool result | 10 |
| B10 | An exception caught and returned as a successful result; a result shaped `{"error": …}` | 11 |
| B11 | An annotation that contradicts the code: `readOnlyHint` on a write, `destructiveHint: false` on a change, `idempotentHint` on a send | 12 |
| B12 | A subject, body, address, file name, search string or person's name in a log call or an exception message | 38 |
| B13 | Outgoing HTML that skips any stage of the pipeline (sanitise → `cid:` images → inline CSS → validate) | 29, 30 |
| B14 | A reply or forward built by replacing the body instead of Graph's `createReply` / `createReplyAll` / `createForward` | 31 |
| B15 | A signed-link check using `==` instead of `hmac.compare_digest`; a signer that works with an empty key; a payload field outside the signature | 32 |
| B16 | A file read or written whole into memory | 33 |
| B17 | A new behaviour or a new negative scenario without a test that fails without it | 36 |
| B18 | A domain, address, tenant, bucket or customer name of a real installation | 41 |
| B19 | A new scope in the connector without a published tool that needs it | 18 |

## Must answer

The author answers each of these in the pull request, even when the answer is "not applicable".

- **What does the model see?** Paste the tool's description and a real (redacted) result.
- **What does it cost?** The catalogue's token count before and after (`make budget`).
- **What can the person not undo?** List every effect visible to other people and how it is gated.
- **What does Graph actually do?** For any new Graph call, the live finding or the recorded fixture
  it was checked against — not the documentation alone.
- **What fails, and how does the person learn of it?** The error text for each refusal.

## Style

- Names follow the manifesto: `list_`, `get_`, `save_`, or an explicit verb; no prefixes.
- Descriptions: when to call it, what it returns, what it will refuse; at least three sentences;
  no capitals for emphasis; no example values the model would copy.
- Arguments: flat; enums inline; every argument whose name is not enough has a description that
  adds something the name does not say.
- One module per area (`mail`, `calendar`, `files`, `teams`, `people`); shared behaviour lives in
  `core/`, not copied.
- Comments explain why, not what. English only.

## Tests a reviewer expects to see

- A negative test for every refusal the change adds, asserting `isError` and the text; the refusal
  as a Must refuse row naming that test, and a mutation for it in `scripts/refusal_evidence.py`.
- A change to `tool-surface.lock.json` that the change meant to make: each line in its diff is a
  change to what the model is offered.
- Contract tests still green: names, descriptions, annotations, budget, `docs/tools.md` in sync.
- Recorded Graph fixtures for new calls; no test that talks to a real tenant outside the `live`
  marker.
- For anything streaming: a fake that reproduces the stream closing, so a premature close fails.
