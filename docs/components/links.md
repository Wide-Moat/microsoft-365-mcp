# Component: signed links

Moves bytes between the agent's sandbox and Microsoft 365 without putting them in the model's
context and without giving the sandbox a Microsoft token.

## Routes

| Route | Method | Does |
|---|---|---|
| `/d/<token>` | `GET` | Streams one object from Graph: a mail attachment, a drive item (optionally converted), a meeting recording, a transcript |
| `/u/<token>` | `PUT` | Streams a body into one target: a draft's HTML body, a new draft attachment, a drive item (create or replace), a Teams message attachment, a OneNote page body |

Nothing else is served on the public hostname.

## Token

The link carries a short id: `secrets.token_urlsafe(16)`, 128 random bits, 22 characters
(`links/store.py`). The id stands for the signed value below, kept in the process's memory
until the link's own expiry, at most 5,000 open links and 300 per person; when full a new link is
refused with a sentence, never by dropping one in use. A restart or a second replica does not know
the id, so the server runs one replica and a restart answers 403 to the open links: the agent asks
for a new one. A full signed value is still accepted and checked the same way.

```
<base64url(payload)>.<base64url(HMAC-SHA256(key_sign, payload))[:22]>
payload = {
  "o": "d" | "u",               # operation, inside the signature
  "k": "<kind>",                # attachment | drive | recording | transcript | draft_body | draft_attachment | chat_file | note_body
  "t": "<target ids>",          # e.g. message id + attachment id
  "p": "<person hash>",         # sha256(tenant id + object id)[:16]
  "e": <unix expiry>,
  "x": "<AES-GCM(key_enc, graph access token)>",
  "m": {"max": <bytes>, "ct": "<allowed content type>"}   # uploads only
}
```

- The signature covers the whole payload. Verification uses `hmac.compare_digest`.
- The Graph token travels encrypted with a separate key and is used only for this one transfer.
- Expiry: download 300 s, upload 180 s, hard cap 900 s, and never later than the token's own
  `exp`. Expiry is checked when the request arrives, not during the transfer.
- Keys: a list per purpose; the first signs or encrypts, all verify or decrypt, so rotation does
  not break links in flight. **No key configured → links are not issued and the tools that need
  them say so; the server does not start in that state if any byte-moving area is enabled.**
- Garbage, a wrong operation, a wrong kind, an expired or unknown link, a person mismatch → `403`
  with one fixed sentence that says to ask for a new link, and no detail of which check failed.
  Never `500`.

## Streaming

- Download: the Graph stream opens inside the route, and is released when the response ends, however
  it ends: by the body generator's `finally` for a client that disconnects mid-way, and by the
  response's own `finally` for one gone before the first chunk, when the generator never starts. A
  context manager must not close the stream before the response body is sent. `HEAD` answers the
  headers and releases the stream without reading the file.
- Upload: the request body is read in chunks into a spooled temporary file (in memory up to 8 MB,
  then disk), counting bytes against `m.max`; over the limit → `413` and the partial file is
  deleted. Then it is sent to Graph: one request under the target's single-request limit, otherwise
  an upload session with fragments of a multiple of 320 KiB (drive) or up to 4 MB (attachments).
- Content type of uploads is checked against `m.ct` from the payload, not trusted from the
  sandbox; HTML bodies then go through the HTML pipeline before they reach the draft.
- An upload link moves one file, once. While a body arrives on it the link is busy and a second
  `PUT` on it answers `409`; once a file was delivered the link is spent and answers `403` like an
  expired one. A body the link refused (413, 415, 408, 503, 507, a Graph failure) leaves it usable for
  the corrected one until it expires.
- Both draft upload kinds refuse (`422`) a message that is no longer a draft or carries the sent
  mark, which `send_draft` sets before it sends: nothing is added to a letter after the person said
  send, nor to its copy in Sent Items. `send_draft` refuses while one of the draft's links is busy
  (an upload past its check of the draft), and spends the draft's links only after a send that went
  through, so a failed send leaves them usable.
- All uploads running at once share one spool budget (1 GiB per process); a body that would pass it
  is dropped with `503` and `Retry-After`, so parallel uploads each under their own limit cannot fill
  the disk. A disk that fails the spool's write (full, or failing) answers `507` with `Retry-After`,
  with the body's bytes given back to the budget and its spool closed.
- A body that sends nothing for 30 s, or is still arriving after 900 s, is cut off with `408`.

## Response to the model

Tools that issue links return the URL and a ready command:

```
curl -sS --fail-with-body -o 'report.pdf' 'https://<links-host>/d/<token>'
curl -sS --fail-with-body -T 'letter.html' -H 'Content-Type: text/html; charset=utf-8' 'https://<links-host>/u/<token>'
```

File names in commands are shell-quoted; the name comes from Graph and is untrusted.
`--fail-with-body` (curl 7.76 or later) makes a refusal exit non-zero and still keeps its sentence:
on stdout for an upload, in the output file for a download. Throttling is passed on as `429` or
`503` with `Retry-After`, never as "not found".

## Tests

- Tampered payload, tampered signature, swapped operation, swapped kind, expired, other person,
  empty string, non-base64 → 403, never 500.
- No key → issuing raises a configuration error; the server refuses to start with byte areas on.
- Key rotation: a link signed with the old key verifies while the old key is in the list.
- Expiry never exceeds the token's `exp`.
- Download: a fake Graph stream that refuses reads after close; a 5-chunk file arrives whole; a
  client disconnect after chunk 2 closes the upstream stream.
- Upload: over `max` → 413 and no temp file left; wrong content type → 415; 10 MB file uses an
  upload session with fragments of the right size.
- Upload: a second `PUT` after a delivered one → 403; one while another runs → 409; past the spool
  budget → 503 with `Retry-After` and the budget given back; a write the disk refuses → 507, the
  budget given back; a stalled body → 408.
- A file name with quotes and `$(…)` is quoted safely in the returned command.
