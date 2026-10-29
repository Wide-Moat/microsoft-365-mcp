# Component: graph client

The only code that speaks HTTP to Microsoft Graph. Every tool and every link route goes through it.
Its job is to make Graph's silent failure modes loud and its quirks invisible to the rest of the
server.

## Interface

```python
class Graph:
    async def request(self, method, path, *, params=None, json=None, content=None, headers=None,
                      retry_server_errors=None) -> httpx.Response
    async def get(self, path, *, params=None, headers=None) -> dict
    async def post(self, path, json=None, *, headers=None) -> dict | None
    async def patch(self, path, json, *, if_match=None, headers=None) -> dict | None
    async def delete(self, path, *, if_match=None) -> None          # only paths on the allowlist
    async def page(self, path_or_next, *, params=None, headers=None) -> Page   # items, next link
    async def immutable_ids(self, ids) -> dict[str, str]
    def stream(self, path, *, headers=None, params=None) -> AsyncContextManager[httpx.Response]
    async def upload_session(self, upload_url, chunks, size) -> dict   # PUT ranges to Graph's uploadUrl
    async def monitor(self, url) -> dict                             # a copy's progress
    async def patch_note_page(self, page, ops) -> None
    async def note_page_html(self, page) -> str
```

`Graph` is constructed per call from the person's token. It holds no state across calls except the
shared pacing limiters, keyed by person.

## What it does on every request

| Behaviour | Why |
|---|---|
| `Prefer: IdType="ImmutableId"` | Outlook ids change on move without it; the header applies only to the request it is on |
| No `Prefer: outlook.timezone` | Times stay in UTC on the wire and the server converts them to the person's zone, read from `mailboxSettings` and kept for an hour per person (`core/context.py`); local wall times from Graph could be read back correctly only for zones the server can name |
| `Prefer: outlook.body-content-type="html"` when reading bodies | One HTML path into the text reducer, never two |
| `client-request-id` (UUID) | Correlates our log line with Microsoft support; returned in errors |
| `User-Agent: m365-mcp/<version>` | Identifies the product to Microsoft |
| Timeout: 10 s to connect, 30 s for each read, write and pool wait, streaming included (no total for a stream) | A hung Graph call must end |

## Retries and pacing

- **429 and 503:** wait `Retry-After` (seconds) when present, otherwise exponential backoff from
  1 s, up to 3 attempts and 30 s total; then a throttling error with the seconds to wait.
- **502 and 504 on idempotent requests:** one retry. Never on `POST` that sends or creates.
- **Pacing before Graph has to throttle:** per person, at most 4 in-flight Outlook requests and 4
  in-flight SharePoint/OneDrive requests; Teams at most 1 request per second per person (a token
  bucket of 1, refill 1/s), member changes serialised per chat.
- Pacing is in-process. Several replicas multiply it; the limits are set so that the default
  replica count stays under Graph's.

## Hosts

Requests with the person's token go to `graph.microsoft.com` only. Graph's own answers name other
hosts — a download's `302`, an upload session's `uploadUrl`, a copy monitor — and those are followed
without the token, and only to Microsoft's file, media and mail hosts (`FILE_HOSTS`:
`sharepoint.com`, `onedrive.com`, `1drv.com`, `svc.ms`, `microsoft.com`, `office.com`,
`office365.com`, `office.net`, `outlook.com`, and their subdomains; the copy monitor a narrower set).
A redirect elsewhere is a `502`, an upload URL elsewhere is refused before any byte is sent.

## Paging

- `page()` requests `$top` = the tool's limit, returns the items and an **opaque cursor** that
  wraps `@odata.nextLink` (or a search `from` offset) together with the tool name and a hash of the
  person. A cursor presented by another tool or person is refused.
- No `$skip` is ever sent where Graph does not support it (`$search` on messages, chats).

## Errors

Every non-2xx response becomes one `GraphError(status, code, request_id, retry_after)`. The text of
the Graph error body is never placed in the exception message — it can contain addresses and
subjects. `core/errors` maps `GraphError` to the person-facing message:

| Graph | Person sees |
|---|---|
| 401, `InvalidAuthenticationToken` | reconnect Microsoft 365 |
| 403, `Authorization_RequestDenied`, `ErrorAccessDenied` with a missing scope | which permission, administrator needed |
| 403, `GraphAccessToTranscriptsDisabled` | the organisation turned off transcript access |
| 404, `ErrorItemNotFound`, `itemNotFound` | not found or not shared with you |
| 409, `nameAlreadyExists` | an item with that name already exists |
| 412 | it changed since it was read; read it again |
| 413, `ErrorMessageSizeExceeded` | the size limit, with numbers |
| 429, 503 after retries | Microsoft is throttling; wait N s |
| anything else | the action did not happen; request id |

## The delete allowlist

`delete()` (any `DELETE` through `request()`) refuses a path that does not match `_DELETE_ALLOWED`
before any HTTP. Allowed: drive items, in OneDrive or a site's drive (they go to the recycle bin);
a drive item's sharing permission (revoking access); To Do tasks; Planner tasks; OneNote pages;
personal contacts, in the default folder or another; SharePoint list items; an attachment of a
message (removing it from a draft); a member of a group chat; and an event (`cancel_event` deletes
only one nobody is invited to; it goes to Deleted Items). Messages are never `DELETE`d — they are
moved — and `permanentDelete` is not reachable. A test refuses a `DELETE` of a message, of an event's
attachment, of a calendar, of a folder and of anything under `/users/`.

## Tests

- Every request carries the immutable-id header and none asks Graph for local times (asserted on a recording transport).
- 429 with `Retry-After: 2` waits 2 s (fake clock) and retries; three 429s produce the throttling
  error; a `POST /send` is never retried on 504.
- Pacing: five concurrent Outlook calls for one person run at most four at a time.
- A cursor from another tool or another person is refused.
- An error body containing an address does not appear in the exception message or the log.
- `delete()` on a message path raises before any HTTP.
- Streaming download releases the connection when the consumer stops early.
