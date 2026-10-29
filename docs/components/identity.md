# Component: identity

Answers "who is calling" from the bearer token the gateway forwards, and nothing more.

- The token is read from `Authorization: Bearer`. A call without it is refused before any tool
  runs.
- Local checks without verifying the signature (Graph does that): it is a JWT, `aud` is Microsoft
  Graph, `exp` is in the future, `tid` and `oid` are present. Failing any → "reconnect".
- **The client the token was issued to**, when the installation names it. Graph accepts any valid
  Graph token, whichever application obtained it, so a token that leaked from another application
  would work here as well as at Graph. `M365_EXPECTED_CLIENT_IDS` (comma-separated application ids,
  compared without case) closes that: the token's `azp` (v2 tokens) or `appid` (v1, which Graph
  usually issues) must be one of them, and a token naming no client is refused. Failing → "reconnect",
  before any Graph request. **Unset, any client is accepted** — the behaviour before the setting
  existed, kept so that an installation that has not looked up its gateway's id keeps working.
- The person is `(tid, oid)`. `person_hash = sha256(tid + ":" + oid)[:16]` keys the bucket, the
  pacing, cursors and links. Nothing about the person is cached from `/me`: a tool that needs their
  name, address or id reads `/me` in that call. What is cached, in memory and by person or tenant, is
  the mailbox time zone (an hour) and the organisation's verified domains (an hour).
- Granted scopes come from the token's `scp` claim; `get_connection` and the area registry use them
  to tell "area on but permission missing" apart from "area off".
- The token is never logged, never put in an exception, never written to disk. It is passed to the
  graph client and, encrypted, into signed links.

## Must refuse

The audience and the client the token was issued to are rows of the shared table in
[requirements](../requirements.md); the shape checks behind them are here.

| Must refuse | The person is told | Test |
|---|---|---|
| A bearer value that is not a JWT (not three parts, or a payload that is not a JSON object) | reconnect Microsoft 365 in the portal | `test_identity.py::test_a_token_that_is_not_a_jwt_is_refused_before_graph` |
| A token without `tid` or `oid` | reconnect Microsoft 365 in the portal | `test_identity.py::test_a_token_without_a_tenant_or_object_id_is_refused` |
| A token naming no client (`azp` and `appid` both absent) while `M365_EXPECTED_CLIENT_IDS` is set | reconnect Microsoft 365 in the portal | `test_identity.py::test_a_token_naming_no_client_is_refused_when_clients_are_configured` |
| A v1 token whose `appid` is another client while `M365_EXPECTED_CLIENT_IDS` is set | reconnect Microsoft 365 in the portal | `test_identity.py::test_a_v1_token_from_another_client_is_refused` |

## Tests

- No header, a non-JWT, a wrong audience, an expired token → "reconnect", no Graph call.
- With `M365_EXPECTED_CLIENT_IDS` set: a token whose `azp`/`appid` is another client, or that names
  none, → "reconnect", no Graph call; the gateway's own client passes in either claim. Unset: any
  client passes.
- Two different people never share a cache entry; the same person with a new token gets one `/me`
  call.
- `scp` without `Mail.Send` → `get_connection` reports mail sending as missing a permission.
- A log capture during a full call contains no part of the token.
