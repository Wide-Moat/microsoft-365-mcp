# Component: mcp

The transport and the catalogue: what the model is offered and what it costs.

## Transport

Streamable HTTP from the official MCP Python SDK, stateless (no server-side session), JSON
responses. One endpoint, `/mcp`. Health endpoints `/healthz` (process up) and `/readyz`
(configuration valid). Link routes are mounted beside it but served on a separate listener so that
the network policy can expose them alone.

## Registry and areas

Each tool is declared once with its area, its annotations and the scopes it needs:

```python
@tool(area="mail", scopes=["Mail.ReadWrite"], read_only=True, idempotent=True)
async def list_messages(...): ...
```

At start-up the registry builds the published set from `M365_AREAS` (the installation) and, per
call, narrows it further by the area header the gateway forwards (the person). A tool outside the
published set does not appear in `tools/list` and is rejected if called.

A parameter can need an area or a switch of its own, declared beside the tool's:

```python
@tool(area="calendar", ..., gates={"include_recording": ("area", "recordings")})
async def get_event(...): ...
```

Such a parameter is removed from the tool's input schema whenever its area is off (for the
installation or for the person) or its switch is off, and a call that sends it anyway is refused
before the tool runs. A depth of an answer that needs its own permission therefore costs nothing
and cannot be attempted while it is off, without a second tool for the same question.

## Budget

`python -m m365_mcp.budget` renders the exact `tools/list` payload for a configuration and counts
tokens with `tiktoken` (`o200k_base`, cached in the repository so the count runs offline). It
prints a table: per tool, per area, total. CI runs it for the default configuration and for "all
areas on", fails above the limits in the manifesto (rule 14), and posts the table on the pull
request with the difference from `main`.

## Contract tests (on the registry, not on hand-picked tools)

- Every tool has an area, a title, annotations, and a description of at least three sentences
  that does not merely restate the name.
- Names start with an allowed verb; no prefixes; no two tools share a name across areas.
- Arguments are flat; enums are inline literals; arguments whose names are not self-explanatory
  have descriptions; no description restates its argument's name.
- Annotations match a declared behaviour table (`reads`, `writes`, `notifies_others`,
  `irreversible`); irreversible tools say so in their description.
- With an area off, none of its tools are listed, and calling one is refused; a parameter gated on
  an area or switch that is off is absent from the schema, and sending it is refused.
- `docs/tools.md` equals the document generated from the registry.
- Every list-returning tool caps `limit` at 25 and returns `cursor` and `has_more`.
