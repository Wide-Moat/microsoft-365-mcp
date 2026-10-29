# Component: tools

One module per area under `src/m365_mcp/tools/`. Their contracts — arguments, results, refusals —
are in [requirements](../requirements.md); this page is about how a tool is built.

## The shape of every tool

```python
@tool(area="mail", scopes=["Mail.ReadWrite"], notifies_others=False, read_only=True)
async def list_messages(ctx: Ctx, folder: str | None = None, query: str | None = None, ...,
                        limit: int = 10, cursor: str | None = None) -> MessageList:
    """When to call it. What it returns. What it refuses."""
    folder_id = await ctx.resolve.mail_folder(folder) if folder else None   # 1. resolve
    page = await ctx.graph.page(...)                                         # 2. graph
    return MessageList.of(page, project=envelope)                            # 3. project
```

A tool is thin: resolve → (guard) → graph → project. Logic shared by two tools moves to `core/`.
A tool body over ~60 lines is a review question.

## Rules a tool cannot break

- No HTTP outside `ctx.graph`; no Graph object outside a projection.
- Anything that notifies others calls `ctx.guard.check(...)` before the Graph call; the decorator
  flag `notifies_others=True` makes the contract test require it.
- Results are pydantic models, so the output schema is published and validated.
- Lists take `limit` (≤25) and `cursor` and return `has_more` and `cursor`.
- Errors are raised as `ToolError` with a catalogue key.

## Tests per tool

Each tool has a test module with recorded Graph fixtures (`tests/fixtures/graph/<area>/…`) and:

- the happy path, asserting the exact Graph requests made (method, path, key params, headers) and
  the projected result;
- one test per row in the tool's **Must refuse** table, asserting `isError`, the message key, and
  that no write request reached Graph;
- for notifying tools, the guard refusals (no confirmation, recipients, bucket).

The live suite (`pytest -m live`) runs the same scenarios against the test tenant with two users
and is never run in CI against a customer tenant.
