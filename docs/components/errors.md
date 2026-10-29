# Component: errors

One exception type reaches the MCP layer: `ToolError(code, message_key, params)`. The SDK turns it
into a result with `isError: true`.

- `message_key` selects a person-facing sentence from a catalogue (`errors/en.toml`); other
  languages can be added as files. The sentence names the next step.
- `params` may contain counts, limits, seconds and candidate lists. They go to the model in the
  result. **They never go into the exception's `str()`**, which is what logs and tracebacks print;
  `str()` is the code alone.
- Graph errors arrive as `GraphError` and are mapped in one table (see the graph client). Anything
  unmapped becomes `unexpected` with the request id.
- A tool never returns an error as data, and never swallows an exception to return partial
  success silently. A partial batch (`update_messages` with 50 ids, 3 failed) returns success with
  `failed: [{id, reason}]` — that is a result, not an error, and the description says so.

## Tests

- Every `message_key` used in code exists in the catalogue (static scan).
- `str(ToolError)` contains no params; a log capture of a failing call contains no candidate
  names or addresses.
- An unmapped Graph error produces `unexpected` with the request id.
