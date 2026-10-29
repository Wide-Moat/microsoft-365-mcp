# m365-mcp

An MCP server that gives an AI assistant one person's Microsoft 365 through Microsoft Graph. It
covers mail, calendar and scheduling, contacts, OneDrive and SharePoint files, SharePoint lists,
Excel, Teams, meeting recordings, To Do and Planner, and OneNote. It acts as that person, with their
own delegated sign-in and nothing more.

We built it for our own agents, after running them against Microsoft 365 every day. It is small on
purpose:

- **39 tools, about 14,800 tokens** of catalogue by default (43 tools and about 16,200 tokens with
  every area on). There is one tool per question a person actually asks, not one per Graph
  endpoint.
- **The server does the work Graph leaves to the caller:**
  - it resolves names;
  - it expands recurrences;
  - it cleans HTML both ways;
  - it streams files through short-lived signed links instead of the model's context.
- **It refuses what must not happen.** Every refusal is written down and has a test that fails when
  its guard is removed.

The rules every change is held to are in the [manifesto](MANIFESTO.md).

## What it does

| Area | Tools | What a person can ask |
|---|---|---|
| Account | `get_connection`, `list_people` | Who am I connected as, what is switched on, which permissions are missing; find a colleague by part of a name |
| Mail | `list_messages`, `get_message`, `save_draft`, `send_draft`, `update_messages`, `save_mail_folder`, `save_automatic_reply` | Find and read mail and threads with attachments; write HTML letters that survive Outlook; send only after a draft was shown; move, flag or delete to Deleted Items |
| Calendar | `list_events`, `get_event`, `save_event`, `respond_to_event`, `cancel_event`, `find_meeting_times` | Read the calendar with recurrences expanded; book, move and cancel meetings; find when people and rooms are free |
| Contacts | `save_contact`, `delete_contact` | Keep the address book |
| Files | `list_files`, `get_file`, `download_file`, `upload_file`, `update_file`, `share_file` | OneDrive and SharePoint: search, read, upload up to 250 MB, convert to PDF, move, copy, share |
| Lists | `list_list_items`, `save_list_item`, `delete_list_item` | SharePoint lists as rows |
| Excel | `get_range`, `save_range` | Read and write cells of a workbook in place |
| Teams | `list_conversations`, `get_conversation`, `send_message`, `update_message`, `save_chat` | Read chats and channels; post, edit and delete messages |
| Tasks | `list_tasks`, `save_task`, `delete_task` | To Do and Planner |
| Notes | `list_notes`, `get_note_page`, `save_note_page`, `delete_note_page` | OneNote |
| Bookings (off by default) | `list_bookings`, `save_booking`, `cancel_booking` | Bookings appointments |

Every tool, with its arguments, is in [`docs/tools.md`](docs/tools.md), which is generated from the
code.

### What it will not do

- It never touches inbox rules, forwarding, delegates, permissions, export, or another person's
  mailbox. There is no `mailbox` argument anywhere.
- It never hard-deletes. Deleted mail goes to Deleted Items and deleted files go to the recycle bin.
- Sending is two steps: the model writes a draft, and the server sends only a draft it created
  itself, after the person confirms.
- Every action that notifies other people is limited:
  - at most 20 recipients per message;
  - no distribution lists;
  - a send allowance per person;
  - an optional list of allowed recipient domains.
- Text written by other people comes back marked as data, not instructions. Hidden HTML
  (zero-size, off-screen, same colour as the background) is dropped before the model sees it.
- The token is never logged, echoed or stored. A test reads the logs of a full mail flow and fails
  if any subject, address or body appears in them.

## How it is tested

| Check | What it proves | Size |
|---|---|---|
| Unit and contract tests (`pytest`) | Every tool against a recorded fake of Graph: arguments, results, errors, refusals | about 1,700 tests |
| [Refusal map](docs/requirements.md) | Every "must refuse" row names an existing test | 200 refusal tests |
| Refusal evidence (`scripts/refusal_evidence.py`) | Each row's guard is removed in turn, and its test must fail | every row |
| Mutation check (`scripts/mutation_check.py`) | Each safety rule is removed in turn, and some test must fail | 69 mutations |
| Token budget (`python -m m365_mcp.budget --check`) | The default catalogue stays under 16,000 tokens; every area under 3,500 and every tool under 600 | every pull request |
| Tool surface lock | A changed argument, enum or annotation is visible in review | every pull request |
| Live suite (`pytest -m live`, [how](docs/live-testing.md)) | The same tools against a real Microsoft 365 tenant with two people. It also measures what Microsoft's documentation does not say ([findings](docs/graph-live-findings.md)) | 37 probes, on demand |
| Secret scan | gitleaks over every commit | every push |
| Purity (`scripts/purity_check.py`) | No real address, Microsoft 365 tenant name or public IP in any tracked file | every push |

## How it compares

Facts as published by each project on 29 October 2026, with sources. Corrections are welcome.

| | **m365-mcp** | [Softeria ms-365-mcp-server][softeria] | [Work IQ MCP][workiq] (Microsoft) | [MCP Server for Enterprise][enterprise] (Microsoft) |
|---|---|---|---|---|
| Shape | 39 task-shaped tools (43 with every area on) | "300+ tools", one per Graph endpoint | 12 generic tools over resource paths (`fetch`, `create_entity`, `do_action`, …) | 3 tools: suggest a query, run a read-only `GET`, list properties |
| Scope | One person's mail, calendar, files, lists, Excel, Teams, tasks, notes | Most of Graph | Microsoft 365 data and Copilot | Entra identity and directory, **read-only** |
| Hosting | Self-hosted container | Self-hosted (npm) | Microsoft-hosted endpoint, or a local CLI | Microsoft-hosted only |
| Cost | Free, Apache-2.0 | Free, MIT | Usage-based, paid in Copilot Credits | No extra cost |
| Sending mail | Draft, then send after confirmation; only drafts the server created | Not described | `do_action /me/sendMail` | Not available (read-only) |
| HTML mail | Cleaned both ways: CSS inlined, active content removed, `data:` images become attachments | Body returned as HTML on request | Not described | n/a |
| Files | Streamed through signed links of at most 15 minutes, never through the model | Attachment download through single-use tickets | `fetch_blob` | n/a |
| Status | Used in production | Community project | Generally available (Copilot Studio integration in preview) | Preview |

[softeria]: https://github.com/Softeria/ms-365-mcp-server
[workiq]: https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/mcp/overview
[enterprise]: https://learn.microsoft.com/en-us/graph/mcp-server/overview

Sources for the comparison:

- **Softeria:** the [README][softeria] ("300+ tools", "maps 1-to-1 to a Graph API endpoint", the
  MIT licence).
- **Work IQ:** the tool list is from the [MCP overview][workiq]. Billing ("usage-based model that
  uses Copilot Credits") is from the
  [Work IQ API licensing section](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/api-overview#licensing-requirements).
  The availability status is from the
  [Agent 365 overview](https://learn.microsoft.com/en-us/microsoft-agent-365/tooling-servers-overview).
- **MCP Server for Enterprise:** the [overview][enterprise] (tools, the read-only scope, "no
  additional cost", the preview status).

## Run

```sh
docker build -t m365-mcp .
docker run -p 8000:8000 -p 8001:8001 \
  -e M365_LINK_SIGNING_KEYS=<base64url, 32+ bytes> \
  -e M365_LINK_ENCRYPTION_KEYS=<base64url, 32+ bytes> \
  -e M365_LINKS_BASE_URL=https://<public host serving port 8001> \
  m365-mcp
```

The two ports:

- **`:8000/mcp`** is the MCP endpoint (Streamable HTTP, stateless). Each request carries
  `Authorization: Bearer <the person's Microsoft Graph token>`, supplied by an MCP gateway that does
  OAuth for each person. An optional `X-M365-Areas: mail,calendar` header narrows the tools for one
  person.
- **`:8001/d/…` and `:8001/u/…`** are the signed download and upload links for the agent's sandbox.

The settings:

- `M365_AREAS` and `M365_SWITCHES` choose which areas and switches are on (see
  [`docs/tools.md`](docs/tools.md)).
- `M365_EXPECTED_CLIENT_IDS` is optional: the gateway's application (client) ids, comma-separated.
  When set, a Graph token issued to any other application is refused
  ([identity](docs/components/identity.md)).
- `M365_ALLOWED_RECIPIENT_DOMAINS` is optional: the domains that mail, invitations, file invitations
  and Teams messages may reach, comma-separated ([send guard](docs/components/guard.md)).

The Graph permissions it asks for, and why, are in
[`docs/requirements.md`](docs/requirements.md#permissions).

## Develop

```sh
uv venv -p 3.12 .venv && uv pip install -p .venv -r requirements-test.lock && uv pip install -p .venv --no-deps -e .
.venv/bin/pytest -q                            # unit and contract tests
.venv/bin/python -m m365_mcp.budget            # what the catalogue costs, per tool and area
.venv/bin/python scripts/mutation_check.py     # every safety rule fails a test when removed
.venv/bin/python scripts/refusal_map.py        # every Must refuse row names an existing test
.venv/bin/python scripts/refusal_evidence.py   # and that test fails with the row's guard removed
.venv/bin/python -m m365_mcp.docs > docs/tools.md
.venv/bin/python -m m365_mcp.surface > tool-surface.lock.json  # after an intended change to the tool surface
```

| Document | What it holds |
|---|---|
| [`MANIFESTO.md`](MANIFESTO.md) | The rules, each with what a violation looks like |
| [`docs/requirements.md`](docs/requirements.md) | Contracts and every negative scenario |
| [`docs/architecture.md`](docs/architecture.md) | Components, token path, byte path, threats |
| [`docs/components/`](docs/components) | One document per component |
| [`docs/graph-reference.md`](docs/graph-reference.md) | What Microsoft Graph documents, with sources and unknowns |

## Who made it

[Wide Moat](https://widemoat.ai) runs AI agents for companies on their own infrastructure: chat,
coding workspaces, a model gateway with per-person keys and connectors like this one, all in one
platform on your cluster. This server is one part of it.

If you want the platform, not only the connector, write to us through the site.

## Licence

[Apache-2.0](LICENSE).
