# m365-mcp

[![validate](https://github.com/Wide-Moat/microsoft-365-mcp/actions/workflows/validate.yml/badge.svg)](https://github.com/Wide-Moat/microsoft-365-mcp/actions/workflows/validate.yml)
[![License: Apache-2.0](https://img.shields.io/badge/license-Apache--2.0-0c0c0d)](LICENSE)

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

Microsoft ships its own MCP servers for Microsoft 365. They are good at what they are built for, and
they are built for something different: Microsoft hosts them, prices them with Copilot, and exposes
Graph generically. This server is self-hosted, free, and shaped around the questions a person asks.
Facts as Microsoft and each project published them on 1 November 2026; every cell has a source below.
Corrections are welcome.

| | **m365-mcp** | [Work IQ MCP][workiq] (Microsoft) | [Mail, Calendar, Teams… servers][mail] (Microsoft, Copilot Studio) | [MCP Server for Enterprise][enterprise] (Microsoft) | [Softeria ms-365-mcp-server][softeria] (community) |
|---|---|---|---|---|---|
| Shape | 39 task-shaped tools (43 with every area on) | 12 generic tools over Graph resource paths (`fetch`, `create_entity`, `do_action`, …) | One tool per Graph operation, per workload (Mail: 10) | 3 tools: suggest a query, run a read-only `GET`, list properties | "300+ tools", one per Graph endpoint |
| Scope | One person's mail, calendar, files, lists, Excel, Teams, tasks, notes | Microsoft 365 data, Copilot and agents | One workload each | Entra identity and directory, **read-only** | Most of Graph |
| Status | Used in production | Generally available | Preview, "only intended for backward compatibility" | Preview | Community project |
| Hosting | Self-hosted container, inside your network | Microsoft-hosted endpoint, or a local CLI | Microsoft-hosted | Microsoft-hosted | Self-hosted (npm) |
| Cost | Free, Apache-2.0 | Usage-based, in Copilot Credits | Copilot Studio | No extra cost | Free, MIT |
| Writes | On, each behind its own guard (below) | Blocked by default by tenant policy; an administrator enables them | Direct (`sendMail` sends) | None | On, or `--read-only` |
| Sending mail | Two calls: `save_draft`, then `send_draft` with `user_confirmed` (a flag the client sets after asking the person); only drafts this server created can be sent; 20 recipients at most, no distribution lists | `do_action` on `/me/sendMail` | `sendMail`, or `sendDraft` | n/a | Not described |
| Files | Streamed through signed, short-lived links (15 minutes at most); bytes never pass through the model | `fetch_blob`: the whole file as Base64 in the result, 4 MB by default | Not described | n/a | Single-use download tickets |
| Throttling | Honours Graph's `Retry-After`, bounded retries | No automatic retries; the client decides | Not described | 100 calls a minute per user | Per-IP rate limit on its HTTP endpoint |
| HTML mail | Cleaned both ways: CSS inlined, active content removed, `data:` images become attachments; hidden text dropped on read | Not described | HTML passed as given | n/a | Body returned as HTML on request |
| Text from other people | Marked as data, not instructions | Not described | Not described | n/a | Not described |

[workiq]: https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/mcp/overview
[workiq-tools]: https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/mcp/tool-reference
[workiq-api]: https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/api-overview#licensing-requirements
[agent365]: https://learn.microsoft.com/en-us/microsoft-agent-365/tooling-servers-overview
[mail]: https://learn.microsoft.com/en-us/microsoft-copilot-studio/mcp-mail-tools
[enterprise]: https://learn.microsoft.com/en-us/graph/mcp-server/overview
[softeria]: https://github.com/Softeria/ms-365-mcp-server

Sources:

- **Work IQ MCP.** The tool list, writes "blocked by default" by tenant policy, `fetch_blob`'s Base64
  and its 4 MB default, and "The server doesn't perform automatic retries" are from the
  [tool reference][workiq-tools]. The design ("Fewer tools, more paths") is from the
  [overview][workiq]. Billing ("usage-based model that uses Copilot Credits") is from the
  [Work IQ API licensing section][workiq-api]. General availability is from the
  [Agent 365 overview][agent365].
- **The workload servers (Mail, Calendar, Teams, …).** The Mail tools, the preview status and "only
  intended for backward compatibility purposes and shouldn't be used in production environments or new
  projects" are from the [Mail reference][mail]; the same note is in the [Agent 365 overview][agent365].
- **MCP Server for Enterprise.** The tools, the read-only scope, "no additional cost", 100 calls a minute
  per user and the preview status are from its [overview][enterprise].
- **Softeria.** "300+ tools", "maps 1-to-1 to a Graph API endpoint", read-only mode, the HTTP rate limit
  and the MIT licence are from its [README][softeria].
- **m365-mcp.** Every claim is a test in this repository: see [How it is tested](#how-it-is-tested) and
  [`docs/requirements.md`](docs/requirements.md).

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

## Contributing and security

[CONTRIBUTING.md](CONTRIBUTING.md) says how a change is made and checked; report a vulnerability
privately as [SECURITY.md](SECURITY.md) describes, never in a public issue.

## Licence

[Apache-2.0](LICENSE).
