"""Break each safety rule on purpose and require the tests to notice.

A green suite proves nothing about a guard unless the suite goes red without it. Each
mutation below removes one guard; the named tests must then fail. A mutation whose text is no
longer found is itself a failure, so a refactor cannot silently retire a check.

    python scripts/mutation_check.py
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src" / "m365_mcp"

# (what it protects, file, [(original, replacement), ...], test selector that must fail)
MUTATIONS = [
    ("the guard asks for confirmation", "core/guard.py",
     [("if needs_confirmation and user_confirmed is not True:", "if False:")],
     "tests/unit/test_core.py"),
    ("send_draft refuses without the person's word", "tools/mail.py",
     [('        if user_confirmed is not True:\n            raise await _unconfirmed(c, id)\n        d = await _own_draft(c, id)',
       "        d = await _own_draft(c, id)"),
      ('receipt = c.guard(kind="mail", user_confirmed=user_confirmed,', 'receipt = c.guard(kind="mail", user_confirmed=True,')],
     "tests/unit/test_mail.py"),
    ("only the boolean true is a confirmation", "server.py",
     [("if param in args and not isinstance(args[param], bool):", "if False:")],
     "tests/unit/test_confirmation.py"),
    ("only drafts made here are changed or sent", "tools/mail.py",
     [('if not any(p.get("id", "").lower() == DRAFT_MARK.lower() for p in d.get("singleValueExtendedProperties") or []):',
       "if False:")],
     "tests/unit/test_mail.py"),
    ("the send bucket", "core/guard.py", [("if b.tokens < 1:", "if False:")], "tests/unit/test_core.py"),
    ("link signatures are checked", "links/tokens.py",
     [("if not sig.isascii() or not any(hmac.compare_digest(self._mac(k, payload), sig) for k in self._sign):",
       "if False:")],
     "tests/unit/test_link_tokens.py"),
    # Two checks hold this (the operation field, and the kind's direction); removing one alone
    # is covered by the other, so both go.
    ("a download link cannot upload", "links/tokens.py",
     [('if body["o"] != op or body["k"] not in KINDS:', 'if body["k"] not in KINDS:'),
      ('            if (op == "d") != (body["k"] in DOWNLOAD_KINDS):\n                raise LinkRefused()\n', "")],
     "tests/unit/test_link_tokens.py"),
    ("outgoing HTML is sanitised", "core/html.py",
     [("    clean = nh3.clean(\n", "    clean = text if True else nh3.clean(\n")],
     "tests/unit/test_html.py"),
    ("hidden text is dropped", "core/html.py",
     [("    if s.get(\"display\") == \"none\" or s.get(\"visibility\") in (\"hidden\", \"collapse\"):\n        return True",
       "    pass")],
     "tests/unit/test_html.py"),
    ("text pushed off the edge or coloured transparent is dropped", "core/html.py",
     [("            self.first_line = len(self.stack) if pushed else None", "            self.first_line = None"),
      ("        if (self.first_line is not None or self.transparent) and not self.hidden and data.strip():",
       "        if self.first_line is not None and not self.hidden and data.strip():")],
     "tests/unit/test_html.py"),
    ("an element a display makes a block takes text-indent", "core/html.py",
     [(' or s.get("display") in _BLOCK_DISPLAY:', ":")], "tests/unit/test_html.py"),
    ("a text-indent in calc() or viewport units is read", "core/html.py",
     [('    value = re.sub(r"^\\s*calc', '    value = value or re.sub(r"^\\s*calc'),
      ('|ch|ex|%|vw|vh)?', '|ch|ex)?')], "tests/unit/test_html.py"),
    ("text Outlook hides with mso-hide is dropped", "core/html.py",
     [('    if s.get("mso-hide") == "all":', '    if False:')], "tests/unit/test_html.py"),
    ("a positioned box pushed off the page is dropped", "core/html.py",
     [('    if s.get("position") in ("absolute", "fixed") and any(_far_off(', '    if False and any(_far_off(')],
     "tests/unit/test_html.py"),
    ("text clipped to nothing is dropped", "core/html.py",
     [("    if _clipped(s):\n", "    if False:\n")], "tests/unit/test_html.py"),
    ("text in a font too small to read is dropped", "core/html.py",
     [('(_is_zero(s["font-size"]) or _tiny(s["font-size"]))', '_is_zero(s["font-size"])')], "tests/unit/test_html.py"),
    ("an outgoing hidden attribute stays hidden", "core/html.py",
     [("    clean, hidden = _hidden_to_style(clean)\n", "    hidden = 0\n")],
     "tests/unit/test_html.py"),
    ("foreign text cannot close its own content block", "core/html.py",
     [('<<<end of content {nonce}>>>"', '<<<end of content>>>"')],
     "tests/unit/test_html.py"),
    ("foreign text loses the brackets of a content marker", "core/html.py",
     [("    text = defang(text)\n", "")],
     "tests/unit/test_html.py"),
    ("text the full body hides stays hidden when uniqueBody lost the style", "tools/mail.py",
     [("and not (full and hides_text(full)) else", "else")],
     "tests/unit/test_mail.py"),
    ("a reply keeps the letter it quotes, even one written here", "tools/mail.py",
     [("insert_new_part(existing, prepared.html, part, replace=False)", "insert_new_part(existing, prepared.html)")],
     "tests/unit/test_mail.py"),
    ("search results carry immutable ids", "tools/mail.py",
     [('                m["id"] = ids.get(m.get("id"), m.get("id"))\n', "                pass\n")],
     "tests/unit/test_mail.py"),
    ("Safe Links are unwrapped", "core/html.py",
     [("        url = target\n", "        break\n")],
     "tests/unit/test_html.py"),
    ("tracking parameters are removed", "core/html.py",
     [('return n.startswith("utm_") or n in _TRACKING', "return False")],
     "tests/unit/test_html.py"),
    ("DELETE only on allowed paths", "graph/client.py",
     [('if method == "DELETE" and not any(', 'if False and not any(')],
     "tests/unit/test_graph_client.py"),
    ("a download redirect is followed only to Microsoft's file hosts", "graph/client.py",
     [("                if not on_file_host(location):\n", "                if not location.startswith(\"https://\"):\n")],
     "tests/unit/test_graph_client.py"),
    ("an upload session goes only to Microsoft's file hosts", "graph/client.py",
     [("        if not on_file_host(upload_url):\n", "        if not upload_url.startswith(\"https://\"):\n")],
     "tests/unit/test_graph_client.py"),
    ("ids cannot rewrite a Graph path", "graph/client.py",
     [("        _check_path(path_or_url)\n", "")],
     "tests/unit/test_path_injection.py"),
    ("only Graph hosts are called", "graph/client.py",
     [("if parts.hostname != GRAPH_HOST or parts.fragment:", "if False:")],
     "tests/unit/test_graph_client.py"),
    ("the HTTP stack does not log URLs", "graph/client.py",
     [("    logging.getLogger(_name).setLevel(logging.WARNING)", "    pass")],
     "tests/unit/test_core.py"),
    ("logs drop content", "observability.py", [("if _CONTENT.search(msg):", "if False:")], "tests/unit/test_core.py"),
    ("a logged traceback loses every exception message", "observability.py",
     [("            record.exc_text = _bare_traceback(record.exc_info)\n            record.exc_info = None\n", "            pass\n")],
     "tests/unit/test_core.py"),
    ("uvicorn's loggers go through the filtered root handlers", "app.py",
     [(", access_log=False, log_config=None)", ", access_log=False)")],
     "tests/unit/test_app.py"),
    ("sharing links are never anonymous-by-default", "tools/files.py",
     [('{"type": role, "scope": "organization"}', '{"type": role}')],
     "tests/unit/test_files.py"),
    ("deleting a folder needs confirmation", "tools/files.py",
     [("if folders and user_confirmed is not True:", "if False:")],
     "tests/unit/test_files.py"),
    ("only the organiser changes a meeting", "tools/calendar.py",
     [('if existing.get("isOrganizer") is not True:', "if False:")],
     "tests/unit/test_calendar.py"),
    ("an empty page ends a Teams list, whatever next link it carries", "tools/teams.py",
     [("        if not peek.items:\n            return None\n", "")],
     "tests/unit/test_teams.py"),
    ("a chat search reads one hit ahead before saying has_more", "tools/teams.py",
     [('            if ahead.get("hits"):\n', "            if True:\n")],
     "tests/unit/test_teams.py"),
    ("search snippets are fenced as someone else's text", "tools/teams.py",
     [('"snippet": content_block(name, snippet) if snippet else None', '"snippet": snippet or None')],
     "tests/unit/test_teams.py"),
    ("a mail subject is fenced as the sender's text", "tools/mail.py",
     [("    return content_line(sender[:60], m[\"subject\"])", "    return m[\"subject\"]")],
     "tests/unit/test_mail.py"),
    ("a one-line fence loses the text's marker brackets", "core/html.py",
     [("    text = \" \".join(defang(text).split())", "    text = \" \".join(text.split())")],
     "tests/unit/test_html.py"),
    ("an event subject is fenced as the organiser's text", "tools/calendar.py",
     [('    return content_line(who(e.get("organizer")) or "the organiser", e["subject"])', '    return e["subject"]')],
     "tests/unit/test_calendar.py"),
    ("a chat's topic is fenced", "tools/teams.py",
     [('        return content_line("a member of the chat", chat["topic"])', '        return chat["topic"]')],
     "tests/unit/test_teams.py"),
    ("task titles are fenced", "tools/tasks.py",
     [('                row["title"] = content_line(', '                row["title"] = row["title"] or content_line(')],
     "tests/unit/test_tasks.py"),
    ("OneNote page titles are fenced", "tools/notes.py",
     [('"title": content_line("a OneNote page", x["title"]) if x.get("title")', '"title": x["title"] if x.get("title")'),
      ('title = content_line("the OneNote page", found) if found', "title = found if found")],
     "tests/unit/test_notes.py"),
    ("list item fields are fenced", "tools/lists.py",
     [('content_block("the list", json.dumps(fields, ensure_ascii=False)) if fields else None', "fields or None")],
     "tests/unit/test_lists.py"),
    ("a folder renamed and then refused its move says it was renamed", "tools/mail.py",
     [('                if "renamed_to" not in result:\n', "                if True:\n")],
     "tests/unit/test_mail.py"),
    ("list_people hands out a personal contact's id", "tools/account.py",
     [('"contact_id": w.contact_id if w.source == "contacts" else None', '"contact_id": None')],
     "tests/unit/test_account.py"),
    ("a name that matches nothing says where to look", "core/errors.py",
     [('        out["where"] = WHERE.get(str(out.get("kind")), "Check the spelling, or list them to see the names.")',
       '        out["where"] = ""')],
     "tests/unit/test_refusals.py"),
    ("a backslash never reaches $search", "tools/mail.py",
     [('.replace("\\\\", " ")', "")],
     "tests/unit/test_mail.py"),
    ("one message id named twice is changed once", "tools/mail.py",
     [("        for mid in dict.fromkeys(ids):", "        for mid in ids:")],
     "tests/unit/test_mail.py"),
    ("an attachment name that matched nothing is reported", "tools/mail.py",
     [("        if remove_attachments and not_removed:\n", "        if False:\n")],
     "tests/unit/test_mail.py"),
    ("an envelope names its folder", "tools/mail.py",
     [("        await _folder_names(c, items)\n", "")],
     "tests/unit/test_mail.py"),
    ("a width attribute over 800 is warned", "core/html.py",
     [(" or any(w > 800 for w in after.widths)", "")],
     "tests/unit/test_html.py"),
    ("a download's upstream is released though the body never started", "links/routes.py",
     [("            await self._done()\n", "            pass\n")],
     "tests/unit/test_link_routes.py"),
    ("a client leaving mid-stream cannot cut the upstream's close short", "links/routes.py",
     [("        with anyio.CancelScope(shield=True):\n", "        with anyio.CancelScope():\n")],
     "tests/unit/test_link_routes.py"),
    ("a draft's upload links are spent only after a send that went through", "tools/mail.py",
     [('        try:\n            await c.graph.request("POST", f"/me/messages/{id}/send"',
       '        c.rt.signer.revoke_draft(c.key, id)\n        try:\n            await c.graph.request("POST", f"/me/messages/{id}/send"')],
     "tests/unit/test_link_routes_mail.py"),
    ("an automatic reply nobody asked to send outside stays inside where domains are limited", "tools/autoreply.py",
     [('audience = "none" if allowed else "contacts"', 'audience = "contacts"')],
     "tests/unit/test_recipient_domains.py"),
    ("HEAD answers headers without pulling the file", "links/routes.py",
     [('    if request.method == "HEAD":\n', "    if False:\n")],
     "tests/unit/test_link_routes.py"),
    ("throttling on a link reaches the sandbox as throttling, with its wait", "links/routes.py",
     [("    if e.status in (429, 503):\n        wait", "    if False:\n        wait")],
     "tests/unit/test_link_routes.py"),
    ("an upload link delivers one file, once", "links/routes.py",
     [("            signer.consume(value, link.expires)\n", "            signer.release(value)\n")],
     "tests/unit/test_link_routes.py"),
    ("one upload at a time on one link", "links/routes.py",
     [("if signer is None or not signer.claim(value):", "if signer is None:")],
     "tests/unit/test_link_routes.py"),
    ("running uploads share one spool budget", "links/routes.py",
     [("        if not budget.take(len(chunk)):\n", "        if False:\n")],
     "tests/unit/test_link_routes.py"),
    ("a spool the disk cannot hold gives its bytes back", "links/routes.py",
     [("            budget.give(len(chunk))\n            return refuse(", "            return refuse(")],
     "tests/unit/test_link_routes.py"),
    ("a stalled upload body is cut off", "links/routes.py",
     [("timeout=max(0.0, min(BODY_IDLE_TIMEOUT, left))", "timeout=None")],
     "tests/unit/test_link_routes.py"),
    ("an upload refused for its size names the limit", "links/routes.py",
     [('at most {limit} bytes "', 'at most some bytes "')],
     "tests/unit/test_link_routes_mail.py"),
    ("list_people continues past its first page", "tools/account.py",
     [('                more = c.cursor_out({"args": asked, "skip": skip + n})', "                more = None")],
     "tests/unit/test_account.py"),
    ("only your own Teams messages change", "tools/teams.py",
     [('if ((m.get("from") or {}).get("user") or {}).get("id") != me["id"]:', "if False:")],
     "tests/unit/test_teams.py"),
    ("Planner writes carry the etag", "tools/tasks.py",
     [('await c.graph.patch(f"/planner/tasks/{id}", fields, if_match=current.get("@odata.etag"))',
       'await c.graph.patch(f"/planner/tasks/{id}", fields)')],
     "tests/unit/test_tasks.py"),
]


class GuardNotFound(Exception):
    """The text a mutation removes is not in the file exactly once: the guard moved or was renamed."""


@contextmanager
def mutated(rel: str, edits: list[tuple[str, str]]) -> Iterator[None]:
    """Apply `edits` to one source file for the duration of the block, then put the file back
    byte for byte, whatever happens inside."""
    path = SRC / rel
    original = path.read_text()
    text = original
    for old, new in edits:
        if text.count(old) != 1:
            raise GuardNotFound(f"the guarded code was not found exactly once in {rel}: {old[:60]!r}")
        text = text.replace(old, new)
    try:
        path.write_text(text)
        yield
    finally:
        path.write_text(original)


def run(*selectors: str) -> int:
    """pytest's exit code: 0 passed, 1 a test failed. Anything else (2 collection error, 4 usage, 5 no
    tests) is not evidence that a test caught the mutation, so it is reported as its own problem.

    Each run compiles into a bytecode cache of its own. Python trusts a .pyc whose recorded source
    mtime (in whole seconds) and size match; two mutations of one file that change it by the same
    number of bytes within one second would otherwise run the previous mutant's bytecode, and a
    guard that is gone reads as caught, or one that is present as not caught."""
    with tempfile.TemporaryDirectory(prefix="mutation-pyc-") as cache:
        env = {**os.environ, "PYTHONPYCACHEPREFIX": cache}
        return subprocess.run([sys.executable, "-m", "pytest", "-q", "-x", "-p", "no:cacheprovider", *selectors],
                              cwd=ROOT, capture_output=True, text=True, env=env).returncode


def main() -> int:
    failures = []
    for what, rel, edits, selector in MUTATIONS:
        try:
            with mutated(rel, edits):
                code = run(selector)
        except GuardNotFound as e:
            failures.append(f"{what}: {e}")
            continue
        status = {0: "NOT CAUGHT", 1: "caught"}.get(code, f"BROKEN({code})")
        print(f"{status:10} {what}")
        if code == 0:
            failures.append(f"{what}: tests still pass with the guard removed")
        elif code != 1:
            failures.append(f"{what}: pytest exited {code}, which is not a failing test")
    if run("tests/unit/test_core.py") != 0:
        failures.append("the suite does not pass after restoring the code")
    print(f"\n{len(MUTATIONS)} mutations, {len(failures)} problems")
    for f in failures:
        print(" -", f)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
