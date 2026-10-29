"""Each Must refuse row's test, seen failing with that row's guard removed.

refusal_map.py proves every row names a test that exists. That says nothing about whether the test
would notice the refusal disappearing: a test asserting only `isError` passes against the wrong
error. So for every test a row names, this removes the code that refuses (the mutation below,
applied with mutation_check's mechanism), runs that one test, and requires pytest to exit 1. The
file is restored after each run; the baseline run first requires every mapped test to pass on the
untouched code, so a test that always fails cannot pass for evidence.

Where two checks hold one refusal (a tool's own limit and the guard's), the mutation removes both,
or the remaining one would hide the missing one.

    python scripts/refusal_evidence.py               # every row
    python scripts/refusal_evidence.py test_mail.py  # only ids containing this text
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from mutation_check import GuardNotFound, mutated, run  # noqa: E402
from refusal_map import ROOT, TEST_ROOT, check, collected, documents, mapped_tests, parse  # noqa: E402

# Shared removals, named once.
CONFIRM_GUARD = ("core/guard.py", [("if needs_confirmation and user_confirmed is not True:", "if False:")])
RECIPIENT_CAP = ("if recipients > MAX_RECIPIENTS:", "if False:")
BUCKET = ("core/guard.py", [("if b.tokens < 1:", "if False:")])
OWN_DRAFT = ("tools/mail.py", [('if not any(p.get("id", "").lower() == DRAFT_MARK.lower() for p in d.get("singleValueExtendedProperties") or []):',
                                "if False:")])
SPOOL_LIMIT = ("links/routes.py", [("if limit is not None and size + len(chunk) > limit:", "if False:")])
CLIENT_CHECK = ("identity.py", [("if clients and _client(c) not in clients:", "if False:")])
GATE_REFUSAL = ("server.py", [("if (params.arguments or {}).get(param) not in (None, False):", "if False:")])
NOT_A_PERSON = ("core/resolve.py", [
    ('return Who(name=known.name if known else "", address=address, is_group=True, source="directory")',
     'return Who(name=known.name if known else "", address=address, is_group=False, source="directory")')])
SHARED_NAME = ("core/resolve.py", [("    if len(found) > 1:\n", "    if False:\n")])
DOMAIN_CHECK = ("core/guard.py", [("        if outside:\n            raise M365Error(\"outside_domains\"", "        if False:\n            raise M365Error(\"outside_domains\"")])

CONFIRM_TYPE = ("server.py", [("if param in args and not isinstance(args[param], bool):", "if False:")])

UPDATE_MEMBERS = ("tools/teams.py", [('if action != "delete" and c.rt.settings.allowed_recipient_domains:', "if False:")])

UPLOAD_RUNNING = ("tools/mail.py", [("    if c.rt.signer is not None and c.rt.signer.draft_busy(c.key, draft_id):\n", "    if False:\n")])

LIST_CURSOR = ("tools/teams.py", [("if state is not None and _cursor_list(state) != _list_of(kind, team, query, unread_only):",
                                   "if False:")])

# test id (relative to tests/unit) -> (file under src/m365_mcp, [(original, replacement), ...])
MUTATIONS: dict[str, tuple[str, list[tuple[str, str]]]] = {
    # -- every tool ---------------------------------------------------------------------------
    "test_account.py::test_no_token_means_reconnect_and_no_graph_call": ("identity.py", [
        ('    if not token:\n        raise M365Error("reconnect")',
         '    if not token:\n        return Person(tenant="t", object_id="o", scopes=frozenset(), expires=2**40, token="")')]),
    "test_account.py::test_expired_token_means_reconnect": ("identity.py", [
        (" or not isinstance(exp, int) or exp <= now:", " or not isinstance(exp, int):")]),
    "test_identity.py::test_a_token_for_another_audience_is_refused_before_graph": ("identity.py", [
        ("if not isinstance(aud, str) or aud not in GRAPH_AUDIENCES or ", "if ")]),
    "test_identity.py::test_a_token_issued_to_another_client_is_refused_before_graph": CLIENT_CHECK,
    "test_refusals.py::test_a_missing_permission_is_named_with_the_administrator_step": ("graph/errors.py", [
        ('        if e.missing_scope:\n            return M365Error("missing_permission", permission=e.missing_scope)\n', "")]),
    "test_mail.py::test_not_found_does_not_say_which": ("graph/errors.py", [
        ('if e.status == 404 or code in {"ErrorItemNotFound", "itemNotFound", "ResourceNotFound"}:', "if False:")]),
    "test_refusals.py::test_throttling_after_retries_says_how_long_to_wait": ("graph/errors.py", [
        ("if e.status in (429, 503):", "if False:")]),
    "test_refusals.py::test_a_refused_write_says_nothing_changed_with_a_reference": ("graph/errors.py", [
        ('return M365Error("failed", request_id=e.request_id or "none")', 'return M365Error("failed", request_id="none")')]),
    "test_refusals.py::test_a_server_failure_on_a_write_never_claims_nothing_changed": ("graph/errors.py", [
        ("if e.status >= 500:", "if False:")]),
    "test_contract.py::test_a_parameter_behind_an_off_switch_is_hidden_and_refused": GATE_REFUSAL,
    # The person's area header is the one input only this test exercises; removing it from the gate
    # both shows the parameter and lets the call through.
    "test_contract.py::test_a_parameter_of_an_area_off_for_the_person_is_hidden_and_refused": ("server.py", [
        (" or (narrowed is not None and gate not in narrowed)", "")]),
    "test_refusals.py::test_a_malformed_id_is_not_found": ("graph/errors.py", [
        ("    if e.status == 400 and (code in INVALID_ID_CODES or inner in INVALID_ID_CODES\n", "    if False and (code in INVALID_ID_CODES or inner in INVALID_ID_CODES\n")]),
    "test_refusals.py::test_an_empty_id_is_refused_before_graph": ("graph/client.py", [
        ('    if "//" in path or path.endswith("/"):\n        raise UnsafePath()\n', "")]),
    "test_refusals.py::test_a_value_outside_an_enum_names_the_allowed_values": ("server.py", [
        ("            if isinstance(cause, ValidationError):\n", "            if False:\n")]),
    "test_refusals.py::test_an_unprovisioned_onedrive_is_not_a_sign_in_problem": ("graph/errors.py", [
        ('if e.status == 403 and inner == "provisioningNotAllowed":', "if False:")]),
    "test_refusals.py::test_a_service_401_with_a_valid_token_does_not_say_reconnect": ("graph/errors.py", [
        ("if e.status == 401 and code.casefold() not in TOKEN_CODES:", "if False:")]),
    "test_refusals.py::test_a_server_failure_on_a_read_says_nothing_changed": ("graph/errors.py", [
        ("        if read_only:\n", "        if False:\n")]),
    # -- identity -----------------------------------------------------------------------------
    "test_identity.py::test_a_token_that_is_not_a_jwt_is_refused_before_graph": ("identity.py", [
        ('    if len(parts) != 3:\n        raise M365Error("reconnect")\n', ""),
        ('    if not isinstance(claims, dict):\n        raise M365Error("reconnect")\n', "")]),
    "test_identity.py::test_a_token_without_a_tenant_or_object_id_is_refused": ("identity.py", [
        ("if not (isinstance(tid, str) and tid and isinstance(oid, str) and oid):", "if False:")]),
    "test_identity.py::test_a_token_naming_no_client_is_refused_when_clients_are_configured": CLIENT_CHECK,
    "test_identity.py::test_a_v1_token_from_another_client_is_refused": CLIENT_CHECK,
    # -- the send guard -----------------------------------------------------------------------
    "test_core.py::test_guard_requires_confirmation": CONFIRM_GUARD,
    "test_core.py::test_guard_caps_recipients_and_refuses_lists": ("core/guard.py", [RECIPIENT_CAP]),
    "test_core.py::test_bucket_bursts_ten_then_refills_one_per_six_seconds": BUCKET,
    "test_recipient_domains.py::test_guard_refuses_an_address_outside_the_allowed_domains": DOMAIN_CHECK,
    "test_recipient_domains.py::test_a_lookalike_domain_is_not_a_subdomain": ("core/guard.py", [
        ('domain.endswith("." + d)', "domain.endswith(d)")]),
    "test_recipient_domains.py::test_an_address_without_a_domain_is_refused": ("core/guard.py", [
        ('bad.add(domain or "(no domain)")', "bad.add(domain) if domain else None")]),
    # -- people and contacts ------------------------------------------------------------------
    "test_account.py::test_list_people_refuses_an_empty_query": ("tools/account.py", [
        ("    if not query.strip():\n        raise invalid(", "    if False:\n        raise invalid(")]),
    # Two checks hold it: an address is refused before Graph, a directory id after Graph's 404.
    "test_contacts.py::test_a_directory_person_is_not_a_contact": ("tools/contacts.py", [
        ('    if "@" in id:\n        raise invalid(NOT_A_CONTACT)\n', ""),
        ("        if e.status in (400, 404):\n            raise invalid(NOT_A_CONTACT) from None\n", "")]),
    "test_contacts.py::test_a_contact_in_deleted_items_is_not_changed_or_deleted_again": ("tools/contacts.py", [
        ("        raise invalid(DELETED)\n    return current", "    return current")]),
    "test_contacts.py::test_a_contact_whose_folder_cannot_be_placed_is_not_changed": ("tools/contacts.py", [
        ("        if not whole or not parent:\n            # In none of the folders read", "        if False:\n            # In none of the folders read")]),
    "test_contacts.py::test_an_email_that_is_not_an_address_is_refused": ("tools/contacts.py", [
        ("    if bad:\n", "    if False:\n")]),
    "test_contacts.py::test_a_fourth_email_names_outlooks_limit_of_three": ("tools/contacts.py", [
        ("            if len(merged) > MAX_EMAILS:\n", "            if False:\n")]),
    "test_contacts.py::test_delete_needs_confirmation": ("tools/contacts.py", [
        ('require(user_confirmed, "Deleting a contact")', "pass")]),
    # -- mail ---------------------------------------------------------------------------------
    "test_mail.py::test_system_folder_is_not_renamed": ("tools/mail.py", [
        ('raise M365Error("system_folder", name=f.get("displayName", folder))', "pass")]),
    "test_mail.py::test_a_folder_name_that_exists_names_the_existing_folder": ("tools/mail.py", [
        ("                if e.status == 409:\n", "                if False:\n")]),
    "test_mail.py::test_search_uses_kql_and_no_orderby": ("tools/mail.py", [
        ("""params["$search"] = '"' + " ".join(terms) + '"'""",
         """params["$search"] = '"' + " ".join(terms) + '"'; params["$orderby"] = "receivedDateTime desc\"""")]),
    "test_mail.py::test_empty_window_is_refused": ("tools/mail.py", [("if s and u and u <= s:", "if False:")]),
    "test_mail.py::test_an_ambiguous_sender_lists_candidates": ("tools/mail.py", [
        ('terms.append(f"from:{_kql(await address_of(c, sender))}")', 'terms.append(f"from:{_kql(sender)}")')]),
    "test_mail.py::test_an_ambiguous_recipient_filter_lists_candidates": ("tools/mail.py", [
        ('terms.append(f"to:{_kql(await address_of(c, recipient))}")', 'terms.append(f"to:{_kql(recipient)}")')]),
    "test_mail.py::test_more_than_50_ids_is_refused": ("tools/mail.py", [("if len(ids) > 50:", "if False:")]),
    "test_mail.py::test_move_without_folder_is_refused": ("tools/mail.py", [('if action == "move" and not folder:', "if False:")]),
    "test_mail.py::test_move_to_a_folder_that_does_not_resolve_is_refused": ("tools/mail.py", [
        ('dest = (await mail_folder(c, folder or ""))["id"]', "dest = folder")]),
    "test_mail.py::test_a_message_the_person_cannot_see_is_not_found": ("tools/mail.py", [
        ('await c.graph.patch(f"/me/messages/{mid}", patch)', 'await c.graph.patch(f"/users/owner/messages/{mid}", patch)')]),
    "test_mail.py::test_changing_a_draft_the_person_wrote_is_refused": OWN_DRAFT,
    "test_mail.py::test_more_than_20_recipients_is_refused": ("tools/mail.py", [("    if count > 20:\n", "    if False:\n")]),
    "test_mail.py::test_distribution_list_is_refused": ("tools/mail.py", [
        ('    if groups:\n        raise M365Error("distribution_list", address=groups[0])\n    return t, cc_, b',
         "    return t, cc_, b")]),
    "test_mail.py::test_reply_without_source_is_refused": ("tools/mail.py", [
        ('if mode in ("reply", "reply_all", "forward") and not source_id and not id:', "if False:")]),
    "test_mail.py::test_long_inline_html_must_be_uploaded": ("tools/mail.py", [
        ('if body and body_format == "html" and len(body) > MAX_INLINE_BODY:', "if False:")]),
    "test_link_routes_mail.py::test_draft_body_not_utf8_is_rejected": ("core/html.py", [
        ('            out.errors.append("The HTML must be UTF-8.")\n            return out\n',
         '            text = raw.decode("latin-1")\n')]),
    "test_link_routes_mail.py::test_draft_body_wrong_content_type_is_415": ("links/routes.py", [
        ("if link.content_type and ctype != link.content_type:", "if False:")]),
    "test_link_routes_mail.py::test_draft_body_over_the_limit_is_413": SPOOL_LIMIT,
    "test_mail.py::test_relative_urls_refuse_the_draft": ("core/html.py", [
        ('out.errors.append(f"Links and images must be absolute https URLs: {u[:200]}")', "pass")]),
    "test_link_routes_mail.py::test_an_attachment_past_the_message_size_limit_is_refused": ("graph/errors.py", [
        ('if e.status == 413 or code in {"ErrorMessageSizeExceeded", "RequestBodyTooLarge"}:', "if False:")]),
    "test_link_routes_mail.py::test_an_attachment_over_the_link_limit_is_413": SPOOL_LIMIT,
    "test_link_routes_mail.py::test_draft_attachment_links_carry_the_150_mb_limit": ("tools/mail.py", [
        ("target={'d': draft_id}, max_bytes=MAX_ATTACHMENT)", "target={'d': draft_id}, max_bytes=None)")]),
    "test_mail.py::test_send_without_confirmation_sends_nothing": ("tools/mail.py", [
        ('        if user_confirmed is not True:\n            raise await _unconfirmed(c, id)\n        d = await _own_draft(c, id)',
         "        d = await _own_draft(c, id)"),
        ('receipt = c.guard(kind="mail", user_confirmed=user_confirmed,', 'receipt = c.guard(kind="mail", user_confirmed=True,')]),
    "test_mail.py::test_send_refuses_a_draft_the_person_wrote": OWN_DRAFT,
    "test_mail.py::test_send_refuses_a_message_that_is_no_longer_a_draft": ("tools/mail.py", [
        ('    if not d.get("isDraft"):\n        raise invalid("That message is not a draft any more: it was sent', '    if False:\n        raise invalid("That message is not a draft any more: it was sent')]),
    "test_mail.py::test_send_refuses_a_list_added_after_the_draft": ("tools/mail.py", [
        ("group_addresses=groups_among(resolved), addresses=addrs)", "group_addresses=None, addresses=addrs)")]),
    "test_mail.py::test_send_refuses_recipients_grown_past_20": ("core/guard.py", [RECIPIENT_CAP]),
    "test_mail.py::test_send_refuses_a_draft_already_sent": ("tools/mail.py", [
        ("    if _mark_value(d) == SENT_MARK:\n", "    if False:\n")]),
    "test_mail.py::test_a_second_send_of_one_draft_sends_nothing": ("tools/mail.py", [
        ("    if _mark_value(d) == SENT_MARK:\n", "    if False:\n")]),
    "test_mail.py::test_bucket_stops_the_eleventh_send": BUCKET,
    "test_recipient_domains.py::test_send_draft_refuses_a_recipient_outside_the_allowed_domains": ("tools/mail.py", [
        ("group_addresses=groups_among(resolved), addresses=addrs)", "group_addresses=groups_among(resolved))")]),
    "test_autoreply.py::test_scheduled_without_both_dates_is_refused": ("tools/autoreply.py", [
        ('if status == "scheduled" and not (start and end):', "if False:")]),
    "test_autoreply.py::test_an_end_in_the_past_is_refused": ("tools/autoreply.py", [
        ("if e <= datetime.now(UTC):", "if False:")]),
    "test_autoreply.py::test_turning_it_on_without_confirmation_writes_nothing": ("tools/autoreply.py", [
        ('c.guard(kind="automatic_reply", user_confirmed=user_confirmed)', 'c.guard(kind="automatic_reply", user_confirmed=True)')]),
    "test_autoreply.py::test_scheduling_without_confirmation_writes_nothing": ("tools/autoreply.py", [
        ('c.guard(kind="automatic_reply", user_confirmed=user_confirmed)', 'c.guard(kind="automatic_reply", user_confirmed=True)')]),
    # -- calendar -----------------------------------------------------------------------------
    "test_calendar.py::test_changing_an_event_someone_else_organises_is_refused": ("tools/calendar.py", [
        ('if existing.get("isOrganizer") is not True:', "if False:")]),
    "test_calendar.py::test_inviting_without_confirmation_writes_nothing": ("tools/calendar.py", [
        ('c.guard(kind="event", user_confirmed=user_confirmed, recipients=invited,\n',
         'c.guard(kind="event", user_confirmed=True, recipients=invited,\n')]),
    "test_calendar.py::test_more_than_20_attendees_are_refused": ("tools/calendar.py", [
        ("if attendees is not None and len(attendees) > MAX_ATTENDEES:", "if False:"),
        ("recipients=invited,\n", "recipients=0,\n")]),
    "test_calendar.py::test_a_distribution_list_is_refused": ("tools/calendar.py", [
        ('            if groups:\n                raise M365Error("distribution_list", address=groups[0])\n', ""),
        ("group_addresses=groups_among(found), addresses=notified)", "group_addresses=None, addresses=notified)")]),
    "test_calendar.py::test_zero_or_negative_duration_is_refused": ("tools/calendar.py", [
        ("if duration_minutes is not None and duration_minutes <= 0:", "if False:")]),
    "test_calendar.py::test_a_recurring_series_cannot_be_changed": ("tools/calendar.py", [
        ('if existing.get("type") == "seriesMaster":', "if False:")]),
    "test_recipient_domains.py::test_inviting_an_attendee_outside_the_allowed_domains_is_refused": ("tools/calendar.py", [
        ("group_addresses=groups_among(found), addresses=notified)", "group_addresses=groups_among(found))")]),
    "test_calendar.py::test_find_meeting_times_more_than_20_people_or_a_group_is_refused": ("tools/calendar.py", [
        ("    if len(names) > MAX_ATTENDEES:\n", "    if False:\n"),
        ('    if groups:\n        raise M365Error("distribution_list", address=groups[0])\n    return found',
         "    return found")]),
    "test_calendar.py::test_find_meeting_times_window_over_62_days_is_refused": ("tools/calendar.py", [
        ("        if e <= s:\n", "        if False:\n"),
        ("if e - s > timedelta(days=MAX_WINDOW_DAYS):", "if False:")]),
    "test_calendar.py::test_responding_to_a_cancelled_event_is_refused": ("tools/calendar.py", [
        ('            raise M365Error("cancelled_event")\n', "            pass\n")]),
    "test_calendar.py::test_the_organiser_does_not_respond": ("tools/calendar.py", [
        ('if e.get("isOrganizer") is True:', "if False:")]),
    "test_calendar.py::test_a_comment_to_the_organiser_needs_confirmation": ("tools/calendar.py", [
        ('c.guard(kind="event_response", user_confirmed=user_confirmed, recipients=1,',
         'c.guard(kind="event_response", user_confirmed=True, recipients=1,')]),
    "test_calendar.py::test_a_comment_longer_than_the_cap_is_refused": ("tools/calendar.py", [
        ("    if comment and len(comment) > MAX_RESPONSE_COMMENT:\n", "    if False:\n")]),
    "test_recipient_domains.py::test_a_comment_to_an_organiser_outside_the_allowed_domains_is_refused": ("tools/calendar.py", [
        ('addresses=[org] if org else [""])', "addresses=None)")]),
    # -- recordings ---------------------------------------------------------------------------
    "test_recordings.py::test_an_event_that_is_not_a_teams_meeting_is_refused": ("tools/recordings.py", [
        ('    if not join_url:\n        raise invalid("That event is not a Teams meeting',
         '    if False:\n        raise invalid("That event is not a Teams meeting')]),
    "test_recordings.py::test_no_recording_yet_is_a_sentence_not_an_error": ("tools/recordings.py", [
        ('if not out["recording_available"] and not notes:', "if False:")]),
    "test_recordings.py::test_a_participant_without_download_rights_is_told_and_given_the_meeting_link": ("tools/recordings.py", [
        ("if e.status == 403 and not e.missing_scope:", "if False:")]),
    "test_recordings.py::test_an_expired_recording_says_it_expired": ("tools/recordings.py", [
        ('        if _expired(e):\n            notes.append("The recording has expired',
         '        if False:\n            notes.append("The recording has expired')]),
    # -- files --------------------------------------------------------------------------------
    "test_files.py::test_download_of_a_folder_is_refused": ("tools/files.py", [
        ('raise invalid(f"{name} is a folder. List it with list_files and download the files inside.")', "pass")]),
    "test_files.py::test_pdf_of_an_unsupported_type_is_refused_with_the_supported_ones": ("tools/files.py", [
        ('if format == "pdf" and ext not in PDF_TYPES:', "if False:")]),
    "test_files.py::test_site_location_without_a_site_is_refused": ("tools/files.py", [
        ('if location == "site" and not site and not cursor and kind != "site":', "if False:")]),
    "test_files.py::test_share_without_confirmation_writes_nothing": ("tools/files.py", [
        ('receipt = c.guard(kind="share", user_confirmed=user_confirmed, recipients=0)',
         'receipt = c.guard(kind="share", user_confirmed=True, recipients=0)'),
        ('receipt = c.guard(kind="share", user_confirmed=user_confirmed, recipients=len(who),',
         'receipt = c.guard(kind="share", user_confirmed=True, recipients=len(who),')]),
    "test_files.py::test_invite_more_than_20_people_is_refused": ("tools/files.py", [
        ("if people and len(people) > MAX_INVITEES:", "if False:"),
        ("recipients=len(who),\n", "recipients=0,\n")]),
    "test_files.py::test_revoke_of_the_owner_is_refused": ("tools/files.py", [
        ('raise invalid("The owner\'s own access cannot be revoked. Nothing was changed.")', "pass")]),
    "test_recipient_domains.py::test_sharing_with_someone_outside_the_allowed_domains_is_refused": ("tools/files.py", [
        ("group_addresses=groups_among(who), addresses=[w.address for w in who])", "group_addresses=groups_among(who))")]),
    # -- excel --------------------------------------------------------------------------------
    "test_excel.py::test_non_workbooks_and_xls_are_refused": ("tools/excel.py", [
        ('if "folder" in m or ext not in ("xlsx", "xlsm"):', "if False:")]),
    "test_excel.py::test_shape_not_matching_the_address_is_refused_with_the_expected_shape": ("tools/excel.py", [
        ("if (er, ec) != (rows, cols):", "if False:")]),
    "test_excel.py::test_more_than_5000_cells_is_refused": ("tools/excel.py", [
        ("if rows * cols > MAX_WRITE_CELLS:", "if False:")]),
    "test_excel.py::test_get_range_over_max_cells_is_refused_with_its_size": ("tools/excel.py", [
        ("if cells > max_cells:", "if False:")]),
    # -- teams --------------------------------------------------------------------------------
    "test_teams.py::test_list_conversations_refuses_options_of_the_other_kind": ("tools/teams.py", [
        ('if unread_only and kind == "channel":', "if False:"), ('if team and kind == "chat":', "if False:")]),
    "test_teams.py::test_get_conversation_needs_exactly_one_conversation": ("tools/teams.py", [
        ("if sum((bool(chat), bool(person), in_channel)) != 1:", "if False:")]),
    "test_teams.py::test_get_conversation_thread_and_channel_need_their_team": ("tools/teams.py", [
        ("if in_channel and not (team and channel):", "if False:")]),
    "test_teams.py::test_get_conversation_members_are_for_chats_only": ("tools/teams.py", [
        ("if in_channel and include_members:", "if False:")]),
    "test_teams.py::test_send_without_confirmation_writes_nothing": ("tools/teams.py", [
        ('        if user_confirmed is not True:\n            # Only reads so far.', '        if False:\n            # Only reads so far.'),
        ('c.guard(kind="teams_message", user_confirmed=user_confirmed, recipients=max(',
         'c.guard(kind="teams_message", user_confirmed=True, recipients=max(')]),
    "test_teams.py::test_a_channel_the_person_cannot_post_to_is_refused": ("graph/errors.py", [
        ('        return M365Error("forbidden")\n', '        return M365Error("uncertain", request_id="none")\n')]),
    "test_teams.py::test_an_empty_bucket_stops_the_message": BUCKET,
    "test_refusals.py::test_a_chat_with_someone_teams_cannot_reach_is_not_a_sign_in_problem": ("graph/errors.py", [
        ('if e.status == 401 and inner == "TenantNotFound":', "if False:")]),
    "test_recipient_domains.py::test_a_teams_message_to_someone_outside_the_allowed_domains_is_refused": ("tools/teams.py", [
        ("group_addresses=groups_among(who_list), addresses=audience)", "group_addresses=groups_among(who_list))")]),
    "test_recipient_domains.py::test_a_message_to_an_existing_chat_with_an_outside_member_is_refused": ("tools/teams.py", [
        ('if kind in ("chat", "channel") and c.rt.settings.allowed_recipient_domains:', "if False:")]),
    "test_teams.py::test_someone_elses_message_is_never_changed": ("tools/teams.py", [
        ('if ((m.get("from") or {}).get("user") or {}).get("id") != me["id"]:', "if False:")]),
    "test_teams.py::test_changing_a_message_without_confirmation_changes_nothing": ("tools/teams.py", [
        ('receipt = c.guard(kind="teams_message", user_confirmed=user_confirmed, recipients=1, addresses=audience)',
         'receipt = c.guard(kind="teams_message", user_confirmed=True, recipients=1, addresses=audience)')]),
    "test_teams.py::test_changing_a_message_draws_from_the_send_allowance": BUCKET,
    "test_teams.py::test_edit_needs_a_body_and_one_place": ("tools/teams.py", [
        ('if action == "edit" and not (body and body.strip()):', "if False:")]),
    "test_teams.py::test_creating_a_group_chat_needs_confirmation_and_two_people": ("tools/teams.py", [
        ('c.guard(kind="chat_members", user_confirmed=user_confirmed,', 'c.guard(kind="chat_members", user_confirmed=True,')]),
    "test_recipient_domains.py::test_an_automatic_reply_to_people_outside_is_refused_when_domains_are_limited": ("tools/autoreply.py", [
        ('            if audience != "none" and allowed:\n', "            if False:\n")]),
    "test_recipient_domains.py::test_a_cancellation_comment_to_an_attendee_outside_the_allowed_domains_is_refused": ("tools/calendar.py", [
        ("addresses=[_address(a) for a in _humans(e)] if comment else None)", "addresses=None)")]),
    "test_recipient_domains.py::test_a_channel_post_with_a_member_outside_the_allowed_domains_is_refused": ("tools/teams.py", [
        ('if kind in ("chat", "channel") and c.rt.settings.allowed_recipient_domains:',
         'if kind == "chat" and c.rt.settings.allowed_recipient_domains:')]),
    "test_recipient_domains.py::test_a_channel_post_whose_members_cannot_be_read_is_refused_when_domains_are_limited": ("tools/teams.py", [
        ('        c.need("ChannelMember.Read.All")\n', "        pass\n")]),
    "test_recipient_domains.py::test_changing_a_chat_message_with_an_outside_member_is_refused": UPDATE_MEMBERS,
    "test_recipient_domains.py::test_changing_a_channel_post_with_an_outside_member_is_refused": UPDATE_MEMBERS,
    "test_recipient_domains.py::test_adding_a_chat_member_outside_the_allowed_domains_is_refused": ("tools/teams.py", [
        ("addresses=[w.address for w in adding])", "addresses=None)")]),
    # -- the shared resolver ------------------------------------------------------------------
    "test_recipients.py::test_save_draft_refuses_an_internal_address_that_is_no_person": NOT_A_PERSON,
    "test_recipients.py::test_send_draft_refuses_a_draft_to_a_group_the_people_index_does_not_know": NOT_A_PERSON,
    "test_recipients.py::test_an_invitation_to_a_group_is_refused": NOT_A_PERSON,
    "test_recipients.py::test_sharing_a_file_with_a_group_is_refused": NOT_A_PERSON,
    "test_recipients.py::test_a_teams_message_to_a_group_is_refused": NOT_A_PERSON,
    "test_recipients.py::test_a_personal_contact_holding_a_group_address_is_refused": ("core/resolve.py", [
        ('return await classify(c, w.address, w) if w.source == "contacts" else w', "return w")]),
    "test_recipients.py::test_a_shared_first_name_is_refused_in_a_draft": SHARED_NAME,
    "test_recipients.py::test_a_shared_first_name_is_refused_in_a_teams_message": SHARED_NAME,
    "test_recipients.py::test_a_shared_first_name_is_refused_when_sharing": SHARED_NAME,
    "test_recipients.py::test_a_shared_first_name_is_refused_as_an_attendee": SHARED_NAME,
    "test_recipients.py::test_a_shared_first_name_is_refused_as_a_sender_filter": SHARED_NAME,
    "test_recipients.py::test_a_shared_first_name_is_refused_when_finding_meeting_times": SHARED_NAME,
    "test_recipients.py::test_a_shared_first_name_is_refused_when_opening_a_one_to_one_conversation": SHARED_NAME,
    "test_recipients.py::test_a_name_is_not_resolved_when_contacts_could_not_all_be_checked": ("core/resolve.py", [
        ("    if not own.complete:\n        # A second", "    if False:\n        # A second")]),
    "test_link_routes_mail.py::test_a_body_cannot_be_uploaded_to_a_draft_already_sent": ("links/routes.py", [
        ("    if _mark_value(d) == SENT_MARK:\n", "    if False:\n")]),
    # Two checks hold this: not a draft, and the sent mark; the parametrised test exercises each alone.
    "test_link_routes_mail.py::test_an_attachment_cannot_be_uploaded_to_a_draft_already_sent": ("links/routes.py", [
        ("    await _unsent_draft(graph, draft_id)\n", "")]),
    # One check, called before the sent mark and after it; removing it removes both.
    "test_link_routes_mail.py::test_send_draft_refuses_while_an_upload_to_the_draft_is_running": UPLOAD_RUNNING,
    "test_link_routes_mail.py::test_an_upload_starting_while_the_draft_is_marked_sent_stops_the_send": UPLOAD_RUNNING,
    "test_link_routes_mail.py::test_sending_a_draft_revokes_its_upload_links": ("tools/mail.py", [
        ("            c.rt.signer.revoke_draft(c.key, id)\n", "            pass\n")]),
    "test_recipients.py::test_a_contact_holding_a_group_people_knows_is_refused_even_outside": ("core/resolve.py", [
        ("replace(w, is_group=True) if key in groups and not w.is_group else w", "w")]),
    "test_recipients.py::test_a_group_people_knows_under_another_name_still_marks_the_contact": ("core/resolve.py", [
        ("replace(w, is_group=True) if key in groups and not w.is_group else w", "w")]),
    "test_recipients.py::test_the_candidates_carry_the_job_title": ("core/resolve.py", [
        ('return f"{w} ({w.job_title})" if w.job_title else str(w)', "return str(w)")]),
    # -- tasks --------------------------------------------------------------------------------
    "test_tasks.py::test_premium_or_foreign_plans_are_not_found": ("tools/tasks.py", [
        ('        if e.key == "unresolved":\n', "        if False:\n")]),
    "test_tasks.py::test_a_plan_the_person_is_not_a_member_of_is_not_found": ("tools/tasks.py", [
        ('        return one(query, plans, kind="plan", name=lambda p: p["title"])',
         '        return {"id": query, "title": query}')]),
    "test_tasks.py::test_delete_task_needs_confirmation": ("tools/tasks.py", [
        ('require(user_confirmed, "Deleting a task")', "pass")]),
    # -- QA on the test tenant, wave 2 ---------------------------------------------------------
    "test_contract.py::test_an_argument_the_tool_does_not_have_is_refused_before_graph": ("server.py", [
        ("        if unknown:\n", "        if False:\n")]),
    "test_confirmation.py::test_a_confirmation_that_is_not_the_boolean_true_is_refused": CONFIRM_TYPE,
    "test_confirmation.py::test_send_draft_with_the_word_yes_sends_nothing": CONFIRM_TYPE,
    "test_mail.py::test_renaming_or_moving_onto_an_existing_name_names_that_folder": ("tools/mail.py", [
        ("        if clash:\n", "        if False:\n")]),
    "test_mail.py::test_a_folder_name_with_a_slash_is_refused": ("tools/mail.py", [
        ('    if "/" in name:\n', "    if False:\n")]),
    "test_mail.py::test_a_folder_is_not_moved_under_itself": ("tools/mail.py", [
        ('        if parent_id and (parent_id == f["id"] or (dest and dest["path"].startswith(own_path + "/"))):\n',
         "        if False:\n")]),
    "test_mail.py::test_the_loser_of_two_marks_at_once_is_told_it_was_already_sent": ("tools/mail.py", [
        ("            if e.status == 412 and _mark_value(await _reread(c, id) or {}) == SENT_MARK:\n", "            if False:\n")]),
    # Two checks hold it: the settling reads, and the read right after the refusal.
    "test_mail.py::test_the_loser_whose_send_exchange_refuses_is_told_it_was_already_sent": ("tools/mail.py", [
        ('                        if now is not None and now.get("isDraft") is False:\n                            raise invalid(ALREADY_SENT)',
         '                        if False:\n                            raise invalid(ALREADY_SENT)'),
        ('                if now is not None and now.get("isDraft") is False:\n                    raise invalid(ALREADY_SENT)',
         '                if False:\n                    raise invalid(ALREADY_SENT)')]),
    "test_teams.py::test_a_message_id_that_is_not_one_is_not_found": ("tools/teams.py", [
        ('    return e.status == 400 and any(i is not None and not i.strip().isdigit() for i in ids)',
         "    return False")]),
    "test_teams.py::test_a_one_to_one_chat_with_oneself_is_refused_with_the_reason": ("tools/teams.py", [
        ("    if w.address.casefold() in _mine(me):\n", "    if False:\n")]),
    "test_teams.py::test_an_unknown_message_in_a_chat_is_not_found_not_forbidden": ("tools/teams.py", [
        ('            if not (channel and action == "restore"):\n                raise M365Error("not_found") from None\n',
         '            if not (channel and action == "restore"):\n                raise\n')]),
    "test_teams.py::test_restoring_a_post_never_in_the_channel_is_not_found": ("tools/teams.py", [
        ("            if ge.status not in (403, 404) or ge.missing_scope:\n", "            if ge.status != 404 or ge.missing_scope:\n")]),
    "test_teams.py::test_a_group_chat_needs_two_different_other_people": ("tools/teams.py", [
        ("        if not chat and len(adding) < 2:\n", "        if False:\n"),
        ("                       if w.address.casefold() not in _mine(me)}.values())", "                       }.values()) + []")]),
    "test_teams.py::test_a_chat_name_too_long_or_blank_is_refused": ("tools/teams.py", [
        ("    if topic is not None and not topic.strip():\n", "    if False:\n"),
        ("    if topic is not None and len(topic.strip()) > MAX_TOPIC:\n", "    if False:\n")]),
    "test_teams.py::test_a_one_to_one_chat_is_not_renamed": ("tools/teams.py", [
        ('            if info.get("chatType") != "group" and topic:\n', "            if False:\n")]),
    "test_mail.py::test_a_message_cursor_beside_other_filters_is_refused": ("tools/mail.py", [
        ('            if any(asked) and state.get("args") is not None and asked != state["args"]:\n', "            if False:\n")]),
    "test_mail.py::test_max_chars_below_one_is_refused": ("tools/mail.py", [
        ("    if max_chars < 1:\n", "    if False:\n")]),
    "test_mail.py::test_a_send_refused_while_the_draft_never_settles_keeps_it_marked_sent": ("tools/mail.py", [
        ('                    if now is None or now.get("isDraft") is not False:\n', "                    if False:\n")]),
    "test_contract.py::test_a_limit_below_one_is_refused": ("core/project.py", [
        ("    if limit < 1:\n", "    if False:\n")]),
    "test_calendar.py::test_a_negative_reminder_is_refused": ("tools/calendar.py", [
        ("if reminder_minutes is not None and reminder_minutes < 0:", "if False:")]),
    "test_calendar.py::test_a_comment_without_sending_the_response_is_refused": ("tools/calendar.py", [
        ("    if comment and not send_response:\n", "    if False:\n")]),
    "test_files.py::test_shared_files_while_microsoft_search_is_down_is_said_plainly": ("tools/files.py", [
        ('                if location == "shared":\n                    raise M365Error("unsupported"',
         '                if False:\n                    raise M365Error("unsupported"')]),
    "test_files.py::test_a_path_with_a_hash_is_refused_with_the_way_to_reach_it": ("core/drive.py", [
        ('    if "#" in p:\n', "    if False:\n")]),
    "test_files.py::test_a_name_with_a_hash_is_refused_everywhere": ("tools/files.py", [
        ("""_BAD_NAME = re.compile(r'[\\\\/:*?"<>|#\\x00-\\x1f]')""", """_BAD_NAME = re.compile(r'[\\\\/:*?"<>|\\x00-\\x1f]')""")]),
    "test_files.py::test_a_folder_that_exists_is_refused_with_its_path": ("tools/files.py", [
        ("                if ge.status != 409:\n", "                if True:\n")]),
    "test_files.py::test_a_change_that_failed_for_every_item_is_an_error": ("tools/files.py", [
        ("        if failed and not done:\n", "        if False:\n")]),
    "test_excel.py::test_a_cell_longer_than_excel_holds_is_refused": ("tools/excel.py", [
        ("    if any(isinstance(v, str) and len(v) > MAX_CELL_TEXT", "    if False and any(isinstance(v, str) and len(v) > MAX_CELL_TEXT")]),
    "test_teams.py::test_a_cursor_from_another_search_or_from_the_plain_list_is_refused": LIST_CURSOR,
    "test_teams.py::test_a_cursor_is_refused_beside_arguments_it_was_not_made_with": LIST_CURSOR,
    "test_teams.py::test_messages_out_of_range_are_refused_below_and_clamped_above_with_a_note": ("tools/teams.py", [
        ("    if messages < 1:\n", "    if False:\n")]),
    # Two checks hold this row: the conversation the cursor names, and a cursor beside before.
    "test_teams.py::test_a_conversation_cursor_is_refused_elsewhere_and_beside_before": ("tools/teams.py", [
        ('    if state is not None and state.get("conv") != conv:\n', "    if False:\n"),
        ("    if cursor and before:\n", "    if False:\n")]),
    "test_teams.py::test_an_edit_longer_than_teams_accepts_is_refused": ("tools/teams.py", [
        ("            if len(prepared.html.encode()) > MAX_MESSAGE_HTML:\n", "            if False:\n")]),
    "test_teams.py::test_renaming_a_group_chat_needs_confirmation": ("tools/teams.py", [
        ("        if adding or removing or topic:\n", "        if adding or removing:\n")]),
    "test_teams.py::test_removing_oneself_from_a_chat_is_refused": ("tools/teams.py", [
        ("                if w.address.casefold() in mine:\n", "                if False:\n")]),
    "test_teams.py::test_a_chat_whose_members_cannot_all_be_read_is_not_changed": ("tools/teams.py", [
        ("                if next_page:\n                    raise invalid(\"This chat has more members", "                if False:\n                    raise invalid(\"This chat has more members")]),
    "test_teams.py::test_someone_elses_deleted_channel_post_is_not_restored": ("tools/teams.py", [
        ('if ((m.get("from") or {}).get("user") or {}).get("id") != me["id"]:', "if False:")]),
    "test_files.py::test_a_site_name_while_search_is_down_is_not_called_missing": ("core/drive.py", [
        ("        if not answered and not known:\n", "        if False:\n")]),
    "test_teams.py::test_looking_for_a_deleted_post_reads_four_pages_and_no_more": ("tools/teams.py", [
        ('    raise invalid(f"That post is not among the latest', '    raise M365Error("not_found")\n    raise invalid(f"That post is not among the latest')]),
    "test_calendar.py::test_working_hours_in_a_zone_the_server_cannot_place_are_not_shifted_to_utc": ("tools/calendar.py", [
        ("        if not known(name):\n            # zone() would", "        if False:\n            # zone() would")]),
    "test_tasks.py::test_a_task_title_over_255_characters_is_refused": ("tools/tasks.py", [
        ("    if units > MAX_TITLE:\n", "    if False:\n")]),
    # -- found by QA, round 2 (calendar, tasks, notes, files, lists, excel) ----------------------
    "test_tasks.py::test_a_title_is_counted_as_microsoft_counts_it": ("tools/tasks.py", [
        ('    units = len(title.encode("utf-16-le")) // 2 if title else 0\n', "    units = len(title) if title else 0\n")]),
    "test_refusals.py::test_an_id_holding_a_separator_is_not_found_before_graph": ("core/ids.py", [
        ('    if _NOT_ONE_SEGMENT.search(v) or v in (".", ".."):\n', "    if False:\n")]),
    "test_refusals.py::test_an_empty_id_on_a_save_is_refused_not_read_as_new": ("core/ids.py", [
        ("    if not v:\n", "    if False:\n")]),
    "test_refusals.py::test_a_todo_delete_of_an_id_to_do_cannot_read_is_not_found": ("tools/tasks.py", [
        ('            await c.graph.get(f"/me/todo/lists/{lst[\'id\']}/tasks/{id}", params={"$select": "id"})\n', "")]),
    "test_notes.py::test_get_note_page_refuses_max_chars_outside_its_range": ("tools/notes.py", [
        ("    if not 1 <= max_chars <= MAX_PAGE_CHARS:\n", "    if False:\n")]),
    "test_notes.py::test_a_page_onenote_lists_but_does_not_open_is_said_plainly": ("tools/notes.py", [
        ("    if e.status != 404:\n        return e\n", "    return e\n")]),
    "test_files.py::test_a_folder_that_is_a_file_is_said_to_be_a_file": ("tools/files.py", [
        ('if ge.status == 422 and ge.code == "notSupported" and folder:', "if False:")]),
    "test_files.py::test_a_name_onedrive_reserves_is_refused_before_a_link_is_made": ("tools/files.py", [
        ('    if _RESERVED_NAME.match(clean) or clean.endswith(".") or "_vti_" in clean.casefold():\n', "    if False:\n")]),
    "test_files.py::test_a_copy_that_failed_at_the_monitor_is_an_error_with_its_reason": ("tools/files.py", [
        ('            for i in [i for i in done if states.get(i) == "failed"]:\n', "            for i in []:\n")]),
    "test_files.py::test_a_folder_into_its_own_subfolder_is_refused_with_why": ("tools/files.py", [
        ("reason = await _into_itself(c, drive, i, dest) if e.status == 400 and dest else None", "reason = None")]),
    "test_lists.py::test_a_text_longer_than_its_column_is_refused_with_the_limit": ("tools/lists.py", [
        ("    if isinstance(limit, int) and len(raw) > limit:\n", "    if False:\n")]),
    "test_excel.py::test_an_address_outside_a_sheet_is_refused_with_the_bounds": ("tools/excel.py", [
        ("        if col is not None and not (1 <= _col_number(col) <= MAX_COLUMN and 1 <= int(row) <= MAX_ROW):\n",
         "        if False:\n")]),
    "test_excel.py::test_a_missing_sheet_or_table_is_named_with_the_real_ones": ("tools/excel.py", [
        ('    if (want or "").strip().casefold() in {str(n).casefold() for n in names}:\n', "    if True:\n")]),
    "test_excel.py::test_an_address_and_a_table_together_are_refused": ("tools/excel.py", [
        ('    if action == "write" and table:\n', "    if False:\n"),
        ('    if action == "append_rows" and address:\n', "    if False:\n")]),
    "test_excel.py::test_a_formula_excel_refuses_is_named": ("tools/excel.py", [
        ('if e.status == 400 and e.code == "InvalidArgument" and any(', "if False and any(")]),
}


def main(argv: list[str]) -> int:
    only = argv[0] if argv else ""
    parsed = parse(documents(), root=ROOT)
    # A row that lost its test contributes no id, and every remaining mutation would still pass:
    # the map is checked first, or a broken map reads as complete evidence.
    broken = check(parsed, collected())
    if broken:
        print("the refusal map is broken, so the evidence would be incomplete:")
        for p in broken:
            print(" -", p)
        return 1
    ids = [t for t in mapped_tests(parsed) if only in t]
    problems: list[str] = []
    missing = [t for t in ids if t not in MUTATIONS]
    problems += [f"{t}: no mutation for its guard" for t in missing]
    if not ids:
        problems.append("no mapped tests selected: nothing would be shown failing")
    started = time.monotonic()
    # Every test must pass on the untouched code, or failing under a mutation proves nothing.
    if ids and run(*(f"{TEST_ROOT}/{t}" for t in ids)) != 0:
        problems.append("the mapped tests do not all pass on the unmutated code")
    caught = 0
    for t in ids:
        if t not in MUTATIONS:
            continue
        rel, edits = MUTATIONS[t]
        try:
            with mutated(rel, edits):
                code = run(f"{TEST_ROOT}/{t}")
        except GuardNotFound as e:
            problems.append(f"{t}: {e}")
            print(f"{'GONE':10} {t}")
            continue
        status = {0: "NOT CAUGHT", 1: "caught"}.get(code, f"BROKEN({code})")
        print(f"{status:10} {t}")
        if code == 1:
            caught += 1
        elif code == 0:
            problems.append(f"{t}: still passes with its guard removed")
        else:
            problems.append(f"{t}: pytest exited {code}, which is not a failing test")
    if ids and run(*(f"{TEST_ROOT}/{t}" for t in ids)) != 0:
        problems.append("the mapped tests do not pass after the code was restored")
    print(f"\n{caught}/{len(ids)} row tests seen failing with their guard removed, "
          f"{len(problems)} problems, {time.monotonic() - started:.0f} s")
    for p in problems:
        print(" -", p)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
