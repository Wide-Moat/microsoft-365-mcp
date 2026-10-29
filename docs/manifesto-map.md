# Manifesto map

Which test holds each rule of the [manifesto](../MANIFESTO.md) (rule 36: every rule has a test that
fails when it is broken). `scripts/manifesto_map.py` reads this table in CI and fails on a rule with
no holder, a test id pytest does not collect, or a `ci:` step that `.github/workflows/validate.yml`
does not have. Test ids are relative to `tests/unit/`. A `ci:` holder is checked by name only: that
the step still does what its name says is for review to see.

| Rule | Held by |
|---|---|
| 1 | `test_surface.py::test_the_registry_matches_the_committed_lock` `ci: the catalogue fits the token budget` |
| 2 | `test_contract.py::test_name_uses_an_allowed_verb_and_no_prefix` |
| 3 | `test_manifesto.py::test_actions_are_closed_enums_inline_in_the_schema` |
| 4 | `test_contract.py::test_name_uses_an_allowed_verb_and_no_prefix` |
| 5 | `test_contract.py::test_description_reads_like_a_man_page` `test_contract.py::test_irreversible_tools_say_so` |
| 6 | `test_contract.py::test_arguments_are_flat_and_described` |
| 7 | `test_mail.py::test_ambiguous_recipient_lists_candidates` `test_mail.py::test_ambiguous_folder_lists_paths` `test_files.py::test_two_sites_with_one_name_are_ambiguous` `test_core.py::test_local_time_uses_the_mailbox_zone` |
| 8 | `test_mail.py::test_list_messages_returns_envelopes_without_bodies` |
| 9 | `test_contract.py::test_lists_take_a_capped_limit` `test_mail.py::test_get_message_truncates_and_says_so` `test_account.py::test_list_people_continues_with_a_cursor` |
| 10 | `test_mail.py::test_get_message_attachments_are_links_not_bytes` |
| 11 | `test_refusals.py::test_the_shared_refusals_name_a_next_step` `test_account.py::test_graph_401_means_reconnect` |
| 12 | `test_contract.py::test_annotations_follow_behaviour` `test_contract.py::test_read_only_tools_do_not_write` |
| 13 | `test_contract.py::test_a_switched_off_area_publishes_nothing` `test_contract.py::test_a_tool_of_an_area_switched_off_is_refused` `test_contract.py::test_per_person_area_header_hides_and_refuses` |
| 14 | `test_contract.py::test_the_catalogue_fits_the_budget` `ci: the catalogue fits the token budget` |
| 15 | `test_manifesto.py::test_no_app_only_credentials_or_token_store` `test_contract.py::test_no_mailbox_or_user_selector_anywhere` |
| 16 | `test_contract.py::test_no_mailbox_or_user_selector_anywhere` |
| 17 | `test_graph_client.py::test_missing_scope_named_in_a_403_is_reported` |
| 18 | `test_manifesto.py::test_every_area_scope_is_needed_by_a_tool_of_that_area` |
| 19 | `test_mail.py::test_delete_moves_to_deleted_items_never_deletes` `test_files.py::test_delete_a_file_goes_to_the_recycle_bin_and_says_where_to_restore` `test_files.py::test_delete_a_folder_without_confirmation_is_refused_and_nothing_is_deleted` |
| 20 | `test_confirmation.py::test_send_draft_with_the_word_yes_sends_nothing` `test_mail.py::test_send_refuses_a_draft_the_person_wrote` |
| 21 | `test_contract.py::test_tools_that_notify_others_go_through_the_guard` `test_core.py::test_guard_requires_confirmation` `test_autoreply.py::test_turning_it_on_without_confirmation_writes_nothing` `test_core.py::test_guard_allows_own_response_without_confirmation` |
| 22 | `test_core.py::test_bucket_bursts_ten_then_refills_one_per_six_seconds` `test_core.py::test_guard_caps_recipients_and_refuses_lists` `test_mail.py::test_distribution_list_is_refused` |
| 23 | `test_calendar.py::test_only_the_organiser_cancels` `test_calendar.py::test_changing_an_event_someone_else_organises_is_refused` `test_teams.py::test_someone_elses_message_is_never_changed` |
| 24 | `test_contract.py::test_forbidden_capabilities_are_absent` |
| 25 | `test_html.py::test_content_block_marks_foreign_text` `test_mail.py::test_get_message_body_cannot_close_its_own_content_block` |
| 26 | `test_contract.py::test_forbidden_capabilities_are_absent` `test_confirmation.py::test_send_draft_with_the_word_yes_sends_nothing` |
| 27 | `test_html.py::test_hidden_text_never_reaches_the_model` `test_html.py::test_tracking_parameters_are_removed_and_the_rest_kept` `test_html.py::test_a_safe_link_is_unwrapped_to_its_target` |
| 28 | `test_html.py::test_oversized_body_is_refused_before_parsing` `test_html.py::test_non_utf8_is_refused` |
| 29 | `test_html.py::test_active_content_is_removed_and_reported` `test_html.py::test_data_image_becomes_an_inline_attachment` `test_html.py::test_style_block_is_inlined` |
| 30 | `test_html.py::test_relative_urls_are_errors` `test_html.py::test_flex_layout_is_a_warning_not_an_error` `test_html.py::test_web_fonts_are_a_warning` |
| 31 | `test_mail.py::test_reply_keeps_quoted_history_byte_for_byte` `test_mail.py::test_changing_a_reply_draft_keeps_the_quoted_letter` |
| 32 | `test_link_tokens.py::test_tampered_or_garbage_links_are_refused` `test_link_tokens.py::test_expired_is_refused` `test_link_tokens.py::test_wrong_operation_is_refused` `test_link_tokens.py::test_no_or_weak_key_refuses_to_sign` `test_link_tokens.py::test_ttl_is_capped` |
| 33 | `test_link_routes_mail.py::test_small_attachment_is_one_post` `test_link_routes_mail.py::test_large_attachment_uses_an_upload_session_in_ranges` `test_link_routes_mail.py::test_an_attachment_past_the_message_size_limit_is_refused` |
| 34 | `test_graph_client.py::test_headers_on_every_request` `test_account.py::test_every_request_carries_immutable_ids_and_a_request_id` |
| 35 | `test_manifesto.py::test_the_repository_is_written_in_english` `test_contract.py::test_every_error_key_used_exists` |
| 36 | `test_refusal_map.py::test_every_must_refuse_row_names_a_collected_test` `ci: every safety rule is caught by a test when it is broken` `ci: every Must refuse row's test fails with its guard removed` `ci: every manifesto rule names a collected test` `test_manifesto_map.py::test_every_rule_names_a_collected_holder` |
| 37 | `ci: the suite is not empty and every collected test ran` |
| 38 | `test_core.py::test_a_full_mail_flow_logs_no_content` `test_core.py::test_log_filter_drops_content_and_counts_it` `test_graph_client.py::test_error_body_text_never_reaches_the_exception` |
| 39 | `test_contract.py::test_only_the_graph_client_speaks_http` `test_graph_client.py::test_429_waits_retry_after_then_succeeds` `test_graph_client.py::test_outlook_concurrency_is_four_per_person` |
| 40 | `test_manifesto.py::test_every_live_finding_is_measured_by_a_probe` |
| 41 | `ci: purity` |
| 42 | `test_manifesto.py::test_the_review_checklist_is_wired_to_the_manifesto` |

Rule 1 has no mechanical test: whether two tools answer one question is a judgement. What a test
can hold is that no tool appears or splits unseen; the lock and the budget table show every such
change on the pull request, and review item B17 decides.
