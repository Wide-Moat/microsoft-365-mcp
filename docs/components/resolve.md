# Component: resolve

Turns the words a person uses into the ids Graph needs. The model never has to know a Graph id it
did not receive from a previous result.

| Kind | Input | Order of lookup | Ambiguity |
|---|---|---|---|
| Person | name, part of a name, address | exact address → personal contacts → Microsoft Search `person` (ranked by the person's own interactions) → directory `users` with `$search` on displayName and mail | error with up to 5 candidates: name, address, job title, department |
| Group or list address | address | not a user in the directory, at one of the organisation's domains | refused where people are expected (distribution lists) |
| Mail folder | well-known name, display name, path `A/B` | well-known names → path walk → name search across the tree | error listing paths |
| Calendar | "my", a colleague's name (shared calendar), a calendar name | own calendars → calendars shared with me | error listing names |
| Site | name or URL | `sites?search=` → followed sites | error listing site names and URLs |
| Drive path | `My files/Reports/2026` or a site-relative path | path addressing `root:/path:` | not ambiguous; missing → not found |
| Team / channel | `Team name/Channel name` | joined teams → channels | error listing `team/channel` pairs |
| Planner plan, To Do list, OneNote notebook/section | name | the person's own containers | error listing names |
| Time | `YYYY-MM-DD HH:MM`, with optional offset; `today`, `tomorrow` are not parsed — the model converts words to dates | mailbox time zone from `mailboxSettings` | — |

Rules:

- For folders, calendars, sites, plans and lists, a single exact match wins over fuzzy matches.
  Two exact matches is still ambiguity.
- For people it does not: a name that several people match is refused with all of them, even when
  one display name equals it exactly. "Megan" is the whole display name of one colleague and the
  first name of another; picking the first sent a Teams message and a file invitation to the wrong
  person on a real tenant. The candidates carry the address, so the next call names one.
- Personal contacts are read from `/me/contacts` and from every contact folder (`/me/contactFolders`
  and the folders inside them), and matched on every word of the name. A contact
  can hold any address, so its address is classified like one typed in.
- An address at one of the organisation's verified domains (`/organization`, `verifiedDomains`) must
  be a user in its directory (`/users?$filter=mail eq … or userPrincipalName eq …`). Anything else
  there — a Microsoft 365 group, a distribution list, a mail-enabled security group, an alias, or no
  such address — is refused where people are expected. Groups are not read directly: that needs
  `Group.Read.All`, which no tool needs otherwise (rule 18); the People index does not know most
  group addresses (a group and a list were sent to on a real tenant). Refusing an unknown internal
  address is deliberate: it errs on the side of what other people see (rule 22), and an alias is
  answered by using the person's name or main address. An address outside the organisation is taken
  as given; whether it is a list there cannot be known, and `M365_ALLOWED_RECIPIENT_DOMAINS` limits it.
- Resolution results are cached per person for 10 minutes (names change rarely; staleness only
  costs one more lookup).
- Candidates go to the model in the error result, never to logs.

## Must refuse

Every tool that reaches people goes through `person()`, so each refusal is tested through the tools.

| Must refuse | The person is told | Test |
|---|---|---|
| An address at the organisation's domains that is no user in its directory, in a draft, at send time, in an invitation, a file invitation or a Teams message | the address, and that it is a group, list, alias or nothing; name the people | `test_recipients.py::test_save_draft_refuses_an_internal_address_that_is_no_person`, `test_recipients.py::test_send_draft_refuses_a_draft_to_a_group_the_people_index_does_not_know`, `test_recipients.py::test_an_invitation_to_a_group_is_refused`, `test_recipients.py::test_sharing_a_file_with_a_group_is_refused`, `test_recipients.py::test_a_teams_message_to_a_group_is_refused` |
| A personal contact whose address is such a group, or one People knows as a group (also outside the organisation) | the same | `test_recipients.py::test_a_personal_contact_holding_a_group_address_is_refused`, `test_recipients.py::test_a_contact_holding_a_group_people_knows_is_refused_even_outside`, `test_recipients.py::test_a_group_people_knows_under_another_name_still_marks_the_contact` |
| A name, when the person has more personal contacts and contact folders than eight requests read (contacts and folder lists together): a second match could hide past them | to give the address | `test_recipients.py::test_a_name_is_not_resolved_when_contacts_could_not_all_be_checked` |
| A name that several people match, even when one display name equals it | up to 5 candidates with name, address and job title | `test_recipients.py::test_a_shared_first_name_is_refused_in_a_draft`, `test_recipients.py::test_a_shared_first_name_is_refused_in_a_teams_message`, `test_recipients.py::test_a_shared_first_name_is_refused_when_sharing`, `test_recipients.py::test_a_shared_first_name_is_refused_as_an_attendee`, `test_recipients.py::test_a_shared_first_name_is_refused_as_a_sender_filter`, `test_recipients.py::test_a_shared_first_name_is_refused_when_finding_meeting_times`, `test_recipients.py::test_a_shared_first_name_is_refused_when_opening_a_one_to_one_conversation`, `test_recipients.py::test_the_candidates_carry_the_job_title` |

## Tests

- Two "Anna" in the directory → error with both; "anna.k@" exact → resolved.
- A distribution list given as a recipient → refused by the guard via `group` detection.
- `Inbox/Projects` resolves through the tree; two "Projects" folders under different parents →
  error listing both paths.
- `2026-10-23 15:00` with mailbox zone Europe/Lisbon → `2026-10-23T15:00:00+01:00`; an explicit
  `+03:00` is respected.
- A cached resolution is not shared between two people.
