# Component: send guard

Runs before every action that other people will see. One place, so no tool can forget it.

## What counts as visible to others

Sending a draft; a Teams message; creating, changing or cancelling an event with attendees;
responding to an invitation (the organiser is notified); sharing a file or a link; inviting people
to a file; turning on or changing an automatic reply; creating or cancelling a Bookings
appointment; adding or removing chat members.

## Checks, in order

1. `user_confirmed is True`, except where the manifesto exempts the action (responding to one's own
   invitation). Otherwise: "Ask the person to confirm", naming who each recipient resolved to where the tool knows, and what did not happen ("Nothing was sent." for messages; saved, changed, shared or booked for the rest).
2. Recipients resolved and counted: at most 20 people in total; no distribution lists or
   Microsoft 365 group addresses. An address at the organisation's own domains must be a user in
   its directory; anything else there (a group, a list, a mail-enabled security group, an alias, an
   address that does not exist) is refused ([resolve](resolve.md)). Otherwise the count or the
   offending address.
3. When the installation sets `M365_ALLOWED_RECIPIENT_DOMAINS`, every address the action reaches is
   at one of those domains or a subdomain of one (`eu.example.com` under `example.com`, never
   `badexample.com`), compared without case. Otherwise: the offending **domains** — never the
   addresses — and "Nothing was sent." Unset, no domain is refused.
4. The bucket: per person, capacity 10, refill 10 per minute. Each action takes one token. Empty →
   "Too many messages in a short time; wait N seconds. Bulk sending is not something the assistant
   does." The checks above run first, so a refused action takes no token.
5. Only then the Graph call.

### Which addresses the domain check sees

| Action | Addresses checked |
|---|---|
| `send_draft` | every `to`, `cc` and `bcc` recipient of the draft, re-read at send time |
| `save_event` with attendees | everyone notified: the attendees after the change and, for a change, those removed (they get a cancellation) |
| `share_file` `invite` | the people invited |
| `send_message` to people | the people the chat is found or created with |
| `send_message` to a chat id | the chat's members, read from Graph (one extra request, only while domains are set); a member without an address is refused |
| `send_message` to a channel | the channel's members, guests and shared-channel participants included, read from Graph while domains are set; this needs `ChannelMember.Read.All`, and without it a channel post is refused |
| `save_chat` | the members being added (removing someone is always allowed) |
| `cancel_event` with a comment | every attendee; without a comment the cancellation is not refused, or the meeting would stay standing with them in it |
| `respond_to_event` with a comment | the organiser; an answer without one is Exchange's own notice |
| `save_automatic_reply` | nobody can be named in advance, so while domains are set an outside audience (`contacts`, `all`) is refused and only `audience=none` is accepted |

Not checked, on purpose: Bookings (customers are outside the organisation by definition), and a
Planner assignment (assignees come from the directory). A `share_file` `link` is organisation-scoped and
reaches nobody outside.

The guard returns a receipt (`kind`, recipients count, bucket left) that the tool includes in its
result, so the model can tell the person what happened.

## Must refuse

The guard's own refusals, below every tool that calls it; each tool's rows are in
[requirements](../requirements.md).

| Must refuse | The person is told | Test |
|---|---|---|
| A visible action without `user_confirmed` | ask the person to confirm; nothing was sent | `test_core.py::test_guard_requires_confirmation` |
| More than 20 recipients, or a group among them | the count, or the group's address | `test_core.py::test_guard_caps_recipients_and_refuses_lists` |
| The person's bucket empty | the seconds to wait; bulk sending is not done here | `test_core.py::test_bucket_bursts_ten_then_refills_one_per_six_seconds` |
| An address outside `M365_ALLOWED_RECIPIENT_DOMAINS`, when it is set | the domains outside it, never the addresses | `test_recipient_domains.py::test_guard_refuses_an_address_outside_the_allowed_domains` |
| A domain that only ends in an allowed one (`evilexample.com` for `example.com`) | that domain | `test_recipient_domains.py::test_a_lookalike_domain_is_not_a_subdomain` |
| An address with no domain, while domains are set | that it has none | `test_recipient_domains.py::test_an_address_without_a_domain_is_refused` |

## State

The bucket lives in process memory, keyed by person hash. Several replicas each hold their own
bucket; that multiplies the burst, not the steady rate beyond what the gateway's own per-key limits
allow. A shared store can replace it behind the same interface if an installation needs a strict
global limit.

## Tests

- Each visible action without confirmation → error, and no Graph request made (recording
  transport stays empty).
- 21 recipients → error; a distribution list → error naming it.
- 10 sends pass, the 11th fails with the wait; after 6 s of fake time one more passes.
- Responding to an invitation without `user_confirmed` passes.
- The receipt is present in every successful visible action's result.
