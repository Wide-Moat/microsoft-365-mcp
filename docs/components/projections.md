# Component: projections

One function per Graph type, from a Graph object to the dict a tool returns. They are the only path
from Graph to a result.

Shape rules for every projection:

- Keys are short, snake_case, stable: `id`, `subject`, `from`, `to`, `received`, `size`.
- People are `"Name <address>"` strings, not nested objects.
- Times are local strings with offset; durations in minutes.
- Booleans only when true (`unread: true`), absent otherwise; empty lists absent.
- No `@odata.*`, `changeKey`, `etag`, `parentFolderId` or internal ids the model cannot use.
- Every free-text field from someone else (body, preview, chat text, page content, a draft's
  preview with its quoted letter, a SharePoint list item's fields) is reduced by the HTML component
  and wrapped in a content block; short ones (mail and event subjects, chat topics) in a content
  line.

Each projection has a declared field list; a test feeds a full recorded Graph object and asserts
the output keys are exactly the declared ones, and that the output is smaller than a fixed share of
the input (measured on recordings: an envelope must be under 15% of the raw message JSON).

Projections exist for: message envelope, message, attachment, mail folder, event, attendee,
schedule item, meeting suggestion, contact, person, drive item, site, list, list item, workbook
range, table, chat, chat message, team, channel, recording, transcript line, To Do list and task,
Planner plan, bucket and task, notebook, section, page, automatic reply setting, Bookings business,
service, staff, appointment.
