# Development notes

## Endpoint mapping

| Tool | Microsoft Graph v1.0 request | Returned data boundary |
|---|---|---|
| ms_auth_test | GET /me | allowlisted profile fields |
| ms_get_profile | GET /me with $select | allowlisted profile fields |
| ms_search_messages | POST /search/query with entityTypes=[message] | hit metadata and 1000-character body preview |
| ms_list_calendar_events | GET /me/calendar/events or /me/calendars/{id}/events | event identity, time, location and response status |
| ms_list_calendar_view | GET /me/calendarView with startDateTime/endDateTime | event identity, time, location and response status |
| ms_search_drive_items | GET /me/drive/root/search(q=...) | file/folder metadata; no content or download URL |
| ms_get_drive_item | GET /me/drive/items/{id} | file/folder metadata; no content or download URL |
| ms_search_chat_messages | POST /search/query with entityTypes=[chatMessage] | sender, timestamps and 1000-character text preview |

The client always selects a fixed field allowlist. It does not accept a caller-supplied Graph path or field selector.

## Cursor handling

Graph collection responses may include an absolute @odata.nextLink containing a skip token. The client verifies the Graph origin and the expected resource path, stores the relative path in an in-memory map, and returns only a random m365 cursor. Cursor input is accepted only if it is present in that map, is no longer than 128 characters, and matches the tool's fixed collection path. The raw Graph nextLink and skip token are never rendered or returned.

## Error handling

Graph error bodies are parsed only for a short alphanumeric error code and request ID. The model-facing error is a generic HTTP status message; the original Graph error message is not copied. Request IDs are limited to 64 sanitized characters. HTTP 429 and 5xx errors are rethrown to allow the host's retry policy to run.

## Release checks

Run the following from this directory:

~~~sh
npm install --ignore-scripts
npm run typecheck
npm run build
npm pack --dry-run
~~~

Do not commit access tokens, generated lib output, or tarballs. The package uses Node.js >=22.19 and publishes lib, documentation, examples, and LICENSE.
