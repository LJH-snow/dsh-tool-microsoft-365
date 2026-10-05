# dsh-tool-microsoft-365

[English](README.md) | [中文](README.zh.md)

A read-only Microsoft 365 plugin for [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness). It uses the Microsoft Graph v1.0 REST API to let an agent inspect the signed-in user's profile, Outlook mail, calendar, OneDrive metadata, and Teams chat search results.

The plugin does not send mail, edit events, upload or download files, change Teams data, or call any Microsoft Graph mutation endpoint.

## Install

~~~sh
npm install @libai168/dsh-tool-microsoft-365
~~~

The package requires @deepseek-ai/cordis (^4.0.1) and @deepseek-ai/dsh-tools (^0.1.0-rc.6) as peer dependencies.

## Authentication and configuration

Provide a delegated Microsoft Graph OAuth access token through the host configuration. The token must be kept outside source control and is never returned by a tool.

~~~yaml
- name: '@libai168/dsh-tool-microsoft-365'
  config:
    accessToken: 'replace-with-a-short-lived-delegated-graph-token'
    # timeoutMs: 15000
~~~

See [examples/cordis.yml](examples/cordis.yml). The client accepts only the production https://graph.microsoft.com/v1.0 endpoint. This prevents a Bearer token from being sent to an arbitrary configured host. OAuth acquisition and refresh belong in the host application.

Recommended delegated permissions are the smallest set needed by the tools you enable:

| Capability | Microsoft Graph delegated permission |
|---|---|
| ms_auth_test, ms_get_profile | User.Read |
| ms_search_messages | Mail.Read |
| ms_list_calendar_events, ms_list_calendar_view | Calendars.Read |
| ms_search_drive_items, ms_get_drive_item | Files.Read |
| ms_search_chat_messages | Chat.Read |

Microsoft Entra consent and Conditional Access policies can require administrator approval. A token with fewer permissions returns the Graph status and a sanitized error code.

## Tools

| Tool | Description | Access |
|---|---|---|
| ms_auth_test | Verify the delegated token through /me and return safe profile fields | read |
| ms_get_profile | Read the signed-in user's allowlisted profile fields | read |
| ms_search_messages | Search Outlook messages through POST /search/query | read |
| ms_list_calendar_events | List default or selected calendar events | read |
| ms_list_calendar_view | List recurring event instances in an ISO 8601 time window | read |
| ms_search_drive_items | Search OneDrive files and folders | read |
| ms_get_drive_item | Read one OneDrive item's safe metadata | read |
| ms_search_chat_messages | Search Teams chat messages through POST /search/query | read |

Search requests use Microsoft Graph's read-only Search API with message or chatMessage entity types. Calendar and OneDrive collections use the Graph @odata.nextLink mechanism internally.

## Pagination

Each list tool returns a short, process-local nextCursor when Graph has another page. Pass that value back as cursor to the same tool. The cursor is random and contains no Graph URL or skip token; it expires when the plugin process exits. The plugin checks that every stored cursor belongs to the fixed collection path for that tool.

Offset search tools return nextFrom and accept it as from. Both offset and page-size values are bounded to avoid accidentally requesting an unbounded result set.

## Security and data boundaries

- Only https://graph.microsoft.com/v1.0 is accepted as baseUrl; arbitrary HTTP hosts and absolute cursor URLs are rejected before a request.
- Every request uses Authorization: Bearer; the token and raw request headers never appear in tool output.
- Graph error messages are not copied into model output. Errors contain HTTP status, a validated short Graph error code, and a truncated request ID when available.
- Profile names, UPN/mail addresses, organization fields, sender fields, subjects, locations, and URLs are length-limited. Outlook and Teams content is restricted to a 1000-character preview with HTML stripped for Teams.
- Drive results omit file content and preauthenticated download URLs. The plugin only returns selected metadata fields.
- Requests honor timeoutMs and the tool execution AbortSignal. HTTP 429 and server errors are rethrown so the host can apply its normal retry policy.

## Limitations

- The MVP supports a delegated token only; application-only credentials and OAuth flows are intentionally outside the plugin.
- Search pagination uses Microsoft Graph offset semantics and can be affected by changes in the mailbox or chat index between calls.
- Calendar and OneDrive cursors are process-local. Restarting the host requires starting a new first-page request.
- No send, write, delete, upload, download, subscription, or administrative tools are included.

## Development

~~~sh
npm install
npm run typecheck
npm run build
npm pack --dry-run
~~~

See [DEVELOPMENT.md](DEVELOPMENT.md) for endpoint mapping and release checks.

## License

[MIT](LICENSE)
