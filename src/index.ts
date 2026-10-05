import type { Context } from '@deepseek-ai/cordis'
import type { ToolCallView } from '@deepseek-ai/dsh-tools'
import { defineTool } from '@deepseek-ai/dsh-tools'
import {
  Microsoft365Client,
  MicrosoftGraphError,
  type Microsoft365ClientOptions,
} from './client.js'

export { Microsoft365Client, MicrosoftGraphError }
export type {
  Microsoft365ClientOptions,
  MicrosoftCalendarEventInfo,
  MicrosoftChatMessageInfo,
  MicrosoftCollectionResult,
  MicrosoftDriveItemInfo,
  MicrosoftMessageInfo,
  MicrosoftProfileInfo,
  MicrosoftSearchResult,
  MicrosoftSenderInfo,
} from './client.js'

export const name = 'dsh-tool-microsoft-365'
export const inject = ['tools']
export interface Microsoft365PluginConfig extends Microsoft365ClientOptions {}

function text(value: string) { return [{ type: 'text' as const, text: value }] }
function unavailable(reason: string) { return { found: false, reason } }
function readKind(title: string, kind: 'read' | 'search' = 'read'): ToolCallView { return { card: 'generic', title, kind } }
function failure(error: unknown): { found: false; reason: string } {
  if (error instanceof MicrosoftGraphError && (error.status === 429 || error.status >= 500)) throw error
  return { found: false, reason: error instanceof Error ? error.message : 'Microsoft Graph request failed.' }
}
function authFailure(error: unknown): { ok: false; reason: string } {
  if (error instanceof MicrosoftGraphError && (error.status === 429 || error.status >= 500)) throw error
  return { ok: false, reason: error instanceof Error ? error.message : 'Microsoft Graph authentication failed.' }
}

const profileProperties = {
  id: { type: 'string' }, displayName: { type: 'string' }, givenName: { type: 'string' }, surname: { type: 'string' },
  userPrincipalName: { type: 'string' }, mail: { type: 'string' }, jobTitle: { type: 'string' }, department: { type: 'string' },
  officeLocation: { type: 'string' }, preferredLanguage: { type: 'string' },
} as const

const senderProperties = { name: { type: 'string' }, address: { type: 'string' } } as const

const messageProperties = {
  id: { type: 'string' }, subject: { type: 'string' }, summary: { type: 'string' }, bodyPreview: { type: 'string' },
  receivedDateTime: { type: 'string' }, sentDateTime: { type: 'string' }, createdDateTime: { type: 'string' },
  lastModifiedDateTime: { type: 'string' }, isRead: { type: 'boolean' }, hasAttachments: { type: 'boolean' },
  importance: { type: 'string' }, sender: { type: 'object', additionalProperties: false, properties: senderProperties },
  webLink: { type: 'string' }, rank: { type: 'integer' },
} as const

const calendarProperties = {
  id: { type: 'string' }, subject: { type: 'string' }, startDateTime: { type: 'string' }, startTimeZone: { type: 'string' },
  endDateTime: { type: 'string' }, endTimeZone: { type: 'string' }, location: { type: 'string' }, isAllDay: { type: 'boolean' },
  organizer: { type: 'object', additionalProperties: false, properties: senderProperties }, responseStatus: { type: 'string' },
  webLink: { type: 'string' },
} as const

const driveProperties = {
  id: { type: 'string' }, name: { type: 'string' }, size: { type: 'integer' }, webUrl: { type: 'string' },
  createdDateTime: { type: 'string' }, lastModifiedDateTime: { type: 'string' }, fileMimeType: { type: 'string' },
  isFolder: { type: 'boolean' }, childCount: { type: 'integer' }, parentPath: { type: 'string' },
} as const

const chatProperties = {
  id: { type: 'string' }, createdDateTime: { type: 'string' }, lastModifiedDateTime: { type: 'string' }, bodyPreview: { type: 'string' },
  sender: { type: 'object', additionalProperties: false, properties: senderProperties }, importance: { type: 'string' },
  webLink: { type: 'string' }, rank: { type: 'integer' },
} as const

function renderProfile(value: { id?: string; displayName?: string; userPrincipalName?: string; mail?: string; jobTitle?: string; department?: string; officeLocation?: string }) {
  return text([
    (value.displayName || '') + ' (' + (value.id || '') + ')',
    'userPrincipalName=' + (value.userPrincipalName || ''),
    value.mail ? 'mail=' + value.mail : '',
    value.jobTitle ? 'jobTitle=' + value.jobTitle : '',
    value.department ? 'department=' + value.department : '',
    value.officeLocation ? 'officeLocation=' + value.officeLocation : '',
  ].filter(Boolean).join('\n'))
}

function renderMessages(items: Array<{ subject?: string; id?: string; receivedDateTime?: string; sender?: { name?: string; address?: string }; hasAttachments?: boolean }>, total = 0, nextFrom = 0) {
  const lines = items.map(item => (item.subject || '(no subject)') + ' (' + (item.id || '') + ') ' + (item.receivedDateTime || '') + ' from=' + (item.sender?.name || item.sender?.address || '') + (item.hasAttachments ? ' attachments' : ''))
  return text(lines.concat('total=' + total + (nextFrom ? ' nextFrom=' + nextFrom : '')).join('\n'))
}

function renderEvents(items: Array<{ subject?: string; id?: string; startDateTime?: string; endDateTime?: string; location?: string; isAllDay?: boolean }>, nextCursor = '') {
  const lines = items.map(item => (item.subject || '(no subject)') + ' (' + (item.id || '') + ') ' + (item.startDateTime || '') + ' -> ' + (item.endDateTime || '') + (item.location ? ' @ ' + item.location : '') + (item.isAllDay ? ' all-day' : ''))
  return text(lines.concat(nextCursor ? 'nextCursor=' + nextCursor : []).join('\n'))
}

function renderDrive(items: Array<{ name?: string; id?: string; fileMimeType?: string; size?: number; lastModifiedDateTime?: string }>, nextCursor = '') {
  const lines = items.map(item => (item.name || '') + ' (' + (item.id || '') + ') ' + (item.fileMimeType || 'folder') + ' size=' + (item.size || 0) + ' modified=' + (item.lastModifiedDateTime || ''))
  return text(lines.concat(nextCursor ? 'nextCursor=' + nextCursor : []).join('\n'))
}

function renderChat(items: Array<{ id?: string; bodyPreview?: string; createdDateTime?: string; sender?: { name?: string; address?: string } }>, total = 0, nextFrom = 0) {
  const lines = items.map(item => (item.createdDateTime || '') + ' ' + (item.sender?.name || item.sender?.address || '') + ' (' + (item.id || '') + ') ' + (item.bodyPreview || ''))
  return text(lines.concat('total=' + total + (nextFrom ? ' nextFrom=' + nextFrom : '')).join('\n'))
}

export function createTools(client: Microsoft365Client) {
  return [
    defineTool({
      name: 'ms_auth_test',
      description: 'Verify a delegated Microsoft Graph access token and return safe signed-in profile fields without exposing the token.',
      parameters: {},
      output: { schema: { type: 'object', additionalProperties: false, properties: { ok: { type: 'boolean' }, reason: { type: 'string' }, ...profileProperties } } as const, render: (_args, value) => value.ok ? renderProfile(value) : text('Microsoft Graph auth failed: ' + value.reason) },
      presentCall(): ToolCallView { return readKind('Verify Microsoft Graph credentials') },
      async execute(_args, exec) {
        if (!client.hasToken()) return { ok: false, reason: 'Microsoft Graph accessToken is not configured.' }
        try { return { ok: true, ...await client.authTest(exec.signal) } } catch (error) { return authFailure(error) }
      },
    }),
    defineTool({
      name: 'ms_get_profile',
      description: 'Read the signed-in Microsoft 365 user profile with a small allowlist of identity and organization fields.',
      parameters: {},
      output: { schema: { type: 'object', additionalProperties: false, properties: { found: { type: 'boolean' }, reason: { type: 'string' }, ...profileProperties } } as const, render: (_args, value) => value.found ? renderProfile(value) : text(value.reason || 'Microsoft Graph is not configured.') },
      presentCall(): ToolCallView { return readKind('Microsoft 365 profile') },
      async execute(_args, exec) {
        if (!client.hasToken()) return unavailable('Microsoft Graph accessToken is not configured.')
        try { return { found: true, ...await client.getProfile(exec.signal) } } catch (error) { return failure(error) }
      },
    }),
    defineTool({
      name: 'ms_search_messages',
      description: 'Search the signed-in user\'s Outlook messages with Microsoft Search and bounded offset pagination. Message bodies are limited to bodyPreview.',
      parameters: { query: { type: 'string', required: true, description: 'Outlook search query string.' }, from: { type: 'integer', description: 'Zero-based result offset from a previous response.' }, size: { type: 'integer', description: 'Results per request, 1-100 (default 25).' } },
      output: { schema: { type: 'object', additionalProperties: false, properties: { found: { type: 'boolean' }, reason: { type: 'string' }, items: { type: 'array', items: { type: 'object', additionalProperties: false, properties: messageProperties } }, total: { type: 'integer' }, moreResultsAvailable: { type: 'boolean' }, nextFrom: { type: 'integer' } } } as const, render: (_args, value) => value.found ? renderMessages(value.items || [], value.total, value.nextFrom) : text(value.reason || 'Microsoft Graph is not configured.') },
      presentCall(args): ToolCallView { return readKind('Search Outlook messages' + (args.query ? ': ' + args.query : ''), 'search') },
      async execute(args, exec) {
        if (!client.hasToken()) return unavailable('Microsoft Graph accessToken is not configured.')
        try { return { found: true, ...await client.searchMessages({ query: args.query as string, from: args.from as number, size: args.size as number, signal: exec.signal }) } } catch (error) { return failure(error) }
      },
    }),
    defineTool({
      name: 'ms_list_calendar_events',
      description: 'List events from the signed-in user\'s default calendar with a bounded opaque cursor.',
      parameters: { calendarId: { type: 'string', description: 'Optional calendar ID; defaults to /me/calendar.' }, top: { type: 'integer', description: 'Results per request, 1-100 (default 25).' }, skipToken: { type: 'string', description: 'Optional opaque skip token.' }, cursor: { type: 'string', description: 'Opaque cursor returned by an earlier response; expires with this plugin process.' } },
      output: { schema: { type: 'object', additionalProperties: false, properties: { found: { type: 'boolean' }, reason: { type: 'string' }, items: { type: 'array', items: { type: 'object', additionalProperties: false, properties: calendarProperties } }, nextCursor: { type: 'string' } } } as const, render: (_args, value) => value.found ? renderEvents(value.items || [], value.nextCursor) : text(value.reason || 'Microsoft Graph is not configured.') },
      presentCall(args): ToolCallView { return readKind('Microsoft calendar events' + (args.calendarId ? ': ' + args.calendarId : ''), 'search') },
      async execute(args, exec) {
        if (!client.hasToken()) return unavailable('Microsoft Graph accessToken is not configured.')
        try { return { found: true, ...await client.listCalendarEvents({ calendarId: args.calendarId as string, top: args.top as number, skipToken: args.skipToken as string, cursor: args.cursor as string, signal: exec.signal }) } } catch (error) { return failure(error) }
      },
    }),
    defineTool({
      name: 'ms_list_calendar_view',
      description: 'List events in a start/end window from the signed-in user\'s calendar, including recurring instances, with a bounded opaque cursor.',
      parameters: { startDateTime: { type: 'string', description: 'ISO 8601 window start; required on the first request.' }, endDateTime: { type: 'string', description: 'ISO 8601 window end; required on the first request.' }, top: { type: 'integer', description: 'Results per request, 1-100 (default 25).' }, skipToken: { type: 'string', description: 'Optional opaque skip token.' }, cursor: { type: 'string', description: 'Opaque cursor returned by an earlier response; expires with this plugin process.' } },
      output: { schema: { type: 'object', additionalProperties: false, properties: { found: { type: 'boolean' }, reason: { type: 'string' }, items: { type: 'array', items: { type: 'object', additionalProperties: false, properties: calendarProperties } }, nextCursor: { type: 'string' } } } as const, render: (_args, value) => value.found ? renderEvents(value.items || [], value.nextCursor) : text(value.reason || 'Microsoft Graph is not configured.') },
      presentCall(args): ToolCallView { return readKind('Microsoft calendar view' + (args.startDateTime ? ': ' + args.startDateTime : ''), 'search') },
      async execute(args, exec) {
        if (!client.hasToken()) return unavailable('Microsoft Graph accessToken is not configured.')
        try { return { found: true, ...await client.listCalendarView({ startDateTime: args.startDateTime as string, endDateTime: args.endDateTime as string, top: args.top as number, skipToken: args.skipToken as string, cursor: args.cursor as string, signal: exec.signal }) } } catch (error) { return failure(error) }
      },
    }),
    defineTool({
      name: 'ms_search_drive_items',
      description: 'Search the signed-in user\'s OneDrive for files and folders with bounded page size and an opaque cursor. Download URLs and file content are never returned.',
      parameters: { query: { type: 'string', description: 'OneDrive search text; required on the first request.' }, top: { type: 'integer', description: 'Results per request, 1-100 (default 25).' }, skipToken: { type: 'string', description: 'Optional opaque skip token.' }, cursor: { type: 'string', description: 'Opaque cursor returned by an earlier response; expires with this plugin process.' } },
      output: { schema: { type: 'object', additionalProperties: false, properties: { found: { type: 'boolean' }, reason: { type: 'string' }, items: { type: 'array', items: { type: 'object', additionalProperties: false, properties: driveProperties } }, nextCursor: { type: 'string' } } } as const, render: (_args, value) => value.found ? renderDrive(value.items || [], value.nextCursor) : text(value.reason || 'Microsoft Graph is not configured.') },
      presentCall(args): ToolCallView { return readKind('Search OneDrive' + (args.query ? ': ' + args.query : ''), 'search') },
      async execute(args, exec) {
        if (!client.hasToken()) return unavailable('Microsoft Graph accessToken is not configured.')
        try { return { found: true, ...await client.searchDriveItems({ query: args.query as string, top: args.top as number, skipToken: args.skipToken as string, cursor: args.cursor as string, signal: exec.signal }) } } catch (error) { return failure(error) }
      },
    }),
    defineTool({
      name: 'ms_get_drive_item',
      description: 'Read safe metadata for one OneDrive item by ID. File content and preauthenticated download URLs are omitted.',
      parameters: { itemId: { type: 'string', required: true, description: 'OneDrive driveItem ID.' } },
      output: { schema: { type: 'object', additionalProperties: false, properties: { found: { type: 'boolean' }, reason: { type: 'string' }, ...driveProperties } } as const, render: (_args, value) => value.found ? renderDrive([value]) : text(value.reason || 'Microsoft Graph is not configured.') },
      presentCall(args): ToolCallView { return readKind('OneDrive item ' + (args.itemId || '')) },
      async execute(args, exec) {
        if (!client.hasToken()) return unavailable('Microsoft Graph accessToken is not configured.')
        try { return { found: true, ...await client.getDriveItem(args.itemId as string, exec.signal) } } catch (error) { return failure(error) }
      },
    }),
    defineTool({
      name: 'ms_search_chat_messages',
      description: 'Search Teams chat messages through Microsoft Search with bounded offset pagination. Message content is reduced to a short preview.',
      parameters: { query: { type: 'string', required: true, description: 'Teams chat message search query string.' }, from: { type: 'integer', description: 'Zero-based result offset from a previous response.' }, size: { type: 'integer', description: 'Results per request, 1-100 (default 25).' } },
      output: { schema: { type: 'object', additionalProperties: false, properties: { found: { type: 'boolean' }, reason: { type: 'string' }, items: { type: 'array', items: { type: 'object', additionalProperties: false, properties: chatProperties } }, total: { type: 'integer' }, moreResultsAvailable: { type: 'boolean' }, nextFrom: { type: 'integer' } } } as const, render: (_args, value) => value.found ? renderChat(value.items || [], value.total, value.nextFrom) : text(value.reason || 'Microsoft Graph is not configured.') },
      presentCall(args): ToolCallView { return readKind('Search Teams chat messages' + (args.query ? ': ' + args.query : ''), 'search') },
      async execute(args, exec) {
        if (!client.hasToken()) return unavailable('Microsoft Graph accessToken is not configured.')
        try { return { found: true, ...await client.searchChatMessages({ query: args.query as string, from: args.from as number, size: args.size as number, signal: exec.signal }) } } catch (error) { return failure(error) }
      },
    }),
  ]
}

export function apply(ctx: Context, config: Microsoft365PluginConfig = {}): void {
  const client = new Microsoft365Client(config)
  for (const tool of createTools(client)) ctx.tools.register(tool)
}
