import { randomUUID } from 'node:crypto'
import { describe, expect, it, vi } from 'vitest'
import { Microsoft365Client, MicrosoftGraphError } from '../src/client.ts'

function jsonResponse(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json', ...headers } })
}

const testToken = process.env.M365_TEST_TOKEN ?? `m365-test-${randomUUID()}`

function client(fetchImpl: ReturnType<typeof vi.fn>, options: { accessToken?: string } = {}) {
  return new Microsoft365Client({ accessToken: options.accessToken ?? testToken, fetchImpl })
}

function searchBody(value: { entityType: string; query: string; from: number; size: number }) {
  return {
    requests: [{
      entityTypes: [value.entityType],
      query: { queryString: value.query },
      from: value.from,
      size: value.size,
      ...(value.entityType === 'chatMessage' ? { enableTopResults: true } : {}),
    }],
  }
}

describe('Microsoft365Client', () => {
  it('pins the base URL to the Microsoft Graph v1.0 endpoint', () => {
    for (const baseUrl of [
      'http://graph.microsoft.com/v1.0',
      'https://graph.microsoft.com/beta',
      'https://graph.microsoft.com/',
      'https://evil.example.com/v1.0',
      'https://graph.microsoft.com:8443/v1.0',
      'https://graph.microsoft.com/v1.0?x=1',
      'https://graph.microsoft.com/v1.0#frag',
      'not-a-url',
    ]) {
      expect(() => new Microsoft365Client({ baseUrl })).toThrow(/baseUrl must be/)
    }
    expect(() => new Microsoft365Client({ timeoutMs: 0 })).toThrow('timeoutMs')
  })

  it('reads the signed-in profile with a bearer token and never exposes the token', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({
      id: 'u-1', displayName: 'Alice', givenName: 'Alice', surname: 'Doe', userPrincipalName: 'alice@contoso.test',
      mail: 'alice@contoso.test', jobTitle: 'Engineer', department: 'Platform', officeLocation: 'SH', preferredLanguage: 'en-US',
    }))
    const result = await client(fetchImpl).getProfile()

    expect(result).toMatchObject({ id: 'u-1', displayName: 'Alice', mail: 'alice@contoso.test' })
    const [url, init] = fetchImpl.mock.calls[0] as unknown as [URL, RequestInit]
    const profileUrl = new URL(String(url))
    expect(profileUrl.origin + profileUrl.pathname).toBe('https://graph.microsoft.com/v1.0/me')
    expect(profileUrl.searchParams.get('$select')).toBe([
      'id', 'displayName', 'givenName', 'surname', 'userPrincipalName', 'mail', 'jobTitle', 'department', 'officeLocation', 'preferredLanguage',
    ].join(','))
    expect((init.headers as Headers).get('authorization')).toBe(`Bearer ${testToken}`)
    expect((init.headers as Headers).get('accept')).toBe('application/json')
    expect(JSON.stringify(result)).not.toContain(testToken)
  })

  it('searches messages through Microsoft Search and maps hits, totals, and offsets', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({
      value: [{
        hitsContainers: [{
          total: 2,
          moreResultsAvailable: true,
          hits: [
            { summary: 'quarterly numbers', rank: 1, resource: { id: 'msg_1', subject: 'Q3 report', bodyPreview: 'quarterly numbers', receivedDateTime: '2026-10-01T00:00:00Z', from: { emailAddress: { name: 'Billing', address: 'billing@contoso.test' } }, hasAttachments: true } },
            { summary: 'second', rank: 2, resource: { id: 'msg_2', subject: 'Hello' } },
          ],
        }],
      }],
    }))
    const result = await client(fetchImpl).searchMessages({ query: 'quarterly', from: 25, size: 10 })

    expect(result.items).toHaveLength(2)
    expect(result.items[0]).toMatchObject({ id: 'msg_1', subject: 'Q3 report', summary: 'quarterly numbers', rank: 1, sender: { address: 'billing@contoso.test' }, hasAttachments: true })
    expect(result.total).toBe(2)
    expect(result.moreResultsAvailable).toBe(true)
    expect(result.nextFrom).toBe(35)
    const [url, init] = fetchImpl.mock.calls[0] as unknown as [URL, RequestInit]
    expect(String(url)).toBe('https://graph.microsoft.com/v1.0/search/query')
    expect(init.method).toBe('POST')
    expect(JSON.parse(init.body as string)).toEqual(searchBody({ entityType: 'message', query: 'quarterly', from: 25, size: 10 }))
    await expect(client(fetchImpl).searchMessages({ query: '   ' })).rejects.toThrow('query is required')
  })

  it('searches Teams chat messages with top results enabled', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({
      value: [{ hitsContainers: [{ total: 0, moreResultsAvailable: false, hits: [] }] }],
    }))
    const result = await client(fetchImpl).searchChatMessages({ query: 'standup' })

    expect(result).toMatchObject({ items: [], total: 0, moreResultsAvailable: false, nextFrom: 0 })
    const [, init] = fetchImpl.mock.calls[0] as unknown as [URL, RequestInit]
    expect(JSON.parse(init.body as string)).toEqual(searchBody({ entityType: 'chatMessage', query: 'standup', from: 0, size: 25 }))
  })

  it('lists calendar events with a clamped page size and encoded calendar ids', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({
      value: [{ id: 'evt_1', subject: 'Sync', start: { dateTime: '2026-10-06T09:00:00.0000000', timeZone: 'UTC' }, end: { dateTime: '2026-10-06T10:00:00.0000000', timeZone: 'UTC' }, location: { displayName: 'Room 4' }, isAllDay: false, organizer: { emailAddress: { name: 'Bob', address: 'bob@contoso.test' } }, responseStatus: { response: 'accepted' } }],
      '@odata.nextLink': '',
    }))
    const result = await client(fetchImpl).listCalendarEvents({ top: 500 })
    expect(result.items[0]).toMatchObject({ id: 'evt_1', subject: 'Sync', startDateTime: '2026-10-06T09:00:00.0000000', startTimeZone: 'UTC', location: 'Room 4', responseStatus: 'accepted' })
    expect(result.nextCursor).toBe('')
    const [defaultUrl] = fetchImpl.mock.calls[0] as unknown as [URL]
    const defaultParsed = new URL(String(defaultUrl))
    expect(defaultParsed.origin + defaultParsed.pathname).toBe('https://graph.microsoft.com/v1.0/me/calendar/events')
    expect(defaultParsed.searchParams.get('$top')).toBe('100')

    await client(fetchImpl).listCalendarEvents({ calendarId: 'cal/1' })
    const [encodedUrl] = fetchImpl.mock.calls[1] as unknown as [URL]
    const encodedParsed = new URL(String(encodedUrl))
    expect(encodedParsed.origin + encodedParsed.pathname).toBe('https://graph.microsoft.com/v1.0/me/calendars/cal%2F1/events')
    expect(encodedParsed.searchParams.get('$top')).toBe('25')
  })

  it('requires a start and end window for calendar view', async () => {
    const fetchImpl = vi.fn()
    const fire = client(fetchImpl)
    await expect(fire.listCalendarView()).rejects.toThrow('startDateTime and endDateTime are required')
    expect(fetchImpl).not.toHaveBeenCalled()

    fetchImpl.mockResolvedValueOnce(jsonResponse({ value: [] }))
    await fire.listCalendarView({ startDateTime: '2026-10-01T00:00:00Z', endDateTime: '2026-10-02T00:00:00Z' })
    const [url] = fetchImpl.mock.calls[0] as unknown as [URL]
    const viewUrl = new URL(String(url))
    expect(viewUrl.origin + viewUrl.pathname).toBe('https://graph.microsoft.com/v1.0/me/calendarView')
    expect(viewUrl.searchParams.get('startDateTime')).toBe('2026-10-01T00:00:00Z')
    expect(viewUrl.searchParams.get('endDateTime')).toBe('2026-10-02T00:00:00Z')
    expect(viewUrl.searchParams.get('$orderby')).toBe('start/dateTime')
  })

  it('searches OneDrive items and maps file and folder metadata', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({
      value: [
        { id: 'item_1', name: 'a.pdf', size: 3, file: { mimeType: 'application/pdf' }, parentReference: { path: '/drive/root:' } },
        { id: 'item_2', name: 'Docs', folder: { childCount: 2 } },
      ],
    }))
    const result = await client(fetchImpl).searchDriveItems({ query: 'pdf' })

    expect(result.items[0]).toMatchObject({ id: 'item_1', name: 'a.pdf', fileMimeType: 'application/pdf', isFolder: false })
    expect(result.items[1]).toMatchObject({ id: 'item_2', name: 'Docs', isFolder: true, childCount: 2 })
    const [url] = fetchImpl.mock.calls[0] as unknown as [URL]
    expect(String(url)).toContain('/me/drive/root/search?q=pdf')
    expect(new URL(String(url)).searchParams.get('$top')).toBe('25')
    await expect(client(fetchImpl).searchDriveItems({ query: '' })).rejects.toThrow('query is required')
  })

  it('reads one drive item by id and rejects empty ids', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({ id: 'item/1', name: 'a.pdf', size: 3, file: { mimeType: 'application/pdf' } }))
    const result = await client(fetchImpl).getDriveItem('item/1')

    expect(result).toMatchObject({ id: 'item/1', name: 'a.pdf', fileMimeType: 'application/pdf' })
    const [url] = fetchImpl.mock.calls[0] as unknown as [URL]
    const itemUrl = new URL(String(url))
    expect(itemUrl.origin + itemUrl.pathname).toBe('https://graph.microsoft.com/v1.0/me/drive/items/item%2F1')
    expect(itemUrl.searchParams.get('$select')).toBe([
      'id', 'name', 'size', 'webUrl', 'createdDateTime', 'lastModifiedDateTime', 'file', 'folder', 'parentReference',
    ].join(','))
    await expect(client(fetchImpl).getDriveItem('  ')).rejects.toThrow('Drive item id is required')
  })

  it('maps Graph error payloads into MicrosoftGraphError with sanitized details', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse(
      { error: { code: 'Bad Request; 0', message: 'nope', innerError: { 'request-id': 'req-9' } } },
      400,
      { 'retry-after': '7' },
    ))
    const error = await client(fetchImpl).getProfile().then(() => undefined, (thrown: unknown) => thrown)

    expect(error).toBeInstanceOf(MicrosoftGraphError)
    const graphError = error as MicrosoftGraphError
    expect(graphError.status).toBe(400)
    expect(graphError.code).toBe('GRAPH_ERROR')
    expect(graphError.requestId).toBe('req-9')
    expect(graphError.retryAfter).toBe(7)
    expect(graphError.message).toContain('HTTP 400')
    expect(graphError.message).toContain('requestId=req-9')
  })

  it('issues an opaque cursor for nextLink and accepts it on the next page', async () => {
    const fetchImpl = vi.fn()
      .mockResolvedValueOnce(jsonResponse({ value: [], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/calendar/events?$top=25&$skiptoken=abc' }))
      .mockResolvedValueOnce(jsonResponse({ value: [{ id: 'evt_2', subject: 'Later' }] }))
    const fire = client(fetchImpl)

    const firstPage = await fire.listCalendarEvents()
    expect(firstPage.nextCursor).toMatch(/^m365_[0-9a-f]{32}$/)

    const secondPage = await fire.listCalendarEvents({ cursor: firstPage.nextCursor })
    expect(secondPage.items[0]).toMatchObject({ id: 'evt_2' })
    const [secondUrl] = fetchImpl.mock.calls[1] as unknown as [URL]
    expect(String(secondUrl)).toBe('https://graph.microsoft.com/v1.0/me/calendar/events?$top=25&$skiptoken=abc')
  })

  it('rejects foreign-host, foreign-resource, and unknown cursors without calling fetch', async () => {
    const fire = client(vi.fn())
    await expect(fire.listCalendarEvents({ cursor: 'm365_unknown' })).rejects.toThrow('cursor is invalid or expired')

    const crossHost = client(vi.fn(async () => jsonResponse({ value: [], '@odata.nextLink': 'https://evil.invalid/v1.0/me/calendar/events' })))
    await expect(crossHost.listCalendarEvents()).rejects.toThrow('nextLink for another host')

    const crossResource = client(vi.fn(async () => jsonResponse({ value: [], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/drive/root/search?q=x' })))
    await expect(crossResource.listCalendarEvents()).rejects.toThrow('nextLink for another resource')

    const page = await client(vi.fn(async () => jsonResponse({ value: [], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/drive/root/search?$skiptoken=x' }))).searchDriveItems({ query: 'a' })
    await expect(fire.listCalendarEvents({ cursor: page.nextCursor })).rejects.toThrow('cursor is invalid or expired')
  })
})
