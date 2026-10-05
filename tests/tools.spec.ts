import { randomUUID } from 'node:crypto'
import { describe, expect, it, vi } from 'vitest'
import { Microsoft365Client } from '../src/client.ts'
import { createTools } from '../src/index.ts'

function jsonResponse(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json', ...headers } })
}

const testToken = process.env.M365_TEST_TOKEN ?? `m365-test-${randomUUID()}`

const clientForTest = () => new Microsoft365Client({ accessToken: testToken })

const exec = { signal: new AbortController().signal }

describe('dsh-tool-microsoft-365 tools', () => {
  it('registers the read-only Microsoft 365 tool set', () => {
    expect(createTools(clientForTest()).map(tool => tool.name)).toEqual([
      'ms_auth_test',
      'ms_get_profile',
      'ms_search_messages',
      'ms_list_calendar_events',
      'ms_list_calendar_view',
      'ms_search_drive_items',
      'ms_get_drive_item',
      'ms_search_chat_messages',
    ])
  })

  it('keeps every call a read or search; the package exposes no edits', () => {
    const requiredArgs: Record<string, Record<string, string>> = {
      ms_search_messages: { query: 'invoice' },
      ms_get_drive_item: { itemId: 'item_1' },
      ms_search_chat_messages: { query: 'standup' },
    }
    for (const tool of createTools(clientForTest())) {
      expect(['read', 'search']).toContain(tool.presentCall(requiredArgs[tool.name] ?? {})?.kind)
    }
  })

  it('renders profile, message, event, and drive output', () => {
    const tools = createTools(clientForTest())

    const profile = tools.find(item => item.name === 'ms_get_profile')!
    expect((profile.output.render({}, { found: true, displayName: 'Alice', id: 'u-1', userPrincipalName: 'alice@contoso.test', mail: 'alice@contoso.test' }) as Array<{ text: string }>)[0].text)
      .toContain('Alice (u-1)')

    const messages = tools.find(item => item.name === 'ms_search_messages')!
    const messageView = (messages.output.render({}, {
      found: true,
      items: [{ id: 'msg_1', subject: 'Q3 report', receivedDateTime: '2026-10-01T00:00:00Z', sender: { name: 'Billing' }, hasAttachments: true }],
      total: 1,
      nextFrom: 0,
    }) as Array<{ text: string }>)[0].text
    expect(messageView).toContain('Q3 report (msg_1)')
    expect(messageView).toContain('from=Billing')
    expect(messageView).toContain('attachments')
    expect(messageView).toContain('total=1')

    const events = tools.find(item => item.name === 'ms_list_calendar_events')!
    const eventView = (events.output.render({}, {
      found: true,
      items: [{ id: 'evt_1', subject: 'Sync', startDateTime: '2026-10-06T09:00:00Z', endDateTime: '2026-10-06T10:00:00Z', location: 'Room 4' }],
      nextCursor: 'm365_next',
    }) as Array<{ text: string }>)[0].text
    expect(eventView).toContain('Sync (evt_1) 2026-10-06T09:00:00Z -> 2026-10-06T10:00:00Z @ Room 4')
    expect(eventView).toContain('nextCursor=m365_next')

    const drive = tools.find(item => item.name === 'ms_search_drive_items')!
    const driveView = (drive.output.render({}, {
      found: true,
      items: [{ id: 'item_1', name: 'a.pdf', fileMimeType: 'application/pdf', size: 3, lastModifiedDateTime: '2026-10-01T00:00:00Z' }],
      nextCursor: '',
    }) as Array<{ text: string }>)[0].text
    expect(driveView).toContain('a.pdf (item_1) application/pdf size=3')
  })

  it('reports a missing token without calling fetch', async () => {
    const fetchImpl = vi.fn()
    const tools = createTools(new Microsoft365Client({ accessToken: '', fetchImpl }))

    const auth = tools.find(item => item.name === 'ms_auth_test')!
    await expect(auth.execute({}, exec)).resolves.toMatchObject({ ok: false, reason: 'Microsoft Graph accessToken is not configured.' })

    const search = tools.find(item => item.name === 'ms_search_messages')!
    await expect(search.execute({ query: 'x' }, exec)).resolves.toMatchObject({ found: false })
    expect(fetchImpl).not.toHaveBeenCalled()
  })

  it('executes auth and search through the tool layer without exposing the token', async () => {
    const fetchImpl = vi.fn()
      .mockResolvedValueOnce(jsonResponse({ id: 'u-1', displayName: 'Alice', userPrincipalName: 'alice@contoso.test' }))
      .mockResolvedValueOnce(jsonResponse({ value: [{ hitsContainers: [{ total: 1, moreResultsAvailable: false, hits: [{ summary: 's', rank: 1, resource: { id: 'msg_1', subject: 'Hi' } }] }] }] }))
    const tools = createTools(new Microsoft365Client({ accessToken: testToken, fetchImpl }))

    const auth = tools.find(item => item.name === 'ms_auth_test')!
    const authResult = await auth.execute({}, exec)
    expect(authResult).toMatchObject({ ok: true, id: 'u-1', displayName: 'Alice' })
    expect(JSON.stringify(authResult)).not.toContain(testToken)

    const search = tools.find(item => item.name === 'ms_search_messages')!
    const searchResult = await search.execute({ query: 'hi' }, exec)
    expect(searchResult).toMatchObject({ found: true, total: 1, items: [{ id: 'msg_1', summary: 's', rank: 1 }] })
    expect(fetchImpl).toHaveBeenCalledTimes(2)
  })

  it('maps client errors into found:false and rethrows server pressure', async () => {
    const failing = vi.fn(async () => jsonResponse({ error: { code: 'BadRequest', message: 'bad' } }, 400))
    const tools = createTools(new Microsoft365Client({ accessToken: testToken, fetchImpl: failing }))
    const search = tools.find(item => item.name === 'ms_search_messages')!
    await expect(search.execute({ query: 'x' }, exec)).resolves.toMatchObject({ found: false, reason: expect.stringContaining('HTTP 400') })

    const throttled = vi.fn(async () => jsonResponse({ error: { code: 'TooManyRequests', message: 'slow down' } }, 429))
    const throttledTools = createTools(new Microsoft365Client({ accessToken: testToken, fetchImpl: throttled }))
    const throttledSearch = throttledTools.find(item => item.name === 'ms_search_messages')!
    await expect(throttledSearch.execute({ query: 'x' }, exec)).rejects.toBeInstanceOf(Error)
  })
})
