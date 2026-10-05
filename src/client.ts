/** Small, read-only Microsoft Graph v1.0 client with injected fetch support. */

import { randomUUID } from 'node:crypto'

export interface Microsoft365ClientOptions {
  /** Delegated Microsoft Graph access token. It is never returned in a result. */
  accessToken?: string
  /** Absolute Graph v1.0 endpoint. Defaults to https://graph.microsoft.com/v1.0. */
  baseUrl?: string
  timeoutMs?: number
  fetchImpl?: typeof fetch
}

export class MicrosoftGraphError extends Error {
  constructor(
    message: string,
    public readonly status: number,
    public readonly code = '',
    public readonly requestId?: string,
    public readonly retryAfter?: number,
  ) {
    super(message)
    this.name = 'MicrosoftGraphError'
  }
}

export interface MicrosoftProfileInfo {
  id: string
  displayName: string
  givenName: string
  surname: string
  userPrincipalName: string
  mail: string
  jobTitle: string
  department: string
  officeLocation: string
  preferredLanguage: string
}

export interface MicrosoftSenderInfo {
  name: string
  address: string
}

export interface MicrosoftMessageInfo {
  id: string
  subject: string
  summary: string
  bodyPreview: string
  receivedDateTime: string
  sentDateTime: string
  createdDateTime: string
  lastModifiedDateTime: string
  isRead: boolean
  hasAttachments: boolean
  importance: string
  sender: MicrosoftSenderInfo
  webLink: string
  rank: number
}

export interface MicrosoftCalendarEventInfo {
  id: string
  subject: string
  startDateTime: string
  startTimeZone: string
  endDateTime: string
  endTimeZone: string
  location: string
  isAllDay: boolean
  organizer: MicrosoftSenderInfo
  responseStatus: string
  webLink: string
}

export interface MicrosoftDriveItemInfo {
  id: string
  name: string
  size: number
  webUrl: string
  createdDateTime: string
  lastModifiedDateTime: string
  fileMimeType: string
  isFolder: boolean
  childCount: number
  parentPath: string
}

export interface MicrosoftChatMessageInfo {
  id: string
  createdDateTime: string
  lastModifiedDateTime: string
  bodyPreview: string
  sender: MicrosoftSenderInfo
  importance: string
  webLink: string
  rank: number
}

export interface MicrosoftSearchResult<T> {
  items: T[]
  total: number
  moreResultsAvailable: boolean
  nextFrom: number
}

export interface MicrosoftCollectionResult<T> {
  items: T[]
  /** Process-local opaque cursor. The Graph nextLink is never returned. */
  nextCursor: string
}

function record(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : {}
}

function array(value: unknown): unknown[] {
  return Array.isArray(value) ? value : []
}

function stringValue(value: unknown): string {
  return typeof value === 'string' ? value : value == null ? '' : String(value)
}

function numberValue(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : Number(value ?? 0) || 0
}

function boolValue(value: unknown): boolean {
  return value === true
}

function clamp(value: number | undefined, min: number, max: number, fallback: number): number {
  return Number.isFinite(value) ? Math.max(min, Math.min(max, Math.trunc(value as number))) : fallback
}

function truncate(value: string, max = 1000): string {
  return value.length > max ? value.slice(0, max) + '…' : value
}

function emailAddress(value: unknown): MicrosoftSenderInfo {
  const outer = record(value)
  const nested = record(outer.emailAddress)
  return {
    name: truncate(stringValue(nested.name) || stringValue(outer.displayName) || stringValue(outer.name), 200),
    address: truncate(stringValue(nested.address) || stringValue(outer.address) || stringValue(outer.email), 320),
  }
}

function sender(value: unknown): MicrosoftSenderInfo {
  const raw = record(value)
  const user = record(raw.user)
  return {
    name: truncate(stringValue(user.displayName) || emailAddress(raw).name || stringValue(raw.displayName), 200),
    address: truncate(stringValue(user.email) || emailAddress(raw).address, 320),
  }
}

function mapProfile(value: unknown): MicrosoftProfileInfo {
  const raw = record(value)
  return {
    id: truncate(stringValue(raw.id), 256),
    displayName: truncate(stringValue(raw.displayName), 200),
    givenName: truncate(stringValue(raw.givenName), 100),
    surname: truncate(stringValue(raw.surname), 100),
    userPrincipalName: truncate(stringValue(raw.userPrincipalName), 320),
    mail: truncate(stringValue(raw.mail), 320),
    jobTitle: truncate(stringValue(raw.jobTitle), 200),
    department: truncate(stringValue(raw.department), 200),
    officeLocation: truncate(stringValue(raw.officeLocation), 200),
    preferredLanguage: truncate(stringValue(raw.preferredLanguage), 40),
  }
}

function mapMessage(value: unknown, summary = '', rank = 0): MicrosoftMessageInfo {
  const raw = record(value)
  const from = emailAddress(raw.from)
  const fallbackSender = emailAddress(raw.sender)
  return {
    id: truncate(stringValue(raw.id), 256),
    subject: truncate(stringValue(raw.subject), 300),
    summary: truncate(summary, 1000),
    bodyPreview: truncate(stringValue(raw.bodyPreview), 1000),
    receivedDateTime: truncate(stringValue(raw.receivedDateTime), 80),
    sentDateTime: truncate(stringValue(raw.sentDateTime), 80),
    createdDateTime: truncate(stringValue(raw.createdDateTime), 80),
    lastModifiedDateTime: truncate(stringValue(raw.lastModifiedDateTime), 80),
    isRead: boolValue(raw.isRead),
    hasAttachments: boolValue(raw.hasAttachments),
    importance: stringValue(raw.importance),
    sender: from.address || from.name ? from : fallbackSender,
    webLink: truncate(stringValue(raw.webLink), 2000),
    rank,
  }
}

function mapCalendarEvent(value: unknown): MicrosoftCalendarEventInfo {
  const raw = record(value)
  const start = record(raw.start)
  const end = record(raw.end)
  const response = record(raw.responseStatus)
  return {
    id: truncate(stringValue(raw.id), 256),
    subject: truncate(stringValue(raw.subject), 300),
    startDateTime: truncate(stringValue(start.dateTime), 80),
    startTimeZone: truncate(stringValue(start.timeZone), 80),
    endDateTime: truncate(stringValue(end.dateTime), 80),
    endTimeZone: truncate(stringValue(end.timeZone), 80),
    location: truncate(stringValue(record(raw.location).displayName), 300),
    isAllDay: boolValue(raw.isAllDay),
    organizer: emailAddress(raw.organizer),
    responseStatus: truncate(stringValue(response.response) || stringValue(response.status), 80),
    webLink: truncate(stringValue(raw.webLink), 2000),
  }
}

function mapDriveItem(value: unknown): MicrosoftDriveItemInfo {
  const raw = record(value)
  const file = record(raw.file)
  const folder = record(raw.folder)
  const parent = record(raw.parentReference)
  return {
    id: truncate(stringValue(raw.id), 256),
    name: truncate(stringValue(raw.name), 300),
    size: numberValue(raw.size),
    webUrl: truncate(stringValue(raw.webUrl), 2000),
    createdDateTime: truncate(stringValue(raw.createdDateTime), 80),
    lastModifiedDateTime: truncate(stringValue(raw.lastModifiedDateTime), 80),
    fileMimeType: truncate(stringValue(file.mimeType), 200),
    isFolder: Object.keys(folder).length > 0,
    childCount: numberValue(folder.childCount),
    parentPath: truncate(stringValue(parent.path), 500),
  }
}

function plainPreview(value: unknown): string {
  const raw = record(value)
  const content = stringValue(raw.content)
  return truncate(content.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim(), 1000)
}

function mapChatMessage(value: unknown, rank = 0): MicrosoftChatMessageInfo {
  const raw = record(value)
  return {
    id: truncate(stringValue(raw.id), 256),
    createdDateTime: truncate(stringValue(raw.createdDateTime), 80),
    lastModifiedDateTime: truncate(stringValue(raw.lastModifiedDateTime), 80),
    bodyPreview: plainPreview(raw.body),
    sender: sender(raw.from),
    importance: truncate(stringValue(raw.importance), 80),
    webLink: truncate(stringValue(raw.webUrl) || stringValue(raw.webLink), 2000),
    rank,
  }
}

function searchHits(body: unknown): { hits: unknown[]; total: number; more: boolean } {
  const values = array(record(body).value)
  const first = record(values[0])
  const containers = array(first.hitsContainers)
  const container = record(containers[0])
  return { hits: array(container.hits), total: numberValue(container.total), more: boolValue(container.moreResultsAvailable) }
}

function errorDetails(value: unknown): { code: string; requestId?: string } {
  const error = record(record(value).error)
  const inner = record(error.innerError)
  return {
    code: stringValue(error.code),
    requestId: stringValue(inner['request-id']) || stringValue(inner.requestId) || undefined,
  }
}

function graphSelect(...fields: string[]): string {
  return fields.join(',')
}

function query(params: Record<string, string | number | undefined>): string {
  const search = new URLSearchParams()
  for (const [key, value] of Object.entries(params)) if (value !== undefined && value !== '') search.set(key, String(value))
  const result = search.toString()
  return result ? '?' + result : ''
}

export class Microsoft365Client {
  private readonly accessToken: string
  private readonly baseUrl: string
  private readonly baseOrigin: string
  private readonly basePath: string
  private readonly timeoutMs: number
  private readonly fetchImpl: typeof fetch
  private readonly cursors = new Map<string, string>()

  constructor(options: Microsoft365ClientOptions = {}) {
    this.accessToken = options.accessToken ?? ''
    const baseUrl = (options.baseUrl ?? 'https://graph.microsoft.com/v1.0').replace(/\/+$/, '')
    let parsed: URL
    try { parsed = new URL(baseUrl) } catch { throw new Error('baseUrl must be the Microsoft Graph v1.0 endpoint.') }
    if (parsed.protocol !== 'https:' || parsed.hostname !== 'graph.microsoft.com' || parsed.port || parsed.search || parsed.hash || parsed.pathname !== '/v1.0') {
      throw new Error('baseUrl must be https://graph.microsoft.com/v1.0; custom hosts are rejected because the Bearer token must not leave Microsoft Graph.')
    }
    this.baseUrl = baseUrl
    this.baseOrigin = parsed.origin
    this.basePath = parsed.pathname
    this.timeoutMs = options.timeoutMs ?? 15_000
    if (!Number.isFinite(this.timeoutMs) || this.timeoutMs <= 0) throw new Error('timeoutMs must be a positive finite number.')
    this.fetchImpl = options.fetchImpl ?? globalThis.fetch
  }

  hasToken(): boolean { return this.accessToken.length > 0 }

  async authTest(signal?: AbortSignal): Promise<MicrosoftProfileInfo> {
    return this.getProfile(signal)
  }

  async getProfile(signal?: AbortSignal): Promise<MicrosoftProfileInfo> {
    const data = await this.request<unknown>('/me' + query({ '$select': graphSelect('id', 'displayName', 'givenName', 'surname', 'userPrincipalName', 'mail', 'jobTitle', 'department', 'officeLocation', 'preferredLanguage') }), {}, signal)
    return mapProfile(data)
  }

  async searchMessages(options: { query: string; from?: number; size?: number; signal?: AbortSignal }): Promise<MicrosoftSearchResult<MicrosoftMessageInfo>> {
    const queryText = (options.query ?? '').trim()
    if (!queryText) throw new Error('Microsoft Graph message search query is required.')
    const from = clamp(options.from, 0, 100_000, 0)
    const size = clamp(options.size, 1, 100, 25)
    const data = await this.search('message', queryText, from, size, options.signal)
    return { items: data.hits.map(hit => { const raw = record(hit); return mapMessage(raw.resource, stringValue(raw.summary), numberValue(raw.rank)) }), total: data.total, moreResultsAvailable: data.more, nextFrom: data.more ? from + size : 0 }
  }

  async searchChatMessages(options: { query: string; from?: number; size?: number; signal?: AbortSignal }): Promise<MicrosoftSearchResult<MicrosoftChatMessageInfo>> {
    const queryText = (options.query ?? '').trim()
    if (!queryText) throw new Error('Microsoft Graph chat message search query is required.')
    const from = clamp(options.from, 0, 100_000, 0)
    const size = clamp(options.size, 1, 100, 25)
    const data = await this.search('chatMessage', queryText, from, size, options.signal)
    return { items: data.hits.map(hit => { const raw = record(hit); return mapChatMessage(raw.resource, numberValue(raw.rank)) }), total: data.total, moreResultsAvailable: data.more, nextFrom: data.more ? from + size : 0 }
  }

  async listCalendarEvents(options: { calendarId?: string; top?: number; skipToken?: string; cursor?: string; signal?: AbortSignal } = {}): Promise<MicrosoftCollectionResult<MicrosoftCalendarEventInfo>> {
    const pathPattern = new RegExp('^/me/(?:calendar/events|calendars/[^/]+/events)$')
    const path = options.cursor ? this.acceptCursor(options.cursor, pathPattern) : (options.calendarId ? '/me/calendars/' + encodeURIComponent(options.calendarId) + '/events' : '/me/calendar/events')
    const url = options.cursor ? path : path + query({ '$top': clamp(options.top, 1, 100, 25), '$skiptoken': options.skipToken })
    const data = await this.request<unknown>(url, {}, options.signal)
    const raw = record(data)
    return { items: array(raw.value).map(mapCalendarEvent), nextCursor: this.normalizeNextLink(raw['@odata.nextLink'], pathPattern) }
  }

  async listCalendarView(options: { startDateTime?: string; endDateTime?: string; top?: number; skipToken?: string; cursor?: string; signal?: AbortSignal } = {}): Promise<MicrosoftCollectionResult<MicrosoftCalendarEventInfo>> {
    if (!options.cursor && (!options.startDateTime || !options.endDateTime)) throw new Error('startDateTime and endDateTime are required for calendar view.')
    const pathPattern = new RegExp('^/me/calendarView$')
    const path = options.cursor ? this.acceptCursor(options.cursor, pathPattern) : '/me/calendarView'
    const url = options.cursor ? path : path + query({ startDateTime: options.startDateTime, endDateTime: options.endDateTime, '$top': clamp(options.top, 1, 100, 25), '$skiptoken': options.skipToken, '$orderby': 'start/dateTime' })
    const data = await this.request<unknown>(url, {}, options.signal)
    const raw = record(data)
    return { items: array(raw.value).map(mapCalendarEvent), nextCursor: this.normalizeNextLink(raw['@odata.nextLink'], pathPattern) }
  }

  async searchDriveItems(options: { query: string; top?: number; skipToken?: string; cursor?: string; signal?: AbortSignal }): Promise<MicrosoftCollectionResult<MicrosoftDriveItemInfo>> {
    const queryText = (options.query ?? '').trim()
    if (!queryText && !options.cursor) throw new Error('Microsoft Graph Drive search query is required.')
    const pathPattern = new RegExp('^/me/drive/root/search$')
    const path = options.cursor ? this.acceptCursor(options.cursor, pathPattern) : '/me/drive/root/search' + query({ q: queryText, '$top': clamp(options.top, 1, 100, 25), '$skiptoken': options.skipToken, '$select': graphSelect('id', 'name', 'size', 'webUrl', 'createdDateTime', 'lastModifiedDateTime', 'file', 'folder', 'parentReference') })
    const data = await this.request<unknown>(path, {}, options.signal)
    const raw = record(data)
    return { items: array(raw.value).map(mapDriveItem), nextCursor: this.normalizeNextLink(raw['@odata.nextLink'], pathPattern) }
  }

  async getDriveItem(itemId: string, signal?: AbortSignal): Promise<MicrosoftDriveItemInfo> {
    if (!(itemId ?? '').trim()) throw new Error('Drive item id is required.')
    const path = '/me/drive/items/' + encodeURIComponent(itemId) + query({ '$select': graphSelect('id', 'name', 'size', 'webUrl', 'createdDateTime', 'lastModifiedDateTime', 'file', 'folder', 'parentReference') })
    return mapDriveItem(await this.request<unknown>(path, {}, signal))
  }

  private async search(entityType: 'message' | 'chatMessage', queryText: string, from: number, size: number, signal?: AbortSignal): Promise<{ hits: unknown[]; total: number; more: boolean }> {
    const body = { requests: [{ entityTypes: [entityType], query: { queryString: queryText }, from, size, ...(entityType === 'chatMessage' ? { enableTopResults: true } : {}) }] }
    return searchHits(await this.request<unknown>('/search/query', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) }, signal))
  }

  private async request<T>(pathOrUrl: string, init: RequestInit = {}, signal?: AbortSignal): Promise<T> {
    const url = this.buildUrl(pathOrUrl)
    const controller = new AbortController()
    const onAbort = () => controller.abort(signal?.reason)
    if (signal) { if (signal.aborted) controller.abort(signal.reason); else signal.addEventListener('abort', onAbort, { once: true }) }
    const timer = setTimeout(() => controller.abort(new Error('Microsoft Graph request timed out after ' + this.timeoutMs + 'ms')), this.timeoutMs)
    try {
      const headers = new Headers(init.headers)
      headers.set('accept', 'application/json')
      headers.set('authorization', 'Bearer ' + this.accessToken)
      if (init.body !== undefined) headers.set('content-type', 'application/json')
      const response = await this.fetchImpl(url, { ...init, headers, signal: controller.signal })
      let body: unknown
      try { body = await response.json() } catch { body = undefined }
      if (!response.ok) {
        const details = errorDetails(body)
        const requestId = details.requestId || response.headers.get('request-id') || response.headers.get('client-request-id') || undefined
        const retry = Number(response.headers.get('retry-after'))
        const safeCode = /^[A-Za-z0-9._-]{1,80}$/.test(details.code) ? details.code : 'GRAPH_ERROR'
        const safeRequestId = requestId ? requestId.slice(0, 64).replace(/[^A-Za-z0-9._:-]/g, '') : ''
        const suffix = safeRequestId ? ', requestId=' + safeRequestId : ''
        throw new MicrosoftGraphError('Microsoft Graph request failed (HTTP ' + response.status + ', code=' + safeCode + suffix + ').', response.status, safeCode, safeRequestId || undefined, Number.isFinite(retry) && retry > 0 ? retry : undefined)
      }
      return body as T
    } finally {
      clearTimeout(timer)
      if (signal) signal.removeEventListener('abort', onAbort)
    }
  }

  private buildUrl(pathOrUrl: string): string {
    if (!pathOrUrl.startsWith('/') || pathOrUrl.startsWith('//') || /^https?:\/\//i.test(pathOrUrl)) throw new MicrosoftGraphError('Microsoft Graph cursor must resolve to a relative v1.0 path.', 400)
    return this.baseUrl + '/' + pathOrUrl.replace(/^\/+/, '')
  }

  private acceptCursor(value: string, pattern: RegExp): string {
    const token = value.trim()
    if (!token || token.length > 128) throw new MicrosoftGraphError('Microsoft Graph cursor is invalid or expired.', 400)
    const path = this.cursors.get(token)
    if (!path || !pattern.test(new URL(path, this.baseOrigin).pathname)) throw new MicrosoftGraphError('Microsoft Graph cursor is invalid or expired.', 400)
    return path
  }

  private normalizeNextLink(value: unknown, pattern: RegExp): string {
    const raw = stringValue(value).trim()
    if (!raw) return ''
    let parsed: URL
    try { parsed = /^https?:\/\//i.test(raw) ? new URL(raw) : new URL(raw, this.baseOrigin) } catch { throw new MicrosoftGraphError('Microsoft Graph returned an invalid nextLink.', 502) }
    if (parsed.origin !== this.baseOrigin) throw new MicrosoftGraphError('Microsoft Graph returned a nextLink for another host.', 502)
    let path = parsed.pathname
    if (path === this.basePath) path = '/'
    else if (path.startsWith(this.basePath + '/')) path = path.slice(this.basePath.length)
    if (!pattern.test(path)) throw new MicrosoftGraphError('Microsoft Graph returned a nextLink for another resource.', 502)
    const relative = path + parsed.search
    const token = 'm365_' + randomUUID().replace(/-/g, '')
    this.cursors.set(token, relative)
    if (this.cursors.size > 1000) this.cursors.delete(this.cursors.keys().next().value as string)
    return token
  }
}
