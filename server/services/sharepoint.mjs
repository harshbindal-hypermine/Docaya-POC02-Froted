import { ClientSecretCredential, ManagedIdentityCredential } from '@azure/identity'
import { basename, extname } from 'node:path'
import { HttpError } from '../middleware/errors.mjs'

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0'
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default'
const SMALL_UPLOAD_LIMIT = 4 * 1024 * 1024
const itemIdPattern = /^[A-Za-z0-9._!:-]{1,300}$/

export class SharePointGraphError extends Error {
  constructor(operation, status, code, message, requestId) {
    super(message || 'Microsoft Graph request failed.')
    this.operation = operation
    this.status = status
    this.code = code || 'graphError'
    this.requestId = requestId || ''
  }
}

export class SharePointService {
  constructor(config, options = {}) {
    this.config = config
    this.fetch = options.fetchImpl || globalThis.fetch
    this.credential = options.credential || createCredential(config)
    this.cache = null
  }

  isConfigured() {
    return Boolean(
      this.config.sharePointHostname &&
        this.config.sharePointSitePath &&
        this.config.sharePointDriveName &&
        ((this.config.sharePointUseManagedIdentity && this.credential) ||
          (this.config.sharePointTenantId && this.config.sharePointClientId && this.config.sharePointClientSecret)),
    )
  }

  configuration() {
    return {
      configured: this.isConfigured(),
      hostname: this.config.sharePointHostname || '',
      sitePath: this.config.sharePointSitePath || '',
      driveName: this.config.sharePointDriveName || '',
      tenantConfigured: Boolean(this.config.sharePointTenantId),
      clientConfigured: Boolean(this.config.sharePointClientId),
      secretConfigured: Boolean(this.config.sharePointClientSecret || this.config.sharePointUseManagedIdentity),
      managedIdentity: Boolean(this.config.sharePointUseManagedIdentity),
    }
  }

  async healthCheck() {
    if (!this.isConfigured()) return { status: 'disabled', configured: false }
    try {
      const resolved = await this.resolve()
      return {
        status: 'ok',
        configured: true,
        site: {
          hostname: this.config.sharePointHostname,
          path: this.config.sharePointSitePath,
          displayName: resolved.site.displayName || resolved.site.name || null,
          webUrl: resolved.site.webUrl || null,
        },
        drive: {
          name: resolved.drive.name,
          webUrl: resolved.drive.webUrl || null,
        },
      }
    } catch (error) {
      return { status: 'error', configured: true, error: publicGraphMessage(error) }
    }
  }

  async getConnectionStatus() {
    const health = await this.healthCheck()
    return {
      ...this.configuration(),
      connected: health.status === 'ok',
      status: health.status === 'ok' ? 'connected' : this.isConfigured() ? 'connection_failed' : 'not_configured',
      site: health.site || null,
      drive: health.drive || null,
    }
  }

  async list(folderId = '') {
    const { drive } = await this.resolve()
    const path = folderId
      ? `/drives/${encodeURIComponent(drive.id)}/items/${encodeGraphPathSegment(validateItemId(folderId))}/children`
      : `/drives/${encodeURIComponent(drive.id)}/root/children`
    const payload = await this.graphPagedJson('list-folder', path)
    return {
      items: (payload.value || []).map(normalizeItem),
      folderId: folderId || null,
      drive: publicDrive(drive),
      site: {
        hostname: this.config.sharePointHostname,
        path: this.config.sharePointSitePath,
      },
    }
  }

  async listFolders() {
    const { drive } = await this.resolve()
    const payload = await this.graphPagedJson(
      'list-sharepoint-folders',
      `/drives/${encodeURIComponent(drive.id)}/root/children?$select=id,name,folder,webUrl,lastModifiedDateTime,parentReference`,
    )
    const roots = (payload.value || []).map(normalizeItem).filter((item) => item.type === 'folder')
    const folders = []
    for (const root of roots) folders.push(...(await this.collectFolderTree(drive.id, root, 0)))
    return {
      items: folders,
      drive: { name: drive.name, webUrl: drive.webUrl || null },
      site: {
        hostname: this.config.sharePointHostname,
        path: this.config.sharePointSitePath,
      },
    }
  }

  async debug() {
    const resolved = await this.resolve()
    const root = await this.list('')
    return {
      site: publicSite(resolved.site, this.config),
      drives: resolved.drives.map(publicDrive),
      drive: publicDrive(resolved.drive),
      rootItems: root.items,
    }
  }

  async getPreview(itemId) {
    const item = await this.getItem(itemId)
    return {
      id: item.id,
      name: item.name,
      type: item.type,
      webUrl: item.webUrl || null,
      previewAvailable: Boolean(item.webUrl),
    }
  }

  async download(itemId) {
    const { drive } = await this.resolve()
    const item = await this.getItem(itemId)
    if (item.type !== 'file') throw new HttpError(422, 'SHAREPOINT_NOT_FILE', 'Downloads are available for files only.')
    const response = await this.graphFetch(
      'download-file',
      `/drives/${encodeURIComponent(drive.id)}/items/${encodeGraphPathSegment(validateItemId(itemId))}/content`,
    )
    const contentType = response.headers.get('content-type') || item.mimeType || 'application/octet-stream'
    const contentLength = response.headers.get('content-length')
    return { item, response, contentType, contentLength }
  }

  async upload({ fileName, mimeType, sizeBytes, stream, folderId = '' }) {
    const safeFileName = validateFileName(fileName)
    validateUpload(this.config, safeFileName, mimeType, sizeBytes)
    const { drive } = await this.resolve()
    const parent = folderId ? { id: validateItemId(folderId) } : await this.ensureUploadFolder()
    const parentPath = `/drives/${encodeURIComponent(drive.id)}/items/${encodeGraphPathSegment(parent.id)}`
    const body = await streamToBuffer(stream, this.config.uploadMaxBytes)
    if (body.length !== sizeBytes) throw new HttpError(400, 'UPLOAD_SIZE_MISMATCH', 'The uploaded file size does not match the declared size.')
    if (sizeBytes <= SMALL_UPLOAD_LIMIT) {
      const item = await this.graphJson(
        'upload-small-file',
        `${parentPath}:/${encodeGraphPathSegment(safeFileName)}:/content`,
        { method: 'PUT', headers: { 'content-type': mimeType }, body },
      )
      return normalizeItem(item)
    }
    const session = await this.graphJson(
      'create-upload-session',
      `${parentPath}:/${encodeGraphPathSegment(safeFileName)}:/createUploadSession`,
      { method: 'POST', body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'rename' } }) },
    )
    if (!session.uploadUrl?.startsWith('https://'))
      throw new HttpError(502, 'SHAREPOINT_UPLOAD_FAILED', 'Upload failed. Please try again.')
    const uploaded = await this.uploadLargeFile(session.uploadUrl, body, mimeType)
    return normalizeItem(uploaded)
  }

  async uploadLargeFile(uploadUrl, body, mimeType) {
    const chunkSize = 8 * 1024 * 1024
    let lastPayload = null
    for (let start = 0; start < body.length; start += chunkSize) {
      const end = Math.min(start + chunkSize, body.length) - 1
      const chunk = body.subarray(start, end + 1)
      const response = await this.fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'content-length': String(chunk.length),
          'content-range': `bytes ${start}-${end}/${body.length}`,
          'content-type': mimeType,
        },
        body: chunk,
      })
      lastPayload = await response.json().catch(() => ({}))
      if (!response.ok) throw graphErrorFromResponse('upload-large-file', response, lastPayload)
    }
    return lastPayload
  }

  async getItem(itemId) {
    const { drive } = await this.resolve()
    const payload = await this.graphJson(
      'get-item',
      `/drives/${encodeURIComponent(drive.id)}/items/${encodeGraphPathSegment(validateItemId(itemId))}`,
    )
    return normalizeItem(payload)
  }

  async collectFolderTree(driveId, folder, depth) {
    const folders = [folder]
    if (depth >= 5) return folders
    const children = await this.listFolderChildren(driveId, folder.id, 'list-sharepoint-folder-children')
    for (const child of children) folders.push(...(await this.collectFolderTree(driveId, child, depth + 1)))
    return folders
  }

  async listFolderChildren(driveId, folderId, operation) {
    const payload = await this.graphPagedJson(
      operation,
      `/drives/${encodeURIComponent(driveId)}/items/${encodeGraphPathSegment(folderId)}/children?$select=id,name,folder,webUrl,lastModifiedDateTime,parentReference`,
    )
    return (payload.value || []).map(normalizeItem).filter((item) => item.type === 'folder')
  }

  async ensureUploadFolder() {
    const folderName = validateFileName(this.config.sharePointUploadFolderName || 'Docaya Uploads')
    const { drive } = await this.resolve()
    try {
      return await this.graphJson(
        'resolve-upload-folder',
        `/drives/${encodeURIComponent(drive.id)}/root:/${encodeGraphPathSegment(folderName)}`,
      )
    } catch (error) {
      if (!(error instanceof SharePointGraphError) || error.status !== 404) throw error
    }
    return this.graphJson('create-upload-folder', `/drives/${encodeURIComponent(drive.id)}/root/children`, {
      method: 'POST',
      body: JSON.stringify({
        name: folderName,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'fail',
      }),
    })
  }

  async resolve() {
    if (this.cache) return this.cache
    if (!this.isConfigured()) throw new HttpError(503, 'SHAREPOINT_NOT_CONFIGURED', 'SharePoint is not configured.')
    const sitePath = encodeSitePath(this.config.sharePointSitePath)
    const site = await this.graphJson(
      'resolve-site',
      `/sites/${encodeURIComponent(this.config.sharePointHostname)}:${sitePath}`,
    )
    const drives = await this.graphJson('resolve-drive', `/sites/${encodeURIComponent(site.id)}/drives`)
    const drive = selectDrive(drives.value || [], this.config)
    if (!drive) throw new HttpError(404, 'SHAREPOINT_DRIVE_NOT_FOUND', 'The configured SharePoint document library was not found.')
    this.logResolvedSharePoint(site, drives.value || [], drive)
    this.cache = { site, drives: drives.value || [], drive }
    return this.cache
  }

  logResolvedSharePoint(site, drives, drive) {
    console.info(
      JSON.stringify({
        level: 'info',
        event: 'sharepoint.resolved',
        site: publicSite(site, this.config),
        configuredDriveName: this.config.sharePointDriveName || null,
        configuredDriveId: this.config.sharePointDriveId || null,
        selectedDrive: publicDrive(drive),
        drives: drives.map(publicDrive),
      }),
    )
  }

  async graphJson(operation, path, init = {}) {
    const response = await this.graphFetch(operation, path, init)
    return response.status === 204 ? {} : await response.json().catch(() => ({}))
  }

  async graphPagedJson(operation, path) {
    const items = []
    let nextPath = path
    for (let page = 0; nextPath && page < 20; page += 1) {
      const payload = nextPath.startsWith('https://graph.microsoft.com/v1.0/')
        ? await this.graphAbsoluteJson(operation, nextPath)
        : await this.graphJson(operation, nextPath)
      items.push(...(payload.value || []))
      nextPath = payload['@odata.nextLink'] || ''
    }
    return { value: items }
  }

  async graphAbsoluteJson(operation, url) {
    const token = await this.getToken()
    const parsed = new URL(url)
    if (parsed.origin !== 'https://graph.microsoft.com' || !parsed.pathname.startsWith('/v1.0/'))
      throw new HttpError(502, 'SHAREPOINT_PAGING_FAILED', 'Documents could not be loaded. Please try again.')
    const response = await this.fetch(parsed.toString(), {
      headers: { accept: 'application/json', authorization: `Bearer ${token}` },
    })
    if (!response.ok) {
      const payload = await response.json().catch(() => ({}))
      throw graphErrorFromResponse(operation, response, payload)
    }
    return response.json().catch(() => ({}))
  }

  async graphFetch(operation, path, init = {}) {
    const token = await this.getToken()
    const response = await this.fetch(`${GRAPH_ROOT}${path}`, {
      ...init,
      headers: {
        accept: 'application/json',
        ...(init.body && typeof init.body === 'string' ? { 'content-type': 'application/json' } : {}),
        ...init.headers,
        authorization: `Bearer ${token}`,
      },
    })
    if (!response.ok) {
      const payload = await response.json().catch(() => ({}))
      throw graphErrorFromResponse(operation, response, payload)
    }
    return response
  }

  async getToken() {
    const token = await this.credential.getToken(GRAPH_SCOPE)
    if (!token?.token) throw new HttpError(503, 'SHAREPOINT_AUTH_FAILED', 'SharePoint authentication failed.')
    return token.token
  }
}

function createCredential(config) {
  if (config.sharePointCredential) return config.sharePointCredential
  if (config.sharePointUseManagedIdentity) return new ManagedIdentityCredential()
  if (!config.sharePointTenantId || !config.sharePointClientId || !config.sharePointClientSecret) return null
  return new ClientSecretCredential(
    config.sharePointTenantId,
    config.sharePointClientId,
    config.sharePointClientSecret,
  )
}

export function normalizeItem(item) {
  const type = item.folder ? 'folder' : 'file'
  const parentPath = item.parentReference?.path || ''
  const pathPrefix = parentPath.includes(':') ? parentPath.slice(parentPath.indexOf(':') + 1) : parentPath
  return {
    id: String(item.id || ''),
    name: String(item.name || ''),
    type,
    size: type === 'folder' ? null : Number(item.size || 0),
    path: `${pathPrefix}/${String(item.name || '')}`.replaceAll('//', '/'),
    webUrl: item.webUrl || null,
    lastModifiedDateTime: item.lastModifiedDateTime || null,
    mimeType: item.file?.mimeType || null,
    parentReference: item.parentReference
      ? {
          id: item.parentReference.id || null,
          path: item.parentReference.path || null,
        }
      : null,
  }
}

function selectDrive(drives, config) {
  if (config.sharePointDriveId) {
    const byId = drives.find((candidate) => candidate.id === config.sharePointDriveId)
    if (byId) return byId
  }
  const configuredName = String(config.sharePointDriveName || '').trim().toLowerCase()
  if (configuredName) {
    const exact = drives.find((candidate) => String(candidate.name || '').toLowerCase() === configuredName)
    if (exact) return exact
  }
  const documentLibrary = drives.find((candidate) => String(candidate.driveType || '').toLowerCase() === 'documentlibrary')
  return documentLibrary || drives[0] || null
}

function publicSite(site, config) {
  return {
    id: site.id,
    displayName: site.displayName || site.name || null,
    webUrl: site.webUrl || null,
    hostname: config.sharePointHostname,
    path: config.sharePointSitePath,
  }
}

function publicDrive(drive) {
  return {
    id: drive.id,
    name: drive.name,
    webUrl: drive.webUrl || null,
    driveType: drive.driveType || null,
  }
}

function validateItemId(value) {
  const itemId = String(value || '').trim()
  if (!itemIdPattern.test(itemId)) throw new HttpError(422, 'INVALID_SHAREPOINT_ITEM_ID', 'The SharePoint item ID is invalid.')
  return itemId
}

function validateFileName(value) {
  const fileName = basename(String(value || '')).normalize('NFKC')
  const hasControlCharacter = [...fileName].some((character) => character.charCodeAt(0) < 32)
  if (!fileName || fileName.length > 240 || /[<>:"/\\|?*]/.test(fileName) || hasControlCharacter)
    throw new HttpError(422, 'INVALID_FILENAME', 'The file name is not allowed.')
  return fileName
}

function validateUpload(config, fileName, mimeType, sizeBytes) {
  const extension = extname(fileName).slice(1).toLowerCase()
  const normalizedMime = String(mimeType || '').toLowerCase()
  if (!config.uploadExtensions.includes(extension))
    throw new HttpError(415, 'EXTENSION_NOT_ALLOWED', 'The file extension is not allowed.')
  if (!config.uploadMimeTypes.includes(normalizedMime))
    throw new HttpError(415, 'MIME_NOT_ALLOWED', 'The file type is not allowed.')
  if (!Number.isSafeInteger(sizeBytes) || sizeBytes <= 0 || sizeBytes > config.uploadMaxBytes)
    throw new HttpError(413, 'FILE_SIZE_INVALID', `The file must be between 1 and ${config.uploadMaxBytes} bytes.`)
}

async function streamToBuffer(stream, limit) {
  const chunks = []
  let size = 0
  for await (const chunk of stream) {
    size += chunk.length
    if (size > limit) throw new HttpError(413, 'PAYLOAD_TOO_LARGE', 'The request body exceeds the allowed size.')
    chunks.push(chunk)
  }
  return Buffer.concat(chunks)
}

function graphErrorFromResponse(operation, response, payload) {
  const graphError = payload?.error || {}
  const requestId =
    graphError.innerError?.['request-id'] ||
    graphError.innerError?.requestId ||
    response.headers.get('request-id') ||
    response.headers.get('client-request-id') ||
    ''
  return new SharePointGraphError(
    operation,
    response.status,
    graphError.code || 'graphError',
    graphError.message || `Microsoft Graph returned HTTP ${response.status}.`,
    requestId,
  )
}

function publicGraphMessage(error) {
  if (error instanceof HttpError) return error.message
  if (error instanceof SharePointGraphError) return 'SharePoint connection failed.'
  return 'SharePoint is unavailable.'
}

export function logSharePointError(error, requestId, fallbackOperation) {
  if (error instanceof SharePointGraphError) {
    console.error(
      JSON.stringify({
        level: 'error',
        event: 'sharepoint.graph_error',
        requestId,
        operation: error.operation || fallbackOperation,
        status: error.status,
        code: error.code,
        graphRequestId: error.requestId,
      }),
    )
    return
  }
  if (!(error instanceof HttpError)) {
    console.error(
      JSON.stringify({
        level: 'error',
        event: 'sharepoint.error',
        requestId,
        operation: fallbackOperation,
        message: error instanceof Error ? error.message : String(error),
      }),
    )
  }
}

function encodeSitePath(sitePath) {
  return String(sitePath || '')
    .split('/')
    .map((part, index) => (index === 0 ? '' : encodeURIComponent(part)))
    .join('/')
}

function encodeGraphPathSegment(segment) {
  return encodeURIComponent(segment).replaceAll('%2F', '')
}
