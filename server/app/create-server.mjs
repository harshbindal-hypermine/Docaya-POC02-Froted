import crypto from 'node:crypto'
import { createServer } from 'node:http'
import { createReadStream, existsSync, statSync } from 'node:fs'
import { extname, resolve, sep } from 'node:path'
import { Readable } from 'node:stream'
import { createAuthenticator, authorize, issueDevelopmentToken } from '../middleware/auth.mjs'
import { assert, HttpError } from '../middleware/errors.mjs'
import { prepareRequest, readJson, sendError, sendJson } from '../middleware/http.mjs'
import { audit } from '../services/audit.mjs'
import { logSharePointError } from '../services/sharepoint.mjs'
import { UploadService } from '../services/uploads.mjs'

const documentStatuses = ['Draft', 'Under Review', 'Under Approval', 'Published', 'Superseded', 'Archived']
const classifications = ['Public', 'Internal', 'Confidential', 'Restricted']
const workflowDecisions = ['Approved', 'Changes requested', 'Rejected']
const assignableRoles = [
  'org-admin',
  'workspace-admin',
  'content-manager',
  'contributor',
  'viewer',
  'guest',
  'auditor',
]
const userStatuses = ['invited', 'active', 'suspended', 'deprovisioned']
const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/
const contentTypes = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
}
const uuid = (prefix) => `${prefix}-${crypto.randomUUID()}`
const cleanText = (value, name, { min = 1, max = 240 } = {}) => {
  const text = String(value ?? '').trim()
  assert(
    text.length >= min && text.length <= max,
    422,
    'VALIDATION_ERROR',
    `${name} must contain ${min}-${max} characters.`,
  )
  return text
}
const parseCursor = (value) => {
  if (!value) return null
  try {
    const decoded = JSON.parse(Buffer.from(value, 'base64url').toString('utf8'))
    assert(
      decoded.id && (decoded.updatedAt || decoded.createdAt),
      400,
      'INVALID_CURSOR',
      'The cursor is invalid.',
    )
    return decoded
  } catch (error) {
    if (error instanceof HttpError) throw error
    throw new HttpError(400, 'INVALID_CURSOR', 'The cursor is invalid.')
  }
}

export function createDocayaServer({
  config,
  repository,
  search = null,
  cache = null,
  events = null,
  assistant = null,
  sharePoint = null,
}) {
  const authenticate = createAuthenticator(config)
  const uploads = new UploadService(config, repository)
  const staticRoot = resolve('dist')
  const checkDependency = async (dependency) => {
    if (!dependency) return { status: 'disabled' }
    let timer
    try {
      const result = await Promise.race([
        dependency.healthCheck(),
        new Promise((_, reject) => {
          timer = setTimeout(() => reject(new Error('Readiness check timed out.')), config.readinessTimeoutMs)
        }),
      ])
      return result && typeof result === 'object' ? result : { status: 'ok' }
    } catch (error) {
      return { status: 'error', error: error instanceof Error ? error.message : String(error) }
    } finally {
      clearTimeout(timer)
    }
  }

  const handler = async (req, res) => {
    try {
      prepareRequest(req, res, config)
      if (req.method === 'OPTIONS') return sendJson(res, 204, {})
      const url = new URL(req.url, 'http://docaya.local')
      const route = url.pathname
      if (req.method === 'GET' && route === '/health/live')
        return sendJson(res, 200, { status: 'ok', timestamp: new Date().toISOString() })
      if (req.method === 'GET' && route === '/health/ready') {
        const checks = Object.fromEntries(
          await Promise.all([
            checkDependency(repository).then((result) => ['database', result]),
            checkDependency(search).then((result) => ['search', result]),
            checkDependency(cache).then((result) => ['cache', result]),
            checkDependency(events).then((result) => ['events', result]),
            checkDependency(sharePoint).then((result) => ['sharePoint', result]),
          ]),
        )
        const ready = Object.values(checks).every(
          (check) => check.status === 'ok' || check.status === 'disabled',
        )
        return sendJson(res, ready ? 200 : 503, {
          status: ready ? 'ready' : 'not_ready',
          timestamp: new Date().toISOString(),
          checks,
        })
      }
      if (req.method === 'POST' && route === '/api/v1/auth/development-token')
        return sendJson(res, 200, {
          accessToken: await issueDevelopmentToken(config),
          tokenType: 'Bearer',
          expiresIn: 3600,
        })

      if (route.startsWith('/api/')) {
        if (!route.startsWith('/api/v1/'))
          throw new HttpError(404, 'API_VERSION_REQUIRED', 'Use the /api/v1 API contract.')
        const actor = await authenticate(req)
        req.actor = actor

        if (route === '/api/v1/auth/me' && req.method === 'GET') {
          await audit(repository, req, actor, 'identity.read', 'user', actor.id)
          return sendJson(res, 200, { user: actor })
        }

        if (route === '/api/v1/documents' && req.method === 'GET') {
          authorize(actor, 'document:read')
          const limit = Math.min(Math.max(Number(url.searchParams.get('limit') || 25), 1), 100)
          const result = await repository.listDocuments(actor, {
            q: url.searchParams.get('q') || '',
            status: url.searchParams.get('status') || undefined,
            cursor: parseCursor(url.searchParams.get('cursor')),
            limit,
          })
          await audit(repository, req, actor, 'document.list', 'document', null, 'success', {
            count: result.items.length,
          })
          return sendJson(res, 200, {
            documents: result.items,
            page: { nextCursor: result.nextCursor, limit },
          })
        }
        if (route === '/api/v1/documents' && req.method === 'POST') {
          authorize(actor, 'document:write')
          assert(
            req.headers['idempotency-key'],
            400,
            'IDEMPOTENCY_KEY_REQUIRED',
            'An Idempotency-Key header is required.',
          )
          const body = await readJson(req, config.bodyLimit)
          const input = {
            id: uuid('doc'),
            title: cleanText(body.title, 'title'),
            number: cleanText(body.number, 'number', { max: 80 }),
            type: cleanText(body.type, 'type', { max: 80 }),
            status: body.status || 'Draft',
            classification: body.classification || 'Internal',
            department: cleanText(body.department, 'department', { max: 120 }),
            library: cleanText(body.library, 'library', { max: 120 }),
            language: cleanText(body.language || 'English', 'language', { max: 80 }),
            sourcePath: body.sourcePath ? cleanText(body.sourcePath, 'sourcePath', { max: 1000 }) : null,
            summary: cleanText(body.summary || '', 'summary', { min: 0, max: 4000 }),
            retention: cleanText(body.retention || '7 years', 'retention', { max: 80 }),
            nextReview: cleanText(body.nextReview, 'nextReview', { max: 10 }),
            reviewer: cleanText(body.reviewer, 'reviewer', { max: 160 }),
            approver: cleanText(body.approver, 'approver', { max: 160 }),
            workflowStep: body.workflowStep || 1,
          }
          assert(
            body.sharePointFolderId,
            422,
            'SHAREPOINT_FOLDER_REQUIRED',
            'Selected SharePoint folder is not available for this department.',
          )
          const folder = await repository.getDepartmentSharePointFolder(
            actor,
            input.department,
            String(body.sharePointFolderId),
          )
          assert(
            folder,
            422,
            'SHAREPOINT_FOLDER_NOT_ALLOWED',
            'Selected SharePoint folder is not available for this department.',
          )
          input.sharePointFolderId = folder.sharePointFolderId
          input.sharePointFolderName = folder.name
          input.sharePointFolderPath = folder.path
          assert(
            documentStatuses.includes(input.status),
            422,
            'VALIDATION_ERROR',
            'Document status is invalid.',
          )
          assert(
            classifications.includes(input.classification),
            422,
            'VALIDATION_ERROR',
            'Classification is invalid.',
          )
          const document = await repository.createDocument(actor, input)
          if (search) await search.indexDocument(document)
          if (events)
            await events.publish({
              type: 'document.created',
              documentId: document.id,
              tenantId: actor.tenantId,
            })
          await audit(repository, req, actor, 'document.register', 'document', document.id)
          return sendJson(res, 201, { document })
        }
        if (route === '/api/v1/trash' && req.method === 'GET') {
          authorize(actor, 'document:read')
          const result = await repository.listDocuments(actor, { deleted: true, limit: 100 })
          return sendJson(res, 200, {
            documents: result.items,
            page: { nextCursor: result.nextCursor, limit: 100 },
          })
        }

        const documentMatch = route.match(/^\/api\/v1\/documents\/([^/]+)$/)
        if (documentMatch && req.method === 'GET') {
          authorize(actor, 'document:read')
          const document = await repository.getDocument(actor, decodeURIComponent(documentMatch[1]))
          assert(document, 404, 'NOT_FOUND', 'Document not found.')
          authorize(actor, 'document:read', document)
          await audit(repository, req, actor, 'document.view', 'document', document.id)
          return sendJson(res, 200, { document })
        }
        if (documentMatch && req.method === 'PATCH') {
          authorize(actor, 'document:write')
          const body = await readJson(req, config.bodyLimit)
          if (body.status !== undefined)
            assert(
              documentStatuses.includes(body.status),
              422,
              'VALIDATION_ERROR',
              'Document status is invalid.',
            )
          if (body.classification !== undefined)
            assert(
              classifications.includes(body.classification),
              422,
              'VALIDATION_ERROR',
              'Classification is invalid.',
            )
          if (body.title !== undefined) body.title = cleanText(body.title, 'title')
          if (body.summary !== undefined)
            body.summary = cleanText(body.summary, 'summary', { min: 0, max: 4000 })
          const document = await repository.updateDocument(actor, decodeURIComponent(documentMatch[1]), body)
          assert(document, 404, 'NOT_FOUND', 'Document not found or no editable fields supplied.')
          await audit(repository, req, actor, 'document.update', 'document', document.id, 'success', {
            fields: Object.keys(body),
          })
          return sendJson(res, 200, { document })
        }
        if (documentMatch && req.method === 'DELETE') {
          authorize(actor, 'document:delete')
          const id = decodeURIComponent(documentMatch[1])
          const result = await repository.softDeleteDocument(actor, id)
          if (result.legalHold)
            throw new HttpError(409, 'LEGAL_HOLD', 'A document on legal hold cannot be deleted.')
          assert(result.changed, 404, 'NOT_FOUND', 'Document not found.')
          await audit(repository, req, actor, 'document.delete', 'document', id)
          return sendJson(res, 200, { success: true })
        }
        const restoreMatch = route.match(/^\/api\/v1\/trash\/([^/]+)\/restore$/)
        if (restoreMatch && req.method === 'POST') {
          authorize(actor, 'document:delete')
          const id = decodeURIComponent(restoreMatch[1])
          assert(
            await repository.restoreDocument(actor, id),
            404,
            'NOT_FOUND',
            'Document not found in trash.',
          )
          await audit(repository, req, actor, 'document.restore', 'document', id)
          return sendJson(res, 200, { success: true })
        }

        if (route === '/api/v1/uploads' && req.method === 'POST') {
          authorize(actor, 'document:write')
          const body = await readJson(req, config.bodyLimit)
          const upload = await uploads.create(actor, body, String(req.headers['idempotency-key'] || ''))
          await audit(repository, req, actor, 'upload.create', 'upload', upload.id)
          return sendJson(res, 201, {
            upload: publicUpload(upload),
            links: {
              chunk: `/api/v1/uploads/${upload.id}/chunks/{number}`,
              complete: `/api/v1/uploads/${upload.id}/complete`,
              cancel: `/api/v1/uploads/${upload.id}`,
            },
          })
        }
        const chunkMatch = route.match(/^\/api\/v1\/uploads\/([^/]+)\/chunks\/(\d+)$/)
        if (chunkMatch && req.method === 'PUT') {
          authorize(actor, 'document:write')
          const upload = await uploads.writeChunk(
            actor,
            decodeURIComponent(chunkMatch[1]),
            chunkMatch[2],
            req,
          )
          return sendJson(res, 200, { upload: publicUpload(upload) })
        }
        const completeMatch = route.match(/^\/api\/v1\/uploads\/([^/]+)\/complete$/)
        if (completeMatch && req.method === 'POST') {
          authorize(actor, 'document:write')
          const body = await readJson(req, config.bodyLimit)
          const result = await uploads.complete(
            actor,
            decodeURIComponent(completeMatch[1]),
            body.metadata || {},
          )
          await audit(repository, req, actor, 'document.create', 'document', result.document.id, 'success', {
            hash: result.document.contentHash,
          })
          return sendJson(res, 201, { upload: publicUpload(result.upload), document: result.document })
        }
        const uploadMatch = route.match(/^\/api\/v1\/uploads\/([^/]+)$/)
        if (uploadMatch && req.method === 'DELETE') {
          authorize(actor, 'document:write')
          const upload = await uploads.cancel(actor, decodeURIComponent(uploadMatch[1]))
          await audit(repository, req, actor, 'upload.cancel', 'upload', upload.id)
          return sendJson(res, 200, { upload: publicUpload(upload) })
        }

        if (route === '/api/v1/search' && req.method === 'POST') {
          authorize(actor, 'document:read')
          const body = await readJson(req, config.bodyLimit)
          const query = cleanText(body.query, 'query', { min: 1, max: 500 })
          if (body.filters?.status !== undefined)
            assert(
              documentStatuses.includes(body.filters.status),
              422,
              'VALIDATION_ERROR',
              'Document status is invalid.',
            )
          const requestedLimit = Number(body.limit || 25)
          assert(
            Number.isInteger(requestedLimit) && requestedLimit > 0,
            422,
            'VALIDATION_ERROR',
            'limit must be a positive integer.',
          )
          const searchLimit = Math.min(requestedLimit, 100)
          const result = await repository.listDocuments(actor, {
            q: query,
            status: body.filters?.status,
            limit: searchLimit,
          })
          if (search) {
            const key = `search:${actor.tenantId}:${query}:${body.filters?.status || ''}:${searchLimit}`
            const cached = cache ? await cache.get(key) : null
            const indexed = cached
              ? JSON.parse(cached)
              : await search.search(query, {
                  top: searchLimit,
                  filter: `tenantId eq '${String(actor.tenantId).replaceAll("'", "''")}'`,
                })
            if (!cached && cache) await cache.set(key, JSON.stringify(indexed), 30)
            if (indexed) {
              await audit(repository, req, actor, 'search.execute', 'search', null, 'success', {
                queryLength: query.length,
                count: indexed.items.length,
              })
              return sendJson(res, 200, { results: indexed.items, page: { nextCursor: null } })
            }
          }
          await audit(repository, req, actor, 'search.execute', 'search', null, 'success', {
            queryLength: query.length,
            count: result.items.length,
          })
          return sendJson(res, 200, { results: result.items, page: { nextCursor: result.nextCursor } })
        }

        if (route === '/api/v1/assistant/ask' && req.method === 'POST') {
          authorize(actor, 'document:read')
          assert(
            assistant,
            503,
            'ASSISTANT_UNAVAILABLE',
            'The AI assistant is not configured for this environment.',
          )
          const body = await readJson(req, config.bodyLimit)
          const question = cleanText(body.question, 'question', { min: 3, max: 500 })
          const references = await findRelevantDocuments(repository, actor, question)
          const contextText = references.length
            ? references
                .map(
                  (document, index) =>
                    `${index + 1}. ${document.title} (${document.number}) — status: ${document.status}, department: ${document.department}, next review: ${document.nextReview}`,
                )
                .join('\n')
            : 'No controlled documents matched this question.'
          const messages = [
            {
              role: 'system',
              content:
                'You are Ask Docaya, an assistant for a document-governance platform. Answer only using the CONTROLLED DOCUMENTS context below. If the context does not answer the question, say you could not find a matching controlled document. Be concise and cite document numbers when relevant.',
            },
            { role: 'user', content: `CONTROLLED DOCUMENTS:\n${contextText}\n\nQUESTION: ${question}` },
          ]
          let answer
          try {
            answer = await assistant.chat(messages)
          } catch (error) {
            console.error(
              JSON.stringify({
                level: 'error',
                requestId: req.requestId,
                event: 'assistant.upstream_error',
                message: error instanceof Error ? error.message : String(error),
              }),
            )
            throw new HttpError(
              502,
              'ASSISTANT_UPSTREAM_ERROR',
              'The AI assistant could not generate a response.',
            )
          }
          await audit(repository, req, actor, 'assistant.ask', 'assistant', null, 'success', {
            questionLength: question.length,
            referencedDocuments: references.length,
          })
          return sendJson(res, 200, { answer, references })
        }

        if (route === '/api/v1/sharepoint/status' && req.method === 'GET') {
          authorize(actor, 'document:read')
          assert(
            sharePoint,
            503,
            'SHAREPOINT_UNAVAILABLE',
            'SharePoint integration is unavailable.',
          )
          try {
            const status = await sharePoint.getConnectionStatus()
            await audit(repository, req, actor, 'sharepoint.status', 'sharepoint', null)
            return sendJson(res, 200, { status })
          } catch (error) {
            if (error instanceof HttpError) throw error
            logSharePointError(error, req.requestId, 'status')
            throw new HttpError(502, 'SHAREPOINT_CONNECTION_FAILED', 'Documents could not be loaded. Please try again.')
          }
        }
        if (route === '/api/v1/sharepoint/explorer' && req.method === 'GET') {
          authorize(actor, 'document:read')
          assert(
            sharePoint,
            503,
            'SHAREPOINT_UNAVAILABLE',
            'SharePoint integration is unavailable.',
          )
          try {
            const folderId = url.searchParams.get('folder_id') || ''
            const explorer = await sharePoint.list(folderId)
            await audit(repository, req, actor, 'sharepoint.list', 'sharepoint', folderId || 'root', 'success', {
              count: explorer.items.length,
            })
            return sendJson(res, 200, explorer)
          } catch (error) {
            if (error instanceof HttpError) throw error
            logSharePointError(error, req.requestId, 'list')
            throw new HttpError(502, 'SHAREPOINT_LIST_FAILED', 'Documents could not be loaded. Please try again.')
          }
        }
        if (route === '/api/v1/sharepoint/debug' && req.method === 'GET') {
          authorize(actor, 'admin:read')
          assert(
            sharePoint,
            503,
            'SHAREPOINT_UNAVAILABLE',
            'SharePoint integration is unavailable.',
          )
          try {
            const debug = await sharePoint.debug()
            await audit(repository, req, actor, 'sharepoint.debug', 'sharepoint', null, 'success', {
              rootItems: debug.rootItems.length,
              drives: debug.drives.length,
            })
            return sendJson(res, 200, debug)
          } catch (error) {
            if (error instanceof HttpError) throw error
            logSharePointError(error, req.requestId, 'debug')
            throw new HttpError(
              502,
              'SHAREPOINT_DEBUG_FAILED',
              'SharePoint folders could not be loaded. Please try again.',
            )
          }
        }
        if (route === '/api/v1/sharepoint/folders' && req.method === 'GET') {
          authorize(actor, 'document:read')
          const department = cleanText(url.searchParams.get('department'), 'department', { max: 120 })
          const folders = await repository.listDepartmentSharePointFolders(actor, department)
          await audit(repository, req, actor, 'sharepoint.department_folders', 'sharepoint', department, 'success', {
            count: folders.length,
          })
          return sendJson(res, 200, { department, items: folders, folders })
        }
        if (route === '/api/v1/sharepoint/upload' && req.method === 'POST') {
          authorize(actor, 'document:write')
          assert(
            sharePoint,
            503,
            'SHAREPOINT_UNAVAILABLE',
            'SharePoint integration is unavailable.',
          )
          try {
            const item = await sharePoint.upload({
              fileName: String(req.headers['x-file-name'] || ''),
              mimeType: String(req.headers['content-type'] || 'application/octet-stream').toLowerCase(),
              sizeBytes: Number(req.headers['x-file-size'] || req.headers['content-length'] || 0),
              folderId: String(req.headers['x-sharepoint-folder-id'] || ''),
              stream: req,
            })
            await audit(repository, req, actor, 'sharepoint.upload', 'sharepoint', item.id, 'success', {
              size: item.size,
              documentId: String(req.headers['x-docaya-document-id'] || ''),
              documentTitle: String(req.headers['x-docaya-document-title'] || ''),
              documentVersion: String(req.headers['x-docaya-document-version'] || ''),
              folderId: String(req.headers['x-sharepoint-folder-id'] || ''),
            })
            return sendJson(res, 201, { item })
          } catch (error) {
            if (error instanceof HttpError) throw error
            logSharePointError(error, req.requestId, 'upload')
            throw new HttpError(502, 'SHAREPOINT_UPLOAD_FAILED', 'Upload failed. Please try again.')
          }
        }
        const sharePointDownloadMatch = route.match(/^\/api\/v1\/sharepoint\/download\/([^/]+)$/)
        if (sharePointDownloadMatch && req.method === 'GET') {
          authorize(actor, 'document:read')
          assert(
            sharePoint,
            503,
            'SHAREPOINT_UNAVAILABLE',
            'SharePoint integration is unavailable.',
          )
          try {
            const download = await sharePoint.download(decodeURIComponent(sharePointDownloadMatch[1]))
            await audit(repository, req, actor, 'sharepoint.download', 'sharepoint', download.item.id)
            res.writeHead(200, {
              'content-type': download.contentType,
              ...(download.contentLength ? { 'content-length': download.contentLength } : {}),
              'content-disposition': `inline; filename*=UTF-8''${encodeURIComponent(download.item.name)}`,
            })
            return download.response.body ? Readable.fromWeb(download.response.body).pipe(res) : res.end()
          } catch (error) {
            if (error instanceof HttpError) throw error
            logSharePointError(error, req.requestId, 'download')
            throw new HttpError(502, 'SHAREPOINT_DOWNLOAD_FAILED', 'Download failed. Please try again.')
          }
        }
        const sharePointPreviewMatch = route.match(/^\/api\/v1\/sharepoint\/preview\/([^/]+)$/)
        if (sharePointPreviewMatch && req.method === 'GET') {
          authorize(actor, 'document:read')
          assert(
            sharePoint,
            503,
            'SHAREPOINT_UNAVAILABLE',
            'SharePoint integration is unavailable.',
          )
          try {
            const preview = await sharePoint.getPreview(decodeURIComponent(sharePointPreviewMatch[1]))
            await audit(repository, req, actor, 'sharepoint.preview', 'sharepoint', preview.id)
            return sendJson(res, 200, { preview })
          } catch (error) {
            if (error instanceof HttpError) throw error
            logSharePointError(error, req.requestId, 'preview')
            throw new HttpError(502, 'SHAREPOINT_PREVIEW_FAILED', 'Preview could not be opened. Please try again.')
          }
        }

        if (route === '/api/v1/workflows' && req.method === 'GET') {
          authorize(actor, 'workflow:read')
          return sendJson(res, 200, { workflows: await repository.listWorkflows(actor) })
        }
        if (route === '/api/v1/workflows' && req.method === 'POST') {
          authorize(actor, 'document:write')
          assert(
            req.headers['idempotency-key'],
            400,
            'IDEMPOTENCY_KEY_REQUIRED',
            'An Idempotency-Key header is required.',
          )
          const body = await readJson(req, config.bodyLimit)
          cleanText(body.documentId, 'documentId')
          cleanText(body.name, 'name')
          assert(
            await repository.getDocument(actor, body.documentId),
            404,
            'NOT_FOUND',
            'Document not found.',
          )
          const workflow = await repository.createWorkflow(actor, { ...body, id: uuid('workflow') })
          await audit(repository, req, actor, 'workflow.start', 'workflow', workflow.id)
          return sendJson(res, 201, { workflow })
        }
        const decisionMatch = route.match(/^\/api\/v1\/workflows\/([^/]+)\/decision$/)
        if (decisionMatch && req.method === 'POST') {
          authorize(actor, 'workflow:approve')
          const body = await readJson(req, config.bodyLimit)
          assert(
            workflowDecisions.includes(body.status),
            422,
            'VALIDATION_ERROR',
            'Workflow decision is invalid.',
          )
          const workflow = await repository.decideWorkflow(
            actor,
            decodeURIComponent(decisionMatch[1]),
            body.status,
          )
          assert(workflow, 404, 'NOT_FOUND', 'Pending workflow not found.')
          await audit(repository, req, actor, 'workflow.decision', 'workflow', workflow.id, 'success', {
            status: body.status,
          })
          return sendJson(res, 200, { workflow })
        }

        if (route === '/api/v1/notifications' && req.method === 'GET') {
          authorize(actor, 'notification:read')
          const limit = Math.min(Math.max(Number(url.searchParams.get('limit') || 25), 1), 100)
          const result = await repository.listNotifications(actor, {
            cursor: parseCursor(url.searchParams.get('cursor')),
            limit,
          })
          return sendJson(res, 200, {
            notifications: result.items,
            page: { nextCursor: result.nextCursor, limit },
          })
        }
        if (route === '/api/v1/notifications/read' && req.method === 'POST') {
          authorize(actor, 'notification:read')
          const body = await readJson(req, config.bodyLimit)
          assert(
            body.ids === undefined || (Array.isArray(body.ids) && body.ids.length <= 100),
            422,
            'VALIDATION_ERROR',
            'ids must be an array of at most 100 items.',
          )
          const updated = await repository.markNotificationsRead(actor, body.ids)
          await audit(repository, req, actor, 'notification.read', 'notification', null, 'success', {
            updated,
          })
          return sendJson(res, 200, { updated })
        }
        const notificationStateMatch = route.match(/^\/api\/v1\/notifications\/([^/]+)\/(read|unread)$/)
        if (notificationStateMatch && req.method === 'POST') {
          authorize(actor, 'notification:read')
          const id = decodeURIComponent(notificationStateMatch[1])
          const updated =
            notificationStateMatch[2] === 'read'
              ? await repository.markNotificationsRead(actor, [id])
              : await repository.markNotificationUnread(actor, id)
          assert(updated, 404, 'NOT_FOUND', 'Notification not found.')
          await audit(repository, req, actor, `notification.${notificationStateMatch[2]}`, 'notification', id)
          return sendJson(res, 200, { updated: Number(updated) })
        }
        if (route === '/api/v1/notifications/preferences' && req.method === 'GET') {
          authorize(actor, 'notification:read')
          return sendJson(res, 200, { preferences: await repository.getNotificationPreferences(actor) })
        }
        if (route === '/api/v1/notifications/preferences' && req.method === 'PATCH') {
          authorize(actor, 'notification:read')
          const body = await readJson(req, config.bodyLimit)
          const allowed = ['workflowEnabled', 'collaborationEnabled', 'securityEnabled', 'systemEnabled']
          const changes = Object.fromEntries(
            Object.entries(body).filter(
              ([key, value]) => allowed.includes(key) && typeof value === 'boolean',
            ),
          )
          assert(
            Object.keys(changes).length > 0,
            422,
            'VALIDATION_ERROR',
            'At least one boolean preference is required.',
          )
          return sendJson(res, 200, {
            preferences: await repository.updateNotificationPreferences(actor, changes),
          })
        }
        if (route === '/api/v1/audit' && req.method === 'GET') {
          authorize(actor, 'audit:read')
          return sendJson(res, 200, {
            events: await repository.listAudit(
              actor,
              Math.min(Number(url.searchParams.get('limit') || 100), 200),
            ),
          })
        }
        if (route === '/api/v1/admin/overview' && req.method === 'GET') {
          authorize(actor, 'admin:read')
          await audit(repository, req, actor, 'admin.overview', 'admin', null)
          return sendJson(res, 200, { overview: await repository.getOverview(actor) })
        }
        if (route === '/api/v1/admin/users' && req.method === 'GET') {
          authorize(actor, 'admin:read')
          return sendJson(res, 200, { users: await repository.listAdminUsers(actor) })
        }
        if (route === '/api/v1/admin/users' && req.method === 'POST') {
          authorize(actor, 'admin:write')
          const body = await readJson(req, config.bodyLimit)
          const name = cleanText(body.name, 'name', { max: 120 })
          const email = cleanText(body.email, 'email', { max: 200 }).toLowerCase()
          assert(emailPattern.test(email), 422, 'VALIDATION_ERROR', 'A valid email address is required.')
          const role = body.role || 'viewer'
          assert(assignableRoles.includes(role), 422, 'VALIDATION_ERROR', 'The selected role is invalid.')
          const status = body.status || 'invited'
          assert(userStatuses.includes(status), 422, 'VALIDATION_ERROR', 'The selected status is invalid.')
          let user
          try {
            user = await repository.createAdminUser(actor, { id: uuid('user'), name, email, role, status })
          } catch (error) {
            if (error?.code === '23505' || /unique/i.test(error?.message || ''))
              throw new HttpError(409, 'DUPLICATE_EMAIL', 'A user with this email already exists.')
            throw error
          }
          await audit(repository, req, actor, 'admin.user.create', 'user', user.id, 'success', {
            role,
            status,
          })
          return sendJson(res, 201, { user })
        }
        const adminUserMatch = route.match(/^\/api\/v1\/admin\/users\/([^/]+)$/)
        if (adminUserMatch && req.method === 'PATCH') {
          authorize(actor, 'admin:write')
          const id = decodeURIComponent(adminUserMatch[1])
          const body = await readJson(req, config.bodyLimit)
          const changes = {}
          if (body.role !== undefined) {
            assert(
              assignableRoles.includes(body.role),
              422,
              'VALIDATION_ERROR',
              'The selected role is invalid.',
            )
            changes.role = body.role
          }
          if (body.status !== undefined) {
            assert(
              userStatuses.includes(body.status),
              422,
              'VALIDATION_ERROR',
              'The selected status is invalid.',
            )
            assert(
              id !== actor.id || body.status === 'active',
              409,
              'SELF_ACTION_DENIED',
              'You cannot suspend or deprovision your own account.',
            )
            changes.status = body.status
          }
          assert(Object.keys(changes).length > 0, 422, 'VALIDATION_ERROR', 'No valid fields were supplied.')
          const user = await repository.updateAdminUser(actor, id, changes)
          assert(user, 404, 'NOT_FOUND', 'User not found.')
          await audit(repository, req, actor, 'admin.user.update', 'user', id, 'success', {
            fields: Object.keys(changes),
          })
          return sendJson(res, 200, { user })
        }
        if (route === '/api/v1/admin/settings' && req.method === 'GET') {
          authorize(actor, 'admin:read')
          return sendJson(res, 200, { settings: await repository.getSettings(actor) })
        }
        if (route === '/api/v1/admin/settings' && req.method === 'PATCH') {
          authorize(actor, 'admin:write')
          const body = await readJson(req, config.bodyLimit)
          const changes = {}
          if (typeof body.defaultLanguage === 'string')
            changes.defaultLanguage = cleanText(body.defaultLanguage, 'defaultLanguage', { max: 50 })
          if (typeof body.defaultRetention === 'string')
            changes.defaultRetention = cleanText(body.defaultRetention, 'defaultRetention', { max: 50 })
          if (typeof body.requireWorkflowOnSubmit === 'boolean')
            changes.requireWorkflowOnSubmit = body.requireWorkflowOnSubmit
          if (typeof body.notifyOnDocumentEvents === 'boolean')
            changes.notifyOnDocumentEvents = body.notifyOnDocumentEvents
          assert(Object.keys(changes).length > 0, 422, 'VALIDATION_ERROR', 'No valid settings were supplied.')
          const settings = await repository.updateSettings(actor, changes)
          await audit(repository, req, actor, 'admin.settings.update', 'settings', actor.tenantId)
          return sendJson(res, 200, { settings })
        }
        if (route === '/api/v1/admin/sharepoint/folders' && req.method === 'GET') {
          authorize(actor, 'admin:read')
          assert(
            sharePoint,
            503,
            'SHAREPOINT_UNAVAILABLE',
            'SharePoint integration is unavailable.',
          )
          try {
            const folderId = url.searchParams.get('folder_id') || ''
            const folders = await sharePoint.list(folderId)
            await audit(repository, req, actor, 'admin.sharepoint.folders', 'sharepoint', null, 'success', {
              count: folders.items.length,
            })
            return sendJson(res, 200, folders)
          } catch (error) {
            if (error instanceof HttpError) throw error
            logSharePointError(error, req.requestId, 'admin-list-folders')
            throw new HttpError(
              502,
              'SHAREPOINT_FOLDER_LIST_FAILED',
              'SharePoint folders could not be loaded. Please try again.',
            )
          }
        }
        if (route === '/api/v1/admin/department-sharepoint-folders' && req.method === 'GET') {
          authorize(actor, 'admin:read')
          return sendJson(res, 200, { items: await repository.listDepartmentSharePointFolderAccess(actor) })
        }
        const departmentFolderRoute = route.match(/^\/api\/v1\/admin\/departments\/([^/]+)\/sharepoint-folders(?:\/([^/]+))?$/)
        if (departmentFolderRoute && req.method === 'GET' && !departmentFolderRoute[2]) {
          authorize(actor, 'admin:read')
          const department = cleanText(decodeURIComponent(departmentFolderRoute[1]), 'department', { max: 120 })
          return sendJson(res, 200, {
            department,
            items: await repository.listDepartmentSharePointFolders(actor, department),
          })
        }
        if (departmentFolderRoute && req.method === 'POST' && !departmentFolderRoute[2]) {
          authorize(actor, 'admin:write')
          const department = cleanText(decodeURIComponent(departmentFolderRoute[1]), 'department', { max: 120 })
          const body = await readJson(req, config.bodyLimit)
          const folderId = cleanText(body.id || body.sharePointFolderId, 'sharePointFolderId', { max: 300 })
          const folder = await repository.addDepartmentSharePointFolder(actor, department, {
            id: uuid('spmap'),
            sharePointFolderId: folderId,
            name: cleanText(body.name, 'name', { max: 240 }),
            path: cleanText(body.path, 'path', { max: 1000 }),
            webUrl: body.webUrl ? cleanText(body.webUrl, 'webUrl', { max: 2000 }) : null,
          })
          await audit(repository, req, actor, 'admin.folder_access.add', 'sharepoint-folder-access', folder.id, 'success', {
            department,
            sharePointFolderId: folder.sharePointFolderId,
          })
          return sendJson(res, 201, { item: folder })
        }
        if (departmentFolderRoute && req.method === 'DELETE' && departmentFolderRoute[2]) {
          authorize(actor, 'admin:write')
          const department = cleanText(decodeURIComponent(departmentFolderRoute[1]), 'department', { max: 120 })
          const folderId = decodeURIComponent(departmentFolderRoute[2])
          const removed = await repository.removeDepartmentSharePointFolder(actor, department, folderId)
          assert(removed, 404, 'NOT_FOUND', 'Folder access mapping not found.')
          await audit(repository, req, actor, 'admin.folder_access.remove', 'sharepoint-folder-access', folderId, 'success', {
            department,
          })
          return sendJson(res, 200, { success: true })
        }

        const knownPath = route.match(
          /^\/api\/v1\/(documents|trash|uploads|search|workflows|notifications|audit|admin|assistant|sharepoint)(?:\/|$)/,
        )
        if (knownPath) {
          res.setHeader('allow', allowedMethods(route))
          throw new HttpError(
            405,
            'METHOD_NOT_ALLOWED',
            'The HTTP method is not supported for this resource.',
          )
        }
        throw new HttpError(404, 'NOT_FOUND', 'Route not found.')
      }

      if (req.method === 'GET' && existsSync(staticRoot)) return serveStatic(route, staticRoot, res)
      throw new HttpError(404, 'NOT_FOUND', 'Route not found.')
    } catch (error) {
      if (req.actor) {
        try {
          await audit(repository, req, req.actor, 'request.failed', 'http', req.url, 'failure', {
            code: error.code || 'INTERNAL_ERROR',
            method: req.method,
          })
        } catch (auditError) {
          console.error(
            JSON.stringify({
              level: 'error',
              event: 'audit.append.failed',
              requestId: req.requestId,
              message: auditError.message,
            }),
          )
        }
      }
      return sendError(res, error, req.requestId || `req_${crypto.randomUUID()}`)
    }
  }
  return createServer(handler)
}

const assistantStopWords = new Set([
  'what',
  'when',
  'where',
  'which',
  'who',
  'whom',
  'whose',
  'why',
  'how',
  'does',
  'did',
  'the',
  'this',
  'that',
  'these',
  'those',
  'and',
  'for',
  'with',
  'from',
  'has',
  'have',
  'was',
  'were',
  'about',
  'please',
  'tell',
  'give',
])
async function findRelevantDocuments(repository, actor, question, limit = 5) {
  const words = (question.toLowerCase().match(/[a-z0-9]{3,}/g) || []).filter(
    (word) => !assistantStopWords.has(word),
  )
  const pool = await repository.listDocuments(actor, { limit: 100 })
  const scored = pool.items
    .map((document) => {
      const haystack =
        `${document.title} ${document.number} ${document.summary} ${document.department} ${document.type}`.toLowerCase()
      const score = words.reduce((total, word) => total + (haystack.includes(word) ? 1 : 0), 0)
      return { document, score }
    })
    .filter((entry) => entry.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, limit)
  return scored.map(({ document }) => ({
    id: document.id,
    number: document.number,
    title: document.title,
    status: document.status,
    department: document.department,
    nextReview: document.nextReview,
  }))
}
function publicUpload(upload) {
  return {
    id: upload.id,
    fileName: upload.file_name,
    mimeType: upload.mime_type,
    sizeBytes: upload.size_bytes,
    chunkSize: upload.chunk_size,
    state: upload.state,
    receivedBytes: upload.received_bytes,
    contentHash: upload.content_hash,
    createdAt: upload.created_at,
    updatedAt: upload.updated_at,
  }
}
function allowedMethods(route) {
  if (route.includes('/chunks/')) return 'PUT'
  if (
    route.endsWith('/complete') ||
    route.endsWith('/restore') ||
    route.endsWith('/decision') ||
    route.endsWith('/read')
  )
    return 'POST'
  if (route.match(/\/documents\/[^/]+$/)) return 'GET, PATCH, DELETE'
  if (route.match(/\/admin\/users\/[^/]+$/)) return 'PATCH'
  if (route.match(/\/uploads\/[^/]+$/)) return 'DELETE'
  return route.endsWith('/search') || route.endsWith('/assistant/ask') ? 'POST' : 'GET, POST'
}
function serveStatic(route, root, res) {
  const target = route === '/' ? 'index.html' : decodeURIComponent(route).replace(/^\/+/, '')
  let path = resolve(root, target)
  if (!path.startsWith(`${root}${sep}`) && path !== root)
    throw new HttpError(404, 'NOT_FOUND', 'Route not found.')
  if (!existsSync(path) || statSync(path).isDirectory()) path = resolve(root, 'index.html')
  res.writeHead(200, {
    'content-type': contentTypes[extname(path)] || 'application/octet-stream',
    'cache-control': path.endsWith('index.html') ? 'no-cache' : 'public, max-age=31536000, immutable',
  })
  createReadStream(path).pipe(res)
}
