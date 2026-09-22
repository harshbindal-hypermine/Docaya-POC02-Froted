import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { createDocayaServer } from '../../server/app/create-server.mjs'
import { SqliteRepository } from '../../server/db/sqlite.mjs'
import { issueDevelopmentToken } from '../../server/middleware/auth.mjs'

const root = mkdtempSync(join(tmpdir(), 'docaya-test-'))
const config = {
  environment: 'test',
  authMode: 'development',
  devTokenSecret: 'a-development-secret-with-32-characters',
  corsOrigins: ['http://localhost:5173'],
  bodyLimit: 1_048_576,
  uploadMaxBytes: 1_048_576,
  uploadExtensions: ['txt'],
  uploadMimeTypes: ['text/plain'],
  localSqlite: true,
  sqlitePath: join(root, 'test.sqlite'),
  uploadPath: join(root, 'uploads'),
}
const repository = new SqliteRepository(config.sqlitePath, { seed: true })
const mockAssistant = {
  async chat(messages) {
    const userMessage = messages.find((entry) => entry.role === 'user')?.content || ''
    if (userMessage.includes('TRIGGER_UPSTREAM_FAILURE')) throw new Error('simulated upstream failure')
    return 'Mock answer grounded in the supplied controlled-document context.'
  },
}
const mockSharePoint = {
  async healthCheck() {
    return { status: 'ok' }
  },
  async listFolders() {
    return {
      items: [
        {
          id: 'root-uploads',
          name: 'Docaya Uploads',
          type: 'folder',
          size: null,
          path: '/Docaya Uploads',
          webUrl: 'https://contoso.sharepoint.com/sites/docaya/Shared%20Documents/Docaya%20Uploads',
          lastModifiedDateTime: '2026-09-20T11:10:24Z',
          mimeType: null,
        },
        {
          id: 'root-library',
          name: 'Document Library',
          type: 'folder',
          size: null,
          path: '/Document Library',
          webUrl: 'https://contoso.sharepoint.com/sites/docaya/Shared%20Documents/Document%20Library',
          lastModifiedDateTime: '2026-09-19T08:32:40Z',
          mimeType: null,
        },
      ],
      drive: { id: 'drive-id', name: 'Documents' },
      site: { hostname: 'hypermineai.sharepoint.com', path: '/sites/DocayaDocLibrary-Dev' },
    }
  },
  async list(folderId = '') {
    if (!folderId) return this.listFolders()
    return {
      folderId,
      items: [
        {
          id: 'hr-policies',
          name: 'HR Policies',
          type: 'folder',
          size: null,
          path: '/Document Library/HR Policies',
          webUrl: 'https://contoso.sharepoint.com/sites/docaya/Shared%20Documents/Document%20Library/HR%20Policies',
          lastModifiedDateTime: '2026-09-20T12:00:00Z',
          mimeType: null,
        },
      ],
      drive: { id: 'drive-id', name: 'Documents' },
      site: { hostname: 'hypermineai.sharepoint.com', path: '/sites/DocayaDocLibrary-Dev' },
    }
  },
  async debug() {
    return {
      site: {
        id: 'site-id',
        displayName: 'Docaya Doc Library - Dev',
        webUrl: 'https://hypermineai.sharepoint.com/sites/DocayaDocLibrary-Dev',
      },
      drives: [{ id: 'drive-id', name: 'Documents' }],
      drive: { id: 'drive-id', name: 'Documents' },
      rootItems: (await this.listFolders()).items,
    }
  },
}
const server = createDocayaServer({ config, repository, assistant: mockAssistant, sharePoint: mockSharePoint })
const unconfiguredServer = createDocayaServer({ config, repository, assistant: null })
let base = ''
let unconfiguredBase = ''
let admin = ''
let viewer = ''

beforeAll(async () => {
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve))
  const address = server.address()
  base = `http://127.0.0.1:${address.port}`
  await new Promise((resolve) => unconfiguredServer.listen(0, '127.0.0.1', resolve))
  const unconfiguredAddress = unconfiguredServer.address()
  unconfiguredBase = `http://127.0.0.1:${unconfiguredAddress.port}`
  admin = await issueDevelopmentToken(config)
  viewer = await issueDevelopmentToken(config, { sub: 'viewer', roles: ['viewer'] })
})
afterAll(async () => {
  await new Promise((resolve) => server.close(resolve))
  await new Promise((resolve) => unconfiguredServer.close(resolve))
  repository.close()
  const resolved = resolve(root)
  if (resolved.startsWith(resolve(tmpdir()))) rmSync(resolved, { recursive: true, force: true })
})
const call = (path, token = admin, init = {}) =>
  fetch(`${base}${path}`, {
    ...init,
    headers: { ...(token ? { authorization: `Bearer ${token}` } : {}), ...init.headers },
  })

describe('v1 API contract', () => {
  const callUnconfigured = (path, token = admin, init = {}) =>
    fetch(`${unconfiguredBase}${path}`, {
      ...init,
      headers: { authorization: `Bearer ${token}`, ...init.headers },
    })
  it('exposes minimal unauthenticated health', async () => {
    const response = await call('/health/live', '')
    expect(response.status).toBe(200)
    expect(await response.json()).toEqual(expect.objectContaining({ status: 'ok' }))
  })
  it('reports readiness with a working database health check', async () => {
    const response = await call('/health/ready', '')
    expect(response.status).toBe(200)
    const payload = await response.json()
    expect(payload.status).toBe('ready')
    expect(payload.checks.database).toMatchObject({ status: 'ok' })
  })
  it('rejects unauthenticated business access with request id', async () => {
    const response = await call('/api/v1/documents', '')
    expect(response.status).toBe(401)
    expect((await response.json()).error).toMatchObject({
      code: 'UNAUTHENTICATED',
      requestId: expect.any(String),
    })
  })
  it('provides deterministic cursor pagination', async () => {
    const first = await call('/api/v1/documents?limit=2')
    const payload = await first.json()
    expect(payload.documents).toHaveLength(2)
    expect(payload.page.nextCursor).toEqual(expect.any(String))
    const second = await call(`/api/v1/documents?limit=2&cursor=${payload.page.nextCursor}`)
    expect((await second.json()).documents).toHaveLength(2)
  })
  it('enforces role capabilities', async () => {
    const response = await call('/api/v1/documents', viewer, {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'idempotency-key': 'viewer-write' },
      body: '{}',
    })
    expect(response.status).toBe(403)
  })
  it('validates document registration', async () => {
    const response = await call('/api/v1/documents', admin, {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'idempotency-key': 'bad-document' },
      body: '{}',
    })
    expect(response.status).toBe(422)
  })
  it('returns 404 for a missing update and 405 for unsupported methods', async () => {
    const missing = await call('/api/v1/documents/missing', admin, {
      method: 'PATCH',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ title: 'Valid title' }),
    })
    expect(missing.status).toBe(404)
    const method = await call('/api/v1/audit', admin, { method: 'POST' })
    expect(method.status).toBe(405)
    expect(method.headers.get('allow')).toBeTruthy()
  })
  it('runs resumable upload, hash, scan, and commit', async () => {
    const bytes = new TextEncoder().encode('safe test file')
    const sessionResponse = await call('/api/v1/uploads', admin, {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'idempotency-key': 'upload-one' },
      body: JSON.stringify({
        fileName: 'record.txt',
        mimeType: 'text/plain',
        sizeBytes: bytes.length,
        metadata: {
          type: 'Record',
          department: 'Legal',
          library: 'Governance',
          summary: 'Safe test',
          retention: '2 years',
          nextReview: '2027-09-03',
          reviewer: 'Reviewer',
          approver: 'Approver',
        },
      }),
    })
    expect(sessionResponse.status).toBe(201)
    const session = await sessionResponse.json()
    expect(
      (
        await call(`/api/v1/uploads/${session.upload.id}/chunks/0`, admin, {
          method: 'PUT',
          headers: { 'content-type': 'application/octet-stream', 'content-length': String(bytes.length) },
          body: bytes,
        })
      ).status,
    ).toBe(200)
    const completed = await call(`/api/v1/uploads/${session.upload.id}/complete`, admin, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ metadata: { title: 'Safe Record' } }),
    })
    expect(completed.status).toBe(201)
    expect((await completed.json()).document.contentHash).toMatch(/^[a-f0-9]{64}$/)
  })
  it('blocks a disallowed origin', async () => {
    const response = await call('/api/v1/documents', admin, { headers: { origin: 'https://evil.example' } })
    expect(response.status).toBe(403)
  })
  it('lists admin users with lifecycle timestamps', async () => {
    const response = await call('/api/v1/admin/users')
    expect(response.status).toBe(200)
    const payload = await response.json()
    expect(payload.users.length).toBeGreaterThan(0)
    expect(payload.users[0]).toMatchObject({ id: expect.any(String), created_at: expect.any(String) })
  })
  it('rejects admin user management for non-admin roles', async () => {
    const response = await call('/api/v1/admin/users', viewer, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ name: 'Nope', email: 'nope@docaya.local', role: 'viewer' }),
    })
    expect(response.status).toBe(403)
  })
  it('validates new admin user input', async () => {
    const response = await call('/api/v1/admin/users', admin, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ name: 'A', email: 'not-an-email', role: 'super-admin' }),
    })
    expect(response.status).toBe(422)
  })
  it('creates, lists and updates an admin user end to end', async () => {
    const created = await call('/api/v1/admin/users', admin, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ name: 'Test Invitee', email: 'invitee@docaya.local', role: 'viewer' }),
    })
    expect(created.status).toBe(201)
    const { user } = await created.json()
    expect(user).toMatchObject({ name: 'Test Invitee', role: 'viewer', status: 'invited' })

    const duplicate = await call('/api/v1/admin/users', admin, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ name: 'Test Invitee', email: 'invitee@docaya.local', role: 'viewer' }),
    })
    expect(duplicate.status).toBe(409)

    const updated = await call(`/api/v1/admin/users/${user.id}`, admin, {
      method: 'PATCH',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ role: 'content-manager', status: 'active' }),
    })
    expect(updated.status).toBe(200)
    const updatedBody = await updated.json()
    expect(updatedBody.user).toMatchObject({ role: 'content-manager', status: 'active' })
    expect(updatedBody.user.updated_at).toEqual(expect.any(String))

    const missing = await call('/api/v1/admin/users/missing-user', admin, {
      method: 'PATCH',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ role: 'viewer' }),
    })
    expect(missing.status).toBe(404)
  })
  it('uses explicit department SharePoint folder mappings for registration', async () => {
    const financeInvoice = {
      id: 'sp-folder-invoices',
      name: 'Invoices',
      path: '/Document Library/Invoices',
      webUrl: 'https://contoso.sharepoint.com/sites/docaya/Shared%20Documents/Invoices',
    }
    const financeReports = {
      id: 'sp-folder-reports',
      name: 'Financial Reports',
      path: '/Document Library/Financial Reports',
      webUrl: 'https://contoso.sharepoint.com/sites/docaya/Shared%20Documents/Reports',
    }
    const hrPolicies = {
      id: 'sp-folder-hr-policies',
      name: 'HR Policies',
      path: '/Document Library/Policies/HR Policies',
      webUrl: 'https://contoso.sharepoint.com/sites/docaya/Shared%20Documents/Policies/HR%20Policies',
    }

    for (const [department, folder] of [
      ['Finance', financeInvoice],
      ['Finance', financeReports],
      ['HR', hrPolicies],
    ]) {
      const response = await call(
        `/api/v1/admin/departments/${encodeURIComponent(department)}/sharepoint-folders`,
        admin,
        {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify(folder),
        },
      )
      expect(response.status).toBe(201)
    }

    const financeFolders = await call('/api/v1/sharepoint/folders?department=Finance')
    expect(financeFolders.status).toBe(200)
    const financePayload = await financeFolders.json()
    expect(financePayload.items.map((folder) => folder.id)).toEqual([
      financeReports.id,
      financeInvoice.id,
    ])
    expect(financePayload.items[1]).toMatchObject({
      id: financeInvoice.id,
      sharePointFolderId: financeInvoice.id,
      name: 'Invoices',
      department: 'Finance',
    })

    const created = await call('/api/v1/documents', admin, {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'idempotency-key': 'finance-invoice-folder' },
      body: JSON.stringify({
        title: 'Quarterly invoice register',
        number: 'FIN-INV-00001',
        type: 'Record',
        department: 'Finance',
        library: 'Corporate Governance',
        language: 'English',
        summary: 'Finance registration using an explicitly assigned SharePoint folder.',
        retention: '7 years',
        nextReview: '2027-09-03',
        reviewer: 'Finance Quality Lead',
        approver: 'Finance Director',
        sharePointFolderId: financeInvoice.id,
      }),
    })
    expect(created.status).toBe(201)
    expect((await created.json()).document).toMatchObject({
      department: 'Finance',
      sharePointFolderId: financeInvoice.id,
      sharePointFolderName: 'Invoices',
      sharePointFolderPath: '/Document Library/Invoices',
    })

    const rejected = await call('/api/v1/documents', admin, {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'idempotency-key': 'finance-hr-folder' },
      body: JSON.stringify({
        title: 'Wrong folder registration',
        number: 'FIN-INV-00002',
        type: 'Record',
        department: 'Finance',
        library: 'Corporate Governance',
        language: 'English',
        summary: 'This should not be allowed.',
        retention: '7 years',
        nextReview: '2027-09-03',
        reviewer: 'Finance Quality Lead',
        approver: 'Finance Director',
        sharePointFolderId: hrPolicies.id,
      }),
    })
    expect(rejected.status).toBe(422)
    expect((await rejected.json()).error).toMatchObject({
      code: 'SHAREPOINT_FOLDER_NOT_ALLOWED',
      message: 'Selected SharePoint folder is not available for this department.',
    })
  })
  it('prevents an admin from suspending or deprovisioning their own account', async () => {
    const me = await (await call('/api/v1/auth/me')).json()
    const response = await call(`/api/v1/admin/users/${me.user.id}`, admin, {
      method: 'PATCH',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ status: 'suspended' }),
    })
    expect(response.status).toBe(409)
  })
  it('rejects unauthenticated assistant access', async () => {
    const response = await call('/api/v1/assistant/ask', '', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ question: 'What is the Data Protection Policy status?' }),
    })
    expect(response.status).toBe(401)
  })
  it('rejects unauthenticated SharePoint explorer access', async () => {
    const response = await call('/api/v1/sharepoint/explorer', '')
    expect(response.status).toBe(401)
  })
  it('lets admins browse SharePoint root folders and child folders by item id', async () => {
    const rootResponse = await call('/api/v1/admin/sharepoint/folders')
    expect(rootResponse.status).toBe(200)
    const rootPayload = await rootResponse.json()
    expect(rootPayload.items.map((item) => item.name)).toEqual(['Docaya Uploads', 'Document Library'])

    const childResponse = await call('/api/v1/admin/sharepoint/folders?folder_id=root-library')
    expect(childResponse.status).toBe(200)
    const childPayload = await childResponse.json()
    expect(childPayload).toMatchObject({ folderId: 'root-library' })
    expect(childPayload.items).toEqual([expect.objectContaining({ id: 'hr-policies', name: 'HR Policies' })])
  })
  it('returns protected safe SharePoint diagnostics', async () => {
    const response = await call('/api/v1/sharepoint/debug')
    expect(response.status).toBe(200)
    const payload = await response.json()
    expect(payload.site).toMatchObject({
      id: 'site-id',
      displayName: 'Docaya Doc Library - Dev',
      webUrl: 'https://hypermineai.sharepoint.com/sites/DocayaDocLibrary-Dev',
    })
    expect(payload.drive).toMatchObject({ id: 'drive-id', name: 'Documents' })
    expect(payload.rootItems.map((item) => item.name)).toEqual(['Docaya Uploads', 'Document Library'])
    expect(JSON.stringify(payload)).not.toContain('token')
    expect(JSON.stringify(payload)).not.toContain('secret')
  })
  it('returns a SharePoint-specific error when the integration is unavailable', async () => {
    const response = await callUnconfigured('/api/v1/sharepoint/explorer')
    expect(response.status).toBe(503)
    expect((await response.json()).error).toMatchObject({
      code: 'SHAREPOINT_UNAVAILABLE',
      message: 'SharePoint integration is unavailable.',
    })
  })
  it('validates the assistant question length', async () => {
    const response = await call('/api/v1/assistant/ask', admin, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ question: 'a' }),
    })
    expect(response.status).toBe(422)
  })
  it('answers a question grounded in controlled-document context', async () => {
    const response = await call('/api/v1/assistant/ask', viewer, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ question: 'What is the Document Control Procedure status?' }),
    })
    expect(response.status).toBe(200)
    const payload = await response.json()
    expect(payload.answer).toEqual(expect.any(String))
    expect(Array.isArray(payload.references)).toBe(true)
  })
  it('returns a bad-gateway error when the upstream model fails', async () => {
    const response = await call('/api/v1/assistant/ask', admin, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ question: 'TRIGGER_UPSTREAM_FAILURE for this question' }),
    })
    expect(response.status).toBe(502)
  })
  it('reports the assistant as unavailable when Azure OpenAI is not configured', async () => {
    const response = await callUnconfigured('/api/v1/assistant/ask', admin, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ question: 'What is the Data Protection Policy status?' }),
    })
    expect(response.status).toBe(503)
    expect((await response.json()).error.code).toBe('ASSISTANT_UNAVAILABLE')
  })
})
