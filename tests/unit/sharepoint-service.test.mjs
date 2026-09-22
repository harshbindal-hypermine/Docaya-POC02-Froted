import { describe, expect, it } from 'vitest'
import { normalizeItem, SharePointGraphError, SharePointService } from '../../server/services/sharepoint.mjs'

const config = {
  sharePointTenantId: 'tenant-id',
  sharePointClientId: 'client-id',
  sharePointClientSecret: 'client-secret',
  sharePointHostname: 'hypermineai.sharepoint.com',
  sharePointSitePath: '/sites/DocayaDocLibrary-Dev',
  sharePointDriveName: 'Shared Documents',
  uploadMaxBytes: 10_000,
  uploadExtensions: ['txt'],
  uploadMimeTypes: ['text/plain'],
}

const credential = { async getToken() { return { token: 'test-token' } } }

function jsonResponse(payload, init = {}) {
  return new Response(JSON.stringify(payload), {
    status: 200,
    headers: { 'content-type': 'application/json', ...init.headers },
    ...init,
  })
}

describe('SharePointService', () => {
  it('normalizes Graph drive items for the frontend', () => {
    expect(
      normalizeItem({
        id: 'folder-1',
        name: 'Policies',
        folder: { childCount: 2 },
        webUrl: 'https://example',
        lastModifiedDateTime: '2026-09-20T10:00:00Z',
      }),
    ).toMatchObject({ id: 'folder-1', name: 'Policies', type: 'folder', size: null })
    expect(
      normalizeItem({
        id: 'file-1',
        name: 'policy.txt',
        size: 12,
        file: { mimeType: 'text/plain' },
      }),
    ).toMatchObject({ id: 'file-1', name: 'policy.txt', type: 'file', size: 12, mimeType: 'text/plain' })
  })

  it('resolves site and drive before listing root children', async () => {
    const seen = []
    const service = new SharePointService(config, {
      credential,
      fetchImpl: async (url) => {
        seen.push(url)
        if (url.includes('/sites/hypermineai.sharepoint.com:/sites/DocayaDocLibrary-Dev'))
          return jsonResponse({ id: 'site-id', displayName: 'Docaya site' })
        if (url.endsWith('/sites/site-id/drives'))
          return jsonResponse({ value: [{ id: 'drive-id', name: 'Shared Documents' }] })
        if (url.endsWith('/drives/drive-id/root/children'))
          return jsonResponse({ value: [{ id: 'file-1', name: 'policy.txt', size: 10, file: { mimeType: 'text/plain' } }] })
        throw new Error(`Unexpected URL ${url}`)
      },
    })

    const result = await service.list()
    expect(result.items).toEqual([expect.objectContaining({ id: 'file-1', type: 'file' })])
    expect(seen).toHaveLength(3)
  })

  it('discovers folder trees without department-name filtering', async () => {
    const service = new SharePointService(config, {
      credential,
      fetchImpl: async (url) => {
        if (url.includes('/sites/hypermineai.sharepoint.com:/sites/DocayaDocLibrary-Dev'))
          return jsonResponse({ id: 'site-id', displayName: 'Docaya site' })
        if (url.endsWith('/sites/site-id/drives'))
          return jsonResponse({ value: [{ id: 'drive-id', name: 'Shared Documents' }] })
        if (url.includes('/drives/drive-id/root/children'))
          return jsonResponse({
            value: [
              { id: 'container-1', name: 'Document Library', folder: {}, parentReference: { path: '/drive/root:' } },
              { id: 'file-1', name: 'notes.txt', file: { mimeType: 'text/plain' }, parentReference: { path: '/drive/root:' } },
            ],
          })
        if (url.includes('/drives/drive-id/items/container-1/children'))
          return jsonResponse({
            value: [
              { id: 'folder-invoices', name: 'Invoices', folder: {}, parentReference: { path: '/drive/root:/Document Library' } },
              { id: 'folder-hr', name: 'HR Policies', folder: {}, parentReference: { path: '/drive/root:/Document Library' } },
            ],
          })
        if (url.includes('/drives/drive-id/items/folder-invoices/children')) return jsonResponse({ value: [] })
        if (url.includes('/drives/drive-id/items/folder-hr/children')) return jsonResponse({ value: [] })
        throw new Error(`Unexpected URL ${url}`)
      },
    })

    const result = await service.listFolders()
    expect(result.items.map((item) => item.id)).toEqual(['container-1', 'folder-invoices', 'folder-hr'])
    expect(result.items.find((item) => item.id === 'folder-invoices')).toMatchObject({ name: 'Invoices' })
  })

  it('falls back to the Graph document library drive when the configured name differs', async () => {
    const seen = []
    const service = new SharePointService({ ...config, sharePointDriveName: 'Shared Documents' }, {
      credential,
      fetchImpl: async (url) => {
        seen.push(url)
        if (url.includes('/sites/hypermineai.sharepoint.com:/sites/DocayaDocLibrary-Dev'))
          return jsonResponse({ id: 'site-id', displayName: 'Docaya site' })
        if (url.endsWith('/sites/site-id/drives'))
          return jsonResponse({
            value: [{ id: 'documents-drive-id', name: 'Documents', driveType: 'documentLibrary' }],
          })
        if (url.endsWith('/drives/documents-drive-id/root/children')) return jsonResponse({ value: [] })
        throw new Error(`Unexpected URL ${url}`)
      },
    })

    const result = await service.list()
    expect(result.drive).toMatchObject({ id: 'documents-drive-id', name: 'Documents' })
    expect(seen.some((url) => url.endsWith('/drives/documents-drive-id/root/children'))).toBe(true)
  })

  it('returns safe SharePoint diagnostics without credentials', async () => {
    const service = new SharePointService(config, {
      credential,
      fetchImpl: async (url) => {
        if (url.includes('/sites/hypermineai.sharepoint.com:/sites/DocayaDocLibrary-Dev'))
          return jsonResponse({
            id: 'site-id',
            displayName: 'Docaya Doc Library - Dev',
            webUrl: 'https://hypermineai.sharepoint.com/sites/DocayaDocLibrary-Dev',
          })
        if (url.endsWith('/sites/site-id/drives'))
          return jsonResponse({ value: [{ id: 'drive-id', name: 'Documents', driveType: 'documentLibrary' }] })
        if (url.endsWith('/drives/drive-id/root/children'))
          return jsonResponse({
            value: [{ id: 'folder-root', name: 'Document Library', folder: {}, parentReference: { path: '/drive/root:' } }],
          })
        throw new Error(`Unexpected URL ${url}`)
      },
    })

    const result = await service.debug()
    expect(result.site).toMatchObject({
      id: 'site-id',
      displayName: 'Docaya Doc Library - Dev',
      webUrl: 'https://hypermineai.sharepoint.com/sites/DocayaDocLibrary-Dev',
    })
    expect(result.drive).toMatchObject({ id: 'drive-id', name: 'Documents' })
    expect(result.drives).toEqual([expect.objectContaining({ id: 'drive-id', name: 'Documents' })])
    expect(result.rootItems).toEqual([expect.objectContaining({ id: 'folder-root', name: 'Document Library' })])
    expect(JSON.stringify(result)).not.toContain('test-token')
    expect(JSON.stringify(result)).not.toContain('client-secret')
  })

  it('reports configured but failed status when Graph denies access', async () => {
    const service = new SharePointService(config, {
      credential,
      fetchImpl: async () =>
        new Response(
          JSON.stringify({ error: { code: 'accessDenied', message: 'Denied', innerError: { 'request-id': 'graph-1' } } }),
          { status: 403, headers: { 'content-type': 'application/json' } },
        ),
    })

    const status = await service.getConnectionStatus()
    expect(status.status).toBe('connection_failed')
    expect(status.connected).toBe(false)
  })

  it('throws Graph errors with operation and request id details', async () => {
    const service = new SharePointService(config, {
      credential,
      fetchImpl: async () =>
        new Response(
          JSON.stringify({ error: { code: 'accessDenied', message: 'Denied', innerError: { 'request-id': 'graph-2' } } }),
          { status: 403, headers: { 'content-type': 'application/json' } },
        ),
    })

    await expect(service.list()).rejects.toMatchObject({
      constructor: SharePointGraphError,
      operation: 'resolve-site',
      status: 403,
      code: 'accessDenied',
      requestId: 'graph-2',
    })
  })

  it('returns not configured without credentials', async () => {
    const service = new SharePointService({ ...config, sharePointClientSecret: '' }, { credential: null })
    expect(service.configuration().configured).toBe(false)
    await expect(service.list()).rejects.toMatchObject({ code: 'SHAREPOINT_NOT_CONFIGURED' })
  })
})
