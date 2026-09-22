import { resolve } from 'node:path'

const integer = (name, fallback) => {
  const value = Number(process.env[name] ?? fallback)
  if (!Number.isSafeInteger(value) || value <= 0) throw new Error(`${name} must be a positive integer.`)
  return value
}

const csv = (value) =>
  String(value || '')
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean)

const sharePointSiteUrlParts = (value) => {
  if (!value) return { hostname: '', sitePath: '' }
  try {
    const url = new URL(value)
    const sitePath = url.pathname.match(/^(\/sites\/[^/]+)/i)?.[1] || ''
    return { hostname: url.hostname, sitePath }
  } catch {
    return { hostname: '', sitePath: '' }
  }
}

export function loadConfig(overrides = {}) {
  const environment = overrides.environment || process.env.DOCAYA_ENV || 'production'
  const development = environment === 'development' || environment === 'test'
  const legacySharePointUrl = sharePointSiteUrlParts(
    overrides.sharePointSiteUrl || process.env.SHAREPOINT_SITE_URL || '',
  )
  const config = {
    environment,
    port: Number(overrides.port ?? process.env.PORT ?? 8787),
    authMode: overrides.authMode || process.env.AUTH_MODE || (development ? 'development' : 'entra'),
    devTokenSecret:
      overrides.devTokenSecret ||
      process.env.DEV_TOKEN_SECRET ||
      (development ? 'docaya-development-secret-change-before-sharing' : ''),
    entraTenantId: overrides.entraTenantId || process.env.ENTRA_TENANT_ID || '',
    entraClientId: overrides.entraClientId || process.env.ENTRA_CLIENT_ID || '',
    corsOrigins:
      overrides.corsOrigins ||
      csv(process.env.CORS_ALLOWED_ORIGINS || (development ? 'http://localhost:5173' : '')),
    bodyLimit: overrides.bodyLimit || integer('REQUEST_BODY_LIMIT_BYTES', 1_048_576),
    rateLimit: overrides.rateLimit || integer('RATE_LIMIT_REQUESTS_PER_MINUTE', 120),
    readinessTimeoutMs: overrides.readinessTimeoutMs || integer('READINESS_TIMEOUT_MS', 2_000),
    uploadMaxBytes: overrides.uploadMaxBytes || integer('UPLOAD_MAX_BYTES', 104_857_600),
    uploadExtensions:
      overrides.uploadExtensions ||
      csv(process.env.UPLOAD_ALLOWED_EXTENSIONS || 'pdf,docx,xlsx,pptx,jpg,jpeg,png,txt'),
    uploadMimeTypes:
      overrides.uploadMimeTypes ||
      csv(
        process.env.UPLOAD_ALLOWED_MIME_TYPES ||
          'application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.openxmlformats-officedocument.presentationml.presentation,image/jpeg,image/png,text/plain',
      ),
    localSqlite: overrides.localSqlite ?? process.env.LOCAL_SQLITE_ENABLED === 'true',
    sqlitePath: resolve(overrides.sqlitePath || process.env.SQLITE_PATH || 'data/docaya-v1.sqlite'),
    uploadPath: resolve(overrides.uploadPath || process.env.UPLOAD_PATH || 'uploads'),
    postgresUrl: overrides.postgresUrl || process.env.DATABASE_URL || process.env.POSTGRES_URL || '',
    searchEndpoint: overrides.searchEndpoint || process.env.AZURE_SEARCH_ENDPOINT || '',
    searchIndex: overrides.searchIndex || process.env.AZURE_SEARCH_INDEX || '',
    searchApiKey: overrides.searchApiKey || process.env.AZURE_SEARCH_API_KEY || '',
    searchManagedIdentity:
      overrides.searchManagedIdentity ?? process.env.AZURE_SEARCH_USE_MANAGED_IDENTITY === 'true',
    redisUrl: overrides.redisUrl || process.env.REDIS_URL || '',
    redisToken: overrides.redisToken || process.env.REDIS_TOKEN || '',
    serviceBusNamespace: overrides.serviceBusNamespace || process.env.SERVICE_BUS_NAMESPACE || '',
    serviceBusTopic: overrides.serviceBusTopic || process.env.SERVICE_BUS_TOPIC || '',
    sharePointTenantId:
      overrides.sharePointTenantId || process.env.TENANT_ID || process.env.ENTRA_TENANT_ID || '',
    sharePointClientId:
      overrides.sharePointClientId || process.env.CLIENT_ID || process.env.ENTRA_CLIENT_ID || '',
    sharePointClientSecret: overrides.sharePointClientSecret || process.env.CLIENT_SECRET || '',
    sharePointHostname:
      overrides.sharePointHostname || process.env.SHAREPOINT_HOSTNAME || legacySharePointUrl.hostname || '',
    sharePointSitePath:
      overrides.sharePointSitePath || process.env.SHAREPOINT_SITE_PATH || legacySharePointUrl.sitePath || '',
    sharePointDriveName: overrides.sharePointDriveName || process.env.SHAREPOINT_DRIVE_NAME || '',
    sharePointUploadFolderName:
      overrides.sharePointUploadFolderName || process.env.SHAREPOINT_UPLOAD_FOLDER_NAME || 'Docaya Uploads',
    sharePointSiteUrl: overrides.sharePointSiteUrl || process.env.SHAREPOINT_SITE_URL || '',
    sharePointDriveId: overrides.sharePointDriveId || process.env.SHAREPOINT_DRIVE_ID || '',
    sharePointUseManagedIdentity:
      overrides.sharePointUseManagedIdentity ??
      overrides.managedIdentity ??
      process.env.SHAREPOINT_USE_MANAGED_IDENTITY === 'true',
    managedIdentity:
      overrides.managedIdentity ??
      overrides.sharePointUseManagedIdentity ??
      process.env.SHAREPOINT_USE_MANAGED_IDENTITY === 'true',
    openaiEndpoint: overrides.openaiEndpoint || process.env.AZURE_OPENAI_ENDPOINT || '',
    openaiDeployment: overrides.openaiDeployment || process.env.AZURE_OPENAI_DEPLOYMENT || '',
    openaiApiKey: overrides.openaiApiKey || process.env.AZURE_OPENAI_API_KEY || '',
    openaiApiVersion:
      overrides.openaiApiVersion || process.env.AZURE_OPENAI_API_VERSION || '2025-01-01-preview',
  }

  if (!Number.isInteger(config.port) || config.port < 1 || config.port > 65535)
    throw new Error('PORT is invalid.')
  if (!['development', 'entra'].includes(config.authMode))
    throw new Error('AUTH_MODE must be development or entra.')
  if (config.authMode === 'development' && !development)
    throw new Error('Development authentication is prohibited outside development/test.')
  if (config.authMode === 'development' && config.devTokenSecret.length < 32)
    throw new Error('DEV_TOKEN_SECRET must contain at least 32 characters.')
  if (config.authMode === 'entra' && (!config.entraTenantId || !config.entraClientId))
    throw new Error('Entra tenant and client IDs are required.')
  if (!development && config.corsOrigins.length === 0)
    throw new Error('A production CORS allowlist is required.')
  if (config.corsOrigins.includes('*')) throw new Error('Wildcard CORS origins are prohibited.')
  if (!development && (config.localSqlite || !config.postgresUrl))
    throw new Error('Production requires PostgreSQL and prohibits local SQLite.')
  if (
    config.sharePointUseManagedIdentity &&
    !(config.sharePointHostname && config.sharePointSitePath && config.sharePointDriveName) &&
    (!config.sharePointSiteUrl || !config.sharePointDriveId)
  )
    throw new Error(
      'Managed-identity SharePoint mode requires explicit SharePoint hostname, site path, and drive name, or a SharePoint site URL and drive ID.',
    )
  return Object.freeze(config)
}
