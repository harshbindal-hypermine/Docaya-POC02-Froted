export type AppPage = 'home' | 'documents' | 'workflows' | 'notifications' | 'admin'

export type Locale = 'en' | 'ar'

export type UserProfile = {
  id?: string
  name: string
  email: string
  initials: string
  role: string
  department: string
}

export type DocumentStatus =
  | 'Draft'
  | 'Under Review'
  | 'Under Approval'
  | 'Published'
  | 'Superseded'
  | 'Archived'

export type Classification = 'Public' | 'Internal' | 'Confidential' | 'Restricted'

export type ControlledDocument = {
  id: string
  number: string
  title: string
  type: string
  library: string
  department: string
  language: string
  sourcePath: string
  owner: string
  classification: Classification
  retention: string
  revision: number
  status: DocumentStatus
  nextReview: string
  updatedAt: string
  summary: string
  workflowStep: number
  reviewer: string
  approver: string
  sharePointFolderId?: string | null
  sharePointFolderName?: string | null
  sharePointFolderPath?: string | null
}

export type WorkflowStatus = 'Pending' | 'Approved' | 'Changes requested' | 'Rejected'

export type WorkflowTask = {
  id: string
  documentId: string
  documentNumber: string
  documentTitle: string
  stage: string
  assignee: string
  requestedBy: string
  dueDate: string
  priority: 'Normal' | 'High' | 'Critical'
  status: WorkflowStatus
}

export type ActivityItem = {
  id: string
  title: string
  detail: string
  time: string
  tone: 'gold' | 'green' | 'red' | 'blue'
}

export type RegistrationInput = Omit<
  ControlledDocument,
  'id' | 'number' | 'revision' | 'status' | 'updatedAt' | 'workflowStep'
> & {
  mimeType?: string
  sizeBytes?: number
}

export type SharePointItemBase = {
  id: string
  name: string
  type: 'file' | 'folder'
  path: string
  webUrl: string | null
  lastModifiedDateTime: string | null
  parentReference?: { driveId?: string; id?: string; path?: string }
}

export type SharePointFolder = SharePointItemBase & {
  type: 'folder'
  size: null
  mimeType: null
}

export type DepartmentSharePointFolder = SharePointFolder & {
  mappingId?: string
  department: string
  sharePointFolderId: string
  createdAt?: string
  updatedAt?: string
}

export type SharePointFile = SharePointItemBase & {
  type: 'file'
  size: number
  mimeType: string | null
}

export type SharePointItem = SharePointFolder | SharePointFile

export type SharePointConnectionStatus = {
  configured: boolean
  connected: boolean
  status: 'connected' | 'not_configured' | 'connection_failed'
  site: { hostname: string; path: string; displayName?: string | null; webUrl?: string | null } | null
  drive: { name: string; webUrl?: string | null } | null
}

export type SharePointExplorerResponse = {
  items: SharePointItem[]
  folderId: string | null
  drive: { name: string; webUrl?: string | null }
  site: { hostname: string; path: string }
}

export type SharePointDepartmentFoldersResponse = {
  department: string
  items: DepartmentSharePointFolder[]
  folders: DepartmentSharePointFolder[]
  drive: { name: string; webUrl?: string | null }
  site: { hostname: string; path: string }
}

export type SharePointUploadResponse = { item: SharePointItem }
