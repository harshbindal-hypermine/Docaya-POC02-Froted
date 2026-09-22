CREATE TABLE IF NOT EXISTS department_sharepoint_folder_access (
  id TEXT PRIMARY KEY,
  tenant_id TEXT NOT NULL,
  department TEXT NOT NULL,
  sharepoint_folder_id TEXT NOT NULL,
  sharepoint_folder_name TEXT NOT NULL,
  sharepoint_folder_path TEXT NOT NULL,
  sharepoint_web_url TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  UNIQUE(tenant_id,department,sharepoint_folder_id)
);
CREATE INDEX IF NOT EXISTS idx_department_sharepoint_access_department ON department_sharepoint_folder_access(tenant_id,department);
CREATE INDEX IF NOT EXISTS idx_department_sharepoint_access_folder ON department_sharepoint_folder_access(tenant_id,sharepoint_folder_id);
