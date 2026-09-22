CREATE TABLE IF NOT EXISTS department_sharepoint_folder_access (
  id text PRIMARY KEY,
  tenant_id text NOT NULL,
  department text NOT NULL,
  sharepoint_folder_id text NOT NULL,
  sharepoint_folder_name text NOT NULL,
  sharepoint_folder_path text NOT NULL,
  sharepoint_web_url text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE(tenant_id,department,sharepoint_folder_id)
);
CREATE INDEX IF NOT EXISTS idx_department_sharepoint_access_department ON department_sharepoint_folder_access(tenant_id,department);
CREATE INDEX IF NOT EXISTS idx_department_sharepoint_access_folder ON department_sharepoint_folder_access(tenant_id,sharepoint_folder_id);
