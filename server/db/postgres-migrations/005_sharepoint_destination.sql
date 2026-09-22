BEGIN;
ALTER TABLE documents ADD COLUMN IF NOT EXISTS sharepoint_folder_id text;
ALTER TABLE documents ADD COLUMN IF NOT EXISTS sharepoint_folder_name text;
ALTER TABLE documents ADD COLUMN IF NOT EXISTS sharepoint_folder_path text;
INSERT INTO schema_migrations(version) VALUES ('005_sharepoint_destination.sql') ON CONFLICT DO NOTHING;
COMMIT;
