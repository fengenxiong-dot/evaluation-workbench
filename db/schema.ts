export const createAppDataTableSql = `
  CREATE TABLE IF NOT EXISTS app_data (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    data_json TEXT NOT NULL,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
  )
`;
