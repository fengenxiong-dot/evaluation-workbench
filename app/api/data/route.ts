import { env } from 'cloudflare:workers';
import { createAppDataTableSql } from '@/db/schema';

type StoredData = {
  templates: unknown[];
  tasks: unknown[];
};

function getDatabase() {
  return (env as unknown as { DB: D1Database }).DB;
}

async function ensureSchema() {
  await getDatabase().prepare(createAppDataTableSql).run();
}

function isStoredData(value: unknown): value is StoredData {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as StoredData;
  return Array.isArray(candidate.templates) && Array.isArray(candidate.tasks);
}

export async function GET() {
  await ensureSchema();
  const row = await getDatabase()
    .prepare('SELECT data_json, updated_at FROM app_data WHERE id = 1')
    .first<{ data_json: string; updated_at: string }>();

  if (!row) return Response.json({ data: null, updatedAt: null });

  try {
    return Response.json({
      data: JSON.parse(row.data_json),
      updatedAt: row.updated_at,
    });
  } catch {
    return Response.json({ error: '服务器数据格式异常' }, { status: 500 });
  }
}

export async function PUT(request: Request) {
  const body = await request.json().catch(() => null);
  const incoming =
    body && typeof body === 'object' && 'data' in body
      ? (body as { data: unknown }).data
      : body;

  if (!isStoredData(incoming)) {
    return Response.json({ error: '提交的数据结构无效' }, { status: 400 });
  }

  await ensureSchema();
  await getDatabase()
    .prepare(
      `INSERT INTO app_data (id, data_json, updated_at)
     VALUES (1, ?, CURRENT_TIMESTAMP)
     ON CONFLICT(id) DO UPDATE SET
       data_json = excluded.data_json,
       updated_at = CURRENT_TIMESTAMP`,
    )
    .bind(JSON.stringify(incoming))
    .run();

  return Response.json({ ok: true });
}
