// 用 Node 內建的 SQLite 模擬 Cloudflare D1（D1 本身就是 SQLite），給測試用
import { DatabaseSync } from 'node:sqlite';

export function fakeD1() {
  const sqlite = new DatabaseSync(':memory:');
  const stmt = (sql: string) => {
    let values: unknown[] = [];
    const s = {
      bind(...v: unknown[]) { values = v; return s; },
      async first<T>() { return (sqlite.prepare(sql).get(...(values as any[])) ?? null) as T | null; },
      async all<T>() { return { results: sqlite.prepare(sql).all(...(values as any[])) as T[] }; },
      async run() {
        const r = sqlite.prepare(sql).run(...(values as any[]));
        return { meta: { changes: Number(r.changes) } };
      },
    };
    return s;
  };
  const db = {
    prepare: stmt,
    async batch(list: ReturnType<typeof stmt>[]) {
      const out = [];
      for (const s of list) out.push(await s.run());
      return out;
    },
  };
  return { db: db as unknown as D1Database, sqlite };
}
