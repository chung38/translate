// ─── 每人每日次數上限 ─────────────────────────────────────────────────────
// 一份文件會分成很多批送出，每批算 1 次（56 頁的掃描 PDF 約 100～150 次）。
// 日期用台灣時間，每天 0 點重新計算；只有放行的那次才會加 1。
// 沒有接 D1 資料庫（env.DB）或資料庫出錯時不限制。
export const DEFAULT_DAILY_LIMIT = 2000;
const readyDbs = new WeakSet<object>();   // 已確認有 api_usage 資料表的資料庫

export const taipeiDay = (now = Date.now()) => new Date(now + 8 * 3600_000).toISOString().slice(0, 10);

export async function takeDailySlot(
  env: { DB?: D1Database; DAILY_LIMIT?: string },
  uid: string,
  now = Date.now()
): Promise<{ ok: boolean; used?: number; limit: number } | null> {
  const db = env.DB;
  if (!db) return null;
  const limit = Number(env.DAILY_LIMIT) || DEFAULT_DAILY_LIMIT;
  try {
    if (!readyDbs.has(db)) {
      await db.prepare(
        'CREATE TABLE IF NOT EXISTS api_usage (uid TEXT NOT NULL, day TEXT NOT NULL, n INTEGER NOT NULL, PRIMARY KEY (uid, day))'
      ).run();
      readyDbs.add(db);
    }
    const row = await db.prepare(
      `INSERT INTO api_usage (uid, day, n) VALUES (?1, ?2, 1)
       ON CONFLICT (uid, day) DO UPDATE SET n = n + 1 WHERE n < ?3
       RETURNING n`
    ).bind(uid, taipeiDay(now), limit).first<{ n: number }>();
    return row ? { ok: true, used: row.n, limit } : { ok: false, limit };
  } catch (e) {
    readyDbs.delete(db);   // 下次重新確認資料表
    console.error('每日次數檢查失敗，本次不限制：', e);
    return null;
  }
}

// 偶爾清掉 90 天前的紀錄，資料表不會一直長大
export async function pruneUsage(db: D1Database | undefined, now = Date.now()) {
  if (!db) return;
  try {
    await db.prepare('DELETE FROM api_usage WHERE day < ?1').bind(taipeiDay(now - 90 * 86400_000)).run();
  } catch { /* 下次再清 */ }
}

