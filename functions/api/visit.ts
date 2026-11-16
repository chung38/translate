// Cloudflare Pages Function — /api/visit：網站瀏覽人數
//
// POST：記一次瀏覽，回傳最新人數（網頁每次開啟時呼叫）
// GET ：只讀人數
//
// 「人數」= 不重複訪客：同一天、同一個 IP + 同一個瀏覽器只算 1 人。
// 不存 IP，只存「日期＋IP＋瀏覽器」的雜湊值，而且隔天就刪；也不用 cookie。
// 資料記在 D1 資料庫（綁定名稱 DB，跟翻譯的每日次數共用）。沒接 D1 時回 { enabled: false }，網頁就不顯示。

interface Env {
  DB?: D1Database;
  VISIT_SALT?: string;   // 選填：雜湊用的隨機字串，讓別人無法反推 IP
}

export interface VisitStats {
  enabled: boolean;
  today?: number;        // 今日不重複訪客
  total?: number;        // 累計人次（每人每天算一次）
  viewsToday?: number;   // 今日瀏覽次數（每次開啟都算）
  viewsTotal?: number;   // 累計瀏覽次數
}

// 搜尋引擎、連結預覽、監控程式不算
const BOT = /bot|crawl|spider|slurp|preview|facebookexternalhit|headless|lighthouse|monitor|curl|wget|python|axios|node-fetch|go-http/i;

export const taipeiDay = (now = Date.now()) => new Date(now + 8 * 3600_000).toISOString().slice(0, 10);

const readyDbs = new WeakSet<object>();
async function ensureTables(db: D1Database) {
  if (readyDbs.has(db)) return;
  await db.batch([
    db.prepare('CREATE TABLE IF NOT EXISTS site_visitors (day TEXT NOT NULL, vid TEXT NOT NULL, PRIMARY KEY (day, vid))'),
    db.prepare('CREATE TABLE IF NOT EXISTS site_stats (k TEXT PRIMARY KEY, n INTEGER NOT NULL)'),
  ]);
  readyDbs.add(db);
}

async function visitorId(day: string, ip: string, ua: string, salt: string) {
  const bytes = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(`${salt}|${day}|${ip}|${ua}`));
  return [...new Uint8Array(bytes)].slice(0, 16).map(b => b.toString(16).padStart(2, '0')).join('');
}

export async function recordVisit(db: D1Database, day: string, vid: string) {
  await ensureTables(db);
  const ins = await db.prepare('INSERT INTO site_visitors (day, vid) VALUES (?1, ?2) ON CONFLICT DO NOTHING').bind(day, vid).run();
  const isNew = ins?.meta?.changes === 1;
  const keys = [`views:${day}`, 'views:total', ...(isNew ? [`visitors:${day}`, 'visitors:total'] : [])];
  await db.batch(keys.map(k =>
    db.prepare('INSERT INTO site_stats (k, n) VALUES (?1, 1) ON CONFLICT (k) DO UPDATE SET n = n + 1').bind(k)
  ));
}

export async function readStats(db: D1Database, day: string): Promise<VisitStats> {
  await ensureTables(db);
  const row = await db.prepare(
    `SELECT
       COALESCE(MAX(CASE WHEN k = ?1 THEN n END), 0) AS today,
       COALESCE(MAX(CASE WHEN k = 'visitors:total' THEN n END), 0) AS total,
       COALESCE(MAX(CASE WHEN k = ?2 THEN n END), 0) AS viewsToday,
       COALESCE(MAX(CASE WHEN k = 'views:total' THEN n END), 0) AS viewsTotal
     FROM site_stats WHERE k IN (?1, ?2, 'visitors:total', 'views:total')`
  ).bind(`visitors:${day}`, `views:${day}`).first<Omit<VisitStats, 'enabled'>>();
  return { enabled: true, today: 0, total: 0, viewsToday: 0, viewsTotal: 0, ...(row ?? {}) };
}

// 只需要今天的訪客名單來判斷重複，前幾天的直接刪掉
export async function pruneVisitors(db: D1Database, now = Date.now()) {
  try {
    await db.prepare('DELETE FROM site_visitors WHERE day < ?1').bind(taipeiDay(now - 86400_000)).run();
  } catch { /* 下次再清 */ }
}

const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' } });

export const onRequestGet: PagesFunction<Env> = async ({ env }) => {
  if (!env.DB) return json({ enabled: false });
  try {
    return json(await readStats(env.DB, taipeiDay()));
  } catch (e) {
    console.error('讀取瀏覽人數失敗：', e);
    return json({ enabled: false });
  }
};

export const onRequestPost: PagesFunction<Env> = async ({ request, env, waitUntil }) => {
  const db = env.DB;
  if (!db) return json({ enabled: false });

  const day = taipeiDay();
  const ua = request.headers.get('user-agent') || '';
  const ip = request.headers.get('cf-connecting-ip') || '';
  // 只算從自己網站發出的請求，其他網站或程式直接打 API 不算
  const origin = request.headers.get('origin');
  const sameSite = !origin || new URL(origin).host === new URL(request.url).host;

  try {
    if (ua && !BOT.test(ua) && sameSite) {
      await recordVisit(db, day, await visitorId(day, ip, ua, env.VISIT_SALT || 'yici-gaoding'));
      if (Math.random() < 0.02) waitUntil(pruneVisitors(db));
    }
    return json(await readStats(db, day));
  } catch (e) {
    console.error('記錄瀏覽人數失敗：', e);
    return json({ enabled: false });
  }
};
