// ─── 額度與付款紀錄（存在 Cloudflare D1，使用者碰不到）──────────────────────
//
// accounts：每人一筆，quota = 總共可以翻幾份檔案、used = 已經用掉幾份
// ledger  ：每一次變動的紀錄（建立帳戶、翻譯扣額度、管理員加額度、記錄付款）
// jobs    ：每開始翻譯一份檔案就發一張「工作單」，之後這份檔案的每一批翻譯都要帶著它；
//           同一張工作單 12 小時內、最多 800 次呼叫，避免一份額度被拿去無限翻譯。
//
// 計費單位：一份上傳的檔案算 1 份，不管選幾種語言、輸出成一個或多個檔案。

export const DEFAULT_FREE_QUOTA = 2;
export const JOB_MAX_CALLS = 800;
export const JOB_TTL_MS = 12 * 3600_000;

export interface Credit { uid: string; email: string | null; quota: number; used: number; updated_at?: string }
export interface LedgerEntry { id: number; uid: string; kind: string; files: number; amount: number | null; note: string | null; by_email: string | null; created_at: string }

export const freeQuota = (env: { FREE_QUOTA?: string }) => {
  const n = Number(env.FREE_QUOTA);
  return Number.isFinite(n) && n >= 0 ? Math.floor(n) : DEFAULT_FREE_QUOTA;
};

const ready = new WeakSet<object>();
export async function ensureCreditTables(db: D1Database) {
  if (ready.has(db)) return;
  await db.batch([
    db.prepare('CREATE TABLE IF NOT EXISTS accounts (uid TEXT PRIMARY KEY, email TEXT, quota INTEGER NOT NULL, used INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL, updated_at TEXT NOT NULL)'),
    db.prepare('CREATE TABLE IF NOT EXISTS ledger (id INTEGER PRIMARY KEY AUTOINCREMENT, uid TEXT NOT NULL, kind TEXT NOT NULL, files INTEGER NOT NULL DEFAULT 0, amount INTEGER, note TEXT, by_email TEXT, created_at TEXT NOT NULL)'),
    db.prepare('CREATE INDEX IF NOT EXISTS ledger_uid ON ledger (uid, id)'),
    db.prepare('CREATE TABLE IF NOT EXISTS jobs (id TEXT PRIMARY KEY, uid TEXT NOT NULL, file TEXT, calls INTEGER NOT NULL DEFAULT 0, max_calls INTEGER NOT NULL, created_at INTEGER NOT NULL)'),
  ]);
  ready.add(db);
}

export async function getCredit(db: D1Database, uid: string): Promise<Credit | null> {
  await ensureCreditTables(db);
  return db.prepare('SELECT uid, email, quota, used, updated_at FROM accounts WHERE uid = ?1').bind(uid).first<Credit>();
}

// 第一次使用時建立帳戶。seed 是從舊資料帶過來的額度與已用份數；已經有帳戶就不動。
export async function createCredit(db: D1Database, uid: string, email: string | null, seed: { quota: number; used: number }, now = Date.now()): Promise<Credit> {
  await ensureCreditTables(db);
  const ts = new Date(now).toISOString();
  const ins = await db.prepare(
    'INSERT INTO accounts (uid, email, quota, used, created_at, updated_at) VALUES (?1, ?2, ?3, ?4, ?5, ?5) ON CONFLICT (uid) DO NOTHING'
  ).bind(uid, email, seed.quota, seed.used, ts).run();
  if (ins?.meta?.changes === 1) {
    await db.prepare("INSERT INTO ledger (uid, kind, files, note, created_at) VALUES (?1, 'init', ?2, ?3, ?4)")
      .bind(uid, seed.quota, seed.used ? `建立帳戶（之前已用 ${seed.used} 份）` : '建立帳戶', ts).run();
  }
  return (await getCredit(db, uid))!;
}

// 開始翻譯一份檔案：扣 1 份額度（charge = false 時不扣，給管理員用），發一張工作單
export async function startJob(db: D1Database, uid: string, file: string, charge: boolean, now = Date.now()):
  Promise<{ ok: true; jobId: string; credit: Credit | null } | { ok: false; credit: Credit | null }> {
  await ensureCreditTables(db);
  const ts = new Date(now).toISOString();
  const name = (file || '未命名檔案').slice(0, 200);
  if (charge) {
    const row = await db.prepare('UPDATE accounts SET used = used + 1, updated_at = ?2 WHERE uid = ?1 AND used < quota RETURNING used')
      .bind(uid, ts).first();
    if (!row) return { ok: false, credit: await getCredit(db, uid) };
  }
  const jobId = crypto.randomUUID();
  const stmts = [db.prepare('INSERT INTO jobs (id, uid, file, calls, max_calls, created_at) VALUES (?1, ?2, ?3, 0, ?4, ?5)').bind(jobId, uid, name, JOB_MAX_CALLS, now)];
  if (charge) stmts.push(db.prepare("INSERT INTO ledger (uid, kind, files, note, created_at) VALUES (?1, 'use', -1, ?2, ?3)").bind(uid, name, ts));
  await db.batch(stmts);
  return { ok: true, jobId, credit: await getCredit(db, uid) };
}

// 每一批翻譯：確認工作單是本人的、還沒過期、次數沒超過，順便記一次
export async function useJob(db: D1Database, uid: string, jobId: string, now = Date.now()): Promise<boolean> {
  if (!jobId) return false;
  await ensureCreditTables(db);
  const row = await db.prepare('UPDATE jobs SET calls = calls + 1 WHERE id = ?1 AND uid = ?2 AND calls < max_calls AND created_at > ?3 RETURNING calls')
    .bind(jobId, uid, now - JOB_TTL_MS).first();
  return !!row;
}

export async function pruneJobs(db: D1Database, now = Date.now()) {
  try { await db.prepare('DELETE FROM jobs WHERE created_at < ?1').bind(now - 7 * 86400_000).run(); } catch { /* 下次再清 */ }
}

// ─── 管理員用 ────────────────────────────────────────────────────────────
export async function listCredits(db: D1Database): Promise<Credit[]> {
  await ensureCreditTables(db);
  const rows = await db.prepare('SELECT uid, email, quota, used, updated_at FROM accounts ORDER BY updated_at DESC').all<Credit>();
  return rows.results ?? [];
}

export async function listLedger(db: D1Database, uid: string, limit = 50): Promise<LedgerEntry[]> {
  await ensureCreditTables(db);
  const rows = await db.prepare('SELECT * FROM ledger WHERE uid = ?1 ORDER BY id DESC LIMIT ?2').bind(uid, limit).all<LedgerEntry>();
  return rows.results ?? [];
}

export interface Adjustment {
  uid: string;
  email?: string | null;
  kind: 'grant' | 'payment' | 'adjust';   // 加額度／記錄一筆付款／直接設定總額度
  files?: number;      // grant、payment：要加幾份（可為負數，扣回）
  setQuota?: number;   // adjust：總額度直接改成多少
  amount?: number;     // payment：收了多少錢（新台幣）
  note?: string;
  seedUsed?: number;   // 這個人還沒有帳戶時，從翻譯紀錄算出的已用份數
}

export async function adjustCredit(db: D1Database, a: Adjustment, byEmail: string, freeQuotaValue: number, now = Date.now()): Promise<Credit> {
  await ensureCreditTables(db);
  if (!a.uid) throw new Error('缺少 uid');
  const ts = new Date(now).toISOString();
  let credit = await getCredit(db, a.uid);
  if (!credit) credit = await createCredit(db, a.uid, a.email ?? null, { quota: freeQuotaValue, used: Math.max(0, Math.floor(a.seedUsed ?? 0)) }, now);

  let delta: number;
  if (a.kind === 'adjust') {
    const target = Math.floor(Number(a.setQuota));
    if (!Number.isFinite(target) || target < 0) throw new Error('總額度要是 0 以上的整數');
    delta = target - credit.quota;
  } else {
    delta = Math.floor(Number(a.files ?? 0));
    if (!Number.isFinite(delta)) throw new Error('份數要是整數');
    if (a.kind === 'grant' && delta === 0) throw new Error('份數不能是 0');
  }
  const amount = a.kind === 'payment' ? Math.round(Number(a.amount) || 0) : null;
  if (a.kind === 'payment' && !(amount && amount > 0)) throw new Error('請填收款金額');
  if (credit.quota + delta < 0) throw new Error('總額度不能小於 0');

  await db.batch([
    db.prepare('UPDATE accounts SET quota = quota + ?2, email = COALESCE(?3, email), updated_at = ?4 WHERE uid = ?1').bind(a.uid, delta, a.email ?? null, ts),
    db.prepare('INSERT INTO ledger (uid, kind, files, amount, note, by_email, created_at) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)')
      .bind(a.uid, a.kind, delta, amount, (a.note || '').slice(0, 300) || null, byEmail, ts),
  ]);
  return (await getCredit(db, a.uid))!;
}
