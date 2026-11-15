// Cloudflare Pages Function — /api/translate
// 自動對應路由：https://your-domain/api/translate

// 前端用的同一份 Firebase 設定；這裡只取 Firestore 資料庫名稱（AI Studio 建的不是預設資料庫）
import firebaseConfig from '../../firebase-applet-config.json';
const FIRESTORE_DB = (firebaseConfig as { firestoreDatabaseId?: string }).firestoreDatabaseId || '(default)';

interface Env {
  DEEPSEEK_API_KEY: string;
  FIREBASE_PROJECT_ID: string;   // Cloudflare → Settings → Variables and Secrets
  ALLOWED_ORIGIN?: string;       // 選填：例如 https://translate-1up.pages.dev，未設定則不限制
  DB?: D1Database;               // 選填：D1 資料庫（Settings → Bindings，名稱 DB），用來算每日次數
  DAILY_LIMIT?: string;          // 選填：每人每天最多呼叫幾次，預設 2000
}

// ─── 帳號狀態與檔案額度 ───────────────────────────────────────────────────
// 用使用者自己的登入 token 讀他自己的 Firestore 資料（跟網頁上看到的一樣），
// 所以不需要額外金鑰。讀不到（網路問題、資料還沒建立）就放行，不擋正常使用。
export interface Account {
  isAdmin: boolean;
  pendingDeletion: boolean;
  quota: number;   // 可翻幾份檔案
  used: number;    // 已翻幾份（翻譯紀錄筆數）
}

const ACCOUNT_TTL = 60_000;   // 同一個人 1 分鐘內只查一次，一份文件會連續呼叫很多次
const accountCache = new Map<string, { at: number; value: Account }>();

export async function getAccount(uid: string, idToken: string, projectId: string): Promise<Account | null> {
  const hit = accountCache.get(uid);
  if (hit && Date.now() - hit.at < ACCOUNT_TTL) return hit.value;

  const docUrl = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/${FIRESTORE_DB}/documents/users/${encodeURIComponent(uid)}`;
  const headers = { Authorization: `Bearer ${idToken}`, 'Content-Type': 'application/json' };
  try {
    const [docRes, countRes] = await Promise.all([
      fetch(docUrl, { headers, signal: AbortSignal.timeout(5000) }),
      fetch(`${docUrl}:runAggregationQuery`, {
        method: 'POST',
        headers,
        signal: AbortSignal.timeout(5000),
        body: JSON.stringify({
          structuredAggregationQuery: {
            structuredQuery: { from: [{ collectionId: 'history' }] },
            aggregations: [{ alias: 'n', count: {} }],
          },
        }),
      }),
    ]);
    if (!docRes.ok || !countRes.ok) return null;

    const f = ((await docRes.json()) as any)?.fields ?? {};
    const agg = (await countRes.json()) as any[];
    const used = Number(agg?.[0]?.result?.aggregateFields?.n?.integerValue);
    if (!Number.isFinite(used)) return null;

    const value: Account = {
      isAdmin: f.role?.stringValue === 'admin',
      pendingDeletion: f.isPendingDeletion?.booleanValue === true,
      quota: Number(f.quota?.integerValue ?? f.quota?.doubleValue) || 2,   // 跟網頁一樣，沒設就是 2
      used,
    };
    accountCache.set(uid, { at: Date.now(), value });
    return value;
  } catch {
    return null;
  }
}

// ─── 每人每日次數上限 ─────────────────────────────────────────────────────
// 一份文件會分成很多批送出，每批算 1 次（56 頁的掃描 PDF 約 100～150 次）。
// 日期用台灣時間，每天 0 點重新計算；只有放行的那次才會加 1。
// 沒有接 D1 資料庫（env.DB）或資料庫出錯時不限制。
export const DEFAULT_DAILY_LIMIT = 2000;
const readyDbs = new WeakSet<object>();   // 已確認有 api_usage 資料表的資料庫

export const taipeiDay = (now = Date.now()) => new Date(now + 8 * 3600_000).toISOString().slice(0, 10);

export async function takeDailySlot(
  env: Pick<Env, 'DB' | 'DAILY_LIMIT'>,
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

// ─── Firebase ID Token 驗證 ────────────────────────────────────────────────
// Workers 環境跑不了 firebase-admin，所以用 WebCrypto 自己驗 RS256。
const JWKS_URL = 'https://www.googleapis.com/service_accounts/v1/jwk/securetoken@system.gserviceaccount.com';
let jwksCache: { keys: any[] | null; exp: number } = { keys: null, exp: 0 };

const b64urlToBytes = (s: string) => {
  const b64 = s.replace(/-/g, '+').replace(/_/g, '/').padEnd(Math.ceil(s.length / 4) * 4, '=');
  return Uint8Array.from(atob(b64), c => c.charCodeAt(0));
};
const b64urlToJson = (s: string) => JSON.parse(new TextDecoder().decode(b64urlToBytes(s)));

async function getSigningKeys(): Promise<any[]> {
  const now = Date.now();
  if (jwksCache.keys && now < jwksCache.exp) return jwksCache.keys;
  const res = await fetch(JWKS_URL);
  if (!res.ok) throw new Error('無法取得 Google 公鑰');
  const jwks = await res.json() as { keys: any[] };
  const maxAge = /max-age=(\d+)/.exec(res.headers.get('cache-control') || '');
  jwksCache = { keys: jwks.keys, exp: now + (maxAge ? Number(maxAge[1]) : 3600) * 1000 };
  return jwks.keys;
}

async function verifyFirebaseIdToken(token: string, projectId: string) {
  const parts = (token || '').split('.');
  if (parts.length !== 3) throw new Error('token 格式錯誤');
  const [h, p, s] = parts;
  const header = b64urlToJson(h);
  const payload = b64urlToJson(p);

  if (header.alg !== 'RS256') throw new Error('演算法不符');
  const now = Math.floor(Date.now() / 1000);
  if (payload.aud !== projectId) throw new Error('aud 不符');
  if (payload.iss !== `https://securetoken.google.com/${projectId}`) throw new Error('iss 不符');
  if (!payload.sub) throw new Error('缺少 sub');
  if (payload.exp <= now) throw new Error('token 已過期');
  if (payload.iat > now + 300) throw new Error('iat 異常');

  const jwk = (await getSigningKeys()).find(k => k.kid === header.kid);
  if (!jwk) throw new Error('找不到對應的公鑰');

  const key = await crypto.subtle.importKey(
    'jwk',
    { kty: 'RSA', n: jwk.n, e: jwk.e, alg: 'RS256', ext: true },
    { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
    false,
    ['verify']
  );
  const ok = await crypto.subtle.verify(
    'RSASSA-PKCS1-v1_5', key, b64urlToBytes(s), new TextEncoder().encode(`${h}.${p}`)
  );
  if (!ok) throw new Error('簽章驗證失敗');

  return { uid: payload.sub as string, email: payload.email as string | undefined, emailVerified: !!payload.email_verified };
}

const buildCors = (env: Env, request: Request) => {
  const origin = request.headers.get('Origin') || '';
  const allowed = env.ALLOWED_ORIGIN;
  return {
    'Access-Control-Allow-Origin': allowed ? (origin === allowed ? origin : allowed) : '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Vary': 'Origin',
  };
};

export const onRequestPost: PagesFunction<Env> = async (context) => {
  const { request, env } = context;
  const corsHeaders = buildCors(env, request);
  const json = (body: unknown, status: number) =>
    new Response(JSON.stringify(body), { status, headers: { ...corsHeaders, 'Content-Type': 'application/json' } });

  // ── 驗證身分：沒有這一段，任何人都能拿這個網址當免費翻譯 proxy ──────────
  if (!env.FIREBASE_PROJECT_ID) {
    return json({ error: { message: '未設定 FIREBASE_PROJECT_ID。' } }, 500);
  }
  const authHeader = request.headers.get('Authorization') || '';
  const idToken = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : '';
  let user: { uid: string; email?: string; emailVerified: boolean };
  try {
    user = await verifyFirebaseIdToken(idToken, env.FIREBASE_PROJECT_ID);
  } catch {
    return json({ error: { message: '請先登入後再使用翻譯功能。' } }, 401);
  }
  if (!user.emailVerified) {
    return json({ error: { message: '請先完成 Email 驗證。' } }, 403);
  }

  // ── 帳號狀態與額度：跟網頁上的檢查一樣，但改在這裡擋，繞過網頁也沒用 ─────
  const account = await getAccount(user.uid, idToken, env.FIREBASE_PROJECT_ID);
  if (account?.pendingDeletion) {
    return json({ error: { message: '此帳號已停用。' } }, 403);
  }
  if (account && !account.isAdmin && account.used >= account.quota) {
    return json({ error: { message: `額度已用完（${account.quota} 份檔案），請聯絡管理員增加額度。` } }, 403);
  }

  // ── 每人每日次數上限（管理員不限）────────────────────────────────────────
  if (!account?.isAdmin) {
    const slot = await takeDailySlot(env, user.uid);
    if (slot && !slot.ok) {
      return json({ error: { message: `今天的翻譯次數已達上限（每人每天 ${slot.limit} 次），明天 0 點會重新計算。` } }, 429);
    }
    if (slot && Math.random() < 0.01) context.waitUntil(pruneUsage(env.DB));
  }

  const apiKey = env.DEEPSEEK_API_KEY;
  if (!apiKey) {
    return json({ error: { message: '未設定 DEEPSEEK_API_KEY，請至 Cloudflare Pages → Settings → Environment variables 新增。' } }, 500);
  }

  let prompt: string;
  try {
    const body = await request.json() as { prompt?: string };
    prompt = body?.prompt ?? '';
  } catch {
    return json({ error: { message: '無效的請求內容' } }, 400);
  }

  if (!prompt || typeof prompt !== 'string') {
    return json({ error: { message: '無效的請求內容' } }, 400);
  }
  if (prompt.length > 200000) {
    return json({ error: { message: '文本內容過長，超過系統單次處理限制。' } }, 400);
  }

  // 呼叫 DeepSeek，最多重試 3 次
  let retries = 3;
  let delay = 1000;

  while (retries > 0) {
    try {
      const response = await fetch('https://api.deepseek.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${apiKey}`,
        },
        body: JSON.stringify({
          model: 'deepseek-chat',
          messages: [
            { role: 'system', content: 'You are a helpful assistant that translates text into multiple languages and outputs only structured JSON.' },
            { role: 'user', content: prompt },
          ],
          response_format: { type: 'json_object' },
          temperature: 0.3,
          max_tokens: 8192,
        }),
      });

      const responseText = await response.text();

      if (!response.ok) {
        if ((response.status === 429 || response.status >= 500) && retries > 1) {
          await new Promise(r => setTimeout(r, delay));
          retries--;
          delay *= 2;
          continue;
        }
        let errorData;
        try { errorData = JSON.parse(responseText); }
        catch { errorData = { error: { message: `DeepSeek API Error (${response.status})` } }; }
        return new Response(JSON.stringify(errorData), {
          status: response.status,
          headers: { ...corsHeaders, 'Content-Type': 'application/json' },
        });
      }

      return new Response(responseText, {
        status: 200,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });

    } catch (err: any) {
      if (retries > 1) {
        await new Promise(r => setTimeout(r, delay));
        retries--;
        delay *= 2;
        continue;
      }
      return json({ error: { message: err.message || 'Internal Server Error' } }, 500);
    }
  }

  return json({ error: { message: '翻譯請求失敗，請稍後重試。' } }, 500);
};

// 同時處理 OPTIONS preflight
export const onRequestOptions: PagesFunction<Env> = async (context) => {
  return new Response(null, { status: 204, headers: buildCors(context.env, context.request) });
};
