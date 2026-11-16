// Cloudflare Pages Function — /api/translate
// 自動對應路由：https://your-domain/api/translate

// 前端用的同一份 Firebase 設定；這裡只取 Firestore 資料庫名稱（AI Studio 建的不是預設資料庫）
import firebaseConfig from '../../firebase-applet-config.json';
const FIRESTORE_DB = (firebaseConfig as { firestoreDatabaseId?: string }).firestoreDatabaseId || '(default)';

interface Env {
  OPENAI_API_KEY?: string;       // 有設就用 OpenAI 翻譯（預設模型 gpt-6-luna）
  OPENAI_MODEL?: string;         // 選填：換別的 OpenAI 模型
  DEEPSEEK_API_KEY?: string;     // 備用：沒設 OPENAI_API_KEY 時改用 DeepSeek
  FIREBASE_PROJECT_ID: string;   // Cloudflare → Settings → Variables and Secrets
  ALLOWED_ORIGIN?: string;       // 選填：允許從哪些網址呼叫（逗號分隔），例如 https://yicipass.com；自己網站一律允許
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

// ─── 翻譯模型 ─────────────────────────────────────────────────────────────
// 有設 OPENAI_API_KEY 就用 OpenAI（預設 gpt-6-luna），沒有才用 DeepSeek。
// 兩家都是同一種 Chat Completions 格式，回傳內容前端不用改。
export const DEFAULT_OPENAI_MODEL = 'gpt-6-luna';
const SYSTEM_PROMPT = 'You are a helpful assistant that translates text into multiple languages and outputs only structured JSON.';

export interface Provider {
  name: 'openai' | 'deepseek';
  url: string;
  key: string;
  model: string;
  options: Record<string, unknown>;   // 除了 model、messages 以外的參數
}

export function pickProvider(env: Pick<Env, 'OPENAI_API_KEY' | 'OPENAI_MODEL' | 'DEEPSEEK_API_KEY'>): Provider | null {
  if (env.OPENAI_API_KEY) {
    return {
      name: 'openai',
      url: 'https://api.openai.com/v1/chat/completions',
      key: env.OPENAI_API_KEY,
      model: env.OPENAI_MODEL || DEFAULT_OPENAI_MODEL,
      options: {
        reasoning_effort: 'none',   // 翻譯不需要「思考」；開著會變慢，推理用的 token 也照輸出計費
        temperature: 0.3,           // 只有 reasoning_effort 為 none 時才能設
        response_format: { type: 'json_object' },
        max_completion_tokens: 8192,
      },
    };
  }
  if (env.DEEPSEEK_API_KEY) {
    return {
      name: 'deepseek',
      url: 'https://api.deepseek.com/v1/chat/completions',
      key: env.DEEPSEEK_API_KEY,
      model: 'deepseek-chat',
      options: { response_format: { type: 'json_object' }, temperature: 0.3, max_tokens: 8192 },
    };
  }
  return null;
}

// 模型不接受某個選填參數（回 400）時，自動調整後重送，並記住，之後的請求直接用調整後的設定。
const ADJUSTABLE = ['temperature', 'reasoning_effort', 'response_format'] as const;
const learned = new Map<string, Record<string, unknown>>();   // 模型名稱 → 調整後的參數

export function adjustForError(options: Record<string, unknown>, status: number, errorBody: any): Record<string, unknown> | null {
  if (status !== 400) return null;
  const param = String(errorBody?.error?.param ?? '');
  const message = String(errorBody?.error?.message ?? '');
  const bad = ADJUSTABLE.find(k => k in options && (param === k || message.includes(k)));
  if (!bad) return null;
  const next = { ...options };
  if (bad === 'reasoning_effort' && next.reasoning_effort === 'none') {
    next.reasoning_effort = 'low';   // 不支援 none 的模型，用最低的 low
  } else {
    delete next[bad];
    if (bad === 'reasoning_effort') delete next.temperature;   // 有在推理就不能設溫度
  }
  return next;
}

export async function callModel(provider: Provider, prompt: string): Promise<Response> {
  let options = learned.get(provider.model) ?? provider.options;
  let adjustments = 0;
  let retries = 3;
  let delay = 1000;

  while (true) {
    let response: Response;
    try {
      response = await fetch(provider.url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${provider.key}` },
        body: JSON.stringify({
          model: provider.model,
          messages: [
            { role: 'system', content: SYSTEM_PROMPT },
            { role: 'user', content: prompt },
          ],
          ...options,
        }),
      });
    } catch (err) {
      if (--retries > 0) {
        await new Promise(r => setTimeout(r, delay));
        delay *= 2;
        continue;
      }
      throw err;
    }

    if (response.ok) return response;

    const text = await response.text();
    let body: any = null;
    try { body = JSON.parse(text); } catch { /* 不是 JSON */ }

    const fixed = adjustments < ADJUSTABLE.length ? adjustForError(options, response.status, body) : null;
    if (fixed) {
      console.warn(`${provider.model} 不接受目前的參數，已自動調整：`, body?.error?.message);
      options = fixed;
      learned.set(provider.model, fixed);
      adjustments++;
      continue;
    }

    if ((response.status === 429 || response.status >= 500) && --retries > 0) {
      await new Promise(r => setTimeout(r, delay));
      delay *= 2;
      continue;
    }

    return new Response(text, { status: response.status, headers: { 'Content-Type': 'application/json' } });
  }
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
  const list = (env.ALLOWED_ORIGIN || '').split(',').map(s => s.trim()).filter(Boolean);
  let sameSite = false;
  try { sameSite = !!origin && new URL(origin).host === new URL(request.url).host; } catch { /* Origin 格式不對 */ }
  return {
    'Access-Control-Allow-Origin': !list.length ? '*' : (sameSite || list.includes(origin)) ? origin : list[0],
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

  const provider = pickProvider(env);
  if (!provider) {
    return json({ error: { message: '尚未設定翻譯服務金鑰，請到 Cloudflare → Settings → Variables and Secrets 新增 OPENAI_API_KEY。' } }, 500);
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

  // ── 每人每日次數上限（管理員不限）────────────────────────────────────────
  if (!account?.isAdmin) {
    const slot = await takeDailySlot(env, user.uid);
    if (slot && !slot.ok) {
      return json({ error: { message: `今天的翻譯次數已達上限（每人每天 ${slot.limit} 次），明天 0 點會重新計算。` } }, 429);
    }
    if (slot && Math.random() < 0.01) context.waitUntil(pruneUsage(env.DB));
  }

  const headers = { ...corsHeaders, 'Content-Type': 'application/json', 'X-Translate-Model': provider.model };
  try {
    const upstream = await callModel(provider, prompt);
    const text = await upstream.text();
    if (upstream.ok) return new Response(text, { status: 200, headers });

    // 金鑰錯誤或沒有權限，跟「使用者沒登入」(401) 分開，免得前端誤會
    if (upstream.status === 401 || upstream.status === 403) {
      console.error(`${provider.name} 金鑰被拒絕：`, text.slice(0, 300));
      const keyName = provider.name === 'openai' ? 'OPENAI_API_KEY' : 'DEEPSEEK_API_KEY';
      return json({ error: { message: `翻譯服務拒絕了金鑰，請檢查 Cloudflare 上的 ${keyName} 是否正確、帳戶是否還有餘額。` } }, 502);
    }
    let errorData: unknown;
    try { errorData = JSON.parse(text); }
    catch { errorData = { error: { message: `翻譯服務錯誤（${upstream.status}）` } }; }
    return new Response(JSON.stringify(errorData), { status: upstream.status, headers });
  } catch (err: any) {
    return json({ error: { message: err?.message || '翻譯請求失敗，請稍後重試。' } }, 500);
  }
};

// 同時處理 OPTIONS preflight
export const onRequestOptions: PagesFunction<Env> = async (context) => {
  return new Response(null, { status: 204, headers: buildCors(context.env, context.request) });
};
