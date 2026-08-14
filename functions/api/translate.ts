// Cloudflare Pages Function — /api/translate
// 自動對應路由：https://your-domain/api/translate

interface Env {
  DEEPSEEK_API_KEY: string;
  FIREBASE_PROJECT_ID: string;   // ← 新增：到 Pages → Settings → Environment variables 設定
  ALLOWED_ORIGIN?: string;       // ← 新增：例如 https://your-domain.pages.dev，未設定則不限制
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
