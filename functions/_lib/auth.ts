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

export async function verifyFirebaseIdToken(token: string, projectId: string) {
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


// ─── 從請求取出登入的使用者 ──────────────────────────────────────────────
export interface User { uid: string; email?: string; emailVerified: boolean }

// 成功時有 user；失敗時有 status 與 message（用 'user' in result 判斷）
export type AuthResult =
  | { user: User; idToken: string }
  | { status: number; message: string };

export async function authenticate(request: Request, env: { FIREBASE_PROJECT_ID?: string }, { requireVerified = true } = {}): Promise<AuthResult> {
  if (!env.FIREBASE_PROJECT_ID) return { status: 500, message: '未設定 FIREBASE_PROJECT_ID。' };
  const header = request.headers.get('Authorization') || '';
  const idToken = header.startsWith('Bearer ') ? header.slice(7) : '';
  let user: User;
  try {
    user = await verifyFirebaseIdToken(idToken, env.FIREBASE_PROJECT_ID);
  } catch {
    return { status: 401, message: '請先登入。' };
  }
  if (requireVerified && !user.emailVerified) return { status: 403, message: '請先完成 Email 驗證。' };
  return { user, idToken };
}

// ─── 管理員 ──────────────────────────────────────────────────────────────
// 只認 Cloudflare 上 ADMIN_EMAILS 列的 Email（必須是已驗證的 Email）。
// 還沒設 ADMIN_EMAILS 時，才暫時沿用 Firestore 的 role（使用者有機會自己改，所以要盡快設定）。
export function adminEmails(env: { ADMIN_EMAILS?: string }): string[] {
  return (env.ADMIN_EMAILS || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
}

export function isAdmin(env: { ADMIN_EMAILS?: string }, user: User, firestoreRoleIsAdmin = false): boolean {
  const list = adminEmails(env);
  if (list.length) return user.emailVerified && !!user.email && list.includes(user.email.toLowerCase());
  return firestoreRoleIsAdmin;
}
