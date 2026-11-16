// Cloudflare Pages Function — /api/admin/credits：管理員查看與調整額度、記錄付款
//   GET               → 所有帳戶的額度與已用份數
//   GET ?uid=xxx      → 這個人的額度變動紀錄
//   POST { uid, kind: 'grant' | 'payment' | 'adjust', files?, setQuota?, amount?, note?, email?, seedUsed? }
import type { Env } from '../../_lib/env';
import { jsonFor, errorBody, corsHeaders } from '../../_lib/http';
import { authenticate, adminEmails } from '../../_lib/auth';
import { loadMe } from '../../_lib/account';
import { adjustCredit, freeQuota, listCredits, listLedger, type Adjustment } from '../../_lib/credits';

async function requireAdmin(request: Request, env: Env) {
  const auth = await authenticate(request, env);
  if (!('user' in auth)) return { error: auth.message, status: auth.status } as const;
  const me = await loadMe(env, auth.user, auth.idToken);
  if (!me.admin) return { error: '只有管理員可以使用。', status: 403 } as const;
  if (!env.DB) return { error: '尚未接上 D1 資料庫（Settings → Bindings，名稱 DB）。', status: 503 } as const;
  return { user: auth.user, db: env.DB } as const;
}

export const onRequestGet: PagesFunction<Env> = async ({ request, env }) => {
  const json = jsonFor(env, request);
  const a = await requireAdmin(request, env);
  if ('error' in a) return json(errorBody(a.error), a.status);
  const uid = new URL(request.url).searchParams.get('uid');
  if (uid) return json({ ledger: await listLedger(a.db, uid) });
  return json({ accounts: await listCredits(a.db), freeQuota: freeQuota(env), adminEmailsSet: adminEmails(env).length > 0 });
};

export const onRequestPost: PagesFunction<Env> = async ({ request, env }) => {
  const json = jsonFor(env, request);
  const a = await requireAdmin(request, env);
  if ('error' in a) return json(errorBody(a.error), a.status);
  let body: Adjustment;
  try { body = await request.json() as Adjustment; } catch { return json(errorBody('無效的請求內容'), 400); }
  if (!['grant', 'payment', 'adjust'].includes(body?.kind)) return json(errorBody('kind 只能是 grant、payment 或 adjust'), 400);
  try {
    const credit = await adjustCredit(a.db, body, a.user.email || a.user.uid, freeQuota(env));
    return json({ account: credit });
  } catch (e) {
    return json(errorBody((e as Error).message || '調整失敗'), 400);
  }
};

export const onRequestOptions: PagesFunction<Env> = async ({ request, env }) =>
  new Response(null, { status: 204, headers: corsHeaders(env, request) });
