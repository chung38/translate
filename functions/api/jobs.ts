// Cloudflare Pages Function — /api/jobs：開始翻譯一份檔案
// 扣 1 份額度（管理員不扣），回傳工作單 jobId；之後這份檔案的每一批翻譯都要帶著它（X-Job-Id）。
import type { Env } from '../_lib/env';
import { jsonFor, errorBody, corsHeaders } from '../_lib/http';
import { authenticate } from '../_lib/auth';
import { loadMe, summary } from '../_lib/account';
import { pruneJobs, startJob } from '../_lib/credits';

export const onRequestPost: PagesFunction<Env> = async ({ request, env, waitUntil }) => {
  const json = jsonFor(env, request);
  const auth = await authenticate(request, env);
  if (!('user' in auth)) return json(errorBody(auth.message), auth.status);

  let file = '';
  try { file = String(((await request.json()) as { file?: string })?.file ?? ''); } catch { /* 沒帶檔名也可以 */ }

  const me = await loadMe(env, auth.user, auth.idToken);
  if (me.pendingDeletion) return json(errorBody('此帳號已停用。'), 403);
  if (!env.DB || !me.credit) return json({ jobId: null, ...summary(me) });   // 沒接 D1：照舊方式翻譯

  try {
    const r = await startJob(env.DB, auth.user.uid, file, !me.admin);
    if (Math.random() < 0.02) waitUntil(pruneJobs(env.DB));
    const s = summary({ ...me, credit: r.credit });
    if (!r.ok) return json({ ...errorBody(`額度已用完（共 ${s.quota} 份），請聯絡我們加購額度。`), ...s }, 403);
    return json({ jobId: r.jobId, ...s });
  } catch (e) {
    console.error('建立翻譯工作失敗：', e);
    return json(errorBody('暫時無法開始翻譯，請稍後再試。'), 500);
  }
};

export const onRequestOptions: PagesFunction<Env> = async ({ request, env }) =>
  new Response(null, { status: 204, headers: corsHeaders(env, request) });
