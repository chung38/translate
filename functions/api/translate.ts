// Cloudflare Pages Function — /api/translate：翻譯一批文字
//
// 每次呼叫依序檢查：
//   1. 已登入、Email 已驗證
//   2. 帳號沒被停用
//   3. 帶著有效的工作單（X-Job-Id，開始翻譯每份檔案時由 /api/jobs 發出並扣額度）
//      沒接 D1 時退回舊方式：比對 Firestore 的額度與翻譯紀錄筆數
//   4. 每人每日次數上限（管理員不限）
// 共用的程式在 functions/_lib/。
import type { Env } from '../_lib/env';
import { jsonFor, errorBody, corsHeaders } from '../_lib/http';
import { authenticate } from '../_lib/auth';
import { loadMe } from '../_lib/account';
import { freeQuota, useJob } from '../_lib/credits';
import { pruneUsage, takeDailySlot } from '../_lib/limits';
import { callModel, pickProvider } from '../_lib/model';

export const onRequestPost: PagesFunction<Env> = async (context) => {
  const { request, env } = context;
  const json = jsonFor(env, request, 'POST, OPTIONS');

  // ── 1. 身分 ────────────────────────────────────────────────────────────
  const auth = await authenticate(request, env);
  if (!('user' in auth)) {
    const message = auth.status === 401 ? '請先登入後再使用翻譯功能。' : auth.message;
    return json(errorBody(message), auth.status);
  }
  const { user, idToken } = auth;

  // ── 2. 帳號狀態 ────────────────────────────────────────────────────────
  const me = await loadMe(env, user, idToken);
  if (me.pendingDeletion) return json(errorBody('此帳號已停用。'), 403);

  const provider = pickProvider(env);
  if (!provider) {
    return json(errorBody('尚未設定翻譯服務金鑰，請到 Cloudflare → Settings → Variables and Secrets 新增 OPENAI_API_KEY。'), 500);
  }

  let prompt: string;
  try {
    const body = await request.json() as { prompt?: string };
    prompt = body?.prompt ?? '';
  } catch {
    return json(errorBody('無效的請求內容'), 400);
  }
  if (!prompt || typeof prompt !== 'string') return json(errorBody('無效的請求內容'), 400);
  if (prompt.length > 200000) return json(errorBody('文本內容過長，超過系統單次處理限制。'), 400);

  // ── 3. 額度：每份檔案要有工作單 ────────────────────────────────────────
  if (env.DB && me.credit) {
    const jobId = request.headers.get('X-Job-Id') || '';
    if (!jobId) return json(errorBody('網站已更新，請重新整理頁面後再翻譯。'), 403);
    let ok = true;
    try {
      ok = await useJob(env.DB, user.uid, jobId);
    } catch (e) {
      console.error('檢查工作單失敗，本次放行：', e);
    }
    if (!ok) return json(errorBody('這份檔案的翻譯授權已失效（超過 12 小時或次數用完），請重新按「開始翻譯」。'), 403);
  } else {
    const p = me.profile;
    const quota = p?.quota ?? freeQuota(env);
    if (p && !me.admin && p.used >= quota) {
      return json(errorBody(`額度已用完（${quota} 份檔案），請聯絡我們加購額度。`), 403);
    }
  }

  // ── 4. 每人每日次數上限（管理員不限）──────────────────────────────────
  if (!me.admin) {
    const slot = await takeDailySlot(env, user.uid);
    if (slot && !slot.ok) {
      return json(errorBody(`今天的翻譯次數已達上限（每人每天 ${slot.limit} 次），明天 0 點會重新計算。`), 429);
    }
    if (slot && Math.random() < 0.01) context.waitUntil(pruneUsage(env.DB));
  }

  // ── 翻譯 ───────────────────────────────────────────────────────────────
  const headers = { ...corsHeaders(env, request, 'POST, OPTIONS'), 'Content-Type': 'application/json', 'X-Translate-Model': provider.model };
  try {
    const upstream = await callModel(provider, prompt);
    const text = await upstream.text();
    if (upstream.ok) return new Response(text, { status: 200, headers });

    // 金鑰錯誤或沒有權限，跟「使用者沒登入」(401) 分開，免得前端誤會
    if (upstream.status === 401 || upstream.status === 403) {
      console.error(`${provider.name} 金鑰被拒絕：`, text.slice(0, 300));
      const keyName = provider.name === 'openai' ? 'OPENAI_API_KEY' : 'DEEPSEEK_API_KEY';
      return json(errorBody(`翻譯服務拒絕了金鑰，請檢查 Cloudflare 上的 ${keyName} 是否正確、帳戶是否還有餘額。`), 502);
    }
    let errorData: unknown;
    try { errorData = JSON.parse(text); }
    catch { errorData = errorBody(`翻譯服務錯誤（${upstream.status}）`); }
    return new Response(JSON.stringify(errorData), { status: upstream.status, headers });
  } catch (err: any) {
    return json(errorBody(err?.message || '翻譯請求失敗，請稍後重試。'), 500);
  }
};

export const onRequestOptions: PagesFunction<Env> = async ({ request, env }) =>
  new Response(null, { status: 204, headers: corsHeaders(env, request, 'POST, OPTIONS') });
