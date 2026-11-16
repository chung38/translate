// Cloudflare Pages Function — /api/me：目前登入者的額度（網頁右上角、開始翻譯前的檢查用）
import type { Env } from '../_lib/env';
import { jsonFor, errorBody, corsHeaders } from '../_lib/http';
import { authenticate } from '../_lib/auth';
import { loadMe, summary } from '../_lib/account';

export const onRequestGet: PagesFunction<Env> = async ({ request, env }) => {
  const json = jsonFor(env, request);
  const auth = await authenticate(request, env, { requireVerified: false });
  if (!('user' in auth)) return json(errorBody(auth.message), auth.status);
  const me = await loadMe(env, auth.user, auth.idToken);
  return json(summary(me));
};

export const onRequestOptions: PagesFunction<Env> = async ({ request, env }) =>
  new Response(null, { status: 204, headers: corsHeaders(env, request) });
