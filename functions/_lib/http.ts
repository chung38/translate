import type { Env } from './env';

// CORS：自己網站發出的請求一律允許；ALLOWED_ORIGIN 可再加其他網站（逗號分隔）
export function corsHeaders(env: Env, request: Request, methods = 'GET, POST, OPTIONS'): Record<string, string> {
  const origin = request.headers.get('Origin') || '';
  const list = (env.ALLOWED_ORIGIN || '').split(',').map(s => s.trim()).filter(Boolean);
  let sameSite = false;
  try { sameSite = !!origin && new URL(origin).host === new URL(request.url).host; } catch { /* Origin 格式不對 */ }
  return {
    'Access-Control-Allow-Origin': !list.length ? '*' : (sameSite || list.includes(origin)) ? origin : list[0],
    'Access-Control-Allow-Methods': methods,
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Job-Id',
    'Vary': 'Origin',
  };
}

export function jsonFor(env: Env, request: Request, methods?: string) {
  const cors = corsHeaders(env, request, methods);
  return (body: unknown, status = 200, extra: Record<string, string> = {}) =>
    new Response(JSON.stringify(body), { status, headers: { ...cors, 'Content-Type': 'application/json', 'Cache-Control': 'no-store', ...extra } });
}

export const errorBody = (message: string) => ({ error: { message } });
