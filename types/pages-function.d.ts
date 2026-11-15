// Cloudflare Pages Functions 的型別（只給 npm run lint 用，不會被打包）。
// 放在 functions/ 外面，免得被 Cloudflare 當成一條路由。
type PagesFunction<Env = unknown> = (context: {
  request: Request;
  env: Env;
  params: Record<string, string | string[]>;
  data: Record<string, unknown>;
  waitUntil(promise: Promise<unknown>): void;
  next(input?: Request | string, init?: RequestInit): Promise<Response>;
}) => Response | Promise<Response>;
