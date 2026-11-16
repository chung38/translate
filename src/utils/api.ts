// 呼叫自己的後端（functions/api/*），自動帶上登入 token
import { auth } from '../firebase';

export interface Credit {
  enabled: boolean;          // 後端有接 D1、額度由伺服器管理
  admin: boolean;
  quota: number | null;      // 總共可以翻幾份
  used: number | null;       // 已經用掉幾份
  remaining: number | null;
}

export async function api<T = any>(path: string, init: RequestInit = {}): Promise<{ ok: boolean; status: number; data: T }> {
  const token = await auth.currentUser?.getIdToken();
  const res = await fetch(path, {
    ...init,
    headers: {
      'Content-Type': 'application/json',
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...(init.headers || {}),
    },
  });
  const data = await res.json().catch(() => ({})) as T;
  return { ok: res.ok, status: res.status, data };
}

export const errorMessage = (data: any, fallback: string) => data?.error?.message || fallback;
