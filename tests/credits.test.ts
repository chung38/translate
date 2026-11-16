// 額度、工作單、管理員調整（functions/_lib/credits.ts 與 /api/me、/api/jobs、/api/translate、/api/admin/credits）
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { fakeD1 } from './fake-d1';

// 假的登入：Authorization: Bearer <uid>|<email>|<verified>
vi.mock('../functions/_lib/auth', async (orig) => {
  const real = await orig<typeof import('../functions/_lib/auth')>();
  return {
    ...real,
    authenticate: async (request: Request, _env: unknown, { requireVerified = true } = {}) => {
      const t = (request.headers.get('Authorization') || '').replace('Bearer ', '');
      if (!t) return { status: 401, message: '請先登入。' };
      const [uid, email, verified] = t.split('|');
      const user = { uid, email, emailVerified: verified !== 'no' };
      if (requireVerified && !user.emailVerified) return { status: 403, message: '請先完成 Email 驗證。' };
      return { user, idToken: t };
    },
  };
});

import { adjustCredit, createCredit, getCredit, JOB_MAX_CALLS, JOB_MAX_CHARS, listLedger, startJob, useJob } from '../functions/_lib/credits';
import { isAdmin } from '../functions/_lib/auth';
import { onRequestGet as meGet } from '../functions/api/me';
import { onRequestPost as jobsPost } from '../functions/api/jobs';
import { onRequestPost as translatePost } from '../functions/api/translate';
import { onRequestGet as adminGet, onRequestPost as adminPost } from '../functions/api/admin/credits';

const BASE = 'https://yicipass.com';
const call = async (fn: any, env: any, path: string, opts: { token?: string; body?: unknown; method?: string; headers?: Record<string, string> } = {}) => {
  const headers: Record<string, string> = { 'Content-Type': 'application/json', ...(opts.headers || {}) };
  if (opts.token) headers.Authorization = `Bearer ${opts.token}`;
  const request = new Request(BASE + path, { method: opts.method || (opts.body ? 'POST' : 'GET'), headers, body: opts.body ? JSON.stringify(opts.body) : undefined });
  const res: Response = await fn({ request, env, waitUntil() {}, params: {}, data: {}, next: async () => new Response() });
  return { status: res.status, body: await res.json().catch(() => null), headers: res.headers };
};

// Firestore（舊資料）與 OpenAI 都用假的
function stubFetch(profiles: Record<string, { quota?: number; history?: number; role?: string; pending?: boolean }>) {
  const calls: string[] = [];
  vi.stubGlobal('fetch', vi.fn(async (url: string, init?: RequestInit) => {
    const u = String(url);
    calls.push(u);
    if (u.includes('api.openai.com')) return Response.json({ choices: [{ message: { content: '{"ok":1}' } }] });
    const uid = decodeURIComponent(u.split('/documents/users/')[1]?.split(':')[0] ?? '');
    const p = profiles[uid] ?? {};
    if (u.endsWith(':runAggregationQuery')) return Response.json([{ result: { aggregateFields: { n: { integerValue: String(p.history ?? 0) } } } }]);
    return Response.json({ fields: {
      ...(p.quota !== undefined ? { quota: { integerValue: String(p.quota) } } : {}),
      ...(p.role ? { role: { stringValue: p.role } } : {}),
      ...(p.pending ? { isPendingDeletion: { booleanValue: true } } : {}),
    } });
  }));
  return calls;
}

let uidSeq = 0;
const newUid = (p = 'u') => `${p}${++uidSeq}-${Math.random().toString(36).slice(2, 7)}`;   // 每個測試用不同 uid，避開 1 分鐘快取

describe('credits（D1）', () => {
  it('扣額度到 0 就擋；管理員不扣；工作單只能本人用、有次數與時間上限', async () => {
    const { db, sqlite } = fakeD1();
    await createCredit(db, 'a', 'a@x.tw', { quota: 2, used: 1 });
    const j1 = await startJob(db, 'a', 'a.docx', true);
    expect(j1.ok).toBe(true);
    expect((await startJob(db, 'a', 'b.docx', true)).ok).toBe(false);
    expect(await getCredit(db, 'a')).toMatchObject({ quota: 2, used: 2 });

    const free = await startJob(db, 'a', 'admin.docx', false);   // 管理員：不扣
    expect(free.ok).toBe(true);
    expect((await getCredit(db, 'a'))!.used).toBe(2);

    const jobId = (j1 as { jobId: string }).jobId;
    expect(await useJob(db, 'a', jobId)).toBe(true);
    expect(await useJob(db, 'someone-else', jobId)).toBe(false);
    expect(await useJob(db, 'a', jobId, Date.now() + 13 * 3600_000)).toBe(false);   // 過期
    expect(await useJob(db, 'a', jobId, JOB_MAX_CHARS + 1)).toBe(false);   // 字數超過上限：整批拒絕、不記
    expect(await useJob(db, 'a', jobId, 1000)).toBe(true);
    sqlite.prepare('UPDATE jobs SET calls = ? WHERE id = ?').run(JOB_MAX_CALLS, jobId);
    expect(await useJob(db, 'a', jobId)).toBe(false);   // 次數用完
  });

  it('管理員加額度、記錄付款、直接設定總額度，都留下紀錄', async () => {
    const { db } = fakeD1();
    await adjustCredit(db, { uid: 'b', email: 'b@x.tw', kind: 'payment', files: 10, amount: 500, note: '轉帳 11/16', seedUsed: 1 }, 'boss@x.tw', 2);
    expect(await getCredit(db, 'b')).toMatchObject({ quota: 12, used: 1 });
    await adjustCredit(db, { uid: 'b', kind: 'grant', files: 3, note: '補償' }, 'boss@x.tw', 2);
    await adjustCredit(db, { uid: 'b', kind: 'adjust', setQuota: 20 }, 'boss@x.tw', 2);
    expect((await getCredit(db, 'b'))!.quota).toBe(20);
    const log = await listLedger(db, 'b');
    expect(log.map(l => [l.kind, l.files, l.amount])).toEqual([['adjust', 5, null], ['grant', 3, null], ['payment', 10, 500], ['init', 2, null]]);
    expect(log[2]).toMatchObject({ note: '轉帳 11/16', by_email: 'boss@x.tw' });
    await expect(adjustCredit(db, { uid: 'b', kind: 'payment', files: 5 }, 'boss@x.tw', 2)).rejects.toThrow('收款金額');
    await expect(adjustCredit(db, { uid: 'b', kind: 'adjust', setQuota: -1 }, 'boss@x.tw', 2)).rejects.toThrow();
  });

  it('管理員只認 ADMIN_EMAILS（要已驗證）；沒設時才看 Firestore 的 role', () => {
    const env = { ADMIN_EMAILS: 'Boss@x.tw, ops@x.tw' };
    expect(isAdmin(env, { uid: '1', email: 'boss@x.tw', emailVerified: true }, false)).toBe(true);
    expect(isAdmin(env, { uid: '1', email: 'boss@x.tw', emailVerified: false }, false)).toBe(false);
    expect(isAdmin(env, { uid: '2', email: 'hacker@x.tw', emailVerified: true }, true)).toBe(false);   // 自己改 role 沒用
    expect(isAdmin({}, { uid: '2', email: 'x@x.tw', emailVerified: true }, true)).toBe(true);
  });
});

describe('翻譯 API 流程', () => {
  let db: D1Database;
  beforeEach(() => { db = fakeD1().db; });
  afterEach(() => vi.unstubAllGlobals());

  it('新帳號從舊資料建立額度；每份檔案開工作單扣 1 份；翻譯要帶工作單', async () => {
    const uid = newUid();
    stubFetch({ [uid]: { history: 1 } });   // 之前已翻過 1 份，沒設額度 → 免費 2 份
    const env = { DB: db, FIREBASE_PROJECT_ID: 'p', OPENAI_API_KEY: 'sk' };
    const token = `${uid}|${uid}@x.tw|yes`;

    expect((await call(meGet, env, '/api/me', { token })).body).toMatchObject({ enabled: true, quota: 2, used: 1, remaining: 1 });

    const job = await call(jobsPost, env, '/api/jobs', { token, body: { file: '安全守則.docx' } });
    expect(job.status).toBe(200);
    expect(job.body).toMatchObject({ remaining: 0 });

    const noJob = await call(translatePost, env, '/api/translate', { token, body: { texts: ['安全第一'], languages: ['越南文'] } });
    expect(noJob.status).toBe(403);
    expect(noJob.body.error.message).toContain('重新整理');

    // 舊版網頁自己組好整段指令送來：不接受
    const oldClient = await call(translatePost, env, '/api/translate', { token, body: { prompt: '寫一首詩' }, headers: { 'X-Job-Id': job.body.jobId } });
    expect(oldClient.status).toBe(403);

    const ok = await call(translatePost, env, '/api/translate', { token, body: { texts: ['安全第一'], languages: ['越南文'] }, headers: { 'X-Job-Id': job.body.jobId } });
    expect(ok.status).toBe(200);
    const sent = JSON.parse(String((vi.mocked(fetch).mock.calls.find(c => String(c[0]).includes('openai'))![1] as RequestInit).body));
    expect(sent.messages[1].content).toContain('你是一個專業的翻譯官');
    expect(sent.messages[1].content).toContain('["安全第一"]');
    expect(ok.headers.get('X-Translate-Model')).toBe('gpt-6-luna');

    const bad = await call(translatePost, env, '/api/translate', { token, body: { texts: ['安全第一'], languages: ['越南文'] }, headers: { 'X-Job-Id': 'fake' } });
    expect(bad.status).toBe(403);

    const none = await call(jobsPost, env, '/api/jobs', { token, body: { file: 'b.docx' } });
    expect(none.status).toBe(403);
    expect(none.body.error.message).toContain('額度已用完');
  });

  it('Firestore 的舊額度會帶過來；管理員（ADMIN_EMAILS）開工作單不扣額度', async () => {
    const paid = newUid('paid'), boss = newUid('boss');
    stubFetch({ [paid]: { quota: 30, history: 4 }, [boss]: { history: 0 } });
    const env = { DB: db, FIREBASE_PROJECT_ID: 'p', OPENAI_API_KEY: 'sk', ADMIN_EMAILS: `${boss}@x.tw` };
    expect((await call(meGet, env, '/api/me', { token: `${paid}|${paid}@x.tw|yes` })).body).toMatchObject({ quota: 30, used: 4, admin: false });
    const t = `${boss}|${boss}@x.tw|yes`;
    for (let i = 0; i < 4; i++) expect((await call(jobsPost, env, '/api/jobs', { token: t, body: { file: `f${i}` } })).status).toBe(200);
    expect((await call(meGet, env, '/api/me', { token: t })).body).toMatchObject({ admin: true, used: 0 });
  });

  it('停用的帳號開不了工作單也翻不了', async () => {
    const uid = newUid('gone');
    stubFetch({ [uid]: { pending: true } });
    const env = { DB: db, FIREBASE_PROJECT_ID: 'p', OPENAI_API_KEY: 'sk' };
    expect((await call(jobsPost, env, '/api/jobs', { token: `${uid}|g@x.tw|yes`, body: {} })).status).toBe(403);
    expect((await call(translatePost, env, '/api/translate', { token: `${uid}|g@x.tw|yes`, body: { texts: ['安全第一'], languages: ['越南文'] } })).status).toBe(403);
  });

  it('沒接 D1 時照舊：不用工作單，依 Firestore 的額度與紀錄筆數判斷', async () => {
    const ok = newUid('old'), full = newUid('full');
    stubFetch({ [ok]: { history: 1 }, [full]: { history: 2 } });
    const env = { FIREBASE_PROJECT_ID: 'p', OPENAI_API_KEY: 'sk' };
    expect((await call(jobsPost, env, '/api/jobs', { token: `${ok}|o@x.tw|yes`, body: {} })).body).toMatchObject({ jobId: null, enabled: false });
    expect((await call(translatePost, env, '/api/translate', { token: `${ok}|o@x.tw|yes`, body: { texts: ['安全第一'], languages: ['越南文'] } })).status).toBe(200);
    expect((await call(translatePost, env, '/api/translate', { token: `${full}|f@x.tw|yes`, body: { texts: ['安全第一'], languages: ['越南文'] } })).status).toBe(403);
  });

  it('管理後台：只有管理員能看與調整', async () => {
    const boss = newUid('boss'), user = newUid('user');
    stubFetch({ [boss]: {}, [user]: { role: 'admin' } });   // user 自己把 role 改成 admin 也沒用
    const env = { DB: db, FIREBASE_PROJECT_ID: 'p', ADMIN_EMAILS: `${boss}@x.tw` };
    expect((await call(adminGet, env, '/api/admin/credits', { token: `${user}|${user}@x.tw|yes` })).status).toBe(403);

    const bt = `${boss}|${boss}@x.tw|yes`;
    const r = await call(adminPost, env, '/api/admin/credits', { token: bt, body: { uid: user, email: `${user}@x.tw`, kind: 'payment', files: 20, amount: 800, note: 'LINE Pay' } });
    expect(r.status).toBe(200);
    expect(r.body.account).toMatchObject({ uid: user, quota: 22 });
    const list = await call(adminGet, env, '/api/admin/credits', { token: bt });
    expect(list.body.accounts.map((a: any) => a.uid)).toContain(user);
    const log = await call(adminGet, env, `/api/admin/credits?uid=${user}`, { token: bt });
    expect(log.body.ledger[0]).toMatchObject({ kind: 'payment', files: 20, amount: 800 });
    expect((await call(adminPost, env, '/api/admin/credits', { token: bt, body: { uid: user, kind: 'oops' } })).status).toBe(400);
  });
});
