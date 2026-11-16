// 翻譯 API 的共用程式：Firestore 資料、每日次數上限、翻譯模型（functions/_lib/）
import { afterEach, describe, expect, it, vi } from 'vitest';
import { fakeD1 } from './fake-d1';
import { getProfile as getAccount } from '../functions/_lib/profile';
import { pruneUsage, taipeiDay, takeDailySlot } from '../functions/_lib/limits';
import { adjustForError, callModel, pickProvider } from '../functions/_lib/model';

describe('taipeiDay', () => {
  it('用台灣時間換日', () => {
    expect(taipeiDay(Date.parse('2026-11-15T15:59:00Z'))).toBe('2026-11-15'); // 台灣 23:59
    expect(taipeiDay(Date.parse('2026-11-15T16:00:00Z'))).toBe('2026-11-16'); // 台灣 00:00
  });
});

describe('takeDailySlot', () => {
  it('沒有接資料庫就不限制', async () => {
    expect(await takeDailySlot({}, 'u')).toBeNull();
  });

  it('到上限就擋，被擋的那次不加次數', async () => {
    const { db, sqlite } = fakeD1();
    const env = { DB: db, DAILY_LIMIT: '3' };
    const now = Date.parse('2026-11-15T02:00:00Z');
    const got = [];
    for (let i = 0; i < 5; i++) got.push(await takeDailySlot(env, 'u1', now));
    expect(got.map(r => r?.ok)).toEqual([true, true, true, false, false]);
    expect(got[2]).toEqual({ ok: true, used: 3, limit: 3 });
    expect(sqlite.prepare('SELECT n FROM api_usage WHERE uid = ?').get('u1')).toEqual({ n: 3 });

    // 別人不受影響；隔天重新計算
    expect((await takeDailySlot(env, 'u2', now))?.ok).toBe(true);
    expect((await takeDailySlot(env, 'u1', now + 86400_000))?.used).toBe(1);
  });

  it('預設每天 2000 次', async () => {
    const { db } = fakeD1();
    expect((await takeDailySlot({ DB: db as any }, 'u'))?.limit).toBe(2000);
  });

  it('資料庫出錯時放行', async () => {
    const broken = { prepare() { throw new Error('D1 掛了'); } };
    vi.spyOn(console, 'error').mockImplementation(() => {});
    expect(await takeDailySlot({ DB: broken as any }, 'u')).toBeNull();
  });

  it('清掉 90 天前的紀錄', async () => {
    const { db, sqlite } = fakeD1();
    const now = Date.parse('2026-11-15T02:00:00Z');
    await takeDailySlot({ DB: db as any }, 'old', now - 100 * 86400_000);
    await takeDailySlot({ DB: db as any }, 'new', now);
    await pruneUsage(db as any, now);
    expect(sqlite.prepare('SELECT uid FROM api_usage').all()).toEqual([{ uid: 'new' }]);
  });
});

describe('getProfile（Firestore 舊資料）', () => {
  afterEach(() => vi.unstubAllGlobals());

  const stub = (doc: unknown, count: number | null, ok = true) => {
    const fetchMock = vi.fn(async (url: string, _init?: RequestInit) => {
      if (!ok) return new Response('{}', { status: 403 });
      if (String(url).endsWith(':runAggregationQuery')) {
        return Response.json(count === null ? [] : [{ result: { aggregateFields: { n: { integerValue: String(count) } } } }]);
      }
      return Response.json(doc);
    });
    vi.stubGlobal('fetch', fetchMock);
    return fetchMock;
  };

  it('讀出角色、額度、已用份數，並查對的資料庫', async () => {
    const f = stub({ fields: { role: { stringValue: 'user' }, quota: { integerValue: '5' } } }, 3);
    expect(await getAccount('acc1', 'tok', 'proj')).toEqual({ isAdmin: false, pendingDeletion: false, quota: 5, used: 3 });
    const [docUrl, init] = f.mock.calls[0] as [string, RequestInit];
    expect(docUrl).toContain('/projects/proj/databases/ai-studio-');
    expect(docUrl).toContain('/documents/users/acc1');
    expect((init.headers as Record<string, string>).Authorization).toBe('Bearer tok');
  });

  it('沒設額度時回 null（由後端決定免費份數）；認得管理員與停用帳號', async () => {
    stub({ fields: { role: { stringValue: 'admin' }, isPendingDeletion: { booleanValue: true } } }, 0);
    expect(await getAccount('acc2', 't', 'p')).toEqual({ isAdmin: true, pendingDeletion: true, quota: null, used: 0 });
  });

  it('讀不到就回 null（放行），而且 1 分鐘內不重查', async () => {
    stub({}, null, false);
    expect(await getAccount('acc3', 't', 'p')).toBeNull();
    const f = stub({ fields: {} }, 1);
    expect(await getAccount('acc4', 't', 'p')).toEqual({ isAdmin: false, pendingDeletion: false, quota: null, used: 1 });
    await getAccount('acc4', 't', 'p');
    expect(f).toHaveBeenCalledTimes(2); // 第二次用快取，沒再打 Firestore
  });
});

describe('翻譯模型', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('有 OPENAI_API_KEY 就用 gpt-6-luna、關掉推理；沒有才用 DeepSeek', () => {
    const o = pickProvider({ OPENAI_API_KEY: 'sk-o', DEEPSEEK_API_KEY: 'sk-d' })!;
    expect(o.name).toBe('openai');
    expect(o.model).toBe('gpt-6-luna');
    expect(o.options).toMatchObject({ reasoning_effort: 'none', max_completion_tokens: 8192 });
    expect(o.options).not.toHaveProperty('max_tokens');
    expect(pickProvider({ OPENAI_API_KEY: 'k', OPENAI_MODEL: 'gpt-6-sol' })!.model).toBe('gpt-6-sol');
    expect(pickProvider({ DEEPSEEK_API_KEY: 'sk-d' })!.name).toBe('deepseek');
    expect(pickProvider({})).toBeNull();
  });

  it('模型不接受某個參數時自動調整', () => {
    const base = { reasoning_effort: 'none', temperature: 0.3, response_format: { type: 'json_object' } };
    expect(adjustForError(base, 400, { error: { param: 'temperature', message: "Unsupported parameter: 'temperature'" } }))
      .toEqual({ reasoning_effort: 'none', response_format: { type: 'json_object' } });
    expect(adjustForError(base, 400, { error: { message: "Unsupported value: 'reasoning_effort' does not support 'none'" } }))
      .toMatchObject({ reasoning_effort: 'low' });
    expect(adjustForError(base, 400, { error: { message: 'something else' } })).toBeNull();
    expect(adjustForError(base, 429, { error: { param: 'temperature' } })).toBeNull();
  });

  it('送出正確的請求；被拒的參數拿掉重送後，之後都記得', async () => {
    const bodies: any[] = [];
    vi.stubGlobal('fetch', vi.fn(async (_url: string, init?: RequestInit) => {
      const b = JSON.parse(String(init!.body));
      bodies.push(b);
      if ('temperature' in b) return Response.json({ error: { param: 'temperature', message: "Unsupported parameter: 'temperature'" } }, { status: 400 });
      return Response.json({ choices: [{ message: { content: '{"ok":true}' } }] });
    }));
    const p = pickProvider({ OPENAI_API_KEY: 'sk-test', OPENAI_MODEL: 'luna-test' })!;
    const res = await callModel(p, '翻譯這句');
    expect(res.ok).toBe(true);
    expect(bodies).toHaveLength(2);
    expect(bodies[0]).toMatchObject({ model: 'luna-test', reasoning_effort: 'none', temperature: 0.3 });
    expect(bodies[0].messages[1]).toEqual({ role: 'user', content: '翻譯這句' });
    expect(bodies[1]).not.toHaveProperty('temperature');

    await callModel(p, '第二句');
    expect(bodies).toHaveLength(3);   // 直接用調整後的設定，不會再被拒一次
    expect(bodies[2]).not.toHaveProperty('temperature');
  });

  it('金鑰錯誤直接回傳，不重試', async () => {
    const f = vi.fn(async () => Response.json({ error: { message: 'Incorrect API key provided' } }, { status: 401 }));
    vi.stubGlobal('fetch', f);
    const res = await callModel(pickProvider({ OPENAI_API_KEY: 'bad' })!, 'x');
    expect(res.status).toBe(401);
    expect(f).toHaveBeenCalledTimes(1);
  });
});
