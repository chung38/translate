// 翻譯 API 的帳號額度與每日次數上限（functions/api/translate.ts）
import { afterEach, describe, expect, it, vi } from 'vitest';
import { DatabaseSync } from 'node:sqlite';
import { getAccount, pruneUsage, taipeiDay, takeDailySlot } from '../functions/api/translate';

// 用 Node 內建的 SQLite 模擬 Cloudflare D1（D1 本身就是 SQLite）
function fakeD1() {
  const sqlite = new DatabaseSync(':memory:');
  const db = {
    prepare(sql: string) {
      let values: unknown[] = [];
      const stmt = {
        bind(...v: unknown[]) { values = v; return stmt; },
        async first<T>() { return (sqlite.prepare(sql).get(...(values as any[])) ?? null) as T | null; },
        async run() { return sqlite.prepare(sql).run(...(values as any[])); },
      };
      return stmt;
    },
  };
  return { db, sqlite };
}

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
    const env = { DB: db as any, DAILY_LIMIT: '3' };
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

describe('getAccount', () => {
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

  it('沒設額度時跟網頁一樣當成 2；認得管理員與停用帳號', async () => {
    stub({ fields: { role: { stringValue: 'admin' }, isPendingDeletion: { booleanValue: true } } }, 0);
    expect(await getAccount('acc2', 't', 'p')).toEqual({ isAdmin: true, pendingDeletion: true, quota: 2, used: 0 });
  });

  it('讀不到就回 null（放行），而且 1 分鐘內不重查', async () => {
    stub({}, null, false);
    expect(await getAccount('acc3', 't', 'p')).toBeNull();
    const f = stub({ fields: {} }, 1);
    expect(await getAccount('acc4', 't', 'p')).toEqual({ isAdmin: false, pendingDeletion: false, quota: 2, used: 1 });
    await getAccount('acc4', 't', 'p');
    expect(f).toHaveBeenCalledTimes(2); // 第二次用快取，沒再打 Firestore
  });
});
