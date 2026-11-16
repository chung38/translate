// 網站瀏覽人數（functions/api/visit.ts）
import { describe, expect, it } from 'vitest';
import { fakeD1 } from './fake-d1';
import { onRequestGet, onRequestPost, pruneVisitors, readStats, recordVisit } from '../functions/api/visit';

const PHONE = 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) Safari/604.1';
const PC = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/140.0 Safari/537.36';

function post(env: any, ua: string, ip = '1.2.3.4', origin?: string) {
  const headers: Record<string, string> = { 'user-agent': ua, 'cf-connecting-ip': ip };
  if (origin) headers.origin = origin;
  const request = new Request('https://yicipass.com/api/visit', { method: 'POST', headers });
  return (onRequestPost as any)({ request, env, waitUntil() {} }).then((r: Response) => r.json());
}

describe('瀏覽人數', () => {
  it('沒接資料庫時回 enabled: false，網頁就不顯示', async () => {
    expect(await post({}, PHONE)).toEqual({ enabled: false });
    expect(await (onRequestGet as any)({ env: {} }).then((r: Response) => r.json())).toEqual({ enabled: false });
  });

  it('同一人重新整理只算 1 人，但瀏覽次數照算；換人才加人數', async () => {
    const { db } = fakeD1();
    const env = { DB: db };
    await post(env, PHONE);
    await post(env, PHONE);
    const s = await post(env, PC);
    expect(s).toMatchObject({ enabled: true, today: 2, total: 2, viewsToday: 3, viewsTotal: 3 });
    // 同一個瀏覽器、不同網路也算另一人（例如手機從 Wi-Fi 換到 4G）
    expect(await post(env, PHONE, '5.6.7.8')).toMatchObject({ today: 3, viewsToday: 4 });
  });

  it('機器人和其他網站發來的不算', async () => {
    const { db } = fakeD1();
    const env = { DB: db };
    await post(env, 'Googlebot/2.1 (+http://www.google.com/bot.html)');
    await post(env, 'facebookexternalhit/1.1');
    await post(env, '');
    await post(env, PHONE, '1.1.1.1', 'https://evil.example.com');
    expect(await post(env, PHONE, '1.1.1.1', 'https://yicipass.com')).toMatchObject({ today: 1, viewsTotal: 1 });
  });

  it('隔天今日人數歸零，累計繼續加；舊的訪客名單會清掉', async () => {
    const { db, sqlite } = fakeD1();
    await recordVisit(db, '2026-11-15', 'a');
    await recordVisit(db, '2026-11-15', 'b');
    await recordVisit(db, '2026-11-16', 'a');
    expect(await readStats(db, '2026-11-16')).toEqual({ enabled: true, today: 1, total: 3, viewsToday: 1, viewsTotal: 3 });
    expect(await readStats(db, '2026-11-17')).toMatchObject({ today: 0, total: 3 });

    await pruneVisitors(db, Date.parse('2026-11-17T02:00:00Z'));   // 台灣 11/17，保留 11/16 之後
    expect(sqlite.prepare('SELECT day FROM site_visitors').all()).toEqual([{ day: '2026-11-16' }]);
  });
});
