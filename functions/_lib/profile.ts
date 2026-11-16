import firebaseConfig from '../../firebase-applet-config.json';

// 前端用的同一份 Firebase 設定；這裡只取 Firestore 資料庫名稱（AI Studio 建的不是預設資料庫）
const FIRESTORE_DB = (firebaseConfig as { firestoreDatabaseId?: string }).firestoreDatabaseId || '(default)';

// 讀使用者自己在 Firestore 的資料（角色、停用標記、舊額度、翻譯紀錄筆數）。
// 用他自己的登入 token 讀，所以不需要額外金鑰。讀不到（網路問題、資料還沒建立）回 null。
export interface Profile {
  isAdmin: boolean;
  pendingDeletion: boolean;
  quota: number | null;   // 管理員在 Firestore 設的額度（沒設是 null）
  used: number;    // 已翻幾份（翻譯紀錄筆數）
}

const TTL = 60_000;   // 同一個人 1 分鐘內只查一次，一份文件會連續呼叫很多次
const cache = new Map<string, { at: number; value: Profile }>();

export async function getProfile(uid: string, idToken: string, projectId: string): Promise<Profile | null> {
  const hit = cache.get(uid);
  if (hit && Date.now() - hit.at < TTL) return hit.value;

  const docUrl = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/${FIRESTORE_DB}/documents/users/${encodeURIComponent(uid)}`;
  const headers = { Authorization: `Bearer ${idToken}`, 'Content-Type': 'application/json' };
  try {
    const [docRes, countRes] = await Promise.all([
      fetch(docUrl, { headers, signal: AbortSignal.timeout(5000) }),
      fetch(`${docUrl}:runAggregationQuery`, {
        method: 'POST',
        headers,
        signal: AbortSignal.timeout(5000),
        body: JSON.stringify({
          structuredAggregationQuery: {
            structuredQuery: { from: [{ collectionId: 'history' }] },
            aggregations: [{ alias: 'n', count: {} }],
          },
        }),
      }),
    ]);
    if (!docRes.ok || !countRes.ok) return null;

    const f = ((await docRes.json()) as any)?.fields ?? {};
    const agg = (await countRes.json()) as any[];
    const used = Number(agg?.[0]?.result?.aggregateFields?.n?.integerValue);
    if (!Number.isFinite(used)) return null;

    const q = Number(f.quota?.integerValue ?? f.quota?.doubleValue);
    const value: Profile = {
      isAdmin: f.role?.stringValue === 'admin',
      pendingDeletion: f.isPendingDeletion?.booleanValue === true,
      quota: Number.isFinite(q) ? q : null,
      used,
    };
    cache.set(uid, { at: Date.now(), value });
    return value;
  } catch {
    return null;
  }
}

