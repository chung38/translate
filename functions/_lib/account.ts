import type { Env } from './env';
import { isAdmin, type User } from './auth';
import { getProfile, type Profile } from './profile';
import { createCredit, freeQuota, getCredit, type Credit } from './credits';

export interface Me {
  admin: boolean;
  pendingDeletion: boolean;
  credit: Credit | null;     // null：沒接 D1 或讀取失敗，這時退回舊的檢查方式
  profile: Profile | null;
}

// 讀出這個人的身分、停用狀態與額度。第一次使用時，從 Firestore 的舊額度與翻譯紀錄筆數建立帳戶。
export async function loadMe(env: Env, user: User, idToken: string): Promise<Me> {
  const profile = await getProfile(user.uid, idToken, env.FIREBASE_PROJECT_ID);
  const admin = isAdmin(env, user, profile?.isAdmin ?? false);
  let credit: Credit | null = null;
  if (env.DB) {
    try {
      credit = await getCredit(env.DB, user.uid)
        ?? await createCredit(env.DB, user.uid, user.email ?? null, { quota: profile?.quota ?? freeQuota(env), used: profile?.used ?? 0 });
    } catch (e) {
      console.error('讀取額度失敗：', e);
    }
  }
  return { admin, pendingDeletion: !!profile?.pendingDeletion, credit, profile };
}

export const summary = (me: Me) => ({
  enabled: !!me.credit,
  admin: me.admin,
  quota: me.credit?.quota ?? null,
  used: me.credit?.used ?? null,
  remaining: me.credit ? Math.max(0, me.credit.quota - me.credit.used) : null,
});
