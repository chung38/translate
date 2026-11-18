// 驗證信相關的提示文字：註冊後、重寄後都用同一套說法

/** 驗證信寄出後給使用者看的提示 */
export function verifySentMessage(email: string | null | undefined): string {
  const to = email ? `驗證信已寄到 ${email}。` : '驗證信已寄出。';
  return `${to}\n通常 1～2 分鐘內會到；沒看到的話，請找「垃圾郵件」或「促銷內容」資料夾（寄件人是 noreply 開頭的地址）。\n點信裡的連結後回到這一頁，系統會自動解鎖翻譯功能。`;
}

/** 寄驗證信失敗時的說明 */
export function verifySendError(err: unknown): string {
  const code = (err as { code?: string } | null)?.code;
  if (code === 'auth/too-many-requests') {
    return '驗證信寄太多次，暫時被擋下。請等 10～30 分鐘再按「重發驗證信」，這段時間可以先到信箱的垃圾郵件找找看。';
  }
  if (code === 'auth/network-request-failed') {
    return '網路連線不穩，驗證信沒有寄出。請確認網路後再按「重發驗證信」。';
  }
  return '驗證信寄送失敗，請稍後再按「重發驗證信」。';
}

/** 還沒驗證就想翻譯時的提示 */
export const VERIFY_FIRST_MESSAGE =
  '還沒完成 Email 驗證。請到信箱點驗證連結（沒看到請找垃圾郵件），完成後這一頁會自動解鎖；沒收到可以按上方的「重發驗證信」。';
