// ─── 翻譯指令（提示詞）──────────────────────────────────────────────────
// 由後端組，瀏覽器只送「要翻的句子、目標語言、產業別」。
// 這樣別人就算繞過網頁，也只能拿這個 API 來翻譯，不能拿去問其他問題。
// 這段文字從原本的前端程式原封不動搬過來（含縮排），改動會影響翻譯品質，請小心。

export const TARGET_LANGUAGES = ['繁體中文', '泰文', '印尼文', '越南文', '英文'] as const;
export const MAX_TEXTS_PER_CALL = 50;
export const MAX_CHARS_PER_TEXT = 20_000;
export const MAX_CHARS_PER_CALL = 40_000;
export const MAX_INDUSTRY_CHARS = 30;

export interface TranslateRequest { texts: string[]; languages: string[]; industry: string }

// 檢查瀏覽器送來的內容；不合格回傳錯誤訊息
export function parseTranslateRequest(body: any): { ok: true; req: TranslateRequest; chars: number } | { ok: false; message: string } {
  if (body && typeof body.prompt === 'string') return { ok: false, message: '網站已更新，請重新整理頁面後再翻譯。' };
  const texts = body?.texts, languages = body?.languages;
  if (!Array.isArray(texts) || !texts.length || texts.length > MAX_TEXTS_PER_CALL || !texts.every((t: unknown) => typeof t === 'string')) {
    return { ok: false, message: '無效的請求內容（texts）' };
  }
  if (texts.some((t: string) => t.length > MAX_CHARS_PER_TEXT)) return { ok: false, message: '有段落太長，超過單段處理上限。' };
  const chars = texts.reduce((n: number, t: string) => n + t.length, 0);
  if (chars > MAX_CHARS_PER_CALL) return { ok: false, message: '文本內容過長，超過系統單次處理限制。' };
  if (!Array.isArray(languages) || !languages.length || languages.length > TARGET_LANGUAGES.length
      || new Set(languages).size !== languages.length
      || !languages.every((l: unknown) => (TARGET_LANGUAGES as readonly unknown[]).includes(l))) {
    return { ok: false, message: '無效的請求內容（languages）' };
  }
  // 產業別會放進指令裡：限制長度、去掉換行與引號類符號，避免被拿來夾帶其他指令
  const industry = typeof body.industry === 'string'
    ? body.industry.replace(/[\r\n\t`"'「」{}\[\]<>]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, MAX_INDUSTRY_CHARS)
    : '';
  return { ok: true, req: { texts, languages, industry }, chars };
}

export function buildTranslatePrompt({ texts, languages: targetLangs, industry }: TranslateRequest): string {
  const industryContext = industry ? `。這是一個關於「${industry}」行業的文件，請使用該行業的專業術語進行翻譯` : '';
  return `你是一個專業的翻譯官${industryContext}。
      請將以下文字陣列中的每一項同時翻譯成以下語言：${targetLangs.join('、')}。

      請嚴格以 JSON 格式回傳結果，格式如下：
      {
        "translations": [
          { ${targetLangs.map(l => `"${l}": "翻譯內容"`).join(', ')} },
          ...
        ]
      }

      要求：
      1. 確保回傳的 "translations" 陣列長度與輸入的文字陣列長度完全一致 (${texts.length})。
      2. 每個物件的鍵 (Key) 必須完全對應目標語言名稱：${targetLangs.map(l => `"${l}"`).join(', ')}。
      3. **極度重要：翻譯後的文字內容中，絕對不要包含任何語言名稱的標籤或前綴（例如絕對不要出現 [英文]、[泰文]、English: 等字樣），只能有翻譯後的純文字。**
      4. 確保翻譯內容完全使用目標語言，不要夾雜原始語言。
      5. **標籤保留與對應：** 譯文必須原封不動地保留原文中的排版標籤 (如 [f0], [f1])。
      6. **每一個原文中出現的 tag 都必須在譯文中出現一次且只出現一次，不能遺失、不能合併、不能改變編號。**
      7. **若原文有 [f0]...[/f0][f1]...[/f1][f2]...[/f2]，譯文也必須保留完全相同的 [f0]、[f1]、[f2] 結構。**
      8. **禁止把兩個 tag 的內容合併成一個 tag，也禁止省略任何短詞內容；像「現場」、「日期」、「地址」這類短詞也必須翻譯。**
      9. **防斷字原則：** 當原文的標籤切斷了同一個詞 (如 \`[f1]設[/f1][f2]備[/f2]\`)，翻譯後的目標語言「同一個音節/單字」內部**絕對不能有空格** (例如 \`[f1]thi[/f1][f2]ết bị[/f2]\`)。
      10. **不同詞彙之間的必定空格 (Critical Spacing Rule)：** 當原文在沒有空格的狀態下切換標籤 (例如 \`[f0]打開[/f0][f1]總電源[/f1]\`)，當翻譯為越南文或英文時，**如果這兩個標籤代表的是兩個不同的單字，你必須在標籤內側或外側加上空格！**
         **[極度重要範例]：**
         ❌ 錯誤：\`[f0]Mở[/f0][f1]tủ điện số PB1[/f1]\` (單字會黏在一起變成 Mởtủ)
         ✅ 正確：\`[f0]Mở [/f0][f1]tủ điện số PB1[/f1]\` (加上了空格，單字分離)
         ❌ 錯誤：\`[f2], sau đó[/f2][f3]bật công tắc[/f3]\` (變成 đóbật)
         ✅ 正確：\`[f2], sau đó [/f2][f3]bật công tắc[/f3]\`
      11. **表單填寫空格保留與字母防散（極度重要）：**
         - 若原文中存在連續多個空白字元（例如做為排版或手寫填寫空間的 \`年   月   日\` 或 \`Name:      \`），你**必須**在翻譯結果中原封不動保留對應的大片空白（例如 \`Năm   Tháng   Ngày\`）。
         - **但是！** 若原文是因為排版對齊，而在同一個詞彙的中文字之間插入了空格（例如 \`申  請  人:\` 或 \`工  作  地:\`），當翻譯成拼音語言（如英文、印文、越文）時，**絕對不准**將空格照抄分布到字母之間！請直接輸出正常拼寫的單字（例如輸出 \`Pemohon:\`，**嚴禁**輸出 \`P e m o h o n:\`；輸出 \`Người yêu cầu:\`，**嚴禁**輸出 \`N g ư ờ i  y ê u  c ầ u:\`）。
      12. 不要包含任何 Markdown 標籤（如 \`\`\`json）或額外文字，只回傳純 JSON 字串。
      13. **完整翻譯檢查：**
         - 原文中的中文正文必須完整翻譯，不得因為句子內含英文專有名詞、設備名稱、縮寫或型號而略過其餘中文。
         - Interlock、LOTO、HMI、AUTO、MANUAL、VETO、Human Limits 等術語可依語意保留，但周圍所有中文必須完整翻譯。
         - 若目標語言為印尼文，譯文不得殘留中文字，也禁止出現「jika 不幸 menimbulkan」這類中印混雜內容。

      待翻譯內容陣列：
      ${JSON.stringify(texts)}`;
}
