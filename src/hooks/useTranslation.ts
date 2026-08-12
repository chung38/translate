import React, { useState, useRef } from 'react';
import { User } from 'firebase/auth';
import { doc, setDoc, collection, Timestamp } from 'firebase/firestore';
import { db } from '../firebase';
import { jsonrepair } from 'jsonrepair';

type TranslationStatus = 'idle' | 'processing' | 'translating' | 'generating' | 'completed' | 'error';

const DEEPSEEK_PROXY_URL = '/api/translate';

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

const extractTagIds = (text: string): number[] => {
  const ids = new Set<number>();
  const regex = /\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g;
  let match;
  while ((match = regex.exec(text)) !== null) {
    ids.add(parseInt(match[1], 10));
  }
  return [...ids].sort((a, b) => a - b);
};

const hasTagMarkup = (text: string) => /\[f\d+\][\s\S]*?\[\/f\d+\]/.test(text);

const hasExactSameTags = (source: string, target: string) => {
  const sourceIds = extractTagIds(source);
  const targetIds = extractTagIds(target);
  return sourceIds.length === targetIds.length && sourceIds.every((id, index) => id === targetIds[index]);
};

const stripAllTags = (text: string) => text.replace(/\[f\d+\]/g, '').replace(/\[\/f\d+\]/g, '');

const rebuildWithSourceTags = (source: string, translated: string) => {
  const srcMatches = [...source.matchAll(/\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g)];
  if (srcMatches.length === 0) return translated;

  const cleanTranslated = stripAllTags(translated).trim();
  if (!cleanTranslated) return source;

  const parts = cleanTranslated.split(/\s+/).filter(Boolean);
  if (parts.length === 0) return source;

  if (srcMatches.length === 1) {
    const id = srcMatches[0][1];
    return `[f${id}]${cleanTranslated}[/f${id}]`;
  }

  const totalSourceChars = srcMatches.reduce((sum, match) => sum + (match[2]?.length || 0), 0) || srcMatches.length;
  let cursor = 0;
  const assigned: string[] = [];

  for (let i = 0; i < srcMatches.length; i++) {
    const srcLen = srcMatches[i][2]?.length || 1;
    const ratio = srcLen / totalSourceChars;
    let takeCount = Math.round(parts.length * ratio);

    if (i === srcMatches.length - 1) {
      takeCount = parts.length - cursor;
    } else {
      const remainingSlots = srcMatches.length - i - 1;
      const remainingWords = parts.length - cursor;
      takeCount = Math.max(1, Math.min(takeCount, remainingWords - remainingSlots));
    }

    assigned.push(parts.slice(cursor, cursor + takeCount).join(' '));
    cursor += takeCount;
  }

  return srcMatches
    .map((match, index) => `[f${match[1]}]${assigned[index] || match[2] || ''}[/f${match[1]}]`)
    .join('');
};

const repairTranslationTags = (sourceText: string, translatedText: string) => {
  if (!hasTagMarkup(sourceText)) return translatedText || '';
  if (hasExactSameTags(sourceText, translatedText || '')) return translatedText || '';
  return rebuildWithSourceTags(sourceText, translatedText || '');
};

export const useTranslation = (
  user: User | null
) => {
  const [status, setStatus] = useState<TranslationStatus>('idle');
  const [statusMessage, setStatusMessage] = useState('');
  const [progress, setProgress] = useState(0);
  const [fileProgress, setFileProgress] = useState<Record<string, number>>({});
  const [error, setError] = useState<string | null>(null);
  const isCancelledRef = useRef(false);

  const saveToFirestore = async (fileName: string, translatedName: string, type: string, targetLanguages: string[], industry: string) => {
    if (!user) return;
    try {
      const historyRef = doc(collection(db, 'users', user.uid, 'history'));
      const newRecord = {
        userId: user.uid,
        fileName,
        translatedFileName: translatedName,
        fileType: type,
        targetLanguages,
        industry,
        timestamp: Timestamp.now(),
        status: 'completed'
      };
      await setDoc(historyRef, newRecord);
    } catch (err) {
      console.error('Error saving to Firestore:', err);
    }
  };

  const translateBatch = async (texts: string[], targetLangs: string[], industry: string, retryCount = 0): Promise<Record<string, string>[]> => {
    if (texts.length === 0 || targetLangs.length === 0) return texts.map(() => ({}));

    try {
      const industryContext = industry ? `。這是一個關於「${industry}」行業的文件，請使用該行業的專業術語進行翻譯` : '';
      const prompt = `你是一個專業的翻譯官${industryContext}。
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

      待翻譯內容陣列：
      ${JSON.stringify(texts)}`;

      const controller = new AbortController();
      // Increase timeout to 5 minutes (300000ms) to handle large batches and server-side retries
      const timeoutId = setTimeout(() => controller.abort(), 300000);

      const response = await fetch(DEEPSEEK_PROXY_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ prompt }),
        signal: controller.signal
      });

      clearTimeout(timeoutId);

      if (!response.ok) {
        let errorMessage = `HTTP error! status: ${response.status}`;
        const errorText = await response.text();
        try {
          const errorData = JSON.parse(errorText);
          errorMessage = errorData.error?.message || errorMessage;
        } catch (e) {
          errorMessage = errorText || errorMessage;
        }
        throw new Error(errorMessage);
      }

      const responseText = await response.text();
      let data;
      try {
        data = JSON.parse(responseText);
      } catch (e) {
        throw new Error(`Invalid JSON from server: ${responseText.substring(0, 200)}`);
      }
      const resultText = data.choices[0].message.content.trim();

      // Extremely robust JSON extraction: find the first { and the last }
      // This bypasses any markdown blocks (```json) or conversational chatter entirely.
      let cleanJson = resultText;
      const firstBrace = cleanJson.indexOf('{');
      const lastBrace = cleanJson.lastIndexOf('}');
      if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
        cleanJson = cleanJson.substring(firstBrace, lastBrace + 1);
      }

      let parsed;
      try {
        parsed = JSON.parse(cleanJson);
      } catch (e) {
        console.error('JSON Parse Error. Raw text:', resultText);
        try {
          const repairedJson = jsonrepair(cleanJson);
          parsed = JSON.parse(repairedJson);
        } catch (repairError) {
          console.error('JSON Repair Error:', repairError);
          const jsonMatch = resultText.match(/\{[\s\S]*\}/);
          if (jsonMatch) {
            try {
              const repairedMatch = jsonrepair(jsonMatch);
              parsed = JSON.parse(repairedMatch);
            } catch (matchRepairError) {
              throw new Error('無法解析 API 回傳的 JSON 格式');
            }
          } else {
            throw new Error('無法解析 API 回傳的 JSON 格式');
          }
        }
      }

      if (!parsed.translations || !Array.isArray(parsed.translations)) {
        throw new Error('API 回傳格式不正確 (缺少 translations 陣列)');
      }

      console.log('=== API RAW JSON OUTPUT ===');
      console.log(JSON.stringify(parsed.translations, null, 2));
      console.log('===========================');

      const normalizedTranslations = parsed.translations.map((item: any, index: number) => {
        const newItem: Record<string, string> = {};
        targetLangs.forEach(lang => {
          let value = '';

          if (item[lang]) {
            value = item[lang];
          } else {
            const keys = Object.keys(item);
            const fuzzyKey = keys.find(k => k.toLowerCase().includes(lang.toLowerCase()) || lang.toLowerCase().includes(k.toLowerCase()));
            if (fuzzyKey) {
              value = item[fuzzyKey];
            } else if (keys.length === 1 && targetLangs.length === 1) {
              value = item[keys];
            }
          }

          const sourceText = texts[index] || '';
          newItem[lang] = repairTranslationTags(sourceText, value || '');
        });
        return newItem;
      });

      return normalizedTranslations;
    } catch (err: any) {
      const isRateLimit = err?.message?.includes('429') || JSON.stringify(err).includes('429');
      const isNetworkError = err?.message?.includes('Load failed') || err?.message?.includes('Failed to fetch') || err?.name === 'TypeError' || err?.message?.includes('NetworkError') || err?.name === 'AbortError';
      const isJsonError = err?.message?.includes('無法解析 API 回傳的 JSON 格式') || err?.message?.includes('API 回傳格式不正確') || err?.message?.includes('Invalid JSON from server');
      const isAuthError = err?.message?.includes('Authentication') || err?.message?.includes('API key') || err?.message?.includes('401') || err?.message?.includes('配置');

      if ((isRateLimit || isNetworkError || isJsonError) && retryCount < 10) {
        const waitTime = isRateLimit
          ? (Math.pow(2, retryCount) * 5000 + Math.random() * 2000)
          : (3000 + Math.random() * 2000);

        console.warn(`DeepSeek ${isRateLimit ? 'Rate limit' : isJsonError ? 'JSON Parse Error' : 'Network error'} hit. Waiting ${Math.round(waitTime / 1000)}s... (Attempt ${retryCount + 1})`);
        await sleep(waitTime);
        return translateBatch(texts, targetLangs, industry, retryCount + 1);
      }

      if (isAuthError) {
        throw new Error(err.message || 'DeepSeek API Key 驗證失敗，請檢查設定是否正確。');
      }

      console.error(`DeepSeek Batch translation error:`, err);
      return texts.map(() => {
        const errorResult: Record<string, string> = {};
        targetLangs.forEach(lang => errorResult[lang] = `(翻譯出錯: ${err?.message || 'API 錯誤'})`);
        return errorResult;
      });
    }
  };

  const cancelTranslation = () => {
    isCancelledRef.current = true;
    setStatus('idle');
    setStatusMessage('翻譯已取消');
    setProgress(0);
    setFileProgress({});
  };

  // The actual processing functions (processDocx, processExcel, processPdf, processPptx)
  // will be defined here or imported.
  // For brevity, I'll export an object that contains the state and functions.

  return {
    status,
    setStatus,
    statusMessage,
    setStatusMessage,
    progress,
    setProgress,
    fileProgress,
    setFileProgress,
    error,
    setError,
    isCancelledRef,
    saveToFirestore,
    translateBatch,
    cancelTranslation
  };
};
