import React, { useState, useRef } from 'react';
import { User } from 'firebase/auth';
import { doc, setDoc, collection, Timestamp } from 'firebase/firestore';
import { db, auth } from '../firebase';
import { jsonrepair } from 'jsonrepair';
import { langBase } from '../utils/documentProcessors';

type TranslationStatus = 'idle' | 'processing' | 'translating' | 'generating' | 'completed' | 'error';

const TRANSLATE_API_URL = '/api/translate';   // 後端：functions/api/translate.ts

// 目前這份檔案的工作單（開始翻譯每份檔案時由 /api/jobs 發出，後端靠它確認已扣過額度）
let currentJobId: string | null = null;
export const setTranslationJob = (jobId: string | null) => { currentJobId = jobId; };

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

const HAN_RE = /[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]/;
const HAN_RE_G = /[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]/g;
const LATIN_RE = /[A-Za-z]/;
const THAI_RE = /[\u0e00-\u0e7f]/;

const containsHan = (text: string) => HAN_RE.test(stripAllTags(text || ''));

const countHan = (text: string) => (text.match(HAN_RE_G) || []).length;
const countVisible = (text: string) => text.replace(/\s/g, '').length;
const hanRatio = (text: string) => countHan(text) / (countVisible(text) || 1);

// 譯文裡允許殘留多少比例的漢字。人名、品牌、料號、分機號碼本來就不該翻，
// 例如「Wakil Manajer 阮敬喬 5116」漢字只佔 15%，屬於正常結果。
const MAX_TARGET_HAN_RATIO = 0.4;

// 譯文若保留了原文八成以上的漢字，代表這段其實沒翻（或只翻了一半）
const MAX_RETAINED_HAN_RATIO = 0.8;

// 譯文裡連續漢字最多允許幾個。人名、品牌通常 2～5 字，
// 連續 6 字以上多半是整句沒翻到。
const MAX_HAN_RUN = 5;

const longestHanRun = (text: string) => {
  const runs: string[] = text.match(/[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]+/g) ?? [];
  return runs.reduce((max: number, run: string) => Math.max(max, run.length), 0);
};

// 譯文裡有沒有出現目標語言該有的字母（拉丁字母／泰文字）
const hasTargetScript = (text: string, base: string) => {
  if (base === 'th') return THAI_RE.test(text);
  if (['id', 'en', 'vi'].includes(base)) return LATIN_RE.test(text);
  return true;
};

const translationNeedsQualityRetry = (
  source: string,
  translated: string,
  targetLang: string
): boolean => {
  const cleanSource = stripAllTags(source || '').trim();
  const cleanTarget = stripAllTags(translated || '').trim();
  if (!cleanTarget) return true;

  // 原文本來就沒有中文（已經是外文、或純數字符號）→ 沒有漏翻的問題
  if (!containsHan(cleanSource)) return false;

  const normalizedSource = cleanSource.replace(/\s+/g, ' ').toLowerCase();
  const normalizedTarget = cleanTarget.replace(/\s+/g, ' ').toLowerCase();

  // 原封不動退回原文：只有在原文「以中文為主」時才算漏翻。
  // 像「• 林逸軒　5102」這種只有姓名和分機的行，照抄回來才是正確結果。
  if (normalizedSource === normalizedTarget) {
    return hanRatio(cleanSource) > MAX_TARGET_HAN_RATIO;
  }

  const base = langBase(targetLang);
  if (!['id', 'en', 'vi', 'th'].includes(base)) return false;

  const sourceHanCount = countHan(cleanSource);
  const targetHanCount = countHan(cleanTarget);
  if (targetHanCount === 0) return false;

  // 原文的漢字幾乎原封不動留在譯文裡 → 這段根本沒翻（或只翻了一半）
  if (sourceHanCount > 0 && targetHanCount / sourceHanCount >= MAX_RETAINED_HAN_RATIO) {
    return true;
  }

  // 譯文裡有一整串沒斷過的中文 → 多半是整句沒翻到
  if (longestHanRun(cleanTarget) > MAX_HAN_RUN) return true;

  // 譯文已經寫成目標語言，只是夾了少量漢字（人名／品牌／料號）→ 視為正常
  if (hasTargetScript(cleanTarget, base) && hanRatio(cleanTarget) <= MAX_TARGET_HAN_RATIO) {
    return false;
  }

  return true;
};

export interface QualityIssue {
  lang: string;
  source: string;
}

export const rebuildWithSourceTags = (source: string, translated: string) => {
  const srcMatches = [...source.matchAll(/\[f(\d+)\]([\s\S]*?)\[\/f\1\]/g)];
  if (srcMatches.length === 0) return translated;

  const emptyTags = () => srcMatches.map(m => `[f${m[1]}][/f${m[1]}]`).join('');

  const cleanTranslated = stripAllTags(translated).trim();
  // 譯文是空的就回傳空標籤。絕對不能退回原文 —— 那會讓中文直接混進泰文／越南文裡。
  if (!cleanTranslated) return emptyTags();

  // 只有「真的有文字」的標籤要分譯文。純空白的標籤是排版用的，
  // 分字給它只會被輸出端丟掉（輸出端會照原文重建空白）。
  const contentIdx = srcMatches
    .map((m, i) => ((m[2] || '').trim().length > 0 ? i : -1))
    .filter(i => i >= 0);

  const assigned: string[] = new Array(srcMatches.length).fill('');

  if (contentIdx.length <= 1) {
    assigned[contentIdx.length === 1 ? contentIdx[0] : 0] = cleanTranslated;
    return srcMatches.map((m, i) => `[f${m[1]}]${assigned[i]}[/f${m[1]}]`).join('');
  }

  const parts = cleanTranslated.split(/\s+/).filter(Boolean);

  if (parts.length < contentIdx.length) {
    // 泰文、日文、中文不用空白分詞，整句切出來往往只有一兩個「詞」。
    // 硬分配會讓後面的標籤分不到字（舊版會因此退回原文，泰文裡才會冒出中文）。
    // 這種情況整句放進第一個有文字的標籤，其餘留空。
    assigned[contentIdx[0]] = cleanTranslated;
  } else {
    const totalSourceChars =
      contentIdx.reduce((sum, i) => sum + (srcMatches[i][2]?.length || 0), 0) || contentIdx.length;
    let cursor = 0;
    contentIdx.forEach((srcIndex, k) => {
      const srcLen = srcMatches[srcIndex][2]?.length || 1;
      let takeCount = Math.round(parts.length * (srcLen / totalSourceChars));
      if (k === contentIdx.length - 1) {
        takeCount = parts.length - cursor;
      } else {
        const remainingSlots = contentIdx.length - k - 1;
        takeCount = Math.max(1, Math.min(takeCount, parts.length - cursor - remainingSlots));
      }
      assigned[srcIndex] = parts.slice(cursor, cursor + takeCount).join(' ');
      cursor += takeCount;
    });
  }

  return srcMatches.map((m, i) => `[f${m[1]}]${assigned[i]}[/f${m[1]}]`).join('');
};

export const repairTranslationTags = (sourceText: string, translatedText: string) => {
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
  // 品質檢查沒過的段落：不再中斷整份翻譯，改成收集起來讓畫面提示人工確認
  const [qualityIssues, setQualityIssues] = useState<QualityIssue[]>([]);
  const qualityIssuesRef = useRef<QualityIssue[]>([]);
  const isCancelledRef = useRef(false);

  const clearQualityIssues = () => {
    qualityIssuesRef.current = [];
    setQualityIssues([]);
  };

  const reportQualityIssues = (issues: QualityIssue[]) => {
    if (issues.length === 0) return;
    const seen = new Set(qualityIssuesRef.current.map(i => `${i.lang}||${i.source}`));
    const added = issues.filter(i => {
      const key = `${i.lang}||${i.source}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
    if (added.length === 0) return;
    qualityIssuesRef.current = [...qualityIssuesRef.current, ...added].slice(0, 200);
    setQualityIssues(qualityIssuesRef.current);
  };

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

  // 把一批文字拆成兩半分別翻，再接回來（順序不變）
  const splitAndTranslate = async (
    texts: string[],
    targetLangs: string[],
    industry: string,
    qualityRetryCount: number
  ): Promise<Record<string, string>[]> => {
    const mid = Math.ceil(texts.length / 2);
    const [head, tail] = await Promise.all([
      translateBatch(texts.slice(0, mid), targetLangs, industry, 0, qualityRetryCount),
      translateBatch(texts.slice(mid), targetLangs, industry, 0, qualityRetryCount),
    ]);
    return [...head, ...tail];
  };

  const translateBatch = async (
    texts: string[],
    targetLangs: string[],
    industry: string,
    retryCount = 0,
    qualityRetryCount = 0
  ): Promise<Record<string, string>[]> => {
    if (texts.length === 0 || targetLangs.length === 0) return texts.map(() => ({}));

    try {
      const body = { texts, languages: targetLangs, industry };   // 翻譯指令由後端組：functions/_lib/prompt.ts

      const controller = new AbortController();
      // Increase timeout to 5 minutes (300000ms) to handle large batches and server-side retries
      const timeoutId = setTimeout(() => controller.abort(), 300000);

      // 後端會驗這個 token；沒有它，任何人都能拿這個端點當免費翻譯 proxy
      const idToken = await auth.currentUser?.getIdToken();
      if (!idToken) {
        clearTimeout(timeoutId);
        throw new Error('登入狀態已失效，請重新登入後再試。');
      }

      const response = await fetch(TRANSLATE_API_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${idToken}`,
          ...(currentJobId ? { 'X-Job-Id': currentJobId } : {}),
        },
        body: JSON.stringify(body),
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
      const resultText = String(data?.choices?.[0]?.message?.content ?? '').trim();

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
              const repairedMatch = jsonrepair(jsonMatch[0]);
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

      if (parsed.translations.length !== texts.length) {
        // 模型偶爾會漏掉（或多回）一筆，整批重送通常還是會漏在同一個地方。
        // 改成拆成兩半分別翻，拆到最後一筆時就不可能對不齊。
        if (texts.length > 1) {
          console.warn(
            `[Translation] 回傳筆數不一致：要求 ${texts.length} 筆、收到 ${parsed.translations.length} 筆，改成拆半重翻`
          );
          return await splitAndTranslate(texts, targetLangs, industry, qualityRetryCount);
        }
        throw new Error(
          `API 回傳筆數不一致：要求 ${texts.length} 筆，實際收到 ${parsed.translations.length} 筆`
        );
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
              value = item[keys[0]];   // 原本是 item[keys]，把陣列當成 key 用
            }
          }

          const sourceText = texts[index] || '';
          newItem[lang] = repairTranslationTags(sourceText, value || '');
        });
        return newItem;
      });

      // 中文殘留或空白譯文只重翻有問題的段落，最多兩次。
      const MAX_QUALITY_RETRIES = 2;
      if (qualityRetryCount < MAX_QUALITY_RETRIES) {
        for (const lang of targetLangs) {
          const badIndices = normalizedTranslations
            .map((item, index) =>
              translationNeedsQualityRetry(texts[index] || '', item[lang] || '', lang)
                ? index
                : -1
            )
            .filter(index => index >= 0);

          if (badIndices.length === 0) continue;

          console.warn(
            `[Translation QA] ${lang} 有 ${badIndices.length} 段不完整，進行第 ${qualityRetryCount + 1} 次重翻：`,
            badIndices
          );

          const retried = await translateBatch(
            badIndices.map(index => texts[index]),
            [lang],
            industry,
            0,
            qualityRetryCount + 1
          );

          badIndices.forEach((originalIndex, retryIndex) => {
            const retryValue = retried[retryIndex]?.[lang];
            if (retryValue) normalizedTranslations[originalIndex][lang] = retryValue;
          });
        }
      }

      // 兩次重翻之後還是沒過的段落：保留現有內容（空白就退回原文），
      // 並記錄下來讓畫面提示，不再整份中斷。
      const unresolved: QualityIssue[] = [];
      normalizedTranslations.forEach((item, index) => {
        const sourceText = texts[index] || '';
        targetLangs.forEach(lang => {
          if (translationNeedsQualityRetry(sourceText, item[lang] || '', lang)) {
            if (!stripAllTags(item[lang] || '').trim()) {
              item[lang] = sourceText;
            }
            unresolved.push({ lang, source: stripAllTags(sourceText).trim().slice(0, 60) });
          }
        });
      });

      if (unresolved.length > 0) {
        console.warn('[Translation QA] 以下段落品質檢查未通過，已保留現有內容：', unresolved);
        reportQualityIssues(unresolved);
      }

      return normalizedTranslations;
    } catch (err: any) {
      const isRateLimit = err?.message?.includes('429') || JSON.stringify(err).includes('429');
      const isNetworkError = err?.message?.includes('Load failed') || err?.message?.includes('Failed to fetch') || err?.name === 'TypeError' || err?.message?.includes('NetworkError') || err?.name === 'AbortError';
      const isJsonError = err?.message?.includes('無法解析 API 回傳的 JSON 格式') || err?.message?.includes('API 回傳格式不正確') || err?.message?.includes('API 回傳筆數不一致') || err?.message?.includes('翻譯回傳筆數不一致') || err?.message?.includes('Invalid JSON from server');
      const isAuthError = err?.message?.includes('Authentication') || err?.message?.includes('API key') || err?.message?.includes('401') || err?.message?.includes('配置');

      // 這一批字數超過後端單次上限：拆半再送
      if (err?.message?.includes('過長') && texts.length > 1) {
        return await splitAndTranslate(texts, targetLangs, industry, qualityRetryCount);
      }

      if ((isRateLimit || isNetworkError || isJsonError) && retryCount < 10) {
        const waitTime = isRateLimit
          ? (Math.pow(2, retryCount) * 5000 + Math.random() * 2000)
          : (3000 + Math.random() * 2000);

        console.warn(`Translate API ${isRateLimit ? 'Rate limit' : isJsonError ? 'JSON Parse Error' : 'Network error'} hit. Waiting ${Math.round(waitTime / 1000)}s... (Attempt ${retryCount + 1})`);
        await sleep(waitTime);
        return translateBatch(texts, targetLangs, industry, retryCount + 1, qualityRetryCount);
      }

      if (isJsonError) {
        // 結構不完整時，先拆半再試一次；真的只剩一筆還是壞的，
        // 就保留原文並列入「建議人工確認」，不要讓整份檔案直接失敗。
        if (texts.length > 1) {
          console.warn('[Translation] 結構持續異常，改成拆半重翻');
          return await splitAndTranslate(texts, targetLangs, industry, qualityRetryCount);
        }
        console.warn('[Translation] 單筆結構持續異常，保留原文：', texts[0]);
        reportQualityIssues(
          targetLangs.map(lang => ({
            lang,
            source: stripAllTags(texts[0] || '').trim().slice(0, 60),
          }))
        );
        return [Object.fromEntries(targetLangs.map(lang => [lang, texts[0] || ''])) as Record<string, string>];
      }

      if (isAuthError) {
        throw new Error(err.message || '翻譯服務金鑰驗證失敗，請檢查設定是否正確。');
      }

      console.error(`Batch translation error:`, err);
      throw err instanceof Error ? err : new Error(err?.message || '翻譯服務發生未知錯誤');
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
    cancelTranslation,
    qualityIssues,
    clearQualityIssues
  };
};
