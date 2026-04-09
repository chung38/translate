import React, { useState, useRef } from 'react';
import { User } from 'firebase/auth';
import { doc, setDoc, collection, Timestamp } from 'firebase/firestore';
import { db } from '../firebase';
import { jsonrepair } from 'jsonrepair';

type TranslationStatus = 'idle' | 'processing' | 'translating' | 'generating' | 'completed' | 'error';

const DEEPSEEK_PROXY_URL = '/api/translate';

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

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
      console.error("Error saving to Firestore:", err);
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
      5. 不要包含任何 Markdown 標籤（如 \`\`\`json）或額外文字，只回傳純 JSON 字串。

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
      
      const cleanJson = resultText.replace(/^```json\n?/, '').replace(/\n?```$/, '').trim();
      
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

      const normalizedTranslations = parsed.translations.map((item: any) => {
        const newItem: any = {};
        targetLangs.forEach(lang => {
          if (item[lang]) {
            newItem[lang] = item[lang];
          } else {
            const keys = Object.keys(item);
            const fuzzyKey = keys.find(k => k.toLowerCase().includes(lang.toLowerCase()) || lang.toLowerCase().includes(k.toLowerCase()));
            if (fuzzyKey) {
              newItem[lang] = item[fuzzyKey];
            } else if (keys.length === 1 && targetLangs.length === 1) {
              newItem[lang] = item[keys[0]];
            }
          }
        });
        return newItem;
      });

      return normalizedTranslations;
    } catch (err: any) {
      const isRateLimit = err?.message?.includes('429') || JSON.stringify(err).includes('429');
      const isNetworkError = err?.message?.includes('Load failed') || err?.message?.includes('Failed to fetch') || err?.name === 'TypeError' || err?.message?.includes('NetworkError') || err?.name === 'AbortError';
      const isJsonError = err?.message?.includes('無法解析 API 回傳的 JSON 格式') || err?.message?.includes('API 回傳格式不正確');
      const isAuthError = err?.message?.includes('Authentication') || err?.message?.includes('API key') || err?.message?.includes('401') || err?.message?.includes('配置');
      
      if ((isRateLimit || isNetworkError || isJsonError) && retryCount < 10) {
        const waitTime = isRateLimit 
          ? (Math.pow(2, retryCount) * 5000 + Math.random() * 2000)
          : (3000 + Math.random() * 2000);
          
        console.warn(`DeepSeek ${isRateLimit ? 'Rate limit' : isJsonError ? 'JSON Parse Error' : 'Network error'} hit. Waiting ${Math.round(waitTime/1000)}s... (Attempt ${retryCount + 1})`);
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
