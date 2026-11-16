// ─── 翻譯模型 ─────────────────────────────────────────────────────────────
// 只用 OpenAI（預設 gpt-6-luna，可用 OPENAI_MODEL 換）。隱私權政策寫的也是 OpenAI，
// 要換或加翻譯服務商時，記得一起更新 public/privacy.html。
export const DEFAULT_OPENAI_MODEL = 'gpt-6-luna';
const SYSTEM_PROMPT = 'You are a helpful assistant that translates text into multiple languages and outputs only structured JSON.';

export interface Provider {
  name: 'openai';
  url: string;
  key: string;
  model: string;
  options: Record<string, unknown>;   // 除了 model、messages 以外的參數
}

export function pickProvider(env: { OPENAI_API_KEY?: string; OPENAI_MODEL?: string }): Provider | null {
  if (env.OPENAI_API_KEY) {
    return {
      name: 'openai',
      url: 'https://api.openai.com/v1/chat/completions',
      key: env.OPENAI_API_KEY,
      model: env.OPENAI_MODEL || DEFAULT_OPENAI_MODEL,
      options: {
        reasoning_effort: 'none',   // 翻譯不需要「思考」；開著會變慢，推理用的 token 也照輸出計費
        temperature: 0.3,           // 只有 reasoning_effort 為 none 時才能設
        response_format: { type: 'json_object' },
        max_completion_tokens: 8192,
      },
    };
  }
  return null;
}

// 模型不接受某個選填參數（回 400）時，自動調整後重送，並記住，之後的請求直接用調整後的設定。
const ADJUSTABLE = ['temperature', 'reasoning_effort', 'response_format'] as const;
const learned = new Map<string, Record<string, unknown>>();   // 模型名稱 → 調整後的參數

export function adjustForError(options: Record<string, unknown>, status: number, errorBody: any): Record<string, unknown> | null {
  if (status !== 400) return null;
  const param = String(errorBody?.error?.param ?? '');
  const message = String(errorBody?.error?.message ?? '');
  const bad = ADJUSTABLE.find(k => k in options && (param === k || message.includes(k)));
  if (!bad) return null;
  const next = { ...options };
  if (bad === 'reasoning_effort' && next.reasoning_effort === 'none') {
    next.reasoning_effort = 'low';   // 不支援 none 的模型，用最低的 low
  } else {
    delete next[bad];
    if (bad === 'reasoning_effort') delete next.temperature;   // 有在推理就不能設溫度
  }
  return next;
}

export async function callModel(provider: Provider, prompt: string): Promise<Response> {
  let options = learned.get(provider.model) ?? provider.options;
  let adjustments = 0;
  let retries = 3;
  let delay = 1000;

  while (true) {
    let response: Response;
    try {
      response = await fetch(provider.url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${provider.key}` },
        body: JSON.stringify({
          model: provider.model,
          messages: [
            { role: 'system', content: SYSTEM_PROMPT },
            { role: 'user', content: prompt },
          ],
          ...options,
        }),
      });
    } catch (err) {
      if (--retries > 0) {
        await new Promise(r => setTimeout(r, delay));
        delay *= 2;
        continue;
      }
      throw err;
    }

    if (response.ok) return response;

    const text = await response.text();
    let body: any = null;
    try { body = JSON.parse(text); } catch { /* 不是 JSON */ }

    const fixed = adjustments < ADJUSTABLE.length ? adjustForError(options, response.status, body) : null;
    if (fixed) {
      console.warn(`${provider.model} 不接受目前的參數，已自動調整：`, body?.error?.message);
      options = fixed;
      learned.set(provider.model, fixed);
      adjustments++;
      continue;
    }

    if ((response.status === 429 || response.status >= 500) && --retries > 0) {
      await new Promise(r => setTimeout(r, delay));
      delay *= 2;
      continue;
    }

    return new Response(text, { status: response.status, headers: { 'Content-Type': 'application/json' } });
  }
}

