// Cloudflare Pages Function — /api/translate
// 自動對應路由：https://your-domain/api/translate

interface Env {
  DEEPSEEK_API_KEY: string;
}

export const onRequestPost: PagesFunction<Env> = async (context) => {
  const { request, env } = context;

  // CORS headers
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };

  // Handle preflight
  if (request.method === 'OPTIONS') {
    return new Response(null, { status: 204, headers: corsHeaders });
  }

  const apiKey = env.DEEPSEEK_API_KEY;
  if (!apiKey) {
    return new Response(
      JSON.stringify({ error: { message: '未設定 DEEPSEEK_API_KEY，請至 Cloudflare Pages → Settings → Environment variables 新增。' } }),
      { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    );
  }

  let prompt: string;
  try {
    const body = await request.json() as { prompt?: string };
    prompt = body?.prompt ?? '';
  } catch {
    return new Response(
      JSON.stringify({ error: { message: '無效的請求內容' } }),
      { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    );
  }

  if (!prompt || typeof prompt !== 'string') {
    return new Response(
      JSON.stringify({ error: { message: '無效的請求內容' } }),
      { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    );
  }

  if (prompt.length > 200000) {
    return new Response(
      JSON.stringify({ error: { message: '文本內容過長，超過系統單次處理限制。' } }),
      { status: 400, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    );
  }

  // 呼叫 DeepSeek，最多重試 3 次
  let retries = 3;
  let delay = 1000;

  while (retries > 0) {
    try {
      const response = await fetch('https://api.deepseek.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${apiKey}`,
        },
        body: JSON.stringify({
          model: 'deepseek-chat',
          messages: [
            { role: 'system', content: 'You are a helpful assistant that translates text into multiple languages and outputs only structured JSON.' },
            { role: 'user', content: prompt },
          ],
          response_format: { type: 'json_object' },
          temperature: 0.3,
          max_tokens: 8192,
        }),
      });

      const responseText = await response.text();

      if (!response.ok) {
        if ((response.status === 429 || response.status >= 500) && retries > 1) {
          await new Promise(r => setTimeout(r, delay));
          retries--;
          delay *= 2;
          continue;
        }
        let errorData;
        try { errorData = JSON.parse(responseText); }
        catch { errorData = { error: { message: `DeepSeek API Error (${response.status})` } }; }
        return new Response(JSON.stringify(errorData), {
          status: response.status,
          headers: { ...corsHeaders, 'Content-Type': 'application/json' },
        });
      }

      return new Response(responseText, {
        status: 200,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });

    } catch (err: any) {
      if (retries > 1) {
        await new Promise(r => setTimeout(r, delay));
        retries--;
        delay *= 2;
        continue;
      }
      return new Response(
        JSON.stringify({ error: { message: err.message || 'Internal Server Error' } }),
        { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
      );
    }
  }

  return new Response(
    JSON.stringify({ error: { message: '翻譯請求失敗，請稍後重試。' } }),
    { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
  );
};

// 同時處理 OPTIONS preflight
export const onRequestOptions: PagesFunction = async () => {
  return new Response(null, {
    status: 204,
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    },
  });
};
