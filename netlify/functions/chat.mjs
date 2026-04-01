// Netlify Serverless Function — Claude API Proxy
// Environment variable: CLAUDE_API_KEY (set in Netlify dashboard)

export default async (req, context) => {
  // CORS preflight
  if (req.method === 'OPTIONS') {
    return new Response('', {
      status: 204,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
      },
    });
  }

  if (req.method !== 'POST') {
    return new Response(JSON.stringify({ error: 'POST only' }), {
      status: 405,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
    });
  }

  // Support both Netlify Functions v2 (Netlify.env) and v1/fallback (process.env)
  const apiKey = (typeof Netlify !== 'undefined' && Netlify.env)
    ? Netlify.env.get('CLAUDE_API_KEY')
    : process.env.CLAUDE_API_KEY;

  if (!apiKey) {
    return new Response(JSON.stringify({ error: 'CLAUDE_API_KEY not configured. Set it in Netlify dashboard → Site configuration → Environment variables.' }), {
      status: 500,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
    });
  }

  try {
    const body = await req.json();
    const { model, max_tokens, system, messages } = body;

    // Use AbortController for client-side timeout (25s to stay under Netlify's 26s limit)
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 25000);

    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: model || 'claude-sonnet-4-20250514',
        max_tokens: Math.min(max_tokens || 4096, 8192),
        system,
        messages,
      }),
      signal: controller.signal,
    });

    clearTimeout(timeout);

    const data = await resp.text();

    return new Response(data, {
      status: resp.status,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
      },
    });
  } catch (e) {
    const isTimeout = e.name === 'AbortError';
    return new Response(JSON.stringify({
      error: isTimeout
        ? 'Claude API 응답 시간 초과 (25초). 더 짧은 요청을 시도하거나 Haiku 모델을 사용해보세요.'
        : e.message
    }), {
      status: isTimeout ? 504 : 500,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
    });
  }
};

export const config = {
  path: '/.netlify/functions/chat',
  preferStatic: true,
};
