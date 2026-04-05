// Netlify Edge Function — Claude API Proxy with Streaming
// Runs on Deno runtime. Up to 50s execution time on Free plan.
// Env var: CLAUDE_API_KEY (set in Netlify dashboard)

export default async (req: Request) => {
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
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
      },
    });
  }

  // Deno: Netlify.env.get or Deno.env.get
  // @ts-ignore
  const apiKey = (typeof Netlify !== 'undefined' && Netlify.env)
    // @ts-ignore
    ? Netlify.env.get('CLAUDE_API_KEY')
    // @ts-ignore
    : Deno.env.get('CLAUDE_API_KEY');

  if (!apiKey) {
    return new Response(JSON.stringify({
      error: 'CLAUDE_API_KEY not configured. Set it in Netlify dashboard → Site configuration → Environment variables.'
    }), {
      status: 500,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
      },
    });
  }

  try {
    const body = await req.json();
    const { model, max_tokens, system, messages } = body;

    // Request streaming response from Anthropic
    const anthropicResp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: model || 'claude-sonnet-4-5-20250929',
        max_tokens: Math.min(max_tokens || 4096, 8192),
        system,
        messages,
        stream: true,
      }),
    });

    if (!anthropicResp.ok) {
      const errText = await anthropicResp.text();
      return new Response(JSON.stringify({
        error: `Anthropic API 오류 (${anthropicResp.status}): ${errText.substring(0, 300)}`
      }), {
        status: anthropicResp.status,
        headers: {
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*',
        },
      });
    }

    // Pipe the SSE stream directly to the client
    return new Response(anthropicResp.body, {
      status: 200,
      headers: {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache, no-transform',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*',
        'X-Accel-Buffering': 'no',
      },
    });
  } catch (e) {
    return new Response(JSON.stringify({
      error: e instanceof Error ? e.message : String(e)
    }), {
      status: 500,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
      },
    });
  }
};

export const config = {
  path: '/api/chat',
};
