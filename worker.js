// Cloudflare Worker — handles /api/pooltracker proxy, falls through to static assets

const PT_BASE = 'https://www.pooltracker.com/w/season/picks_matrix.asp';
const PT_POOL_ID   = '256835';
const PT_PLAYER_ID = '3593112';
const PT_T         = '1851';

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === 'OPTIONS') {
      return new Response(null, {
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET',
          'Access-Control-Max-Age': '86400',
        },
      });
    }

    if (url.pathname === '/api/pooltracker') {
      const week = parseInt(url.searchParams.get('week') || '1');
      if (isNaN(week) || week < 1 || week > 22) {
        return new Response(JSON.stringify({ error: 'Invalid week (1-22)' }), {
          status: 400,
          headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
        });
      }

      const ptUrl = `${PT_BASE}?poolid=${PT_POOL_ID}&playerid=${PT_PLAYER_ID}&week=${week}&t=${PT_T}`;

      try {
        const resp = await fetch(ptUrl, {
          headers: {
            'User-Agent':
              'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
            Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
          },
        });

        if (!resp.ok) {
          return new Response(
            JSON.stringify({ error: 'pooltracker.com returned HTTP ' + resp.status }),
            {
              status: 502,
              headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
            }
          );
        }

        const html = await resp.text();

        return new Response(html, {
          headers: {
            'Content-Type': 'text/html; charset=utf-8',
            'Access-Control-Allow-Origin': '*',
            'X-Pooltracker-Week': String(week),
            // Don't let browsers cache the proxy result — scores change
            'Cache-Control': 'no-store',
          },
        });
      } catch (e) {
        return new Response(JSON.stringify({ error: e.message }), {
          status: 500,
          headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
        });
      }
    }

    // All other requests → static assets (index.html, manifest.json, etc.)
    return env.ASSETS.fetch(request);
  },
};
