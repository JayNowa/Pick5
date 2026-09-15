// Cloudflare Worker — handles /api/pooltracker proxy, falls through to static assets

const PT_BASE = 'https://www.pooltracker.com/w/season/picks_matrix.asp';
const PT_POOL_ID   = '256835';
const PT_PLAYER_ID = '3593112';
const PT_T         = '1851';

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbw6dH6UV1BVH6Nzm_z_BwDi6oecp4AGcJ-0OkJhhUzSS07QD1cb84LCHjhQiSSremAd6Q/exec';

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

    // Email proxy — forwards JSON POST to Apps Script, handling the POST→GET redirect
    if (url.pathname === '/api/email') {
      if (request.method === 'OPTIONS') {
        return new Response(null, { headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'POST', 'Access-Control-Allow-Headers': 'Content-Type' } });
      }
      try {
        const body = await request.text();
        // Step 1: hit /exec — expect a redirect
        const initial = await fetch(APPS_SCRIPT_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: body,
          redirect: 'manual',
        });
        // Step 2: follow redirect as POST (preserving body)
        const target = initial.headers.get('Location') || APPS_SCRIPT_URL;
        const final = await fetch(target, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: body,
        });
        const text = await final.text();
        return new Response(text, { headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'text/plain' } });
      } catch (e) {
        return new Response(JSON.stringify({ error: e.message }), { status: 500, headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' } });
      }
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
