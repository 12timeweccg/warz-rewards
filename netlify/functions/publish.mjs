import { getStore } from '@netlify/blobs';

// POST /.netlify/functions/publish
// Header: x-publish-token must match the PUBLISH_TOKEN env var.
// Body: { events, codes, items }
// Writes the data to the "warz" blob store under key "live".
export default async (req) => {
  if (req.method !== 'POST') {
    return new Response('Method not allowed', { status: 405 });
  }

  const token = req.headers.get('x-publish-token');
  const expected = process.env.PUBLISH_TOKEN;
  if (!expected || token !== expected) {
    return new Response(JSON.stringify({ error: 'unauthorized' }), {
      status: 401,
      headers: { 'content-type': 'application/json' },
    });
  }

  let body;
  try {
    body = await req.json();
  } catch {
    return new Response(JSON.stringify({ error: 'bad json' }), {
      status: 400,
      headers: { 'content-type': 'application/json' },
    });
  }

  const publishedAt = new Date().toISOString();
  const payload = {
    events: Array.isArray(body.events) ? body.events : [],
    codes:  Array.isArray(body.codes)  ? body.codes  : [],
    items:  Array.isArray(body.items)  ? body.items  : [],
    publishedAt,
  };

  try {
    const store = getStore('warz');
    await store.setJSON('live', payload);
  } catch (err) {
    return new Response(JSON.stringify({ error: 'store failed', detail: String(err) }), {
      status: 500,
      headers: { 'content-type': 'application/json' },
    });
  }

  return new Response(JSON.stringify({
    ok: true,
    publishedAt,
    counts: { events: payload.events.length, codes: payload.codes.length, items: payload.items.length },
  }), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
};
