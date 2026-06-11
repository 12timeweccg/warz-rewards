// POST /api/save  → save shared data to KV (requires x-publish-token == PUBLISH_TOKEN)
export async function onRequestPost(context) {
  const token = (context.request.headers.get('x-publish-token') || '').trim();
  const expected = (context.env.PUBLISH_TOKEN || '').trim();
  if (!expected || token !== expected) {
    return new Response(JSON.stringify({ error: 'unauthorized' }), {
      status: 401, headers: { 'content-type': 'application/json' },
    });
  }

  let body;
  try { body = await context.request.json(); }
  catch { return new Response(JSON.stringify({ error: 'bad json' }), { status: 400, headers: { 'content-type': 'application/json' } }); }

  const savedAt = new Date().toISOString();
  const payload = {
    events: Array.isArray(body.events) ? body.events : [],
    codes:  Array.isArray(body.codes)  ? body.codes  : [],
    items:  Array.isArray(body.items)  ? body.items  : [],
    savedAt,
  };

  try {
    await context.env.warz_data.put('live', JSON.stringify(payload));
  } catch (err) {
    return new Response(JSON.stringify({ error: 'store failed', detail: String(err) }), {
      status: 500, headers: { 'content-type': 'application/json' },
    });
  }

  return new Response(JSON.stringify({
    ok: true, savedAt,
    counts: { events: payload.events.length, codes: payload.codes.length, items: payload.items.length },
  }), { status: 200, headers: { 'content-type': 'application/json' } });
}
