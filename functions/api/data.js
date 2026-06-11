// GET /api/data  → return the shared published data from KV, or {empty:true}
export async function onRequestGet(context) {
  try {
    const data = await context.env.warz_data.get('live', { type: 'json' });
    if (!data) return json({ empty: true });
    return json(data);
  } catch (err) {
    return json({ empty: true });
  }
}

function json(obj) {
  return new Response(JSON.stringify(obj), {
    status: 200,
    headers: { 'content-type': 'application/json', 'cache-control': 'no-store' },
  });
}
