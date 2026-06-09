import { getStore } from '@netlify/blobs';

// GET /.netlify/functions/data
// Returns the latest published data, or { empty: true } if nothing has been
// published yet (the public site then falls back to the baked-in events-data.js).
const emptyResponse = () =>
  new Response(JSON.stringify({ empty: true }), {
    status: 200,
    headers: { 'content-type': 'application/json', 'cache-control': 'no-store' },
  });

export default async () => {
  try {
    const store = getStore('warz');
    const data = await store.get('live', { type: 'json' });
    if (!data) return emptyResponse();
    return new Response(JSON.stringify(data), {
      status: 200,
      headers: { 'content-type': 'application/json', 'cache-control': 'no-store' },
    });
  } catch (err) {
    // On any error, behave as "nothing published" so the site uses its fallback
    return emptyResponse();
  }
};
