// Vercel serverless function: GET/POST/DELETE /api/storage
// Backed by Upstash Redis. Every request must include a valid admin session token
// (issued by /api/admin-auth) — this is the main site's actual data, so nobody
// should be able to read or write it without having logged in first.

import { redis, NS, isValidAdminToken } from '../lib/serverAuth.js';

export default async function handler(req, res) {
  // Belt-and-suspenders: prevent any browser/CDN layer from caching this response,
  // since these values change frequently and every client needs the live data.
  res.setHeader('Cache-Control', 'no-store, must-revalidate');

  if (!redis) {
    return res.status(500).json({
      error: 'Redis credentials not found. Check that UPSTASH_REDIS_REST_URL/TOKEN ' +
        '(or KV_REST_API_URL/TOKEN) are set in this Vercel project\'s Environment ' +
        'Variables, then redeploy.',
    });
  }

  const adminToken = req.method === 'GET' || req.method === 'DELETE' ? req.query.token : (req.body || {}).token;
  const authorized = await isValidAdminToken(adminToken);
  if (!authorized) {
    return res.status(401).json({ error: 'Not authorized. Please log in.' });
  }

  try {
    if (req.method === 'GET') {
      const { key } = req.query;
      if (!key) return res.status(400).json({ error: 'key is required' });
      const value = await redis.get(NS + key);
      if (value === null || value === undefined) {
        return res.status(404).json({ error: 'not found' });
      }
      return res.status(200).json({ key, value });
    }

    if (req.method === 'POST') {
      const { key, value } = req.body || {};
      if (!key) return res.status(400).json({ error: 'key is required' });
      await redis.set(NS + key, value);
      return res.status(200).json({ key, value });
    }

    if (req.method === 'DELETE') {
      const { key } = req.query;
      if (!key) return res.status(400).json({ error: 'key is required' });
      await redis.del(NS + key);
      return res.status(200).json({ key, deleted: true });
    }

    res.setHeader('Allow', 'GET, POST, DELETE');
    return res.status(405).json({ error: 'method not allowed' });
  } catch (err) {
    console.error('storage api error', err);
    return res.status(500).json({ error: String(err) });
  }
}
