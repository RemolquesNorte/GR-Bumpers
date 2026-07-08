// Vercel serverless function: GET/POST /api/storage
// Backed by Upstash Redis. Vercel's Upstash Marketplace integration can name the
// injected credentials one of a couple ways depending on how the database was
// created, so we check both:
//   UPSTASH_REDIS_REST_URL / UPSTASH_REDIS_REST_TOKEN   (Upstash's own naming)
//   KV_REST_API_URL / KV_REST_API_TOKEN                  (legacy Vercel KV naming,
//                                                          still used by some
//                                                          Marketplace flows)

import { Redis } from '@upstash/redis';

const url = process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL;
const token = process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN;

const redis = url && token ? new Redis({ url, token }) : null;

// Namespaced so this Redis database could safely be reused by other projects later.
const NS = 'gr-bumpers:';

export default async function handler(req, res) {
  if (!redis) {
    return res.status(500).json({
      error: 'Redis credentials not found. Check that UPSTASH_REDIS_REST_URL/TOKEN ' +
        '(or KV_REST_API_URL/TOKEN) are set in this Vercel project\'s Environment ' +
        'Variables, then redeploy.',
    });
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
