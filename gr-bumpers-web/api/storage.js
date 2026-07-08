// Vercel serverless function: GET/POST /api/storage
// Backed by Upstash Redis. Requires these env vars to be set in your Vercel project:
//   UPSTASH_REDIS_REST_URL
//   UPSTASH_REDIS_REST_TOKEN
// (These are added automatically if you connect an Upstash Redis database to your
// project from the Vercel dashboard's Storage tab. Otherwise copy them from your
// Upstash console into Vercel's Environment Variables settings.)

import { Redis } from '@upstash/redis';

const redis = Redis.fromEnv();

// Namespaced so this Redis database could safely be reused by other projects later.
const NS = 'gr-bumpers:';

export default async function handler(req, res) {
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
