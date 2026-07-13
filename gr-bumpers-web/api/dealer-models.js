// Vercel serverless function: /api/dealer-models
// Returns just the list of model codes for the order form dropdown — never
// quantities, so a dealer can't infer your stock levels from this.

import { Redis } from '@upstash/redis';

const url = process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL;
const token = process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN;
const redis = url && token ? new Redis({ url, token }) : null;

const INVENTORY_KEY = 'gr-bumpers:bumper-inventory';

export default async function handler(req, res) {
  if (!redis) {
    return res.status(500).json({ error: 'Redis credentials not configured.' });
  }
  if (req.method !== 'GET') {
    res.setHeader('Allow', 'GET');
    return res.status(405).json({ error: 'method not allowed' });
  }
  try {
    const inventory = (await redis.get(INVENTORY_KEY)) || [];
    const skus = inventory.map(r => r.sku).sort();
    return res.status(200).json({ skus });
  } catch (err) {
    console.error('dealer-models error', err);
    return res.status(500).json({ error: String(err) });
  }
}
