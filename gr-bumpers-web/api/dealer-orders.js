// Vercel serverless function: /api/dealer-orders
// Every request must include a valid session token. The server looks up which
// dealer that token belongs to and only ever reads/writes that dealer's own
// orders — the dealer's browser never sees anyone else's data, inventory
// counts, or model stock levels.

import { Redis } from '@upstash/redis';

const url = process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL;
const token = process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN;
const redis = url && token ? new Redis({ url, token }) : null;

const ORDERS_KEY = 'gr-bumpers:bumper-orders';
const SESSIONS_KEY = 'gr-bumpers:dealer-sessions';
const INVENTORY_KEY = 'gr-bumpers:bumper-inventory';

async function dealerForToken(sessionToken) {
  if (!sessionToken) return null;
  const sessions = (await redis.get(SESSIONS_KEY)) || {};
  const session = sessions[sessionToken];
  return session ? session.dealerName : null;
}

export default async function handler(req, res) {
  if (!redis) {
    return res.status(500).json({ error: 'Redis credentials not configured.' });
  }

  try {
    if (req.method === 'GET') {
      const { token: sessionToken } = req.query;
      const dealerName = await dealerForToken(sessionToken);
      if (!dealerName) return res.status(401).json({ error: 'Not logged in.' });

      const allOrders = (await redis.get(ORDERS_KEY)) || [];
      const mine = allOrders.filter(o => (o.dealer || '').trim().toLowerCase() === dealerName.trim().toLowerCase());
      return res.status(200).json({ dealerName, orders: mine });
    }

    if (req.method === 'POST') {
      const { token: sessionToken, sku, qty, po } = req.body || {};
      const dealerName = await dealerForToken(sessionToken);
      if (!dealerName) return res.status(401).json({ error: 'Not logged in.' });

      const cleanQty = parseInt(qty, 10);
      if (!sku || !cleanQty || cleanQty <= 0) return res.status(400).json({ error: 'A model and a valid quantity are required.' });

      // Confirm the model is real, but never send inventory counts back to the dealer.
      const inventory = (await redis.get(INVENTORY_KEY)) || [];
      if (!inventory.some(r => r.sku === sku)) return res.status(400).json({ error: 'Unknown model.' });

      const allOrders = (await redis.get(ORDERS_KEY)) || [];
      const nextId = 'D' + Date.now();
      const today = new Date();
      const dateStr = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;
      const newOrder = {
        id: nextId, sku, dealer: dealerName, date: dateStr, due: '', num: '',
        po: po || '', qty: cleanQty, invoiced: 0, backordered: cleanQty,
      };
      await redis.set(ORDERS_KEY, [newOrder, ...allOrders]);
      return res.status(200).json({ ok: true, order: newOrder });
    }

    res.setHeader('Allow', 'GET, POST');
    return res.status(405).json({ error: 'method not allowed' });
  } catch (err) {
    console.error('dealer-orders error', err);
    return res.status(500).json({ error: String(err) });
  }
}
