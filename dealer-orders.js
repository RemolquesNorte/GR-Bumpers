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
const PENDING_KEY = 'gr-bumpers:bumper-new-orders';
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

      const [allOrders, allPending] = await Promise.all([
        redis.get(ORDERS_KEY),
        redis.get(PENDING_KEY),
      ]);
      const matches = o => (o.dealer || '').trim().toLowerCase() === dealerName.trim().toLowerCase();
      const mineOpen = (allOrders || []).filter(matches).map(o => ({ ...o, pending: false }));
      const minePending = (allPending || []).filter(matches).map(o => ({ ...o, pending: true }));
      return res.status(200).json({ dealerName, orders: [...minePending, ...mineOpen] });
    }

    if (req.method === 'POST') {
      const { token: sessionToken, po, items } = req.body || {};
      const dealerName = await dealerForToken(sessionToken);
      if (!dealerName) return res.status(401).json({ error: 'Not logged in.' });

      const lines = Array.isArray(items) ? items : [];
      const cleanLines = lines
        .map(l => ({ sku: l.sku, qty: parseInt(l.qty, 10) }))
        .filter(l => l.sku && l.qty > 0);
      if (cleanLines.length === 0) return res.status(400).json({ error: 'At least one model and quantity are required.' });

      // Confirm every model is real, but never send inventory counts back to the dealer.
      const inventory = (await redis.get(INVENTORY_KEY)) || [];
      const validSkus = new Set(inventory.map(r => r.sku));
      const unknown = cleanLines.find(l => !validSkus.has(l.sku));
      if (unknown) return res.status(400).json({ error: `Unknown model: ${unknown.sku}` });

      const allPending = (await redis.get(PENDING_KEY)) || [];
      const today = new Date();
      const dateStr = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;
      const newOrders = cleanLines.map((l, i) => ({
        id: 'D' + Date.now() + '_' + i, sku: l.sku, dealer: dealerName, date: dateStr, due: '', num: '',
        po: po || '', qty: l.qty, invoiced: 0, backordered: l.qty,
      }));
      await redis.set(PENDING_KEY, [...newOrders, ...allPending]);
      return res.status(200).json({ ok: true, orders: newOrders, pending: true });
    }

    res.setHeader('Allow', 'GET, POST');
    return res.status(405).json({ error: 'method not allowed' });
  } catch (err) {
    console.error('dealer-orders error', err);
    return res.status(500).json({ error: String(err) });
  }
}
