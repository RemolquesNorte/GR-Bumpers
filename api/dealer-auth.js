// Vercel serverless function: /api/dealer-auth
// Handles dealer login, and lets the internal (staff) app set/reset a dealer's
// username + password. Uses Node's built-in crypto (scrypt) for hashing — no
// extra dependency needed.
//
// Storage shape (Redis key "gr-bumpers:dealer-accounts"):
//   { [username]: { dealerName, salt, hash } }
// Storage shape (Redis key "gr-bumpers:dealer-sessions"):
//   { [token]: { username, dealerName, createdAt } }

import { Redis } from '@upstash/redis';
import crypto from 'crypto';

const url = process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL;
const token = process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN;
const redis = url && token ? new Redis({ url, token }) : null;

const ACCOUNTS_KEY = 'gr-bumpers:dealer-accounts';
const SESSIONS_KEY = 'gr-bumpers:dealer-sessions';

function hashPassword(password, salt) {
  return crypto.scryptSync(password, salt, 64).toString('hex');
}

function makeToken() {
  return crypto.randomBytes(24).toString('hex');
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, must-revalidate');
  if (!redis) {
    return res.status(500).json({ error: 'Redis credentials not configured.' });
  }
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ error: 'method not allowed' });
  }

  const { action } = req.body || {};

  try {
    // --- Dealer login ---
    if (action === 'login') {
      const { username, password } = req.body;
      if (!username || !password) return res.status(400).json({ error: 'Username and password required.' });
      const accounts = (await redis.get(ACCOUNTS_KEY)) || {};
      const account = accounts[username.trim().toLowerCase()];
      if (!account) return res.status(401).json({ error: 'Incorrect username or password.' });
      const hash = hashPassword(password, account.salt);
      if (hash !== account.hash) return res.status(401).json({ error: 'Incorrect username or password.' });

      const sessions = (await redis.get(SESSIONS_KEY)) || {};
      const sessionToken = makeToken();
      sessions[sessionToken] = { username: username.trim().toLowerCase(), dealerName: account.dealerName, createdAt: Date.now() };
      await redis.set(SESSIONS_KEY, sessions);

      return res.status(200).json({ token: sessionToken, dealerName: account.dealerName });
    }

    // --- Validate an existing session token (used by the portal on reload) ---
    if (action === 'session') {
      const { token: sessionToken } = req.body;
      if (!sessionToken) return res.status(400).json({ error: 'Token required.' });
      const sessions = (await redis.get(SESSIONS_KEY)) || {};
      const session = sessions[sessionToken];
      if (!session) return res.status(401).json({ error: 'Session expired or invalid.' });
      return res.status(200).json({ dealerName: session.dealerName });
    }

    // --- Log out ---
    if (action === 'logout') {
      const { token: sessionToken } = req.body;
      if (sessionToken) {
        const sessions = (await redis.get(SESSIONS_KEY)) || {};
        delete sessions[sessionToken];
        await redis.set(SESSIONS_KEY, sessions);
      }
      return res.status(200).json({ ok: true });
    }

    // --- Admin: set or reset a dealer's login (called from the staff app) ---
    if (action === 'set-credentials') {
      const { dealerName, username, password } = req.body;
      if (!dealerName || !username || !password) {
        return res.status(400).json({ error: 'dealerName, username, and password are required.' });
      }
      const cleanUsername = username.trim().toLowerCase();
      const accounts = (await redis.get(ACCOUNTS_KEY)) || {};
      // If this dealer already had a different username, drop the old one.
      for (const [u, acc] of Object.entries(accounts)) {
        if (acc.dealerName === dealerName && u !== cleanUsername) delete accounts[u];
      }
      const salt = crypto.randomBytes(16).toString('hex');
      accounts[cleanUsername] = { dealerName, salt, hash: hashPassword(password, salt) };
      await redis.set(ACCOUNTS_KEY, accounts);
      return res.status(200).json({ ok: true, username: cleanUsername });
    }

    // --- Admin: remove a dealer's login ---
    if (action === 'remove-credentials') {
      const { dealerName } = req.body;
      const accounts = (await redis.get(ACCOUNTS_KEY)) || {};
      for (const [u, acc] of Object.entries(accounts)) {
        if (acc.dealerName === dealerName) delete accounts[u];
      }
      await redis.set(ACCOUNTS_KEY, accounts);
      return res.status(200).json({ ok: true });
    }

    // --- Admin: list which dealers currently have a login (no password data returned) ---
    if (action === 'list') {
      const accounts = (await redis.get(ACCOUNTS_KEY)) || {};
      const list = Object.entries(accounts).map(([username, acc]) => ({ username, dealerName: acc.dealerName }));
      return res.status(200).json({ accounts: list });
    }

    return res.status(400).json({ error: 'Unknown action.' });
  } catch (err) {
    console.error('dealer-auth error', err);
    return res.status(500).json({ error: String(err) });
  }
}
