// Vercel serverless function: /api/dealer-auth
// Handles dealer login (public), plus admin-only actions that let the staff app
// set/reset a dealer's username + password. The admin-only actions require a valid
// admin session token (from /api/admin-auth) — a dealer can never call those.
//
// Storage shape (Redis key "gr-bumpers:dealer-accounts"):
//   { [username]: { dealerName, salt, hash } }
// Storage shape (Redis key "gr-bumpers:dealer-sessions"):
//   { [token]: { username, dealerName, createdAt } }

import crypto from 'crypto';
import { redis, getAdminSession, checkRateLimit, recordFailedAttempt } from '../lib/serverAuth.js';

const ACCOUNTS_KEY = 'gr-bumpers:dealer-accounts';
const SESSIONS_KEY = 'gr-bumpers:dealer-sessions';
const ATTEMPTS_PREFIX = 'gr-bumpers:dealer-login-attempts:';

const SESSION_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000; // 30 days
const MAX_ATTEMPTS = 5;
const LOCKOUT_MS = 15 * 60 * 1000; // 15 minutes

const ADMIN_ONLY_ACTIONS = new Set(['set-credentials', 'remove-credentials', 'rename-dealer', 'list']);

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

  if (ADMIN_ONLY_ACTIONS.has(action)) {
    const session = await getAdminSession((req.body || {}).token);
    const canManage = session && (session.role === 'owner' || session.permissions?.manageDealerLogins);
    if (!canManage) return res.status(401).json({ error: 'You don\'t have permission to manage dealer logins. Ask the owner to grant it in the Staff section.' });
  }

  try {
    // --- Dealer login (public, rate-limited) ---
    if (action === 'login') {
      const { username, password } = req.body;
      if (!username || !password) return res.status(400).json({ error: 'Username and password required.' });

      const cleanUsername = username.trim().toLowerCase();
      const rlKey = ATTEMPTS_PREFIX + cleanUsername;
      const rl = await checkRateLimit(rlKey, MAX_ATTEMPTS, LOCKOUT_MS);
      if (rl.locked) {
        return res.status(429).json({ error: `Too many failed attempts. Try again in about ${rl.waitMin} minute${rl.waitMin === 1 ? '' : 's'}.` });
      }

      const accounts = (await redis.get(ACCOUNTS_KEY)) || {};
      const account = accounts[cleanUsername];
      const hash = account ? hashPassword(password, account.salt) : null;
      if (!account || hash !== account.hash) {
        await recordFailedAttempt(rlKey, rl.attempts);
        return res.status(401).json({ error: 'Incorrect username or password.' });
      }
      await redis.del(rlKey);

      const sessions = (await redis.get(SESSIONS_KEY)) || {};
      const sessionToken = makeToken();
      sessions[sessionToken] = { username: cleanUsername, dealerName: account.dealerName, createdAt: Date.now() };
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
      if (Date.now() - session.createdAt > SESSION_MAX_AGE_MS) {
        delete sessions[sessionToken];
        await redis.set(SESSIONS_KEY, sessions);
        return res.status(401).json({ error: 'Session expired.' });
      }
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

    // --- Admin: set or reset a dealer's login ---
    if (action === 'set-credentials') {
      const { dealerName, username, password } = req.body;
      if (!dealerName || !username || !password) {
        return res.status(400).json({ error: 'dealerName, username, and password are required.' });
      }
      const cleanUsername = username.trim().toLowerCase();
      const accounts = (await redis.get(ACCOUNTS_KEY)) || {};
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

      const sessions = (await redis.get(SESSIONS_KEY)) || {};
      let changed = false;
      for (const [tok, session] of Object.entries(sessions)) {
        if (session.dealerName === dealerName) { delete sessions[tok]; changed = true; }
      }
      if (changed) await redis.set(SESSIONS_KEY, sessions);

      return res.status(200).json({ ok: true });
    }

    // --- Admin: keep a dealer's portal login working after a rename in the main app ---
    if (action === 'rename-dealer') {
      const { oldName, newName } = req.body;
      if (!oldName || !newName) return res.status(400).json({ error: 'oldName and newName are required.' });

      const accounts = (await redis.get(ACCOUNTS_KEY)) || {};
      let accChanged = false;
      for (const acc of Object.values(accounts)) {
        if (acc.dealerName === oldName) { acc.dealerName = newName; accChanged = true; }
      }
      if (accChanged) await redis.set(ACCOUNTS_KEY, accounts);

      const sessions = (await redis.get(SESSIONS_KEY)) || {};
      let sessChanged = false;
      for (const session of Object.values(sessions)) {
        if (session.dealerName === oldName) { session.dealerName = newName; sessChanged = true; }
      }
      if (sessChanged) await redis.set(SESSIONS_KEY, sessions);

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
