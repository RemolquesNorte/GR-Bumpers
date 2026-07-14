// Vercel serverless function: /api/admin-auth
// Single shared password for staff access to the main site (not per-user accounts —
// just a wall so a stray link isn't the same as full access). Includes basic rate
// limiting so the password can't be brute-forced by hammering this endpoint.

import crypto from 'crypto';
import { redis, ADMIN_SESSIONS_KEY, checkRateLimit, recordFailedAttempt } from '../lib/serverAuth.js';

const ATTEMPTS_KEY = 'gr-bumpers:admin-login-attempts';
const MAX_ATTEMPTS = 5;
const LOCKOUT_MS = 15 * 60 * 1000; // 15 minutes

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
    if (action === 'login') {
      const { password } = req.body;
      const adminPassword = process.env.ADMIN_PASSWORD;
      if (!adminPassword) {
        return res.status(500).json({
          error: 'ADMIN_PASSWORD is not set on the server yet. Add it in Vercel → Settings → Environments, then redeploy.',
        });
      }

      const rl = await checkRateLimit(ATTEMPTS_KEY, MAX_ATTEMPTS, LOCKOUT_MS);
      if (rl.locked) {
        return res.status(429).json({ error: `Too many failed attempts. Try again in about ${rl.waitMin} minute${rl.waitMin === 1 ? '' : 's'}.` });
      }

      if (!password || password !== adminPassword) {
        await recordFailedAttempt(ATTEMPTS_KEY, rl.attempts);
        return res.status(401).json({ error: 'Incorrect password.' });
      }

      await redis.del(ATTEMPTS_KEY);
      const sessions = (await redis.get(ADMIN_SESSIONS_KEY)) || {};
      const sessionToken = makeToken();
      sessions[sessionToken] = { createdAt: Date.now() };
      await redis.set(ADMIN_SESSIONS_KEY, sessions);
      return res.status(200).json({ token: sessionToken });
    }

    if (action === 'session') {
      const { token: sessionToken } = req.body;
      if (!sessionToken) return res.status(400).json({ error: 'Token required.' });
      const sessions = (await redis.get(ADMIN_SESSIONS_KEY)) || {};
      const session = sessions[sessionToken];
      if (!session) return res.status(401).json({ error: 'Session expired or invalid.' });
      if (Date.now() - session.createdAt > 30 * 24 * 60 * 60 * 1000) {
        delete sessions[sessionToken];
        await redis.set(ADMIN_SESSIONS_KEY, sessions);
        return res.status(401).json({ error: 'Session expired.' });
      }
      return res.status(200).json({ ok: true });
    }

    if (action === 'logout') {
      const { token: sessionToken } = req.body;
      if (sessionToken) {
        const sessions = (await redis.get(ADMIN_SESSIONS_KEY)) || {};
        delete sessions[sessionToken];
        await redis.set(ADMIN_SESSIONS_KEY, sessions);
      }
      return res.status(200).json({ ok: true });
    }

    return res.status(400).json({ error: 'Unknown action.' });
  } catch (err) {
    console.error('admin-auth error', err);
    return res.status(500).json({ error: String(err) });
  }
}
