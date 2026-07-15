// Vercel serverless function: /api/admin-auth
// Two kinds of logins:
//   - Owner: username "admin" (or blank), password = the ADMIN_PASSWORD env var.
//     Always has every permission, including managing staff accounts.
//   - Staff: a named account created by the owner, with its own username/password
//     and a specific set of permissions (right now just "manage dealer logins").
// Includes basic rate limiting so passwords can't be brute-forced by hammering this
// endpoint.

import crypto from 'crypto';
import { redis, ADMIN_SESSIONS_KEY, STAFF_ACCOUNTS_KEY, getAdminSession, checkRateLimit, recordFailedAttempt } from '../lib/serverAuth.js';

const MAX_ATTEMPTS = 5;
const LOCKOUT_MS = 15 * 60 * 1000; // 15 minutes

function hashPassword(password, salt) {
  return crypto.scryptSync(password, salt, 64).toString('hex');
}
function makeToken() {
  return crypto.randomBytes(24).toString('hex');
}

function normalizePermissions(permissions) {
  const p = permissions || {};
  return {
    manageDealerLogins: !!p.manageDealerLogins,
    processNewOrders: !!p.processNewOrders,
    createOrders: !!p.createOrders,
    editOrders: !!p.editOrders,
    editModelsAndDealers: !!p.editModelsAndDealers,
    receiveOrders: !!p.receiveOrders,
  };
}

async function requireOwner(req) {
  const session = await getAdminSession((req.body || {}).token);
  if (!session || session.role !== 'owner') return null;
  return session;
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
    // --- Login (owner or staff) ---
    if (action === 'login') {
      const { password } = req.body;
      const username = (req.body.username || '').trim().toLowerCase();
      const isOwnerLogin = !username || username === 'admin';

      const rlKey = `gr-bumpers:admin-login-attempts:${isOwnerLogin ? 'owner' : username}`;
      const rl = await checkRateLimit(rlKey, MAX_ATTEMPTS, LOCKOUT_MS);
      if (rl.locked) {
        return res.status(429).json({ error: `Too many failed attempts. Try again in about ${rl.waitMin} minute${rl.waitMin === 1 ? '' : 's'}.` });
      }

      if (isOwnerLogin) {
        const adminPassword = process.env.ADMIN_PASSWORD;
        if (!adminPassword) {
          return res.status(500).json({ error: 'ADMIN_PASSWORD is not set on the server yet. Add it in Vercel → Settings → Environments, then redeploy.' });
        }
        if (!password || password !== adminPassword) {
          await recordFailedAttempt(rlKey, rl.attempts);
          return res.status(401).json({ error: 'Incorrect password.' });
        }
        await redis.del(rlKey);
        const sessions = (await redis.get(ADMIN_SESSIONS_KEY)) || {};
        const sessionToken = makeToken();
        sessions[sessionToken] = {
          role: 'owner', username: 'admin',
          permissions: { manageDealerLogins: true, processNewOrders: true, createOrders: true, editOrders: true, editModelsAndDealers: true, receiveOrders: true, manageStaff: true },
          createdAt: Date.now(),
        };
        await redis.set(ADMIN_SESSIONS_KEY, sessions);
        return res.status(200).json({ token: sessionToken, role: 'owner', username: 'admin', permissions: sessions[sessionToken].permissions });
      }

      // Staff login
      const staff = (await redis.get(STAFF_ACCOUNTS_KEY)) || {};
      const account = staff[username];
      const hash = account ? hashPassword(password, account.salt) : null;
      if (!account || hash !== account.hash) {
        await recordFailedAttempt(rlKey, rl.attempts);
        return res.status(401).json({ error: 'Incorrect username or password.' });
      }
      await redis.del(rlKey);
      const sessions = (await redis.get(ADMIN_SESSIONS_KEY)) || {};
      const sessionToken = makeToken();
      const permissions = account.permissions || {};
      sessions[sessionToken] = { role: 'staff', username, permissions, createdAt: Date.now() };
      await redis.set(ADMIN_SESSIONS_KEY, sessions);
      return res.status(200).json({ token: sessionToken, role: 'staff', username, permissions });
    }

    if (action === 'session') {
      const { token: sessionToken } = req.body;
      if (!sessionToken) return res.status(400).json({ error: 'Token required.' });
      const session = await getAdminSession(sessionToken);
      if (!session) return res.status(401).json({ error: 'Session expired or invalid.' });
      return res.status(200).json({ role: session.role, username: session.username, permissions: session.permissions });
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

    // --- Owner-only: staff account management ---
    if (action === 'list-staff') {
      if (!(await requireOwner(req))) return res.status(401).json({ error: 'Owner login required.' });
      const staff = (await redis.get(STAFF_ACCOUNTS_KEY)) || {};
      const list = Object.entries(staff).map(([username, acc]) => ({
        username, permissions: acc.permissions || {}, createdAt: acc.createdAt,
      }));
      return res.status(200).json({ staff: list });
    }

    if (action === 'create-staff') {
      if (!(await requireOwner(req))) return res.status(401).json({ error: 'Owner login required.' });
      const { username, password, permissions } = req.body;
      const cleanUsername = (username || '').trim().toLowerCase();
      if (!cleanUsername || !password) return res.status(400).json({ error: 'Username and password are required.' });
      if (cleanUsername === 'admin') return res.status(400).json({ error: '"admin" is reserved for the owner login.' });
      const staff = (await redis.get(STAFF_ACCOUNTS_KEY)) || {};
      const salt = crypto.randomBytes(16).toString('hex');
      staff[cleanUsername] = {
        salt, hash: hashPassword(password, salt),
        permissions: normalizePermissions(permissions),
        createdAt: Date.now(),
      };
      await redis.set(STAFF_ACCOUNTS_KEY, staff);
      return res.status(200).json({ ok: true, username: cleanUsername });
    }

    if (action === 'update-staff-permissions') {
      if (!(await requireOwner(req))) return res.status(401).json({ error: 'Owner login required.' });
      const { username, permissions } = req.body;
      const cleanUsername = (username || '').trim().toLowerCase();
      const staff = (await redis.get(STAFF_ACCOUNTS_KEY)) || {};
      if (!staff[cleanUsername]) return res.status(404).json({ error: 'No such staff account.' });
      staff[cleanUsername].permissions = normalizePermissions(permissions);
      await redis.set(STAFF_ACCOUNTS_KEY, staff);

      // Update any of their active sessions immediately, not just future logins.
      const sessions = (await redis.get(ADMIN_SESSIONS_KEY)) || {};
      let changed = false;
      for (const session of Object.values(sessions)) {
        if (session.role === 'staff' && session.username === cleanUsername) {
          session.permissions = staff[cleanUsername].permissions;
          changed = true;
        }
      }
      if (changed) await redis.set(ADMIN_SESSIONS_KEY, sessions);

      return res.status(200).json({ ok: true });
    }

    if (action === 'delete-staff') {
      if (!(await requireOwner(req))) return res.status(401).json({ error: 'Owner login required.' });
      const { username } = req.body;
      const cleanUsername = (username || '').trim().toLowerCase();
      const staff = (await redis.get(STAFF_ACCOUNTS_KEY)) || {};
      delete staff[cleanUsername];
      await redis.set(STAFF_ACCOUNTS_KEY, staff);

      // Kick them out of any active session right away.
      const sessions = (await redis.get(ADMIN_SESSIONS_KEY)) || {};
      let changed = false;
      for (const [tok, session] of Object.entries(sessions)) {
        if (session.role === 'staff' && session.username === cleanUsername) { delete sessions[tok]; changed = true; }
      }
      if (changed) await redis.set(ADMIN_SESSIONS_KEY, sessions);

      return res.status(200).json({ ok: true });
    }

    return res.status(400).json({ error: 'Unknown action.' });
  } catch (err) {
    console.error('admin-auth error', err);
    return res.status(500).json({ error: String(err) });
  }
}
