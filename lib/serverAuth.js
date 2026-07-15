// Shared server-side helpers used by multiple /api routes.
// Lives outside the api/ folder so Vercel doesn't treat it as a route itself.

import { Redis } from '@upstash/redis';

const url = process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL;
const token = process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN;

export const redis = url && token ? new Redis({ url, token }) : null;

export const NS = 'gr-bumpers:';
export const ADMIN_SESSIONS_KEY = 'gr-bumpers:admin-sessions';
export const STAFF_ACCOUNTS_KEY = 'gr-bumpers:staff-accounts';
export const ADMIN_SESSION_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000; // 30 days

// Returns the session record { role: 'owner'|'staff', username, permissions, createdAt }
// or null if the token is missing/invalid/expired. Sessions created before the staff
// system existed have no `role` field — those were only ever issued via the master
// password, so we treat a missing role as 'owner' for backward compatibility.
export async function getAdminSession(sessionToken) {
  if (!redis || !sessionToken) return null;
  const sessions = (await redis.get(ADMIN_SESSIONS_KEY)) || {};
  const session = sessions[sessionToken];
  if (!session) return null;
  if (Date.now() - session.createdAt > ADMIN_SESSION_MAX_AGE_MS) return null;
  if (!session.role) return { ...session, role: 'owner', permissions: { manageDealerLogins: true, manageStaff: true } };
  return session;
}

export async function isValidAdminToken(sessionToken) {
  return !!(await getAdminSession(sessionToken));
}

// Re-checks the CURRENT session holder's own password (owner's ADMIN_PASSWORD, or a
// staff member's own account password) — used as a "confirm it's really you" step
// before sensitive changes like resetting a dealer's login.
export async function verifyAdminPassword(sessionToken, password) {
  const session = await getAdminSession(sessionToken);
  if (!session || !password) return false;
  if (session.role === 'owner') {
    return password === process.env.ADMIN_PASSWORD;
  }
  const crypto = await import('crypto');
  const staff = (await redis.get(STAFF_ACCOUNTS_KEY)) || {};
  const account = staff[session.username];
  if (!account) return false;
  const hash = crypto.scryptSync(password, account.salt, 64).toString('hex');
  return hash === account.hash;
}

// Basic Redis-backed rate limiter for login endpoints. Returns { locked, waitMin, attempts }.
export async function checkRateLimit(key, maxAttempts, lockoutMs) {
  const attempts = (await redis.get(key)) || { count: 0, first: 0 };
  const now = Date.now();
  if (attempts.count >= maxAttempts && now - attempts.first < lockoutMs) {
    const waitMin = Math.ceil((lockoutMs - (now - attempts.first)) / 60000);
    return { locked: true, waitMin, attempts };
  }
  if (now - attempts.first > lockoutMs) {
    return { locked: false, attempts: { count: 0, first: now } };
  }
  return { locked: false, attempts };
}

export async function recordFailedAttempt(key, attempts) {
  const now = Date.now();
  const next = attempts.count === 0 ? { count: 1, first: now } : { count: attempts.count + 1, first: attempts.first };
  await redis.set(key, next);
}
