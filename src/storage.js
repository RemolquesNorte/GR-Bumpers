// Thin client for the shared key-value store used by the app.
// Talks to /api/storage, which is backed by Upstash Redis (see api/storage.js).
// Every request now carries the admin session token — the server rejects anything
// without one. setAdminToken() is called once by App.jsx right after login.

const API = '/api/storage';

let adminToken = null;
let onUnauthorized = null;

export function setAdminToken(token) {
  adminToken = token;
}
export function getAdminToken() {
  return adminToken;
}
export function onUnauthorizedRequest(handler) {
  onUnauthorized = handler;
}

export async function storageGet(key) {
  try {
    const res = await fetch(`${API}?key=${encodeURIComponent(key)}&token=${encodeURIComponent(adminToken || '')}&_=${Date.now()}`, { cache: 'no-store' });
    if (res.status === 401) { onUnauthorized?.(); return null; }
    if (res.status === 404) return null;
    if (!res.ok) return null;
    const data = await res.json();
    return data.value ?? null;
  } catch (e) {
    console.error('storage get failed', key, e);
    return null;
  }
}

export async function storageSet(key, value) {
  try {
    const res = await fetch(API, {
      method: 'POST',
      cache: 'no-store',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ key, value, token: adminToken }),
    });
    if (res.status === 401) { onUnauthorized?.(); return; }
    if (!res.ok) console.error('storage set failed', key, await res.text());
  } catch (e) {
    console.error('storage set failed', key, e);
  }
}
