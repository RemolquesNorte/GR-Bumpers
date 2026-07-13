// Thin client for the shared key-value store used by the app.
// Talks to /api/storage, which is backed by Upstash Redis (see api/storage.js).
// The `shared` argument that App.jsx passes everywhere is accepted but ignored —
// every key in this app is meant to be shared across the whole team, so there's
// only one namespace on the server.

const API = '/api/storage';

export async function storageGet(key) {
  try {
    const res = await fetch(`${API}?key=${encodeURIComponent(key)}`);
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
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ key, value }),
    });
    if (!res.ok) console.error('storage set failed', key, await res.text());
  } catch (e) {
    console.error('storage set failed', key, e);
  }
}
