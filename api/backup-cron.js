// Vercel serverless function: /api/backup-cron
// Triggered automatically once a day by Vercel Cron (see vercel.json). Snapshots
// every important data key into a dated backup, and prunes anything older than
// KEEP_DAYS so this doesn't grow forever. This is the "even if something breaks,
// we can roll back" safety net — separate from the manual "Download backup" button
// staff can use any time from the app itself.

import { redis, NS } from '../lib/serverAuth.js';

const DATA_KEYS = [
  'bumper-inventory', 'bumper-toolbox', 'bumper-dealers', 'bumper-orders',
  'bumper-sales', 'bumper-production-batches', 'bumper-production-log', 'bumper-new-orders',
];
const BACKUP_INDEX_KEY = 'gr-bumpers:backup-index';
const KEEP_DAYS = 14;

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, must-revalidate');

  // Verify this is really Vercel's scheduler, not a public request to this URL.
  const authHeader = req.headers.authorization;
  if (!process.env.CRON_SECRET || authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return res.status(401).json({ error: 'Unauthorized.' });
  }

  if (!redis) {
    return res.status(500).json({ error: 'Redis credentials not configured.' });
  }

  try {
    const snapshot = {};
    for (const key of DATA_KEYS) {
      snapshot[key] = await redis.get(NS + key);
    }

    const dateStr = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
    const backupKey = `gr-bumpers:backups:${dateStr}`;
    await redis.set(backupKey, { takenAt: Date.now(), data: snapshot });

    const index = (await redis.get(BACKUP_INDEX_KEY)) || [];
    const nextIndex = [dateStr, ...index.filter(d => d !== dateStr)];
    const toKeep = nextIndex.slice(0, KEEP_DAYS);
    const toDelete = nextIndex.slice(KEEP_DAYS);
    await redis.set(BACKUP_INDEX_KEY, toKeep);
    for (const oldDate of toDelete) {
      await redis.del(`gr-bumpers:backups:${oldDate}`);
    }

    return res.status(200).json({ ok: true, date: dateStr, keysBackedUp: DATA_KEYS.length });
  } catch (err) {
    console.error('backup-cron error', err);
    return res.status(500).json({ error: String(err) });
  }
}
