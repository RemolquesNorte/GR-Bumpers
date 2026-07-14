// Vercel serverless function: /api/backup-manage
// Admin-only. Lets the staff app list the automatic daily backups (see
// api/backup-cron.js) and restore the live data from one of them if something
// ever goes wrong.

import { redis, NS, isValidAdminToken } from '../lib/serverAuth.js';

const BACKUP_INDEX_KEY = 'gr-bumpers:backup-index';

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, must-revalidate');
  if (!redis) {
    return res.status(500).json({ error: 'Redis credentials not configured.' });
  }
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ error: 'method not allowed' });
  }

  const { action, token: adminToken } = req.body || {};
  const authorized = await isValidAdminToken(adminToken);
  if (!authorized) return res.status(401).json({ error: 'Not authorized. Please log in.' });

  try {
    if (action === 'list') {
      const index = (await redis.get(BACKUP_INDEX_KEY)) || [];
      const details = [];
      for (const dateStr of index) {
        const backup = await redis.get(`gr-bumpers:backups:${dateStr}`);
        if (backup) details.push({ date: dateStr, takenAt: backup.takenAt });
      }
      return res.status(200).json({ backups: details });
    }

    if (action === 'restore') {
      const { date } = req.body;
      if (!date) return res.status(400).json({ error: 'date is required' });
      const backup = await redis.get(`gr-bumpers:backups:${date}`);
      if (!backup) return res.status(404).json({ error: 'That backup no longer exists.' });

      for (const [key, value] of Object.entries(backup.data)) {
        if (value !== null && value !== undefined) {
          await redis.set(NS + key, value);
        }
      }
      return res.status(200).json({ ok: true, restoredFrom: date });
    }

    return res.status(400).json({ error: 'Unknown action.' });
  } catch (err) {
    console.error('backup-manage error', err);
    return res.status(500).json({ error: String(err) });
  }
}
