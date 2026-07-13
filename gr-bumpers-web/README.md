# GR Bumpers

Internal inventory / dealer / order management tool for GR Bumpers. React + Vite frontend,
with a tiny serverless API (`/api/storage`) backed by Upstash Redis so everyone on the team
sees the same live data.

## What's here

```
gr-bumpers/
├── api/
│   └── storage.js       # serverless function: get/set key-value pairs in Redis
├── src/
│   ├── App.jsx           # the whole app (all views, all logic)
│   ├── storage.js        # tiny client that talks to /api/storage
│   └── main.jsx          # React entry point
├── index.html
├── package.json
└── vite.config.js
```

There's no separate database schema to manage — the app stores a handful of JSON blobs
under fixed keys (inventory, dealers, orders, sales log, production log, etc.), the same
way it did before. The first time anyone opens the site, it seeds itself with your original
Mexico/GR inventory, dealer list, and orders — same as it did before.

## 1. Get the code onto GitHub

From this project folder:

```bash
git init
git add .
git commit -m "Initial commit"
```

Create a new empty repo on GitHub (github.com → New repository — don't initialize it with
a README), then:

```bash
git remote add origin https://github.com/YOUR-USERNAME/gr-bumpers.git
git branch -M main
git push -u origin main
```

## 2. Create the database (Upstash Redis)

The easiest path is entirely inside Vercel:

1. Go to your Vercel dashboard → **Storage** tab → **Create Database**.
2. Choose **Upstash** → **Redis** (it's on Vercel's Marketplace, free tier is plenty for this).
3. Once created, click **Connect Project** and pick this project (do step 3 first if you
   haven't imported the project yet — you can also come back and connect it afterward).

This automatically adds two environment variables to your Vercel project:
`UPSTASH_REDIS_REST_URL` and `UPSTASH_REDIS_REST_TOKEN`. That's all `api/storage.js` needs.

*(Alternative: create a free database directly at [upstash.com](https://upstash.com), then
manually copy those same two values into your Vercel project's Settings → Environment
Variables.)*

## 3. Deploy on Vercel

1. Go to [vercel.com/new](https://vercel.com/new) and import the GitHub repo you just pushed.
2. Vercel will auto-detect it as a Vite project — no build settings to change.
3. Make sure the Upstash env vars from step 2 are attached to this project (Settings →
   Environment Variables should show them).
4. Click **Deploy**.

Once it's live, open the URL Vercel gives you. The app will load empty the very first time,
then immediately seed itself with the original inventory/dealer/order data and save it to
Redis — after that, everyone who opens the link sees the same shared data.

## Troubleshooting: "no env vars showed up after connecting"

This happens because Vercel's Upstash integration doesn't always name the credentials
`UPSTASH_REDIS_REST_URL`/`TOKEN` — depending on how the database was created it can also
use the older `KV_REST_API_URL`/`KV_REST_API_TOKEN` naming. `api/storage.js` already checks
for both, so you don't need to rename anything — but you do need to confirm they're actually
attached to the project:

1. In Vercel, open your project → **Settings → Environment Variables**. You should see
   either pair listed there. If you see neither:
   - Go to the **Storage** tab → click into your Upstash database → **Projects** tab →
     make sure this project is listed as connected. If it isn't, connect it there.
2. **Redeploy** after connecting — env vars only take effect on the next deployment, not
   retroactively. Go to the **Deployments** tab → click the **⋯** menu on the latest
   deployment → **Redeploy**.
3. Reload the site. If it still doesn't work, check the error: opening the browser dev
   tools (F12) → Network tab → look at the `/api/storage` request. If it returns the
   "Redis credentials not found" message, the env vars still aren't attached — repeat
   step 1.

## Dealer Portal (new)

There's now a second, completely separate page at **`/portal`** for dealers. It has no
sidebar, no inventory, no other dealers' data — just a login, a simple order form, and a
list of *that dealer's own* orders.

**How it works:**
- Staff create a login for a dealer from the main site: **Dealers** → **Edit dealers** →
  click the key icon next to a dealer → set a username and password.
- The dealer goes to `yoursite.vercel.app/portal`, logs in, and can place an order (pick a
  model, quantity, optional PO) or see their order history. New orders show up immediately
  in your main Orders view, Bumper Lookup, Dealer Lookup, and Production Planning — they're
  the same order records, just entered by the dealer instead of you.
- The dealer never sees stock quantities, other dealers, models management, or anything
  else — the portal only ever talks to three narrow API routes (`/api/dealer-auth`,
  `/api/dealer-orders`, `/api/dealer-models`) that are scoped server-side to that one
  dealer's own data.

**Security notes, since this is a first pass and not a hardened system:**
- Passwords are hashed (not stored in plain text), but there's no password-reset flow,
  no expiring sessions, and no lockout after repeated failed attempts.
- Anyone who knows a dealer's username/password can log in from anywhere — there's no
  extra verification step (email, 2FA, etc.).
- If a dealer's login needs to be revoked, remove it from the Dealers section (same key
  icon → Remove login).

If you ever want this tightened up — expiring sessions, forced password resets, rate
limiting on login attempts — that's a reasonable next step, just ask.

## Local development

```bash
npm install
npm run dev
```

This runs the UI on `localhost:5173`, but plain `vite dev` doesn't run the `/api` serverless
function, so imports/saves that touch the database won't work locally that way. To test the
full thing locally (including the API), use the Vercel CLI instead:

```bash
npm install -g vercel
vercel dev
```

`vercel dev` will ask you to link the project to Vercel the first time — say yes, and it'll
pull down the same environment variables so the database calls work locally too.

## Making changes later

Any time you want Claude (or you, by hand) to change something, edit the files and push:

```bash
git add .
git commit -m "describe the change"
git push
```

Vercel redeploys automatically on every push to `main`.

## Notes

- All data is shared across everyone who opens the site — there's no login/user system.
  If you ever want per-user accounts or permissions, that's a bigger change (auth + likely
  a real database instead of a flat key-value store) — just ask and we can add it.
- The XLSX importers (Mexico inventory, GR/US inventory, Orders) run entirely in the
  browser — no file ever gets uploaded to a server other than the parsed JSON that gets
  saved to Redis.
