# Deploy PonPower to agenticpropertyos.com

The app is a Vite SPA + a Vercel serverless function (`api/share.ts`) backed
by a Neon Postgres database. Three dashboards, one push.

## 1. Neon — new database

1. Open https://console.neon.tech (same account that hosts PropCortex).
2. New project → name it `ponpower` (or add a new database to an existing project).
3. Open the SQL Editor and paste the contents of [`db/schema.sql`](db/schema.sql), run it.
4. From the project's **Connection Details**, copy the pooled connection string
   (looks like `postgres://user:pwd@ep-xxx-pooler.neon.tech/dbname?sslmode=require`).
   You'll paste this into Vercel as `DATABASE_URL`.

## 2. Vercel — free `agenticpropertyos.com` from the old project

1. Open the **PropCortex** project in Vercel.
2. Settings → Domains → remove `agenticpropertyos.com` and `www.agenticpropertyos.com`.
   (`propcortex.net` stays. Nothing else changes on PropCortex.)

## 3. Vercel — new project for PonPower

1. New Project → import this repo (`ksvendetta/PonPower`).
2. Framework preset: **Vite** (auto-detected). Build command and output dir come from
   [`vercel.json`](vercel.json) — no manual override needed.
3. Environment Variables → add **DATABASE_URL** = (the Neon connection string from step 1).
4. Deploy.
5. Settings → Domains → add **`agenticpropertyos.com`** and **`www.agenticpropertyos.com`**.
   Vercel will pick up the existing DNS automatically since you just removed it from
   PropCortex.

That's it. Share links generated from the app will be:

    https://www.agenticpropertyos.com/map.html?id=<8 chars>

The viewer fetches the map data from `/api/share?id=...`, decompresses, and renders.

## Local dev with the API route

`npm run dev` only runs the Vite client + Express dev server (no `/api/*`).
To exercise the share endpoint locally:

    npx vercel dev

This runs Vite + `/api/*` functions on one port. You'll need `DATABASE_URL` set
(create `.env.local` with `DATABASE_URL=...` — Vercel CLI reads it).
