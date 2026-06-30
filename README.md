# WeRoCon Lab — Procurement & Inventory System

Automatic procurement, document generation, and inventory tracking for the
WeRoCon Lab. Static site (GitHub Pages) backed by Supabase for real,
permanent, shared data.

## What this does

- Enter a component once → fund approval, quotation comparison, Non-GeM
  certificate, and payment receipt all auto-generate from that data as
  real `.docx` files (editable in Word), matching the IIT Jodhpur
  Acct/R&D-01 format.
- One-click "Download complete procurement pack" → a `.zip` with all
  applicable documents + your uploaded quotation/invoice files.
- Inventory table — permanent delete (no restore bugs), shared live
  across every device via Supabase.
- Shipment tracking — manual entry always works; auto-sync via
  TrackCourier.io once you deploy the Edge Function (steps below).
- Shared password gate for you and Dr. Saurav Kumar (see Security note).

## 1. Supabase setup (you've likely already done this)

1. Create a project at supabase.com.
2. SQL Editor → run `supabase_schema.sql` (creates tables + RLS policies).
3. Storage → create a **public** bucket named `procurement-files`.
4. Project Settings → API → copy your **Project URL** and **publishable
   (anon) key** into `js/config.js`.

## 2. TrackCourier.io auto-tracking setup

This app uses [TrackCourier.io](https://api.trackcourier.io/), a
multi-courier tracking API with strong coverage of Indian couriers
(BlueDart, Delhivery, DTDC, India Post, Ecom Express, etc.) — one
API key tracks any of them by courier slug + tracking number.

Even though TrackCourier.io's own demo page allows direct browser
calls, this app still routes through a small Supabase Edge Function so
your **paid API key never ends up in the public GitHub repo** (anyone
viewing your repo's source could otherwise copy it and burn your quota).

Install the Supabase CLI once:
```bash
npm install -g supabase
supabase login
```

From the `werocon/` folder:
```bash
supabase link --project-ref lzmcblzqoakxfmgexyar
supabase secrets set TRACKCOURIER_API_KEY=tc_live_xxxxxxxxxxxxxxxxxxxx
supabase functions deploy track-shipment
```

That's it — `js/config.js` already points `TRACKING_FUNCTION_URL` at the
deployed function. Once deployed, the "Sync now" button on the Shipment
Tracking tab pulls live status from TrackCourier.io and writes it into
Supabase (visible to both you and Dr. Saurav instantly).

**Note on status mapping:** TrackCourier.io's exact response field names
weren't available to verify while building this (their docs page is a
JS-rendered Swagger UI), so the Edge Function maps statuses with a
best-effort keyword match (looks for "DELIVERED", "IN TRANSIT", etc.
anywhere in the raw response). Every raw API response is saved to
`shipment_tracking.raw_status_payload` in Supabase — after your first
real sync, check that column; if statuses aren't mapping correctly,
share one example raw payload and the mapping in
`supabase/functions/track-shipment/index.ts` can be tightened.

**Free tier note:** TrackCourier.io's free plan is 100 requests/month —
fine for lab volume, but "Sync now" is a manual button on purpose;
don't wire it to auto-fire on every page load.

**Courier name → slug mapping:** the Edge Function maps the courier
names in this app's dropdown (`js/tracking.js` COURIERS list) to
TrackCourier.io's slugs (`bluedart`, `delhivery`, `dtdc`, `indiapost`,
`ecomexpress`, `fedex`, etc.). If you add a courier to the dropdown that
isn't in `COURIER_SLUG_MAP` in the Edge Function, add it there too —
check `GET /v1/couriers` on TrackCourier.io for the exact slug.

## 3. Set your access codes

Open `js/auth.js` and change:
```js
const ACCESS_CODES = {
  'werocon2026': 'Lab member',
  'professor2026': 'Professor (view)'
};
```
to whatever codes you want. Give Dr. Saurav Kumar the
`'Professor (view)'` code if you want him to only view (not edit/delete);
give him the same code as you if he should have full access — just add
another key mapped to `'Lab member'`.

**Security note (read this):** this is a shared password, not real
per-person login. Anyone with the code gets the same access level; it
deters casual access, not a determined person. The Supabase database
itself isn't gated by this — its protection is the publishable key +
RLS policies. If you later want real separate logins (so the system
knows *who* deleted what), swap this for Supabase Auth — ask Claude,
it's a moderate-sized follow-up, not a rewrite.

## 4. Deploy to GitHub Pages

```bash
cd werocon
git init
git add .
git commit -m "Initial WeRoCon procurement system"
git remote add origin https://github.com/<your-username>/werocon-lab.git
git push -u origin main
```

Then on GitHub: **Settings → Pages → Source: main branch, / (root)**.
Your site will be live at `https://<your-username>.github.io/werocon-lab/`.

**Important:** don't commit your TrackCourier.io API key anywhere — it
only ever lives in `supabase secrets set`, never in this repo. The
Supabase URL and publishable key in `config.js` are safe to commit (see
comment in that file for why).

## File structure

```
werocon/
├── index.html
├── css/style.css
├── js/
│   ├── config.js        — your Supabase URL/key, edit this
│   ├── auth.js           — access codes, edit this
│   ├── db.js              — Supabase data layer
│   ├── docgen.js          — generates the 4 .docx documents + zip bundle
│   ├── tracking.js        — tracking constants + live-sync stub
│   ├── app.js             — UI logic
│   ├── supabase.js        — vendored supabase-js (don't edit)
│   ├── docx.iife.js       — vendored docx.js (don't edit)
│   └── jszip.min.js       — vendored JSZip (don't edit)
├── supabase_schema.sql   — run once in Supabase SQL Editor
├── supabase/functions/track-shipment/index.ts — TrackCourier.io proxy
└── assets/logo.png
```

## Known limitations (so there are no surprises later)

- Access gate is a shared password, not real accounts (see Security note above).
- TrackCourier.io status mapping is best-effort until verified against a
  live response (see note above) — check raw_status_payload after first sync.
- TrackCourier.io free tier is 100 requests/month; "Sync now" is manual.
- The public Storage bucket means anyone with a direct file link can view
  it — fine for internal lab documents, not for anything truly sensitive.
- Document generation runs entirely in the browser; very large component
  lists (100+) may take a few seconds to package into the `.zip`.
