# BoviBerry Studio — Project Context

## What this is
Website + business management platform for BoviBerry, a digital art/design
studio. Andy (this user) handles all technical work; his girlfriend is the
studio owner/client and manages content via the admin panel.

## Stack
- Frontend: Pure HTML/CSS/JS — index.html, pricing.html, reviews.html,
  moodboard.html, faq.html, privacy.html, boards.html, thank-you.html,
  admin.html, robots.txt
- Backend: Google Apps Script — **apps-script/Code.js is the live, real
  backend, linked via clasp.** Do NOT reference or edit any other file
  named apps-script-backend.js — that was a stale, outdated reference copy
  and has been removed.
- Database: Google Sheets (Spreadsheet ID: 1tpezSsLBbe3Ytx7Es4KO2iiTQnyQzpAyP97WP3rk6zc)
- Web App URL: https://script.google.com/macros/s/AKfycbzekrL5Rt0MpbpSxUtoicKFDGRK3_oVpNqN60kARVH3grNpZa0RJho2Gbpj6_CpZ5ip/exec
- Payments: Stripe
- Spam protection: Cloudflare Turnstile (site key: 0x4AAAAAAD6ZQz0wbCYFxNZR) + honeypot
- Hosting: Netlify at boviberry.com (domain via Cloudflare Registrar)

## Deploying backend changes
This folder is linked to the live Apps Script project via clasp.
- After editing apps-script/Code.js, run `clasp push` from inside the
  apps-script/ folder to deploy it live.
- If a change is ever made directly in the online Apps Script editor,
  run `clasp pull` from that folder to bring it back down locally.

## Catalog-driven architecture
Prices, sale status, and visibility for everything sold on the site live
in the Catalog sheet, not hardcoded in HTML. `active=false` items now
show a dimmed "Out of Stock" badge on index.html/pricing.html (recently
fixed — they used to just disappear, which was unreliable) and are
excluded from the commission form dropdown.

## Brand
Green #2C7449, lime #9BD16B, amber #FFAA1B, cocoa #412A13, clay #815121
Fonts: Cormorant Garamond + Jost

## Currently pending
- Logo swap: replacing styled "BoviBerry" text with an image logo —
  waiting on a corrected file from the client (last upload came through
  as a solid black square, no visible artwork).
- Six new Stripe products needed for the restructured subscription tiers.
- Catalog sheet cleanup: old blueberry/strawberry rows need manual
  deletion before running syncCatalogDefaults and syncCatalogMeta.