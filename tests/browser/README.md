# Browser tests (Playwright, no Firebase project needed)

`fakefb.js` replaces the Firebase compat SDK with an in-memory fake, so every screen runs against seeded data
without touching the real project. `harness.js` serves `index.html` from a local HTTP server and routes
`vendor/firebase/*` to the fake.

```
cd <repo root>
python3 -m http.server 8765 &
npm i -g playwright            # or use an existing Playwright + Chromium
node tests/browser/xss.js       # stored-XSS sweep: payload in every text field of every collection, every view/tab/drawer/print window
ROLE=w node tests/browser/xss.js
node tests/browser/auth.js      # e-mail verification gate, Google sign-in (login + verify screen + re-auth), re-authentication before role changes, self-demotion guard, per-department records, idle sign-out
node tests/browser/pwa.js       # notifications bell, calendar, version footer + what’s new, error boundary + clientErrors report, admin health panel, service worker install (no reload on first install)
node tests/browser/govern.js  # access review with re-auth, account register CSV, access matrix, retention periods, legal hold + record history, SHA-256 approval stamp and tamper detection
node tests/browser/polish.js    # external seed file + empty-database banner, skip link + live toast region, sequential meeting decision numbers
node tests/browser/offline.js   # offline writes resolve on the device, pending counter + status chip, sync toast, late rejection reported, cache reads, offline sign-in with the saved role, logout guard
node tests/browser/sms055.js    # SMS055 dashboard: import, parts, KPIs, previous-period comparison, replace overlap, CSV, permissions, English
node tests/browser/debts.js     # debts dashboard: template sheets, every KPI to the dinar, six aging buckets, government vs private, actions, customer statement, replace, monthly statement, large month, EN + dark
node tests/browser/units.js     # quantity conversions per ASTM D1250-04 Table 54B / 53B / 56, LPG, measured density, spec ranges (no browser)
node tests/browser/closures.js  # sales / transfers / receipts dashboards: import of the three closing reports, units (m³ conversion, densities), comparison, replace overlap, wrong file, permissions, English + dark
node tests/browser/margin.js    # profit margin: Sales_Report import, m³, price board with effective dates + history, two price lists (local / foreign × rate), profit = hand calculation, price-list import, simulator, calculator, permissions, XSS, English + dark
node tests/browser/meetnotify.js # meeting notices in the bell for invitees (new/changed/cancelled), «notify everyone» dialog: WhatsApp, e-mail Bcc, .ics, sent marks
node tests/browser/powerbi.js   # Power BI button beside every Excel export/import, export dialog + .zip pack (CSV, Power Query), CSV import, CSV-injection guard
node tests/browser/formats.js   # import in every format: xlsx/xlsm/xlsb/xls, WPS et/ett, ods/fods, Excel 2003 XML, HTML, CSV (UTF-8, ;, Windows-1256), TSV, UTF-16 text, JSON; accept list on every import input; SMS055 end to end from .et/.ods/.xlsb; a broken PDF gets a clear message
node tests/browser/pdf.js       # import from PDF: Calc and Chrome prints of the sample match the Excel file (Arabic, ligatures, numbers, codes, repeated headers, footers, second table); sales/transfers/receipts/SMS055/margin dashboards from Excel-style PDF prints; scanned, password-protected and damaged PDFs get clear messages; pdf.js loaded on demand
node tests/browser/palette.js   # command palette (Ctrl+K), record search with Arabic-Indic digits, theme action, Alt+1, g-k chord, help dialog, no shortcuts while typing
```
`xss.js` prints `hits: []` when nothing rendered as HTML; `auth.js` prints a JSON object where every value must be `true` (or the expected list) and ends with `AUTH OK`; `pwa.js` ends with `PWA/NOTIFY OK` and `palette.js` with `PALETTE OK`, `govern.js` with `GOVERN OK`, `polish.js` with `POLISH OK`, `offline.js` with `OFFLINE OK`, `sms055.js` with `SMS055 OK`, `units.js` with `UNITS OK`, `closures.js` with `CLOSURES OK`, `margin.js` with `MARGIN OK`, `debts.js` with `DEBTS OK`, `formats.js` with `FORMATS OK`, `pdf.js` with `PDF OK`.
