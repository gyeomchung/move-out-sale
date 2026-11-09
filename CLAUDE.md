# Move-Out Sale — how this project works

Owner: Gyeom (gyeomcc@gmail.com). Sale ends Dec 3, 2026. Downtown Brooklyn, 1 min from Hoyt–Schermerhorn.

## Pieces
- `index.html` — the whole public site, served by GitHub Pages at https://gyeomchung.github.io/move-out-sale/.
  Pushing to `main` publishes it in ~1 min. Item data is embedded as a fallback snapshot (`<script id="data">`),
  but live data comes from the Apps Script API (`CONFIG.apiUrl`).
- `backend/Code.gs` + `backend/appsscript.json` — Google Apps Script web app bound to nothing (opens the sheet by ID).
  **It self-deploys**: every 5 min it pulls these two files from `main` on raw.githubusercontent.com, replaces its own
  code via the Apps Script API and moves the web-app deployment to the new version. To deploy instantly after a push,
  POST `{"action":"sync","sha":"<40-char commit sha>"}` to the web app URL (e.g. from the browser pane on the
  site's origin) — the SHA bypasses raw.githubusercontent.com caching of `main`.
  Never break `selfUpdate_` / `doGet` — a broken updater can't fix itself. Always `node --check` before pushing.
- Google Sheet `1OctIxn2ZInOi_aG52P9nQw3arQ-1KNTbg22BYzFSdCk`
  - `Move Out Sale` tab: items. Status dropdown: Posted / Reserved / Sold / On Hold. Column M = Photo (Drive link).
  - `Comments`, `Reservations`, `Waitlist` tabs are written by the backend.

## Rules
- Reserved is computed: qty − active reservations − live waitlist holds (24h). Backend syncs Posted ⇄ Reserved.
- Owner cancels by setting Reservations Status = Cancelled (processed by `tick` within 10 min).
- Pickup slots: Nov 9 – Dec 3; weekdays 6:30–9:30pm starts, weekends 9am–9:30pm starts, 30-min steps, one booking per slot.
