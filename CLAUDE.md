# Move-Out Sale — how this project works

Owner: Gyeom (notification email lives in the sheet's Settings!B1 — never put it in this public repo). Sale ends Dec 3, 2026. Downtown Brooklyn, 1 min from Hoyt–Schermerhorn.

## Pieces
- `index.html` — the whole public site, served by GitHub Pages at https://gyeomchung.github.io/move-out-sale/.
  Pushing to `main` publishes it in ~1 min. Item data is embedded as a fallback snapshot (`<script id="data">`),
  but live data comes from the Apps Script API (`CONFIG.apiUrl`).
- `backend/Code.gs` + `backend/appsscript.json` — Google Apps Script web app bound to nothing (opens the sheet by ID).
  **It self-deploys**: every 5 min it pulls these two files from `main` on raw.githubusercontent.com, replaces its own
  code via the Apps Script API and moves the web-app deployment to the new version. To deploy instantly after a push,
  POST `{"action":"sync","sha":"<40-char commit sha>"}` to the web app URL (e.g. from the browser pane on the
  site's origin) — the SHA bypasses raw.githubusercontent.com caching of `main`. The `Check backend` GitHub Action
  does this automatically after every push that touches `backend/`.
  Never break `selfUpdate_` / `doGet` — a broken updater can't fix itself. Always `node --check` before pushing.
- Google Sheet `1OctIxn2ZInOi_aG52P9nQw3arQ-1KNTbg22BYzFSdCk`
  - `Move Out Sale` tab: items. Status dropdown: Posted / Reserved / Sold / On Hold. Column M = Photo (Drive link).
    Public price = "Initial Price" (older name "Price"). "Final Price" = private negotiated sale price, never shown.
    Columns are found by header name — if Gyeom renames a header, update the lookups in Code.gs and index.html.
  - `Comments`, `Reservations`, `Waitlist` tabs are written by the backend.

## Rules
- Reserved is computed: qty − active reservations − live waitlist holds (24h). Backend syncs Posted ⇄ Reserved.
- Owner cancels by setting Reservations Status = Cancelled (processed by `tick` within 10 min).
- Pickup slots: Nov 9 – Dec 3; weekdays 6:30–9:30pm starts, weekends 9am–9:30pm starts, 30-min steps, one booking per slot.

## Photos
- Sheet column M holds a Drive link per item. The `Sync photos` GitHub Action (on push, every 15 min, or manual)
  downloads each linked photo to `images/<driveId>.jpg`; the site loads that same-origin copy and only falls back
  to Drive if the copy doesn't exist yet. (Drive hotlinks break on some phones, e.g. Samsung Internet.)
- Commit as `Gyeom Chung <gyeomchung@users.noreply.github.com>` — never put a personal email in this public repo.

## Calendar feed
- `<web app URL>?ics=<key>` serves active reservations as an iCalendar feed ("Move Out Sale"), which Gyeom subscribes
  to in Apple Calendar. The key is in the sheet's Settings!B3 — never commit it.
- Reservation, cancellation and owner-side waitlist emails (joined / offered) also go to the addresses in Settings!B5 (`reservationEmails_`).
- Blocked pickup days: sheet Settings!B4 (comma-separated YYYY-MM-DD, store as text). Backend rejects them and
  returns them as `blocked`; the site hides those days.

## Owner changes without touching the sheet
- Gyeom does NOT open the sheet — Claude makes every data change. Claude has no sheet/owner-key access from its
  sandbox, so changes ship as one-shot ops in `backend/Code.gs` (`OWNER_OPS`, plus older `OWNER_SEEDS`): each runs once
  on the next `doGet` after deploy (Script Property `OP_<id>`), and its result (IDs/counts only, no names) is returned
  as `ops` in the public GET. Add a new op with the next id; never reuse or edit an id that already ran.
- Verify with the `Peek items` GitHub Action (`gh workflow run "Peek items"`), which prints `ops` and matching items.
