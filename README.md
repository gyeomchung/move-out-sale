# Move-Out Sale

Live sale page for our Downtown Brooklyn move-out sale.

- `index.html` — the whole site. Items load live from the Google Sheet through the Apps Script web app (`CONFIG.apiUrl`), with an embedded snapshot as fallback.
- Q&A comments and pickup reservations are stored in the sheet's `Comments` and `Reservations` tabs.
- Owner mode: open the site with `#owner` and enter the owner key.
