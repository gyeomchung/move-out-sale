/**
 * Move-Out Sale backend (Google Apps Script web app)
 *  - Serves items (with live availability) from the "Move Out Sale" tab
 *  - Anonymous Q&A per item, owner replies (owner key)
 *  - Pickup reservations → items become "Reserved"; owner gets an email
 *  - Waitlist per item: when a reservation is cancelled, the next person is emailed
 *    a claim link that holds the item for 24 hours, then it moves to the next person
 *
 * Auto-deploy: this file lives in github.com/gyeomchung/move-out-sale/backend.
 *  The script pulls backend/Code.gs + backend/appsscript.json from GitHub (every 5 min, or instantly via
 *  POST {action:'sync'}), replaces its own code through the Apps Script API, and points the web app
 *  deployment at the new version. Edit the code on GitHub, not in the editor.
 *
 * One-time setup: turn on the Apps Script API (script.google.com/home/usersettings), paste this file and
 *  appsscript.json, run setup(), then Deploy → Manage deployments → ✏️ → New version → Deploy.
 *
 * Owner actions in the sheet:
 *  - Cancel a reservation: set its Status (Reservations tab) to "Cancelled".
 *    Within ~10 minutes the items are released and the waitlist is notified.
 *  - Sold: set the item's Status to "Sold" in the Move Out Sale tab.
 */

const SITE_URL = 'https://gyeomchung.github.io/move-out-sale/';
const SHEET_ID = '1OctIxn2ZInOi_aG52P9nQw3arQ-1KNTbg22BYzFSdCk';
const ITEMS_TAB = 'Move Out Sale', COMMENTS_TAB = 'Comments', RES_TAB = 'Reservations', WAIT_TAB = 'Waitlist';
const TZ = 'America/New_York';
const HOLD_HOURS = 24;
const BACKEND_VERSION = '2026-11-13.5';
const REPO_RAW = 'https://raw.githubusercontent.com/gyeomchung/move-out-sale/main/backend/';
const DEPLOYMENT_ID = 'AKfycbz3KhKAzisUaMzxDfwU4BhsvnoH68_72IV1nV2qgzuWifbI9KT_vqxpRxsMn7b3tYLt';

const RES_HEAD = ['Time', 'ID', 'Items', 'Total', 'Pickup Date', 'Pickup Time', 'Name', 'Contact', 'Note', 'Status', 'Token', 'Released'];
const WAIT_HEAD = ['Time', 'ID', 'Row', 'Item', 'Email', 'Status', 'Offered At', 'Token'];

// Owner's notification email lives in the sheet (Settings!B1), not in this public code
// Owner + extra people (Settings!B5) for reservation emails: new pickups and cancellations
function reservationEmails_() {
  let extra = '';
  try { extra = String(ss_().getSheetByName('Settings').getRange('B5').getValue()); } catch (e) {}
  return [ownerEmail_()].concat(extra.split(/[,;\s]+/).filter(isEmail_)).join(',');
}

function ownerEmail_() {
  try { return String(ss_().getSheetByName('Settings').getRange('B1').getValue()).trim(); } catch (e) { return Session.getEffectiveUser().getEmail(); }
}
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function sh_(name) { return ss_().getSheetByName(name); }

/* ======================= setup ======================= */

function setup() {
  const ss = ss_();
  const ensure = (name, head) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  };
  ensure(COMMENTS_TAB, ['Time', 'ID', 'Item', 'Name', 'Text', 'Reply To', 'Owner', 'Hidden', 'Token', 'Edited']);
  ensure(RES_TAB, RES_HEAD);
  ensure(WAIT_TAB, WAIT_HEAD);
  const res = ss.getSheetByName(RES_TAB);
  res.getRange('J2:J1000').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['New', 'Confirmed', 'Picked Up', 'Cancelled'], true).build());

  ScriptApp.getProjectTriggers().filter(t => ['tick', 'autoUpdate'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('autoUpdate').timeBased().everyMinutes(5).create();

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('OWNER_KEY')) props.setProperty('OWNER_KEY', Utilities.getUuid().slice(0, 8));
  Logger.log('Owner key: ' + props.getProperty('OWNER_KEY'));
  Logger.log('Auto-deploy check: ' + JSON.stringify(selfUpdate_(true)));
}

/* ======================= web endpoints ======================= */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.claim) return json_(claimInfo_(p.claim));
  if (p.res) return json_(resInfo_(p.res));
  if (p.ics) return icsFeed_(p.ics);
  runOwnerSeeds_(); runOwnerOps_();
  return json_({ ops: opResults_(), v: BACKEND_VERSION, items: items_(), comments: comments_(), booked: booked_(), blocked: blockedDays_() });
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad request' }); }
  if (b.website) return json_({ ok: true });                       // honeypot
  if (b.action === 'sync') return json_(syncNow_(b.sha));
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (b.action) {
      case 'comment': return json_(comment_(b));
      case 'reply': return json_(reply_(b));
      case 'hide': return json_(hide_(b));
      case 'editComment': return json_(editComment_(b));
      case 'deleteComment': return json_(deleteComment_(b));
      case 'reserve': return json_(reserve_(b));
      case 'cancel': return json_(cancel_(b));
      case 'waitlist': return json_(waitlist_(b));
      case 'checkKey': return json_({ ok: isOwner_(b.key) });
      case 'testNotify': return json_(testNotify_(b));
      case 'ownerReserve': return json_(ownerReserve_(b));
      default: return json_({ ok: false, error: 'unknown action' });
    }
  } finally { lock.releaseLock(); }
}

/* ======================= data model ======================= */

function itemRows_() {
  const v = sh_(ITEMS_TAB).getDataRange().getDisplayValues();
  const h = v[0].map(s => String(s).trim());
  const c = (...names) => { for (const n of names) { const i = h.indexOf(n); if (i >= 0) return i; } return -1; };   // header names, first match wins
  const money = s => { const n = parseFloat(String(s).replace(/[$,\s]/g, '')); return isFinite(n) ? n : null; };
  return { statusCol: c('Status') + 1, list: v.slice(1).map((r, i) => ({
    row: i + 2,
    category: r[c('Category')], sub: r[c('Subcategory')], name: String(r[c('Item')]).trim(),
    // 'Final Price' is the private negotiated sale price — never shown on the site. Only an explicit 'Price Drop' column lowers the public price.
    price: money(r[c('Initial Price', 'Price')]), qty: parseInt(r[c('QTY')], 10) || 1, drop: c('Price Drop') >= 0 ? money(r[c('Price Drop')]) : null,
    condition: r[c('Condition')], original: money(r[c('Original Price + Tax')]),
    links: String(r[c('Original Item Link')] || '').split(/\s+/).map(s => s.replace(/^-/, '')).filter(s => /^https?:/.test(s)),
    status: String(r[c('Status')] || '').trim(),
    photo: c('Photo') >= 0 ? String(r[c('Photo')] || '').trim() : ''
  })).filter(x => x.name && x.price != null) };
}

// Reservation items cell: one line per item → "row|count|name"
function parseResItems_(cell) {
  return String(cell || '').split('\n').map(l => {
    const m = /^(\d+)\|(\d+)\|(.*)$/.exec(l.trim());
    return m ? { row: +m[1], n: +m[2], name: m[3] } : null;
  }).filter(Boolean);
}

function resRows_() {
  const sh = sh_(RES_TAB); if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, RES_HEAD.length).getDisplayValues().map((r, i) => ({
    sheetRow: i + 2, id: r[1], items: parseResItems_(r[2]), date: r[4], time: r[5], name: r[6], contact: r[7],
    status: r[9], token: r[10], released: r[11] === 'TRUE'
  })).filter(r => r.id);
}

function waitRows_() {
  const sh = sh_(WAIT_TAB); if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, WAIT_HEAD.length).getValues().map((r, i) => ({
    sheetRow: i + 2, id: r[1], row: +r[2], name: r[3], email: r[4], status: r[5], offeredAt: r[6], token: r[7]
  })).filter(r => r.id);
}

const activeRes_ = r => !/cancel/i.test(r.status);
const liveOffer_ = w => w.status === 'Offered' && w.offeredAt instanceof Date && (Date.now() - w.offeredAt.getTime()) < HOLD_HOURS * 3600e3;

// Units still free for a row = qty − units in active reservations − units held by live waitlist offers
function availability_(items, res, wait) {
  const used = {};
  res.filter(activeRes_).forEach(r => r.items.forEach(it => used[it.row] = (used[it.row] || 0) + it.n));
  wait.filter(liveOffer_).forEach(w => used[w.row] = (used[w.row] || 0) + 1);
  const out = {};
  items.forEach(it => out[it.row] = Math.max(0, it.qty - (used[it.row] || 0)));
  return out;
}

function items_() {
  const { list } = itemRows_(), res = resRows_(), wait = waitRows_();
  const avail = availability_(list, res, wait);
  const waiting = {};
  wait.filter(w => w.status === 'Waiting').forEach(w => waiting[w.row] = (waiting[w.row] || 0) + 1);
  return list.map(it => Object.assign(it, { avail: avail[it.row], waiting: waiting[it.row] || 0 }));
}

// Keep the sheet's Status in step: Posted ⇄ Reserved (never touches On Hold / Sold)
function syncStatuses_() {
  const { list, statusCol } = itemRows_(), res = resRows_(), wait = waitRows_();
  const avail = availability_(list, res, wait), sh = sh_(ITEMS_TAB);
  list.forEach(it => {
    if (!/^(posted|reserved)$/i.test(it.status)) return;
    const want = avail[it.row] > 0 ? 'Posted' : 'Reserved';
    if (it.status !== want) sh.getRange(it.row, statusCol).setValue(want);
  });
}

/* ======================= Q&A ======================= */

function comments_() {
  const sh = sh_(COMMENTS_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  // Columns: Time, ID, Item, Name, Text, Reply To, Owner, Hidden, Token (private, never returned), Edited
  return sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues()
    .filter(r => r[1] && r[7] !== true)
    .map(r => ({ time: r[0] instanceof Date ? r[0].toISOString() : r[0], id: r[1], item: r[2], name: r[3], text: r[4], parent: r[5], owner: r[6] === true,
      edited: r[9] instanceof Date ? r[9].toISOString() : '' }));
}

function comment_(b) {
  const text = clean_(b.text, 600), item = clean_(b.item, 200), label = clean_(b.label, 200) || item;
  if (!text || !item) return { ok: false, error: 'empty' };
  const id = Utilities.getUuid().slice(0, 8), name = 'Guest ' + id.slice(0, 4).toUpperCase();
  const token = Utilities.getUuid().replace(/-/g, '');   // lets this (anonymous) author edit/delete later
  sh_(COMMENTS_TAB).appendRow([new Date(), id, item, name, text, '', false, false, token, '']);
  MailApp.sendEmail(ownerEmail_(), '[Move-Out Sale] New question: ' + label,
    name + ' asked about "' + label + '":\n\n' + text + '\n\nReply in owner mode: ' + SITE_URL + '#owner');
  return { ok: true, id: id, token: token };
}

function reply_(b) {
  if (!isOwner_(b.key)) return { ok: false, error: 'not owner' };
  const text = clean_(b.text, 1000);
  if (!text) return { ok: false, error: 'empty' };
  const id = Utilities.getUuid().slice(0, 8);
  sh_(COMMENTS_TAB).appendRow([new Date(), id, clean_(b.item, 200), 'Owner', text, clean_(b.parent, 20), true, false, '', '']);
  return { ok: true, id: id };
}

// Edit/delete your own comment: anonymous authors prove it with the token they got when posting;
// the owner (owner key) can edit their own replies and delete anything.
function commentRow_(id) {
  const sh = sh_(COMMENTS_TAB);
  if (!id || sh.getLastRow() < 2) return null;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  for (let i = 0; i < rows.length; i++) if (rows[i][1] === id) return { sh: sh, row: i + 2, owner: rows[i][6] === true, token: String(rows[i][8] || '') };
  return null;
}
function mayChange_(c, b, forDelete) {
  if (!c) return false;
  if (isOwner_(b.key) && (c.owner || forDelete)) return true;
  return !c.owner && !!c.token && b.token === c.token;
}
function editComment_(b) {
  const c = commentRow_(clean_(b.id, 20)), text = clean_(b.text, c && c.owner ? 1000 : 600);
  if (!mayChange_(c, b, false)) return { ok: false, error: 'not allowed' };
  if (!text) return { ok: false, error: 'empty' };
  c.sh.getRange(c.row, 5).setValue(text);
  c.sh.getRange(c.row, 10).setValue(new Date());
  return { ok: true };
}
function deleteComment_(b) {
  const c = commentRow_(clean_(b.id, 20));
  if (!mayChange_(c, b, true)) return { ok: false, error: 'not allowed' };
  c.sh.getRange(c.row, 8).setValue(true);   // soft delete: hidden from the site, kept in the sheet
  return { ok: true };
}

function hide_(b) {
  if (!isOwner_(b.key)) return { ok: false, error: 'not owner' };
  const sh = sh_(COMMENTS_TAB);
  if (sh.getLastRow() < 2) return { ok: true };
  const ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  ids.forEach((r, i) => { if (r[0] === b.id) sh.getRange(i + 2, 8).setValue(true); });
  return { ok: true };
}

/* ======================= reservations ======================= */

function booked_() {
  return resRows_().filter(activeRes_).map(r => r.date + ' ' + r.time);
}

function reserve_(b) {
  const name = clean_(b.name, 80), contact = clean_(b.contact, 120), note = clean_(b.note, 500);
  const date = clean_(b.date, 10), time = clean_(b.time, 10);
  const want = (b.items || []).slice(0, 50).map(x => typeof x === 'string' ? { row: 0, n: 1, name: clean_(x, 200) }
    : { row: +x.row, n: Math.max(1, +x.n || 1), name: clean_(x.name, 200) });
  if (!name || !contact || !date || !time || !want.length) return { ok: false, error: 'missing fields' };
  if (!validSlot_(date, time)) return { ok: false, error: 'invalid slot' };
  if (booked_().indexOf(date + ' ' + time) >= 0) return { ok: false, error: 'taken' };

  const { list } = itemRows_(), res = resRows_(), wait = waitRows_();
  const avail = availability_(list, res, wait);
  const claim = b.claim ? wait.find(w => w.token === b.claim && liveOffer_(w)) : null;
  if (claim) avail[claim.row] += 1;                                 // the held unit belongs to this person

  const lines = [], gone = [];
  let total = 0;
  want.forEach(x => {
    const it = list.find(i => i.row === x.row && i.name === x.name) || list.find(i => i.name === x.name);
    if (!it || /sold|on hold/i.test(it.status) || avail[it.row] < x.n) { gone.push(x.name); return; }
    avail[it.row] -= x.n;
    const unit = (it.drop != null && it.drop > 0 && it.drop < it.price) ? it.drop : it.price;
    total += unit * x.n;
    lines.push({ row: it.row, n: x.n, name: it.name, unit: unit });
  });
  if (gone.length) return { ok: false, error: 'unavailable', items: gone };

  const id = Utilities.getUuid().slice(0, 6).toUpperCase(), token = Utilities.getUuid().replace(/-/g, '');
  sh_(RES_TAB).appendRow([new Date(), id, lines.map(l => l.row + '|' + l.n + '|' + l.name).join('\n'), total,
    "'" + date, "'" + time, name, contact, note, 'New', token, false]);
  if (claim) sh_(WAIT_TAB).getRange(claim.sheetRow, 6).setValue('Claimed');
  syncStatuses_();

  const nice = niceDate_(date), list_ = lines.map(l => '• ' + l.name + (l.n > 1 ? ' × ' + l.n : '') + ' — $' + (l.unit * l.n)).join('\n');
  const cancelUrl = SITE_URL + '?cancel=' + token;
  MailApp.sendEmail(reservationEmails_(), '[Move-Out Sale] Pickup reserved — ' + nice + ' ' + time12_(time) + ' · ' + name,
    'Reservation #' + id + (claim ? ' (from waitlist)' : '') + '\n\nName: ' + name + '\nContact: ' + contact +
    '\nPickup: ' + nice + ' at ' + time12_(time) + '\n\nItems:\n' + list_ + '\nTotal: $' + total + (note ? '\n\nNote: ' + note : '') +
    '\n\nTo cancel: set Status to "Cancelled" in the Reservations tab, or open ' + cancelUrl);
  if (isEmail_(contact)) {
    MailApp.sendEmail(contact, 'Your move-out sale pickup — ' + nice + ' at ' + time12_(time),
      'Hi ' + name + ',\n\nThanks! Your pickup is reserved (#' + id + ').\n\nPickup: ' + nice + ' at ' + time12_(time) +
      '\nDowntown Brooklyn, 1 min from Hoyt–Schermerhorn station (exact address will be sent when we confirm).\n\n' +
      list_ + '\nTotal: $' + total + '\n\nPlans changed? Cancel here so the next person can have it:\n' + cancelUrl);
  }
  return { ok: true, id: id, token: token, total: total };
}

function resInfo_(token) {
  const r = resRows_().find(x => x.token === token);
  if (!r) return { ok: false, error: 'not found' };
  return { ok: true, id: r.id, status: r.status, date: r.date, time: r.time, items: r.items.map(i => i.name + (i.n > 1 ? ' × ' + i.n : '')) };
}

function cancel_(b) {
  const r = resRows_().find(x => x.token === b.token);
  if (!r) return { ok: false, error: 'not found' };
  if (/cancel/i.test(r.status)) return { ok: true, already: true };
  if (/picked/i.test(r.status)) return { ok: false, error: 'picked up' };
  const sh = sh_(RES_TAB);
  sh.getRange(r.sheetRow, 10).setValue('Cancelled');
  sh.getRange(r.sheetRow, 12).setValue(true);
  MailApp.sendEmail(reservationEmails_(), '[Move-Out Sale] Reservation cancelled — ' + r.name,
    'Reservation #' + r.id + ' (' + niceDate_(r.date) + ' ' + time12_(r.time) + ') was cancelled by the buyer.\n\nItems:\n' +
    r.items.map(i => '• ' + i.name).join('\n') + '\n\nWaitlisted buyers (if any) are being notified automatically.');
  release_(r.items.map(i => i.row));
  return { ok: true };
}

/* ======================= waitlist ======================= */

function waitlist_(b) {
  const email = clean_(b.email, 120), row = +b.row, name = clean_(b.name, 200);
  if (!isEmail_(email)) return { ok: false, error: 'email' };
  const { list } = itemRows_();
  const it = list.find(i => i.row === row && i.name === name) || list.find(i => i.name === name);
  if (!it || /sold/i.test(it.status)) return { ok: false, error: 'unavailable' };
  const wait = waitRows_();
  const mine = wait.find(w => w.row === it.row && String(w.email).toLowerCase() === email.toLowerCase() && /waiting|offered/i.test(w.status));
  const ahead = wait.filter(w => w.row === it.row && w.status === 'Waiting').length;
  if (mine) return { ok: true, position: ahead, already: true };
  sh_(WAIT_TAB).appendRow([new Date(), Utilities.getUuid().slice(0, 8), it.row, it.name, email, 'Waiting', '', '']);
  MailApp.sendEmail(email, "You're on the waitlist — " + it.name,
    "Thanks! You're #" + (ahead + 1) + ' on the waitlist for "' + it.name + '".\n\n' +
    "If it frees up, we'll email you a link to reserve it. You'll have " + HOLD_HOURS + ' hours to book before it goes to the next person.\n\n' + SITE_URL);
  MailApp.sendEmail(ownerEmail_(), '[Move-Out Sale] Waitlist: ' + it.name, email + ' joined the waitlist (#' + (ahead + 1) + ') for "' + it.name + '".');
  release_([it.row]);    // in case a unit is actually free right now
  return { ok: true, position: ahead + 1 };
}

// For each row: while a unit is free and someone is waiting, offer it to the next person
function release_(rows) {
  const seen = {};
  rows.forEach(row => {
    if (seen[row]) return; seen[row] = true;
    for (let guard = 0; guard < 20; guard++) {
      const { list } = itemRows_(), wait = waitRows_();
      const it = list.find(i => i.row === row);
      if (!it || /sold|on hold/i.test(it.status)) break;
      const free = availability_(list, resRows_(), wait)[row];
      const next = wait.find(w => w.row === row && w.status === 'Waiting');
      if (free <= 0 || !next) break;
      const token = Utilities.getUuid().replace(/-/g, '');
      const sh = sh_(WAIT_TAB);
      sh.getRange(next.sheetRow, 6, 1, 3).setValues([['Offered', new Date(), token]]);
      MailApp.sendEmail(next.email, 'It\'s available! Reserve "' + it.name + '" within ' + HOLD_HOURS + ' hours',
        'Good news — "' + it.name + '" from the move-out sale is available, and we\'re holding it for you.\n\n' +
        'Pick a pickup time here within ' + HOLD_HOURS + ' hours:\n' + SITE_URL + '?claim=' + token + '\n\n' +
        "If you don't reserve by then, it will be offered to the next person on the waitlist.");
      MailApp.sendEmail(ownerEmail_(), '[Move-Out Sale] Offered to waitlist: ' + it.name,
        '"' + it.name + '" was offered to ' + next.email + ' (holds for ' + HOLD_HOURS + 'h).');
    }
  });
  syncStatuses_();
}

function claimInfo_(token) {
  const w = waitRows_().find(x => x.token === token);
  if (!w) return { ok: false, error: 'not found' };
  if (w.status === 'Claimed') return { ok: false, error: 'claimed' };
  if (!liveOffer_(w)) return { ok: false, error: 'expired' };
  const { list } = itemRows_();
  const it = list.find(i => i.row === w.row);
  return { ok: true, row: w.row, name: it ? it.name : w.name, expires: new Date(w.offeredAt.getTime() + HOLD_HOURS * 3600e3).toISOString() };
}

/* ======================= timer (every 10 min) ======================= */

function tick() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const rows = [];
    // 1) Reservations the owner cancelled in the sheet
    const sh = sh_(RES_TAB);
    resRows_().forEach(r => {
      if (/cancel/i.test(r.status) && !r.released) { sh.getRange(r.sheetRow, 12).setValue(true); r.items.forEach(i => rows.push(i.row)); }
    });
    // 2) Waitlist offers that ran out
    const ws = sh_(WAIT_TAB);
    waitRows_().forEach(w => {
      if (w.status === 'Offered' && !liveOffer_(w)) {
        ws.getRange(w.sheetRow, 6).setValue('Expired');
        MailApp.sendEmail(w.email, 'Your hold on "' + w.name + '" has expired',
          'The ' + HOLD_HOURS + '-hour hold on "' + w.name + '" ended, so it has moved to the next person. Thanks for your interest!');
        rows.push(w.row);
      }
    });
    if (rows.length) release_(rows); else syncStatuses_();
  } finally { lock.releaseLock(); }
}

/* ======================= helpers ======================= */

// Days with no pickups, from the sheet: Settings!B4, comma-separated YYYY-MM-DD
function blockedDays_() {
  try {
    const v = ss_().getSheetByName('Settings').getRange('B4').getDisplayValue();
    return String(v).split(/[,\s]+/).filter(d => /^\d{4}-\d{2}-\d{2}$/.test(d));
  } catch (e) { return []; }
}

function validSlot_(date, time) {
  if (blockedDays_().indexOf(date) >= 0) return false;
  if (!/^2026-(11|12)-\d{2}$/.test(date) || date < '2026-11-09' || date > '2026-12-03') return false;
  const m = /^(\d{1,2}):(\d{2})$/.exec(time); if (!m) return false;
  const mins = (+m[1]) * 60 + (+m[2]); if (mins % 30) return false;
  const dow = new Date(date + 'T12:00:00').getDay(), weekend = dow === 0 || dow === 6;
  return weekend ? (mins >= 9 * 60 && mins <= 21 * 60 + 30) : (mins >= 18 * 60 + 30 && mins <= 21 * 60 + 30);
}
function niceDate_(d) { return Utilities.formatDate(new Date(d + 'T12:00:00'), TZ, 'EEE, MMM d'); }
function time12_(t) { const m = /^(\d+):(\d+)$/.exec(t); if (!m) return t; const h = +m[1]; return ((h + 11) % 12 + 1) + ':' + m[2] + (h < 12 ? 'am' : 'pm'); }
function isEmail_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s)); }
function isOwner_(key) { return !!key && key === PropertiesService.getScriptProperties().getProperty('OWNER_KEY'); }
function clean_(s, n) { return String(s == null ? '' : s).replace(/[<>]/g, '').trim().slice(0, n); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ======================= self-update from GitHub ======================= */

function autoUpdate() { selfUpdate_(false); }

function syncNow_(sha) {
  const cache = CacheService.getScriptCache();
  if (cache.get('sync')) return { ok: true, skipped: 'rate limited' };
  cache.put('sync', '1', 20);
  return selfUpdate_(false, sha);
}

function selfUpdate_(force, sha) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { ok: false, error: 'busy' };
  try {
    // raw.githubusercontent.com caches `main` for ~5 min; a commit SHA (sent by `sync` right after a push) is always fresh
    const base_ = /^[0-9a-f]{40}$/.test(String(sha || '')) ? REPO_RAW.replace('/main/', '/' + sha + '/') : REPO_RAW;
    const code = UrlFetchApp.fetch(base_ + 'Code.gs', { muteHttpExceptions: true });
    const manifest = UrlFetchApp.fetch(base_ + 'appsscript.json', { muteHttpExceptions: true });
    if (code.getResponseCode() !== 200 || manifest.getResponseCode() !== 200) return { ok: false, error: 'fetch ' + code.getResponseCode() };
    const src = code.getContentText(), man = manifest.getContentText();
    if (src.indexOf('function doGet') < 0 || src.indexOf('function selfUpdate_') < 0) return { ok: false, error: 'sanity check failed' };
    JSON.parse(man);
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, src + man));
    const props = PropertiesService.getScriptProperties();
    if (!force && props.getProperty('DEPLOYED_HASH') === hash) return { ok: true, changed: false };

    const id = ScriptApp.getScriptId(), base = 'https://script.googleapis.com/v1/projects/' + id;
    const call = (method, url, body) => {
      const r = UrlFetchApp.fetch(url, { method: method, contentType: 'application/json', payload: JSON.stringify(body),
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
      if (r.getResponseCode() >= 300) throw new Error(method + ' ' + url.replace(base, '') + ' → ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 300));
      return JSON.parse(r.getContentText() || '{}');
    };
    call('put', base + '/content', { files: [
      { name: 'Code', type: 'SERVER_JS', source: src },
      { name: 'appsscript', type: 'JSON', source: man }
    ] });
    const ver = call('post', base + '/versions', { description: 'auto ' + new Date().toISOString() });
    call('put', base + '/deployments/' + DEPLOYMENT_ID, { deploymentConfig: {
      scriptId: id, versionNumber: ver.versionNumber, manifestFileName: 'appsscript', description: 'auto v' + ver.versionNumber } });
    props.setProperty('DEPLOYED_HASH', hash);
    return { ok: true, changed: true, version: ver.versionNumber };
  } catch (err) {
    const props = PropertiesService.getScriptProperties(), msg = String(err && err.message || err);
    if (props.getProperty('LAST_UPDATE_ERROR') !== msg) {
      props.setProperty('LAST_UPDATE_ERROR', msg);
      MailApp.sendEmail(ownerEmail_(), '[Move-Out Sale] Auto-deploy failed', msg);
    }
    return { ok: false, error: msg };
  } finally { lock.releaseLock(); }
}

/* ======================= calendar feed (owner only) ======================= */
// Apple/Google Calendar subscribe to  <web app URL>?ics=<key>  — the key lives in the sheet (Settings!B3), not here.

function icsFeed_(key) {
  let want = '';
  try { want = String(ss_().getSheetByName('Settings').getRange('B3').getValue()).trim(); } catch (e) {}
  if (!want || key !== want) return ContentService.createTextOutput('not found');
  const esc = s => String(s == null ? '' : s).replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
  const utc = d => Utilities.formatDate(d, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const fold = line => { const out = []; while (line.length > 73) { out.push(line.slice(0, 73)); line = ' ' + line.slice(73); } out.push(line); return out.join('\r\n'); };
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//move-out-sale//pickups//EN', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
    'X-WR-CALNAME:Move Out Sale', 'X-WR-TIMEZONE:' + TZ, 'X-APPLE-CALENDAR-COLOR:#8E5CF7',
    'REFRESH-INTERVAL;VALUE=DURATION:PT15M', 'X-PUBLISHED-TTL:PT15M'];
  const now = utc(new Date());
  resRows_().filter(activeRes_).forEach(r => {
    if (!r.date) return;   // offline deal with no pickup date yet
    const allDay = !r.time;   // date set, time not yet → all-day event
    let start;
    try { start = Utilities.parseDate(r.date + ' ' + (r.time || '12:00'), TZ, 'yyyy-MM-dd HH:mm'); } catch (e) { return; }
    const end = new Date(start.getTime() + 30 * 60e3);
    const what = r.items.map(i => i.name + (i.n > 1 ? ' × ' + i.n : '')).join(', ');
    const picked = /picked/i.test(r.status) ? '✓ ' : '';
    const desc = 'Reservation #' + r.id + ' (' + r.status + ')\nBuyer: ' + r.name + '\nContact: ' + r.contact +
      '\nItems:\n' + r.items.map(i => '• ' + i.name + (i.n > 1 ? ' × ' + i.n : '')).join('\n') + (r.note ? '\nNote: ' + r.note : '');
    lines.push('BEGIN:VEVENT', 'UID:res-' + r.id + '@move-out-sale', 'DTSTAMP:' + now, allDay ? 'DTSTART;VALUE=DATE:' + r.date.replace(/-/g, '') : 'DTSTART:' + utc(start),
      allDay ? 'DTEND;VALUE=DATE:' + Utilities.formatDate(new Date(start.getTime() + 864e5), TZ, 'yyyyMMdd') : 'DTEND:' + utc(end),
      fold('SUMMARY:' + esc(picked + '픽업: ' + r.name + ' — ' + what)), fold('DESCRIPTION:' + esc(desc)),
      'LOCATION:Downtown Brooklyn (home)', 'STATUS:CONFIRMED',
      'BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:' + esc('Pickup: ' + r.name), 'TRIGGER:-PT60M', 'END:VALARM', 'END:VEVENT');
  });
  lines.push('END:VCALENDAR');
  return ContentService.createTextOutput(lines.join('\r\n') + '\r\n').setMimeType(ContentService.MimeType.ICAL);
}

// Owner-only: send a test reservation notification to everyone on the reservation list
function testNotify_(b) {
  if (!isOwner_(b.key)) return { ok: false, error: 'not owner' };
  const to = reservationEmails_();
  MailApp.sendEmail(to, '[Move-Out Sale] Test — reservation notifications are working',
    'This is a test from the move-out sale site.\n\nFrom now on, every new pickup reservation and every cancellation is emailed to:\n' +
    to.split(',').join('\n') + '\n\nNo action needed.');
  return { ok: true, to: to.split(',').length };
}

// Owner-only: record a deal made outside the site (any date/time or none, no slot rules). Optional b.id (e.g. 'OWNER3'). Items are matched by exact name.
// Appears in the calendar feed and marks the items Reserved, like the hand-made OWNER rows.
function ownerReserve_(b) {
  if (!isOwner_(b.key)) return { ok: false, error: 'not owner' };
  return recordOwnerReserve_(b);
}

// One-shot offline deals Gyeom asked Claude to record (no owner key in the repo, so they run once inside the script).
// Each runs at most once (Script Property SEED_<id>); an existing row with that ID counts as done.
const OWNER_SEEDS = [
  { id: 'OWNER3', match: /^outdoor\s+tent\s+bundle$/i, note: 'Arranged offline' }
];
function runOwnerSeeds_() {
  const props = PropertiesService.getScriptProperties();
  const todo = OWNER_SEEDS.filter(s => !props.getProperty('SEED_' + s.id));
  if (!todo.length) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    const ids = resRows_().map(r => r.id), { list } = itemRows_();
    todo.forEach(s => {
      if (props.getProperty('SEED_' + s.id)) return;
      let out;
      if (ids.indexOf(s.id) >= 0) out = { ok: true, skipped: 'id exists' };
      else {
        const hits = list.filter(i => s.match.test(i.name));
        out = hits.length === 1 ? recordOwnerReserve_({ id: s.id, items: [hits[0].name], note: s.note })
                                : { ok: false, error: 'matched ' + hits.length + ' items' };
      }
      props.setProperty('SEED_' + s.id, JSON.stringify(out));
    });
  } catch (e) { /* never break the public feed */ } finally { lock.releaseLock(); }
}

// One-shot sheet edits Gyeom asks Claude for in chat (Gyeom doesn't edit the sheet). Each runs once (Script Property
// OP_<id>); its result — IDs and counts only, never names or contacts — is returned publicly as `ops` by doGet.
const OWNER_OPS = [
  // 2026-11-13: today's buyer didn't show for Monstera Cathy Mint → move it to Bonnie's Sunday (11/15) pickup
  { id: 'OP1', run: () => moveItem_(/monstera.*cathy.*mint/i, '2026-11-13', /bonnie/i, 'Bonnie', '2026-11-15', 'No-show 11/13') }
];
function runOwnerOps_() {
  const props = PropertiesService.getScriptProperties();
  const todo = OWNER_OPS.filter(o => !props.getProperty('OP_' + o.id));
  if (!todo.length) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    todo.forEach(o => {
      if (props.getProperty('OP_' + o.id)) return;
      let out; try { out = o.run(); } catch (e) { out = { ok: false, error: String(e && e.message || e).slice(0, 200) }; }
      props.setProperty('OP_' + o.id, JSON.stringify(out));
    });
  } finally { lock.releaseLock(); }
}
function opResults_() {
  const props = PropertiesService.getScriptProperties(), out = {};
  OWNER_OPS.forEach(o => { const v = props.getProperty('OP_' + o.id); if (v) try { out[o.id] = JSON.parse(v); } catch (e) {} });
  return out;
}
function unitPrice_(it) { return (it.drop != null && it.drop > 0 && it.drop < it.price) ? it.drop : it.price; }
const lineStr_ = i => i.row + '|' + i.n + '|' + i.name;

// Move one item from a no-show reservation (on fromDate) to the buyer matching toName on toDate.
// Adds to the receiver first so the unit is never free in between (no waitlist offer). No buyer emails are sent.
function moveItem_(itemRe, fromDate, toRe, toLabel, toDate, note) {
  const { list } = itemRows_();
  const hits = list.filter(i => itemRe.test(i.name));
  if (hits.length !== 1) return { ok: false, error: 'item matched ' + hits.length };
  const it = hits[0], unit = unitPrice_(it);
  const live = resRows_().filter(r => activeRes_(r) && !/picked/i.test(r.status));
  const from = live.filter(r => r.date === fromDate && !toRe.test(r.name) && r.items.some(i => i.row === it.row));
  if (from.length !== 1) return { ok: false, error: 'no-show reservation matched ' + from.length };
  const toAll = live.filter(r => toRe.test(r.name));
  const to = toAll.filter(r => r.date === toDate);
  if (to.length > 1) return { ok: false, error: 'receiver matched ' + to.length };
  const sh = sh_(RES_TAB), out = { ok: true, item: it.name, from: from[0].id };
  if (to.length === 1) {
    const r = to[0];
    if (!r.items.some(i => i.row === it.row)) {
      sh.getRange(r.sheetRow, 3).setValue(r.items.concat([{ row: it.row, n: 1, name: it.name }]).map(lineStr_).join('\n'));
      const tot = parseFloat(String(sh.getRange(r.sheetRow, 4).getValue()).replace(/[$,]/g, '')) || 0;
      sh.getRange(r.sheetRow, 4).setValue(tot + unit);
    }
    out.to = r.id; out.toTime = r.time || ''; out.toItems = r.items.length + 1;
  } else {
    const made = recordOwnerReserve_({ items: [it.name], name: toLabel, date: toDate, note: 'Moved from no-show' });
    if (!made.ok) return made;
    out.to = made.id; out.created = true; out.toTime = ''; out.receiverOtherDates = toAll.length;
  }
  const f = from[0], left = f.items.filter(i => i.row !== it.row);
  const fNote = sh.getRange(f.sheetRow, 9).getValue();
  sh.getRange(f.sheetRow, 9).setValue((fNote ? fNote + ' · ' : '') + note + ': ' + it.name + ' moved');
  if (left.length) {
    sh.getRange(f.sheetRow, 3).setValue(left.map(lineStr_).join('\n'));
    const tot = parseFloat(String(sh.getRange(f.sheetRow, 4).getValue()).replace(/[$,]/g, '')) || 0;
    sh.getRange(f.sheetRow, 4).setValue(Math.max(0, tot - unit));
    out.fromItemsLeft = left.length;
  } else {
    sh.getRange(f.sheetRow, 10).setValue('Cancelled');
    sh.getRange(f.sheetRow, 12).setValue(true);
    out.fromCancelled = true;
  }
  syncStatuses_();
  return out;
}

function recordOwnerReserve_(b) {
  // Date/time are optional: a deal with no pickup set yet just holds the items (no calendar event until a date is added).
  const date = clean_(b.date, 10), time = clean_(b.time, 5);
  if (date && !/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: 'bad date' };
  if (time && (!date || !/^\d{2}:\d{2}$/.test(time))) return { ok: false, error: 'bad time' };
  const { list } = itemRows_();
  const lines = [], missing = [];
  let total = 0;
  (b.items || []).forEach(n => {
    const it = list.find(i => i.name === clean_(n, 200));
    if (!it) { missing.push(n); return; }
    lines.push(it.row + '|1|' + it.name); total += it.price;
  });
  if (missing.length || !lines.length) return { ok: false, error: 'item not found', items: missing };
  const ids = resRows_().map(r => r.id);
  let id = clean_(b.id, 20);
  if (id && (!/^OWNER\d+$/.test(id) || ids.indexOf(id) >= 0)) return { ok: false, error: 'id taken or invalid', id: id };
  if (!id) { let n = 1; while (ids.indexOf('OWNER' + n) >= 0) n++; id = 'OWNER' + n; }
  sh_(RES_TAB).appendRow([new Date(), id, lines.join('\n'), total, date ? "'" + date : '', time ? "'" + time : '', clean_(b.name, 80) || 'Offline buyer', '',
    clean_(b.note, 300) || 'Arranged offline', 'Confirmed', '', false]);
  syncStatuses_();
  return { ok: true, id: id, total: total };
}
