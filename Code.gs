/**
 * RMAS POSM Deployment Tracker — Google Sheets backend (Apps Script)
 * Replaces the old POSM KPI Command Center backend in the SAME Apps Script project,
 * so the Web app URL stays the same and the TS KPI Grading tool needs no change.
 *
 * Setup: open the old Command Center Apps Script → replace all code with this file → Save →
 *        Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy.
 *
 * Tabs it creates: DT_Stores, DT_Status, DT_Log (the old Command Center tabs are left as they are).
 * Settings (password, unlocked months, close day) are in Script Properties: DT_PASS, DT_UNLOCKED, DT_CLOSE_DAY.
 *
 * Read API (used by the tracker and by the TS KPI Grading tool):
 *   ?action=all                        everything the tracker needs
 *   ?action=tsScores&period=2026-10    deployment score per Trade Supervisor for one month
 *                                      (period optional → current grading month)
 *   ?action=tsScores&period=2026-10&ts=CARLOS, LUIS VERGEL   one TS only
 *   ?action=getAll[&period=2026-10]    store × BU rows in the format the TS KPI Grading tool reads
 *                                      (same as the old Deployment KPI Tool, so only the URL changes)
 */

var BUS = ['AV', 'HA', 'AC'];
var START_PERIOD = '2026-10';
var TZ = 'Asia/Manila';
var GRADES = [[95, 'Outstanding'], [90, 'Excellent'], [85, 'Very Good'], [80, 'Good'], [70, 'Needs Improvement'], [0, 'Poor']];
var STATUS_LABEL = { C: 'Complete', XS: 'Complete – no stock', XO: 'Complete – outdated display', I: 'Incomplete POSM', NA: 'N/A (no display)' };

/* ---------- entry points ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'tsScores') return out_(tsScores_(p.period, p.ts));
    if (p.action === 'stores') return out_(storeScores_(p.period, p.ts));
    if (p.action === 'getAll') return out_(gradingFeed_(p.period));
    return out_(all_());
  } catch (err) { return out_({ ok: false, error: String(err && err.message || err) }); }
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { return out_({ ok: false, error: 'Bad request' }); }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    if (!checkPass_(body.pass)) return out_({ ok: false, error: 'Wrong password.', auth: false });
    switch (body.action) {
      case 'login': return out_({ ok: true, hasStores: sheet_(TAB.Stores).getLastRow() > 1 });
      case 'set': return out_(setItems_(body.items || []));
      case 'importStores': return out_(importStores_(body.stores || [], body.note || ''));
      case 'unlock': return out_(setUnlock_(body.period, true, body.reason));
      case 'relock': return out_(setUnlock_(body.period, false));
      case 'changePass': return out_(changePass_(body.newPass));
      case 'setCloseDay': return out_(setCloseDay_(body.closeDay));
      default: return out_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- sheets ---------- */
var TAB = { Stores: 'DT_Stores', Status: 'DT_Status', Log: 'DT_Log' };
var HEADERS = {
  DT_Stores: ['id', 'name', 'ts', 'region', 'dealer', 'agency'],
  DT_Status: ['period', 'storeId', 'field', 'value', 'by', 'at'],
  DT_Log: ['at', 'event']
};
function sheet_(name) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function rows_(name) {
  var sh = sheet_(name), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HEADERS[name].length).getDisplayValues();
}
function log_(text) {
  sheet_(TAB.Log).appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), text]);
}

/* ---------- settings ---------- */
function props_() { return PropertiesService.getScriptProperties(); }
function checkPass_(p) { return String(p || '') === (props_().getProperty('DT_PASS') || 'rmas2026'); }
function unlocked_() { try { return JSON.parse(props_().getProperty('DT_UNLOCKED') || '{}'); } catch (x) { return {}; } }
function closeDay_() { var v = props_().getProperty('DT_CLOSE_DAY'); return v === null ? 7 : Number(v); }

/* ---------- periods & locking ----------
 * A grading month runs until day CLOSE_DAY of the next month (default 7):
 * e.g. October 2026 = everything entered up to Nov 7; from Nov 8 entries count for November.
 */
function currentPeriod_() {
  var d = new Date(Date.now() - closeDay_() * 86400000);
  var p = Utilities.formatDate(d, TZ, 'yyyy-MM');
  return p < START_PERIOD ? START_PERIOD : p;
}
function isLocked_(p) { return p < currentPeriod_() && !unlocked_()[p]; }
function closesOn_(p) {
  var y = +p.slice(0, 4), m = +p.slice(5, 7) + 1; if (m > 12) { m = 1; y++; }
  var cd = closeDay_();
  if (cd === 0) { var last = new Date(y, m - 1, 0); return Utilities.formatDate(last, TZ, 'yyyy-MM-dd'); }
  return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + cd).slice(-2);
}
function validPeriod_(p) { return /^\d{4}-\d{2}$/.test(p || '') && p >= START_PERIOD && p <= currentPeriod_(); }

/* ---------- reads ---------- */
function stores_() {
  return rows_(TAB.Stores).map(function (r) { return { id: r[0], name: r[1], ts: r[2], region: r[3], dealer: r[4], agency: r[5] }; });
}
function months_(onlyPeriod) {
  var months = {};
  rows_(TAB.Status).forEach(function (r) {
    if (!r[3]) return;
    if (onlyPeriod && r[0] !== onlyPeriod) return;
    (months[r[0]] = months[r[0]] || {})[r[1] + '|' + r[2]] = { s: r[3], by: r[4], at: r[5] };
  });
  return months;
}
function all_() {
  var log = rows_(TAB.Log).slice(-150).map(function (r) { return r[0] + ' · ' + r[1]; });
  return { ok: true, stores: stores_(), months: months_(), unlocked: unlocked_(), closeDay: closeDay_(),
           currentPeriod: currentPeriod_(), log: log, serverTime: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') };
}
function grade_(pct) { if (pct === null) return 'No data'; for (var i = 0; i < GRADES.length; i++) if (pct >= GRADES[i][0]) return GRADES[i][1]; }
function tally_(list, m) {
  var t = { complete: 0, exempt: 0, incomplete: 0, notChecked: 0, na: 0 };
  list.forEach(function (s) {
    BUS.forEach(function (bu) {
      var v = (m[s.id + '|' + bu] || {}).s || '';
      if (v === 'C') t.complete++; else if (v === 'XS' || v === 'XO') t.exempt++; else if (v === 'I') t.incomplete++; else if (v === 'NA') t.na++; else t.notChecked++;
    });
  });
  t.checked = t.complete + t.exempt + t.incomplete;
  t.pct = t.checked ? Math.round((t.complete + t.exempt) / t.checked * 1000) / 10 : null;
  t.grade = grade_(t.pct);
  return t;
}
function topOf_(o) { var best = '', n = -1; for (var k in o) if (o[k] > n) { n = o[k]; best = k; } return best; }
function tsScores_(period, onlyTs) {
  period = period || currentPeriod_();
  var m = months_(period)[period] || {};
  var byTs = {};
  stores_().forEach(function (s) {
    var x = byTs[s.ts] = byTs[s.ts] || { stores: [], r: {}, a: {} };
    x.stores.push(s); x.r[s.region] = (x.r[s.region] || 0) + 1; if (s.agency) x.a[s.agency] = (x.a[s.agency] || 0) + 1;
  });
  var rows = Object.keys(byTs).sort().filter(function (n) { return !onlyTs || n === String(onlyTs).toUpperCase().trim(); }).map(function (n) {
    var x = byTs[n], t = tally_(x.stores, m), perBu = {};
    BUS.forEach(function (bu) {
      var tb = { complete: 0, exempt: 0, incomplete: 0 };
      x.stores.forEach(function (s) { var v = (m[s.id + '|' + bu] || {}).s || ''; if (v === 'C') tb.complete++; else if (v === 'XS' || v === 'XO') tb.exempt++; else if (v === 'I') tb.incomplete++; });
      var c = tb.complete + tb.exempt + tb.incomplete; perBu[bu] = c ? Math.round((tb.complete + tb.exempt) / c * 1000) / 10 : null;
    });
    return { ts: n, region: topOf_(x.r), agency: topOf_(x.a), stores: x.stores.length, complete: t.complete, exempt: t.exempt,
             incomplete: t.incomplete, notChecked: t.notChecked, checked: t.checked, pct: t.pct, grade: t.grade, perBu: perBu };
  });
  return { ok: true, period: period, locked: isLocked_(period), closesOn: closesOn_(period), rows: rows };
}
function storeScores_(period, onlyTs) {
  period = period || currentPeriod_();
  var m = months_(period)[period] || {};
  var rows = stores_().filter(function (s) { return !onlyTs || s.ts === String(onlyTs).toUpperCase().trim(); }).map(function (s) {
    var t = tally_([s], m), o = { id: s.id, store: s.name, ts: s.ts, region: s.region, dealer: s.dealer, agency: s.agency, pct: t.pct, remarks: (m[s.id + '|note'] || {}).s || '' };
    BUS.forEach(function (bu) { o[bu] = STATUS_LABEL[(m[s.id + '|' + bu] || {}).s] || 'Not checked'; });
    return o;
  });
  return { ok: true, period: period, locked: isLocked_(period), rows: rows };
}

/* Feed for the TS KPI Grading tool: one row per store per BU that RMAS has checked this month.
 * Not-checked and N/A displays are left out, so the grading counts the same displays the tracker does. */
var FEED_STATUS = { C: 'Complete', XS: 'Incomplete - Full Deployment (No Stock)', XO: 'Incomplete - Full Deployment (Outdated Display)', I: 'Incomplete' };
function gradingFeed_(period) {
  period = validPeriod_(period) ? period : currentPeriod_();
  var m = months_(period)[period] || {}, stores = stores_(), tsRegion = {};
  tsScores_(period).rows.forEach(function (r) { tsRegion[r.ts] = r.region; });
  var records = [];
  stores.forEach(function (s) {
    BUS.forEach(function (bu) {
      var v = (m[s.id + '|' + bu] || {}).s || '';
      if (!FEED_STATUS[v]) return;
      records.push({ ts: s.ts, region: tsRegion[s.ts] || s.region, store: s.name, bu: bu, status: FEED_STATUS[v],
                     remarks: (m[s.id + '|note'] || {}).s || '', agency: s.agency, dealer: s.dealer, storeRegion: s.region });
    });
  });
  return { ok: true, source: 'POSM Deployment Tracker', period: period, locked: isLocked_(period), records: records, masterChanges: [] };
}

/* ---------- writes ---------- */
function setItems_(items) {
  var sh = sheet_(TAB.Status), data = rows_(TAB.Status), index = {};
  data.forEach(function (r, i) { index[r[0] + '|' + r[1] + '|' + r[2]] = i + 2; });
  var now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), applied = 0, rejected = [], appends = [];
  items.forEach(function (it) {
    var p = it.period, k = String(it.key || ''), parts = k.split('|');
    if (!validPeriod_(p) || isLocked_(p)) { rejected.push(k); return; }
    if (parts.length !== 2 || (BUS.indexOf(parts[1]) < 0 && parts[1] !== 'note')) { rejected.push(k); return; }
    var v = String(it.value || '').slice(0, 500), row = index[p + '|' + k];
    if (parts[1] !== 'note' && v && !STATUS_LABEL[v]) { rejected.push(k); return; }
    var rec = [p, parts[0], parts[1], v, it.by || 'RMAS', now];
    if (row) sh.getRange(row, 1, 1, 6).setValues([rec]);
    else if (v) { appends.push(rec); index[p + '|' + k] = -1; }
    applied++;
  });
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, 6).setValues(appends);
  return { ok: true, applied: applied, rejected: rejected, lockedPeriod: rejected.length ? true : false };
}
function importStores_(stores, note) {
  var clean = stores.filter(function (s) { return s && s.id && s.name && s.ts; }).map(function (s) {
    return [s.id, s.name, s.ts, s.region || '', s.dealer || '', s.agency || ''];
  });
  if (!clean.length) return { ok: false, error: 'No stores in the upload.' };
  var sh = sheet_(TAB.Stores);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  sh.getRange(2, 1, clean.length, 6).setNumberFormat('@').setValues(clean);
  log_('Store list updated: ' + clean.length + ' stores' + (note ? ' (' + note + ')' : ''));
  return { ok: true, count: clean.length };
}
function setUnlock_(p, on, reason) {
  if (!/^\d{4}-\d{2}$/.test(p || '') || p >= currentPeriod_()) return { ok: false, error: 'Only past months can be unlocked.' };
  if (on && !reason) return { ok: false, error: 'Please give a reason.' };
  var u = unlocked_();
  if (on) u[p] = true; else delete u[p];
  props_().setProperty('DT_UNLOCKED', JSON.stringify(u));
  log_(p + (on ? ' unlocked: ' + reason : ' re-locked'));
  return { ok: true, unlocked: u };
}
function changePass_(np) {
  if (!np || String(np).length < 4) return { ok: false, error: 'New password needs at least 4 characters.' };
  props_().setProperty('DT_PASS', String(np));
  log_('RMAS password changed');
  return { ok: true };
}
function setCloseDay_(d) {
  d = Number(d);
  if (!(d >= 0 && d <= 15)) return { ok: false, error: 'Close day must be 0 to 15.' };
  props_().setProperty('DT_CLOSE_DAY', String(d));
  log_('Monthly close day set to ' + (d === 0 ? 'last day of the month' : 'day ' + d + ' of the next month'));
  return { ok: true, closeDay: d, currentPeriod: currentPeriod_() };
}
