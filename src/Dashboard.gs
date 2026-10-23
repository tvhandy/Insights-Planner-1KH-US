/**
 * Insights Planner dashboard – server side
 *
 * Serves the DashboardPage HTML file two ways:
 *  - Basecamp > Open dashboard (a large window inside the Sheet)
 *  - As a web app link (Deploy > New deployment > Web app)
 *
 * The page reads the Planner, Team, Lists and Notes tabs live, writes Stage
 * and Booking changes back to the Planner, and adds card notes to the Notes tab.
 */

const DASH_TITLE = 'Insights Planner | 1KH US [Beta]';          // the browser tab's name
// The browser tab's icon: a public .png or .ico link. This is the 1000heads site icon; swap in another if you prefer.
const DASH_FAVICON = 'https://1000heads.com/wp-content/uploads/2022/12/1000heads_logo_purple-1.png';
const PLANNER_SHEET_ID = '1as50OMGMFfcaBEgopWL9aaRfAxHLuH8_oBBiVuLWgIg';

// Works from the Sheet's menu and from the web app
function ss_() {
  return SpreadsheetApp.getActive() || SpreadsheetApp.openById(PLANNER_SHEET_ID);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.state && (e.parameter.code || e.parameter.error)) return finishConnect_(e);
  const page = HtmlService.createHtmlOutputFromFile('DashboardPage')
    .setTitle(DASH_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  try { page.setFaviconUrl(DASH_FAVICON); } catch (e) { console.log('Tab icon not set: ' + e.message); }
  return page;
}

function openDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('DashboardPage').setWidth(1500).setHeight(920);
  SpreadsheetApp.getUi().showModalDialog(html, DASH_TITLE);
}

// Brief and replies for one card, loaded when it opens
function getCardDetail(id) {
  const pl = ss_().getSheetByName('Planner');
  const last = pl.getLastRow();
  const ids = pl.getRange(2, 1, Math.max(last - 1, 1), 1).getValues().map(r => String(r[0]));
  const k = ids.indexOf(String(id));
  if (k === -1) throw new Error('Could not find ' + id + ' in the Planner.');
  const row = pl.getRange(k + 2, 1, 1, 28).getValues()[0];
  return { brief: String(row[23] || ''), replies: replies_(row[27]) };
}

function replyFlag_(cell) {
  try { const j = JSON.parse(String(cell || '') || 'null'); return j && j.flag ? j.flag : null; } catch (e) { return null; }
}
function replies_(cell) {
  try { const j = JSON.parse(String(cell || '') || 'null'); return j && Array.isArray(j.list) ? j.list : []; } catch (e) { return []; }
}

// Everything the page needs, in one call
function getDashboardData(opts) {
  const lite = !!(opts && opts.lite);
  const ss = ss_();
  const tz = ss.getSpreadsheetTimeZone();
  const iso = v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : '';
  const str = v => (v === null || v === undefined) ? '' : String(v).trim();

  // Planner: A ID … X Brief, Y Booking, Z Team on it, AA Split, AB Basecamp replies (JSON from the sync)
  const pl = ss.getSheetByName('Planner');
  const last = pl.getLastRow();
  const rows = last < 2 ? [] : pl.getRange(2, 1, Math.max(last - 1, 1), 31).getValues();
  const items = [];
  rows.forEach((r, i) => {
    if (str(r[3]) === '') return;
    items.push({
      row: i + 2, id: str(r[0]), account: str(r[1]), project: str(r[2]), title: str(r[3]),
      type: str(r[4]), scale: str(r[5]), est: Number(r[7]) || 0,
      owner: str(r[8]), assigned: str(r[9]), stage: str(r[10]) || 'Not started',
      start: iso(r[11]), due: iso(r[12]), workStart: iso(r[13]),
      priority: str(r[16]) || 'Medium', link: str(r[17]), notes: str(r[18]),
      source: str(r[19]), added: iso(r[20]), brief: str(r[23]), booking: str(r[24]) || 'Confirmed', split: str(r[26]), replies: replies_(r[27]), flag: replyFlag_(r[27]), actual: r[28] === '' || r[28] === null ? null : Number(r[28]), likelihood: str(r[29]), stageSince: iso(r[30])
    });
  });

  // Team: people, time off log, company holidays
  const tm = ss.getSheetByName('Team');
  const team = tm.getRange('A5:P14').getValues().filter(r => str(r[0])).map(r => ({
    name: str(r[0]), role: str(r[1]), bill: Number(r[2]) || 0,
    start: iso(r[3]), last: iso(r[4]), include: str(r[5]) === 'Yes',
    projectRole: str(r[15]) === 'Oversight' ? 'Oversight' : 'Hands-on'
  }));
  // Time off lives on its own tab (ID, Name, Start, End, Type, Day, Coverage / notes, Added by, Added)
  const tos = ss.getSheetByName(TIMEOFF_SHEET);
  const timeOff = tos && tos.getLastRow() > 1
    ? tos.getRange(2, 1, tos.getLastRow() - 1, 9).getValues().map((r, k) => ({ r: r, row: k + 2 })).filter(x => str(x.r[1]) && x.r[2] instanceof Date).map(x => ({
        id: str(x.r[0]), row: x.row, name: str(x.r[1]), start: iso(x.r[2]), end: iso(x.r[3] instanceof Date ? x.r[3] : x.r[2]),
        type: str(x.r[4]) || 'Out', half: str(x.r[5]).indexOf('Half day') === 0, part: str(x.r[5]).indexOf('AM') !== -1 ? 'AM' : str(x.r[5]).indexOf('Half day') === 0 ? 'PM' : '', note: str(x.r[6]), by: str(x.r[7])
      }))
    : tm.getRange('A21:E45').getValues().filter(r => str(r[0]) && r[1] instanceof Date).map(r => ({
        name: str(r[0]), start: iso(r[1]), end: iso(r[2] instanceof Date ? r[2] : r[1]), type: str(r[3]) || 'Out', half: false, note: str(r[4])
      }));
  const holidays = tm.getRange('H21:I45').getValues().filter(r => r[0] instanceof Date).map(r => ({
    date: iso(r[0]), name: str(r[1]) || 'Holiday'
  }));

  // Lists for dropdowns
  const ls = ss.getSheetByName('Lists');
  const col = a1 => ls.getRange(a1).getValues().map(r => str(r[0])).filter(String);
  const typeRanges = {};
  ls.getRange('D2:F30').getValues().forEach(r => { if (str(r[0])) typeRanges[str(r[0])] = [Number(r[1]) || 0, Number(r[2]) || 0]; });
  const scaleFactors = {};
  ls.getRange('H2:I4').getValues().forEach(r => { if (str(r[0])) scaleFactors[str(r[0])] = Number(r[1]) || 0; });
  const leadRaw = Number(ls.getRange('Q2').getValue());
  const summerFridays = str(ls.getRange('Q5').getValue()) !== 'No';
  // Review stages book only a small revision allowance (Lists tab P7:Q11)
  const rv = ls.getRange('Q7:Q11').getValues().map(r => r[0]);
  const num = (v, dflt) => { const n = Number(v); return v === '' || isNaN(n) ? dflt : n; };
  const listOf = (v, dflt) => (str(v) || dflt).split(/\s*,\s*/).filter(String);
  const review = { pct: (p => p > 1 ? p / 100 : p)(num(rv[0], 0.1)), min: num(rv[1], 0.1), max: num(rv[2], 0.5),
    stages: listOf(rv[3], 'Internal review, Client review, Revisions'), paused: listOf(rv[4], 'Blocked, On hold') };
  const leadShare = leadRaw > 0 && leadRaw <= 1 ? leadRaw : 0.75;

  // Card notes (Notes tab: Timestamp, ID, Deliverable, Author, Note)
  const ns = ss.getSheetByName('Notes');
  const nLast = ns ? ns.getLastRow() : 1;
  const notes = nLast < 2 ? [] : ns.getRange(2, 1, nLast - 1, 5).getValues()
    .filter(r => str(r[1]) && str(r[4]))
    .map(r => ({ at: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "MMM d, h:mm a") : str(r[0]), id: str(r[1]), author: str(r[3]), text: str(r[4]) }));

  // Activity feed: the latest 300 events, newest first
  const as = ss.getSheetByName('Activity');
  const aLast = as ? as.getLastRow() : 1;
  const activity = aLast < 2 ? [] : as.getRange(Math.max(2, aLast - 299), 1, Math.min(300, aLast - 1), 7).getValues()
    .filter(r => r[0] instanceof Date).reverse()
    .map(r => ({ at: r[0].toISOString(), type: str(r[1]), id: str(r[2]), title: str(r[3]), who: str(r[4]), text: str(r[5]), link: str(r[6]) }));

  let connected = [], sync = null;
  try { connected = services_().map(c => c.name); } catch (e) {}
  try { sync = syncStatus_(); } catch (e) { sync = { error: e.message }; }

  // Light payload for the page: parsed schedule and a reply summary instead of full briefs and replies
  if (lite) {
    const isoD = v => Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    const P2 = s => { if (!s) return null; const a = s.split('-').map(Number); return new Date(a[0], a[1] - 1, a[2]); };
    const weekAgo = new Date(Date.now() - 7 * 864e5).toISOString();
    items.forEach(i => {
      const anchor = P2(i.due) || new Date();
      let w = null, from = '';
      for (let k = i.replies.length - 1; k >= 0 && !w; k--) { w = parseWbs_(i.replies[k].text, anchor); if (w) from = 'from a reply by ' + i.replies[k].by + ', ' + Utilities.formatDate(new Date(i.replies[k].at), tz, 'MMM d'); }
      if (!w) { w = parseWbs_(i.brief, anchor); if (w) from = 'from the brief'; }
      i.wbsSteps = w ? w.map(x => ({ label: x.label, s: isoD(x.s), e: isoD(x.e), kind: x.kind, note: x.note })) : null;
      i.wbsFromText = from;
      const last = i.replies[i.replies.length - 1];
      i.lastReply = last ? { by: last.by, at: last.at, text: String(last.text || '').slice(0, 200) } : null;
      i.replyCount = i.replies.length;
      i.newReplies = i.replies.filter(x => x.at && x.at >= weekAgo).length;
      i.brief = ''; i.replies = []; i.lite = true;
    });
  }

  return {
    items, team, timeOff, holidays, notes, activity, connected, sync, me: displayName_(),
    lists: { accounts: col('A2:A30'), types: col('D2:D30'), scales: col('H2:H4'), stages: col('K2:K15'), typeRanges: typeRanges, scaleFactors: scaleFactors },
    version: typeof PLANNER_VERSION === 'string' ? PLANNER_VERSION : '', leadShare: leadShare, summerFridays: summerFridays, review: review, timesheets: timesheets_(ss, iso, str),
    today: iso(new Date()),
    sheetUrl: ss.getUrl()
  };
}

// Adds a note to a card; the Notes tab keeps the full history
function addNote(id, text) {
  const t = String(text || '').trim();
  if (!t) throw new Error('Write a note first.');
  if (t.length > 5000) throw new Error('Notes are limited to 5,000 characters.');
  const ss = ss_();
  let ns = ss.getSheetByName('Notes');
  if (!ns) { ns = ss.insertSheet('Notes'); ns.appendRow(['Timestamp', 'ID', 'Deliverable', 'Author', 'Note']); }
  const pl = ss.getSheetByName('Planner');
  const ids = pl.getRange(2, 1, Math.max(pl.getLastRow() - 1, 1), 4).getValues();
  const row = ids.find(r => String(r[0]) === String(id));
  if (!row) throw new Error('Could not find ' + id + ' in the Planner.');
  const author = displayName_();
  const now = new Date();
  ns.appendRow([now, id, row[3], author, t]);
  logActivity_('note', id, row[3], matchTeamName_(author) || author, t.slice(0, 400), '');
  return { at: Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MMM d, h:mm a"), id: id, author: author, text: t };
}

// The viewer's name: from their Basecamp connection if they have one, otherwise their email
function displayName_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) return 'Someone';
  try {
    const reg = registry_();
    const hit = Object.keys(reg).map(k => reg[k]).find(u => String(u.email).toLowerCase() === email);
    if (hit && hit.name && hit.name !== hit.email) return hit.name;
  } catch (e) {}
  return email.split('@')[0].split(/[._-]/).map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
}

// Write-backs check the row still holds the same ID, so a sort in the Sheet can't misfile a change
function setStage(row, id, stage) {
  const r = writeCell_(row, id, 11, stage);
  logRow_(r.row, 'stage', 'Moved to ' + stage);
  return r;
}

function setBooking(row, id, booking) {
  const r = writeCell_(row, id, 25, booking);
  logRow_(r.row, 'stage', 'Booking set to ' + booking);
  return r;
}

// Bulk clean-up: one stage for many rows
function setStages(list, stage) {
  const allowed = ['Delivered', 'Dismissed', 'Not started', 'On hold'];
  if (allowed.indexOf(stage) === -1) throw new Error('Not a bulk stage: ' + stage);
  const done = [];
  (list || []).forEach(x => { try { const r = writeCell_(x.row, x.id, 11, stage); logRow_(r.row, 'stage', 'Moved to ' + stage + ' (bulk clean-up)'); done.push(x.id); } catch (e) {} });
  return { done: done, stage: stage };
}

// How long a delivered piece of work actually took, for checking estimates
function setActual(row, id, days) {
  const v = days === '' || days === null ? '' : Math.round(Number(days) * 100) / 100;
  if (v !== '' && !(v >= 0 && v <= 200)) throw new Error('Enter days between 0 and 200.');
  const r = writeCell_(row, id, 29, v);
  if (v !== '') logRow_(r.row, 'stage', 'Actual days: ' + v);
  return { row: r.row, id: id, actual: v };
}

// Adds a dashboard change to the Activity feed, attributed to the person who made it
// ===== Time off: add, edit and delete from the dashboard =====
const TIMEOFF_SHEET = 'Time off';
const TIMEOFF_TYPES = ['PTO', 'Sick', 'OOO', 'Leave', 'Other'];
function timeOffSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(TIMEOFF_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TIMEOFF_SHEET);
    sh.getRange(1, 1, 1, 9).setValues([['ID', 'Name', 'Start', 'End', 'Type', 'Day', 'Coverage / notes', 'Added by', 'Added']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function cleanTimeOff_(f) {
  const str = v => String(v == null ? '' : v).trim();
  const P = s => { const a = str(s).split('-').map(Number); return a.length === 3 && a[0] ? new Date(a[0], a[1] - 1, a[2], 12) : null; };
  const name = str(f.name), start = P(f.start), end = P(f.end) || start;
  if (!name) throw new Error('Pick who is out.');
  if (!start) throw new Error('Add a start date.');
  if (end < start) throw new Error('The end date is before the start date.');
  const type = TIMEOFF_TYPES.indexOf(str(f.type)) !== -1 ? str(f.type) : 'PTO';
  const half = str(f.day).indexOf('Half day') === 0, part = str(f.day).indexOf('AM') !== -1 ? 'AM' : 'PM';
  if (half && +end !== +start) throw new Error('A half day has to be a single day. Use two entries for two half days.');
  const cover = str(f.cover), note = str(f.note).replace(/\s*coverage\s*:[^.;\n]*[.;]?/i, '').trim();
  // Midday dates so the day can't slip when the script and the Sheet are in different time zones
  return { name, start, end, type, day: half ? 'Half day (' + part + ')' : 'Full day', notes: [cover ? 'Coverage: ' + cover.replace(/\.$/, '') + '.' : '', note].filter(String).join(' ') };
}
function addTimeOff(f) {
  const x = cleanTimeOff_(f);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = timeOffSheet_();
    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => Number(String(r[0]).replace(/\D/g, '')) || 0) : [];
    const id = 'TO-' + ('00' + (Math.max(0, ...ids) + 1)).slice(-3);
    sh.getRange(last + 1, 1, 1, 9).setValues([[id, x.name, x.start, x.end, x.type, x.day, x.notes, displayName_(), new Date()]]);
    sh.getRange(last + 1, 3, 1, 2).setNumberFormat('ddd, mmm d, yyyy');
    sh.getRange(last + 1, 9).setNumberFormat('mmm d, yyyy h:mm am/pm');
    logActivity_('timeoff', id, x.name + ' · ' + x.type, displayName_(), 'Added ' + x.type.toLowerCase() + (x.day !== 'Full day' ? ' (' + x.day.toLowerCase() + ')' : '') + ': ' + timeOffRange_(x), '');
    return id;
  } finally { lock.releaseLock(); }
}
function timeOffRow_(sh, id) {
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
  const k = ids.indexOf(String(id));
  if (k === -1) throw new Error('Could not find ' + id + '. Someone may have changed it. Refresh and try again.');
  return k + 2;
}
function updateTimeOff(id, f) {
  const x = cleanTimeOff_(f);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = timeOffSheet_(), r = timeOffRow_(sh, id);
    sh.getRange(r, 2, 1, 6).setValues([[x.name, x.start, x.end, x.type, x.day, x.notes]]);
    logActivity_('timeoff', id, x.name + ' · ' + x.type, displayName_(), 'Changed to ' + timeOffRange_(x) + (x.day !== 'Full day' ? ' (' + x.day.toLowerCase() + ')' : ''), '');
    return id;
  } finally { lock.releaseLock(); }
}
function deleteTimeOff(id) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = timeOffSheet_(), r = timeOffRow_(sh, id);
    const v = sh.getRange(r, 1, 1, 6).getValues()[0];
    sh.deleteRow(r);
    logActivity_('timeoff', id, v[1] + ' · ' + v[4], displayName_(), 'Removed ' + String(v[4]).toLowerCase() + ': ' + timeOffRange_({ start: v[2], end: v[3] }), '');
    return id;
  } finally { lock.releaseLock(); }
}
function timeOffRange_(x) {
  const tz = ss_().getSpreadsheetTimeZone(), f = d => Utilities.formatDate(d, tz, 'EEE, MMM d');
  return +x.start === +x.end ? f(x.start) : f(x.start) + ' – ' + f(x.end);
}

// ===== Timesheets: hours people log (Timesheets tab), compared with booked days on the Forecast =====
// Columns: Week of (the Monday), Name, Client, Hours. Paste an export in, or type it.
function timesheets_(ss, iso, str) {
  const sh = ss.getSheetByName('Timesheets');
  if (!sh || sh.getLastRow() < 2) return [];
  const since = new Date(Date.now() - 26 * 7 * 864e5);
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(r => r[0] instanceof Date && r[0] >= since && str(r[1]) && Number(r[3]) > 0)
    .map(r => ({ week: iso(r[0]), name: str(r[1]), client: str(r[2]), hours: Number(r[3]) }));
}

// Adds the every-2-hours backup schedule under the person viewing (once), so the planner never depends on one account
function ensureMyBackup() {
  try { return ensureBackupTrigger_(); } catch (e) { return false; }
}

// Card drawer: set a deliverable's client. Optionally the same for every row from its Basecamp project, and
// remember the project so future to-dos from it get this client automatically.
function setClient(row, id, account, everyRow) {
  account = String(account || '').trim();
  if (!account) throw new Error('Type or pick a client.');
  const sheet = ss_().getSheetByName('Planner');
  const r = writeCell_(row, id, 2, account);
  const project = String(sheet.getRange(r.row, 3).getValue()).trim();
  let others = 0;
  const generic = /^(Recurring \(projected\)|Planning \(booked in dashboard\))$/.test(project);
  addClient_(account, everyRow && project && !generic ? cleanProjectName_(project) : '');
  if (everyRow && project && !generic) {
    const last = sheet.getLastRow();
    const vals = sheet.getRange(2, 1, last - 1, 3).getValues();
    vals.forEach((v, k) => {
      const c = String(v[1]).trim();
      if (k + 2 !== r.row && String(v[2]).trim() === project && (!c || c === 'Other')) { sheet.getRange(k + 2, 2).setValue(account); others++; }
    });
  }
  logRow_(r.row, 'stage', 'Client set to ' + account + (others ? ' (and ' + others + ' more from ' + project + ')' : ''));
  return { row: r.row, others: others };
}

// Feedback button: saved to the Feedback tab and the Activity feed; optionally posted in the Hub's Campfire
function sendFeedback(f) {
  const text = String((f && f.text) || '').trim().slice(0, 2000);
  if (!text) throw new Error('Add a few words first.');
  const ss = ss_();
  let sh = ss.getSheetByName('Feedback');
  if (!sh) {
    sh = ss.insertSheet('Feedback');
    sh.getRange(1, 1, 1, 5).setValues([['When', 'Who', 'Tab', 'Feedback', 'Status']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const who = displayName_();
  sh.appendRow([new Date(), who, String(f.tab || ''), text, 'New']);
  logActivity_('feedback', '', 'Planner feedback (' + (f.tab || '') + ')', who, text.slice(0, 400), '');
  if (f.share) { try { campfirePost_('💬 Planner feedback from ' + who + ' (' + (f.tab || 'dashboard') + '): ' + text); } catch (e) { console.log('Campfire post failed: ' + e.message); } }
  return true;
}

// Estimate check: write a new day range for a deliverable type to the Lists tab
function applyTypeRange(type, min, max) {
  const ls = ss_().getSheetByName('Lists');
  const names = ls.getRange('D2:D30').getValues().map(r => String(r[0]).trim());
  const k = names.indexOf(String(type).trim());
  if (k === -1) throw new Error('Could not find ' + type + ' on the Lists tab.');
  if (!(min > 0) || !(max >= min)) throw new Error('The range needs a minimum above 0 and a maximum at least as big.');
  const old = ls.getRange(k + 2, 5, 1, 2).getValues()[0];
  ls.getRange(k + 2, 5, 1, 2).setValues([[min, max]]);
  logActivity_('settings', '', type, displayName_(), 'Day range changed from ' + old[0] + '–' + old[1] + ' to ' + min + '–' + max + ' days (from actual days)', '');
  return { type: type, min: min, max: max };
}

function logRow_(row, type, text) {
  try {
    const v = ss_().getSheetByName('Planner').getRange(row, 1, 1, 18).getValues()[0];
    const full = displayName_();
    logActivity_(type, v[0], v[3], matchTeamName_(full) || full, text, v[17]);
  } catch (e) { console.log('Could not log: ' + e.message); }
}

// Saves a per-person split ("Pranav P. 4; Tom H. 0.5") and sets Override days to its total
function setSplit(row, id, text, total) {
  const r = writeCell_(row, id, 27, String(text || ''));
  if (text && Number(total) > 0) writeCell_(r.row, id, 7, Math.round(Number(total) * 100) / 100);
  logRow_(r.row, 'split', text ? 'Days per person set: ' + text : 'Back to the default split by role');
  return { row: r.row, id: id, split: String(text || ''), total: Number(total) || 0 };
}

// Hands a deliverable from someone leaving to a teammate: new split, and a new Owner when needed
function handOver(row, id, text, total, newOwner, from, to) {
  let r = { row: row };
  if (text) r = setSplit(row, id, text, total);
  if (newOwner) r = writeCell_(r.row, id, 9, newOwner);
  if (from && to) logRow_(r.row, 'handover', 'Handed ' + from + '\'s work to ' + to);
  return { row: r.row, id: id };
}

// Book work from the Planning tab: a Tentative row with no Basecamp link
function bookWork(f) {
  const title = String(f.title || '').trim();
  if (!title) throw new Error('Give the work a name.');
  if (!f.due) throw new Error('Add a due date.');
  const sheet = ss_().getSheetByName('Planner');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const r = firstEmptyRow_(sheet);
    const people = (f.people || []).filter(String);
    const P = s => { const a = String(s).split('-').map(Number); return new Date(a[0], a[1] - 1, a[2], 12); };
    const v = {};
    v[COL.id] = nextId_(sheet); v[COL.client] = f.account || 'Other'; v[COL.project] = 'Planning (booked in dashboard)';
    v[COL.deliverable] = title; v[COL.type] = f.type || inferType_(title); v[COL.scale] = f.scale || '';
    if (Number(f.days) > 0) v[COL.override] = Number(f.days);
    v[COL.owner] = people[0] || ''; v[COL.assigned] = people.join(', ');
    v[COL.stage] = NEW_STAGE; v[COL.start] = f.start ? P(f.start) : ''; v[COL.due] = P(f.due);
    v[COL.priority] = DEFAULT_PRIORITY; v[COL.notes] = String(f.notes || '');
    v[COL.source] = 'Booked in dashboard'; v[COL.added] = new Date(); v[COL.booking] = 'Tentative';
    v[COL.likelihood] = f.likelihood || 'Likely'; v[COL.stageSince] = new Date();
    Object.keys(v).forEach(c => sheet.getRange(r, Number(c)).setValue(v[c]));
    const full = displayName_();
    logActivity_('new', v[COL.id], title, matchTeamName_(full) || full, 'Booked as Tentative (' + v[COL.likelihood] + '), due ' + f.due + (people.length ? ' · ' + people.join(', ') : ' · needs an owner'), '');
    return { id: v[COL.id], row: r };
  } finally { lock.releaseLock(); }
}

// Planning actions: assign people, move dates, set likelihood
function setPlanFields(row, id, f) {
  const P = s => { const a = String(s).split('-').map(Number); return new Date(a[0], a[1] - 1, a[2], 12); };
  let r = { row: row }; const said = [];
  if (f.people) { const p = f.people.filter(String); r = writeCell_(r.row, id, 9, p[0] || ''); writeCell_(r.row, id, 10, p.join(', ')); said.push(p.length ? 'Assigned to ' + p.join(', ') : 'Unassigned'); }
  if (f.start !== undefined) { r = writeCell_(r.row, id, 12, f.start ? P(f.start) : ''); }
  if (f.due) { r = writeCell_(r.row, id, 13, P(f.due)); said.push('Dates: ' + (f.start || '…') + ' to ' + f.due); }
  if (f.likelihood) { r = writeCell_(r.row, id, 30, f.likelihood); said.push('Likelihood: ' + f.likelihood); }
  if (said.length) logRow_(r.row, 'stage', said.join(' · '));
  return { row: r.row, id: id };
}

function writeCell_(row, id, col, value) {
  const sheet = ss_().getSheetByName('Planner');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    let r = row;
    if (String(sheet.getRange(r, 1).getValue()) !== String(id)) {
      const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(v => String(v[0]));
      const i = ids.indexOf(String(id));
      if (i === -1) throw new Error('Could not find ' + id + ' in the Planner.');
      r = i + 2;
    }
    sheet.getRange(r, col).setValue(value);
    if (col === 11) sheet.getRange(r, 31).setValue(new Date());   // Stage since
    return { row: r, id: id, value: value };
  } finally {
    lock.releaseLock();
  }
}

// ===== Settings tab: Basecamp connections and sync status =====
function readProp_(k) {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(k) || 'null'); } catch (e) { return null; }
}

// A fresh Basecamp sign-in link for whoever clicks Connect. Links expire, so never reuse one baked into the page.
function getConnectUrl() {
  return connectUrlFor_(userKey_(String(currentEmail_()).toLowerCase()));
}

function syncStatus_() {
  const reg = registry_();
  const last = readProp_('SYNC_LAST');
  const lastUsers = {};
  ((last && last.users) || []).forEach(u => { lastUsers[u.key] = u; });
  const meEmail = String(currentEmail_()).toLowerCase();
  const meKey = userKey_(meEmail);
  const nameFromEmail = e => String(e || '').split('@')[0].replace(/[._-]+/g, ' ');

  const conns = Object.keys(reg).map(k => {
    const r = reg[k];
    const access = serviceFor_(k).hasAccess();
    const lu = lastUsers[k];
    return {
      key: k, email: r.email || '', bcName: r.name || r.email || '', connectedAt: r.connectedAt || '', backup: !!r.backup,
      member: matchTeamName_(r.name) || matchTeamName_(nameFromEmail(r.email)) || '',
      status: !access ? 'reconnect' : (lu && !lu.ok ? 'error' : 'connected'),
      error: lu ? lu.error : '', found: lu ? lu.found : null, added: lu ? lu.added : null,
      me: k === meKey
    };
  });

  const mine = conns.find(c => c.me);
  let connectUrl = '';
  if (!mine || mine.status === 'reconnect') {
    connectUrl = 'on-click';   // made fresh when the button is clicked (getConnectUrl)
  }
  let legacy = false;
  try { legacy = serviceFor_('legacy').hasAccess(); } catch (e) {}

  return {
    conns: conns, last: last, log: readProp_('SYNC_LOG') || [], owner: readProp_('SYNC_OWNER'),
    hours: SYNC_HOURS, connectUrl: connectUrl, redirectUrl: (function () { try { return redirectUrl_(); } catch (e) { return ''; } })(), legacy: legacy,
    poll: readProp_('POLL_LAST'), backupLast: readProp_('BACKUP_LAST'),
    digest: { to: (function () { try { return digestRecipients_(); } catch (e) { return []; } })(), last: readProp_('DIGEST_LAST'), hour: DIGEST_HOUR, fixed: !!PropertiesService.getScriptProperties().getProperty('DIGEST_TO'), project: digestProject_() },
    me: { email: meEmail, connected: !!mine && mine.status !== 'reconnect', name: displayName_() }
  };
}

// Called from the Settings tab
function refreshSyncStatus() { return syncStatus_(); }

function runSyncFromDashboard() {
  const full = displayName_();
  const run = runSync_(matchTeamName_(full) || full);
  return { run: run, sync: syncStatus_() };
}

function disconnectMine() {
  const key = userKey_(currentEmail_());
  serviceFor_(key).reset();
  const reg = registry_(); delete reg[key]; saveRegistry_(reg);
  return syncStatus_();
}

// Moves the twice-daily sync (and the paste-a-link handler) to the person clicking
function takeOverSync() {
  setup();
  return syncStatus_();
}

// Settings tab: email the Monday digest to just the person clicking
function postDigestPreviewFromDashboard() {
  const info = postDigestPreview();
  return { info: info, sync: syncStatus_() };
}
function sendDigestPreviewFromDashboard() {
  const info = sendDigestPreview();
  return { info: info, sync: syncStatus_() };
}
