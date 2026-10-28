/**
 * Insights Planner | 1000heads Insights US – Basecamp link sync
 *
 * What it does
 *  - Basecamp > Add deliverable opens a short form (Basecamp link, account,
 *    type, scale, owner, dates) and adds the row, filling the rest from Basecamp.
 *  - Or paste a Basecamp to-do, sub-task or card link into column R (Basecamp
 *    link) on an empty row of the Planner tab → the script fills ID, Account,
 *    Basecamp project, Deliverable, Owner, Assigned to, Stage, Start date,
 *    Due date, Priority, Notes, Source and Added on.
 *  - Duplicate links are flagged instead of logged twice.
 *  - Team sync: every teammate connects their own Basecamp once (Basecamp >
 *    Connect my Basecamp). Twice a day the script pulls each connected
 *    person's open assignments and adds any to-do or card the planner doesn't
 *    have yet, so nothing depends on one person being tagged.
 *  - The same run re-checks every open row that has a Basecamp link: updates
 *    dates, assignees and the brief, and marks completed items Delivered.
 *  - "Basecamp" menu: Open dashboard, Add deliverable, Sync now, Connect my
 *    Basecamp, Who's connected, Disconnect my Basecamp.
 *
 * Setup: see the steps sent with this file.
 */

// ===== Settings =====
const BC_ACCOUNT_ID = '4104244';
const PLANNER_VERSION = '2026.10.29';                // bump on every code release; matches the folder name in Drive > Insights planner code
const CONTACT = 'tom.handy@1000heads.com';               // Basecamp requires a contact in the User-Agent
const SHEET_NAME = 'Planner';
const FIRST_ROW = 2;
const DEFAULT_PRIORITY = 'Medium';
const SOURCE_LABEL = 'Basecamp link';
const NEW_STAGE = 'Not started';
const DONE_STAGE = 'Delivered';
const SKIP_STAGES = ['Delivered', 'Dismissed'];   // never re-checked or re-added
const STALE_DAYS = 30;                            // to-dos due longer ago than this are ignored, and open rows past it expire
const SYNC_SOURCE = 'Basecamp sync';
const SYNC_HOURS = [7, 13];                       // daily sync times (script time zone)
const SYNC_LOG_MAX = 10;                          // recent runs kept for the Settings tab
const CHECKIN_HOUR = 10;                         // Wednesday team check-in in Campfire, around 10am
const WRAP_HOUR = 17;                            // Friday wrap-up in Campfire, around 5pm (script time zone)
const DIGEST_HOUR = 8;                            // Monday digest goes out around 8am (script time zone)

// Planner columns (1-based). H, N, O, P, V and W are formulas – never written.
const COL = {
  id: 1, client: 2, project: 3, deliverable: 4, type: 5, scale: 6, override: 7,
  owner: 9, assigned: 10, stage: 11, start: 12, due: 13, priority: 17, link: 18,
  notes: 19, source: 20, added: 21, brief: 24, booking: 25, replies: 28, actual: 29, likelihood: 30, stageSince: 31, trackerLogged: 32, clientDate: 33
};

// Sets a stage and records when it changed (for "days in stage" on the Board)
function setStage_(sheet, r, stage) {
  sheet.getRange(r, COL.stage).setValue(stage);
  sheet.getRange(r, COL.stageSince).setValue(new Date());
}
const BRIEF_MAX = 20000; // characters kept from each to-do description
const GONE_NOTE = "Basecamp can't show this link (deleted, moved or private)";
const REPLIES_KEPT = 10;  // latest Basecamp replies kept per to-do
const REPLY_MAX = 1500;   // characters kept from each reply
// A file or link whose name has EXT as its own word (e.g. "Chrome_Sept_Report_EXT.pdf") means it went to the client
const EXT_RX = /(^|[^A-Za-z0-9])EXT([^A-Za-z0-9]|$)/i;

// ===== Menu =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Basecamp')
    .addItem('Open dashboard', 'openDashboard')
    .addItem('Add deliverable', 'showAddForm')
    .addItem('Sync now', 'syncNow')
    .addSeparator()
    .addItem('Connect my Basecamp', 'connectBasecamp')
    .addItem("Who's connected", 'showConnections')
    .addItem('Disconnect my Basecamp', 'disconnectBasecamp')
    .addToUi();
}

// ===== OAuth: one Basecamp connection per teammate (uses the OAuth2 library) =====
// Tokens are kept in script properties so the timed sync can use everyone's connection.
function serviceFor_(userKey) {
  const props = PropertiesService.getScriptProperties();
  return OAuth2.createService(userKey === 'legacy' ? 'basecamp' : 'bc_' + userKey)
    .setAuthorizationBaseUrl('https://launchpad.37signals.com/authorization/new')
    .setTokenUrl('https://launchpad.37signals.com/authorization/token')
    .setClientId(props.getProperty('BC_CLIENT_ID'))
    .setClientSecret(props.getProperty('BC_CLIENT_SECRET'))
    .setCallbackFunction('authCallback')
    .setPropertyStore(props)
    .setCache(CacheService.getScriptCache())
    .setLock(LockService.getScriptLock())
    .setParam('type', 'web_server')
    .setTokenPayloadHandler(function (payload) {
      payload.type = payload.grant_type === 'refresh_token' ? 'refresh' : 'web_server';
      return payload;
    });
}

// ===== Connecting: Basecamp sends people back to the dashboard web app (not Apps Script's usercallback) =====
// The dashboard link already works for everyone, so finishing the sign-in there avoids Google's
// "Authorization is required" and "state token" errors. The Basecamp integration's Redirect URI must match redirectUrl_().
function redirectUrl_() {
  const p = PropertiesService.getScriptProperties().getProperty('BC_REDIRECT_URL');
  if (p) return p;
  return ScriptApp.getService().getUrl();
}
function connectUrlFor_(key) {
  const nonce = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  CacheService.getScriptCache().put('bcstate_' + nonce, key, 1800);   // link works for 30 minutes
  const id = PropertiesService.getScriptProperties().getProperty('BC_CLIENT_ID');
  return 'https://launchpad.37signals.com/authorization/new?type=web_server&client_id=' + encodeURIComponent(id) +
    '&redirect_uri=' + encodeURIComponent(redirectUrl_()) + '&state=' + nonce;
}
function connectPage_(title, body) {
  return HtmlService.createHtmlOutput('<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5;padding:24px;max-width:520px">' +
    '<p style="font-size:16px"><b>' + title + '</b></p><p>' + body + '</p></div>').setTitle('Connect my Basecamp');
}
// Called from doGet when Basecamp sends someone back with ?code=…&state=…
function finishConnect_(e) {
  const prm = e.parameter || {};
  const cache = CacheService.getScriptCache();
  const key = cache.get('bcstate_' + prm.state);
  if (!key) { logActivity_('settings', '', 'Basecamp connection', currentEmail_(), 'Tried to connect Basecamp, but the sign-in link had expired (over 30 minutes old).', ''); return connectPage_('That sign-in link has expired.', 'Close this tab, go back to the dashboard and click Connect my Basecamp again.'); }
  cache.remove('bcstate_' + prm.state);
  if (prm.error) { logActivity_('settings', '', 'Basecamp connection', currentEmail_(), 'Basecamp access was not approved (' + prm.error + ').', ''); return connectPage_('Basecamp access was not approved.', 'Close this tab and try again if that was a mistake.'); }
  const service = serviceFor_(key);
  try {
    if (typeof service.setRedirectUri !== 'function') throw new Error('Update the OAuth2 library to the latest version (Libraries > OAuth2 > Version).');
    service.setRedirectUri(redirectUrl_());
    service.handleCallback({ parameter: { code: prm.code } });
  } catch (err) {
    logActivity_('settings', '', 'Basecamp connection', currentEmail_(), 'Could not finish connecting Basecamp: ' + (err.message || err), '');
    return connectPage_('Could not finish connecting.', String(err.message || err) + '<br><br>Send this message to whoever looks after the planner.');
  }
  const who = recordConnection_(key, service);
  // Pin the address Basecamp just accepted, so every later sign-in sends exactly the same one
  try { const props = PropertiesService.getScriptProperties(); if (!props.getProperty('BC_REDIRECT_URL')) props.setProperty('BC_REDIRECT_URL', redirectUrl_()); } catch (err) {}
  try { ensureBackupTrigger_(); } catch (err) { console.log('Backup schedule not added: ' + err.message); }
  return connectPage_('Connected as ' + who + '.', 'Close this tab. On the dashboard, open Settings and click Refresh. Your to-dos come in with the next sync.');
}
function recordConnection_(key, service) {
  let name = '', bcEmail = '';
  const email = currentEmail_();
  try {
    const res = UrlFetchApp.fetch('https://launchpad.37signals.com/authorization.json', {
      headers: { Authorization: 'Bearer ' + service.getAccessToken(), 'User-Agent': userAgent_() }, muteHttpExceptions: true });
    const id = JSON.parse(res.getContentText()).identity || {};
    name = [id.first_name, id.last_name].filter(String).join(' ');
    bcEmail = id.email_address || '';
  } catch (e) { /* identity is optional */ }
  if (key !== 'legacy') {
    const reg = registry_();
    reg[key] = { email: email, bcEmail: bcEmail, name: name || email, connectedAt: new Date().toISOString() };
    saveRegistry_(reg);
  }
  logActivity_('settings', '', 'Basecamp connection', name || email, 'Connected their Basecamp' + (bcEmail ? ' (' + bcEmail + ')' : '') + ' from ' + email, '');
  return name || email;
}

function currentEmail_() {
  return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'unknown';
}
// Short key per Google account. Basecamp rejects long OAuth "state" values, so keep service names tiny.
function userKey_(email) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(email).toLowerCase());
  return bytes.slice(0, 5).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

function registry_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('BC_USERS') || '{}'); }
  catch (e) { return {}; }
}
function saveRegistry_(reg) {
  PropertiesService.getScriptProperties().setProperty('BC_USERS', JSON.stringify(reg));
}

// Every working connection, plus the original single connection if it is still there
function services_() {
  const out = [];
  const reg = registry_();
  Object.keys(reg).forEach(k => {
    const svc = serviceFor_(k);
    if (svc.hasAccess()) out.push({ key: k, name: reg[k].name || reg[k].email, service: svc });
  });
  const legacy = serviceFor_('legacy');
  if (legacy.hasAccess()) out.push({ key: 'legacy', name: 'Original connection', service: legacy });
  return out;
}

// Kept for older code paths: the first working connection
function getService_() {
  const all = services_();
  if (!all.length) throw new Error('No Basecamp connections yet – use Basecamp > Connect my Basecamp');
  return all[0].service;
}

function connectBasecamp() {
  const email = currentEmail_();
  const key = userKey_(email);
  const service = serviceFor_(key);
  const ui = SpreadsheetApp.getUi();
  if (service.hasAccess()) {
    ui.alert('Your Basecamp is already connected (' + email + ').');
    return;
  }
  const url = connectUrlFor_(key);
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:14px;line-height:1.5">' +
    '<p>Sign in to Basecamp as <b>yourself</b> and approve access. The planner will then pick up every to-do and card assigned to you.</p>' +
    '<p><a href="' + url + '" target="_blank" style="display:inline-block;background:#6100DB;color:#fff;padding:10px 16px;border-radius:12px;text-decoration:none;font-weight:bold">Connect my Basecamp</a></p>' +
    '<p style="color:#6b7280">Close this window when Basecamp says you are connected.</p></div>'
  ).setWidth(380).setHeight(220);
  ui.showModalDialog(html, 'Connect my Basecamp');
}

function authCallback(request) {
  // The OAuth2 library passes the service name back to us ("bc_<key>", or "basecamp" for the original connection)
  const sname = String(request.parameter.serviceName || '');
  const key = sname === 'basecamp' ? 'legacy' : sname.replace(/^bc_/, '');
  if (!key) return HtmlService.createHtmlOutput('Something went wrong. Close this tab and try Connect my Basecamp again.');
  const service = serviceFor_(key);
  const ok = service.handleCallback(request);
  if (!ok) return HtmlService.createHtmlOutput('Access denied. Close this tab and try again.');
  const who = recordConnection_(key, service);
  return HtmlService.createHtmlOutput('<div style="font-family:Arial;font-size:14px;line-height:1.5;padding:16px"><p><b>Connected as ' + who + '.</b></p>' +
    '<p>Close this tab. On the dashboard, open Settings and click Refresh to see your status. Your to-dos come in with the next sync.</p></div>');
}

function disconnectBasecamp() {
  const key = userKey_(currentEmail_());
  serviceFor_(key).reset();
  const reg = registry_(); delete reg[key]; saveRegistry_(reg);
  SpreadsheetApp.getUi().alert('Your Basecamp connection was removed.');
}

function showConnections() {
  const list = services_().map(s => '• ' + s.name).join('\n');
  SpreadsheetApp.getUi().alert("Basecamp connections", list || 'Nobody is connected yet.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function userAgent_() { return '1000heads Insights Planner (' + CONTACT + ')'; }

// ===== One-time setup: triggers =====
// Whoever runs setup() becomes the sync owner. Anyone else's old triggers then stand down.
function setup() {
  const ss = ss_();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss).onEdit().create();
  SYNC_HOURS.forEach(h => ScriptApp.newTrigger('dailySync').timeBased().everyDays(1).atHour(h).create());
  ScriptApp.newTrigger('pollReplies').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('weeklyDigest').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(DIGEST_HOUR).create();
  ScriptApp.newTrigger('wednesdayCheckin').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(CHECKIN_HOUR).create();
  if (registry_()[userKey_(String(currentEmail_()).toLowerCase())]) ScriptApp.newTrigger('backupRun').timeBased().everyHours(BACKUP_EVERY_HOURS).create();
  ScriptApp.newTrigger('fridayWrap').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(WRAP_HOUR).nearMinute(0).create();
  const email = currentEmail_();
  const full = typeof displayName_ === 'function' ? displayName_() : email;
  const name = matchTeamName_(full) || full;   // "Tom Handy" → "Tom H." when on the Team tab
  PropertiesService.getScriptProperties().setProperty('SYNC_OWNER',
    JSON.stringify({ email: email, name: name, at: new Date().toISOString() }));
  return { email: email, name: name };
}

// True when this account owns the triggers (or no owner is recorded yet)
function isSyncOwner_() {
  let owner = null;
  try { owner = JSON.parse(PropertiesService.getScriptProperties().getProperty('SYNC_OWNER') || 'null'); } catch (e) {}
  if (!owner || !owner.email) return true;
  const me = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  return !me || me === String(owner.email).toLowerCase();
}

// ===== Paste handler =====
function handleEdit(e) {
  if (!isSyncOwner_()) return;   // an older owner's trigger: the current owner's trigger handles this edit
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;
  const c1 = range.getColumn(), c2 = range.getLastColumn();
  // Stage typed in the Sheet: note when it changed
  if (COL.stage >= c1 && COL.stage <= c2) {
    for (let r = Math.max(range.getRow(), FIRST_ROW); r <= range.getLastRow(); r++) sheet.getRange(r, COL.stageSince).setValue(new Date());
  }
  if (COL.link < c1 || COL.link > c2) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    for (let r = range.getRow(); r <= range.getLastRow(); r++) {
      if (r < FIRST_ROW) continue;
      const url = String(sheet.getRange(r, COL.link).getValue()).trim();
      if (!url) continue;
      if (sheet.getRange(r, COL.deliverable).getValue() !== '') continue; // existing row, leave it
      logFromLink_(sheet, r, url);
    }
  } finally {
    lock.releaseLock();
  }
}

function logFromLink_(sheet, row, url) {
  const ref = parseLink_(url);
  const notes = sheet.getRange(row, COL.notes);
  if (!ref) { notes.setValue('Not a Basecamp to-do, sub-task or card link'); return; }

  const dupRow = findDuplicate_(sheet, ref.key, row);
  if (dupRow) { notes.setValue('Duplicate: already logged in row ' + dupRow); return; }

  let item;
  try { item = fetchItem_(ref); }
  catch (err) { notes.setValue('Could not read from Basecamp: ' + err.message); return; }

  const names = (item.assignees || []).map(a => a.name);
  const team = mapTeam_(names);
  const desc = stripHtml_(item.description || item.content_description || '');
  const noteParts = [];
  if (item.parent && item.parent.title) noteParts.push('List: ' + item.parent.title);
  if (desc) noteParts.push(desc.slice(0, 200));

  const values = {};
  values[COL.id] = nextId_(sheet);
  values[COL.client] = clientFor_(item.bucket ? item.bucket.name : '');
  values[COL.project] = item.bucket ? item.bucket.name : '';
  values[COL.deliverable] = item.title || item.content || '';
  values[COL.type] = inferType_(values[COL.deliverable]);
  values[COL.owner] = pickOwner_(team, item.due_on);
  values[COL.assigned] = assignedText_(names);
  values[COL.stage] = item.completed ? DONE_STAGE : NEW_STAGE;
  values[COL.start] = toDate_(item.starts_on);
  values[COL.due] = toDate_(item.due_on);
  values[COL.priority] = DEFAULT_PRIORITY;
  values[COL.source] = SOURCE_LABEL;
  values[COL.link] = item.app_url || url;
  values[COL.notes] = noteParts.join(' · ');
  values[COL.added] = new Date();
  values[COL.brief] = briefOf_(item);

  Object.keys(values).forEach(c => sheet.getRange(row, Number(c)).setValue(values[c]));
}

// ===== Add deliverable form =====
function showAddForm() {
  const ss = ss_();
  const list = (sheet, a1) => ss.getSheetByName(sheet).getRange(a1).getValues()
    .map(r => String(r[0]).trim()).filter(String);
  const opts = arr => '<option value=""></option>' +
    arr.map(v => '<option>' + v.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</option>').join('');
  const html = `
<style>
  body{font-family:Arial,sans-serif;font-size:13px;margin:0;padding:12px;color:#1f2937}
  label{display:block;font-weight:bold;margin:10px 0 4px}
  input,select{width:100%;box-sizing:border-box;padding:6px;border:1px solid #cbd5e1;border-radius:4px;font-size:13px}
  .row{display:flex;gap:10px}.row>div{flex:1}
  .hint{color:#6b7280;font-size:11px;margin-top:3px}
  button{margin-top:16px;background:#6100DB;color:#fff;border:0;padding:9px 16px;border-radius:12px;font-size:13px;cursor:pointer}
  #msg{margin-top:10px;color:#065f46}
</style>
<label>Basecamp link</label>
<input id="link" placeholder="https://3.basecamp.com/4104244/buckets/.../todos/...">
<div class="hint">Optional. Title, assignees, project and dates are pulled from Basecamp.</div>
<label>Deliverable</label>
<input id="deliverable" placeholder="Leave blank to use the Basecamp title">
<div class="row">
  <div><label>Account</label><select id="account">${opts(list('Lists', 'A2:A30'))}</select></div>
  <div><label>Owner</label><select id="owner">${opts(list('Team', 'A5:A14'))}</select></div>
</div>
<div class="row">
  <div><label>Type</label><select id="type">${opts(list('Lists', 'D2:D30'))}</select></div>
  <div><label>Scale</label><select id="scale">${opts(list('Lists', 'H2:H4'))}</select></div>
</div>
<div class="row">
  <div><label>Start date</label><input id="start" type="date"></div>
  <div><label>Due date</label><input id="due" type="date"></div>
</div>
<div class="hint">Blank fields are filled from Basecamp where possible.</div>
<button id="go" onclick="submitForm()">Add to planner</button>
<div id="msg"></div>
<script>
  function submitForm(){
    const f = {};
    ['link','deliverable','account','owner','type','scale','start','due'].forEach(k => f[k] = document.getElementById(k).value.trim());
    if (!f.link && !f.deliverable) { document.getElementById('msg').textContent = 'Add a Basecamp link or a deliverable name.'; return; }
    document.getElementById('go').disabled = true;
    document.getElementById('msg').textContent = 'Adding…';
    google.script.run
      .withSuccessHandler(m => { document.getElementById('msg').textContent = m; setTimeout(() => google.script.host.close(), 1800); })
      .withFailureHandler(e => { document.getElementById('msg').textContent = e.message; document.getElementById('go').disabled = false; })
      .addDeliverable(f);
  }
</script>`;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(460).setHeight(560), 'Add deliverable');
}

function addDeliverable(f) {
  const sheet = ss_().getSheetByName(SHEET_NAME);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let item = null, ref = null;
    if (f.link) {
      ref = parseLink_(f.link);
      if (!ref) throw new Error('That is not a Basecamp to-do, sub-task or card link.');
      const dup = findDuplicate_(sheet, ref.key, -1);
      if (dup) throw new Error('Already in the planner (row ' + dup + ').');
      item = fetchItem_(ref);
    }
    // First row with no deliverable
    const last = sheet.getLastRow();
    const names = sheet.getRange(FIRST_ROW, COL.deliverable, Math.max(last - FIRST_ROW + 1, 1), 1).getValues();
    let row = names.findIndex(r => r[0] === '');
    row = row === -1 ? last + 1 : FIRST_ROW + row;

    const assignees = item ? (item.assignees || []).map(a => a.name) : [];
    const team = mapTeam_(assignees);
    const notes = [];
    if (item && item.parent && item.parent.title) notes.push('List: ' + item.parent.title);
    const desc = item ? stripHtml_(item.description || '') : '';
    if (desc) notes.push(desc.slice(0, 200));

    const v = {};
    v[COL.id] = nextId_(sheet);
    v[COL.client] = f.account || (item && item.bucket ? clientFor_(item.bucket.name) : '');
    v[COL.project] = item && item.bucket ? item.bucket.name : '';
    v[COL.deliverable] = f.deliverable || (item ? (item.title || item.content || '') : '');
    v[COL.type] = f.type || inferType_(v[COL.deliverable]);
    v[COL.scale] = f.scale;
    v[COL.owner] = f.owner || pickOwner_(team, item && item.due_on);
    v[COL.assigned] = assignedText_(assignees);
    v[COL.stage] = item && item.completed ? DONE_STAGE : NEW_STAGE;
    v[COL.start] = f.start || toDate_(item && item.starts_on);
    v[COL.due] = f.due || toDate_(item && item.due_on);
    v[COL.priority] = DEFAULT_PRIORITY;
    v[COL.link] = item ? (item.app_url || f.link) : '';
    v[COL.notes] = notes.join(' · ');
    v[COL.source] = item ? SOURCE_LABEL : 'Manual';
    v[COL.added] = new Date();
    v[COL.brief] = item ? briefOf_(item) : '';
    Object.keys(v).forEach(c => sheet.getRange(row, Number(c)).setValue(v[c]));
    return 'Added ' + v[COL.id] + ': ' + v[COL.deliverable];
  } finally {
    lock.releaseLock();
  }
}

// ===== Team sync =====
function syncNow() { dailySync(); }

// Timed runs pass an event; menu and dashboard runs don't
function dailySync(e) {
  const scheduled = !!(e && e.triggerUid);
  if (scheduled && !isSyncOwner_()) { console.log('Skipped: another account now runs the daily sync'); return null; }
  return runSync_(scheduled ? 'Scheduled' : (typeof displayName_ === 'function' ? displayName_() : currentEmail_()));
}

// Runs the full sync and records the result for the dashboard's Settings tab
function runSync_(by) {
  const started = Date.now();
  let prev = null;
  try { prev = JSON.parse(PropertiesService.getScriptProperties().getProperty('SYNC_LAST') || 'null'); } catch (e) {}
  const run = { at: new Date().toISOString(), by: by, ok: true, added: 0, checked: 0, completed: 0, updated: 0, failed: 0, users: [], error: '' };
  try {
    const team = syncTeamAssignments_();
    run.added = team.added.length;
    run.users = team.users;
    const r = refreshLinkedRows(team.added) || {};
    try { const rec = generateRecurring_(); run.pencilled = rec.added.length; run.retired = rec.retired.length; } catch (e) { console.log('Recurring failed: ' + e.message); }
    try { run.reclassified = reclassifyClients_(); } catch (e) { console.log('Client match failed: ' + e.message); }
    try { run.expired = expireOld_(); } catch (e) { console.log('Expiry failed: ' + e.message); }
    try { run.logged = logToTracker_(); } catch (e) { console.log('Tracker log failed: ' + e.message); run.trackerError = e.message; }
    try { writeTeamCapacity_(); } catch (e) { console.log('Team tab capacity failed: ' + e.message); }
    run.checked = r.checked || 0; run.completed = r.completed || 0; run.updated = r.updated || 0; run.failed = r.failed || 0; run.gone = r.gone || 0;
    if (team.users.some(u => !u.ok) || run.failed) run.ok = false;
  } catch (err) {
    run.ok = false;
    run.error = /lock/i.test(err.message) ? 'Another sync was already running. Try again in a minute.' : err.message;
  }
  run.seconds = Math.round((Date.now() - started) / 1000);
  saveRun_(run);
  // Email only when things go from fine to broken, so a lasting problem doesn't send an alert twice a day
  if (!run.ok && (!prev || prev.ok)) { try { sendSyncAlert_(run); } catch (e) { console.log('Alert failed: ' + e.message); } }
  return run;
}

function saveRun_(run) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SYNC_LAST', JSON.stringify(run));
  let log = [];
  try { log = JSON.parse(props.getProperty('SYNC_LOG') || '[]'); } catch (e) {}
  log.unshift({ at: run.at, by: run.by, ok: run.ok, added: run.added, completed: run.completed, updated: run.updated, failed: run.failed, error: run.error,
                problems: run.users.filter(u => !u.ok).map(u => u.name) });
  props.setProperty('SYNC_LOG', JSON.stringify(log.slice(0, SYNC_LOG_MAX)));
}

// Pulls every connected teammate's open Basecamp assignments and adds any the planner is missing
function syncTeamAssignments_() {
  const sheet = ss_().getSheetByName(SHEET_NAME);
  const last = Math.max(sheet.getLastRow(), FIRST_ROW);
  const have = {};
  sheet.getRange(FIRST_ROW, COL.link, last - FIRST_ROW + 1, 1).getValues().forEach(v => {
    const ref = parseLink_(String(v[0])); if (ref) have[ref.key] = true;
  });
  const cutoff = Utilities.formatDate(new Date(Date.now() - STALE_DAYS * 864e5), ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const added = [];
  const users = [];
  // Teammates whose Basecamp access has expired or been revoked
  const reg = registry_();
  Object.keys(reg).forEach(k => {
    if (!serviceFor_(k).hasAccess()) users.push({ key: k, name: reg[k].name || reg[k].email, ok: false, found: 0, added: 0, error: 'Needs to reconnect' });
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    services_().forEach(conn => {
      const u = { key: conn.key, name: conn.name, ok: true, found: 0, added: 0, error: '' };
      users.push(u);
      let list = [];
      try {
        const data = apiGet_(conn.service, 'https://3.basecampapi.com/' + BC_ACCOUNT_ID + '/my/assignments.json');
        list = (data.priorities || []).concat(data.non_priorities || []);
      } catch (e) {
        console.log('Assignments failed for ' + conn.name + ': ' + e.message);
        u.ok = false;
        u.error = /HTTP 401/.test(e.message) ? 'Basecamp rejected the connection. Reconnect.' : e.message;
        return;
      }
      u.found = list.filter(a => !a.completed).length;
      list.forEach(a => {
        const ref = parseLink_(String(a.app_url || ''));
        if (!ref || have[ref.key] || a.completed) return;
        if (a.due_on && a.due_on < cutoff) return;           // long overdue: assume stale
        let item;
        try { item = fetchItem_(ref, conn.service); } catch (e) { return; }
        Utilities.sleep(250);
        if (item.completed) return;
        const row = firstEmptyRow_(sheet);
        writeNewRow_(sheet, row, item, item.app_url || a.app_url, SYNC_SOURCE + ' (' + conn.name + ')');
        logActivity_('new', sheet.getRange(row, COL.id).getValue(), item.title || item.content || '', conn.name, 'New to-do from ' + conn.name + '\'s Basecamp' + (item.due_on ? ', due ' + item.due_on : ''), item.app_url || a.app_url);
        have[ref.key] = true;
        u.added++;
        added.push(sheet.getRange(row, COL.id).getValue());
      });
    });
  } finally {
    lock.releaseLock();
  }
  console.log('Team sync added ' + added.length + ': ' + added.join(', '));
  return { added: added, users: users };
}

function firstEmptyRow_(sheet) {
  const last = sheet.getLastRow();
  const names = sheet.getRange(FIRST_ROW, COL.deliverable, Math.max(last - FIRST_ROW + 1, 1), 1).getValues();
  const i = names.findIndex(r => r[0] === '');
  return i === -1 ? last + 1 : FIRST_ROW + i;
}

function writeNewRow_(sheet, row, item, link, source) {
  const names = (item.assignees || []).map(a => a.name);
  const v = {};
  v[COL.id] = nextId_(sheet);
  v[COL.client] = clientFor_(item.bucket ? item.bucket.name : '');
  v[COL.project] = item.bucket ? item.bucket.name : '';
  v[COL.deliverable] = item.title || item.content || '';
  v[COL.type] = inferType_(v[COL.deliverable]);
  v[COL.owner] = pickOwner_(mapTeam_(names), item.due_on);
  v[COL.assigned] = assignedText_(names);
  v[COL.stage] = NEW_STAGE;
  v[COL.start] = toDate_(item.starts_on);
  v[COL.due] = toDate_(item.due_on);
  v[COL.priority] = DEFAULT_PRIORITY;
  v[COL.link] = link;
  v[COL.notes] = item.parent && item.parent.title ? 'List: ' + item.parent.title : '';
  v[COL.source] = source;
  v[COL.added] = new Date();
  v[COL.brief] = briefOf_(item);
  Object.keys(v).forEach(c => sheet.getRange(row, Number(c)).setValue(v[c]));
}

// Owner = first assigned teammate (Team tab order) who is in capacity and not out on the due date
function pickOwner_(team, dueIso) {
  if (!team.length) return '';
  const t = ss_().getSheetByName('Team');
  const people = t.getRange('A5:P14').getValues();
  // Full-day time off from the Time off tab (Name, Start, End), else the old Team tab table
  const tos = ss_().getSheetByName('Time off');
  const off = tos && tos.getLastRow() > 1
    ? tos.getRange(2, 2, tos.getLastRow() - 1, 5).getValues().filter(r => String(r[4]).trim().indexOf('Half day') !== 0).map(r => [r[0], r[1], r[2]])
    : t.getRange('A21:C45').getValues();
  const tz = ss_().getSpreadsheetTimeZone();
  const due = dueIso || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const iso = d => d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : '';
  const ok = name => {
    const p = people.find(r => String(r[0]).trim() === name);
    if (p && String(p[5]) !== 'Yes') return false;
    if (p && p[4] instanceof Date && iso(p[4]) < due) return false;
    return !off.some(o => String(o[0]).trim() === name && iso(o[1]) <= due && iso(o[2] || o[1]) >= due);
  };
  // Hands-on people lead; someone in Oversight (the Director) only owns it when nobody hands-on is on it
  const order = people.map(r => String(r[0]).trim());
  const oversight = name => { const p = people.find(r => String(r[0]).trim() === name); return !!p && String(p[15]).trim() === 'Oversight'; };
  const sorted = team.slice().sort((a, b) => (oversight(a) - oversight(b)) || (order.indexOf(a) - order.indexOf(b)));
  return sorted.find(ok) || sorted[0];
}

// ===== Daily refresh =====
function refreshLinkedRows(newIds) {
  const ss = ss_();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const last = sheet.getLastRow();
  if (last < FIRST_ROW) return { checked: 0, completed: 0, updated: 0, failed: 0 };
  const tz = ss.getSpreadsheetTimeZone();
  const data = sheet.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, COL.replies).getValues();
  const iso = v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : '';
  let checked = 0, completed = [], redated = [], reassigned = [], briefed = [], failed = [], typed = [], replied = [], extDelivered = [], gone = [];

  data.forEach((row, i) => {
    const stage = row[COL.stage - 1];
    const ref = parseLink_(String(row[COL.link - 1]));
    if (!ref || SKIP_STAGES.indexOf(stage) !== -1 || !row[COL.deliverable - 1]) return;
    const id = row[COL.id - 1];
    let item;
    try { item = fetchItem_(ref); }
    catch (err) {
      // Basecamp no longer shows it (deleted, moved, or in a project nobody connected can see): note it once, don't count it as a sync failure
      if (/404|403/.test(err.message)) {
        gone.push(id);
        const r0 = FIRST_ROW + i, prevNote = String(row[COL.notes - 1] || '');
        if (prevNote.indexOf(GONE_NOTE) === -1) sheet.getRange(r0, COL.notes).setValue(prevNote ? prevNote + ' · ' + GONE_NOTE : GONE_NOTE);
      } else failed.push(id + ' (' + err.message + ')');
      return;
    }
    Utilities.sleep(250); // stay under Basecamp's rate limit
    checked++;
    const r = FIRST_ROW + i;

    if (!row[COL.type - 1]) {
      const guess = inferType_(row[COL.deliverable - 1]);
      if (guess) { sheet.getRange(r, COL.type).setValue(guess); typed.push(id); }
    }

    let dateChanged = false;
    if (item.due_on && item.due_on !== iso(row[COL.due - 1])) {
      sheet.getRange(r, COL.due).setValue(toDate_(item.due_on)); dateChanged = true;
      const was = iso(row[COL.due - 1]);
      logActivity_('timeline', id, row[COL.deliverable - 1], 'Basecamp', 'Due date ' + (was ? 'changed: ' + was + ' → ' + item.due_on : 'set: ' + item.due_on), row[COL.link - 1]);
      if (was) {
        let st = null; try { st = JSON.parse(String(row[COL.replies - 1] || 'null')); } catch (e) {}
        st = st || { n: 0, list: [] };
        st.flag = { kind: 'due', at: new Date().toISOString(), by: 'Basecamp', changes: ['Due date: ' + was + ' → ' + item.due_on] };
        sheet.getRange(r, COL.replies).setValue(JSON.stringify(st));
        row[COL.replies - 1] = JSON.stringify(st);
      }
    }
    if (item.starts_on && item.starts_on !== iso(row[COL.start - 1])) {
      sheet.getRange(r, COL.start).setValue(toDate_(item.starts_on)); dateChanged = true;
    }
    if (dateChanged) redated.push(id);

    const assigned = assignedText_((item.assignees || []).map(a => a.name));
    if (assigned && assigned !== row[COL.assigned - 1]) {
      sheet.getRange(r, COL.assigned).setValue(assigned);
      if (!row[COL.owner - 1]) {
        const owner = pickOwner_(mapTeam_((item.assignees || []).map(a => a.name)), item.due_on);
        if (owner) sheet.getRange(r, COL.owner).setValue(owner);
      }
      reassigned.push(id);
    }

    const brief = briefOf_(item);
    if (brief && brief !== String(row[COL.brief - 1])) {
      const due = row[COL.due - 1] instanceof Date ? row[COL.due - 1] : new Date();
      const was = parseWbs_(String(row[COL.brief - 1] || ''), due), now = parseWbs_(brief, due);
      sheet.getRange(r, COL.brief).setValue(brief);
      briefed.push(id);
      const changes = now ? (was ? diffWbs_(was, now) : (row[COL.brief - 1] ? ['Work back schedule added to the brief'] : [])) : [];
      if (changes.length) {
        logActivity_('timeline', id, row[COL.deliverable - 1], 'Basecamp', 'Brief\'s work back schedule changed: ' + changes.join('; '), row[COL.link - 1]);
        let st = null; try { st = JSON.parse(String(row[COL.replies - 1] || 'null')); } catch (e) {}
        st = st || { n: 0, list: [] };
        st.flag = { kind: 'wbs', at: new Date().toISOString(), by: 'Brief', changes: changes.slice(0, 8) };
        sheet.getRange(r, COL.replies).setValue(JSON.stringify(st));
        row[COL.replies - 1] = JSON.stringify(st);
      }
    }

    // Basecamp replies: re-read only when the count changed since last sync
    let stored = null;
    try { stored = JSON.parse(String(row[COL.replies - 1] || 'null')); } catch (e) {}
    const count = Number(item.comments_count || 0);
    if (count && (!stored || stored.n !== count)) {
      try {
        const list = fetchReplies_(item);
        const known = {}; ((stored && stored.list) || []).forEach(x => { known[x.url || x.at] = 1; });
        const fresh = list.filter(x => !known[x.url || x.at]);
        const newest = ((stored && stored.list) || []).reduce((m, x) => x.at > m ? x.at : m, '');
        const res = applyNewReplies_(sheet, r, row, stored, fresh, tz, count, !stored || !newest ? true : newest);
        replied.push(id);
        if (res.delivered) extDelivered.push(id);
      } catch (e) { failed.push(id + ' replies (' + e.message + ')'); }
    } else if (!count && stored) {
      sheet.getRange(r, COL.replies).setValue('');
    }

    if (item.completed) {
      setStage_(sheet, r, DONE_STAGE);
      completed.push(id);
      logActivity_('stage', id, row[COL.deliverable - 1], 'Basecamp sync', 'Completed in Basecamp, marked Delivered', row[COL.link - 1]);
    }
  });

  const list = a => a.length ? ' (' + a.join(', ') + ')' : '';
  if (!Array.isArray(newIds)) newIds = null;
  const summary = (newIds ? newIds.length + ' new from teammates\' Basecamp' + list(newIds) + '.\n' : '') +
    'Checked ' + checked + ' rows.\n' +
    completed.length + ' marked Delivered' + list(completed) + '.\n' +
    redated.length + ' dates updated' + list(redated) + '.\n' +
    reassigned.length + ' assignee lists updated' + list(reassigned) + '.\n' +
    briefed.length + ' briefs updated' + list(briefed) + '.\n' +
    typed.length + ' given a Type from their title' + list(typed) + '.\n' +
    replied.length + ' with new Basecamp replies' + list(replied) + '.\n' +
    extDelivered.length + ' marked Delivered from an EXT file' + list(extDelivered) + '.\n' +
    gone.length + ' links Basecamp no longer shows' + list(gone) + (gone.length ? '. Set them to Dismissed if they\'re done, or fix the link' : '') + '.\n' +
    failed.length + ' could not be read' + (failed.length ? ': ' + failed.join(', ') : '') + '.';
  console.log(summary);
  try { SpreadsheetApp.getUi().alert('Basecamp sync', summary, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { /* running on the daily timer or from the dashboard – no pop-up */ }
  const updated = {};
  redated.concat(reassigned, briefed).forEach(id => { updated[id] = 1; });
  return { checked: checked, completed: completed.length + extDelivered.length, updated: Object.keys(updated).length, failed: failed.length, gone: gone.length, replies: replied.length, ext: extDelivered };
}

// ===== Basecamp API =====
function parseLink_(url) {
  const m = url.match(/\/(\d+)\/buckets\/(\d+)\/(todos|subtasks|card_tables\/cards|cards)\/(\d+)/);
  if (!m) return null;
  // Sub-tasks can be to-dos or card steps, so both endpoints are tried
  const kinds = m[3] === 'todos' ? ['todos'] : m[3] === 'subtasks' ? ['todos', 'card_tables/steps'] : ['card_tables/cards'];
  return { account: m[1], bucket: m[2], kind: kinds[0], kinds: kinds, id: m[4], key: m[2] + '/' + m[4] };
}

// Reads a to-do or card, trying the preferred connection first and then everyone else's
function fetchItem_(ref, preferred) {
  const conns = services_().map(c => c.service);
  if (preferred) conns.unshift(preferred);
  if (!conns.length) throw new Error('No Basecamp connections yet – use Basecamp > Connect my Basecamp');
  let lastErr = null;
  for (const kind of (ref.kinds || [ref.kind])) {
    const api = 'https://3.basecampapi.com/' + (ref.account || BC_ACCOUNT_ID) +
                '/buckets/' + ref.bucket + '/' + kind + '/' + ref.id + '.json';
    for (const svc of conns) {
      try { return apiGet_(svc, api); } catch (e) { lastErr = e; if (!/404|403/.test(e.message)) throw e; }
    }
  }
  throw lastErr || new Error('not found');
}

function apiGet_(service, url) {
  const opts = {
    headers: { Authorization: 'Bearer ' + service.getAccessToken(), 'User-Agent': userAgent_() },
    muteHttpExceptions: true
  };
  let res = UrlFetchApp.fetch(url, opts);
  if (res.getResponseCode() === 429) {
    Utilities.sleep(Number(res.getHeaders()['Retry-After'] || 10) * 1000);
    res = UrlFetchApp.fetch(url, opts);
  }
  const code = res.getResponseCode();
  if (code !== 200) throw new Error('HTTP ' + code + (code === 404 ? ' (not found or no access)' : ''));
  return JSON.parse(res.getContentText());
}

// ===== Helpers =====
// Guesses a deliverable Type from its title (Lists column D names). First match wins.
const TYPE_RULES = [
  [/4\s*c'?s\b|four\s*cs?\b/i, '4Cs audit'],
  [/brand\s*lift|\b(?:pre|post)[\s-]*(?:lift\s*)?survey|survey\s*(?:launch|recap)/i, 'Brand lift study'],
  [/social\s*listening|listening\s*report/i, 'Social listening report'],
  [/talent\s*(?:valuation|roi)|creator\s*scor/i, 'Talent valuation'],
  [/\b24\s*-?\s*h(?:ou)?rs?\b/i, '24hr campaign report'],
  [/\b48\s*-?\s*h(?:ou)?rs?\b/i, '48hr campaign report'],
  [/\b72\s*-?\s*h(?:ou)?rs?\b/i, '72hr campaign report'],
  [/\b(?:1|one)[\s-]*week\b/i, '1-week campaign report'],
  [/\b(?:1|one)[\s-]*month\b.*(?:campaign|recap)|\b30[\s-]*day\b/i, '1-month campaign report'],
  [/\bmonthly\b|\b(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\s+20\d{2}\s+(?:dashboard|report)/i, 'Monthly report'],
  [/quarterly|\bqbr\b|\bq[1-4]\b.*\b(?:report|review|dashboard)\b/i, 'Quarterly report'],
  [/\bwrap\b|end[\s-]of[\s-]campaign|campaign\s*recap/i, 'End-of-campaign/wrap report'],
  [/\brf[ip]\b/i, 'RFI insights/audit']
];
function inferType_(title) {
  const t = String(title || '');
  const hit = TYPE_RULES.find(r => r[0].test(t));
  return hit ? hit[1] : '';
}

function findDuplicate_(sheet, key, skipRow) {
  const last = sheet.getLastRow();
  if (last < FIRST_ROW) return 0;
  const links = sheet.getRange(FIRST_ROW, COL.link, last - FIRST_ROW + 1, 1).getValues();
  for (let i = 0; i < links.length; i++) {
    const r = FIRST_ROW + i;
    if (r === skipRow) continue;
    const ref = parseLink_(String(links[i][0]));
    if (ref && ref.key === key) return r;
  }
  return 0;
}

function nextId_(sheet) {
  const last = sheet.getLastRow();
  const ids = sheet.getRange(FIRST_ROW, COL.id, Math.max(last - FIRST_ROW + 1, 1), 1).getValues();
  let max = 0;
  ids.forEach(v => { const m = String(v[0]).match(/^T-(\d+)$/); if (m) max = Math.max(max, Number(m[1])); });
  return 'T-' + String(max + 1).padStart(3, '0');
}

let TEAM_CACHE_ = null;
function teamList_() {
  if (!TEAM_CACHE_) {
    TEAM_CACHE_ = ss_().getSheetByName('Team')
      .getRange('A5:A20').getValues().map(v => String(v[0]).trim()).filter(String);
  }
  return TEAM_CACHE_;
}

// "Tayler Morgan" → "Tayler M." if that name is on the Team tab
function matchTeamName_(fullName) {
  const parts = String(fullName).replace(/\(.*?\)/g, ' ').trim().split(/\s+/);
  const first = (parts[0] || '').toLowerCase();
  const initial = (parts[1] || '').charAt(0).toLowerCase();
  return teamList_().find(t => {
    const tp = t.split(/\s+/);
    return tp[0].toLowerCase() === first &&
           (!initial || !tp[1] || tp[1].charAt(0).toLowerCase() === initial);
  }) || '';
}

// "Tayler Morgan" → "Tayler M."
function shortName_(fullName) {
  const p = String(fullName).replace(/\(.*?\)/g, ' ').trim().split(/\s+/);
  return p.length > 1 ? p[0] + ' ' + p[p.length - 1].charAt(0) + '.' : p[0];
}

function assignedText_(names) {
  return names.map(shortName_).sort().join(', ');
}

function mapTeam_(names) {
  const out = [];
  names.forEach(n => { const m = matchTeamName_(n); if (m && out.indexOf(m) === -1) out.push(m); });
  return out;
}

// Picks the account whose keyword (Lists column B, options separated by |)
// appears in the Basecamp project name
function matchClient_(projectName) {
  const p = String(projectName).toLowerCase();
  if (!p) return '';
  const rows = ss_().getSheetByName('Lists')
    .getRange('A2:B30').getValues().filter(r => r[0] && r[1]);
  for (const [account, keys] of rows) {
    if (String(keys).split('|').some(k => k.trim() && p.indexOf(k.trim().toLowerCase()) !== -1)) return account;
  }
  return '';
}

// The client for a Basecamp project: a keyword match on the Lists tab, else the project's own name
// ("Ameranth: Blue Light [40112]" → "Ameranth"), added to the Lists tab so the dropdowns and future to-dos use it.
function cleanProjectName_(name) {
  return String(name || '').replace(/\[[^\]]*\]/g, ' ').split(':')[0].replace(/[\u{1F300}-\u{1FAFF}\u2600-\u27BF]/gu, '').replace(/\s+/g, ' ').trim();
}
function clientFor_(projectName) {
  const hit = matchClient_(projectName);
  if (hit) return hit;
  const name = cleanProjectName_(projectName);
  if (!name) return 'Other';
  return addClient_(name, name);
}
// Adds an account (or a keyword to an existing one) on the Lists tab; returns the account name as listed
function addClient_(account, keyword) {
  const ls = ss_().getSheetByName('Lists');
  const vals = ls.getRange('A2:B30').getValues();
  const k = vals.findIndex(r => String(r[0]).trim().toLowerCase() === String(account).trim().toLowerCase());
  const kw = String(keyword || '').trim();
  if (k !== -1) {
    const keys = String(vals[k][1] || '').split('|').map(x => x.trim()).filter(String);
    if (kw && keys.map(x => x.toLowerCase()).indexOf(kw.toLowerCase()) === -1 && String(vals[k][0]).trim() !== 'Other') {
      ls.getRange(k + 2, 2).setValue(keys.concat([kw]).join('|'));
    }
    return String(vals[k][0]).trim();
  }
  const empty = vals.findIndex(r => !String(r[0]).trim());
  if (empty === -1) return 'Other';   // the list is full (29 accounts); pick one by hand
  ls.getRange(empty + 2, 1, 1, 2).setValues([[account, kw]]);
  return account;
}
// Rows still showing Other or a blank client: try again from their Basecamp project
function reclassifyClients_() {
  const sheet = ss_().getSheetByName(SHEET_NAME);
  const last = sheet.getLastRow(); if (last < FIRST_ROW) return 0;
  const rows = sheet.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, COL.project).getValues();
  let n = 0;
  rows.forEach((row, i) => {
    const client = String(row[COL.client - 1]).trim(), project = String(row[COL.project - 1]).trim();
    if (!row[COL.deliverable - 1] || (client && client !== 'Other') || !project || /^(Recurring \(projected\)|Planning \(booked in dashboard\))$/.test(project)) return;
    const next = clientFor_(project);
    if (next && next !== 'Other' && next !== client) { sheet.getRange(FIRST_ROW + i, COL.client).setValue(next); n++; }
  });
  return n;
}

// Returns a plain date the Sheet reads in its own time zone (avoids off-by-one-day shifts)
function toDate_(iso) {
  if (!iso) return '';
  return iso; // e.g. "2026-10-30" – Sheets converts it to a date at midnight, sheet time
}

// Full to-do description (the project brief) as plain text, keeping paragraphs and bullets
function briefOf_(item) {
  return htmlText_(item.description || item.content_description || '', BRIEF_MAX);
}
function htmlText_(raw, max) {
  const html = String(raw || '');
  if (!html) return '';
  let t = html
    .replace(/<bc-attachment[^>]*filename="([^"]*)"[^>]*>(?:<\/bc-attachment>)?/gi, '\n📎 $1\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<(ul|ol)[^>]*>/gi, '\n').replace(/<li[^>]*>/gi, '• ')
    .replace(/<\/(p|div|h[1-6]|li|ul|ol|blockquote)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  if (t.length > max) t = t.slice(0, max) + '…';
  return t;
}

// File names and link texts in a reply that mark a client delivery
function extFiles_(html) {
  const out = [];
  const h = String(html || '');
  h.replace(/<bc-attachment[^>]*filename="([^"]*)"/gi, (m, f) => { if (EXT_RX.test(f)) out.push(f); return m; });
  h.replace(/<a\b[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, (m, href, text) => {
    const t = text.replace(/<[^>]+>/g, '').trim();
    if (EXT_RX.test(t) || EXT_RX.test(safeDecode_(href.split('?')[0].split('/').pop() || ''))) out.push(t || href);
    return m;
  });
  return out;
}

function safeDecode_(t) { try { return decodeURIComponent(t); } catch (e) { return t; } }

// Reads every page of a Basecamp list (comments come back in pages)
function apiGetAll_(service, url) {
  const out = [];
  let next = url, guard = 0;
  while (next && guard++ < 20) {
    const res = UrlFetchApp.fetch(next, { headers: { Authorization: 'Bearer ' + service.getAccessToken(), 'User-Agent': userAgent_() }, muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code === 429) { Utilities.sleep(Number(res.getHeaders()['Retry-After'] || 10) * 1000); guard--; continue; }
    if (code !== 200) throw new Error('HTTP ' + code);
    JSON.parse(res.getContentText()).forEach(x => out.push(x));
    const link = String(res.getHeaders()['Link'] || res.getHeaders()['link'] || '');
    const m = link.match(/<([^>]+)>;\s*rel="next"/);
    next = m ? m[1] : null;
    if (next) Utilities.sleep(250);
  }
  return out;
}

// Latest replies on a to-do, read with whichever connection can see it
function fetchReplies_(item) {
  const url = item.comments_url;
  if (!url) return [];
  let lastErr = null;
  for (const c of services_()) {
    try {
      const all = apiGetAll_(c.service, url);
      return all.sort((a, b) => a.created_at < b.created_at ? -1 : 1).slice(-REPLIES_KEPT).map(x => ({
        at: x.created_at, by: x.creator ? x.creator.name : '', text: htmlText_(x.content, REPLY_MAX), ext: extFiles_(x.content), url: x.app_url || ''
      }));
    } catch (e) { lastErr = e; if (!/404|403/.test(e.message)) break; }
  }
  throw lastErr || new Error('replies not found');
}

function stripHtml_(html) {
  return String(html).replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&').replace(/\s+/g, ' ').trim();
}


// ===== Emails: sync alerts and the Monday digest =====
function dashboardUrl_() {
  const p = PropertiesService.getScriptProperties().getProperty('DASHBOARD_URL');
  if (p) return p;
  try { return ScriptApp.getService().getUrl() || ss_().getUrl(); } catch (e) { return ss_().getUrl(); }
}

// Everyone connected to Basecamp, unless DIGEST_TO (comma-separated emails) is set in script properties
function digestRecipients_() {
  const fixed = PropertiesService.getScriptProperties().getProperty('DIGEST_TO');
  if (fixed) return fixed.split(/[,;\s]+/).filter(x => /@/.test(x));
  const reg = registry_();
  const out = [];
  Object.keys(reg).forEach(k => { const e = String(reg[k].email || '').toLowerCase(); if (/@/.test(e) && out.indexOf(e) === -1) out.push(e); });
  return out;
}

const MAIL_CSS_ = 'font-family:Arial,sans-serif;font-size:14px;line-height:1.5;color:#242624';
function mailButton_(url, label) {
  return '<p><a href="' + url + '" style="display:inline-block;background:#6100DB;color:#fff;padding:10px 16px;border-radius:10px;text-decoration:none;font-weight:bold">' + label + '</a></p>';
}
function esc_(t) { return String(t == null ? '' : t).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

// Tells the people whose connection broke, plus whoever runs the sync
function sendSyncAlert_(run) {
  const reg = registry_();
  let owner = null;
  try { owner = JSON.parse(PropertiesService.getScriptProperties().getProperty('SYNC_OWNER') || 'null'); } catch (e) {}
  const to = [];
  const add = e => { e = String(e || '').toLowerCase(); if (/@/.test(e) && to.indexOf(e) === -1) to.push(e); };
  (run.users || []).filter(u => !u.ok).forEach(u => add(reg[u.key] && reg[u.key].email));
  if (owner) add(owner.email);
  if (!to.length) return;
  const problems = (run.users || []).filter(u => !u.ok).map(u => '<li><b>' + esc_(u.name) + '</b>: ' + esc_(u.error || 'could not read assignments') + '</li>').join('');
  const html = '<div style="' + MAIL_CSS_ + '"><p>The Insights planner\'s Basecamp sync ran into a problem, so some to-dos may be missing until it\'s fixed.</p>' +
    (problems ? '<ul>' + problems + '</ul><p>If your name is listed, open the dashboard, go to <b>Settings</b> and click <b>Reconnect my Basecamp</b>.</p>' : '') +
    (run.error ? '<p>Error: ' + esc_(run.error) + '</p>' : '') +
    (run.failed ? '<p>' + run.failed + ' rows could not be read from Basecamp.</p>' : '') +
    mailButton_(dashboardUrl_(), 'Open Settings') +
    '<p style="color:#5F6775;font-size:12px">You\'ll get this once when the sync breaks, not on every run.</p></div>';
  MailApp.sendEmail({ to: to.join(','), subject: 'Insights planner: Basecamp sync needs attention', htmlBody: html, name: 'Insights planner' });
}

// ===== Monday digest =====
// Posted to a Basecamp project's Message Board (1000heads mail is on Outlook, which quarantines email Google sends as @1000heads.com).
// Set DIGEST_PROJECT in script properties to use another project. Set DIGEST_TO to also send it by email.
const DIGEST_PROJECT_DEFAULT = '1KH Insights US Hub';
function digestProject_() { return PropertiesService.getScriptProperties().getProperty('DIGEST_PROJECT') || DIGEST_PROJECT_DEFAULT; }

function weeklyDigest(e) {
  if (e && e.triggerUid && !isSyncOwner_()) return null;
  const info = { at: new Date().toISOString(), preview: false };
  info.post = postDigestToBasecamp_(false);
  const fixed = PropertiesService.getScriptProperties().getProperty('DIGEST_TO');
  if (fixed) { sendDigest_(digestRecipients_(), false); info.count = digestRecipients_().length; }
  PropertiesService.getScriptProperties().setProperty('DIGEST_LAST', JSON.stringify(info));
  return info;
}

// Email preview to the person clicking (kept for anyone whose mail is on Gmail)
function sendDigestPreview() {
  return sendDigest_([currentEmail_()], true);
}

// A different quote each week: about 60 motivational, 25 lighthearted, cycling every 90 weeks.
// Keep additions work-safe and inclusive. Format: [quote, who].
const DIGEST_QUOTES = [
  ["Well done is better than well said.", "Benjamin Franklin"],
  ["I love deadlines. I love the whooshing noise they make as they go by.", "Douglas Adams"],
  ["The best way out is always through.", "Robert Frost"],
  ["The goal is to turn data into information, and information into insight.", "Carly Fiorina"],
  ["Alone we can do so little; together we can do so much.", "Helen Keller"],
  ["I'm not superstitious, but I am a little stitious.", "Michael Scott, The Office"],
  ["Do what you can, with what you have, where you are.", "Theodore Roosevelt"],
  ["Nothing will work unless you do.", "Maya Angelou"],
  ["Sometimes I'll start a sentence and I don't even know where it's going. I just hope I find it along the way.", "Michael Scott, The Office"],
  ["Perfection is not attainable, but if we chase perfection we can catch excellence.", "Vince Lombardi"],
  ["Research is formalized curiosity. It is poking and prying with a purpose.", "Zora Neale Hurston"],
  ["The only way to do great work is to love what you do.", "Steve Jobs"],
  ["I always wanted to be somebody, but now I realize I should have been more specific.", "Lily Tomlin"],
  ["Energy and persistence conquer all things.", "Benjamin Franklin"],
  ["Great things are done by a series of small things brought together.", "Vincent van Gogh"],
  ["The trouble with having an open mind, of course, is that people will insist on coming along and trying to put things in it.", "Terry Pratchett"],
  ["The important thing is not to stop questioning. Curiosity has its own reason for existing.", "Albert Einstein"],
  ["Get your facts first, and then you can distort them as you please.", "Mark Twain"],
  ["If you want to go fast, go alone. If you want to go far, go together.", "African proverb"],
  ["Coffee is a language in itself.", "Jackie Chan"],
  ["You miss 100% of the shots you don't take.", "Wayne Gretzky"],
  ["Act as if what you do makes a difference. It does.", "William James"],
  ["Talent wins games, but teamwork and intelligence win championships.", "Michael Jordan"],
  ["A day without sunshine is like, you know, night.", "Steve Martin"],
  ["Errors using inadequate data are much less than those using no data at all.", "Charles Babbage"],
  ["The best preparation for tomorrow is doing your best today.", "H. Jackson Brown Jr."],
  ["Done is better than perfect.", "Sheryl Sandberg"],
  ["Don't Panic.", "Douglas Adams, The Hitchhiker's Guide to the Galaxy"],
  ["Curiosity is the engine of achievement.", "Ken Robinson"],
  ["Start where you are. Use what you have. Do what you can.", "Arthur Ashe"],
  ["I knew exactly what to do. But in a much more real sense, I had no idea what to do.", "Michael Scott, The Office"],
  ["Data is a precious thing and will last longer than the systems themselves.", "Tim Berners-Lee"],
  ["If I have seen further, it is by standing on the shoulders of giants.", "Isaac Newton"],
  ["The more that you read, the more things you will know.", "Dr. Seuss"],
  ["I am so clever that sometimes I don't understand a single word of what I am saying.", "Oscar Wilde"],
  ["Plans are nothing; planning is everything.", "Dwight D. Eisenhower"],
  ["Feedback is the breakfast of champions.", "Ken Blanchard"],
  ["I can resist everything except temptation.", "Oscar Wilde"],
  ["The world cannot be understood without numbers. But the world cannot be understood with numbers alone.", "Hans Rosling"],
  ["Hard work beats talent when talent doesn't work hard.", "Tim Notke"],
  ["Ideas are easy. Implementation is hard.", "Guy Kawasaki"],
  ["The early bird gets the worm, but the second mouse gets the cheese.", "Steven Wright"],
  ["You can't use up creativity. The more you use, the more you have.", "Maya Angelou"],
  ["Happiness lies in the joy of achievement and the thrill of creative effort.", "Franklin D. Roosevelt"],
  ["All models are wrong, but some are useful.", "George E. P. Box"],
  ["Experience is the name everyone gives to their mistakes.", "Oscar Wilde"],
  ["Simplicity is prerequisite for reliability.", "Edsger W. Dijkstra"],
  ["We are what we repeatedly do. Excellence, then, is not an act, but a habit.", "Will Durant"],
  ["Never put off till tomorrow what may be done day after tomorrow just as well.", "Mark Twain"],
  ["Strive not to be a success, but rather to be of value.", "Albert Einstein"],
  ["Clarity affords focus.", "Thomas Leonard"],
  ["It is a capital mistake to theorize before one has data.", "Arthur Conan Doyle (Sherlock Holmes)"],
  ["There is only one thing in the world worse than being talked about, and that is not being talked about.", "Oscar Wilde"],
  ["To improve is to change; to be perfect is to change often.", "Winston Churchill"],
  ["The best way to have a good idea is to have lots of ideas.", "Linus Pauling"],
  ["I have made this longer than usual because I have not had time to make it shorter.", "Blaise Pascal"],
  ["Creativity takes courage.", "Henri Matisse"],
  ["Work hard, be kind, and amazing things will happen.", "Conan O'Brien"],
  ["Numbers have an important story to tell. They rely on you to give them a clear and convincing voice.", "Stephen Few"],
  ["I'm writing a book. I've got the page numbers done.", "Steven Wright"],
  ["Be so good they can't ignore you.", "Steve Martin"],
  ["Nobody can go back and start a new beginning, but anyone can start today and make a new ending.", "Maria Robinson"],
  ["A meeting is an event at which the minutes are kept and the hours are lost.", "Unknown"],
  ["Knowing is not enough; we must apply. Willing is not enough; we must do.", "Johann Wolfgang von Goethe"],
  ["Doubt kills more dreams than failure ever will.", "Suzy Kassem"],
  ["The greatest value of a picture is when it forces us to notice what we never expected to see.", "John Tukey"],
  ["Procrastination is the art of keeping up with yesterday.", "Don Marquis"],
  ["Small deeds done are better than great deeds planned.", "Peter Marshall"],
  ["I never dreamed about success. I worked for it.", "Estée Lauder"],
  ["Don't watch the clock; do what it does. Keep going.", "Sam Levenson"],
  ["We have to remember what's important in life: friends, waffles, and work. Or waffles, friends, work. Doesn't matter, but work is third.", "Leslie Knope, Parks and Recreation"],
  ["Optimism is the faith that leads to achievement.", "Helen Keller"],
  ["Far better an approximate answer to the right question, which is often vague, than an exact answer to the wrong question, which can always be made precise.", "John Tukey"],
  ["It's kind of fun to do the impossible.", "Walt Disney"],
  ["Taking on a challenge is a lot like riding a horse, isn't it? If you're comfortable while you're doing it, you're probably doing it wrong.", "Ted Lasso"],
  ["Before anything else, preparation is the key to success.", "Alexander Graham Bell"],
  ["Most of the important things in the world have been accomplished by people who have kept on trying when there seemed to be no hope at all.", "Dale Carnegie"],
  ["Do. Or do not. There is no try.", "Yoda, The Empire Strikes Back"],
  ["If you're not making mistakes, then you're not doing anything.", "John Wooden"],
  ["Above all else show the data.", "Edward Tufte"],
  ["Be quick, but don't hurry.", "John Wooden"],
  ["Just keep swimming.", "Dory, Finding Nemo"],
  ["Almost everything will work again if you unplug it for a few minutes, including you.", "Anne Lamott"],
  ["Courage doesn't always roar. Sometimes courage is the quiet voice at the end of the day saying, 'I will try again tomorrow.'", "Mary Anne Radmacher"],
  ["Life moves pretty fast. If you don't stop and look around once in a while, you could miss it.", "Ferris Bueller, Ferris Bueller's Day Off"],
  ["None of us is as smart as all of us.", "Ken Blanchard"],
  ["Make it work, make it right, make it fast.", "Kent Beck"],
  ["Most people do not listen with the intent to understand; they listen with the intent to reply.", "Stephen R. Covey"],
  ["Time is an illusion. Lunchtime doubly so.", "Douglas Adams"],
  ["Be curious, not judgmental.", "Ted Lasso"]
];
const TRACKER_URL = 'https://docs.google.com/spreadsheets/d/1YD7IShR49vSqLxBMK25NR-A96H4E1_Bu0mSJGR1qsG8/edit?gid=0#gid=0';

// Everything the digest says, computed once for both Basecamp and email
function digestData_() {
  const d = getDashboardData();
  const m = teamModel_(d);
  const tz = ss_().getSpreadsheetTimeZone();
  const T = m.T, P = m.P, add = m.add;
  const mon = add(T, -((T.getDay() + 6) % 7)), fri = add(mon, 4), sun = add(mon, 6);
  const nmon = add(mon, 7), nfri = add(mon, 11);
  const f = x => Utilities.formatDate(x, tz, 'EEE, MMM d');
  const fs = x => Utilities.formatDate(x, tz, 'MMM d');
  const CLOSED = ['Delivered', 'Dismissed'];
  const teamOf = i => { const t = []; [i.owner].concat(String(i.assigned).split(/,\s*/).map(matchTeamName_)).forEach(n => { if (n && t.indexOf(n) === -1) t.push(n); }); return t; };
  const people = d.team.filter(p => p.include && !(p.last && P(p.last) < T));
  const until = x => { const n = Math.round((x - T) / 864e5); return n <= 0 ? 'today' : n === 1 ? 'tomorrow' : n < 7 ? 'in ' + n + ' days' : n < 14 ? 'in 1 week' : 'in ' + Math.round(n / 7) + ' weeks'; };
  const nextWd = x => { let y = add(x, 1); while (!m.isWd(y)) y = add(y, 1); return y; };

  // Who's in and out
  const cal = { out: [], soon: [], back: [], hires: [], leaving: [], holidays: [] };
  const coverOf = o => { const c = m.covers.find(c => c.from === o.name && +c.s === +P(o.start)); return c ? ' · cover: ' + esc_(c.to) : ''; };
  const range = (s, e) => +s === +e ? f(s) : f(s) + ' – ' + f(e);
  d.timeOff.forEach(o => {
    const s = P(o.start), e = P(o.end); if (!s || !e || e < T) return;
    const long = m.wd(s, e) >= 5;
    const line = '<strong>' + esc_(o.name) + '</strong> · ' + esc_(o.type) + (o.half ? ' (' + (o.part || 'PM') + ' half day)' : '') + ', ' + range(s, e) + coverOf(o);
    if (s <= sun) cal.out.push(line + (long ? ' · back ' + f(nextWd(e)) : ''));
    else if (s <= add(T, 42)) cal.soon.push(line + (long ? ' · back ' + f(nextWd(e)) : ''));
    if (long && s <= T) cal.back.push('<strong>' + esc_(o.name) + '</strong> is back ' + f(nextWd(e)) + ' (' + until(nextWd(e)) + ')');
  });
  d.team.forEach(p => {
    const st = P(p.start), la = P(p.last);
    if (st && st >= add(T, -14) && st <= add(T, 90)) cal.hires.push('<strong>' + esc_(p.name) + '</strong>' + (p.role ? ' (' + esc_(p.role) + ')' : '') + (st > T ? ' starts ' + f(st) + ' (' + until(st) + ')' : ' joined ' + f(st) + '. Welcome!'));
    if (la && la >= T && la <= add(T, 90)) cal.leaving.push('<strong>' + esc_(p.name) + '</strong>\'s last day is ' + f(la) + ' (' + until(la) + ')');
  });
  d.holidays.forEach(h => { const x = P(h.date); if (x && x >= T && x <= add(T, 42)) cal.holidays.push(esc_(h.name) + ' · ' + f(x)); });
  // Summer Fridays: note this week's, and flag when the season starts or ends in the next 6 weeks
  if (d.summerFridays) {
    if (isSummerFriday_(fri) && m.isWd(fri)) cal.out.push('☀️ <strong>Summer Friday</strong> · half day for everyone, ' + f(fri));
    const sr = summerFridayRange_(T.getFullYear());
    if (sr.s > sun && sr.s <= add(T, 42)) cal.holidays.push('☀️ Summer Fridays start ' + f(sr.s) + ' (half days through ' + fs(sr.e) + ')');
    if (sr.e > fri && sr.e <= add(T, 42)) cal.holidays.push('☀️ Last Summer Friday is ' + f(sr.e));
  }

  // Work this week and coming up
  const open = d.items.filter(i => CLOSED.indexOf(i.stage) === -1);
  const due = open.filter(i => i.booking !== 'Tentative' && i.due && P(i.due) >= mon && P(i.due) <= sun).sort((a, b) => a.due < b.due ? -1 : 1);
  const dueIds = due.map(i => i.id);
  const active = m.all.filter(x => x.src.booking !== 'Tentative' && x.s && x.e && x.s <= sun && x.e >= mon && dueIds.indexOf(x.src.id) === -1 && CLOSED.indexOf(x.src.stage) === -1)
    .sort((a, b) => a.e - b.e);
  const upcoming = m.all.filter(x => x.s && x.s > sun && x.s <= add(sun, 28)).sort((a, b) => a.s - b.s);

  // Capacity (estimated)
  const pct = (nm, a, b) => { const av = m.avail(nm, a, b); return av ? Math.round(m.booked(nm, a, b) / av * 100) : null; };
  const cap = people.filter(p => !(P(p.start) && P(p.start) > nfri)).map(p => ({ name: p.name, now: P(p.start) && P(p.start) > fri ? 'new' : pct(p.name, mon, fri), next: pct(p.name, nmon, nfri) }));

  // Delivered last week (see deliveredIn_ for how catch-up updates are left out)
  const wkAgo = add(mon, -7);
  const delivered = deliveredIn_(d, P, wkAgo, mon);

  // Long leave with named cover, happening now: "who to ask"
  const away = m.covers.filter(c => c.s <= sun && c.e >= T && m.wd(c.s, c.e) >= 5).map(c => ({ from: c.from, to: c.to, until: c.e, back: nextWd(c.e) }));

  // TL;DR: a few plain sentences for anyone who only reads the top
  const list = a => a.length <= 1 ? a.join('') : a.slice(0, -1).join(', ') + ' and ' + a[a.length - 1];
  const plural = (n, w) => n + ' ' + w + (n === 1 ? '' : 's');
  const clientsOf = arr => arr.map(i => i.account || 'Other').filter((c, k, all) => all.indexOf(c) === k);
  const tldr = [];
  const dc = clientsOf(due);
  tldr.push(due.length ? plural(due.length, 'deliverable') + ' due this week' + (dc.length <= 3 ? ' for ' + list(dc) : ' across ' + dc.length + ' clients') + '.' : 'Nothing due this week.');
  if (delivered.length) tldr.push(plural(delivered.length, 'deliverable') + ' went out last week' + (clientsOf(delivered).length <= 3 ? ' (' + list(clientsOf(delivered)) + ')' : '') + '.');
  const more = [active.length ? active.length + ' more in progress' : '', upcoming.length ? upcoming.length + ' starting in the next 4 weeks' : ''].filter(String);
  if (more.length) tldr.push(list(more).replace(/^./, c => c.toUpperCase()) + '.');
  const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const outWk = d.timeOff.filter(o => P(o.end) >= mon && P(o.start) <= fri).map(o => {
    const a = P(o.start) < mon ? mon : P(o.start), b = P(o.end) > fri ? fri : P(o.end);
    return o.name + ' (' + (+a === +mon && +b === +fri ? 'all week' : +a === +b ? DOW[a.getDay()] + (o.half ? ' ' + (o.part || 'PM') : '') : DOW[a.getDay()] + '–' + DOW[b.getDay()]) + ')';
  });
  const soon = [];
  d.team.forEach(p => {
    const st = P(p.start), la = P(p.last);
    if (st && st > T && st <= add(T, 14)) soon.push(p.name + ' starts ' + fs(st));
    if (la && la >= T && la <= add(T, 14)) soon.push(p.name + '\'s last day is ' + fs(la));
  });
  d.holidays.forEach(h => { const x = P(h.date); if (x && x >= T && x <= add(T, 14)) soon.push(h.name + ' is ' + fs(x)); });
  tldr.push((outWk.length ? 'Out this week: ' + list(outWk) + '.' : 'Everyone\'s in this week.') + (soon.length ? ' ' + list(soon).replace(/^./, c => c.toUpperCase()) + '.' : ''));

  const quote = DIGEST_QUOTES[Math.floor((mon - new Date(2024, 0, 1)) / (7 * 864e5)) % DIGEST_QUOTES.length];
  const last = d.sync && d.sync.last;
  const subject = 'Insights week of ' + fs(mon) + ': ' + due.length + ' due';
  return { tz, mon, f, fs, teamOf, quote, cal, due, active, upcoming, cap, P, delivered, away, tldr, outWk, syncBad: last && !last.ok, subject, url: dashboardUrl_() };
}

function sendDigest_(to, preview) {
  if (!to.length) throw new Error('Nobody to send to yet.');
  const g = digestData_();
  const html = '<div style="' + MAIL_CSS_ + '">' + (preview ? '<p style="background:#F2E4FF;color:#6100DB;padding:8px 12px;border-radius:8px">Preview: only you received this.</p>' : '') + digestBasecampHtml_(g, false) + '</div>';
  MailApp.sendEmail({ to: to.join(','), subject: (preview ? '[Preview] ' : '') + g.subject, htmlBody: html, name: 'Insights planner' });
  return { at: new Date().toISOString(), count: to.length, preview: !!preview };
}

// Basecamp rich text allows only simple tags (div, strong, em, a, ul, li, br, h1, blockquote), so no tables or styles
// When a deliverable actually went out, as best the planner can tell:
// - a Basecamp reply with an EXT file: the reply's date
// - moved to Delivered: the earlier of the due date and the day it was marked.
//   A row due Oct 1 but marked Delivered on Oct 22 (catch-up) counts as Oct 1, so it doesn't show up as "last week".
function deliveredOn_(i, P) {
  if (i.stage === 'Dismissed') return null;
  const ext = (i.replies || []).filter(x => x.at && x.ext && x.ext.length).map(x => new Date(x.at));
  if (ext.length) return new Date(Math.max.apply(null, ext));
  if (i.stage !== 'Delivered' || !P(i.stageSince)) return null;
  const marked = P(i.stageSince), due = i.due ? P(i.due) : null;
  return due && due < marked ? due : marked;
}
function deliveredIn_(d, P, a, b) {
  return d.items.filter(i => { const t = deliveredOn_(i, P); return t && t >= a && t < b; });
}

// "8 deliverables across 4 clients went out: Google (4, including 2 monthly reports and 1 quarterly report), Valvoline (2 monthly reports) and Ancestry (1)"
function deliveredSummary_(items) {
  const list = a => a.length <= 1 ? a.join('') : a.slice(0, -1).join(', ') + ' and ' + a[a.length - 1];
  const plural = (n, w) => n === 1 ? w : /y$/i.test(w) && !/[aeiou]y$/i.test(w) ? w.slice(0, -1) + 'ies' : w + 's';
  const typeName = t => t.charAt(0).toLowerCase() + t.slice(1);
  const groups = {};
  items.forEach(i => { const c = i.account || 'Other'; (groups[c] = groups[c] || []).push(i); });
  const clients = Object.keys(groups).sort((a, b) => groups[b].length - groups[a].length || a.localeCompare(b));
  const part = c => {
    const g = groups[c], types = {};
    let untyped = 0;
    g.forEach(i => { const t = String(i.type || '').trim(); if (!t || t === 'Other') untyped++; else types[t] = (types[t] || 0) + 1; });
    const named = Object.keys(types).sort((a, b) => types[b] - types[a]).map(t => (types[t] === 1 ? (/^[aeiou]/i.test(t) ? 'an' : 'a') : types[t]) + ' ' + plural(types[t], typeName(t)));
    if (!named.length) return c + ' (' + g.length + ')';
    return c + ' (' + (untyped ? g.length + ', including ' : '') + list(named) + ')';
  };
  const n = items.length;
  return n + ' deliverable' + (n === 1 ? '' : 's') + (clients.length > 1 ? ' across ' + clients.length + ' clients' : '') + ' went out: ' + list(clients.map(part));
}

function digestBasecampHtml_(g, preview, tags) {
  const link = i => i.link ? '<a href="' + esc_(i.link) + '">' + esc_(i.title) + '</a>' : esc_(i.title);
  // tags: team name -> Basecamp person sgid. Tagged people get a Basecamp mention; anyone else shows as plain text.
  const tag = n => tags && tags[n] ? '<bc-attachment sgid="' + esc_(tags[n]) + '" content-type="application/vnd.basecamp.mention"></bc-attachment>' : esc_(n);
  const who = i => { const t = g.teamOf(i); return t.length ? t.map(tag).join(' ') : 'No owner'; };
  // Groups lines under each client, A to Z, keeping the order within each client
  const byClient = (list, src, line) => {
    const groups = {};
    list.forEach(x => { const c = src(x).account || 'Other'; (groups[c] = groups[c] || []).push(x); });
    return '<ul>' + Object.keys(groups).sort().map(c => '<li><strong>' + esc_(c) + '</strong><ul>' + groups[c].map(x => '<li>' + line(x) + '</li>').join('') + '</ul></li>').join('') + '</ul>';
  };
  const head = (emoji, t) => '<div><br><strong>' + emoji + ' ' + t + '</strong></div>';
  const ul = items => '<ul>' + items.map(x => '<li>' + x + '</li>').join('') + '</ul>';
  const none = t => '<div><em>' + t + '</em></div>';
  const c = g.cal;
  const calBlocks = [
    ['Out this week', c.out], ['Coming up (next 6 weeks)', c.soon], ['Back from leave', c.back],
    ['New starters', c.hires], ['Last days', c.leaving], ['Holidays', c.holidays]
  ].filter(x => x[1].length);
  // 🟢 under 70% (has room) · 🟡 70–100% (near full) · 🔴 over 100% · ⚪ out
  const dot = v => v === null ? '⚪' : v > 100 ? '🔴' : v >= 70 ? '🟡' : '🟢';
  const wk = (v, label) => v === 'new' ? 'starts next week' : dot(v) + ' ' + (v === null ? 'out' : v + '%') + ' ' + label;
  const capLine = r => '<strong>' + esc_(r.name) + '</strong>: ' + wk(r.now, 'this week') + ' · ' + wk(r.next, 'next week');
  return '<div>' +
    (preview ? '<div><em>Preview: a draft only you can see. Delete it when you\'re done.</em><br><br></div>' : '') +
    '<blockquote>"' + esc_(g.quote[0]) + '"<br>— ' + esc_(g.quote[1]) + '</blockquote>' +
    '<div>Good morning, team. Here\'s the week of ' + g.fs(g.mon) + '.</div>' +
    '<div><br><strong>TL;DR</strong></div>' + ul(g.tldr.map(esc_)) +
    g.away.map(a => '<div><br>🌴 <strong>' + esc_(a.from) + '</strong> is out until ' + g.f(a.until) + ' (back ' + g.fs(a.back) + '). Questions go to ' + tag(a.to) + '.</div>').join('') +

    head('🗓️', 'Who\'s in and out') +
    (calBlocks.length ? calBlocks.map(b => '<div><em>' + b[0] + '</em></div>' + ul(b[1])).join('') : none('Everyone\'s in this week, and nothing coming up.')) +

    (g.delivered.length ? head('🎉', 'Delivered last week') +
      '<div>' + esc_(deliveredSummary_(g.delivered)) + '. Thanks to ' +
      g.delivered.reduce((a, i) => a.concat(g.teamOf(i)), []).filter((n, k, all) => all.indexOf(n) === k).map(tag).join(', ') +
      ' for getting them out the door.</div>' : '') +

    head('📌', 'Due this week (' + g.due.length + ')') +
    (g.due.length ? byClient(g.due, i => i, i => link(i) + ' · due ' + g.f(g.P(i.due)) + ' · ' + who(i) + ' · <em>' + esc_(i.stage) + '</em>') : none('Nothing due this week.')) +

    head('🛠️', 'Also in progress this week (' + g.active.length + ')') +
    (g.active.length ? byClient(g.active, x => x.src, x => link(x.src) + ' · due ' + g.fs(x.e) + ' · ' + who(x.src)) : none('Nothing else in progress.')) +

    head('🔭', 'Coming up in the next 4 weeks (' + g.upcoming.length + ')') +
    (g.upcoming.length ? byClient(g.upcoming, x => x.src, x => link(x.src) + ' · starts ' + g.fs(x.s) + (x.e ? ', due ' + g.fs(x.e) : '') + ' · ' + who(x.src) + (x.src.booking === 'Tentative' ? ' · <em>tentative</em>' : '')) : none('Nothing booked yet.')) +

    head('✅', 'Reminders') +
    ul(['Complete your timesheet for last week if you haven\'t yet.',
        g.delivered.length
          ? 'Last week\'s deliveries were added to the <a href="' + TRACKER_URL + '">Centralized Insights Deliverables Tracker</a> automatically. Add the report and dataset links, and log anything delivered outside Basecamp.'
          : 'Log every delivered project in the <a href="' + TRACKER_URL + '">Centralized Insights Deliverables Tracker</a>.']) +

    head('📊', 'Team capacity') +
    ul(g.cap.map(capLine)) +
    '<div>🟢 has room (under 70%) · 🟡 near full (70–100%) · 🔴 over capacity · ⚪ out</div>' +
    '<div><em>Capacity is an estimate: each deliverable\'s estimated days, split by role and spread over its work dates, against each person\'s billable time after PTO and holidays. Use it as a guide, not a timesheet.</em></div>' +

    (g.syncBad ? '<div><br>⚠️ The last Basecamp sync had problems. Check Settings on the dashboard.</div>' : '') +
    '<div><br><a href="' + esc_(g.url) + '">Open the planner</a></div></div>';
}

// The connection to post with: the person running this (the sync owner on Mondays), else the first working one
function postingService_() {
  const all = services_();
  if (!all.length) throw new Error('No Basecamp connections yet.');
  const mine = all.find(c => c.key === userKey_(String(currentEmail_()).toLowerCase()));
  return (mine || all[0]).service;
}
function bcSend_(service, method, url, body) {
  const res = UrlFetchApp.fetch(url, { method: method, contentType: 'application/json', payload: body ? JSON.stringify(body) : undefined,
    headers: { Authorization: 'Bearer ' + service.getAccessToken(), 'User-Agent': userAgent_() }, muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code >= 300) throw new Error('Basecamp said HTTP ' + code + ': ' + res.getContentText().slice(0, 200));
  return code === 204 ? {} : JSON.parse(res.getContentText());
}
// Finds the project's Message Board (cached) and the Insights team's Basecamp people ids to notify
function digestTarget_(service) {
  const props = PropertiesService.getScriptProperties();
  const name = digestProject_();
  let t = null;
  try { t = JSON.parse(props.getProperty('DIGEST_TARGET') || 'null'); } catch (e) {}
  if (!t || t.name !== name || t.chat === undefined) {
    const projects = apiGetAll_(service, 'https://3.basecampapi.com/' + BC_ACCOUNT_ID + '/projects.json');
    const want = name.trim().toLowerCase();
    const proj = projects.find(p => String(p.name).trim().toLowerCase() === want) || projects.find(p => String(p.name).toLowerCase().indexOf(want) !== -1);
    if (!proj) throw new Error('Could not find a Basecamp project called "' + name + '" that this account can see.');
    const board = (proj.dock || []).find(x => x.name === 'message_board' && x.enabled !== false);
    if (!board) throw new Error('"' + proj.name + '" has no Message Board turned on.');
    const chat = (proj.dock || []).find(x => x.name === 'chat' && x.enabled !== false);
    t = { name: name, project: proj.name, bucket: proj.id, board: board.id, chat: chat ? chat.id : null, url: proj.app_url };
    props.setProperty('DIGEST_TARGET', JSON.stringify(t));
  }
  return t;
}
// The Insights team's Basecamp people on the digest project: ids to notify, and sgids to tag them by Team tab name
function digestPeople_(service, t) {
  const people = apiGetAll_(service, 'https://3.basecampapi.com/' + BC_ACCOUNT_ID + '/projects/' + t.bucket + '/people.json');
  const reg = registry_();
  const byEmail = {};
  Object.keys(reg).forEach(k => { const r = reg[k]; const nm = matchTeamName_(r.name) || ''; [r.bcEmail, r.email].forEach(e => { if (e && nm) byEmail[String(e).toLowerCase()] = nm; }); });
  const ids = [], tags = {};
  people.forEach(p => {
    const nm = byEmail[String(p.email_address || '').toLowerCase()] || matchTeamName_(p.name);
    if (!nm) return;
    ids.push(p.id);
    if (p.attachable_sgid && !tags[nm]) tags[nm] = p.attachable_sgid;
  });
  return { ids: ids, tags: tags };
}
function postDigestToBasecamp_(preview) {
  const service = postingService_();
  const t = digestTarget_(service);
  const g = digestData_();
  let team = { ids: [], tags: {} };
  try { team = digestPeople_(service, t); } catch (e) { console.log('Could not read project people: ' + e.message); }
  const body = { subject: (preview ? '[Preview] ' : '') + g.subject, content: digestBasecampHtml_(g, preview, team.tags) };
  if (!preview) {
    body.status = 'active';   // without a status the message stays a draft only the poster can see
    if (team.ids.length) body.subscriptions = team.ids;
  }
  const msg = bcSend_(service, 'post', 'https://3.basecampapi.com/' + BC_ACCOUNT_ID + '/buckets/' + t.bucket + '/message_boards/' + t.board + '/messages.json', body);
  const out = { at: new Date().toISOString(), project: t.project, url: msg.app_url || t.url, preview: !!preview };
  // Then a short, no-tag line in the project's Campfire pointing to it (real posts only; a failure here never blocks the digest)
  if (!preview && t.chat) {
    try {
      bcSend_(service, 'post', 'https://3.basecampapi.com/' + BC_ACCOUNT_ID + '/buckets/' + t.bucket + '/chats/' + t.chat + '/lines.json', { content: campfireLine_(g, out.url) });
      out.campfire = true;
    } catch (e) { out.campfire = false; console.log('Campfire post failed: ' + e.message); }
  }
  return out;
}
// "📬 The Insights digest for the week of Oct 26 is up. 4 due this week, 2 delivered last week, Leana L. out Thu–Fri. Read it here: <link>"
function campfireLine_(g, url) {
  const bits = [g.due.length ? g.due.length + ' due this week' : 'nothing due this week'];
  if (g.delivered.length) bits.push(g.delivered.length + ' delivered last week');
  if (g.outWk.length) bits.push(g.outWk.join(', ') + ' out');
  return '📬 The Insights digest for the week of ' + g.fs(g.mon) + ' is up. ' + bits.join(', ').replace(/^./, c => c.toUpperCase()) + '. Read it here: ' + url;
}
// ===== Campfire: Wednesday check-in and Friday wrap-up =====
function campfirePost_(text) {
  const service = postingService_();
  const t = digestTarget_(service);
  if (!t.chat) throw new Error('"' + t.project + '" has no Campfire turned on.');
  bcSend_(service, 'post', 'https://3.basecampapi.com/' + BC_ACCOUNT_ID + '/buckets/' + t.bucket + '/chats/' + t.chat + '/lines.json', { content: text });
  return { at: new Date().toISOString(), project: t.project, text: text };
}
function weekNumber_() { const now = new Date(); return Math.floor((now - new Date(2024, 0, 1)) / (7 * 864e5)); }

// Wednesday: a wellbeing nudge, different each week. Warm, optional to answer, and about looking out for each other.
// Keep additions work-safe and inclusive: invite sharing, never require it.
const CHECKIN_MESSAGES = [
  '🌿 Midweek check-in! How\'s everyone\'s week going? If your plate feels full, say so here. Someone may have room to lend a hand.',
  '☕ Happy Wednesday! Take two minutes today to ask a teammate how their week is going. A quick check-in goes a long way.',
  '🤝 Halfway there! Anyone stuck on something or waiting on someone? Drop it here and let\'s get it moving together.',
  '🌤️ Midweek moment: what\'s one thing that\'s gone well this week? Big or small, share it if you\'d like.',
  '💬 Wednesday check-in: if you\'ve got a bit of room this week, offer it up. If you could use a second pair of eyes, ask. That\'s what the team\'s for.',
  '🙌 Shout-out Wednesday! Has a teammate helped you out lately? Give them a thank-you here.',
  '🧘 Midweek reminder to step away from the screen for a few minutes today. Stretch, grab a coffee, get some air. The data will still be there.',
  '🌱 How\'s the week treating you? If anything feels off, reach out to a teammate or your manager. No issue is too small.',
  '🧩 Midweek check-in: anything you\'re working on that someone else on the team might know a trick for? Ask away.',
  '🎧 Happy hump day! What\'s keeping you going this week: a playlist, a podcast, a good snack? Share if you\'d like.',
  '📣 Wednesday nudge: if a deadline is looking tight, flag it early. It\'s much easier to help on Wednesday than on Friday.',
  '💡 Midweek check-in: learned something useful this week? A shortcut, a source, a formula? Pass it on here.',
  '🌻 Halfway through the week! Take a moment to check in on someone you haven\'t talked to in a few days.',
  '🫶 Midweek reminder: you don\'t have to carry it alone. If your week is heavier than expected, let the team know.',
  '🏁 Wednesday check-in: what\'s the one thing you most want to get done by Friday? Saying it out loud helps.',
  '😄 Happy Wednesday! Here\'s your reminder that good work and good humor go together. How\'s everyone doing?',
  '🔄 Midweek check-in: any priorities shift since Monday? Share here so everyone has the same picture.',
  '🌈 Halfway there! What\'s something you\'re looking forward to, this week or this weekend? Share if you\'d like.',
  '🤗 Wednesday check-in: ask a teammate if there\'s anything you can take off their plate, or anything they can help you with.',
  '⏸️ Midweek pause: take a breath. You\'re doing great work. Anything the team can do to make the rest of this week easier?'
];
function wednesdayCheckin(e) {
  if (e && e.triggerUid && !isSyncOwner_()) return null;
  const out = campfirePost_(CHECKIN_MESSAGES[weekNumber_() % CHECKIN_MESSAGES.length]);
  PropertiesService.getScriptProperties().setProperty('CHECKIN_LAST', out.at);
  return out;
}

// Friday: what went out this week, the timesheet and tracker reminders, and a peek at next week
function fridayWrapText_() {
  const d = getDashboardData();
  const m = teamModel_(d);
  const tz = ss_().getSpreadsheetTimeZone();
  const T = m.T, P = m.P, add = m.add;
  const mon = add(T, -((T.getDay() + 6) % 7)), sat = add(mon, 5);
  const nmon = add(mon, 7), nsun = add(mon, 13);
  const fs = x => Utilities.formatDate(x, tz, 'MMM d');
  const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const list = a => a.length <= 1 ? a.join('') : a.slice(0, -1).join(', ') + ' and ' + a[a.length - 1];
  const delivered = deliveredIn_(d, P, mon, sat);
  const dueNext = d.items.filter(i => i.booking !== 'Tentative' && ['Delivered', 'Dismissed'].indexOf(i.stage) === -1 && i.due && P(i.due) >= nmon && P(i.due) <= nsun);
  const outNext = d.timeOff.filter(o => P(o.end) >= nmon && P(o.start) <= add(nmon, 4)).map(o => o.name);
  const hol = d.holidays.filter(h => P(h.date) >= nmon && P(h.date) <= nsun).map(h => h.name + ' ' + DOW[P(h.date).getDay()]);
  const parts = [];
  parts.push(delivered.length
    ? '🎉 Week wrapped! ' + deliveredSummary_(delivered).replace(/went out$/, 'went out this week') + '. Nice work, team.'
    : '🎉 Week wrapped! Thanks for all the work this week, team.');
  parts.push('Before you log off: please complete your timesheet, and add report links to this week\'s rows in the Centralized Insights Deliverables Tracker (' + TRACKER_URL + '). Deliveries from Basecamp are logged there automatically.');
  const next = [dueNext.length ? dueNext.length + ' due' : 'nothing due'].concat(hol).concat(outNext.length ? [list(outNext.filter((n, k) => outNext.indexOf(n) === k)) + ' out'] : []);
  parts.push('Next week (' + fs(nmon) + '): ' + next.join(', ') + '. Have a great weekend!');
  return parts.join(' ');
}
function fridayWrap(e) {
  if (e && e.triggerUid && !isSyncOwner_()) return null;
  const out = campfirePost_(fridayWrapText_());
  PropertiesService.getScriptProperties().setProperty('WRAP_LAST', out.at);
  return out;
}

function postDigestPreview() { return postDigestToBasecamp_(true); }

// ===== Activity feed and timeline changes =====
// Every event lands on the Activity tab: replies, client deliveries, timeline changes, new to-dos, stage moves, notes
const ACTIVITY_SHEET = 'Activity';
const ACTIVITY_KEEP = 1000;
function activitySheet_() {
  const ss = ss_();
  let a = ss.getSheetByName(ACTIVITY_SHEET);
  if (!a) {
    a = ss.insertSheet(ACTIVITY_SHEET);
    a.appendRow(['When', 'Type', 'ID', 'Deliverable', 'Who', 'What', 'Link']);
    a.setFrozenRows(1);
    a.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#622E90').setFontColor('#FFFFFF');
  }
  return a;
}
function logActivity_(type, id, title, who, text, link, when) {
  try {
    activitySheet_().appendRow([when ? new Date(when) : new Date(), type, id || '', title || '', who || '', String(text || '').slice(0, 1000), link || '']);
  } catch (e) { console.log('Activity log failed: ' + e.message); }
}
function trimActivity_() {
  const a = activitySheet_(), n = a.getLastRow() - 1;
  if (n > ACTIVITY_KEEP + 200) a.deleteRows(2, n - ACTIVITY_KEEP);
}

// Work back schedule parser (same rules as the dashboard)
// A heading is any short line that mentions WBS, work back, timeline or schedule
// ("Tentative WBS:", "Here's the WBS we discussed", "Revised timeline (pending sign-off):")
const WBS_HEAD_ = /\b(?:wbs|work\s*-?\s*backs?|workbacks?|timelines?|schedules?)\b/i;
const MONTHS_ = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, sept: 8, oct: 9, nov: 10, dec: 11 };
const DAYW_ = '(?:mon|tue|tues|wed|weds|thu|thur|thurs|fri|sat|sun)[a-z]*\\.?,?\\s*';
const RX_NAMED_ = new RegExp('(?:' + DAYW_ + ')?([a-z]{3,9})\\.?\\s+(\\d{1,2})(?:st|nd|rd|th)?(?:\\s*(?:-|–|—|to|through|thru)\\s*(?:' + DAYW_ + ')?(?:([a-z]{3,9})\\.?\\s+)?(\\d{1,2})(?:st|nd|rd|th)?)?', 'i');
const RX_NUM_ = new RegExp('(\\d{1,2})/(\\d{1,2})(?:/\\d{2,4})?(?:\\s*(?:-|–|—|to|through|thru)\\s*(?:' + DAYW_ + ')?(?:(\\d{1,2})/)?(\\d{1,2})(?:/\\d{2,4})?)?', 'i'); // "11/6—Monday, 11/9"
const WORK_RX_ = /develop|draft|build|analy|revis|edit|final|prep|writ|pull|data|design|updat|production|creat|research|setup|set up/i;
const MS_RX_ = /\bko\b|kick\s*-?\s*off|review|present|deliver|\bdue\b|\bsend\b|\bshare\b|approv|meeting|\bcall\b|launch|go\s*live/i;
// Single-day steps that mark a hand-off are milestones even if they also say "final" or "report" ("Final report shared: 11/9")
const MS_STRONG_ = /\b(?:ko|kick\s*-?\s*off|shar(?:e|ed|ing)|sen[dt]|sending|deliver(?:ed|y|s)?|due|present(?:ed|ation|s)?|launch(?:es|ed)?|go(?:es)?\s*live)\b/i;
function nearYear_(m, day, anchor) {
  let best = null;
  [-1, 0, 1].forEach(k => { const c = new Date(anchor.getFullYear() + k, m, day); if (!best || Math.abs(c - anchor) < Math.abs(best - anchor)) best = c; });
  return best;
}
function parseWhen_(text, anchor) {
  const t = text.replace(/\(.*?\)/g, ' ').trim();
  let m = t.match(RX_NAMED_);
  if (m && MONTHS_[m[1].toLowerCase().slice(0, 4)] === undefined && MONTHS_[m[1].toLowerCase().slice(0, 3)] === undefined) m = null;
  if (m) {
    const mo = MONTHS_[m[1].toLowerCase().slice(0, 4)] ?? MONTHS_[m[1].toLowerCase().slice(0, 3)];
    const s = nearYear_(mo, +m[2], anchor);
    if (!m[4]) return { s, e: s };
    const mo2 = m[3] ? (MONTHS_[m[3].toLowerCase().slice(0, 4)] ?? MONTHS_[m[3].toLowerCase().slice(0, 3)]) : mo;
    if (mo2 === undefined) return { s, e: s };
    let e = new Date(s.getFullYear(), mo2, +m[4]); if (e < s) e = new Date(s.getFullYear() + 1, mo2, +m[4]);
    return { s, e };
  }
  const n = t.match(RX_NUM_);
  if (n) {
    const s = nearYear_(+n[1] - 1, +n[2], anchor);
    if (!n[4]) return { s, e: s };
    let e = new Date(s.getFullYear(), (n[3] ? +n[3] : +n[1]) - 1, +n[4]); if (e < s) e = new Date(s.getFullYear() + 1, e.getMonth(), e.getDate());
    return { s, e };
  }
  return null;
}
function wbsSteps_(lines, from, anchor) {
  const steps = [];
  let misses = 0;
  for (let k = from; k < lines.length && steps.length < 20; k++) {
    const raw = lines[k].replace(/^[\s•*\-–·]+/, '').trim();
    if (!raw) continue;
    const c = raw.indexOf(':');
    const when = c > 0 ? parseWhen_(raw.slice(c + 1), anchor) : null;
    if (!when) { if (steps.length && ++misses >= 2) break; continue; }
    misses = 0;
    const label = raw.slice(0, c).trim();
    const span = Math.round((when.e - when.s) / 864e5);
    const kind = span === 0 && MS_STRONG_.test(label) ? 'milestone' : WORK_RX_.test(label) ? 'work' : MS_RX_.test(label) ? 'milestone' : span > 0 ? 'work' : 'milestone';
    steps.push({ label, s: when.s, e: when.e, kind, note: (raw.slice(c + 1).match(/\(([^)]*)\)/) || [])[1] || '' });
  }
  return steps;
}
function parseWbs_(brief, anchor) {
  if (!brief) return null;
  const lines = String(brief).split(/\n/);
  // 1) Steps under a heading line that mentions WBS / work back / timeline / schedule
  for (let at = 0; at < lines.length; at++) {
    const l = lines[at].replace(/^[\s•*\-–·]+/, '').trim();
    if (!l || l.length > 160 || !WBS_HEAD_.test(l)) continue;
    const c = l.indexOf(':');
    if (c > 0 && parseWhen_(l.slice(c + 1), anchor)) continue; // "Timeline: Oct 28" is a step, not a heading
    const steps = wbsSteps_(lines, at + 1, anchor);
    if (steps.length >= 2) return wbsFillGaps_(steps.sort((a, b) => a.s - b.s));
  }
  // 2) No heading: three or more "Phase: date" lines in a row still count
  const steps = wbsSteps_(lines, 0, anchor);
  return steps.length >= 3 ? wbsFillGaps_(steps.sort((a, b) => a.s - b.s)) : null;
}
// Weekdays between two steps that no step covers are work time (e.g. between "INT KO: 11/2" and "Internal review: 11/5").
// Without this, a WBS with no development phase would book all the effort into revisions.
function wbsFillGaps_(steps) {
  const day = 864e5, out = [];
  steps.forEach((st, k) => {
    if (k) {
      const prevEnd = steps.slice(0, k).reduce((m, z) => z.e > m ? z.e : m, steps[0].e);
      const a = new Date(prevEnd.getFullYear(), prevEnd.getMonth(), prevEnd.getDate() + 1), b = new Date(st.s.getFullYear(), st.s.getMonth(), st.s.getDate() - 1);
      let s0 = null, e0 = null;
      for (let t = new Date(a); t <= b; t = new Date(t.getFullYear(), t.getMonth(), t.getDate() + 1)) if (t.getDay() % 6) { if (!s0) s0 = t; e0 = t; }
      if (s0) out.push({ label: 'Work time (not named in the WBS)', s: s0, e: e0, kind: 'work', note: 'Gap between ' + steps[k - 1].label + ' and ' + st.label });
    }
    out.push(st);
  });
  return out;
}

const TIMELINE_WORDS_ = /\b(push(?:ed|ing)?|mov(?:e|ed|ing)\s+(?:to|up|back|out)|shift(?:ed|ing)?|delay(?:ed|s)?|extend(?:ed)?|reschedul\w*|postpon\w*|bump(?:ed)?|new\s+(?:due\s+)?date|due\s+date|deadline|timeline|timing)\b/i;
const DATEISH_ = /\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s+\d{1,2}\b|\b\d{1,2}\/\d{1,2}\b|\b(?:mon|tues|wednes|thurs|fri)day\b|\bnext\s+week\b|\bEO[DW]\b/i;

function rangeText_(x) {
  const M = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const f = d => M[d.getMonth()] + ' ' + d.getDate();
  return +x.s === +x.e ? f(x.s) : f(x.s) + (x.s.getMonth() === x.e.getMonth() ? '–' + x.e.getDate() : '–' + f(x.e));
}
// What changed between two schedules, in plain words
function diffWbs_(a, b) {
  const key = x => x.label.toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  const A = {}, B = {}, out = [];
  a.forEach(x => { A[key(x)] = x; }); b.forEach(x => { B[key(x)] = x; });
  b.forEach(x => { const o = A[key(x)]; if (!o) out.push(x.label + ' added: ' + rangeText_(x)); else if (+o.s !== +x.s || +o.e !== +x.e) out.push(x.label + ': ' + rangeText_(o) + ' → ' + rangeText_(x)); });
  a.forEach(x => { if (!B[key(x)]) out.push(x.label + ' removed'); });
  return out;
}

// Handles replies the planner hasn't seen: logs them, flags timeline changes, marks EXT deliveries.
// quiet = true for a first read of a to-do's history (no feed entries or flags, only the EXT check),
// or an ISO time: replies at or before it are filed quietly (older history the planner hadn't stored yet).
function applyNewReplies_(sheet, r, row, stored, fresh, tz, count, quiet) {
  const id = row[COL.id - 1], title = row[COL.deliverable - 1], link = row[COL.link - 1];
  const due = row[COL.due - 1] instanceof Date ? row[COL.due - 1] : new Date();
  const seen = (stored && stored.list) ? stored.list.slice() : [];
  let flag = stored && stored.flag ? stored.flag : null, ext = null;
  const latestWbs = list => { for (let k = list.length - 1; k >= 0; k--) { const w = parseWbs_(list[k].text, due); if (w) return w; } return parseWbs_(String(row[COL.brief - 1] || ''), due); };
  fresh.sort((x, y) => x.at < y.at ? -1 : 1).forEach(x => {
    const hush = quiet === true || (typeof quiet === 'string' && x.at <= quiet);
    if (!hush) {
      logActivity_(x.ext && x.ext.length ? 'delivered' : 'reply', id, title, x.by, x.text.slice(0, 400), x.url || link, x.at);
      const w = parseWbs_(x.text, due);
      if (w) {
        const prev = latestWbs(seen);
        const changes = prev ? diffWbs_(prev, w) : ['New work back schedule posted'];
        if (changes.length) {
          flag = { kind: 'wbs', at: x.at, by: x.by, changes: changes.slice(0, 8) };
          logActivity_('timeline', id, title, x.by, 'Work back schedule ' + (prev ? 'changed: ' : 'posted: ') + changes.join('; '), x.url || link, x.at);
        }
      } else if (TIMELINE_WORDS_.test(x.text) && DATEISH_.test(x.text)) {
        flag = { kind: 'mention', at: x.at, by: x.by, changes: [x.text.replace(/\s+/g, ' ').slice(0, 220)] };
        logActivity_('timeline', id, title, x.by, 'Possible timeline change: ' + x.text.replace(/\s+/g, ' ').slice(0, 220), x.url || link, x.at);
      }
    }
    if (x.ext && x.ext.length) ext = x;
    seen.push(x);
  });
  const out = { n: count || ((stored && stored.n) || 0) + fresh.length, list: seen.slice(-REPLIES_KEPT), flag: flag };
  sheet.getRange(r, COL.replies).setValue(JSON.stringify(out));
  let delivered = false;
  if (ext && row[COL.stage - 1] !== DONE_STAGE) {
    setStage_(sheet, r, DONE_STAGE);
    const when = Utilities.formatDate(new Date(ext.at), tz, 'MMM d');
    const note = 'Delivered to client ' + when + ' (' + ext.ext[0] + ', posted by ' + shortName_(ext.by) + ')';
    const prevNote = String(row[COL.notes - 1] || '');
    if (prevNote.indexOf(note) === -1) sheet.getRange(r, COL.notes).setValue(prevNote ? prevNote + ' · ' + note : note);
    logActivity_('stage', id, title, 'Basecamp sync', 'Marked Delivered: ' + ext.ext[0] + ' posted by ' + ext.by, link);
    delivered = true;
  }
  return { stored: out, delivered: delivered };
}

// Every 15 minutes: one request per connection for the latest replies across Basecamp
function pollReplies(e) {
  if (e && e.triggerUid && !isSyncOwner_()) return null;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return null;   // a full sync is running; it reads replies itself
  try {
    const ss = ss_(), tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const last = sheet.getLastRow();
    if (last < FIRST_ROW) return null;
    const data = sheet.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, COL.replies).getValues();
    const byKey = {};
    data.forEach((row, i) => { const ref = parseLink_(String(row[COL.link - 1] || '')); if (ref && row[COL.deliverable - 1] && row[COL.stage - 1] !== 'Dismissed') byKey[ref.key] = i; });
    const props = PropertiesService.getScriptProperties();
    const found = {}, ids = {};
    services_().forEach(c => {
      const pk = 'POLL_SINCE_' + c.key;
      const since = props.getProperty(pk) || new Date(Date.now() - 3600e3).toISOString();
      let newest = since, url = 'https://3.basecampapi.com/' + BC_ACCOUNT_ID + '/projects/recordings.json?type=Comment&sort=created_at&direction=desc', pages = 0;
      try {
        while (url && pages++ < 3) {
          const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + c.service.getAccessToken(), 'User-Agent': userAgent_() }, muteHttpExceptions: true });
          if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
          const list = JSON.parse(res.getContentText());
          let older = false;
          list.forEach(x => {
            if (x.created_at <= since) { older = true; return; }
            if (x.created_at > newest) newest = x.created_at;
            const k = x.bucket && x.parent ? x.bucket.id + '/' + x.parent.id : '';
            if (byKey[k] === undefined || ids[x.id]) return;
            ids[x.id] = 1;
            (found[k] = found[k] || []).push({ at: x.created_at, by: x.creator ? x.creator.name : '', text: htmlText_(x.content, REPLY_MAX), ext: extFiles_(x.content), url: x.app_url || '' });
          });
          if (older) break;
          const m = String(res.getHeaders()['Link'] || res.getHeaders()['link'] || '').match(/<([^>]+)>;\s*rel="next"/);
          url = m ? m[1] : null;
        }
        props.setProperty(pk, newest);
      } catch (err) { console.log('Reply check failed for ' + c.name + ': ' + err.message); }
    });
    let n = 0;
    Object.keys(found).forEach(k => {
      const i = byKey[k], row = data[i], r = FIRST_ROW + i;
      let stored = null;
      try { stored = JSON.parse(String(row[COL.replies - 1] || 'null')); } catch (err) {}
      const known = {}; ((stored && stored.list) || []).forEach(x => { known[x.url || x.at] = 1; });
      const fresh = found[k].filter(x => !known[x.url || x.at]);
      if (!fresh.length) return;
      applyNewReplies_(sheet, r, row, stored, fresh, tz, stored && stored.n ? stored.n + fresh.length : 0, false);
      n += fresh.length;
    });
    props.setProperty('POLL_LAST', JSON.stringify({ at: new Date().toISOString(), replies: n }));
    trimActivity_();
    return n;
  } finally {
    lock.releaseLock();
  }
}


// ===== Recurring work: pencil in the next 90 days as Tentative =====
const RECURRING_SOURCE = 'Recurring';
const MONTH_NAMES_ = ['January','February','March','April','May','June','July','August','September','October','November','December'];
function recurringPeriods_(cadence, dueDay, today, horizon) {
  const out = [];
  // Day of the month, moved back to Friday when it lands on a weekend
  const clampDay = (y, m, d) => { const x = new Date(y, m, Math.min(d, new Date(y, m + 1, 0).getDate())); const w = x.getDay(); if (w === 6) x.setDate(x.getDate() - 1); if (w === 0) x.setDate(x.getDate() - 2); return x; };
  if (cadence === 'Quarterly') {
    const q0 = Math.floor(today.getMonth() / 3);
    for (let k = -1; k <= 2; k++) {
      const qs = new Date(today.getFullYear(), (q0 + k) * 3, 1);
      const q = Math.floor(qs.getMonth() / 3) + 1, y = qs.getFullYear();
      const due = clampDay(y, qs.getMonth() + 3, dueDay);
      out.push({ label: 'Q' + q + ' ' + y, key: y + '-Q' + q, due: due, q: q, y: y });
    }
  } else {
    for (let k = -1; k <= 4; k++) {
      const ms = new Date(today.getFullYear(), today.getMonth() + k, 1);
      const due = clampDay(ms.getFullYear(), ms.getMonth() + 1, dueDay);
      out.push({ label: MONTH_NAMES_[ms.getMonth()] + " '" + String(ms.getFullYear()).slice(2), key: ms.getFullYear() + '-' + ('0' + (ms.getMonth() + 1)).slice(-2), due: due, m: ms.getMonth(), y: ms.getFullYear() });
    }
  }
  return out.filter(p => p.due >= today && p.due <= horizon);
}
function titleMatchesPeriod_(title, p) {
  const t = String(title || '').toLowerCase();
  const yr = new RegExp("\\b" + p.y + "\\b|['’]" + String(p.y).slice(2) + "\\b");
  const otherYear = /\b20\d{2}\b|['’]\d{2}\b/.test(t) && !yr.test(t);
  if (p.q) return new RegExp('\\bq' + p.q + '\\b').test(t) && !otherYear;
  const full = MONTH_NAMES_[p.m].toLowerCase(), short = full.slice(0, 3);
  return (new RegExp('\\b' + full + '\\b').test(t) || new RegExp('\\b' + short + '\\b').test(t)) && !otherYear;
}

function generateRecurring_() {
  const ss = ss_(), rs = ss.getSheetByName('Recurring');
  const result = { added: [], retired: [] };
  if (!rs || rs.getLastRow() < 4) return result;
  const str = v => String(v == null ? '' : v).trim();
  const series = rs.getRange(4, 1, rs.getLastRow() - 3, 8).getValues()
    .filter(r => str(r[0]) && str(r[1]) && str(r[7]) === 'Yes')
    .map(r => ({ name: str(r[0]), account: str(r[1]), type: str(r[2]), team: str(r[3]).split(/[;,]+/).map(x => x.trim()).filter(String), cadence: str(r[4]) || 'Monthly', day: Number(r[5]) || 15, est: Number(r[6]) || '' }));
  if (!series.length) return result;
  const sheet = ss.getSheetByName(SHEET_NAME);
  const tz = ss.getSpreadsheetTimeZone();
  const now = new Date(), today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const horizon = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 90);
  const last = sheet.getLastRow();
  const rows = last < FIRST_ROW ? [] : sheet.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, COL.booking).getValues();
  const keyOf = row => (String(row[COL.notes - 1]).match(/\[recurring:([^\]]+)\]/) || [])[1] || '';
  const near = (a, b) => a instanceof Date && Math.abs(a - b) <= 12 * 864e5;
  const words = t => String(t || '').toLowerCase().replace(/[^a-z0-9 ]+/g, ' ').split(/\s+/).filter(String);
  const nameIn = (name, title) => { const tw = words(title); return words(name).every(w => tw.indexOf(w) !== -1); };
  const matches = (row, s, p) => str(row[COL.client - 1]) === s.account && str(row[COL.stage - 1]) !== 'Dismissed' &&
    (str(row[COL.type - 1]) === s.type || nameIn(s.name, row[COL.deliverable - 1])) &&
    (near(row[COL.due - 1], p.due) || titleMatchesPeriod_(row[COL.deliverable - 1], p));

  series.forEach(s => {
    recurringPeriods_(s.cadence, s.day, today, horizon).forEach(p => {
      const key = s.name + '|' + p.key;
      const projected = rows.map((row, i) => ({ row, i })).filter(x => keyOf(x.row) === key);
      const real = rows.find(row => str(row[COL.source - 1]) !== RECURRING_SOURCE && matches(row, s, p));
      if (real) {
        // The Basecamp to-do arrived: retire the pencilled-in row
        projected.forEach(x => {
          const st = str(x.row[COL.stage - 1]);
          if (st === 'Dismissed' || st === DONE_STAGE) return;
          const r = FIRST_ROW + x.i;
          setStage_(sheet, r, 'Dismissed');
          sheet.getRange(r, COL.notes).setValue(str(x.row[COL.notes - 1]) + ' · Replaced by ' + str(real[COL.id - 1]));
          x.row[COL.stage - 1] = 'Dismissed';
          result.retired.push(str(x.row[COL.id - 1]));
          logActivity_('stage', str(x.row[COL.id - 1]), str(x.row[COL.deliverable - 1]), 'Recurring', 'Pencilled-in row replaced by the real to-do ' + str(real[COL.id - 1]), '');
        });
        return;
      }
      if (projected.length) return;   // already pencilled in (or dismissed by someone: leave it)
      const r = firstEmptyRow_(sheet);
      const v = {};
      v[COL.id] = nextId_(sheet);
      v[COL.client] = s.account;
      v[COL.project] = 'Recurring (projected)';
      v[COL.deliverable] = s.name + ': ' + p.label + ' (projected)';
      v[COL.type] = s.type;
      if (s.est) v[COL.override] = s.est;
      v[COL.owner] = s.team[0] || '';
      v[COL.assigned] = s.team.join(', ');
      v[COL.stage] = NEW_STAGE;
      v[COL.due] = p.due;
      v[COL.priority] = DEFAULT_PRIORITY;
      v[COL.notes] = 'Pencilled in from the Recurring tab [recurring:' + key + ']';
      v[COL.source] = RECURRING_SOURCE;
      v[COL.added] = new Date();
      v[COL.booking] = 'Tentative';
      Object.keys(v).forEach(c => sheet.getRange(r, Number(c)).setValue(v[c]));
      const newRow = []; newRow[COL.client - 1] = s.account; newRow[COL.type - 1] = s.type; newRow[COL.deliverable - 1] = v[COL.deliverable];
      newRow[COL.stage - 1] = NEW_STAGE; newRow[COL.due - 1] = p.due; newRow[COL.notes - 1] = v[COL.notes]; newRow[COL.source - 1] = RECURRING_SOURCE; newRow[COL.id - 1] = v[COL.id];
      rows[r - FIRST_ROW] = newRow;
      result.added.push(v[COL.id]);
      logActivity_('new', v[COL.id], v[COL.deliverable], 'Recurring', 'Pencilled in as Tentative, due ' + Utilities.formatDate(p.due, tz, 'MMM d'), '');
    });
  });
  return result;
}

// ===== Team tab: booked days with the dashboard's rules (role split, splits, WBS phases, leave cover) =====
function writeTeamCapacity_() {
  const tm = ss_().getSheetByName('Team');
  const m = teamModel_(getDashboardData());
  let wEnd = m.T, n = m.isWd(m.T) ? 1 : 0;
  while (n < 10) { wEnd = m.add(wEnd, 1); if (m.isWd(wEnd)) n++; }
  // Working days (after time off, half days and Summer Fridays), available days, booked days, and "Out this week?"
  const r1 = x => Math.round(x * 10) / 10;
  const team = tm.getRange('A5:F14').getValues();
  const mon = m.add(m.T, -((m.T.getDay() + 6) % 7)), fri = m.add(mon, 4);
  const d = m.d;
  tm.getRange('G5:I14').setValues(team.map(r => {
    const nm = String(r[0]).trim();
    if (!nm) return ['', '', ''];
    if (String(r[5]).trim() !== 'Yes') return [0, 0, r1(m.booked(nm, m.T, wEnd))];
    return [r1(m.daysIn(nm, m.T, wEnd)), r1(m.avail(nm, m.T, wEnd)), r1(m.booked(nm, m.T, wEnd))];
  }));
  tm.getRange('O5:O14').setValues(team.map(r => {
    const nm = String(r[0]).trim(); if (!nm) return [''];
    const p = d.team.find(x => x.name === nm);
    if (p && p.last && m.P(p.last) < m.T) return ['Left'];
    const o = d.timeOff.filter(x => x.name === nm && m.P(x.start) <= fri && m.P(x.end) >= mon);
    return [o.length ? (o.every(x => x.half) ? 'Half day' : 'Yes') : ''];
  }));
}

// Summer Fridays: half days from the Friday before Memorial Day through the Friday before Labor Day
// ===== Expiry: open rows due more than STALE_DAYS ago leave the planner =====
function expireOld_() {
  const ss = ss_(), sheet = ss.getSheetByName(SHEET_NAME), tz = ss.getSpreadsheetTimeZone();
  const last = sheet.getLastRow(); if (last < FIRST_ROW) return 0;
  const cutoff = new Date(Date.now() - STALE_DAYS * 864e5);
  const rows = sheet.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, COL.notes).getValues();
  let n = 0;
  rows.forEach((row, i) => {
    const stage = String(row[COL.stage - 1]).trim(), due = row[COL.due - 1];
    if (!row[COL.deliverable - 1] || !(due instanceof Date) || due >= cutoff || stage === DONE_STAGE || stage === 'Dismissed') return;
    const r = FIRST_ROW + i, when = Utilities.formatDate(due, tz, 'MMM d');
    setStage_(sheet, r, 'Dismissed');
    const note = 'Expired: due ' + when + ', more than ' + STALE_DAYS + ' days ago';
    sheet.getRange(r, COL.notes).setValue(row[COL.notes - 1] ? row[COL.notes - 1] + ' · ' + note : note);
    logActivity_('stage', row[COL.id - 1], row[COL.deliverable - 1], 'Planner', note + '. Set it back to a stage if it\'s still live.', row[COL.link - 1]);
    n++;
  });
  return n;
}

// ===== Backup schedule: every connected teammate's account checks every 2 hours =====
// If the main sync (run by the sync owner) hasn't run for 20 hours, or a team post was missed, the backup does it.
// So the planner keeps working if the owner's account is out, suspended or their triggers stop.
const BACKUP_EVERY_HOURS = 2;
function backupRun(e) {
  const cache = CacheService.getScriptCache();
  if (cache.get('backup_running')) return null;
  cache.put('backup_running', '1', 1500);
  const props = PropertiesService.getScriptProperties();
  const read = k => { try { return JSON.parse(props.getProperty(k) || 'null'); } catch (err) { return null; } };
  const tz = ss_().getSpreadsheetTimeZone(), now = new Date();
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd'), hour = Number(Utilities.formatDate(now, tz, 'H')), dow = Number(Utilities.formatDate(now, tz, 'u'));
  const dayOf = iso => iso ? Utilities.formatDate(new Date(iso), tz, 'yyyy-MM-dd') : '';
  const who = 'Backup (' + (typeof displayName_ === 'function' ? displayName_() : currentEmail_()) + ')';
  const did = [];
  try {
    const last = read('SYNC_LAST');
    if (!last || (now - new Date(last.at)) / 36e5 >= 20) { runSync_(who); did.push('sync'); }
    const dg = read('DIGEST_LAST');
    if (dow === 1 && hour >= DIGEST_HOUR + 2 && dayOf(dg && dg.at) !== today) { weeklyDigest(); did.push('digest'); }
    if (dow === 3 && hour >= CHECKIN_HOUR + 2 && dayOf(props.getProperty('CHECKIN_LAST')) !== today) { wednesdayCheckin(); did.push('check-in'); }
    if (dow === 5 && hour >= WRAP_HOUR + 1 && dayOf(props.getProperty('WRAP_LAST')) !== today) { fridayWrap(); did.push('wrap'); }
    if (did.length) props.setProperty('BACKUP_LAST', JSON.stringify({ at: now.toISOString(), by: who, did: did }));
  } finally { cache.remove('backup_running'); }
  return did;
}
// Adds the backup schedule under the current account (once). Called when someone connects and when they open the dashboard.
function ensureBackupTrigger_() {
  const key = userKey_(String(currentEmail_()).toLowerCase());
  const reg = registry_();
  if (!reg[key]) return false;
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'backupRun')) {
    ScriptApp.newTrigger('backupRun').timeBased().everyHours(BACKUP_EVERY_HOURS).create();
  }
  if (!reg[key].backup) { reg[key].backup = new Date().toISOString(); saveRegistry_(reg); }
  return true;
}

// ===== Auto-log deliveries to the 1KH Centralized Insights Deliverables Tracker =====
// Each Delivered row (from the last 30 days) gets one row on the tracker's client tab. Turn off with Lists tab Q14 = No.
const TRACKER_ID = '1YD7IShR49vSqLxBMK25NR-A96H4E1_Bu0mSJGR1qsG8';
const TRACKER_TYPES = [
  [/monthly/i, 'Monthly'], [/quarterly|qbr/i, 'Quarterly'], [/listening/i, 'Social listening'], [/talent/i, 'Talent audit'],
  [/4cs|audit/i, 'Audit'], [/brand lift|survey/i, 'Survey'], [/24hr|48hr|72hr/i, 'Ad-hoc'], [/week|month|wrap|campaign/i, 'Campaign Recap']
];
function trackerTarget_(account, project) {
  const a = String(account || ''), p = String(project || ''), both = a + ' ' + p;
  if (/^google/i.test(a) || (!a && /google/i.test(p))) {
    const client = /chrome/i.test(both) ? 'Chrome' : /workspace/i.test(both) ? 'Workspace' : /cloud/i.test(both) ? 'Google Cloud'
      : /pixel/i.test(both) ? 'Pixel' : /photos/i.test(both) ? 'Google Photos' : /labs/i.test(both) ? 'Google Labs' : /startups|\bgfs\b/i.test(both) ? 'GFS'
      : /googlebook/i.test(both) ? 'Googlebook' : (a.replace(/^Google\s*[–-]\s*/i, '').trim() || 'Unknown');
    return { tab: 'Google', client: client };
  }
  const tabs = { 'nature made': 'NatureMade', naturemade: 'NatureMade', ancestry: 'Ancestry', valvoline: 'Valvoline', henkel: 'Henkel', verizon: 'Verizon' };
  const hit = Object.keys(tabs).find(k => a.toLowerCase().indexOf(k) === 0);
  return hit ? { tab: tabs[hit], client: a } : { tab: 'Ad-hoc', client: a && a !== 'Other' ? a : p };
}
function logToTracker_() {
  const ss = ss_();
  const ls = ss.getSheetByName('Lists');
  if (String(ls.getRange('Q14').getValue()).trim() === 'No') return 0;
  const sheet = ss.getSheetByName(SHEET_NAME), tz = ss.getSpreadsheetTimeZone();
  const last = sheet.getLastRow(); if (last < FIRST_ROW) return 0;
  const rows = sheet.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, COL.trackerLogged).getValues();
  // Only deliveries from when auto-logging was switched on, so past work isn't logged twice
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TRACKER_START')) props.setProperty('TRACKER_START', new Date().toISOString());
  const since = new Date(props.getProperty('TRACKER_START'));
  const extOf = row => { try { const st = JSON.parse(String(row[COL.replies - 1] || 'null')); const hit = st && (st.list || []).slice().reverse().find(y => y.ext && y.ext.length); return hit ? hit.ext[0] : ''; } catch (e) { return ''; } };
  const todo = rows.map((row, i) => ({ row, r: FIRST_ROW + i })).filter(x => {
    const row = x.row, when = row[COL.stageSince - 1];
    if (!row[COL.deliverable - 1] || String(row[COL.stage - 1]).trim() !== DONE_STAGE || row[COL.trackerLogged - 1] || !(when instanceof Date) || when < since) return false;
    // A client deliverable: an EXT file went out, or it has a real Type (not blank or Other)
    const type = String(row[COL.type - 1] || '').trim();
    return !!extOf(row) || (type && type !== 'Other');
  });
  if (!todo.length) return 0;
  const tr = SpreadsheetApp.openById(TRACKER_ID);
  const team = ss.getSheetByName('Team').getRange('A5:P14').getValues().filter(r => String(r[0]).trim());
  const oversight = n => team.some(t => String(t[0]).trim() === n && String(t[15]).trim() === 'Oversight');
  const titlesCache = {};
  let n = 0;
  todo.forEach(x => {
    const row = x.row;
    const tgt = trackerTarget_(row[COL.client - 1], row[COL.project - 1]);
    const tab = tr.getSheetByName(tgt.tab) || tr.getSheetByName('Ad-hoc');
    if (!tab) return;
    const when = row[COL.stageSince - 1];
    const year = Number(Utilities.formatDate(when, tz, 'yyyy')), q = 'Q' + (Math.floor((Number(Utilities.formatDate(when, tz, 'M')) - 1) / 3) + 1);
    const title = String(row[COL.deliverable - 1]).replace(/\s*\(projected\)\s*$/, '').trim();
    // Already on the tracker (same title, same year)? Mark it logged and move on
    const key = tab.getName();
    if (!titlesCache[key]) titlesCache[key] = tab.getLastRow() > 1 ? tab.getRange(2, 2, tab.getLastRow() - 1, 5).getValues().map(v => String(v[0]) + '|' + String(v[4]).trim().toLowerCase()) : [];
    const people = [row[COL.owner - 1]].concat(String(row[COL.assigned - 1]).split(/,\s*/).map(matchTeamName_)).filter(String);
    const lead = people.find(p => team.some(t => String(t[0]).trim() === p) && !oversight(p)) || people[0] || '';
    const ext = extOf(row);
    if (titlesCache[key].indexOf(year + '|' + title.toLowerCase()) === -1) {
      const typeName = String(row[COL.type - 1] || '');
      const type = (TRACKER_TYPES.find(t => t[0].test(typeName + ' ' + title)) || [null, 'Ad-hoc'])[1];
      const link = String(row[COL.link - 1] || '');
      const note = 'Auto-logged by the Insights planner (' + row[COL.id - 1] + ')' + (link ? ' from Basecamp: ' + link : '') + '. Add the report and dataset links.';
      tab.appendRow([q, year, tgt.client, type, lead, title, '', '', note, '', '', ext, '', '']);
      titlesCache[key].push(year + '|' + title.toLowerCase());
      logActivity_('tracker', row[COL.id - 1], title, 'Planner', 'Logged to the tracker (' + tab.getName() + ' tab)' + (ext ? ' with ' + ext : '') + '. Add the report link there.', link);
      n++;
    }
    sheet.getRange(x.r, COL.trackerLogged).setValue(new Date());
  });
  return n;
}

function summerFridayRange_(year) {
  const memorial = new Date(year, 4, 31); while (memorial.getDay() !== 1) memorial.setDate(memorial.getDate() - 1);
  const labor = new Date(year, 8, 1); while (labor.getDay() !== 1) labor.setDate(labor.getDate() + 1);
  return { s: new Date(year, 4, memorial.getDate() - 3), e: new Date(year, 8, labor.getDate() - 3) };
}
function isSummerFriday_(x) {
  if (x.getDay() !== 5) return false;
  const r = summerFridayRange_(x.getFullYear());
  return x >= r.s && x <= r.e;
}

// The dashboard's load model on the server: role split, custom splits, WBS work phases, leave cover
// Split text: "Leana L. 4 @dev (competitive set); Tayler M. 2 @rev; oversight 10%" (same rules as the dashboard)
function parseSplit_(text, match) {
  if (!text) return null;
  const out = { alloc: {}, win: {}, owns: {}, ov: null };
  String(text).split(/[;\n]+/).forEach(part => {
    part = part.trim(); if (!part) return;
    let m = part.match(/^oversight\s*[:=]?\s*(\d+(?:\.\d+)?)\s*%?$/i);
    if (m) { out.ov = Number(m[1]) / 100; return; }
    m = part.match(/^(.*?)[\s:=]+(\d+(?:\.\d+)?)\s*(?:d|days?)?\s*(?:@\s*([a-z+ ]+?))?\s*(?:\((.*)\))?\s*$/i);
    if (!m) return;
    const nm = match(m[1].trim());
    if (!nm || !(+m[2] > 0)) return;
    out.alloc[nm] = (out.alloc[nm] || 0) + Number(m[2]);
    const w = String(m[3] || '').trim().toLowerCase();
    out.win[nm] = /^dev/.test(w) ? 'dev' : /^rev/.test(w) ? 'rev' : 'all';
    if (m[4]) out.owns[nm] = m[4].trim();
  });
  return Object.keys(out.alloc).length || out.ov !== null ? out : null;
}

function teamModel_(d) {
  const P = s => { if (!s) return null; const a = String(s).split('-').map(Number); return new Date(a[0], a[1] - 1, a[2]); };
  const add = (x, n) => new Date(x.getFullYear(), x.getMonth(), x.getDate() + n);
  const K = x => x.getFullYear() + '-' + x.getMonth() + '-' + x.getDate();
  const HOL = {}; d.holidays.forEach(h => { const x = P(h.date); if (x) HOL[K(x)] = 1; });
  const isWd = x => x.getDay() > 0 && x.getDay() < 6 && !HOL[K(x)];
  const wd = (a, b) => { let n = 0; for (let x = a; x <= b; x = add(x, 1)) if (isWd(x)) n++; return n; };
  const maxD = (a, b) => a > b ? a : b, minD = (a, b) => a < b ? a : b;
  const T = P(d.today);

  const roster = d.team.map(p => p.name);
  const match = nm => roster.find(r => r === nm) || roster.find(r => r.split(' ')[0] === String(nm).split(' ')[0] && (r.split(' ')[1] || '').charAt(0) === (String(nm).split(' ')[1] || '').charAt(0));
  const isOv = nm => (d.team.find(p => p.name === nm) || {}).projectRole === 'Oversight';
  const coverOf = (note, self) => {
    const m = String(note || '').match(/cover(?:age|ed\s*by|ing)?\s*[:\-–]?\s*([^.;,\n]+)/i); if (!m) return '';
    const t = m[1].trim().toLowerCase();
    const hit = d.team.find(p => p.name !== self && (t.indexOf(p.name.toLowerCase()) === 0 || t.split(/\s+/)[0] === p.name.split(' ')[0].toLowerCase()));
    return hit ? hit.name : '';
  };
  const covers = d.timeOff.map(o => ({ from: o.name, to: coverOf(o.note, o.name), s: P(o.start), e: P(o.end) })).filter(c => c.to && c.s && c.e);
  const CLOSED = ['Delivered', 'On hold', 'Needs triage', 'Dismissed'];

  let items = d.items.filter(i => CLOSED.indexOf(i.stage) === -1).map(i => {
    const x = { src: i, s: P(i.workStart) || P(i.start) || P(i.due), e: P(i.due), est: i.est, team: [], alloc: {} };
    [i.owner].concat(String(i.assigned).split(/,\s*/).map(match)).forEach(nm => { if (nm && x.team.indexOf(nm) === -1) x.team.push(nm); });
    // Same rules as the dashboard: the estimate is hands-on days (a Split sets them per person, with an optional
    // window, @dev or @rev); oversight is added on top at the deliverable's, account's or default %
    const sp = parseSplit_(i.split, match);
    x.win = sp ? sp.win : {};
    const accOv = (d.accountOv || {})[i.account];
    const ovPct = sp && sp.ov !== null ? sp.ov : accOv !== undefined ? accOv : (d.ovDefault !== undefined ? d.ovDefault : 0.25);
    let hands, overs, handsEst;
    if (sp && Object.keys(sp.alloc).length) {
      Object.keys(sp.alloc).forEach(nm => { if (x.team.indexOf(nm) === -1) x.team.push(nm); x.alloc[nm] = sp.alloc[nm]; });
      hands = Object.keys(sp.alloc); handsEst = hands.reduce((t, nm) => t + sp.alloc[nm], 0);
      overs = x.team.filter(nm => isOv(nm) && hands.indexOf(nm) === -1);
    } else {
      hands = x.team.filter(nm => !isOv(nm)); overs = x.team.filter(isOv);
      if (!hands.length) { hands = overs; overs = []; }
      handsEst = Number(x.est) || 0;
      hands.forEach(nm => { x.alloc[nm] = handsEst / hands.length; });
    }
    overs.forEach(nm => { x.alloc[nm] = handsEst * ovPct / overs.length; x.win[nm] = 'all'; });
    x.est = handsEst + (overs.length ? handsEst * ovPct : 0);
    x.team = x.team.filter(nm => x.alloc[nm] > 0);
    x.hasWin = Object.keys(x.win).some(nm => x.win[nm] !== 'all');
    let w = null;
    for (let k = (i.replies || []).length - 1; k >= 0 && !w; k--) w = parseWbs_(i.replies[k].text, x.e || T);
    if (!w) w = parseWbs_(i.brief, x.e || T);
    if (w && x.est) {
      const work = w.filter(st => st.kind === 'work'), days = work.reduce((t, st) => t + wd(st.s, st.e), 0);
      x.s = w[0].s; x.e = maxD(x.e || w[0].e, w.reduce((m, st) => maxD(m, st.e), w[0].e));
      if (days) { x.work = work; x.rate = x.est / days; }
    }
    if (!x.rate) x.rate = x.s && x.e && x.est ? x.est / Math.max(1, wd(x.s, x.e)) : 0;
    // Development runs up to internal review (from the WBS; else the last quarter of working days is review and revisions)
    if (x.s && x.e) {
      const ir = w ? w.find(st => /internal\s*review|\bIR\b/i.test(st.label)) : null;
      let rs = ir ? ir.s : null;
      if (!rs) { const n = wd(x.s, x.e), k = Math.max(1, Math.ceil(n * 0.75)); let c = 0; for (let y = x.s; y <= x.e; y = add(y, 1)) if (isWd(y) && ++c === k + 1) { rs = y; break; } }
      if (rs && rs > x.s && rs <= x.e) { x.rwS = rs; let y = add(rs, -1); while (y > x.s && !isWd(y)) y = add(y, -1); x.devE = y; }
    }
    // In review: only a revision allowance is left, spread from today to the due date. Paused stages book nothing.
    const RV = d.review || { pct: 0.1, min: 0.1, max: 0.5, stages: ['Internal review', 'Client review', 'Revisions'], paused: ['Blocked', 'On hold'] };
    x.paused = RV.paused.indexOf(i.stage) !== -1;
    const pastIR = w && ['Not started', 'In progress'].indexOf(i.stage) !== -1 && w.some(st => /internal\s*review|\bIR\b/i.test(st.label) && st.e < T);
    if (x.est && (RV.stages.indexOf(i.stage) !== -1 || pastIR)) {
      x.revAllow = Math.min(RV.max, Math.max(RV.min, x.est * RV.pct));
      x.revS = T; x.revE = x.e && x.e > T ? x.e : T;
    }
    return x;
  });
  const all = items.slice();
  items = items.filter(x => x.rate && x.s && x.e);

  const load = (x, a, b) => x.paused ? 0 : x.revAllow !== undefined
    ? (x.revE >= a && x.revS <= b ? wd(maxD(x.revS, a), minD(x.revE, b)) * x.revAllow / Math.max(1, wd(x.revS, x.revE)) : 0)
    : x.work ? x.work.reduce((t, st) => t + (st.e >= a && st.s <= b ? wd(maxD(st.s, a), minD(st.e, b)) : 0), 0) * x.rate : (x.e >= a && x.s <= b ? wd(maxD(x.s, a), minD(x.e, b)) * x.rate : 0);
  const share = (x, nm) => x.est ? (x.alloc[nm] || 0) / x.est : 0;
  // One person's days between a and b, honoring their window (whole project, development, or review and revisions)
  const winOf = (x, nm) => (x.win && x.win[nm]) || 'all';
  const pload = (x, nm, a, b) => {
    const wn = winOf(x, nm);
    if (x.paused) return 0;
    if (x.revAllow !== undefined && x.hasWin) {
      if (wn === 'dev') return 0;
      const pool = x.team.filter(n => winOf(x, n) !== 'dev').reduce((t, n) => t + (x.alloc[n] || 0), 0);
      return pool ? load(x, a, b) * (x.alloc[nm] || 0) / pool : 0;
    }
    if (wn === 'all' || !x.est) return load(x, a, b) * share(x, nm);
    const ws = wn === 'dev' ? x.s : (x.rwS || x.s), we = wn === 'dev' ? (x.devE || x.e) : x.e;
    if (we < a || ws > b || we < ws) return 0;
    return (x.alloc[nm] || 0) * wd(maxD(ws, a), minD(we, b)) / Math.max(1, wd(ws, we));
  };
  // Booked days for one person between a and b
  const LIKE = { 'Likely': 0.9, 'Maybe': 0.5, 'Long shot': 0.2 };
  const booked = (nm, a, b, opts) => {
    const away = covers.filter(c => c.from === nm), inbound = covers.filter(c => c.to === nm);
    let t = 0;
    items.forEach(x => {
      // Tentative work counts in full (same as the dashboard's Overview) unless opts.weighted or opts.confirmedOnly
      const w = x.src.booking !== 'Tentative' ? 1 : opts && opts.confirmedOnly ? 0 : opts && opts.weighted ? (LIKE[x.src.likelihood] || 0.9) : 1;
      if (!w) return;
      if (x.team.indexOf(nm) !== -1) {
        let days = pload(x, nm, a, b);
        away.forEach(c => { if (c.e >= a && c.s <= b) days -= pload(x, nm, maxD(a, c.s), minD(b, c.e)); });
        t += Math.max(0, days) * w;
      }
      inbound.forEach(c => { if (x.team.indexOf(c.from) !== -1 && c.e >= a && c.s <= b) t += pload(x, c.from, maxD(a, c.s), minD(b, c.e)) * w; });
    });
    return t;
  };
  // Working days a person is in between a and b: minus time off (half days count 0.5) and Summer Friday afternoons
  const daysIn = (nm, a, b) => {
    const p = d.team.find(x => x.name === nm); if (!p) return 0;
    const s = P(p.start) ? maxD(a, P(p.start)) : a, e = P(p.last) ? minD(b, P(p.last)) : b;
    if (e < s) return 0;
    let n = 0;
    for (let x = s; x <= e; x = add(x, 1)) {
      if (!isWd(x)) continue;
      let off = 0;
      d.timeOff.forEach(o => { if (o.name === nm && P(o.start) <= x && P(o.end) >= x) off += o.half ? 0.5 : 1; });
      if (d.summerFridays && isSummerFriday_(x)) off += 0.5;
      n += Math.max(0, 1 - off);
    }
    return n;
  };
  // Available days: days in, times billable %
  const avail = (nm, a, b) => {
    const p = d.team.find(x => x.name === nm); if (!p) return 0;
    const bill = p.bill > 1 ? p.bill / 100 : p.bill;
    return daysIn(nm, a, b) * bill;
  };
  return { d, T, P, add, isWd, wd, items, all, booked, avail, daysIn, covers };
}
