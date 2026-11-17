/**
 * Suds Construction Monitoring — backend (Google Apps Script + Google Sheets + Drive)
 *
 * The page is published by Cloudflare (folder public/). Every button on that page
 * sends { fn, args } to this web app (doPost). Data lives in the Google Sheet this
 * script is attached to; photos, videos and documents go to a Drive folder.
 *
 * First time: run setup() once from the editor (see README).
 */

// ───────────────────────── Settings ─────────────────────────
var TRADES = ['General Contractor', 'CCTV', 'WiFi / Network', 'Fire Sprinkler', 'Smoke Detector / Fire Alarm',
              'LPG Line', 'Aircon', 'Glass & Aluminum', 'Signage', 'Others'];

var DEFAULT_MILESTONES = {
  'General Contractor': ['Mobilization & site prep', 'Demolition / hacking', 'Electrical & plumbing rough-in',
                         'Ceiling & walls', 'Flooring & tiles', 'Painting & finishes', 'Fixtures & fit-out', 'Final cleaning & punch list'],
  'CCTV': ['Cabling', 'Camera & NVR installation', 'Configuration & remote view', 'Testing & handover'],
  'WiFi / Network': ['Cabling', 'Router / AP installation', 'Configuration', 'Speed test & handover'],
  'Fire Sprinkler': ['Pipe installation', 'Sprinkler heads', 'Pressure test', 'Test certificate'],
  'Smoke Detector / Fire Alarm': ['Wiring', 'Device installation', 'Panel programming', 'Testing & certificate'],
  'LPG Line': ['Pipe routing', 'Regulator & valves', 'Leak / pressure test', 'Test certificate'],
  'Aircon': ['Delivery of units', 'Piping & bracket installation', 'Unit installation', 'Test run & handover'],
  'Glass & Aluminum': ['Measurement & fabrication', 'Frame installation', 'Glass installation', 'Sealant & cleanup'],
  'Signage': ['Design approval', 'Fabrication', 'Installation', 'Lighting test'],
  'Others': ['Start of work', 'Main work', 'Completion']
};

var TURNOVER_ITEMS = {
  _all: ['Punch list cleared', 'As-built / completion photos', 'Warranty certificate'],
  'General Contractor': ['Final cleaning done', 'As-built plans'],
  'Fire Sprinkler': ['Pressure test certificate'],
  'Smoke Detector / Fire Alarm': ['Test certificate (for BFP)'],
  'LPG Line': ['Leak test certificate'],
  'Aircon': ['Unit manuals & serial numbers'],
  'CCTV': ['Login credentials handed to owner'],
  'WiFi / Network': ['Login credentials handed to owner']
};

var SCHEMA = {
  Owners:        ['OwnerID', 'Name', 'Email', 'Mobile', 'OdooKey', 'Active', 'CreatedAt'],
  Suppliers:     ['SupplierID', 'Company', 'ContactPerson', 'Mobile', 'Email', 'TIN', 'Address', 'Trades', 'PinHash',
                  'Status', 'Docs', 'Remarks', 'CreatedAt', 'AccreditedAt'],
  Projects:      ['ProjectID', 'OwnerID', 'Name', 'Branch', 'Address', 'StartDate', 'TargetOpening', 'Status', 'Notes', 'CreatedAt', 'Plans',
                  'MapLink', 'SiteContact', 'SiteMobile'],
  Packages:      ['PackageID', 'ProjectID', 'Trade', 'Scope', 'Status', 'AwardedSupplierID', 'AwardedQuoteID', 'ContractAmount',
                  'DPRequired', 'RetentionPct', 'WarrantyMonths', 'AwardedAt', 'TurnedOverAt', 'WarrantyEnd', 'CreatedAt',
                  'UpdateFreq', 'UpdateDay', 'UpdateDueDate', 'UpdateNote'],
  Invitations:   ['InviteID', 'PackageID', 'ProjectID', 'SupplierID', 'Brief', 'Attachments', 'Deadline', 'Status',
                  'SentAt', 'ViewedAt', 'RespondedAt', 'DeclineReason', 'SupplierEmail', 'SupplierName'],
  Quotes:        ['QuoteID', 'InviteID', 'PackageID', 'ProjectID', 'SupplierID', 'Rev', 'Lines', 'Total', 'DurationDays',
                  'PaymentTerms', 'WarrantyMonths', 'Notes', 'Attachments', 'IsFinal', 'Status', 'OwnerRemarks',
                  'CounterAmount', 'AwardStatus', 'NotifiedAt', 'CreatedAt', 'UpdatedAt'],
  Payments:      ['PaymentID', 'PackageID', 'ProjectID', 'SupplierID', 'Type', 'RequestID', 'Amount', 'PaidDate', 'Method',
                  'RefNo', 'Proof', 'Notes', 'Status', 'RecordedAt', 'ConfirmedAt'],
  PayRequests:   ['RequestID', 'PackageID', 'ProjectID', 'SupplierID', 'Type', 'Amount', 'Milestone', 'Notes',
                  'Attachments', 'Status', 'RequestedAt', 'DecidedAt', 'RejectReason'],
  Mobilization:  ['PackageID', 'ProjectID', 'SupplierID', 'MobDate', 'DurationDays', 'TargetEnd', 'Workers', 'Foreman',
                  'ForemanMobile', 'Materials', 'Notes', 'SubmittedAt'],
  Milestones:    ['MilestoneID', 'PackageID', 'ProjectID', 'Seq', 'Name', 'Weight', 'Done', 'DoneAt', 'Photos'],
  Updates:       ['UpdateID', 'PackageID', 'ProjectID', 'SupplierID', 'WeekOf', 'Percent', 'WorkDone', 'NextPlan',
                  'Manpower', 'Delays', 'Media', 'VideoLink', 'SubmittedAt'],
  Concerns:      ['ConcernID', 'PackageID', 'ProjectID', 'RaisedBy', 'RaisedByName', 'Category', 'Description', 'Media',
                  'Status', 'Response', 'RespondedBy', 'CreatedAt', 'UpdatedAt'],
  ChangeOrders:  ['COID', 'PackageID', 'ProjectID', 'SupplierID', 'Description', 'Amount', 'AddDays', 'Reason',
                  'Attachments', 'Status', 'OwnerRemarks', 'RequestedAt', 'DecidedAt'],
  Turnover:      ['PackageID', 'ProjectID', 'SupplierID', 'Checklist', 'Docs', 'Notes', 'SubmittedAt'],
  EmailLog:      ['At', 'To', 'Subject', 'Ref', 'Result']
};

var JSON_FIELDS = { Docs: 1, Attachments: 1, Lines: 1, Proof: 1, Photos: 1, Media: 1, Checklist: 1, Trades: 1, Plans: 1 };

// ───────────────────────── Web entry points ─────────────────────────
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'Suds Construction Monitoring API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var fn = String(req.fn || '');
    if (!/^api_/.test(fn) || typeof this[fn] !== 'function') throw new Error('Unknown action: ' + fn);
    var data = this[fn].apply(this, req.args || []);
    out = { ok: true, data: data };
    // Send the refreshed data back in the same response (saves a second trip).
    if (req.withData && fn !== 'api_data') {
      var tok = data && data.token ? data.token : (req.args || [])[0];
      try { out.fresh = api_data(tok); } catch (e2) { }
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────── One-time setup ─────────────────────────
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p = props_();
  p.setProperty('SHEET_ID', ss.getId());
  Object.keys(SCHEMA).forEach(function (n) { sh_(n); });
  var s1 = ss.getSheetByName('Sheet1'); if (s1 && ss.getSheets().length > 1 && s1.getLastRow() === 0) ss.deleteSheet(s1);
  if (!p.getProperty('ROOT_FOLDER_ID')) {
    p.setProperty('ROOT_FOLDER_ID', DriveApp.createFolder('Suds Construction Monitoring – Files').getId());
  }
  if (!p.getProperty('ADMIN_PIN')) p.setProperty('ADMIN_PIN', String(Math.floor(100000 + Math.random() * 900000)));
  if (!p.getProperty('ADMIN_EMAIL')) p.setProperty('ADMIN_EMAIL', Session.getEffectiveUser().getEmail());
  if (!p.getProperty('APP_URL')) p.setProperty('APP_URL', 'https://suds-construction.pages.dev');
  if (!p.getProperty('OWNER_URL')) p.setProperty('OWNER_URL', 'https://YOUR-COMPANY.odoo.com/my');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyJobs') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyJobs').timeBased().everyDays(1).atHour(8).create();
  Logger.log('Setup done.\nADMIN PIN: ' + p.getProperty('ADMIN_PIN') +
             '\nADMIN EMAIL: ' + p.getProperty('ADMIN_EMAIL') +
             '\nNow set APP_URL in Project Settings → Script properties to your Cloudflare link.');
}

// ───────────────────────── Sheet helpers ─────────────────────────
var _rows = {};
function props_() { return PropertiesService.getScriptProperties(); }
var _ss = null, _sheets = {};
function ss_() {
  if (_ss) return _ss;
  var id = props_().getProperty('SHEET_ID');
  _ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  return _ss;
}
var _heads = {};
function sh_(name) {
  if (_sheets[name]) return _sheets[name];
  var ss = ss_(), s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name]]).setFontWeight('bold').setBackground('#FFF0E3');
    s.setFrozenRows(1);
  }
  _sheets[name] = s;
  return s;
}
// Real column order of the sheet (adds any new columns from SCHEMA at the end, so old sheets keep working).
function head_(name) {
  if (!_heads[name]) rows_(name);
  return _heads[name];
}
function tz_() { return Session.getScriptTimeZone() || 'Asia/Manila'; }
function now_() { return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd HH:mm"); }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function id_(prefix) { return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase(); }

function norm_(k, v) {
  if (v instanceof Date) {
    var hasTime = v.getHours() || v.getMinutes();
    return Utilities.formatDate(v, tz_(), hasTime ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
  }
  if (JSON_FIELDS[k]) { if (v === '' || v == null) return []; try { return JSON.parse(v); } catch (e) { return []; } }
  return v;
}
function ser_(k, v) {
  if (v === undefined || v === null) return '';
  if (JSON_FIELDS[k] || (typeof v === 'object')) return JSON.stringify(v);
  if (typeof v === 'boolean') return v;
  if (typeof v === 'string' && /^[+0-9][0-9 .\-\/]*$/.test(v) && !/^-?\d+(\.\d+)?$/.test(v) || /^0\d+$/.test(v)) return "'" + v;
  return v;
}
function rows_(name) {
  if (_rows[name]) return _rows[name];
  var s = sh_(name), vals = s.getDataRange().getValues();
  var head = (vals.shift() || []).map(String);
  var missing = SCHEMA[name].filter(function (k) { return head.indexOf(k) < 0; });
  if (missing.length) {   // new columns from SCHEMA are added at the end, so old sheets keep working
    s.getRange(1, head.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#FFF0E3');
    head = head.concat(missing);
  }
  _heads[name] = head;
  var out = [];
  vals.forEach(function (r, i) {
    if (r[0] === '' || r[0] === null) return;
    var o = { _row: i + 2 };
    head.forEach(function (h, j) { if (h) o[h] = norm_(h, r[j]); });
    out.push(o);
  });
  _rows[name] = out;
  return out;
}
function find_(name, id) {
  var key = SCHEMA[name][0];
  for (var i = 0, a = rows_(name); i < a.length; i++) if (String(a[i][key]) === String(id)) return a[i];
  return null;
}
function insert_(name, obj) {
  var s = sh_(name);
  s.appendRow(head_(name).map(function (k) { return ser_(k, obj[k]); }));
  delete _rows[name];
  return obj;
}
function update_(name, id, patch) {
  var r = find_(name, id);
  if (!r) throw new Error(name + ' record not found: ' + id);
  var s = sh_(name), head = head_(name);
  Object.keys(patch).forEach(function (k) {
    var c = head.indexOf(k);
    if (c < 0) throw new Error('Unknown field ' + k);
    s.getRange(r._row, c + 1).setValue(ser_(k, patch[k]));
    r[k] = patch[k];
  });
  delete _rows[name];
  return r;
}
function clean_(o) { var c = {}; Object.keys(o).forEach(function (k) { if (k !== '_row' && k !== 'PinHash' && k !== 'OdooKey') c[k] = o[k]; }); return c; }
function num_(v) { var n = Number(String(v).replace(/[^0-9.\-]/g, '')); return isFinite(n) ? n : 0; }
function str_(v, max) { return String(v == null ? '' : v).trim().slice(0, max || 2000); }
function withLock_(fn) {
  var l = LockService.getScriptLock(); l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'suds-cm|' + s)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

// ───────────────────────── Sessions & roles ─────────────────────────
function newSession_(me) {
  var t = 'S' + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put(t, JSON.stringify(me), 6 * 3600);
  return t;
}
function auth_(token, roles) {
  var raw = token && CacheService.getScriptCache().get(token);
  if (!raw) throw new Error('SESSION_EXPIRED');
  var me = JSON.parse(raw);
  CacheService.getScriptCache().put(token, raw, 6 * 3600);
  if (roles && roles.indexOf(me.role) < 0) throw new Error('Not allowed for your account.');
  return me;
}
function project_(id) { var p = find_('Projects', id); if (!p) throw new Error('Project not found.'); return p; }
function pkg_(id) { var p = find_('Packages', id); if (!p) throw new Error('Package not found.'); return p; }
function ownsProject_(me, projectId) {
  if (me.role === 'admin') return true;
  if (me.role === 'owner') return String(project_(projectId).OwnerID) === String(me.id);
  return false;
}
function needOwner_(me, projectId) { if (!ownsProject_(me, projectId)) throw new Error('Not allowed for this project.'); }
function needAwardee_(me, pk) {
  if (me.role === 'admin') return;
  if (me.role !== 'supplier' || String(pk.AwardedSupplierID) !== String(me.id)) throw new Error('Not allowed for this package.');
}
function canSeePackage_(me, pk) {
  if (ownsProject_(me, pk.ProjectID)) return true;
  if (me.role === 'supplier') {
    if (String(pk.AwardedSupplierID) === String(me.id)) return true;
    return rows_('Invitations').some(function (i) { return i.PackageID === pk.PackageID && String(i.SupplierID) === String(me.id); });
  }
  return false;
}

function api_adminLogin(pin) {
  if (!pin || String(pin) !== String(props_().getProperty('ADMIN_PIN'))) throw new Error('Wrong admin PIN.');
  var me = { role: 'admin', id: 'ADMIN', name: 'SPFC Admin', email: props_().getProperty('ADMIN_EMAIL') };
  return { token: newSession_(me), me: me };
}
function api_ownerLogin(key) {
  key = str_(key, 200);
  var o = key && rows_('Owners').filter(function (r) { return r.OdooKey && String(r.OdooKey) === key; })[0];
  if (!o || String(o.Active).toUpperCase() === 'FALSE' || o.Active === false)
    throw new Error('Your construction monitoring access is not set up yet. Please contact SPFC.');
  var me = { role: 'owner', id: o.OwnerID, name: o.Name, email: o.Email };
  return { token: newSession_(me), me: me };
}
function api_supplierLogin(email, pin) {
  email = str_(email, 200).toLowerCase();
  var s = rows_('Suppliers').filter(function (r) { return String(r.Email).toLowerCase() === email; })[0];
  if (!s || s.PinHash !== hash_(str_(pin, 20))) throw new Error('Wrong email or PIN.');
  linkInvites_(s);
  var me = { role: 'supplier', id: s.SupplierID, name: s.Company, email: s.Email };
  return { token: newSession_(me), me: me };
}
function api_supplierRegister(d) {
  return withLock_(function () {
    var email = str_(d.Email, 200).toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Please enter a valid email.');
    if (!str_(d.Company)) throw new Error('Company name is required.');
    if (rows_('Suppliers').some(function (r) { return String(r.Email).toLowerCase() === email; }))
      throw new Error('This email is already registered. Use "Forgot PIN" if you lost your PIN.');
    var pin = String(Math.floor(100000 + Math.random() * 900000));
    var s = insert_('Suppliers', {
      SupplierID: id_('SUP'), Company: str_(d.Company, 200), ContactPerson: str_(d.ContactPerson, 200),
      Mobile: str_(d.Mobile, 40), Email: email, TIN: str_(d.TIN, 40), Address: str_(d.Address, 500),
      Trades: (d.Trades || []).filter(function (t) { return TRADES.indexOf(t) >= 0; }),
      PinHash: hash_(pin), Status: 'Pending', Docs: [], Remarks: '', CreatedAt: now_(), AccreditedAt: ''
    });
    mail_(email, 'Your Suds supplier account – PIN', box_('Welcome, ' + esc_(s.Company) + '!',
      '<p>Your supplier registration was received and is <b>pending accreditation</b> by SPFC.</p>' +
      '<p>Your login PIN: <b style="font-size:20px;letter-spacing:3px">' + pin + '</b></p>' +
      '<p>Log in with your email and this PIN to upload your documents (business permit, DTI/SEC, BIR 2303).</p>'), s.SupplierID);
    mail_(adminEmail_(), 'New supplier registration: ' + s.Company, box_('New supplier registration',
      '<p><b>' + esc_(s.Company) + '</b> (' + esc_((s.Trades || []).join(', ')) + ') registered and is waiting for accreditation.</p>'), s.SupplierID);
    linkInvites_(s);
    var me = { role: 'supplier', id: s.SupplierID, name: s.Company, email: s.Email };
    return { token: newSession_(me), me: me, pin: pin };
  });
}
function api_supplierResetPin(email) {
  email = str_(email, 200).toLowerCase();
  var s = rows_('Suppliers').filter(function (r) { return String(r.Email).toLowerCase() === email; })[0];
  if (s) {
    var pin = String(Math.floor(100000 + Math.random() * 900000));
    update_('Suppliers', s.SupplierID, { PinHash: hash_(pin) });
    mail_(email, 'Your new Suds supplier PIN', box_('New PIN', '<p>Your new login PIN: <b style="font-size:20px;letter-spacing:3px">' + pin + '</b></p>'), s.SupplierID);
  }
  return { sent: true };
}

// ───────────────────────── Load everything the user may see ─────────────────────────
function api_data(token) {
  var me = auth_(token);
  var all = {};
  ['Owners', 'Suppliers', 'Projects', 'Packages', 'Invitations', 'Quotes', 'Payments', 'PayRequests', 'Mobilization',
   'Milestones', 'Updates', 'Concerns', 'ChangeOrders', 'Turnover'].forEach(function (n) { all[n] = rows_(n).map(clean_); });

  var out = { me: me, trades: TRADES, turnoverItems: TURNOVER_ITEMS, today: today_() };
  if (me.role === 'admin') {
    Object.keys(all).forEach(function (k) { out[k] = all[k]; });
    out.Owners = rows_('Owners').map(function (o) { var c = clean_(o); c.OdooKey = o.OdooKey; return c; });
    return out;
  }
  var byProj, byPkg, supplierBasic = function (s) {
    return { SupplierID: s.SupplierID, Company: s.Company, ContactPerson: s.ContactPerson, Mobile: s.Mobile,
             Email: s.Email, Trades: s.Trades, Status: s.Status, AccreditedAt: s.AccreditedAt, Address: s.Address };
  };
  if (me.role === 'owner') {
    var pids = {};
    out.Projects = all.Projects.filter(function (p) { if (String(p.OwnerID) === String(me.id)) { pids[p.ProjectID] = 1; return true; } });
    byProj = function (r) { return pids[r.ProjectID]; };
    ['Packages', 'Invitations', 'Quotes', 'Payments', 'PayRequests', 'Mobilization', 'Milestones', 'Updates', 'Concerns',
     'ChangeOrders', 'Turnover'].forEach(function (k) { out[k] = all[k].filter(byProj); });
    var involved = {};
    out.Invitations.forEach(function (i) { involved[i.SupplierID] = 1; });
    out.Suppliers = all.Suppliers.filter(function (s) { return s.Status === 'Accredited' || involved[s.SupplierID]; }).map(supplierBasic);
    return out;
  }
  // supplier
  var mine = function (r) { return String(r.SupplierID) === String(me.id); };
  var self = all.Suppliers.filter(mine)[0];
  out.Profile = self;
  out.Invitations = all.Invitations.filter(mine);
  out.Quotes = all.Quotes.filter(mine);
  var seePkg = {}, awarded = {};
  out.Invitations.forEach(function (i) { seePkg[i.PackageID] = 1; });
  all.Packages.forEach(function (p) { if (String(p.AwardedSupplierID) === String(me.id)) { awarded[p.PackageID] = 1; seePkg[p.PackageID] = 1; } });
  out.Packages = all.Packages.filter(function (p) { return seePkg[p.PackageID]; }).map(function (p) {
    if (awarded[p.PackageID]) return p;
    return { PackageID: p.PackageID, ProjectID: p.ProjectID, Trade: p.Trade, Scope: p.Scope,
             Status: p.AwardedSupplierID ? 'Awarded to another contractor' : p.Status };
  });
  var seeProj = {}; out.Packages.forEach(function (p) { seeProj[p.ProjectID] = 1; });
  out.Projects = all.Projects.filter(function (p) { return seeProj[p.ProjectID]; }).map(function (p) {
    return { ProjectID: p.ProjectID, Name: p.Name, Branch: p.Branch, Address: p.Address, StartDate: p.StartDate, Status: p.Status,
             TargetOpening: p.TargetOpening, MapLink: p.MapLink, SiteContact: p.SiteContact, SiteMobile: p.SiteMobile };
  });
  byPkg = function (r) { return awarded[r.PackageID]; };
  ['Payments', 'PayRequests', 'Mobilization', 'Milestones', 'Updates', 'Concerns', 'ChangeOrders', 'Turnover']
    .forEach(function (k) { out[k] = all[k].filter(byPkg); });
  out.Suppliers = [];
  return out;
}

// ───────────────────────── File uploads (Drive) ─────────────────────────
function folderFor_(me, ctx) {
  ctx = ctx || {};
  var root = DriveApp.getFolderById(props_().getProperty('ROOT_FOLDER_ID'));
  var path;
  if (ctx.type === 'supplier') {
    var sid = me.role === 'supplier' ? me.id : ctx.supplierId;
    var s = find_('Suppliers', sid); if (!s) throw new Error('Supplier not found.');
    path = ['Suppliers', s.Company + ' (' + s.SupplierID + ')'];
  } else {
    var pk = ctx.packageId ? pkg_(ctx.packageId) : null;
    var pid = pk ? pk.ProjectID : ctx.projectId;
    var pr = project_(pid);
    if (pk ? !canSeePackage_(me, pk) : !ownsProject_(me, pid)) throw new Error('Not allowed to upload here.');
    path = ['Projects', (pr.Branch || pr.Name) + ' (' + pr.ProjectID + ')'];
    if (pk) path.push(pk.Trade + ' (' + pk.PackageID + ')');
    path.push(str_(ctx.kind || 'Files', 60).replace(/[\/\\]/g, '-'));
  }
  var f = root;
  path.forEach(function (name) {
    var it = f.getFoldersByName(name);
    f = it.hasNext() ? it.next() : f.createFolder(name);
  });
  return f;
}
function fileInfo_(file) {
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* domain may block */ }
  return { id: file.getId(), name: file.getName(), mime: file.getMimeType(), url: file.getUrl() };
}
function api_uploadSmall(token, ctx, name, mime, b64) {
  var me = auth_(token);
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', str_(name, 200) || 'file');
  return fileInfo_(folderFor_(me, ctx).createFile(blob));
}
function api_uploadStart(token, ctx, name, mime, size) {
  var me = auth_(token);
  if (size > 300 * 1024 * 1024) throw new Error('File is too big (max 300 MB). Please send a shorter video or a link.');
  var folder = folderFor_(me, ctx);
  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true', {
    method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': mime || 'application/octet-stream',
               'X-Upload-Content-Length': String(size) },
    payload: JSON.stringify({ name: str_(name, 200) || 'file', parents: [folder.getId()] })
  });
  if (res.getResponseCode() !== 200) throw new Error('Could not start upload (' + res.getResponseCode() + ').');
  var h = res.getHeaders(), loc = h.Location || h.location;
  var up = 'U' + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put(up, JSON.stringify({ loc: loc, size: size, mime: mime, uid: me.id }), 6 * 3600);
  return { token: up };
}
function api_uploadChunk(token, upToken, b64, pos) {
  var me = auth_(token);
  var c = JSON.parse(CacheService.getScriptCache().get(upToken) || 'null');
  if (!c || c.uid !== me.id) throw new Error('Upload expired. Please try again.');
  var bytes = Utilities.base64Decode(b64), end = pos + bytes.length - 1;
  var res = UrlFetchApp.fetch(c.loc, {
    method: 'put', contentType: c.mime || 'application/octet-stream', payload: bytes, muteHttpExceptions: true,
    headers: { 'Content-Range': 'bytes ' + pos + '-' + end + '/' + c.size }
  });
  var code = res.getResponseCode();
  if (code === 308) return { done: false };
  if (code === 200 || code === 201) return { done: true, file: fileInfo_(DriveApp.getFileById(JSON.parse(res.getContentText()).id)) };
  throw new Error('Upload failed (' + code + ').');
}

// ───────────────────────── Admin ─────────────────────────
function api_setSupplierStatus(token, supplierId, status, remarks) {
  auth_(token, ['admin']);
  if (['Accredited', 'Rejected', 'Pending', 'Suspended'].indexOf(status) < 0) throw new Error('Bad status');
  var s = update_('Suppliers', supplierId, { Status: status, Remarks: str_(remarks, 1000),
                                             AccreditedAt: status === 'Accredited' ? now_() : '' });
  var body = status === 'Accredited'
    ? '<p>Good news! <b>' + esc_(s.Company) + '</b> is now an <b>accredited Suds supplier</b>. You will receive an email when an owner invites you to quote.</p>'
    : '<p>Your supplier account status is now <b>' + esc_(status) + '</b>.</p>' + (remarks ? '<p>Remarks: ' + esc_(remarks) + '</p>' : '');
  mail_(s.Email, 'Supplier accreditation: ' + status, box_('Accreditation update', body), supplierId);
  return clean_(s);
}
function api_saveSupplierProfile(token, d) {
  var me = auth_(token, ['supplier', 'admin']);
  var id = me.role === 'supplier' ? me.id : d.SupplierID;
  var patch = {};
  ['Company', 'ContactPerson', 'Mobile', 'TIN', 'Address'].forEach(function (k) { if (d[k] !== undefined) patch[k] = str_(d[k], 500); });
  if (d.Trades) patch.Trades = d.Trades.filter(function (t) { return TRADES.indexOf(t) >= 0; });
  if (d.Docs) patch.Docs = d.Docs;
  return clean_(update_('Suppliers', id, patch));
}
function api_saveOwner(token, d) {
  auth_(token, ['admin']);
  return withLock_(function () {
    var email = str_(d.Email, 200).toLowerCase();
    if (!str_(d.Name) || !email) throw new Error('Name and email are required.');
    if (d.OwnerID) {
      var u = update_('Owners', d.OwnerID, { Name: str_(d.Name, 200), Email: email, Mobile: str_(d.Mobile, 40), Active: d.Active !== false });
      return u;
    }
    return insert_('Owners', { OwnerID: id_('OWN'), Name: str_(d.Name, 200), Email: email, Mobile: str_(d.Mobile, 40),
                               OdooKey: newKey_(), Active: true, CreatedAt: now_() });
  });
}
function api_regenOwnerKey(token, ownerId) {
  auth_(token, ['admin']);
  return update_('Owners', ownerId, { OdooKey: newKey_() });
}
function newKey_() { return 'CM-' + Utilities.getUuid().replace(/-/g, '').toUpperCase(); }
function api_ownerLinkAsAdmin(token, ownerId) {
  auth_(token, ['admin']);
  var o = find_('Owners', ownerId);
  var me = { role: 'owner', id: o.OwnerID, name: o.Name, email: o.Email, viaAdmin: true };
  return { token: newSession_(me), me: me };
}

// ───────────────────────── Projects & packages ─────────────────────────
function api_saveProject(token, d) {
  var me = auth_(token, ['admin', 'owner']);
  return withLock_(function () {
    var base = { Name: str_(d.Name, 200), Branch: str_(d.Branch, 200), Address: str_(d.Address, 500),
                 StartDate: str_(d.StartDate, 10), TargetOpening: str_(d.TargetOpening, 10), Notes: str_(d.Notes, 2000),
                 MapLink: /^https?:\/\//i.test(str_(d.MapLink, 500)) ? str_(d.MapLink, 500) : '', SiteContact: str_(d.SiteContact, 200),
                 SiteMobile: str_(d.SiteMobile, 40) };
    if (!base.Name) throw new Error('Project name is required.');
    if (d.ProjectID) {
      needOwner_(me, d.ProjectID);
      if (me.role === 'admin' && d.OwnerID) base.OwnerID = d.OwnerID;
      if (d.Status) base.Status = d.Status;
      return clean_(update_('Projects', d.ProjectID, base));
    }
    base.ProjectID = id_('PRJ');
    base.OwnerID = me.role === 'owner' ? me.id : str_(d.OwnerID, 40);
    if (!base.OwnerID) throw new Error('Please choose the owner.');
    base.Status = 'Planning'; base.CreatedAt = now_();
    return insert_('Projects', base);
  });
}
function api_savePackage(token, d) {
  var me = auth_(token, ['admin', 'owner']);
  needOwner_(me, d.ProjectID);
  return withLock_(function () {
    if (TRADES.indexOf(d.Trade) < 0) throw new Error('Please choose a trade.');
    if (d.PackageID) {
      var pk = pkg_(d.PackageID);
      if (pk.ProjectID !== d.ProjectID) throw new Error('Mismatch.');
      return update_('Packages', d.PackageID, { Trade: d.Trade, Scope: str_(d.Scope, 3000) });
    }
    return insert_('Packages', { PackageID: id_('PKG'), ProjectID: d.ProjectID, Trade: d.Trade, Scope: str_(d.Scope, 3000),
      Status: 'Draft', AwardedSupplierID: '', AwardedQuoteID: '', ContractAmount: 0, DPRequired: false, RetentionPct: 0,
      WarrantyMonths: '', AwardedAt: '', TurnedOverAt: '', WarrantyEnd: '', CreatedAt: now_() });
  });
}

// ───────────────────────── Invitations ─────────────────────────
// Invite suppliers (registered or just an email) to quote one or more packages of a project.
// d = { packageIds:[], supplierIds:[], emails:[{email,name}], brief, deadline, attachments:[], planIds:[] (only these plans are shared) }
function api_inviteProject(token, projectId, d) {
  var me = auth_(token, ['admin', 'owner']);
  needOwner_(me, projectId);
  d = d || {};
  var pr = project_(projectId);
  var pkgs = (d.packageIds || []).map(pkg_).filter(function (p) { return p.ProjectID === projectId && !p.AwardedSupplierID; });
  if (!pkgs.length) throw new Error('Please tick at least one package to be quoted.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.deadline || '')) throw new Error('Please set the quote deadline.');
  return withLock_(function () {
    // who: registered suppliers + typed emails (matched to an existing supplier when the email is registered)
    var targets = {};
    (d.supplierIds || []).forEach(function (sid) {
      var s = find_('Suppliers', sid);
      if (s) targets[String(s.Email).toLowerCase()] = { email: String(s.Email).toLowerCase(), name: s.Company, sup: s };
    });
    (d.emails || []).forEach(function (e) {
      var email = str_(e.email, 200).toLowerCase();
      if (!email) return;
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Invalid email: ' + email);
      var s = rows_('Suppliers').filter(function (r) { return String(r.Email).toLowerCase() === email; })[0];
      targets[email] = { email: email, name: s ? s.Company : str_(e.name, 200), sup: s || null };
    });
    var list = Object.keys(targets).map(function (k) { return targets[k]; });
    if (!list.length) throw new Error('Please enter a supplier email or choose a supplier.');
    var pick = (d.planIds || []).map(String);
    var plans = (pr.Plans || []).filter(function (f) { return pick.indexOf(String(f.id)) >= 0; });
    var att = plans.concat(d.attachments || []);   // the invitation keeps exactly what was shared
    var made = [];
    list.forEach(function (t) {
      var newPk = [];
      pkgs.forEach(function (pk) {
        var dup = rows_('Invitations').filter(function (i) {
          return i.PackageID === pk.PackageID && (t.sup ? i.SupplierID === t.sup.SupplierID : String(i.SupplierEmail).toLowerCase() === t.email) &&
                 i.Status !== 'Declined' && i.Status !== 'Expired';
        })[0];
        if (dup) return;
        made.push(insert_('Invitations', { InviteID: id_('INV'), PackageID: pk.PackageID, ProjectID: projectId,
          SupplierID: t.sup ? t.sup.SupplierID : '', SupplierEmail: t.email, SupplierName: t.name || '',
          Brief: str_(d.brief, 4000), Attachments: att, Deadline: d.deadline, Status: 'Invited', SentAt: now_(),
          ViewedAt: '', RespondedAt: '', DeclineReason: '' }));
        if (pk.Status === 'Draft') update_('Packages', pk.PackageID, { Status: 'Quoting' });
        newPk.push(pk);
      });
      if (!newPk.length) return;
      var appUrl = props_().getProperty('APP_URL') || '#';
      var link = t.sup ? appUrl : appUrl + (appUrl.indexOf('?') < 0 ? '?' : '&') + 'invite=' + encodeURIComponent(t.email);
      var files = att;
      mail_(t.email, 'Invitation to quote: ' + (pr.Branch || pr.Name), box_('You are invited to quote',
        '<p>Hi ' + esc_(t.name || '') + ',</p><p>You are invited to submit a quotation for the following:</p>' +
        table_([['Project', pr.Name + (pr.Branch ? ' – ' + pr.Branch : '')], ['Location', pr.Address],
                ['Site contact', [pr.SiteContact, pr.SiteMobile].filter(String).join(' – ')],
                ['Package(s) to quote', newPk.map(function (p) { return p.Trade + (p.Scope ? ' – ' + p.Scope : ''); }).join('\n')],
                ['Brief', d.brief], ['Quote deadline', d.deadline]]) +
        (pr.MapLink ? '<p>📍 <a href="' + esc_(pr.MapLink) + '"><b>Open site location in Google Maps</b></a></p>' : '') +
        (files.length ? '<p><b>Plans & attachments:</b><br>' + files.map(function (f) {
          return '<a href="' + f.url + '">' + esc_(f.name) + '</a>'; }).join('<br>') + '</p>' : '') +
        (t.sup ? '<p>Log in to the supplier portal to view the details and submit your quotation.</p>'
               : '<p><b>To submit your quotation, register as a supplier using this email address (' + esc_(t.email) +
                 ')</b>. After registering you will see this invitation and all the plans in the portal.</p>') +
        '<p style="margin:18px 0"><a href="' + link + '" style="background:#FF6B00;color:#fff;padding:11px 20px;border-radius:8px;text-decoration:none;font-weight:600">' +
        (t.sup ? 'Open supplier portal' : 'Register & view invitation') + '</a></p>'), projectId);
    });
    if (!made.length) throw new Error('No new invitations were sent (already invited for these packages).');
    return made;
  });
}
// After a supplier registers or logs in, attach invitations that were sent to their email.
function linkInvites_(s) {
  var email = String(s.Email).toLowerCase();
  rows_('Invitations').forEach(function (i) {
    if (!i.SupplierID && String(i.SupplierEmail).toLowerCase() === email) update_('Invitations', i.InviteID, { SupplierID: s.SupplierID });
  });
}
function api_savePlans(token, projectId, plans) {
  var me = auth_(token, ['admin', 'owner']);
  needOwner_(me, projectId);
  return update_('Projects', projectId, { Plans: (plans || []).map(function (f) {
    return { id: str_(f.id, 100), name: str_(f.name, 200), mime: str_(f.mime, 100), url: str_(f.url, 500) }; }) });
}
function api_extendInvite(token, inviteId, deadline) {
  var me = auth_(token, ['admin', 'owner']);
  var inv = find_('Invitations', inviteId); needOwner_(me, inv.ProjectID);
  var s = find_('Suppliers', inv.SupplierID) || { Email: inv.SupplierEmail }, pk = pkg_(inv.PackageID);
  var r = update_('Invitations', inviteId, { Deadline: deadline, Status: inv.Status === 'Expired' ? 'Invited' : inv.Status });
  mail_(s.Email, 'Quote deadline extended: ' + pk.Trade, box_('Deadline extended', '<p>The quote deadline for <b>' + esc_(pk.Trade) +
        '</b> was moved to <b>' + esc_(deadline) + '</b>.</p>' + btn_('Open supplier portal')), inviteId);
  return r;
}
function api_inviteViewed(token, inviteId) {
  var me = auth_(token, ['supplier']);
  var inv = find_('Invitations', inviteId);
  if (!inv || inv.SupplierID !== me.id) throw new Error('Not found.');
  if (inv.Status === 'Invited') return update_('Invitations', inviteId, { Status: 'Viewed', ViewedAt: now_() });
  return inv;
}
function api_declineInvite(token, inviteId, reason) {
  var me = auth_(token, ['supplier']);
  var inv = find_('Invitations', inviteId);
  if (!inv || inv.SupplierID !== me.id) throw new Error('Not found.');
  var r = update_('Invitations', inviteId, { Status: 'Declined', RespondedAt: now_(), DeclineReason: str_(reason, 1000) });
  notifyOwner_(inv.ProjectID, 'Supplier declined: ' + me.name, '<p><b>' + esc_(me.name) + '</b> declined the invitation for <b>' +
    esc_(pkg_(inv.PackageID).Trade) + '</b>.</p><p>Reason: ' + esc_(reason || '—') + '</p>', inviteId);
  return r;
}

// ───────────────────────── Quotes & negotiation ─────────────────────────
function api_submitQuote(token, inviteId, q, isFinal) {
  var me = auth_(token, ['supplier']);
  return withLock_(function () {
    var inv = find_('Invitations', inviteId);
    if (!inv || inv.SupplierID !== me.id) throw new Error('Invitation not found.');
    if (inv.Status === 'Declined') throw new Error('You declined this invitation.');
    if (inv.Status === 'Expired' || (inv.Deadline && inv.Deadline < today_() && !hasQuote_(inviteId)))
      throw new Error('The quote deadline has passed. Ask the owner to extend it.');
    var pk = pkg_(inv.PackageID);
    if (pk.AwardedSupplierID) throw new Error('This package was already awarded.');
    var prev = latestQuote_(inviteId);
    if (prev && (prev.IsFinal === true || String(prev.IsFinal).toUpperCase() === 'TRUE')) throw new Error('Your final quote is already locked.');
    if (prev && prev.Status === 'Declined') throw new Error('Your quote was declined by the owner.');
    var lines = (q.Lines || []).map(function (l) {
      var qty = num_(l.qty), price = num_(l.price);
      return { desc: str_(l.desc, 300), qty: qty, unit: str_(l.unit, 20), price: price, amount: Math.round(qty * price * 100) / 100 };
    }).filter(function (l) { return l.desc; });
    if (!lines.length) throw new Error('Please add at least one line item.');
    var total = lines.reduce(function (a, l) { return a + l.amount; }, 0);
    if (prev) update_('Quotes', prev.QuoteID, { Status: 'Superseded', UpdatedAt: now_() });
    var nq = insert_('Quotes', { QuoteID: id_('QT'), InviteID: inviteId, PackageID: inv.PackageID, ProjectID: inv.ProjectID,
      SupplierID: me.id, Rev: prev ? Number(prev.Rev) + 1 : 0, Lines: lines, Total: Math.round(total * 100) / 100,
      DurationDays: num_(q.DurationDays), PaymentTerms: str_(q.PaymentTerms, 500), WarrantyMonths: num_(q.WarrantyMonths),
      Notes: str_(q.Notes, 3000), Attachments: q.Attachments || [], IsFinal: !!isFinal,
      Status: isFinal ? 'Final' : 'Submitted', OwnerRemarks: '', CounterAmount: '', AwardStatus: 'Pending', NotifiedAt: '',
      CreatedAt: now_(), UpdatedAt: now_() });
    update_('Invitations', inviteId, { Status: 'Quoted', RespondedAt: now_() });
    var label = isFinal ? 'FINAL quote' : (prev ? 'revised quote (Rev ' + nq.Rev + ')' : 'quote');
    notifyOwner_(inv.ProjectID, 'New ' + label + ' from ' + me.name + ' – ' + pk.Trade,
      '<p><b>' + esc_(me.name) + '</b> submitted a ' + label + ' for <b>' + esc_(pk.Trade) + '</b>.</p>' +
      table_([['Total', peso_(nq.Total)], ['Duration', nq.DurationDays + ' days'], ['Payment terms', nq.PaymentTerms]]) +
      btn_('Review in portal'), nq.QuoteID);
    return nq;
  });
}
function hasQuote_(inviteId) { return rows_('Quotes').some(function (q) { return q.InviteID === inviteId; }); }
function latestQuote_(inviteId) {
  var l = rows_('Quotes').filter(function (q) { return q.InviteID === inviteId; });
  l.sort(function (a, b) { return Number(b.Rev) - Number(a.Rev); });
  return l[0] || null;
}
function api_quoteAction(token, quoteId, action, remarks, counterAmount) {
  var me = auth_(token, ['admin', 'owner']);
  return withLock_(function () {
    var q = find_('Quotes', quoteId); if (!q) throw new Error('Quote not found.');
    needOwner_(me, q.ProjectID);
    var pk = pkg_(q.PackageID), s = find_('Suppliers', q.SupplierID), pr = project_(q.ProjectID);
    if (pk.AwardedSupplierID) throw new Error('This package is already awarded.');
    if (q.Status === 'Superseded') throw new Error('This is an old revision.');
    var patch = { OwnerRemarks: str_(remarks, 2000), UpdatedAt: now_() }, subj, body;
    if (action === 'accept') {
      patch.Status = 'Accepted'; subj = 'Your quote was accepted – ' + pk.Trade;
      body = '<p>The owner <b>accepted</b> your quote of <b>' + peso_(q.Total) + '</b> for <b>' + esc_(pk.Trade) + '</b> (' +
             esc_(pr.Branch || pr.Name) + '). You will be notified once the package is awarded.</p>';
    } else if (action === 'negotiate') {
      patch.Status = 'Negotiating'; patch.CounterAmount = counterAmount ? num_(counterAmount) : '';
      subj = 'Price negotiation – ' + pk.Trade;
      body = '<p>The owner would like to <b>negotiate</b> your quote for <b>' + esc_(pk.Trade) + '</b> (' + esc_(pr.Branch || pr.Name) + ').</p>' +
             table_([['Your quote', peso_(q.Total)], ['Owner counter-offer', patch.CounterAmount ? peso_(patch.CounterAmount) : '—'],
                     ['Remarks', remarks || '—']]) +
             '<p>Please log in and send a revised quote. Tick <b>"This is my final quote"</b> when it is your best price.</p>' + btn_('Revise quote');
    } else if (action === 'decline') {
      patch.Status = 'Declined'; patch.AwardStatus = 'Not Awarded'; subj = 'Quotation update – ' + pk.Trade;
      body = '<p>Thank you for your quotation for <b>' + esc_(pk.Trade) + '</b> – ' + esc_(pr.Branch || pr.Name) +
             '. After review, the owner will not proceed with your quote.</p>' + (remarks ? '<p>Remarks: ' + esc_(remarks) + '</p>' : '');
    } else throw new Error('Unknown action');
    var r = update_('Quotes', quoteId, patch);
    mail_(s.Email, subj, box_(subj, body), quoteId);
    return r;
  });
}
function api_award(token, packageId, quoteId, opts) {
  var me = auth_(token, ['admin', 'owner']);
  opts = opts || {};
  return withLock_(function () {
    var pk = pkg_(packageId); needOwner_(me, pk.ProjectID);
    if (pk.AwardedSupplierID) throw new Error('Already awarded. Cancel the award first to re-award.');
    var q = find_('Quotes', quoteId);
    if (!q || q.PackageID !== packageId) throw new Error('Quote not found.');
    if (['Accepted', 'Final', 'Submitted', 'Negotiating'].indexOf(q.Status) < 0) throw new Error('This quote cannot be awarded.');
    var pr = project_(pk.ProjectID), win = find_('Suppliers', q.SupplierID);
    var amount = opts.amount ? num_(opts.amount) : num_(q.Total);
    update_('Packages', packageId, { Status: 'Awarded', AwardedSupplierID: q.SupplierID, AwardedQuoteID: quoteId, ContractAmount: amount,
      DPRequired: !!opts.dpRequired, RetentionPct: num_(opts.retentionPct), WarrantyMonths: num_(q.WarrantyMonths), AwardedAt: now_() });
    update_('Packages', packageId, cleanSched_(opts.schedule));
    update_('Quotes', quoteId, { Status: 'Awarded', AwardStatus: 'Awarded', NotifiedAt: now_(), UpdatedAt: now_() });
    if (pr.Status === 'Planning') update_('Projects', pr.ProjectID, { Status: 'Ongoing' });
    // milestones template
    if (!rows_('Milestones').some(function (m) { return m.PackageID === packageId; })) {
      var ms = DEFAULT_MILESTONES[pk.Trade] || DEFAULT_MILESTONES.Others, w = Math.round(100 / ms.length * 100) / 100;
      ms.forEach(function (n, i) {
        insert_('Milestones', { MilestoneID: id_('MS'), PackageID: packageId, ProjectID: pk.ProjectID, Seq: i + 1, Name: n,
                                Weight: w, Done: false, DoneAt: '', Photos: [] });
      });
    }
    mail_(win.Email, 'Award notice: ' + pk.Trade + ' – ' + (pr.Branch || pr.Name), box_('Congratulations – you have been awarded!',
      '<p>Hi ' + esc_(win.ContactPerson || win.Company) + ',</p><p>We are pleased to inform you that your quotation has been <b>awarded</b>.</p>' +
      table_([['Project', pr.Name + (pr.Branch ? ' – ' + pr.Branch : '')], ['Location', pr.Address], ['Package', pk.Trade],
              ['Contract amount', peso_(amount)], ['Payment terms', q.PaymentTerms], ['Downpayment', opts.dpRequired ? 'Yes – you will be notified once it is sent' : 'None'],
              ['Retention', num_(opts.retentionPct) ? num_(opts.retentionPct) + '% until turnover' : 'None'],
              ['Required progress updates', schedText_(pkg_(packageId))]]) +
      '<p><b>Next step:</b> ' + (opts.dpRequired ? 'wait for the downpayment, then submit your mobilization details in the portal.'
                                                 : 'log in and submit your mobilization date, manpower and materials.') + '</p>' + btn_('Open supplier portal')), quoteId);
    // regret notices to everyone else who quoted
    var others = {};
    rows_('Quotes').forEach(function (o) {
      if (o.PackageID === packageId && o.SupplierID !== q.SupplierID && o.Status !== 'Superseded') others[o.SupplierID] = o;
    });
    Object.keys(others).forEach(function (sid) {
      var o = others[sid], s = find_('Suppliers', sid);
      update_('Quotes', o.QuoteID, { Status: o.Status === 'Declined' ? 'Declined' : 'Not Awarded', AwardStatus: 'Not Awarded', NotifiedAt: now_(), UpdatedAt: now_() });
      if (o.Status === 'Declined') return; // already told
      mail_(s.Email, 'Quotation result: ' + pk.Trade + ' – ' + (pr.Branch || pr.Name), box_('Thank you for your quotation',
        '<p>Hi ' + esc_(s.ContactPerson || s.Company) + ',</p><p>Thank you for your quotation for <b>' + esc_(pk.Trade) + '</b> – ' +
        esc_(pr.Name + (pr.Branch ? ' – ' + pr.Branch : '')) + '. After review, the project has been awarded to another contractor.</p>' +
        (opts.regretRemark ? '<p>Owner\'s note: ' + esc_(opts.regretRemark) + '</p>' : '') +
        '<p>We appreciate your time and hope to work with you on future projects.</p>'), o.QuoteID);
    });
    // close open invitations
    rows_('Invitations').forEach(function (i) {
      if (i.PackageID === packageId && (i.Status === 'Invited' || i.Status === 'Viewed')) update_('Invitations', i.InviteID, { Status: 'Closed' });
    });
    return pkg_(packageId);
  });
}
function api_cancelAward(token, packageId, reason) {
  var me = auth_(token, ['admin', 'owner']);
  return withLock_(function () {
    var pk = pkg_(packageId); needOwner_(me, pk.ProjectID);
    if (!pk.AwardedSupplierID) throw new Error('Not awarded.');
    if (rows_('Payments').some(function (p) { return p.PackageID === packageId; }))
      throw new Error('Payments were already recorded for this package. Settle them with the supplier first (contact SPFC).');
    var s = find_('Suppliers', pk.AwardedSupplierID), pr = project_(pk.ProjectID);
    update_('Quotes', pk.AwardedQuoteID, { Status: 'Award Cancelled', AwardStatus: 'Cancelled', UpdatedAt: now_() });
    update_('Packages', packageId, { Status: 'Quoting', AwardedSupplierID: '', AwardedQuoteID: '', ContractAmount: 0, AwardedAt: '' });
    rows_('Quotes').forEach(function (o) {   // let the other quotes be awardable again
      if (o.PackageID === packageId && o.Status === 'Not Awarded') update_('Quotes', o.QuoteID, { Status: 'Submitted', AwardStatus: 'Pending' });
    });
    mail_(s.Email, 'Award cancelled: ' + pk.Trade + ' – ' + (pr.Branch || pr.Name), box_('Award cancelled',
      '<p>The award for <b>' + esc_(pk.Trade) + '</b> – ' + esc_(pr.Branch || pr.Name) + ' has been <b>cancelled</b>.</p><p>Reason: ' + esc_(reason || '—') + '</p>'), packageId);
    return pkg_(packageId);
  });
}

// ───────────────────────── Payments ─────────────────────────
function api_requestPayment(token, d) {
  var me = auth_(token, ['supplier']);
  var pk = pkg_(d.PackageID); needAwardee_(me, pk);
  var amt = num_(d.Amount); if (amt <= 0) throw new Error('Please enter the amount.');
  var r = insert_('PayRequests', { RequestID: id_('PR'), PackageID: pk.PackageID, ProjectID: pk.ProjectID, SupplierID: me.id,
    Type: str_(d.Type, 40) || 'Progress', Amount: amt, Milestone: str_(d.Milestone, 300), Notes: str_(d.Notes, 2000),
    Attachments: d.Attachments || [], Status: 'Requested', RequestedAt: now_(), DecidedAt: '', RejectReason: '' });
  notifyOwner_(pk.ProjectID, 'Payment request from ' + me.name + ': ' + peso_(amt), '<p><b>' + esc_(me.name) + '</b> (' + esc_(pk.Trade) +
    ') is requesting a <b>' + esc_(r.Type) + '</b> payment.</p>' + table_([['Amount', peso_(amt)], ['For', r.Milestone], ['Notes', r.Notes]]) +
    '<p>Upload your proof of settlement in the portal once paid.</p>' + btn_('Open portal'), r.RequestID);
  return r;
}
function api_rejectRequest(token, requestId, reason) {
  var me = auth_(token, ['admin', 'owner']);
  var r = find_('PayRequests', requestId); needOwner_(me, r.ProjectID);
  var u = update_('PayRequests', requestId, { Status: 'Rejected', DecidedAt: now_(), RejectReason: str_(reason, 1000) });
  var s = find_('Suppliers', r.SupplierID);
  mail_(s.Email, 'Payment request not approved', box_('Payment request update', '<p>Your payment request of <b>' + peso_(r.Amount) +
    '</b> was not approved.</p><p>Reason: ' + esc_(reason || '—') + '</p>'), requestId);
  return u;
}
function api_recordPayment(token, d) {
  var me = auth_(token, ['admin', 'owner']);
  var pk = pkg_(d.PackageID); needOwner_(me, pk.ProjectID);
  if (!pk.AwardedSupplierID) throw new Error('Award the package first.');
  var amt = num_(d.Amount); if (amt <= 0) throw new Error('Please enter the amount.');
  return withLock_(function () {
    var p = insert_('Payments', { PaymentID: id_('PAY'), PackageID: pk.PackageID, ProjectID: pk.ProjectID, SupplierID: pk.AwardedSupplierID,
      Type: str_(d.Type, 40) || 'Progress', RequestID: str_(d.RequestID, 40), Amount: amt, PaidDate: str_(d.PaidDate, 10) || today_(),
      Method: str_(d.Method, 60), RefNo: str_(d.RefNo, 100), Proof: d.Proof || [], Notes: str_(d.Notes, 1000),
      Status: 'Sent', RecordedAt: now_(), ConfirmedAt: '' });
    if (d.RequestID) update_('PayRequests', d.RequestID, { Status: 'Paid', DecidedAt: now_() });
    var s = find_('Suppliers', pk.AwardedSupplierID), pr = project_(pk.ProjectID);
    var proofLinks = (p.Proof || []).map(function (f) { return '<a href="' + f.url + '">' + esc_(f.name) + '</a>'; }).join('<br>');
    mail_(s.Email, 'Payment sent: ' + peso_(amt) + ' – ' + pk.Trade, box_('Payment sent',
      '<p>The owner recorded a <b>' + esc_(p.Type) + '</b> payment for <b>' + esc_(pk.Trade) + '</b> – ' + esc_(pr.Branch || pr.Name) + '.</p>' +
      table_([['Amount', peso_(amt)], ['Date', p.PaidDate], ['Method', p.Method], ['Reference no.', p.RefNo]]) +
      (proofLinks ? '<p>Proof of settlement:<br>' + proofLinks + '</p>' : '') +
      '<p>Please log in and click <b>Confirm received</b>.' + (p.Type === 'Downpayment' ? ' Then submit your mobilization details.' : '') + '</p>' + btn_('Open supplier portal')), p.PaymentID);
    return p;
  });
}
function api_confirmPayment(token, paymentId) {
  var me = auth_(token, ['supplier', 'admin']);
  var p = find_('Payments', paymentId);
  if (me.role === 'supplier' && p.SupplierID !== me.id) throw new Error('Not allowed.');
  var u = update_('Payments', paymentId, { Status: 'Received', ConfirmedAt: now_() });
  notifyOwner_(p.ProjectID, 'Payment confirmed received: ' + peso_(p.Amount), '<p><b>' + esc_(find_('Suppliers', p.SupplierID).Company) +
    '</b> confirmed receiving your payment of <b>' + peso_(p.Amount) + '</b> (' + esc_(p.Type) + ', ref ' + esc_(p.RefNo || '—') + ').</p>', paymentId);
  return u;
}

// ───────────────────────── Mobilization & milestones ─────────────────────────
function api_saveMobilization(token, d) {
  var me = auth_(token, ['supplier', 'admin']);
  var pk = pkg_(d.PackageID); needAwardee_(me, pk);
  var dpRequired = pk.DPRequired === true || String(pk.DPRequired).toUpperCase() === 'TRUE';
  if (dpRequired && !rows_('Payments').some(function (p) { return p.PackageID === pk.PackageID && p.Type === 'Downpayment'; }))
    throw new Error('Mobilization opens after the downpayment is sent.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.MobDate || '')) throw new Error('Please set the mobilization date.');
  var days = num_(d.DurationDays); if (days <= 0) throw new Error('Please enter how many days the work will take.');
  var end = new Date(d.MobDate + 'T00:00:00'); end.setDate(end.getDate() + days);
  var rec = { PackageID: pk.PackageID, ProjectID: pk.ProjectID, SupplierID: pk.AwardedSupplierID, MobDate: d.MobDate, DurationDays: days,
    TargetEnd: Utilities.formatDate(end, tz_(), 'yyyy-MM-dd'), Workers: num_(d.Workers), Foreman: str_(d.Foreman, 200),
    ForemanMobile: str_(d.ForemanMobile, 40), Materials: str_(d.Materials, 4000), Notes: str_(d.Notes, 2000), SubmittedAt: now_() };
  return withLock_(function () {
    var exists = find_('Mobilization', pk.PackageID);
    if (exists) update_('Mobilization', pk.PackageID, rec); else insert_('Mobilization', rec);
    if (['Awarded', 'Mobilized'].indexOf(pk.Status) >= 0) update_('Packages', pk.PackageID, { Status: 'Mobilized' });
    notifyOwner_(pk.ProjectID, 'Mobilization schedule: ' + pk.Trade, '<p><b>' + esc_(find_('Suppliers', pk.AwardedSupplierID).Company) + '</b> ' +
      (exists ? 'updated' : 'submitted') + ' the mobilization details for <b>' + esc_(pk.Trade) + '</b>.</p>' +
      table_([['Mobilization date', rec.MobDate], ['Duration', days + ' days'], ['Target completion', rec.TargetEnd], ['Workers', rec.Workers],
              ['Foreman', rec.Foreman + ' ' + rec.ForemanMobile], ['Materials', rec.Materials]]), pk.PackageID);
    return rec;
  });
}
function api_saveMilestones(token, packageId, list) {
  var me = auth_(token);
  var pk = pkg_(packageId);
  if (!ownsProject_(me, pk.ProjectID)) needAwardee_(me, pk);
  return withLock_(function () {
    var existing = rows_('Milestones').filter(function (m) { return m.PackageID === packageId; });
    var keep = {};
    (list || []).forEach(function (m, i) {
      if (!str_(m.Name)) return;
      if (m.MilestoneID && find_('Milestones', m.MilestoneID)) {
        keep[m.MilestoneID] = 1;
        update_('Milestones', m.MilestoneID, { Seq: i + 1, Name: str_(m.Name, 200), Weight: num_(m.Weight) });
      } else insert_('Milestones', { MilestoneID: id_('MS'), PackageID: packageId, ProjectID: pk.ProjectID, Seq: i + 1,
                                     Name: str_(m.Name, 200), Weight: num_(m.Weight), Done: false, DoneAt: '', Photos: [] });
    });
    existing.forEach(function (m) {
      if (!keep[m.MilestoneID] && !(m.Done === true || String(m.Done).toUpperCase() === 'TRUE')) sh_('Milestones').getRange(m._row, 1).setValue('');
    });
    delete _rows.Milestones;
    return rows_('Milestones').filter(function (m) { return m.PackageID === packageId; }).map(clean_);
  });
}
function api_tickMilestone(token, milestoneId, done, photos) {
  var me = auth_(token, ['supplier', 'admin']);
  var m = find_('Milestones', milestoneId); var pk = pkg_(m.PackageID); needAwardee_(me, pk);
  if (done && !(photos && photos.length) && !(m.Photos && m.Photos.length)) throw new Error('Please attach at least one photo for this milestone.');
  var u = update_('Milestones', milestoneId, { Done: !!done, DoneAt: done ? now_() : '', Photos: (m.Photos || []).concat(photos || []) });
  if (done) notifyOwner_(pk.ProjectID, 'Milestone done: ' + m.Name + ' (' + pk.Trade + ')', '<p><b>' + esc_(pk.Trade) +
    '</b> milestone <b>' + esc_(m.Name) + '</b> was marked done.</p>' + photoLinks_(photos), milestoneId);
  return u;
}

// ───────────────────────── Weekly updates & concerns ─────────────────────────
function api_submitUpdate(token, d) {
  var me = auth_(token, ['supplier']);
  var pk = pkg_(d.PackageID); needAwardee_(me, pk);
  var media = d.Media || [];
  if (media.filter(function (f) { return /^image\//.test(f.mime || ''); }).length < 3) throw new Error('Please attach at least 3 photos.');
  if (!str_(d.WorkDone)) throw new Error('Please describe the work done.');
  var u = insert_('Updates', { UpdateID: id_('UPD'), PackageID: pk.PackageID, ProjectID: pk.ProjectID, SupplierID: me.id,
    WeekOf: str_(d.WeekOf, 10) || today_(), Percent: Math.max(0, Math.min(100, num_(d.Percent))), WorkDone: str_(d.WorkDone, 4000),
    NextPlan: str_(d.NextPlan, 4000), Manpower: num_(d.Manpower), Delays: str_(d.Delays, 3000), Media: media,
    VideoLink: str_(d.VideoLink, 500), SubmittedAt: now_() });
  if (['Awarded', 'Mobilized'].indexOf(pk.Status) >= 0) update_('Packages', pk.PackageID, { Status: 'Ongoing' });
  notifyOwner_(pk.ProjectID, 'Progress update: ' + pk.Trade + ' – ' + u.Percent + '%', '<p><b>' + esc_(me.name) + '</b> posted a progress update for <b>' +
    esc_(pk.Trade) + '</b>.</p>' + table_([['Progress', u.Percent + '%'], ['Work done', u.WorkDone], ['Next week', u.NextPlan],
    ['Manpower', u.Manpower], ['Delays', u.Delays || 'None']]) + photoLinks_(media) + btn_('View in portal'), u.UpdateID);
  return u;
}
function api_raiseConcern(token, d) {
  var me = auth_(token);
  var pk = pkg_(d.PackageID);
  if (!ownsProject_(me, pk.ProjectID)) needAwardee_(me, pk);
  if (!str_(d.Description)) throw new Error('Please describe the concern.');
  var c = insert_('Concerns', { ConcernID: id_('CN'), PackageID: pk.PackageID, ProjectID: pk.ProjectID,
    RaisedBy: me.role === 'supplier' ? 'Supplier' : me.role === 'owner' ? 'Owner' : 'Admin', RaisedByName: me.name,
    Category: str_(d.Category, 60), Description: str_(d.Description, 4000), Media: d.Media || [], Status: 'Open',
    Response: '', RespondedBy: '', CreatedAt: now_(), UpdatedAt: now_() });
  var subj = 'Concern raised – ' + pk.Trade + ': ' + c.Category;
  var body = '<p><b>' + esc_(me.name) + '</b> raised a concern on <b>' + esc_(pk.Trade) + '</b>.</p>' +
             table_([['Category', c.Category], ['Details', c.Description]]) + photoLinks_(c.Media) + btn_('Respond in portal');
  if (me.role === 'supplier') notifyOwner_(pk.ProjectID, subj, body, c.ConcernID);
  else if (pk.AwardedSupplierID) mail_(find_('Suppliers', pk.AwardedSupplierID).Email, subj, box_('Concern raised', body), c.ConcernID);
  return c;
}
function api_respondConcern(token, concernId, response, status) {
  var me = auth_(token);
  var c = find_('Concerns', concernId), pk = pkg_(c.PackageID);
  if (!ownsProject_(me, pk.ProjectID)) needAwardee_(me, pk);
  var resp = (c.Response ? c.Response + '\n' : '') + '[' + now_() + ' ' + me.name + '] ' + str_(response, 3000);
  var u = update_('Concerns', concernId, { Response: resp, RespondedBy: me.name, UpdatedAt: now_(),
                                           Status: ['Open', 'Responded', 'Resolved'].indexOf(status) >= 0 ? status : 'Responded' });
  var subj = 'Concern ' + u.Status.toLowerCase() + ' – ' + pk.Trade + ': ' + c.Category;
  var body = '<p><b>' + esc_(me.name) + '</b> replied:</p><p>' + esc_(response) + '</p><p>Status: <b>' + esc_(u.Status) + '</b></p>';
  if (me.role === 'supplier') notifyOwner_(pk.ProjectID, subj, body, concernId);
  else if (pk.AwardedSupplierID) mail_(find_('Suppliers', pk.AwardedSupplierID).Email, subj, box_('Concern update', body), concernId);
  return u;
}

// ───────────────────────── Change orders ─────────────────────────
function api_requestCO(token, d) {
  var me = auth_(token, ['supplier']);
  var pk = pkg_(d.PackageID); needAwardee_(me, pk);
  if (!str_(d.Description)) throw new Error('Please describe the change.');
  var co = insert_('ChangeOrders', { COID: id_('CO'), PackageID: pk.PackageID, ProjectID: pk.ProjectID, SupplierID: me.id,
    Description: str_(d.Description, 3000), Amount: num_(d.Amount), AddDays: num_(d.AddDays), Reason: str_(d.Reason, 2000),
    Attachments: d.Attachments || [], Status: 'Pending', OwnerRemarks: '', RequestedAt: now_(), DecidedAt: '' });
  notifyOwner_(pk.ProjectID, 'Change order request – ' + pk.Trade, '<p><b>' + esc_(me.name) + '</b> requested a change order.</p>' +
    table_([['Change', co.Description], ['Additional cost', peso_(co.Amount)], ['Additional days', co.AddDays], ['Reason', co.Reason]]) +
    btn_('Approve or reject in portal'), co.COID);
  return co;
}
function api_decideCO(token, coId, approve, remarks) {
  var me = auth_(token, ['admin', 'owner']);
  return withLock_(function () {
    var co = find_('ChangeOrders', coId); needOwner_(me, co.ProjectID);
    if (co.Status !== 'Pending') throw new Error('Already decided.');
    var u = update_('ChangeOrders', coId, { Status: approve ? 'Approved' : 'Rejected', OwnerRemarks: str_(remarks, 2000), DecidedAt: now_() });
    if (approve && num_(co.AddDays)) {
      var mob = find_('Mobilization', co.PackageID);
      if (mob && mob.TargetEnd) {
        var d = new Date(mob.TargetEnd + 'T00:00:00'); d.setDate(d.getDate() + num_(co.AddDays));
        update_('Mobilization', co.PackageID, { TargetEnd: Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'),
                                                DurationDays: num_(mob.DurationDays) + num_(co.AddDays) });
      }
    }
    mail_(find_('Suppliers', co.SupplierID).Email, 'Change order ' + u.Status.toLowerCase(), box_('Change order ' + u.Status.toLowerCase(),
      '<p>Your change order "<b>' + esc_(co.Description) + '</b>" (' + peso_(co.Amount) + ', +' + co.AddDays + ' days) was <b>' +
      u.Status.toLowerCase() + '</b>.</p>' + (remarks ? '<p>Remarks: ' + esc_(remarks) + '</p>' : '')), coId);
    return u;
  });
}

// ───────────────────────── Turnover ─────────────────────────
function api_submitTurnover(token, d) {
  var me = auth_(token, ['supplier']);
  var pk = pkg_(d.PackageID); needAwardee_(me, pk);
  if (pk.TurnedOverAt) throw new Error('Already turned over.');
  var rec = { PackageID: pk.PackageID, ProjectID: pk.ProjectID, SupplierID: me.id, Checklist: d.Checklist || [],
              Docs: d.Docs || [], Notes: str_(d.Notes, 3000), SubmittedAt: now_() };
  return withLock_(function () {
    if (find_('Turnover', pk.PackageID)) update_('Turnover', pk.PackageID, rec); else insert_('Turnover', rec);
    update_('Packages', pk.PackageID, { Status: 'For Turnover' });
    notifyOwner_(pk.ProjectID, 'Ready for turnover: ' + pk.Trade, '<p><b>' + esc_(me.name) + '</b> marked <b>' + esc_(pk.Trade) +
      '</b> as complete and submitted the turnover checklist.</p><p>Please inspect the work, then click <b>Turned Over</b> in the portal.</p>' +
      btn_('Review turnover'), pk.PackageID);
    return rec;
  });
}
function api_turnOver(token, packageId, remarks) {
  var me = auth_(token, ['admin', 'owner']);
  return withLock_(function () {
    var pk = pkg_(packageId); needOwner_(me, pk.ProjectID);
    if (!find_('Turnover', packageId)) throw new Error('The contractor has not submitted the turnover checklist yet.');
    var openC = rows_('Concerns').filter(function (c) { return c.PackageID === packageId && c.Status !== 'Resolved'; }).length;
    if (openC) throw new Error('There are ' + openC + ' unresolved concern(s). Resolve them first.');
    var wEnd = '';
    if (num_(pk.WarrantyMonths)) { var d = new Date(); d.setMonth(d.getMonth() + num_(pk.WarrantyMonths)); wEnd = Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
    var u = update_('Packages', packageId, { Status: 'Turned Over', TurnedOverAt: now_(), WarrantyEnd: wEnd });
    var pr = project_(pk.ProjectID);
    mail_(find_('Suppliers', pk.AwardedSupplierID).Email, 'Turned over: ' + pk.Trade + ' – ' + (pr.Branch || pr.Name), box_('Project turned over',
      '<p>The owner accepted the turnover of <b>' + esc_(pk.Trade) + '</b>. Thank you for your work!</p>' +
      table_([['Turned over', u.TurnedOverAt], ['Warranty until', wEnd || '—'], ['Remarks', remarks || '—']]) +
      (num_(pk.RetentionPct) ? '<p>Retention (' + pk.RetentionPct + '%) can be requested after the warranty period.</p>' : '')), packageId);
    var all = rows_('Packages').filter(function (p) { return p.ProjectID === pk.ProjectID && p.Status !== 'Draft'; });
    if (all.length && all.every(function (p) { return p.Status === 'Turned Over'; })) update_('Projects', pk.ProjectID, { Status: 'Completed' });
    return u;
  });
}

// ───────────────────────── Daily reminders (trigger) ─────────────────────────
function dailyJobs() {
  var today = today_();
  var tmr = addDays_(today, 1);
  rows_('Invitations').forEach(function (i) {
    if (i.Status !== 'Invited' && i.Status !== 'Viewed') return;
    var s = find_('Suppliers', i.SupplierID) || (i.SupplierEmail ? { Email: i.SupplierEmail } : null), pk = find_('Packages', i.PackageID);
    if (!s || !pk) return;
    if (i.Deadline && i.Deadline < today) update_('Invitations', i.InviteID, { Status: 'Expired' });
    else if (i.Deadline === tmr) mail_(s.Email, 'Reminder: quote due tomorrow – ' + pk.Trade, box_('Quote due tomorrow',
      '<p>Your quotation for <b>' + esc_(pk.Trade) + '</b> is due on <b>' + i.Deadline + '</b>.</p>' + btn_('Submit quote')), i.InviteID);
  });
  rows_('Packages').forEach(function (pk) {
    if (['Awarded', 'Mobilized', 'Ongoing'].indexOf(pk.Status) < 0) return;
    var st = updState_(pk, today);
    if (!st.active) return;
    var s = find_('Suppliers', pk.AwardedSupplierID); if (!s) return;
    if (st.remind)
      mail_(s.Email, 'Reminder: progress update due – ' + pk.Trade, box_('Progress update due',
        '<p>Please submit your progress update for <b>' + esc_(pk.Trade) + '</b>' + (st.next ? ' (due <b>' + st.next + '</b>)' : '') + ': % complete, work done, next plan and at least 3 photos.</p>' +
        '<p class="small">Required updates: ' + esc_(schedText_(pk)) + '</p>' + (pk.UpdateNote ? '<p>Owner\'s note: ' + esc_(pk.UpdateNote) + '</p>' : '') +
        btn_('Submit update')), pk.PackageID);
    if (st.alert)
      notifyOwner_(pk.ProjectID, 'Progress update overdue: ' + pk.Trade, '<p><b>' + esc_(s.Company) + '</b> has not submitted the required progress update for <b>' +
        esc_(pk.Trade) + '</b> (due ' + (st.next || '—') + ', last update: ' + (st.last || 'none') + ').</p><p>Required updates: ' + esc_(schedText_(pk)) + '</p>', pk.PackageID);
  });
}

// ───────────────────────── Required progress updates (set by the owner) ─────────────────────────
var UPDATE_FREQS = ['Daily', 'Weekly', 'Every 2 weeks', 'One time', 'None'];
var WEEKDAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
function cleanSched_(sc) {
  sc = sc || {};
  var f = UPDATE_FREQS.indexOf(sc.freq) >= 0 ? sc.freq : 'Weekly';
  return { UpdateFreq: f, UpdateDay: WEEKDAYS.indexOf(sc.day) >= 0 ? sc.day : 'Friday',
           UpdateDueDate: /^\d{4}-\d{2}-\d{2}$/.test(sc.dueDate || '') ? sc.dueDate : '', UpdateNote: str_(sc.note, 1000) };
}
function schedText_(pk) {
  var f = pk.UpdateFreq || 'Weekly';
  if (f === 'Weekly' || f === 'Every 2 weeks') return f + ' (every ' + (pk.UpdateDay || 'Friday') + ')';
  if (f === 'One time') return 'One time only' + (pk.UpdateDueDate ? ', due ' + pk.UpdateDueDate : '');
  if (f === 'Daily') return 'Daily';
  return 'No regular updates';
}
function addDays_(ymd, n) { var d = new Date(ymd + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function daysDiff_(a, b) { return Math.round((new Date(b + 'T00:00:00Z') - new Date(a + 'T00:00:00Z')) / 864e5); }
function weekday_(ymd) { return WEEKDAYS[(new Date(ymd + 'T00:00:00Z').getUTCDay() + 6) % 7]; }
// Same rules as pkgUpdateState() in public/index.html – keep both in sync.
function updState_(pk, today) {
  var ups = rows_('Updates').filter(function (u) { return u.PackageID === pk.PackageID; })
    .map(function (u) { return String(u.SubmittedAt).slice(0, 10); }).sort();
  var last = ups.length ? ups[ups.length - 1] : '';
  var mob = find_('Mobilization', pk.PackageID);
  return updateRule_(pk.UpdateFreq || 'Weekly', pk.UpdateDay || 'Friday', pk.UpdateDueDate, last, mob ? mob.MobDate : '', today);
}
function updateRule_(freq, day, dueDate, last, mobDate, today) {
  var out = { active: false, next: '', overdue: false, remind: false, alert: false, last: last };
  if (freq === 'None') return out;
  if (freq === 'One time') {
    if (last || !dueDate) return out;
    out.active = true; out.next = dueDate; out.overdue = today > dueDate;
    out.remind = today === dueDate || today === addDays_(dueDate, -1);
    out.alert = today === addDays_(dueDate, 1);
    return out;
  }
  if (!mobDate || mobDate > today) return out;           // periodic updates start at mobilization
  out.active = true;
  var ref = last || mobDate;
  if (freq === 'Daily') {
    out.next = last && last >= today ? addDays_(today, 1) : today;
    out.remind = !(last && last >= today);
    out.overdue = daysDiff_(ref, today) >= 2;
    out.alert = daysDiff_(ref, today) >= 2 && (daysDiff_(ref, today) - 2) % 7 === 0;
    return out;
  }
  var d = addDays_(ref, freq === 'Every 2 weeks' ? 8 : 1);
  while (weekday_(d) !== day) d = addDays_(d, 1);
  out.next = d;
  out.remind = today >= d && weekday_(today) === day;
  out.overdue = today > d;
  out.alert = today > d && daysDiff_(d, today) % 7 === 1;
  return out;
}
function api_setUpdateSchedule(token, packageId, sc) {
  var me = auth_(token, ['admin', 'owner']);
  var pk = pkg_(packageId); needOwner_(me, pk.ProjectID);
  var u = update_('Packages', packageId, cleanSched_(sc));
  if (pk.AwardedSupplierID) {
    var s = find_('Suppliers', pk.AwardedSupplierID);
    mail_(s.Email, 'Required progress updates – ' + pk.Trade, box_('Progress update schedule',
      '<p>The owner set the required progress updates for <b>' + esc_(pk.Trade) + '</b>:</p>' +
      table_([['Schedule', schedText_(u)], ['Note', u.UpdateNote || '—']]) +
      '<p>You will get an email reminder when an update is due.</p>' + btn_('Open supplier portal')), packageId);
  }
  return u;
}

// ───────────────────────── Email helpers ─────────────────────────
function adminEmail_() { return props_().getProperty('ADMIN_EMAIL'); }
function notifyOwner_(projectId, subject, html, ref) {
  var pr = find_('Projects', projectId), o = pr && find_('Owners', pr.OwnerID);
  var appUrl = props_().getProperty('APP_URL'), ownerUrl = props_().getProperty('OWNER_URL');
  if (ownerUrl && appUrl) html = html.split(appUrl).join(ownerUrl); // owners open the app from the Odoo portal
  if (o && o.Email) mail_(o.Email, subject, box_(subject, html), ref, adminEmail_());
  else mail_(adminEmail_(), subject, box_(subject, html), ref);
}
function mail_(to, subject, html, ref, cc) {
  var res = 'sent';
  try {
    var opt = { htmlBody: html, name: 'Suds Construction Monitoring' };
    if (cc && cc !== to) opt.cc = cc;
    MailApp.sendEmail(to, subject, subject, opt);
  } catch (e) { res = 'error: ' + e.message; }
  try { insert_('EmailLog', { At: now_(), To: to, Subject: subject, Ref: ref || '', Result: res }); } catch (e) { }
}
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }).replace(/\n/g, '<br>'); }
function peso_(n) { return '₱' + num_(n).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function table_(rows) {
  return '<table style="border-collapse:collapse;margin:10px 0;font-size:14px">' + rows.map(function (r) {
    return '<tr><td style="padding:6px 12px 6px 0;color:#7A7A7A;vertical-align:top;white-space:nowrap">' + esc_(r[0]) +
           '</td><td style="padding:6px 0;font-weight:600">' + esc_(r[1] == null || r[1] === '' ? '—' : r[1]) + '</td></tr>';
  }).join('') + '</table>';
}
function btn_(label) {
  var url = props_().getProperty('APP_URL') || '#';
  return '<p style="margin:18px 0"><a href="' + url + '" style="background:#FF6B00;color:#fff;padding:11px 20px;border-radius:8px;text-decoration:none;font-weight:600">' + esc_(label) + '</a></p>';
}
function photoLinks_(files) {
  if (!files || !files.length) return '';
  return '<p>' + files.map(function (f) { return '<a href="' + f.url + '">' + esc_(f.name) + '</a>'; }).join(' · ') + '</p>';
}
function box_(title, inner) {
  return '<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;border:1px solid #ECECEC;border-radius:12px;overflow:hidden">' +
    '<div style="background:#FF6B00;color:#fff;padding:16px 20px;font-size:18px;font-weight:700">' + esc_(title) + '</div>' +
    '<div style="padding:18px 20px;color:#2B2B2B;font-size:14px;line-height:1.5">' + inner + '</div>' +
    '<div style="padding:10px 20px;background:#FAFAF8;color:#9A9A9A;font-size:12px">Suds Construction Monitoring · automated message</div></div>';
}
