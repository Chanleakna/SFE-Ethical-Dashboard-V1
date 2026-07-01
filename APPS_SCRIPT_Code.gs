/**
 * SFE KPI Dashboard — Backend (Apps Script)
 *
 * RE-DEPLOY when you change this code:
 *   Deploy → Manage deployments → pencil icon → Version: New version → Deploy
 *
 * FIXES IN THIS VERSION:
 *   1. Ethical filter is now case-insensitive (handles "Ethical", "ETHICAL", " ethical ", etc.)
 *      Trade rows are excluded properly in all 3 calculations (Sales, Active, New Listing).
 *   2. "Thay Dararith - Vacancy" is now treated as a DIFFERENT SR from "Thay Dararith".
 *      Removed the vacancy-strip from tokenize so SR Codes match correctly.
 *
 * NEW LISTING LOGIC:
 *   Customer matching is by NAME (not Customer Code), because 2025 data uses
 *   placeholder code 281000000 while 2026 uses real codes. Name-based matching
 *   correctly detects prior history across this code change.
 *   ALL OTHER KPIs continue to use Customer Code as before.
 */

// =============================================================================
// CONFIGURATION
// =============================================================================
const SHEET_IDS = {
  ims:      '1ocanfdc_9R6m0X0sH9W0dlCvc2U9EoeT50Wul61rrEk',
  daily:    '1-iJM9eU-cPqnjvHGBDKT8i59XN53rz9ykmyIqNpvkz8',
  material: '1GPxPYtkUD4kx2iCQMIUdctvbBrKS7XhCKp2JGy3NTr4',
};

const TAB_NAMES = {
  targets:      'Target Set',
  shopDetails:  'Target - Shop Around Details',
  activeTarget: 'Target-Active Cus 3 Months',
  newTarget:    'Target - New Listing',
  leadTarget:   'Target-Lead',
  leadActual:   'Actual-Lead',
  users:        'Users',
  loginLog:     'Login_Log',
  shared:       'Shared_Customers',
  shopCoverage: 'Shop_Coverage',
  hcpHco:       'HCP_HCO_Lookup',
  daily:        'Export',
  materials:    'Maser Material Code',
};

const MONTHS = [1,2,3,4,5,6,7,8,9,10,11,12];
const CACHE_SECONDS = 21600; // 6h — data changes once a day (morning import clears it)

// === Daily email import (auto-ingest the morning sales email) ===
// NOTE: Apps Script can only read GMAIL (the Google account that owns this
// script) — it cannot read Outlook/Microsoft mailboxes. If your daily sales
// email arrives in Outlook, auto-forward (or redirect) it to this Gmail address,
// then match it below. Configure at least ONE of: gmailLabel / senderEmail /
// subjectContains. gmailLabel (a Gmail filter + label) is the most reliable for
// forwarded mail. The script reads the latest matching Excel attachment and
// REPLACES the Export tab with its contents, then refreshes the dashboard cache.
const DAILY_IMPORT = {
  gmailLabel:      '',  // e.g. 'daily-sales' — recommended: add a Gmail filter that labels the forwarded email
  senderEmail:     '',  // e.g. the original sender, OR your Outlook address if you forward it
  subjectContains: '',  // optional subject text (matches even with FW:/RE: prefixes)
  searchWindowDays: 2,  // look back this many days for the latest matching email
};

// =============================================================================
// MAIN ENDPOINT
// =============================================================================
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    if (!action) {
      return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('SFE KPI Dashboard')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    }
    if (action === 'login')   return jsonResponse(handleLogin(e.parameter.code));
    if (action === 'data')    return jsonResponse(getDashboardData());
    if (action === 'accessLog') return jsonResponse(getAccessLog(e.parameter.code));
    if (action === 'health')  return jsonResponse({ ok: true, time: new Date().toISOString() });
    if (action === 'clearCache') {
      clearChunkedCache_(CacheService.getScriptCache(), 'dashboard_data');
      return jsonResponse({ ok: true, message: 'Cache cleared' });
    }
    return jsonResponse({ error: 'unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.message, stack: err.stack });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================================
// LOGIN
// =============================================================================
function handleLogin(code) {
  if (!code) return { ok: false, error: 'no code' };
  const ss = SpreadsheetApp.openById(SHEET_IDS.ims);
  const sh = ss.getSheetByName(TAB_NAMES.users);
  if (!sh) return { ok: false, error: 'Users tab not found in IMS sheet' };
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const codeCol = headers.indexOf('Code');
  const roleCol = headers.indexOf('Role');
  const nameCol = headers.indexOf('Name');
  const flmCol  = headers.indexOf('FLM');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][codeCol]) === String(code)) {
      const user = {
        role: data[i][roleCol],
        name: data[i][nameCol],
        flm:  data[i][flmCol] || null,
      };
      logAccess(ss, { code: code, name: user.name, role: user.role, flm: user.flm, status: 'success' });
      return { ok: true, user: user };
    }
  }
  logAccess(ss, { code: code, name: '', role: '', flm: '', status: 'invalid code' });
  return { ok: false, error: 'Invalid code' };
}

// Append a login attempt to the Login_Log tab (auto-creates it). Never throws.
function logAccess(ss, entry) {
  try {
    let log = ss.getSheetByName(TAB_NAMES.loginLog);
    if (!log) {
      log = ss.insertSheet(TAB_NAMES.loginLog);
      log.appendRow(['Timestamp', 'Name', 'Role', 'FLM', 'Code', 'Status']);
      log.setFrozenRows(1);
    }
    log.appendRow([
      new Date(),
      entry.name || '',
      entry.role || '',
      entry.flm || '',
      String(entry.code || ''),
      entry.status || '',
    ]);
  } catch (e) {
    Logger.log('logAccess failed: ' + e.message);
  }
}

// Returns the login history. Admin-only: the caller's code must map to an Admin user.
function getAccessLog(code) {
  if (!code) return { ok: false, error: 'no code' };
  const ss = SpreadsheetApp.openById(SHEET_IDS.ims);
  const users = ss.getSheetByName(TAB_NAMES.users);
  if (!users) return { ok: false, error: 'Users tab not found in IMS sheet' };
  const udata = users.getDataRange().getValues();
  const uh = udata[0];
  const cCol = uh.indexOf('Code'), rCol = uh.indexOf('Role');
  let role = null;
  for (let i = 1; i < udata.length; i++) {
    if (String(udata[i][cCol]) === String(code)) { role = udata[i][rCol]; break; }
  }
  if (String(role).trim().toLowerCase() !== 'admin') {
    return { ok: false, error: 'Admin access required' };
  }

  const log = ss.getSheetByName(TAB_NAMES.loginLog);
  if (!log) return { ok: true, rows: [] };
  const data = log.getDataRange().getValues();
  if (data.length < 2) return { ok: true, rows: [] };
  const h = data[0];
  const tsCol = h.indexOf('Timestamp'), nCol = h.indexOf('Name'),
        roCol = h.indexOf('Role'), fCol = h.indexOf('FLM'),
        coCol = h.indexOf('Code'), sCol = h.indexOf('Status');
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const ts = data[i][tsCol];
    rows.push({
      ts: ts instanceof Date ? ts.toISOString() : String(ts),
      name: data[i][nCol] || '',
      role: data[i][roCol] || '',
      flm: data[i][fCol] || '',
      code: String(data[i][coCol] || ''),
      status: data[i][sCol] || '',
    });
  }
  return { ok: true, rows: rows };
}

// =============================================================================
// DAILY EMAIL IMPORT
// Pulls the latest Excel attachment from DAILY_IMPORT.senderEmail and replaces
// the Export tab with its contents, then clears the dashboard cache.
//
// ONE-TIME SETUP:
//   0. Outlook only: add a rule to auto-forward/redirect the daily sales email
//      to the Gmail address that owns this script. (Optionally add a Gmail
//      filter that applies a label to it, e.g. 'daily-sales'.)
//   1. Apps Script editor → Services (＋) → add "Drive API" (advanced service).
//   2. Set DAILY_IMPORT.gmailLabel / senderEmail / subjectContains (top of file).
//   3. Run importDailyFromEmail() once manually to grant Gmail/Drive access.
//   4. Run installDailyImportTrigger() once to schedule it every morning.
// =============================================================================
function importDailyFromEmail() {
  const cfg = DAILY_IMPORT;
  if (!cfg.gmailLabel && !cfg.senderEmail && !cfg.subjectContains) {
    throw new Error('Configure at least one of DAILY_IMPORT.gmailLabel / senderEmail / subjectContains.');
  }

  const parts = ['has:attachment', 'newer_than:' + cfg.searchWindowDays + 'd'];
  if (cfg.gmailLabel)      parts.push('label:' + cfg.gmailLabel);
  if (cfg.senderEmail)     parts.push('from:' + cfg.senderEmail);
  if (cfg.subjectContains) parts.push('subject:("' + cfg.subjectContains + '")');
  const query = parts.join(' ');

  const threads = GmailApp.search(query, 0, 20);
  if (!threads.length) {
    Logger.log('No matching emails for query: ' + query);
    return { ok: false, error: 'No matching email found' };
  }

  // Find the newest message that has an Excel attachment.
  let latestDate = 0, latestMsg = null, excel = null;
  threads.forEach(function (t) {
    t.getMessages().forEach(function (m) {
      const d = m.getDate().getTime();
      if (d <= latestDate) return;
      const xlsx = m.getAttachments().filter(function (a) {
        return /\.xlsx?$/i.test(a.getName());
      })[0];
      if (xlsx) { latestDate = d; latestMsg = m; excel = xlsx; }
    });
  });
  if (!excel) {
    Logger.log('No Excel attachment found in matching emails.');
    return { ok: false, error: 'No Excel attachment found' };
  }

  // Convert the Excel blob to a temporary Google Sheet so we can read its cells.
  // Requires the "Drive API" advanced service to be enabled (see setup notes).
  const tempFile = Drive.Files.insert(
    { title: 'tmp_daily_import_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS },
    excel.copyBlob()
  );
  let values;
  try {
    const tmpSheet = SpreadsheetApp.openById(tempFile.id).getSheets()[0];
    values = tmpSheet.getDataRange().getValues();
  } finally {
    try { Drive.Files.remove(tempFile.id); } catch (e) {}
  }
  if (!values || values.length < 2) {
    return { ok: false, error: 'Attachment had no data rows' };
  }

  // === SAFETY VALIDATION — never wipe a good Export with a bad/mismatched file ===
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const REQUIRED = ['Customer Code', 'Total Act. Sales', 'Year', 'Short Cut', 'Dep'];
  const missing = REQUIRED.filter(function (h) { return headers.indexOf(h) < 0; });
  if (missing.length) {
    Logger.log('Import ABORTED — attachment missing columns: ' + missing.join(', '));
    return { ok: false, error: 'Attachment missing expected columns (' + missing.join(', ') + '). Export left unchanged.' };
  }

  const ss = SpreadsheetApp.openById(SHEET_IDS.daily);
  let sheet = ss.getSheetByName(TAB_NAMES.daily);
  const newRows = values.length - 1;
  const curRows = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
  // Refuse to replace good data with a suspiciously small file.
  if (newRows < 50 || (curRows > 0 && newRows < curRows * 0.5)) {
    Logger.log('Import ABORTED — new file has ' + newRows + ' rows vs current ' + curRows + '.');
    return { ok: false, error: 'New file had only ' + newRows + ' rows (current ' + curRows + '). Looks wrong — Export left unchanged.' };
  }

  // Back up the current Export before replacing, so a bad import is recoverable in-sheet.
  if (sheet && curRows > 0) {
    const oldBak = ss.getSheetByName('Export_BACKUP');
    if (oldBak) ss.deleteSheet(oldBak);
    sheet.copyTo(ss).setName('Export_BACKUP');
  }

  // Replace the Export tab with the new data (full dataset each morning).
  if (!sheet) sheet = ss.insertSheet(TAB_NAMES.daily);
  sheet.clearContents();
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);

  // Serve the fresh data immediately, then pre-warm the cache so the first
  // user after the morning import doesn't wait for a cold rebuild.
  try { clearChunkedCache_(CacheService.getScriptCache(), 'dashboard_data'); } catch (e) {}
  try { getDashboardData(); } catch (e) {}

  Logger.log('Imported ' + newRows + ' rows (email dated ' + latestMsg.getDate() + ').');
  return { ok: true, rows: newRows, emailDate: latestMsg.getDate() };
}

// Schedule importDailyFromEmail() to run every morning (~7am script timezone).
// Run this once from the editor; safe to re-run (it replaces the old trigger).
function installDailyImportTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'importDailyFromEmail') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importDailyFromEmail').timeBased().everyDays(1).atHour(7).create();
  Logger.log('Daily import trigger installed (runs ~7am).');
}


// =============================================================================
function getDashboardData() {
  const cache = CacheService.getScriptCache();
  const cached = readChunkedCache_(cache, 'dashboard_data');
  if (cached) {
    try {
      const obj = JSON.parse(cached);
      if (isPayloadHealthy_(obj)) return obj; // only trust a healthy cached payload
    } catch (e) {}
  }
  const data = buildDashboardPayload();
  // Only cache a HEALTHY payload. The morning import clears the cache, and a
  // read taken while the source tab is still rebuilding produces near-zero
  // actuals — caching that would serve zeros for CACHE_SECONDS (6h). Skipping
  // the cache on a bad read means the next request simply rebuilds.
  if (isPayloadHealthy_(data)) {
    try {
      writeChunkedCache_(cache, 'dashboard_data', JSON.stringify(data), CACHE_SECONDS);
    } catch (e) {}
  }
  return data;
}

// A build is "healthy" if it carries the core arrays and its total actuals are
// at least half the best total ever recorded (a monotonic high-water mark in
// Script Properties). This filters out partial reads taken mid-rebuild without
// needing a fixed threshold. Self-heals: the first full read sets the mark.
function isPayloadHealthy_(data) {
  if (!data || !Array.isArray(data.srs) || !data.srs.length) return false;
  if (!Array.isArray(data.actuals)) return false;
  let total = 0;
  for (let i = 0; i < data.actuals.length; i++) total += (data.actuals[i].v || 0);
  try {
    const props = PropertiesService.getScriptProperties();
    const hw = Number(props.getProperty('actuals_highwater')) || 0;
    if (total >= hw * 0.5) {
      if (total > hw) props.setProperty('actuals_highwater', String(total));
      return true;
    }
    return false;
  } catch (e) {
    return total > 0; // if Properties is unavailable, accept any non-empty build
  }
}

// === Chunked cache ===
// CacheService caps each item at ~100KB, so the (large) dashboard payload could
// never be cached as one value — which meant it was rebuilt on EVERY request
// (slow logins). Split it across several <100KB chunks so caching actually works.
const CACHE_CHUNK = 90000;

function writeChunkedCache_(cache, key, str, ttl) {
  const n = Math.ceil(str.length / CACHE_CHUNK);
  const parts = {};
  for (let i = 0; i < n; i++) {
    parts[key + '_' + i] = str.substring(i * CACHE_CHUNK, (i + 1) * CACHE_CHUNK);
  }
  cache.putAll(parts, ttl);
  cache.put(key + '_n', String(n), ttl);
}

function readChunkedCache_(cache, key) {
  const meta = cache.get(key + '_n');
  const n = Number(meta);
  if (!n) return null;
  const keys = [];
  for (let i = 0; i < n; i++) keys.push(key + '_' + i);
  const got = cache.getAll(keys);
  let out = '';
  for (let i = 0; i < n; i++) {
    const part = got[key + '_' + i];
    if (part == null) return null; // any missing chunk → treat as a cache miss
    out += part;
  }
  return out;
}

function clearChunkedCache_(cache, key) {
  const n = Number(cache.get(key + '_n')) || 0;
  const keys = [key + '_n'];
  for (let i = 0; i < n + 50; i++) keys.push(key + '_' + i); // +50 to clear stale extras
  cache.removeAll(keys);
}

function buildDashboardPayload() {
  const imsSS    = SpreadsheetApp.openById(SHEET_IDS.ims);
  const dailySS  = SpreadsheetApp.openById(SHEET_IDS.daily);
  const matSS    = SpreadsheetApp.openById(SHEET_IDS.material);

  const daily      = readSheet(dailySS,  TAB_NAMES.daily);
  const targets    = readSheet(imsSS,    TAB_NAMES.targets);
  const materials  = readSheet(matSS,    TAB_NAMES.materials);
  const shared     = readSheet(imsSS,    TAB_NAMES.shared);
  let shopCoverage = [];
  try { shopCoverage = readSheet(imsSS, TAB_NAMES.shopCoverage); }
  catch (e) { Logger.log('Note: Shop_Coverage tab not found'); }

  let hcpHcoLookup = [];
  try { hcpHcoLookup = readSheet(imsSS, TAB_NAMES.hcpHco); }
  catch (e) { Logger.log('Note: HCP_HCO_Lookup tab not found'); }

  const shopRaw    = readSheet(imsSS,    TAB_NAMES.shopDetails);
  const acTarget   = readSheet(imsSS,    TAB_NAMES.activeTarget);
  const nlTarget   = readSheet(imsSS,    TAB_NAMES.newTarget);
  const ldTarget   = readSheet(imsSS,    TAB_NAMES.leadTarget);

  const matMap = {};
  materials.forEach(r => {
    if (r['Material Code']) {
      matMap[Number(r['Material Code'])] = {
        sb: r['Sub-Brand'] || null,
        cat: r['Category'] || null,
      };
    }
  });

  const FLM_TYPOS = {
    'Thong Knaha': 'Thong Kanha',
    'Thong knaha': 'Thong Kanha',
    'Chhay Mengkong': 'Chay Mengkong',
    'Sem sokhom': 'Sem Sokhom',
    'In lena': 'In Lena',
    'Um phana': 'Um Phana',
    'Um Phanna': 'Um Phana',
    'Thong kanha': 'Thong Kanha',
  };
  const normFlm = (s) => {
    if (!s) return null;
    let v = String(s).trim();
    if (FLM_TYPOS[v]) return FLM_TYPOS[v];
    v = v.replace(/Chhay/g, 'Chay').replace(/Phanna/g, 'Phana').replace(/Knaha/g, 'Kanha');
    return v;
  };

  // === FIX: case-insensitive Ethical filter ===
  // Returns true if the row's Dep field equals "Ethical" (case- and whitespace-insensitive)
  const isEthical = (depValue) => {
    return String(depValue || '').trim().toUpperCase() === 'ETHICAL';
  };

  // Customer name normalization — used ONLY for New Listing (case + whitespace insensitive)
  const normName = (s) => {
    if (s === null || s === undefined || s === '') return null;
    return String(s).trim().toUpperCase().replace(/\s+/g, ' ');
  };

  const SHARED = {};
  shared.forEach(r => {
    const c = Number(r['Customer Code']);
    if (!c) return;
    if (!SHARED[c]) SHARED[c] = [];
    SHARED[c].push({
      sr: Number(r['SR Code']),
      cat: r['Category'],
      w: Number(r['Weight']) || 1,
    });
  });

  const SHOP_COVERAGE = {};
  shopCoverage.forEach(r => {
    const c = Number(r['Customer Code']);
    const sr = Number(r['SR Code']);
    if (!c || !sr) return;
    if (!SHOP_COVERAGE[c]) SHOP_COVERAGE[c] = [];
    SHOP_COVERAGE[c].push(sr);
  });

  const HCP_TO_HCO = {};
  hcpHcoLookup.forEach(r => {
    const hcp = Number(r['HCP ID'] || r['HCP Code'] || r['Customer Code']);
    const hco = r['HCO: Account Name'] || r['HCO Name'] || r['HCO Code'] || r['HCO'];
    if (!hcp || hco === undefined || hco === null || hco === '') return;
    HCP_TO_HCO[hcp] = String(hco).trim();
  });
  Logger.log('HCP_TO_HCO map size: ' + Object.keys(HCP_TO_HCO).length);

  const KPIS = ['SM','SIM','STC','PED PWD','PED RPB','ENS PWD','ENS RPB','GLU PWD','GLU RPB','PRO'];
  const ethical = targets.filter(r => r['Department'] === 'Ethical');

  const srSeen = {};
  const srMaster = [];
  // Strip '& Other Name' suffix from Seller Name display (e.g., "Chho Phanny & Ly Vuthea" -> "Chho Phanny")
  // The full name is preserved in the sheet; this only changes how it's displayed on the dashboard.
  const cleanSellerName = (s) => {
    if (!s) return s;
    return String(s).split('&')[0].trim();
  };
  
  ethical.forEach(r => {
    const code = Number(r['SR Code']);
    if (!code || srSeen[code]) return;
    srSeen[code] = true;
    srMaster.push({
      code: code,
      name: cleanSellerName(r['Seller Name']),
      flm: normFlm(r['ASM/SM Name']),
    });
  });
  const srToFlm = {};
  srMaster.forEach(s => srToFlm[s.code] = s.flm);

  const allTargets = [];
  ethical.forEach(r => {
    const d = new Date(r['Date']);
    const m = d.getDate();
    KPIS.forEach(k => {
      const v = Number(r[k]);
      if (v > 0) {
        allTargets.push({ m: m, sr: Number(r['SR Code']), k: k, t: Math.round(v) });
      }
    });
  });

  // === FIX: tokenize NO LONGER strips "Vacancy" ===
  // This keeps "Thay Dararith - Vacancy" SEPARATE from "Thay Dararith"
  // so each gets matched to its own SR Code in Target Set.
  const tokenize = (s) => {
    if (!s) return [];
    return String(s).toLowerCase()
      .replace(/&/g, ' ')
      .split(/[,;\s\-]+/).filter(p => p && p.length > 1);
  };

  const srMatch = {};
  const dailySrNames = {};
  const dailySrFlmCount = {}; // SR name -> { FLM -> count }, to pick each SR's FLM from the data
  daily.forEach(r => {
    if (r['SR']) {
      const first = String(r['SR']).split(/[,;]/)[0].trim();
      if (!first) return;
      dailySrNames[first] = true;
      const f = normFlm(r['FLM']);
      if (f) {
        if (!dailySrFlmCount[first]) dailySrFlmCount[first] = {};
        dailySrFlmCount[first][f] = (dailySrFlmCount[first][f] || 0) + 1;
      }
    }
  });

  Object.keys(dailySrNames).forEach(daiName => {
    const dt = tokenize(daiName);
    if (dt.length === 0) return;
    let best = null, bestScore = 0;
    srMaster.forEach(sr => {
      const tt = tokenize(sr.name);
      if (tt.length === 0) return;
      const dailyInMaster = dt.every(x => tt.indexOf(x) >= 0);
      const masterInDaily = tt.every(x => dt.indexOf(x) >= 0);
      if (!dailyInMaster && !masterInDaily) return;
      let score = (dt.length + tt.length) / (2 * Math.max(dt.length, tt.length));
      if (score > bestScore) { bestScore = score; best = sr; }
    });
    if (best && bestScore >= 0.7) {
      srMatch[daiName] = best.code;
    }
  });

  // Auto-include SRs that appear in the daily sales but are NOT in the Target Set
  // (e.g. Heng Norm). Without this they fall through as "Unassigned" and their
  // sales/active never show under their name or FLM. Each gets a stable synthetic
  // code and its FLM from the data, so every selling SR appears and the totals
  // reconcile to the raw file at total / FLM / SR level. Synthetic SRs simply
  // have no targets (target-based % shows against 0).
  Object.keys(dailySrNames).sort().forEach(function (nm) {
    if (srMatch[nm]) return; // already matched to a Target Set SR
    const code = 900000000 + (srMaster.length); // 9-digit, won't collide with real 8-digit codes
    let bestFlm = null, bestN = -1;
    const fc = dailySrFlmCount[nm] || {};
    Object.keys(fc).forEach(function (f) { if (fc[f] > bestN) { bestN = fc[f]; bestFlm = f; } });
    srMatch[nm] = code;
    srToFlm[code] = bestFlm;
    srMaster.push({ code: code, name: nm, flm: bestFlm, synthetic: true });
  });

  // MONTH_MAP handles BOTH 'Apr' and 'April' (your data has both spellings)
  const MONTH_MAP = {Jan:1,Feb:2,Mar:3,April:4,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
  // Robust month parser: accepts 'Jun', 'June', 'JUNE', ' june ', 'Sept', etc.
  // (any casing / full or short name) so no month's sales are silently dropped.
  const M3 = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
  const mapMonth = function (s) {
    if (s == null || s === '') return undefined;
    if (MONTH_MAP[s] != null) return MONTH_MAP[s];
    return M3[String(s).trim().toLowerCase().slice(0, 3)];
  };

  // Robust numeric parse. Some Export cells arrive as TEXT — "1,234", " 1,234.50 ",
  // currency "$1,234", or accounting negatives "(500)". Number() turns those into
  // NaN, which was being coerced to 0 and silently dropped, so the dashboard
  // undercounted vs the raw Excel column. Strip separators/symbols so every sale
  // is counted exactly as the sheet shows it.
  const numVal = function (x) {
    if (typeof x === 'number') return x;
    if (x == null || x === '') return 0;
    var s = String(x).trim();
    if (s === '') return 0;
    var neg = /^\(.*\)$/.test(s);            // accounting negative: (1,234) => -1234
    s = s.replace(/[^0-9.\-]/g, '');
    var n = parseFloat(s);
    if (!isFinite(n)) return 0;
    return neg ? -Math.abs(n) : n;
  };

  const custDict = {};
  daily.forEach(r => {
    const c = Number(r['Customer Code']);
    if (!c || numVal(r['Year']) !== 2026 || !r['FLM']) return;
    if (!custDict[c]) {
      const srFirst = r['SR'] ? String(r['SR']).split(/[,;]/)[0].trim() : null;
      custDict[c] = {
        flm: normFlm(r['FLM']),
        name: r['Customer Name'],
        sr: srMatch[srFirst] || null,
      };
    }
  });

  const allocated = [];
  daily.forEach(r => {
    // === FIX: case-insensitive Ethical filter ===
    if (!isEthical(r['Dep'])) return;
    const matCode = Number(r['Material Code']);
    const mat = matMap[matCode];
    if (!mat || !mat.sb || KPIS.indexOf(mat.sb) < 0) return;
    const monthNum = mapMonth(r['Short Cut']);
    if (!monthNum) return;
    if (numVal(r['Year']) !== 2026) return;
    const cust = Number(r['Customer Code']);
    const cat = mat.cat;
    // Sales actuals match the Export tab exactly — negatives (returns/credit
    // notes) are included, same as the raw data.
    const sales = numVal(r['Total Act. Sales']);
    const flm = normFlm(r['FLM']);
    const srFirst = r['SR'] ? String(r['SR']).split(/[,;]/)[0].trim() : null;
    const dailySr = srMatch[srFirst] || null;

    if (SHARED[cust]) {
      const matching = SHARED[cust].filter(rule => rule.cat === cat);
      if (matching.length > 0) {
        matching.forEach(rule => {
          allocated.push({
            m: monthNum, sr: rule.sr, f: srToFlm[rule.sr],
            k: mat.sb, cust: cust, cat: cat,
            v: sales * rule.w,
          });
        });
        return;
      }
    }
    if (dailySr) {
      allocated.push({
        m: monthNum, sr: dailySr, f: srToFlm[dailySr] || flm,
        k: mat.sb, cust: cust, cat: cat,
        v: sales,
      });
    }
  });

  const aggMap = {};
  allocated.forEach(r => {
    const key = r.m + '|' + r.sr + '|' + r.f + '|' + r.k;
    aggMap[key] = (aggMap[key] || 0) + r.v;
  });
  const actuals = [];
  Object.keys(aggMap).forEach(key => {
    const parts = key.split('|');
    const v = Math.round(aggMap[key]);
    if (v > 0) actuals.push({
      m: Number(parts[0]),
      sr: Number(parts[1]),
      f: parts[2] === 'null' || parts[2] === 'undefined' ? null : parts[2],
      k: parts[3],
      v: v,
    });
  });

  // =========================================================================
  // === TOTAL ETHICAL SALES per month × SR (the headline "Sales" figure) ===
  // The Sales number must equal the Daily Sales sheet's Ethical total EXACTLY,
  // so — unlike the per-product KPIs — it does NOT filter by tracked material /
  // sub-brand and does NOT drop unmatched-SR rows. Every Ethical 2026 row is
  // counted once. Sales are attributed to SR (shared customers split by category
  // weight, else the row's SR, else the customer's mapped SR); anything still
  // unattributable lands in the 'Unassigned' bucket (sr code 0) so bySr always
  // sums to total and the grand total ties out to the sheet.
  // =========================================================================
  const salesByMonth = {};
  MONTHS.forEach(m => { salesByMonth[m] = { bySr: {}, byFlm: {}, total: 0 }; });
  daily.forEach(r => {
    if (!isEthical(r['Dep'])) return;
    const monthNum = mapMonth(r['Short Cut']);
    if (!monthNum) return;
    if (numVal(r['Year']) !== 2026) return;
    const sales = numVal(r['Total Act. Sales']);
    if (!sales) return;
    const cust = Number(r['Customer Code']);
    const mat = matMap[Number(r['Material Code'])];
    const cat = mat ? mat.cat : null;
    const bucket = salesByMonth[monthNum];
    bucket.total += sales;
    // FLM breakdown comes straight from the row's FLM column, so by-FLM ties to
    // the raw data exactly (every row's full sale lands in one FLM).
    const rowFlm = normFlm(r['FLM']) || 'Unassigned';
    bucket.byFlm[rowFlm] = (bucket.byFlm[rowFlm] || 0) + sales;

    const addToSr = (sr, amt) => {
      bucket.bySr[sr] = (bucket.bySr[sr] || 0) + amt;
    };

    let done = false;
    if (SHARED[cust] && cat) {
      const matching = SHARED[cust].filter(rule => rule.cat === cat);
      if (matching.length > 0) {
        matching.forEach(rule => addToSr(rule.sr, sales * rule.w));
        done = true;
      }
    }
    if (!done) {
      const srFirst = r['SR'] ? String(r['SR']).split(/[,;]/)[0].trim() : null;
      const sr = srMatch[srFirst] || (custDict[cust] && custDict[cust].sr) || 0;
      addToSr(sr, sales);
    }
  });


  //  - dailyByPeriodCust: ETHICAL only — used by Active 3-mo.
  //  - shopByPeriodCust:  ALL departments — used by Shop Around. EXCEPTION:
  //    shop-around customers usually book their purchases under TRADE (not
  //    Ethical), so Shop Around must count their sales regardless of department.
  const dailyByPeriodCust = {};
  const shopByPeriodCust = {};
  const srByPeriodCust = {};  // period -> cust -> { srCode: true } — who actually sold (Ethical)
  daily.forEach(function (r) {
    const m = mapMonth(r['Short Cut']);
    const period = (r['Year'] || 0) * 100 + m;
    const c = Number(r['Customer Code']);
    if (!c) return;
    const sales = numVal(r['Total Act. Sales']);  // negatives included (matches Export)
    if (!shopByPeriodCust[period]) shopByPeriodCust[period] = {};
    shopByPeriodCust[period][c] = (shopByPeriodCust[period][c] || 0) + sales;
    if (!isEthical(r['Dep'])) return;  // everything below is Ethical-only
    if (!dailyByPeriodCust[period]) dailyByPeriodCust[period] = {};
    dailyByPeriodCust[period][c] = (dailyByPeriodCust[period][c] || 0) + sales;
    // Track which SR actually sold to this customer (from the daily SR column),
    // so Active can credit the selling SR — matching how Sales attributes.
    const srName = r['SR'] ? String(r['SR']).split(/[,;]/)[0].trim() : null;
    const srCode = srName ? srMatch[srName] : null;
    if (srCode) {
      if (!srByPeriodCust[period]) srByPeriodCust[period] = {};
      if (!srByPeriodCust[period][c]) srByPeriodCust[period][c] = {};
      srByPeriodCust[period][c][srCode] = true;
    }
  });

  // === SHOP AROUND ===
  // Step 1: Count how many distinct SR codes share each customer (across Shop Around Details)
  const shopSharedCount = {};
  {
    const seenPairs = {};
    shopRaw.forEach(r => {
      const c = Number(r['Customer Code']);
      const sr = Number(r['SR Code']);
      if (!c || !sr) return;
      if (!seenPairs[c]) seenPairs[c] = {};
      seenPairs[c][sr] = true;
    });
    Object.keys(seenPairs).forEach(c => {
      shopSharedCount[Number(c)] = Object.keys(seenPairs[c]).length;
    });
  }
  
  const shopByMonth = {};
  MONTHS.forEach(m => {
    const period = 202600 + m;
    const items = [];
    shopRaw.forEach(r => {
      const srPrimary = Number(r['SR Code']);
      if (!srPrimary) return;
      const rawCust = r['Customer Code'];
      const numCust = Number(rawCust);
      const isPlaceholder = !numCust || isNaN(numCust);
      
      let target = 0;
      for (const col in r) {
        const colDate = new Date(col);
        if (!isNaN(colDate.getTime()) && colDate.getMonth() === (m-1) && colDate.getFullYear() === 2026) {
          target = Number(r[col]) || 0;
          break;
        }
      }
      
      const fullActual = isPlaceholder ? 0
        : ((shopByPeriodCust[period] && shopByPeriodCust[period][numCust]) || 0);
      
      const shareCount = (!isPlaceholder && shopSharedCount[numCust]) ? shopSharedCount[numCust] : 1;
      
      const actual = shareCount > 1 ? (fullActual / shareCount) : fullActual;
      
      if (target === 0 && actual === 0) return;
      
      items.push({
        sr: srPrimary,
        f: srToFlm[srPrimary] || normFlm(r['FLM']),
        c: isPlaceholder ? null : numCust,
        cn: isPlaceholder
          ? (String(rawCust || 'New') + ' Prospect')
          : String(r['Customer Name'] || ''),
        t: Math.round(target),
        v: Math.round(actual),
        isNew: isPlaceholder,
        shared: shareCount > 1,
        coverCount: shareCount,
      });
    });
    shopByMonth[m] = items;
  });

  const activeByMonth = {};
  MONTHS.forEach(m => {
    const periods = [];
    [m-2, m-1, m].forEach(prev => {
      if (prev >= 1) periods.push(202600 + prev);
      else periods.push(202500 + (12 + prev));
    });
    // Active 3-mo is the ONLY place negative sales are applied: sum each
    // customer's NET sales across the window and count them active only if the
    // net is positive (returns/credit notes net out a customer's activity).
    // Reuses dailyByPeriodCust (built once above) instead of re-scanning the
    // whole daily sheet 12 times — much faster on large data.
    const custNet = {};
    periods.forEach(function (p) {
      const pc = dailyByPeriodCust[p];
      if (!pc) return;
      Object.keys(pc).forEach(function (c) { custNet[c] = (custNet[c] || 0) + pc[c]; });
    });
    // SRs who actually sold to each customer anywhere in this 3-month window.
    const soldSrs = {};
    periods.forEach(function (p) {
      const sc = srByPeriodCust[p];
      if (!sc) return;
      Object.keys(sc).forEach(function (cc) {
        if (!soldSrs[cc]) soldSrs[cc] = {};
        Object.keys(sc[cc]).forEach(function (sr) { soldSrs[cc][sr] = true; });
      });
    });
    const custs = {};
    Object.keys(custNet).forEach(c => { if (custNet[c] > 0) custs[c] = true; });
    const flmCount = {}, srCount = {};
    const custSrsMap = {}; // customer code -> [sr codes crediting them] (one source of truth for all tabs)
    Object.keys(custs).forEach(cs => {
      const c = Number(cs);
      // Credit the SR(s) who actually sold to the customer (daily SR column) plus
      // any shared partners; fall back to the customer master only if neither
      // exists. This is why a customer Heng Norm sold to now counts for him.
      let srs = [];
      if (soldSrs[cs]) srs = srs.concat(Object.keys(soldSrs[cs]).map(Number));
      if (SHARED[c]) srs = srs.concat(SHARED[c].map(r => r.sr));
      if (srs.length === 0 && custDict[c] && custDict[c].sr) srs = [custDict[c].sr];
      // Dedupe: count each customer ONCE per SR even if they're shared with the
      // same SR across multiple categories (otherwise the active count inflates).
      srs = srs.filter(function (v, i) { return srs.indexOf(v) === i; });
      if (srs.length === 0) {
        flmCount['Unassigned'] = (flmCount['Unassigned'] || 0) + 1;
        return;
      }
      custSrsMap[c] = srs;
      // Counts are SUMMED, not distinct: a customer shared across SRs counts once
      // for EACH SR (so two SRs sharing the same 3 customers both show 3), and the
      // FLM total is the sum of its SRs' counts (not a distinct customer count).
      srs.forEach(sr => {
        srCount[sr] = (srCount[sr] || 0) + 1;
        const f = srToFlm[sr];
        if (f) flmCount[f] = (flmCount[f] || 0) + 1;
      });
    });
    // Overall total = sum of every FLM bucket (= sum of all SR counts + Unassigned),
    // matching the summed semantics above rather than a distinct customer count.
    let activeTotalSum = 0;
    Object.keys(flmCount).forEach(f => { activeTotalSum += flmCount[f]; });
    activeByMonth[m] = {
      total: activeTotalSum,
      byFlm: flmCount, bySr: srCount,
      customers: Object.keys(custs).map(Number),
      custSrs: custSrsMap,
    };
  });

  const findCol = (obj, names) => {
    for (let i = 0; i < names.length; i++) if (names[i] in obj) return names[i];
    return null;
  };

  const activeTargetByMonth = {};
  const MONTH_LABELS_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  MONTHS.forEach(m => {
    const lbl = MONTH_LABELS_SHORT[m-1];
    const candidates = [
      lbl + '-26\nFINAL', lbl + '-26 FINAL', lbl + '_Final',
      lbl + '-26\n(Act)', lbl + '-26 (Act)', lbl + '_Act',
      lbl + '-26', lbl + ' 2026', lbl + '-2026',
    ];
    const bySr = {}, byFlm = {};
    acTarget.forEach(r => {
      const sr = Number(r['ID']);
      if (!sr) return;
      const col = findCol(r, candidates);
      const v = col ? Number(r[col]) || 0 : 0;
      bySr[sr] = v;
      const f = normFlm(r['FLM']);
      if (f) byFlm[f] = (byFlm[f] || 0) + v;
    });
    activeTargetByMonth[m] = { bySr: bySr, byFlm: byFlm };
  });

  // =========================================================================
  // === NEW LISTING — Name-based matching (case-insensitive) ===
  // =========================================================================
  // Build entity -> set of periods (from Ethical rows only).
  // Customer Code is the primary key, so the same customer counts as ONE entity
  // even when the Customer Name text varies across rows. Fall back to name only
  // for placeholder/no-code rows (281000000) so genuine new prospects still work.
  const entityPeriods = {};   // key -> { periods:{}, code:number|null, name:string|null }

  daily.forEach(r => {
    // === FIX: case-insensitive Ethical filter ===
    if (!isEthical(r['Dep'])) return;
    const nm = normName(r['Customer Name']);
    const sc = String(r['Short Cut'] || '').trim();
    const mn = mapMonth(sc);
    if (!mn) return;
    const yr = Number(r['Year']) || 0;
    if (!yr) return;
    const p = yr * 100 + mn;

    const c = Number(r['Customer Code']);
    const hasRealCode = c && c !== 281000000;
    const key = hasRealCode ? ('C:' + c) : (nm ? ('N:' + nm) : null);
    if (!key) return;

    if (!entityPeriods[key]) entityPeriods[key] = { periods: {}, code: hasRealCode ? c : null, name: nm || null };
    entityPeriods[key].periods[p] = true;
    if (hasRealCode) entityPeriods[key].code = c;
    if (nm) entityPeriods[key].name = nm;
  });
  Logger.log('entityPeriods map size: ' + Object.keys(entityPeriods).length);

  const newByMonth = {};
  MONTHS.forEach(m => {
    const targetPeriod = 202600 + m;
    
    const priorPeriods = {};
    for (let i = 1; i <= 11; i++) {
      const dm = m - i;
      if (dm >= 1) priorPeriods[202600 + dm] = true;
      else priorPeriods[202500 + (12 + dm)] = true;
    }

    const items = [];
    const flmCount = {}, srCount = {};

    Object.keys(entityPeriods).forEach(key => {
      const ent = entityPeriods[key];
      const purchases = ent.periods;

      if (!purchases[targetPeriod]) return;

      let hasPrior = false;
      for (const pStr in priorPeriods) {
        if (purchases[Number(pStr)]) { hasPrior = true; break; }
      }
      if (hasPrior) return;

      const repCode = ent.code;

      let srs = [];
      if (repCode && SHARED[repCode] && SHARED[repCode].length > 0) {
        const seen = {};
        SHARED[repCode].forEach(rule => {
          if (!seen[rule.sr]) { seen[rule.sr] = true; srs.push(rule.sr); }
        });
      } else if (repCode && custDict[repCode] && custDict[repCode].sr) {
        srs = [custDict[repCode].sr];
      }

      srs.forEach(sr => {
        srCount[sr] = (srCount[sr] || 0) + 1;
      });

      const flmsSeen = {};
      srs.forEach(sr => {
        const f = srToFlm[sr];
        if (f) flmsSeen[f] = true;
      });
      if (Object.keys(flmsSeen).length === 0) {
        flmCount['Unassigned'] = (flmCount['Unassigned'] || 0) + 1;
      } else {
        Object.keys(flmsSeen).forEach(f => {
          flmCount[f] = (flmCount[f] || 0) + 1;
        });
      }

      items.push({
        c: repCode || null,
        n: (repCode && custDict[repCode] && custDict[repCode].name) || ent.name,
        f: srs[0] ? srToFlm[srs[0]] : null,
        sr: srs[0] || null,
        srs: srs,
      });
    });

    newByMonth[m] = { items: items, byFlm: flmCount, bySr: srCount };
  });

  const newTargetByMonth = {};
  MONTHS.forEach(m => {
    const lbl = MONTH_LABELS_SHORT[m-1];
    const candidates = [
      lbl + '-26\nFINAL', lbl + '-26 FINAL', lbl + '_Final',
      lbl + '-26\n(Act)', lbl + '-26 (Act)', lbl + '_Act',
      lbl + '-26', lbl + ' 2026', lbl + '-2026',
    ];
    const bySr = {}, byFlm = {};
    nlTarget.forEach(r => {
      const sr = Number(r['ID']);
      if (!sr) return;
      const col = findCol(r, candidates);
      const v = col ? Number(r[col]) || 0 : 0;
      bySr[sr] = v;
      const f = normFlm(r['FLM']);
      if (f) byFlm[f] = (byFlm[f] || 0) + v;
    });
    newTargetByMonth[m] = { bySr: bySr, byFlm: byFlm };
  });

  const leadTargetByMonth = {};
  MONTHS.forEach(m => {
    const lbl = MONTH_LABELS_SHORT[m-1];
    const candidates = [
      lbl + '-26\nFINAL', lbl + '-26 FINAL', lbl + '_Final',
      lbl + '-26\n(Act)', lbl + '-26 (Act)', lbl + '_Act',
      lbl + '-26', lbl + ' 2026', lbl + '-2026',
    ];
    const bySr = {}, byFlm = {};
    ldTarget.forEach(r => {
      const sr = Number(r['ID']);
      if (!sr) return;
      const col = findCol(r, candidates);
      const v = col ? Number(r[col]) || 0 : 0;
      bySr[sr] = v;
      const f = normFlm(r['FLM']);
      if (f) byFlm[f] = (byFlm[f] || 0) + v;
    });
    leadTargetByMonth[m] = { bySr: bySr, byFlm: byFlm };
  });

  const leadTargetPerSr  = leadTargetByMonth[4] ? leadTargetByMonth[4].bySr  : {};
  const leadTargetPerFlm = leadTargetByMonth[4] ? leadTargetByMonth[4].byFlm : {};

  const leadActualByMonth = {};
  let ldActual = [];
  try { ldActual = readSheet(imsSS, TAB_NAMES.leadActual); }
  catch (e) { Logger.log('Note: Actual-Lead tab not found'); }
  MONTHS.forEach(m => {
    const lbl = MONTH_LABELS_SHORT[m-1];
    const candidates = [
      lbl + '-26\nFINAL', lbl + '-26 FINAL', lbl + '_Final',
      lbl + '-26\n(Act)', lbl + '-26 (Act)', lbl + '_Act',
      lbl + '-26', lbl + ' 2026', lbl + '-2026',
    ];
    const bySr = {}, byFlm = {};
    ldActual.forEach(r => {
      const sr = Number(r['ID']);
      if (!sr) return;
      const col = findCol(r, candidates);
      const v = col ? Number(r[col]) || 0 : 0;
      bySr[sr] = v;
      const f = normFlm(r['FLM']);
      if (f) byFlm[f] = (byFlm[f] || 0) + v;
    });
    leadActualByMonth[m] = { bySr: bySr, byFlm: byFlm };
  });

  const subBrandCategory = {
    'PED PWD':'PND','PED RPB':'PND','SIM':'PND','STC':'PND','SM':'PND',
    'ENS PWD':'MND','ENS RPB':'MND','GLU PWD':'MND','GLU RPB':'MND','PRO':'MND',
  };
  const mndPndByMonth = {};
  MONTHS.forEach(m => {
    let pndA = 0, mndA = 0, pndT = 0, mndT = 0;
    actuals.filter(r => r.m === m).forEach(r => {
      const c = subBrandCategory[r.k];
      if (c === 'PND') pndA += r.v;
      else if (c === 'MND') mndA += r.v;
    });
    allTargets.filter(r => r.m === m).forEach(r => {
      const c = subBrandCategory[r.k];
      if (c === 'PND') pndT += r.t;
      else if (c === 'MND') mndT += r.t;
    });
    mndPndByMonth[m] = { pnd_a: pndA, pnd_t: pndT, mnd_a: mndA, mnd_t: mndT };
  });

  const customers = [];
  Object.keys(custDict).forEach(c => {
    customers.push({
      c: Number(c), n: custDict[c].name,
      f: custDict[c].flm, sr: custDict[c].sr,
    });
  });

  const kpiCustByMonth = {};
  MONTHS.forEach(m => {
    const subAlloc = allocated.filter(r => r.m === m);
    const aggCust = {};
    subAlloc.forEach(r => {
      const key = r.sr + '|' + r.cust + '|' + r.k;
      aggCust[key] = (aggCust[key] || 0) + r.v;
    });
    const recs = [];
    Object.keys(aggCust).forEach(key => {
      const parts = key.split('|');
      const v = Math.round(aggCust[key]);
      if (v > 0) recs.push({
        sr: Number(parts[0]), c: Number(parts[1]),
        k: parts[2], v: v,
      });
    });
    kpiCustByMonth[m] = recs;
  });

  // =========================================================================
  // === CUSTOMER MONTHLY SALES MATRIX (Jan 2025 → latest period present) ===
  // Per-customer Ethical sales by Year-Month, used by the Customer Sales tab
  // (heatmap matrix + 3M/6M/12M "did not purchase" analysis). Negative sales
  // (returns/credit notes) are clamped to 0, consistent with the other actuals.
  // =========================================================================
  const custMonthly = {};
  let maxPeriod = 202501; // YYYYMM
  daily.forEach(r => {
    if (!isEthical(r['Dep'])) return;
    const c = Number(r['Customer Code']);
    if (!c) return;
    const mn = mapMonth(r['Short Cut']);
    if (!mn) return;
    const yr = Number(r['Year']) || 0;
    const period = yr * 100 + mn;
    if (period < 202501) return; // from Jan 2025 onward
    // Match the Export tab exactly — negatives included.
    const sales = numVal(r['Total Act. Sales']);
    if (!custMonthly[c]) {
      const srFirst = r['SR'] ? String(r['SR']).split(/[,;]/)[0].trim() : null;
      custMonthly[c] = {
        c: c,
        n: r['Customer Name'] || ('Customer ' + c),
        sr: srMatch[srFirst] || (custDict[c] && custDict[c].sr) || null,
        f: normFlm(r['FLM']) || (custDict[c] && custDict[c].flm) || null,
        p: {},
      };
    }
    custMonthly[c].p[period] = (custMonthly[c].p[period] || 0) + sales;
    if (period > maxPeriod) maxPeriod = period;
  });

  return {
    kpis: KPIS,
    flms: ['Chay Mengkong','In Lena','Sem Sokhom','Thong Kanha','Um Phana'],
    srs: srMaster, customers: customers,
    targets: allTargets, actuals: actuals,
    salesByMonth: salesByMonth,
    months: MONTHS,
    customerMonthly: Object.keys(custMonthly).map(k => custMonthly[k]),
    customerMonthlyMaxPeriod: maxPeriod,
    kpiCustByMonth: kpiCustByMonth,
    shopByMonth: shopByMonth,
    activeByMonth: activeByMonth,
    activeTargetByMonth: activeTargetByMonth,
    newByMonth: newByMonth,
    newTargetByMonth: newTargetByMonth,
    leadTargetPerSr: leadTargetPerSr,
    leadTargetPerFlm: leadTargetPerFlm,
    leadTargetByMonth: leadTargetByMonth,
    leadActualByMonth: leadActualByMonth,
    mndPndByMonth: mndPndByMonth,
    subBrandCategory: subBrandCategory,
    sharedCustomers: SHARED,
    generatedAt: new Date().toISOString(),
  };
}

// =============================================================================
// HELPERS
// =============================================================================
function readSheet(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet/tab not found: "' + name + '" in "' + ss.getName() + '"');
  const range = sh.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h));
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = values[i][j];
    }
    rows.push(obj);
  }
  return rows;
}

// =============================================================================
// ADMIN UTILITIES — run from Apps Script editor
// =============================================================================
function clearCache() {
  clearChunkedCache_(CacheService.getScriptCache(), 'dashboard_data');
  Logger.log('Cache cleared');
}

function testConnections() {
  try {
    const ims = SpreadsheetApp.openById(SHEET_IDS.ims);
    Logger.log('OK IMS Targets: ' + ims.getName());
    const daily = SpreadsheetApp.openById(SHEET_IDS.daily);
    Logger.log('OK Daily Sales: ' + daily.getName());
    const mat = SpreadsheetApp.openById(SHEET_IDS.material);
    Logger.log('OK Material Code: ' + mat.getName());
    Logger.log('SUCCESS: All 3 sheets accessible');
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
  }
}

function testBuild() {
  const data = buildDashboardPayload();
  Logger.log('SRs: ' + data.srs.length);
  Logger.log('Customers: ' + data.customers.length);
  Logger.log('Targets: ' + data.targets.length);
  Logger.log('Actuals: ' + data.actuals.length);
  const aprT = data.targets.filter(r => r.m === 4).reduce((s, r) => s + r.t, 0);
  const aprA = data.actuals.filter(r => r.m === 4).reduce((s, r) => s + r.v, 0);
  Logger.log('April Target: ' + aprT + ' / Actual: ' + aprA);

  Logger.log('--- New Listing per month (name-based matching) ---');
  MONTHS.forEach(m => {
    const nb = data.newByMonth[m];
    if (!nb) return;
    const total = nb.items.length;
    const credits = Object.values(nb.bySr).reduce((s,v)=>s+v,0);
    if (total > 0) Logger.log('  Month ' + m + ': ' + total + ' new HCPs, ' + credits + ' SR credits');
  });

  Logger.log('--- Lead per month ---');
  MONTHS.forEach(m => {
    const tT = data.leadTargetByMonth[m] ? Object.values(data.leadTargetByMonth[m].bySr).reduce((s,v)=>s+v,0) : 0;
    const tA = data.leadActualByMonth[m] ? Object.values(data.leadActualByMonth[m].bySr).reduce((s,v)=>s+v,0) : 0;
    if (tT > 0 || tA > 0) {
      Logger.log('  Month ' + m + ': target=' + tT + ' / actual=' + tA);
    }
  });
}

// === Diagnostic to check if Trade rows are leaking through ===
function debugEthicalFilter() {
  const dailySS = SpreadsheetApp.openById(SHEET_IDS.daily);
  const daily = readSheet(dailySS, TAB_NAMES.daily);
  
  const uniqueDeps = {};
  daily.forEach(r => {
    const raw = r['Dep'];
    const key = '"' + String(raw === null || raw === undefined ? '(empty)' : raw) + '"';
    uniqueDeps[key] = (uniqueDeps[key] || 0) + 1;
  });
  
  Logger.log('=== Unique values in Dep column (and row counts) ===');
  Object.keys(uniqueDeps).forEach(k => {
    Logger.log('  ' + k + ' : ' + uniqueDeps[k] + ' rows');
  });
  
  // Now check April 2026 totals - both ways
  let aprEthicalSum = 0;
  let aprAllSum = 0;
  let aprEthicalRows = 0;
  let aprAllRows = 0;
  
  daily.forEach(r => {
    if (numVal(r['Year']) !== 2026) return;
    const mn = String(r['Short Cut']).trim();
    if (mn !== 'Apr' && mn !== 'April') return;
    const sales = numVal(r['Total Act. Sales']);
    aprAllSum += sales;
    aprAllRows++;
    const dep = String(r['Dep'] || '').trim().toUpperCase();
    if (dep === 'ETHICAL') {
      aprEthicalSum += sales;
      aprEthicalRows++;
    }
  });
  
  Logger.log('');
  Logger.log('=== April 2026 ===');
  Logger.log('  All rows: ' + aprAllRows + ' = ' + aprAllSum.toLocaleString());
  Logger.log('  Ethical rows: ' + aprEthicalRows + ' = ' + aprEthicalSum.toLocaleString());
  Logger.log('  Trade (everything else): ' + (aprAllSum - aprEthicalSum).toLocaleString());
}

function debugThayDararith() {
  const data = buildDashboardPayload();
  Logger.log('=== SRs matching "Thay Dararith" ===');
  data.srs.forEach(s => {
    if (String(s.name || '').toLowerCase().indexOf('thay dararith') >= 0) {
      Logger.log('  Code: ' + s.code + ' | Name: "' + s.name + '" | FLM: ' + s.flm);
    }
  });
  
  Logger.log('');
  Logger.log('=== April Actuals for Thay Dararith SRs ===');
  data.actuals.filter(r => r.m === 4).forEach(r => {
    const sr = data.srs.find(s => s.code === r.sr);
    if (sr && String(sr.name || '').toLowerCase().indexOf('thay dararith') >= 0) {
      Logger.log('  SR ' + r.sr + ' (' + sr.name + ') | KPI ' + r.k + ' | Value ' + r.v);
    }
  });
}
