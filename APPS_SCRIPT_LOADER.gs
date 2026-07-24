/**
 * SFE KPI Dashboard — SELF-UPDATING LOADER  (paste ONCE, then never again)
 * =====================================================================
 * WHAT THIS IS
 *   A tiny bootstrap. It does NOT contain the dashboard logic itself.
 *   Instead it PULLS the latest backend logic from your public site and
 *   runs it. That means: once you paste this and deploy a New version one
 *   time, every future backend change goes live automatically — you never
 *   open Apps Script again.
 *
 * ONE-TIME SETUP
 *   1. In the Apps Script editor, open your Code.gs (the server file).
 *   2. Select all, delete, and paste THIS file's contents.
 *   3. Deploy → Manage deployments → ✏️ → Version: New version → Deploy.
 *   That's the last manual deploy you'll ever do.
 *
 * SAFETY
 *   The logic is only cached as "known-good" after it compiles cleanly, and
 *   if a fetch ever fails the loader falls back to the last known-good copy —
 *   so a bad update can't take your live dashboard down.
 */

// Where the latest backend logic lives (public). First that responds wins.
var LOGIC_URLS = [
  'https://sfe-ethical-dashboard-v1-kpwp.vercel.app/backend-logic.js',
  'https://raw.githubusercontent.com/chanleakna/sfe-ethical-dashboard-v1/main/APPS_SCRIPT_Code.gs'
];
var LOGIC_PROP_KEY = 'lastGoodLogic';   // last known-good source (ScriptProperties)
var LOGIC_SRC_CACHE = 'logicSrc';        // short-lived source cache (CacheService)
var LOGIC_SRC_TTL = 600;                 // 10 min — how often we re-check the site

function _json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Fetch the logic source, preferring the short cache so most requests skip the
// network. Returns the source string, or null if nothing could be fetched.
function fetchLogicSource_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(LOGIC_SRC_CACHE);
  if (cached) return cached;
  for (var i = 0; i < LOGIC_URLS.length; i++) {
    try {
      var resp = UrlFetchApp.fetch(LOGIC_URLS[i], { muteHttpExceptions: true, followRedirects: true });
      if (resp.getResponseCode() === 200) {
        var txt = resp.getContentText();
        if (txt && txt.indexOf('buildDashboardPayload') >= 0) {
          cache.put(LOGIC_SRC_CACHE, txt, LOGIC_SRC_TTL);
          return txt;
        }
      }
    } catch (e) { /* try the next URL */ }
  }
  return null;
}

// Compile the source into a module object exposing the functions the loader
// needs. Throws if the code is broken or incomplete.
function compileLogic_(src) {
  var factory = new Function(
    src +
    '\n;return {' +
    ' getDashboardData: getDashboardData,' +
    ' handleLogin: handleLogin,' +
    ' getAccessLog: getAccessLog,' +
    ' debugNaAssign: (typeof debugNaAssign !== "undefined" ? debugNaAssign : null),' +
    ' clearCache: clearCache,' +
    ' setupAutoRefresh: setupAutoRefresh,' +
    ' refreshDashboardCache: refreshDashboardCache,' +
    ' onSheetChange: onSheetChange,' +
    ' importDailyFromEmail: (typeof importDailyFromEmail !== "undefined" ? importDailyFromEmail : null),' +
    ' CODE_VERSION: (typeof CODE_VERSION !== "undefined" ? CODE_VERSION : null)' +
    '};'
  );
  var mod = factory();
  if (!mod || typeof mod.getDashboardData !== 'function') {
    throw new Error('logic missing getDashboardData');
  }
  return mod;
}

// Load the logic, caching only KNOWN-GOOD source and falling back to the last
// known-good copy if the fetch/compile fails.
function loadLogic_() {
  var props = PropertiesService.getScriptProperties();
  var src = fetchLogicSource_();
  if (src) {
    try {
      var mod = compileLogic_(src);
      props.setProperty(LOGIC_PROP_KEY, src);  // only persist code that compiled
      return mod;
    } catch (e) {
      Logger.log('Fetched logic failed to compile, using last known-good: ' + e.message);
    }
  }
  var cached = props.getProperty(LOGIC_PROP_KEY);
  if (cached) return compileLogic_(cached);
  throw new Error('No backend logic available (fetch failed and no cached copy).');
}

function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action;
    if (action === 'loaderPing') return _json_({ ok: true, loader: 'v1', urls: LOGIC_URLS });
    // clearCache forces a fresh pull of both the code and the data.
    if (action === 'clearCache') CacheService.getScriptCache().remove(LOGIC_SRC_CACHE);

    var mod = loadLogic_();

    if (!action) {
      return HtmlService.createHtmlOutput(
        '<html><body style="font-family:sans-serif;padding:24px;color:#374151">' +
        '<h3>SFE KPI Dashboard — data API</h3>' +
        '<p>This URL serves data to the dashboard. Open the dashboard here: ' +
        '<a href="https://sfe-ethical-dashboard-v1-kpwp.vercel.app">sfe-ethical-dashboard-v1-kpwp.vercel.app</a></p>' +
        '</body></html>'
      ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    if (action === 'login')     return _json_(mod.handleLogin(e.parameter.code));
    if (action === 'data')      return _json_(mod.getDashboardData());
    if (action === 'accessLog') return _json_(mod.getAccessLog(e.parameter.code));
    if (action === 'health')    return _json_({ ok: true, codeVersion: mod.CODE_VERSION, loader: true, time: new Date().toISOString() });
    if (action === 'clearCache') { mod.clearCache(); return _json_({ ok: true, message: 'Cache + logic refreshed' }); }
    if (action === 'debugNa')   return _json_(mod.debugNaAssign ? mod.debugNaAssign() : { error: 'debugNa not available in current logic' });
    return _json_({ error: 'unknown action' });
  } catch (err) {
    return _json_({ error: err.message, stack: err.stack });
  }
}

// Trigger targets must be top-level functions in THIS project. They just
// delegate to the loaded logic, so the auto-refresh keeps working unchanged.
function refreshDashboardCache() { try { loadLogic_().refreshDashboardCache(); } catch (e) { Logger.log('refreshDashboardCache: ' + e.message); } }
function onSheetChange()         { try { loadLogic_().onSheetChange(); }        catch (e) { Logger.log('onSheetChange: ' + e.message); } }
function importDailyFromEmail()  { var m = loadLogic_(); return m.importDailyFromEmail ? m.importDailyFromEmail() : null; }

// Run this ONCE from the editor (Run menu) to install the auto-refresh triggers.
function setupAutoRefresh() { return loadLogic_().setupAutoRefresh(); }
