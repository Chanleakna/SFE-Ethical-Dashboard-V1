/**
 * SFE KPI Dashboard — Backend (Apps Script)
 * 
 * SETUP:
 * 1. This script is attached to your IMS Targets sheet
 * 2. It reads from 3 Google Sheets via their IDs (already configured below)
 * 3. Add 2 new tabs to your IMS Targets sheet: 'Users' and 'Shared_Customers'
 * 
 * DEPLOY:
 *   Deploy → New deployment → Web app → Execute as: Me → Access: Anyone → Deploy
 * 
 * RE-DEPLOY when you change this code:
 *   Deploy → Manage deployments → pencil icon → Version: New version → Deploy
 */

// =============================================================================
// CONFIGURATION — Sheet IDs (already filled in for your sheets)
// =============================================================================
const SHEET_IDS = {
  ims:      '1ocanfdc_9R6m0X0sH9W0dlCvc2U9EoeT50Wul61rrEk',
  daily:    '1-iJM9eU-cPqnjvHGBDKT8i59XN53rz9ykmyIqNpvkz8',
  material: '1GPxPYtkUD4kx2iCQMIUdctvbBrKS7XhCKp2JGy3NTr4',
};

const TAB_NAMES = {
  // In IMS Targets sheet
  targets:      'Target Set',
  shopDetails:  'Target - Shop Around Details',
  activeTarget: 'Target-Active Cus 3 Months',
  newTarget:    'Target - New Listing',
  leadTarget:   'Target-Lead',
  leadActual:   'Actual-Lead',            // ← Lead actuals (manual entry per month)
  users:        'Users',
  shared:       'Shared_Customers',
  
  // In Daily Sales sheet
  daily: 'Export',
  
  // In Master Material Code sheet
  materials: 'Maser Material Code',
};

// Months we support (1=Jan, 12=Dec). Year is 2026.
const MONTHS = [1,2,3,4,5,6,7,8,9,10,11,12];

const CACHE_SECONDS = 300;

// =============================================================================
// MAIN ENDPOINT
// =============================================================================
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    
    // No action = serve the dashboard HTML
    if (!action) {
      return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('SFE KPI Dashboard')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    }
    
    if (action === 'login')   return jsonResponse(handleLogin(e.parameter.code));
    if (action === 'data')    return jsonResponse(getDashboardData());
    if (action === 'health')  return jsonResponse({ ok: true, time: new Date().toISOString() });
    if (action === 'clearCache') {
      CacheService.getScriptCache().remove('dashboard_data');
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
      return {
        ok: true,
        user: {
          role: data[i][roleCol],
          name: data[i][nameCol],
          flm:  data[i][flmCol] || null,
        },
      };
    }
  }
  return { ok: false, error: 'Invalid code' };
}

// =============================================================================
// DATA BUILDER
// =============================================================================
function getDashboardData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dashboard_data');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const data = buildDashboardPayload();
  try {
    cache.put('dashboard_data', JSON.stringify(data), CACHE_SECONDS);
  } catch (e) {}
  return data;
}

function buildDashboardPayload() {
  const imsSS    = SpreadsheetApp.openById(SHEET_IDS.ims);
  const dailySS  = SpreadsheetApp.openById(SHEET_IDS.daily);
  const matSS    = SpreadsheetApp.openById(SHEET_IDS.material);
  
  const daily      = readSheet(dailySS,  TAB_NAMES.daily);
  const targets    = readSheet(imsSS,    TAB_NAMES.targets);
  const materials  = readSheet(matSS,    TAB_NAMES.materials);
  const shared     = readSheet(imsSS,    TAB_NAMES.shared);
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
  
  const normFlm = (s) => {
    if (!s) return null;
    return String(s).trim().replace(/Chhay/g, 'Chay').replace(/Phanna/g, 'Phana');
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
  
  const KPIS = ['SM','SIM','STC','PED PWD','PED RPB','ENS PWD','ENS RPB','GLU PWD','GLU RPB','PRO'];
  const ethical = targets.filter(r => r['Department'] === 'Ethical');
  
  const srSeen = {};
  const srMaster = [];
  ethical.forEach(r => {
    const code = Number(r['SR Code']);
    if (!code || srSeen[code]) return;
    srSeen[code] = true;
    srMaster.push({
      code: code,
      name: r['Seller Name'],
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
  
  const tokenize = (s) => {
    if (!s) return [];
    return String(s).toLowerCase()
      .replace(/\s*-\s*vacancy.*/g, '').replace(/&/g, ' ')
      .split(/[,;\s]+/).filter(p => p && p.length > 1);
  };
  
  const srMatch = {};
  const dailySrNames = {};
  daily.forEach(r => {
    if (r['SR']) {
      const first = String(r['SR']).split(/[,;]/)[0].trim();
      dailySrNames[first] = true;
    }
  });
  
  Object.keys(dailySrNames).forEach(daiName => {
    const dt = tokenize(daiName);
    if (dt.length === 0) return;
    let best = null, bestScore = 0;
    srMaster.forEach(sr => {
      const tt = tokenize(sr.name);
      if (tt.length === 0) return;
      const common = dt.filter(x => tt.indexOf(x) >= 0);
      if (common.length === 0) return;
      let score = common.length / Math.max(dt.length, tt.length);
      if (dt.every(x => tt.indexOf(x) >= 0) || tt.every(x => dt.indexOf(x) >= 0)) score += 0.5;
      if (score > bestScore) { bestScore = score; best = sr; }
    });
    if (best && bestScore >= 0.4) {
      srMatch[daiName] = best.code;
    }
  });
  
  // Add unmatched daily SRs as "Vacancy" entries (no FLM, sales tracked but not credited to team)
  Object.keys(dailySrNames).forEach(daiName => {
    if (srMatch[daiName]) return; // already matched
    // Generate a stable negative SR code from the name (so it persists across runs)
    let hash = 0;
    for (let i = 0; i < daiName.length; i++) {
      hash = ((hash << 5) - hash) + daiName.charCodeAt(i);
      hash = hash | 0;
    }
    const vacCode = -Math.abs(hash) % 100000;
    srMatch[daiName] = vacCode;
    srMaster.push({
      code: vacCode,
      name: daiName + ' (Vacancy)',
      flm: 'Vacancy',  // Special FLM label — not in main FLM list
    });
    srToFlm[vacCode] = 'Vacancy';
  });
  
  const MONTH_MAP = {Jan:1,Feb:2,Mar:3,April:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
  
  const custDict = {};
  daily.forEach(r => {
    const c = Number(r['Customer Code']);
    if (!c || r['Year'] !== 2026 || !r['FLM']) return;
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
    const matCode = Number(r['Material Code']);
    const mat = matMap[matCode];
    if (!mat || !mat.sb || KPIS.indexOf(mat.sb) < 0) return;
    const monthNum = MONTH_MAP[r['Short Cut']];
    if (!monthNum) return;
    const cust = Number(r['Customer Code']);
    const cat = mat.cat;
    const sales = Number(r['Total Act. Sales']) || 0;
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
  
  const dailyByPeriodCust = {};
  daily.forEach(r => {
    const m = MONTH_MAP[r['Short Cut']];
    const period = (r['Year'] || 0) * 100 + m;
    const c = Number(r['Customer Code']);
    if (!c) return;
    if (!dailyByPeriodCust[period]) dailyByPeriodCust[period] = {};
    dailyByPeriodCust[period][c] = (dailyByPeriodCust[period][c] || 0) + (Number(r['Total Act. Sales']) || 0);
  });
  
  const shopByMonth = {};
  MONTHS.forEach(m => {
    const period = 202600 + m;
    const items = [];
    shopRaw.forEach(r => {
      const sr = Number(r['SR Code']);
      if (!sr) return;
      
      // Customer Code may be a real number OR text like "New" (prospecting placeholder)
      const rawCust = r['Customer Code'];
      const numCust = Number(rawCust);
      const isPlaceholder = !numCust || isNaN(numCust);  // "New" or empty
      
      // Find target column matching this month
      let target = 0;
      for (const col in r) {
        const colDate = new Date(col);
        if (!isNaN(colDate.getTime()) && colDate.getMonth() === (m-1) && colDate.getFullYear() === 2026) {
          target = Number(r[col]) || 0;
          break;
        }
      }
      
      // Actual: only for real customer codes (placeholders stay at 0)
      const actual = isPlaceholder ? 0
        : ((dailyByPeriodCust[period] && dailyByPeriodCust[period][numCust]) || 0);
      
      if (target === 0 && actual === 0) return;
      
      items.push({
        sr: sr,
        f: srToFlm[sr] || normFlm(r['FLM']),
        c: isPlaceholder ? null : numCust,           // null = placeholder
        cn: isPlaceholder
          ? (String(rawCust || 'New') + ' Prospect')  // e.g., "New Prospect"
          : String(r['Customer Name'] || ''),
        t: Math.round(target),
        v: Math.round(actual),
        isNew: isPlaceholder,                         // flag for frontend
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
    const custs = {};
    daily.forEach(r => {
      const mn = MONTH_MAP[r['Short Cut']];
      const p = (r['Year'] || 0) * 100 + mn;
      if (periods.indexOf(p) < 0) return;
      const c = Number(r['Customer Code']);
      if (c) custs[c] = true;
    });
    const flmCount = {}, srCount = {};
    Object.keys(custs).forEach(cs => {
      const c = Number(cs);
      let srs = [];
      if (SHARED[c]) srs = SHARED[c].map(r => r.sr);
      else if (custDict[c] && custDict[c].sr) srs = [custDict[c].sr];
      if (srs.length === 0) {
        flmCount['Unassigned'] = (flmCount['Unassigned'] || 0) + 1;
        return;
      }
      const flmsForCust = {};
      srs.forEach(sr => {
        srCount[sr] = (srCount[sr] || 0) + 1;
        const f = srToFlm[sr];
        if (f) flmsForCust[f] = true;
      });
      Object.keys(flmsForCust).forEach(f => flmCount[f] = (flmCount[f] || 0) + 1);
    });
    activeByMonth[m] = {
      total: Object.keys(custs).length,
      byFlm: flmCount, bySr: srCount,
      customers: Object.keys(custs).map(Number),
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
    // Try multiple column header formats — Apr historically used "FINAL" naming;
    // other months may use either pattern as you fill them in
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
  
  const pairsByPeriod = {};
  daily.forEach(r => {
    const matCode = Number(r['Material Code']);
    const mat = matMap[matCode];
    if (!mat || KPIS.indexOf(mat.sb) < 0) return;
    const mn = MONTH_MAP[r['Short Cut']];
    const p = (r['Year'] || 0) * 100 + mn;
    const c = Number(r['Customer Code']);
    if (!c) return;
    if (!pairsByPeriod[p]) pairsByPeriod[p] = {};
    pairsByPeriod[p][c + '|' + mat.sb] = mat.cat;
  });
  
  const newByMonth = {};
  MONTHS.forEach(m => {
    const targetPeriod = 202600 + m;
    const priorPeriods = [];
    for (let i = 1; i <= 12; i++) {
      const dm = m - i;
      if (dm >= 1) priorPeriods.push(202600 + dm);
      else priorPeriods.push(202500 + (12 + dm));
    }
    const priorPairs = {};
    priorPeriods.forEach(p => {
      Object.keys(pairsByPeriod[p] || {}).forEach(k => priorPairs[k] = true);
    });
    const items = [];
    const flmCount = {}, srCount = {};
    const currentPairs = pairsByPeriod[targetPeriod] || {};
    Object.keys(currentPairs).forEach(key => {
      if (priorPairs[key]) return;
      const parts = key.split('|');
      const c = Number(parts[0]);
      const kpi = parts[1];
      const cat = currentPairs[key];
      let srs = [];
      if (SHARED[c]) srs = SHARED[c].filter(r => r.cat === cat).map(r => r.sr);
      if (srs.length === 0 && custDict[c] && custDict[c].sr) srs = [custDict[c].sr];
      srs.forEach(sr => srCount[sr] = (srCount[sr] || 0) + 1);
      const flmsSeen = {};
      srs.forEach(sr => {
        const f = srToFlm[sr];
        if (f) flmsSeen[f] = true;
      });
      if (Object.keys(flmsSeen).length === 0) {
        flmCount['Unassigned'] = (flmCount['Unassigned'] || 0) + 1;
      }
      Object.keys(flmsSeen).forEach(f => flmCount[f] = (flmCount[f] || 0) + 1);
      items.push({
        c: c, k: kpi, cat: cat,
        f: srs[0] ? srToFlm[srs[0]] : null,
        sr: srs[0] || null, srs: srs,
        n: (custDict[c] && custDict[c].name) || null,
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
  
  // ---- Lead targets per month (auto-extends as future month columns are filled) ----
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
  
  // Backward-compat: keep Apr-only fields used by older dashboard code paths
  const leadTargetPerSr  = leadTargetByMonth[4] ? leadTargetByMonth[4].bySr  : {};
  const leadTargetPerFlm = leadTargetByMonth[4] ? leadTargetByMonth[4].byFlm : {};
  
  // ---- Lead ACTUALS per month (from Actual-Lead tab) ----
  const leadActualByMonth = {};
  let ldActual = [];
  try {
    ldActual = readSheet(imsSS, TAB_NAMES.leadActual);
  } catch (e) {
    Logger.log('Note: Actual-Lead tab not found - Lead actuals will be empty');
  }
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
  
  return {
    kpis: KPIS,
    flms: ['Chay Mengkong','In Lena','Sem Sokhom','Thong Kanha','Um Phana'],
    srs: srMaster, customers: customers,
    targets: allTargets, actuals: actuals,
    months: MONTHS,
    kpiCustByMonth: kpiCustByMonth,
    shopByMonth: shopByMonth,
    activeByMonth: activeByMonth,
    activeTargetByMonth: activeTargetByMonth,
    newByMonth: newByMonth,
    newTargetByMonth: newTargetByMonth,
    leadTargetPerSr: leadTargetPerSr,         // backward-compat (Apr only)
    leadTargetPerFlm: leadTargetPerFlm,       // backward-compat (Apr only)
    leadTargetByMonth: leadTargetByMonth,     // NEW: per-month lead targets
    leadActualByMonth: leadActualByMonth,     // NEW: per-month lead actuals
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

// Run after editing data in sheets to force a fresh fetch
function clearCache() {
  CacheService.getScriptCache().remove('dashboard_data');
  Logger.log('Cache cleared');
}

// Run once to verify all 3 sheets are accessible
function testConnections() {
  try {
    const ims = SpreadsheetApp.openById(SHEET_IDS.ims);
    Logger.log('OK IMS Targets: ' + ims.getName());
    Logger.log('  Tabs: ' + ims.getSheets().map(s => s.getName()).join(' | '));
    
    const daily = SpreadsheetApp.openById(SHEET_IDS.daily);
    Logger.log('OK Daily Sales: ' + daily.getName());
    Logger.log('  Tabs: ' + daily.getSheets().map(s => s.getName()).join(' | '));
    
    const mat = SpreadsheetApp.openById(SHEET_IDS.material);
    Logger.log('OK Material Code: ' + mat.getName());
    Logger.log('  Tabs: ' + mat.getSheets().map(s => s.getName()).join(' | '));
    
    Logger.log('SUCCESS: All 3 sheets accessible');
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
  }
}

// Run to test full data build
function testBuild() {
  const data = buildDashboardPayload();
  Logger.log('SRs: ' + data.srs.length);
  Logger.log('Customers: ' + data.customers.length);
  Logger.log('Targets: ' + data.targets.length);
  Logger.log('Actuals: ' + data.actuals.length);
  const aprT = data.targets.filter(r => r.m === 4).reduce((s, r) => s + r.t, 0);
  const aprA = data.actuals.filter(r => r.m === 4).reduce((s, r) => s + r.v, 0);
  Logger.log('April Target: ' + aprT);
  Logger.log('April Actual: ' + aprA);
  Logger.log('April PND/MND: ' + JSON.stringify(data.mndPndByMonth[4]));
  
  // NEW: Lead targets and actuals per month
  Logger.log('--- Lead per month ---');
  [1,2,3,4,5,6,7,8,9,10,11,12].forEach(m => {
    const tT = data.leadTargetByMonth[m] ? Object.values(data.leadTargetByMonth[m].bySr).reduce((s,v)=>s+v,0) : 0;
    const tA = data.leadActualByMonth[m] ? Object.values(data.leadActualByMonth[m].bySr).reduce((s,v)=>s+v,0) : 0;
    if (tT > 0 || tA > 0) {
      Logger.log('  Month ' + m + ': target=' + tT + ' / actual=' + tA);
    }
  });
}
