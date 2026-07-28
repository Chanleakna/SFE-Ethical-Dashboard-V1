import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, LabelList
} from "recharts";

// === CONFIGURE THIS LINE ===
// Replace with your Apps Script Web App URL (ends in /exec)
const API_URL = "https://script.google.com/macros/s/AKfycbyXkZNFHARUMpbJ1i47BV5Dwaq1E-0cgOrn7CKIcRW8PMJuzywqH5WQIXWfQ7JKXiei/exec";

// Version tag shown in the header. Keep in step with CODE_VERSION in
// APPS_SCRIPT_Code.gs — the header shows both so a stale backend is obvious.
const BUILD_TAG = "R33";

// Refresh interval for live data (seconds). Daily-sales data doesn't change by
// the minute; a longer cadence keeps it live while reducing load on the slow
// sheet and the chance of catching a mid-rebuild partial read.
const REFRESH_INTERVAL = 300;
// localStorage key for the last-good payload (instant open while refetching)
const CACHE_KEY = "sfe_dashboard_payload_v1";

const KPI_DEFS = [
  { key: "SM",      label: "SM",      brand: "SM",        color: "#94a3b8" },
  { key: "SIM",     label: "SIM",     brand: "Similac",   color: "#0ea5e9" },
  { key: "STC",     label: "STC",     brand: "STC",       color: "#06b6d4" },
  { key: "PED PWD", label: "PED PWD", brand: "Pediasure", color: "#f59e0b" },
  { key: "PED RPB", label: "PED RPB", brand: "Pediasure", color: "#fbbf24" },
  { key: "ENS PWD", label: "ENS PWD", brand: "Ensure",    color: "#10b981" },
  { key: "ENS RPB", label: "ENS RPB", brand: "Ensure",    color: "#34d399" },
  { key: "GLU PWD", label: "GLU PWD", brand: "Glucerna",  color: "#8b5cf6" },
  { key: "GLU RPB", label: "GLU RPB", brand: "Glucerna",  color: "#a78bfa" },
  { key: "PRO",     label: "PRO",     brand: "Prosure",   color: "#ec4899" },
];

const SUB_BRANDS = [
  { name: "Pediasure", color: "#f59e0b", kpis: ["PED PWD", "PED RPB"] },
  { name: "Ensure",    color: "#10b981", kpis: ["ENS PWD", "ENS RPB"] },
  { name: "Glucerna",  color: "#8b5cf6", kpis: ["GLU PWD", "GLU RPB"] },
  { name: "Similac",   color: "#0ea5e9", kpis: ["SIM"] },
  { name: "STC",       color: "#06b6d4", kpis: ["STC"] },
  { name: "Prosure",   color: "#ec4899", kpis: ["PRO"] },
  { name: "SM",        color: "#94a3b8", kpis: ["SM"] },
];

// Division grouping of the 10 KPI sub-brands (from the Material Code master).
const KPI_CATEGORY = {
  "SM":"PND", "SIM":"PND", "STC":"PND", "PED PWD":"PND", "PED RPB":"PND",
  "ENS PWD":"MND", "ENS RPB":"MND", "GLU PWD":"MND", "GLU RPB":"MND", "PRO":"MND",
};

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

// === Excel export helper ===
// Uses the SheetJS library when available; otherwise falls back to a built-in
// CSV download that needs no external CDN (so it still works on locked-down
// company networks that block cdn.sheetjs.com). CSV opens directly in Excel.
const downloadCSV = (rows, filename) => {
  const headers = Object.keys(rows[0]);
  const esc = (v) => {
    const s = v == null ? "" : String(v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const csv = [headers.join(",")]
    .concat(rows.map(r => headers.map(h => esc(r[h])).join(",")))
    .join("\r\n");
  const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = String(filename || "export.xlsx").replace(/\.xlsx$/i, ".csv");
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};

const exportToExcel = (rows, filename, sheetName) => {
  if (!rows || rows.length === 0) {
    alert("No data to export with current filters.");
    return;
  }
  if (typeof XLSX === "undefined") {
    // Library blocked/unavailable — fall back to CSV so export still works.
    downloadCSV(rows, filename);
    return;
  }
  const ws = XLSX.utils.json_to_sheet(rows);
  const cols = Object.keys(rows[0]).map(k => {
    const maxLen = Math.max(k.length, ...rows.map(r => String(r[k] ?? "").length));
    return { wch: Math.min(maxLen + 2, 40) };
  });
  ws["!cols"] = cols;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName || "Data");
  XLSX.writeFile(wb, filename);
};

const ExportBtn = ({ onClick, label = "Export Excel" }) => (
  <button onClick={onClick} style={{
    background: "#16a34a", color: "#fff", border: "none",
    borderRadius: 6, padding: "5px 11px", fontSize: 11,
    fontWeight: 600, cursor: "pointer", marginLeft: "auto",
  }}>⬇ {label}</button>
);

const fmt = (n) => {
  if (n == null || isNaN(n)) return "—";
  if (n === 0) return "0";
  return Math.round(n).toLocaleString();
};
const pct = (a, t) => (t > 0 ? (a / t) * 100 : 0);
// Latest month (1-12) that has any sales — used as the default month so the
// dashboard opens on the most recent month that actually has data (and only
// advances to a new month once that month has its first sales).
const latestDataMonth = (raw) => {
  const fallback = new Date().getMonth() + 1;
  if (!raw || !Array.isArray(raw.actuals)) return fallback;
  const totals = {};
  raw.actuals.forEach(r => { totals[r.m] = (totals[r.m] || 0) + (r.v || 0); });
  const months = Object.keys(totals).map(Number).filter(m => totals[m] > 0);
  if (!months.length) return fallback;
  return Math.max.apply(null, months);
};
const uniqArr = (a) => a.filter((v, i) => a.indexOf(v) === i);
// Build a "by SR → Customer, one column per month Jan→latest + YTD Total" export.
// recordsByMonth(m) returns an array of { sr, c, cn?, v }; values are summed per (sr, c, month).
function ytdCustomerExport(raw, year, flmSel, srSel, recordsByMonth, filename, sheet) {
  const fMatch = (f) => !flmSel.length || flmSel.indexOf(f) >= 0;
  const sMatch = (code) => !srSel.length || srSel.indexOf(Number(code)) >= 0;
  const latest = latestDataMonth(raw);
  const months = []; for (let m = 1; m <= latest; m++) months.push(m);
  const srName = {}, srFlm = {};
  (raw.srs || []).forEach(s => { srName[s.code] = s.name; srFlm[s.code] = s.flm; });
  const custName = {}; (raw.customers || []).forEach(c => { custName[c.c] = c.n; });
  const data = {};
  months.forEach(m => (recordsByMonth(m) || []).forEach(r => {
    if (r.c == null || r.sr == null) return;
    const key = r.sr + "|" + r.c;
    if (!data[key]) data[key] = { cn: r.cn, vals: {} };
    if (r.cn && !data[key].cn) data[key].cn = r.cn;
    data[key].vals[m] = (data[key].vals[m] || 0) + (r.v || 0);
  }));
  const list = Object.keys(data).map(key => {
    const parts = key.split("|"); const sr = Number(parts[0]), c = Number(parts[1]);
    let ytd = 0; months.forEach(m => { ytd += (data[key].vals[m] || 0); });
    return { sr, c, cn: data[key].cn || custName[c] || "", flm: srFlm[sr] || "—", srNm: srName[sr] || ("SR " + sr), vals: data[key].vals, ytd };
  })
    .filter(r => fMatch(r.flm))
    .filter(r => sMatch(r.sr))
    .sort((a, b) => (a.sr - b.sr) || (b.ytd - a.ytd));
  const out = list.map(r => {
    const o = { "FLM": r.flm, "SR Code": r.sr, "SR Name": r.srNm, "Customer Code": r.c, "Customer Name": r.cn };
    months.forEach(m => { o[MONTH_NAMES[m - 1] + "-" + String(year).slice(2)] = Math.round(r.vals[m] || 0); });
    o["YTD Total"] = Math.round(r.ytd);
    return o;
  });
  exportToExcel(out, filename + "_Jan-" + MONTH_NAMES[latest - 1] + "-" + year + ".xlsx", sheet);
}
const pctColor = (p) => {
  if (p >= 100) return "#059669";
  if (p >= 80)  return "#d97706";
  if (p > 0)    return "#dc2626";
  return "#9ca3af";
};
const heatBg = (p) => {
  if (p === 0 || p == null) return "#f3f4f6";
  if (p >= 120) return "#a7f3d0";
  if (p >= 100) return "#d1fae5";
  if (p >= 80)  return "#fef3c7";
  if (p >= 50)  return "#fee2e2";
  return "#fecaca";
};
const heatFg = (p) => {
  if (p === 0 || p == null) return "#9ca3af";
  if (p >= 100) return "#065f46";
  if (p >= 80)  return "#92400e";
  return "#991b1b";
};

export default function App() {
  const [user, setUser] = useState(null);
  const [raw, setRaw] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);

  // Fetch data after login
  const hasDataRef = useRef(false);
  const lastMagRef = useRef(null);   // magnitude (total actuals) of last accepted payload
  const lowStreakRef = useRef(0);    // consecutive suspiciously-low reads
  const [refreshNonce, setRefreshNonce] = useState(0);
  const forceRef = useRef(false);
  // Force-refresh: clears the 6h server cache, then refetches — so daily data
  // you just entered shows immediately instead of waiting for the cache.
  const forceRefresh = () => { forceRef.current = true; setRefreshNonce(n => n + 1); };
  useEffect(() => {
    if (!user) return;
    let cancelled = false;

    // A complete payload must carry the core arrays. We only ever render / cache
    // data that passes this check, so a slow or half-built Apps Script response
    // can't poison the view with empty (very low) numbers.
    const isComplete = (d) => d && !d.error && Array.isArray(d.srs) && d.srs.length > 0
      && Array.isArray(d.actuals) && Array.isArray(d.targets);
    // Total actuals — a stable magnitude used to spot mid-rebuild partial reads.
    const magnitude = (d) => (d.actuals || []).reduce((s, r) => s + (r.v || 0), 0);

    const accept = (data) => {
      setRaw(data);
      hasDataRef.current = true;
      lastMagRef.current = magnitude(data);
      lowStreakRef.current = 0;
      setError(null);
      try { localStorage.setItem(CACHE_KEY, JSON.stringify(data)); } catch (e) { /* quota — skip */ }
    };

    // Instant open: render the last-good payload from cache while we refetch in
    // the background. The Apps Script call recomputes the whole sheet and can be
    // slow, so this keeps the dashboard from blocking on a blank loading screen.
    try {
      const cached = localStorage.getItem(CACHE_KEY);
      if (cached) {
        const data = JSON.parse(cached);
        if (isComplete(data)) { setRaw(data); hasDataRef.current = true; lastMagRef.current = magnitude(data); }
      }
    } catch (e) { /* ignore corrupt / oversized cache */ }

    const fetchData = async () => {
      if (!hasDataRef.current) setLoading(true);
      try {
        const res = await fetch(API_URL + "?action=data");
        if (!res.ok) throw new Error("HTTP " + res.status);
        const data = await res.json();
        if (cancelled) return;
        setLoading(false);
        if (!isComplete(data)) {
          // Nothing good to show yet — surface why instead of rendering empties.
          if (!hasDataRef.current) setError((data && data.error) || "Data looks incomplete — the Apps Script may still be deploying or returned partial data. Refresh in a moment.");
          return;
        }
        const prev = lastMagRef.current;
        const mag = magnitude(data);
        // Guard against a mid-rebuild partial read: a sudden >50% collapse vs the
        // last good total is almost certainly the source sheet being rewritten,
        // not a real change. Keep showing the last good data. If it stays low for
        // 3 reads in a row, accept it (genuine corrections still get through).
        if (prev != null && prev > 0 && mag < prev * 0.5) {
          lowStreakRef.current += 1;
          if (lowStreakRef.current < 3) return; // ignore this suspicious read
        }
        accept(data);
      } catch (err) {
        if (cancelled) return;
        setLoading(false);
        // Don't replace a working (cached) view with an error screen on a
        // transient refresh failure — only surface errors when we have nothing.
        if (!hasDataRef.current) setError(err.message || "Failed to load data");
      }
    };

    const start = async () => {
      // On an explicit Refresh, clear the server cache first so we rebuild from
      // the current sheet (picks up daily data entered since the last cache).
      if (forceRef.current) {
        forceRef.current = false;
        setLoading(true);
        try { await fetch(API_URL + "?action=clearCache"); } catch (e) { /* ignore */ }
        if (cancelled) return;
      }
      fetchData();
    };
    start();
    const id = setInterval(fetchData, REFRESH_INTERVAL * 1000);
    return () => { cancelled = true; clearInterval(id); };
  }, [user, refreshNonce]);

  if (!user) return <Login onLogin={(u) => setUser(u)} />;
  
  if (error) return (
    <div style={{padding:40, fontFamily:"system-ui", color:"#dc2626"}}>
      <h2>Cannot load data</h2>
      <pre style={{background:"#fef2f2", padding:12, borderRadius:6}}>{error}</pre>
      <p style={{color:"#6b7280", fontSize:13}}>
        Check that the Apps Script is deployed and accessible.
      </p>
      <button onClick={() => setUser(null)}>Sign out</button>
    </div>
  );
  
  if (!raw) return (
    <div style={{padding:40, fontFamily:"system-ui", textAlign:"center", color:"#6b7280"}}>
      Loading data from Google Sheets…
    </div>
  );

  return (
    <ErrorBoundary onReset={() => setUser(null)}>
      <Dashboard user={user} raw={raw} onLogout={() => setUser(null)} onRefresh={forceRefresh} refreshing={loading} />
    </ErrorBoundary>
  );
}

// Catches render errors in any tab so a crash shows a message instead of a blank
// white screen — and surfaces the exact error so it can be diagnosed.
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(error) { return { error }; }
  componentDidCatch(error, info) { try { console.error("Dashboard crash:", error, info); } catch (e) {} }
  render() {
    if (this.state.error) {
      return (
        <div style={{padding:24, fontFamily:"system-ui", maxWidth:720, margin:"40px auto"}}>
          <h2 style={{color:"#dc2626", margin:"0 0 8px"}}>Something went wrong on this view</h2>
          <p style={{color:"#6b7280", fontSize:13}}>
            The page hit an error and stopped rendering. Try reloading. If it keeps happening,
            please screenshot the message below so it can be fixed.
          </p>
          <pre style={{background:"#fef2f2", color:"#991b1b", padding:12, borderRadius:8,
            fontSize:12, whiteSpace:"pre-wrap", overflowX:"auto"}}>
            {String(this.state.error && this.state.error.message || this.state.error)}
          </pre>
          <div style={{display:"flex", gap:8, marginTop:10}}>
            <button onClick={() => { this.setState({ error: null }); }}
              style={{background:"#2563eb", color:"#fff", border:"none", borderRadius:6, padding:"8px 14px", cursor:"pointer"}}>
              Try again
            </button>
            <button onClick={() => window.location.reload()}
              style={{background:"#fff", color:"#111827", border:"1px solid #d1d5db", borderRadius:6, padding:"8px 14px", cursor:"pointer"}}>
              Reload page
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

function Login({ onLogin }) {
  const [pw, setPw] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const submit = async () => {
    if (!pw) return;
    setLoading(true);
    setError("");
    try {
      const res = await fetch(API_URL + "?action=login&code=" + encodeURIComponent(pw));
      const data = await res.json();
      setLoading(false);
      if (data.ok && data.user) {
        onLogin({ ...data.user, code: pw });
      } else {
        setError(data.error || "Incorrect password");
        setPw("");
      }
    } catch (err) {
      setLoading(false);
      setError("Connection failed: " + err.message);
    }
  };

  return (
    <div style={{
      fontFamily: "'Inter', system-ui, sans-serif",
      background: "#f9fafb", minHeight: "100vh",
      display: "flex", alignItems: "center", justifyContent: "center",
      padding: 20,
    }}>
      <div style={{
        background: "#fff", border: "1px solid #e5e7eb", borderRadius: 12,
        padding: "32px 28px", width: "100%", maxWidth: 360,
        boxShadow: "0 1px 3px rgba(0,0,0,0.04)",
      }}>
        <div style={{textAlign: "center", marginBottom: 24}}>
          <div style={{
            display: "inline-flex", alignItems: "center", justifyContent: "center",
            width: 48, height: 48, borderRadius: 12, background: "#eff6ff",
            marginBottom: 12,
          }}>
            <div style={{fontSize: 22}}>🔐</div>
          </div>
          <h1 style={{fontSize: 18, fontWeight: 700, margin: 0, color: "#111827"}}>
            Ethical SFE Dashboard
          </h1>
          <p style={{fontSize: 12, color: "#6b7280", marginTop: 4}}>
            Enter your access code
          </p>
        </div>

        <input
          type="password"
          inputMode="numeric"
          pattern="[0-9]*"
          autoFocus
          value={pw}
          onChange={(e) => { setPw(e.target.value); setError(""); }}
          onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
          placeholder="••••"
          style={{
            width: "100%", boxSizing: "border-box",
            padding: "12px 14px", fontSize: 18, textAlign: "center",
            letterSpacing: 8, fontFamily: "monospace",
            border: error ? "1px solid #dc2626" : "1px solid #d1d5db",
            borderRadius: 8, outline: "none",
            marginBottom: 8,
          }}
        />

        {error && (
          <div style={{
            fontSize: 11, color: "#dc2626", textAlign: "center", marginBottom: 8,
          }}>{error}</div>
        )}

        <button onClick={submit} disabled={loading}
          style={{
            width: "100%", padding: "10px 14px", fontSize: 13, fontWeight: 600,
            background: loading ? "#93c5fd" : "#2563eb", color: "#fff", border: "none",
            borderRadius: 8, cursor: loading ? "wait" : "pointer",
          }}>
          {loading ? "Signing in…" : "Sign in"}
        </button>

        <div style={{fontSize: 10, color: "#9ca3af", marginTop: 16, textAlign: "center"}}>
          Contact your administrator if you don't have an access code.
        </div>
      </div>
    </div>
  );
}

function Dashboard({ user, raw, onLogout, onRefresh, refreshing }) {
  // Show only Ethical-master SRs. Auto-detected SRs (found in daily sales but
  // not in the Target Set) are assigned negative codes by the backend, so we
  // filter them out here too — they disappear even before the backend redeploys.
  const RAW = useMemo(() => ({
    ...raw,
    srs: (raw.srs || []).filter(s => Number(s.code) > 0),
  }), [raw]);
  const [tick, setTick] = useState(0);
  const [auto, setAuto] = useState(true);
  const [tab, setTab] = useState("summary");

  // Default to the current month WHEN it already has sales data; otherwise fall
  // back to the latest month that does — so a freshly-started month (still empty)
  // doesn't open on all-zeros. Auto-advances to the new month once its data lands.
  const [year, setYear] = useState(new Date().getFullYear());
  const [month, setMonth] = useState(() => Math.min(new Date().getMonth() + 1, latestDataMonth(raw)));
  // Multi-select filters — each is an array of selected values; [] means "All".
  const [flmSel, setFlmSel] = useState(user.role === "FLM" ? [user.flm] : []);
  const [srSel, setSrSel] = useState([]);
  const [custSel, setCustSel] = useState([]);
  const [custSearch, setCustSearch] = useState("");
  const allFlm = flmSel.length === 0, allSr = srSel.length === 0, allCust = custSel.length === 0;
  const flmMatch = (f) => allFlm || flmSel.indexOf(f) >= 0;
  const srMatch = (code) => allSr || srSel.indexOf(Number(code)) >= 0;
  const custMatch = (code) => allCust || custSel.indexOf(Number(code)) >= 0;
  const flmList = allFlm ? (RAW.flms || []) : flmSel;
  const [custLookup, setCustLookup] = useState(""); // Active 1M: customer code lookup
  const [lapseFilter, setLapseFilter] = useState("all"); // Customer Sales tab: all / 3 / 6 / 12
  const [custSortKey, setCustSortKey] = useState("total"); // Customer Sales sort column
  const [custSortDir, setCustSortDir] = useState("desc");

  // Expanded FLM rows
  const [expanded, setExpanded] = useState({});

  const [lastRefresh, setLastRefresh] = useState(new Date());

  // Admin-only access log
  const [accessLog, setAccessLog] = useState(null);
  const [accessErr, setAccessErr] = useState(null);
  useEffect(() => {
    if (tab !== "access" || user.role !== "Admin") return;
    let cancelled = false;
    setAccessErr(null);
    setAccessLog(null);
    fetch(API_URL + "?action=accessLog&code=" + encodeURIComponent(user.code || ""))
      .then(r => r.json())
      .then(d => {
        if (cancelled) return;
        if (d && d.ok) setAccessLog(d.rows || []);
        else setAccessErr((d && d.error) || "Failed to load access log");
      })
      .catch(e => { if (!cancelled) setAccessErr(e.message || "Failed to load access log"); });
    return () => { cancelled = true; };
  }, [tab, tick]);

  useEffect(() => {
    if (!auto) return;
    const id = setInterval(() => {
      setTick(t => t + 1);
      setLastRefresh(new Date());
    }, 5000);
    return () => clearInterval(id);
  }, [auto]);

  // Drop any selected SRs that aren't in the selected FLM(s) when FLM changes.
  useEffect(() => {
    if (allFlm) return;
    setSrSel(sel => sel.filter(code => {
      const sr = RAW.srs.find(s => s.code === Number(code));
      return sr && flmSel.indexOf(sr.flm) >= 0;
    }));
  }, [flmSel]);

  const C = useMemo(() => {
    // === Filtered SR set ===
    let srs = RAW.srs;
    if (!allFlm) srs = srs.filter(s => flmMatch(s.flm));
    if (!allSr) srs = srs.filter(s => srMatch(s.code));
    const srSet = new Set(srs.map(s => s.code));

    // === Target & actual for selected month ===
    const targetMap = new Map();
    RAW.targets.filter(r => r.m === month).forEach(r => {
      targetMap.set(r.sr + "|" + r.k, r.t);
    });
    const actualMap = new Map();
    RAW.actuals.filter(r => r.m === month).forEach(r => {
      const key = (r.sr ?? "X") + "|" + r.k;
      actualMap.set(key, (actualMap.get(key) || 0) + r.v);
    });

    // === SR scorecards (full detail) ===
    const srScorecards = srs.map(s => {
      const card = { code: s.code, name: s.name, flm: s.flm,
        kpis: {}, subBrands: {}, totalT: 0, totalA: 0 };
      card.mndT = 0; card.mndA = 0; card.pndT = 0; card.pndA = 0;
      KPI_DEFS.forEach(k => {
        const t = targetMap.get(s.code + "|" + k.key) || 0;
        const a = actualMap.get(s.code + "|" + k.key) || 0;
        card.kpis[k.key] = { target: t, actual: a, pct: pct(a, t), variance: a - t };
        card.totalT += t; card.totalA += a;
        const cat = KPI_CATEGORY[k.key];
        if (cat === "MND") { card.mndT += t; card.mndA += a; }
        else if (cat === "PND") { card.pndT += t; card.pndA += a; }
      });
      SUB_BRANDS.forEach(b => {
        let t = 0, a = 0;
        b.kpis.forEach(kk => {
          t += card.kpis[kk].target;
          a += card.kpis[kk].actual;
        });
        card.subBrands[b.name] = { target: t, actual: a, pct: pct(a, t) };
      });
      card.totalPct = pct(card.totalA, card.totalT);
      // Coverage actuals/targets per SR for current month
      const shopItems = (RAW.shopByMonth[month] || []).filter(x => x.sr === s.code);
      card.shopTarget = shopItems.reduce((s2, x) => s2 + x.t, 0);
      // Only count a customer's actual when THAT customer has a Shop Around target.
      card.shopActual = shopItems.reduce((s2, x) => s2 + (x.t > 0 ? x.v : 0), 0);
      card.activeActual = (RAW.activeByMonth[month]?.bySr?.[s.code]) || 0;
      card.activeTarget = (RAW.activeTargetByMonth[month]?.bySr?.[s.code]) || 0;
      card.newActual = (RAW.newByMonth[month]?.bySr?.[s.code]) || 0;
      card.newTarget = (RAW.newTargetByMonth[month]?.bySr?.[s.code]) || 0;
      card.leadTarget = (RAW.leadTargetByMonth && RAW.leadTargetByMonth[month])
        ? (RAW.leadTargetByMonth[month].bySr[s.code] || 0)
        : (RAW.leadTargetPerSr?.[s.code] || 0);
      card.leadActual = (RAW.leadActualByMonth && RAW.leadActualByMonth[month])
        ? (RAW.leadActualByMonth[month].bySr[s.code] || 0)
        : 0;
      return card;
    });

    // KPI totals (filtered scope)
    const kpiTotals = KPI_DEFS.map(k => {
      let t = 0, a = 0;
      srs.forEach(s => {
        t += targetMap.get(s.code + "|" + k.key) || 0;
        a += actualMap.get(s.code + "|" + k.key) || 0;
      });
      return { ...k, target: t, actual: a, pct: pct(a, t) };
    });

    const subBrandTotals = SUB_BRANDS.map(b => {
      let t = 0, a = 0;
      srs.forEach(s => {
        b.kpis.forEach(kk => {
          t += targetMap.get(s.code + "|" + kk) || 0;
          a += actualMap.get(s.code + "|" + kk) || 0;
        });
      });
      return { ...b, target: t, actual: a, pct: pct(a, t) };
    });

    const totalTarget = kpiTotals.reduce((s, k) => s + k.target, 0);
    const totalActual = kpiTotals.reduce((s, k) => s + k.actual, 0);

    // === Headline "Sales" = full Ethical sheet total (all materials, all rows) ===
    // Unlike totalActual (sum of tracked KPIs), this equals the Daily Sales sheet's
    // Ethical figure. For All/All we use the exact month total; when filtered we sum
    // the SRs in scope (unattributed sales sit in the Unassigned bucket, sr code 0).
    const sbm = (RAW.salesByMonth && RAW.salesByMonth[month]) || null;
    let salesActualFull = 0;
    if (sbm) {
      if (allFlm && allSr) salesActualFull = sbm.total;
      else srs.forEach(s => { salesActualFull += (sbm.bySr[s.code] || 0); });
    } else {
      salesActualFull = totalActual; // fallback before backend redeploy
    }

    // === FLM-level rollup with SR list nested ===
    const flmList = allFlm ? (RAW.flms || []) : flmSel;
    const flmRollup = flmList.map(f => {
      const flmSrs = RAW.srs.filter(s => s.flm === f &&
        (srMatch(s.code)));
      let t = 0, a = 0;
      flmSrs.forEach(s => {
        KPI_DEFS.forEach(k => {
          t += targetMap.get(s.code + "|" + k.key) || 0;
          a += actualMap.get(s.code + "|" + k.key) || 0;
        });
      });
      return {
        flm: f, srs: flmSrs, target: t, actual: a,
        pct: pct(a, t), variance: a - t,
      };
    });

    // === Coverage rollups ===
    const shopItems = (RAW.shopByMonth[month] || []).filter(x =>
      (flmMatch(x.f)) &&
      (srMatch(x.sr)) &&
      (custMatch(x.c))
    );

    // FLM-level coverage
    const flmCoverage = flmList.map(f => {
      const sList = RAW.srs.filter(s => s.flm === f);
      const shopT = shopItems.filter(x => x.f === f).reduce((s2, x) => s2 + x.t, 0);
      const shopA = shopItems.filter(x => x.f === f).reduce((s2, x) => s2 + x.v, 0);
      const activeT = RAW.activeTargetByMonth[month]?.byFlm?.[f] || 0;
      const activeA = RAW.activeByMonth[month]?.byFlm?.[f] || 0;
      const newT = RAW.newTargetByMonth[month]?.byFlm?.[f] || 0;
      const newA = RAW.newByMonth[month]?.byFlm?.[f] || 0;
      const leadT = (RAW.leadTargetByMonth && RAW.leadTargetByMonth[month])
        ? (RAW.leadTargetByMonth[month].byFlm[f] || 0)
        : (RAW.leadTargetPerFlm?.[f] || 0);
      const leadA = (RAW.leadActualByMonth && RAW.leadActualByMonth[month])
        ? (RAW.leadActualByMonth[month].byFlm[f] || 0)
        : 0;
      return { flm: f, srs: sList, shopT, shopA, activeT, activeA, newT, newA, leadT, leadA };
    });

    // === FLM × Division (MND / PND) rollup ===
    const flmDivision = flmList.map(f => {
      const cards = srScorecards.filter(c => c.flm === f);
      const sum = (key) => cards.reduce((s, c) => s + c[key], 0);
      return {
        flm: f, srs: cards,
        mndT: sum("mndT"), mndA: sum("mndA"),
        pndT: sum("pndT"), pndA: sum("pndA"),
      };
    });
    const divisionTotals = {
      mndT: srScorecards.reduce((s, c) => s + c.mndT, 0),
      mndA: srScorecards.reduce((s, c) => s + c.mndA, 0),
      pndT: srScorecards.reduce((s, c) => s + c.pndT, 0),
      pndA: srScorecards.reduce((s, c) => s + c.pndA, 0),
    };

    return {
      srs, srScorecards, kpiTotals, subBrandTotals,
      totalTarget, totalActual, salesActualFull,
      flmRollup, flmCoverage, shopItems,
      flmDivision, divisionTotals,
    };
  }, [tick, year, month, flmSel, srSel, custSel]);

  const overallPct = pct(C.totalActual, C.totalTarget);
  const salesPct = pct(C.salesActualFull, C.totalTarget);
  const activeTotal = (allFlm && allSr)
    ? RAW.activeByMonth[month]?.total
    : C.flmCoverage.reduce((s, r) => s + r.activeA, 0);
  const activeTotalT = C.flmCoverage.reduce((s, r) => s + r.activeT, 0);
  const newTotal = C.flmCoverage.reduce((s, r) => s + r.newA, 0);
  const newTotalT = C.flmCoverage.reduce((s, r) => s + r.newT, 0);
  // Shop Around actual is gated at CUSTOMER level (only a customer with a target
  // contributes its actual) so the headline card matches the Summary Coverage
  // matrix and the Shop Around tab. No-target customers contribute zero.
  const shopTotal = C.shopItems.reduce((s, x) => s + (x.t > 0 ? x.v : 0), 0);
  const shopTotalT = C.shopItems.reduce((s, x) => s + x.t, 0);
  const leadTotal = C.flmCoverage.reduce((s, r) => s + r.leadT, 0);
  const leadActualTotal = C.flmCoverage.reduce((s, r) => s + r.leadA, 0);

  // Totals for the new Active-1M and L&L tabs (respect the FLM / SR filters).
  const flmScope = flmList;
  const sumScope = (o) => {
    if (!o) return 0;
    if (!allSr) return srSel.reduce((s, code) => s + ((o.bySr && o.bySr[Number(code)]) || 0), 0);
    return flmScope.reduce((s, f) => s + ((o.byFlm && o.byFlm[f]) || 0), 0);
  };
  const active1AM = (RAW.active1ByMonth && RAW.active1ByMonth[month]) || { bySr:{}, byFlm:{}, total:0 };
  const active1TM = (RAW.activeTarget1ByMonth && RAW.activeTarget1ByMonth[month]) || { bySr:{}, byFlm:{} };
  const active1Total = (allFlm && allSr) ? (active1AM.total || 0) : sumScope(active1AM);
  const active1TotalT = sumScope(active1TM);
  const llTM = (RAW.llTargetByMonth && RAW.llTargetByMonth[month]) || { bySr:{}, byFlm:{} };
  const llAM = (RAW.llActualByMonth && RAW.llActualByMonth[month]) || { bySr:{}, byFlm:{} };
  const llTotalT = sumScope(llTM);
  const llActualTotal = sumScope(llAM);

  // SR options (filtered by FLM)
  const srOptions = (allFlm ? RAW.srs : RAW.srs.filter(s => flmMatch(s.flm)))
    .sort((a, b) => a.name.localeCompare(b.name));

  // Customer search filter
  const filteredCustomers = useMemo(() => {
    // Use the fullest customer set (everyone with sales since Jan 2025) so the
    // filter lists every customer, not just those with a 2026 FLM record.
    let custs = (RAW.customerMonthly && RAW.customerMonthly.length)
      ? RAW.customerMonthly : (RAW.customers || []);
    if (!allFlm) custs = custs.filter(c => flmMatch(c.f));
    if (!allSr) custs = custs.filter(c => srMatch(c.sr) || (c.srs || []).some(srMatch) || (c.asg || []).some(srMatch));
    if (custSearch) {
      const q = custSearch.toLowerCase();
      custs = custs.filter(c => String(c.c).includes(q) || String(c.n || "").toLowerCase().includes(q));
    }
    return custs; // no cap — the dropdown renders a window and its search covers all
  }, [flmSel, srSel, custSearch]);

  return (
    <div style={{
      fontFamily: "'Inter', system-ui, sans-serif",
      background: "#f9fafb", minHeight: "100vh", color: "#111827", padding: 16,
    }}>
      {/* Header */}
      <div style={{
        display: "flex", justifyContent: "space-between", alignItems: "center",
        marginBottom: 12, paddingBottom: 10, borderBottom: "1px solid #e5e7eb",
        flexWrap: "wrap", gap: 10,
      }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{
              width: 9, height: 9, borderRadius: "50%",
              background: auto ? "#10b981" : "#9ca3af",
              animation: auto ? "pulse 2s infinite" : "none",
            }} />
            <h1 style={{ fontSize: 20, fontWeight: 700, margin: 0, letterSpacing: "-.5px" }}>
              Ethical SFE — SR Performance Dashboard
            </h1>
          </div>
          <p style={{ fontSize: 11, color: "#6b7280", margin: "3px 0 0 19px" }}>
            {MONTH_NAMES[month-1]}-{String(year).slice(2)} · {C.srs.length} SRs in scope ·
            refreshed {lastRefresh.toLocaleTimeString()}
            <span style={{ marginLeft: 8, padding: "1px 6px", background: "#dcfce7",
              color: "#166534", borderRadius: 4, fontWeight: 600 }}>build {BUILD_TAG} ✓</span>
            {(() => {
              const dv = RAW.codeVersion;
              const ok = dv === BUILD_TAG;
              return (
                <span title={ok ? "Backend logic is up to date" : "Backend is behind — it will catch up automatically within ~10 min, or click Refresh / open ?action=clearCache to force it now"}
                  style={{ marginLeft: 6, padding: "1px 6px", borderRadius: 4, fontWeight: 600,
                    background: ok ? "#dcfce7" : "#fef3c7", color: ok ? "#166534" : "#92400e" }}>
                  data {dv || "?"} {ok ? "✓" : "⟳ updating…"}
                </span>
              );
            })()}
          </p>
        </div>
        <div style={{display:"flex", gap:6, alignItems:"center"}}>
          <div style={{
            fontSize:11, color:"#374151", padding:"4px 10px",
            background:"#f3f4f6", borderRadius:6, marginRight:4,
          }}>
            <span style={{color:"#6b7280"}}>{user.role}:</span>{" "}
            <strong>{user.name}</strong>
          </div>
          <button onClick={() => { if (onRefresh) onRefresh(); setTick(t => t + 1); setLastRefresh(new Date()); }}
            disabled={refreshing}
            title="Pulls fresh data from the sheet (clears the cache)"
            style={{ ...btnStyle, background: refreshing ? "#93c5fd" : "#2563eb", borderColor: "#2563eb",
              color: "#fff", cursor: refreshing ? "wait" : "pointer" }}>
            {refreshing ? "↻ Refreshing…" : "↻ Refresh"}
          </button>
          <button onClick={() => setAuto(a => !a)}
            style={{
              ...btnStyle,
              background: auto ? "#ecfdf5" : "#fff",
              borderColor: auto ? "#10b981" : "#d1d5db",
              color: auto ? "#065f46" : "#374151",
            }}>
            {auto ? "● Live" : "○ Paused"}
          </button>
          <button onClick={onLogout}
            style={{...btnStyle, background:"#fff", color:"#6b7280"}}>
            Sign out
          </button>
        </div>
      </div>

      {/* === FILTER BAR === */}
      <div style={{
        background:"#fff", border:"1px solid #e5e7eb", borderRadius:8,
        padding:"10px 12px", marginBottom:12,
        display:"grid", gridTemplateColumns:"repeat(5, 1fr)", gap:10,
      }}>
        <FilterField label="Year">
          <select value={year} onChange={e => setYear(Number(e.target.value))} style={selectStyle}>
            {Array.from(new Set([2026, new Date().getFullYear()]))
              .sort((a, b) => a - b)
              .map(y => <option key={y} value={y}>{y}</option>)}
          </select>
        </FilterField>
        <FilterField label="Month">
          <select value={month} onChange={e => setMonth(Number(e.target.value))} style={selectStyle}>
            {MONTH_NAMES.map((m, i) =>
              <option key={i} value={i+1}>{m}-{String(year).slice(2)}</option>)}
          </select>
        </FilterField>
        <FilterField label="FLM">
          <MultiSelect
            options={RAW.flms.filter(f => user.role !== "FLM" || f === user.flm).map(f => ({ value: f, label: f }))}
            selected={flmSel} onChange={setFlmSel}
            disabled={user.role === "FLM"} allLabel="All FLMs" />
        </FilterField>
        <FilterField label="SR">
          <MultiSelect
            options={srOptions.map(s => ({ value: s.code, label: s.name }))}
            selected={srSel} onChange={setSrSel} allLabel="All SRs" />
        </FilterField>
        <FilterField label="Customer">
          <MultiSelect
            options={filteredCustomers.map(c => ({ value: c.c, label: c.c + " — " + String(c.n || "").substring(0, 24) }))}
            selected={custSel} onChange={setCustSel} allLabel={"All (" + filteredCustomers.length + ")"} searchable />
        </FilterField>
      </div>

      {/* === SUMMARY METRICS — 6 cards === */}
      <div style={{display:"grid", gridTemplateColumns:"repeat(6, 1fr)", gap:8, marginBottom:12}}>
        <Metric label="Sales · Actual / Target" value={fmt(C.salesActualFull)}
          actual={fmt(C.totalTarget)} pct={salesPct} accent="#2563eb"
          secondaryLabel="tgt" />
        <Metric label="Active 3-mo · Actual / Target" value={fmt(activeTotal)}
          actual={fmt(activeTotalT)} pct={pct(activeTotal, activeTotalT)} accent="#8b5cf6"
          secondaryLabel="tgt" info={ACTIVE_3M_INFO} />
        <Metric label="Shop Around · Actual / Target" value={fmt(shopTotal)}
          actual={fmt(shopTotalT)} pct={pct(shopTotal, shopTotalT)} accent="#0ea5e9"
          secondaryLabel="tgt" />
        <Metric label="New Listing · Actual / Target" value={fmt(newTotal)}
          actual={fmt(newTotalT)} pct={pct(newTotal, newTotalT)} accent="#ec4899"
          secondaryLabel="tgt" />
        <Metric label="NU · Actual / Target" value={fmt(leadActualTotal)}
          actual={fmt(leadTotal)} pct={pct(leadActualTotal, leadTotal)} accent="#f59e0b"
          secondaryLabel="tgt" />
        <Metric label="Variance" value={fmt(C.salesActualFull - C.totalTarget)}
          actual={salesPct.toFixed(1) + "%"} pct={null}
          accent={C.salesActualFull >= C.totalTarget ? "#059669" : "#dc2626"}
          sub={C.salesActualFull >= C.totalTarget ? "Above target" : "Below target"} />
      </div>

      {/* === TABS === */}
      <div style={{
        display:"flex", gap:2, marginBottom:12, borderBottom:"1px solid #e5e7eb",
        flexWrap:"wrap", background:"#fff", borderRadius:"8px 8px 0 0",
        padding:"0 4px", border:"1px solid #e5e7eb",
      }}>
        <TabBtn label="Summary" v="summary" cur={tab} on={setTab} />
        <TabBtn label="SR Scorecards" v="scorecards" cur={tab} on={setTab} />
        <Sep />
        {KPI_DEFS.map(k =>
          <TabBtn key={k.key} label={k.label} v={"kpi:"+k.key} cur={tab} on={setTab} small />
        )}
        <Sep />
        <TabBtn label="Shop Around" v="shop" cur={tab} on={setTab} />
        <TabBtn label="Active Cust." v="active" cur={tab} on={setTab} />
        <TabBtn label="Active 1M" v="active1" cur={tab} on={setTab} />
        <TabBtn label="New Listing" v="new" cur={tab} on={setTab} />
        <TabBtn label="NU" v="leads" cur={tab} on={setTab} />
        <TabBtn label="L&L" v="ll" cur={tab} on={setTab} />
          <TabBtn label="📈 Trend" v="trend" cur={tab} on={setTab} />
          <TabBtn label="📅 Cust Sales" v="custsales" cur={tab} on={setTab} />
        {user.role === "Admin" && <>
          <Sep />
          <TabBtn label="🔑 Access Log" v="access" cur={tab} on={setTab} />
        </>}
      </div>

      {/* ============ SUMMARY ============ */}
      {tab === "summary" && (
        <>
          <Panel title="Sub-Brand Performance">
            <div style={{display:"grid", gridTemplateColumns:"repeat(7,1fr)", gap:8, marginBottom:10}}>
              {C.subBrandTotals.map(b => (
                <div key={b.name} style={{
                  background:"#fff", border:"1px solid #e5e7eb",
                  borderTop:"3px solid " + b.color, borderRadius:6, padding:"8px 10px",
                }}>
                  <div style={{fontSize:10, fontWeight:600, color:b.color}}>{b.name}</div>
                  <div style={{fontSize:15, fontWeight:700, marginTop:3}}>{fmt(b.actual)}</div>
                  <div style={{fontSize:9, color:"#6b7280"}}>of {fmt(b.target)}</div>
                  <div style={{fontSize:10, fontWeight:700, marginTop:3, color:pctColor(b.pct)}}>
                    {b.pct.toFixed(0)}%
                  </div>
                  <Bar2 pct={b.pct} />
                </div>
              ))}
            </div>
          </Panel>

          <Panel title="FLM × KPI Matrix — with MND / PND (click ▸ to expand SR detail)"
            action={<ExportBtn onClick={() => {
              // YTD by SR → Customer: one column per month Jan → latest + YTD total.
              const latest = latestDataMonth(RAW);
              const months = []; for (let m = 1; m <= latest; m++) months.push(m);
              const custName = {}; (RAW.customers || []).forEach(c => { custName[c.c] = c.n; });
              const srName = {}, srFlm = {}; RAW.srs.forEach(s => { srName[s.code] = s.name; srFlm[s.code] = s.flm; });
              const data = {};
              months.forEach(m => (RAW.kpiCustByMonth[m] || []).forEach(r => {
                const key = r.sr + "|" + r.c;
                (data[key] = data[key] || {})[m] = (data[key][m] || 0) + r.v;
              }));
              const list = Object.keys(data).map(key => {
                const parts = key.split("|"); const sr = Number(parts[0]), c = Number(parts[1]);
                const vals = data[key]; let ytd = 0; months.forEach(m => { ytd += (vals[m] || 0); });
                return { sr, c, flm: srFlm[sr] || "—", srNm: srName[sr] || ("SR " + sr), vals, ytd };
              })
                .filter(r => flmMatch(r.flm))
                .filter(r => allSr || srMatch(r.sr))
                .sort((a, b) => (a.sr - b.sr) || (b.ytd - a.ytd));
              const out = list.map(r => {
                const o = { "FLM": r.flm, "SR Code": r.sr, "SR Name": r.srNm, "Customer Code": r.c, "Customer Name": custName[r.c] || "" };
                months.forEach(m => { o[MONTH_NAMES[m - 1] + "-" + String(year).slice(2)] = Math.round(r.vals[m] || 0); });
                o["YTD Total"] = Math.round(r.ytd);
                return o;
              });
              exportToExcel(out, `Sales_byCustomer_YTD_Jan-${MONTH_NAMES[latest-1]}-${year}.xlsx`, "Sales by Customer");
            }} />}>
            <div style={{overflowX:"auto"}}>
              <table style={tblStyle}>
                <thead><tr style={{background:"#f9fafb"}}>
                  <th style={thStyle}>FLM / SR</th>
                  <th style={thStyleR}>Target</th>
                  <th style={thStyleR}>Actual</th>
                  <th style={thStyleR}>Var</th>
                  <th style={thStyleR}>%</th>
                  <th style={{...thStyleR, borderLeft:"1px solid #e5e7eb"}}>MND Tgt</th>
                  <th style={thStyleR}>MND Act</th>
                  <th style={thStyleR}>MND %</th>
                  <th style={{...thStyleR, borderLeft:"1px solid #e5e7eb"}}>PND Tgt</th>
                  <th style={thStyleR}>PND Act</th>
                  <th style={thStyleR}>PND %</th>
                  <th style={{...thStyle, borderLeft:"1px solid #e5e7eb"}}>Progress</th>
                </tr></thead>
                <tbody>
                  {C.flmRollup.map(f => {
                    const dv = C.flmDivision.find(d => d.flm === f.flm) || {mndT:0,mndA:0,pndT:0,pndA:0};
                    const mp = pct(dv.mndA, dv.mndT), pp = pct(dv.pndA, dv.pndT);
                    return (
                    <React.Fragment key={f.flm}>
                      <tr style={{borderTop:"1px solid #e5e7eb", background:"#fafbfc",
                        cursor:"pointer"}}
                        onClick={() => setExpanded(e => ({ ...e, [f.flm]: !e[f.flm] }))}>
                        <td style={{...tdStyle, fontWeight:600}}>
                          <span style={{display:"inline-block", width:14, color:"#6b7280"}}>
                            {expanded[f.flm] ? "▾" : "▸"}
                          </span>
                          {f.flm}
                          <span style={{fontSize:10, color:"#9ca3af", marginLeft:6, fontWeight:400}}>
                            {f.srs.length} SRs
                          </span>
                        </td>
                        <td style={tdStyleR}>{fmt(f.target)}</td>
                        <td style={tdStyleR}>{fmt(f.actual)}</td>
                        <td style={{...tdStyleR, color: f.actual >= f.target ? "#059669" : "#dc2626"}}>
                          {fmt(f.variance)}
                        </td>
                        <td style={{...tdStyleR, color:pctColor(f.pct), fontWeight:700}}>
                          {f.pct.toFixed(1)}%
                        </td>
                        <td style={{...tdStyleR, borderLeft:"1px solid #e5e7eb"}}>{fmt(dv.mndT)}</td>
                        <td style={tdStyleR}>{fmt(dv.mndA)}</td>
                        <td style={{...tdStyleR, color:pctColor(mp), fontWeight:600}}>{dv.mndT > 0 ? mp.toFixed(0)+"%" : "—"}</td>
                        <td style={{...tdStyleR, borderLeft:"1px solid #e5e7eb"}}>{fmt(dv.pndT)}</td>
                        <td style={tdStyleR}>{fmt(dv.pndA)}</td>
                        <td style={{...tdStyleR, color:pctColor(pp), fontWeight:600}}>{dv.pndT > 0 ? pp.toFixed(0)+"%" : "—"}</td>
                        <td style={{...tdStyle, borderLeft:"1px solid #e5e7eb"}}><Bar2 pct={f.pct} /></td>
                      </tr>
                      {expanded[f.flm] && f.srs.map(sr => {
                        const card = C.srScorecards.find(c => c.code === sr.code);
                        if (!card) return null;
                        const cmp = pct(card.mndA, card.mndT), cpp = pct(card.pndA, card.pndT);
                        return (
                          <tr key={sr.code} style={{borderTop:"1px solid #f3f4f6", background:"#fff"}}>
                            <td style={{...tdStyle, paddingLeft:32, fontSize:11}}>
                              <span style={{color:"#9ca3af", fontFamily:"monospace", fontSize:10, marginRight:6}}>
                                {sr.code}
                              </span>
                              {sr.name}
                            </td>
                            <td style={tdStyleR}>{fmt(card.totalT)}</td>
                            <td style={tdStyleR}>{fmt(card.totalA)}</td>
                            <td style={{...tdStyleR, color: card.totalA >= card.totalT ? "#059669" : "#dc2626"}}>
                              {fmt(card.totalA - card.totalT)}
                            </td>
                            <td style={{...tdStyleR, color:pctColor(card.totalPct), fontWeight:600}}>
                              {card.totalPct.toFixed(0)}%
                            </td>
                            <td style={{...tdStyleR, borderLeft:"1px solid #e5e7eb"}}>{fmt(card.mndT)}</td>
                            <td style={tdStyleR}>{fmt(card.mndA)}</td>
                            <td style={{...tdStyleR, color:pctColor(cmp), fontWeight:600}}>{card.mndT > 0 ? cmp.toFixed(0)+"%" : "—"}</td>
                            <td style={{...tdStyleR, borderLeft:"1px solid #e5e7eb"}}>{fmt(card.pndT)}</td>
                            <td style={tdStyleR}>{fmt(card.pndA)}</td>
                            <td style={{...tdStyleR, color:pctColor(cpp), fontWeight:600}}>{card.pndT > 0 ? cpp.toFixed(0)+"%" : "—"}</td>
                            <td style={{...tdStyle, borderLeft:"1px solid #e5e7eb"}}><Bar2 pct={card.totalPct} /></td>
                          </tr>
                        );
                      })}
                    </React.Fragment>
                    );
                  })}
                  <tr style={{borderTop:"2px solid #d1d5db", background:"#f9fafb", fontWeight:700}}>
                    <td style={tdStyle}>TOTAL</td>
                    <td style={tdStyleR}>{fmt(C.totalTarget)}</td>
                    <td style={tdStyleR}>{fmt(C.totalActual)}</td>
                    <td style={{...tdStyleR, color: C.totalActual >= C.totalTarget ? "#059669" : "#dc2626"}}>
                      {fmt(C.totalActual - C.totalTarget)}
                    </td>
                    <td style={{...tdStyleR, color:pctColor(overallPct)}}>
                      {overallPct.toFixed(1)}%
                    </td>
                    <td style={{...tdStyleR, borderLeft:"1px solid #e5e7eb"}}>{fmt(C.divisionTotals.mndT)}</td>
                    <td style={tdStyleR}>{fmt(C.divisionTotals.mndA)}</td>
                    <td style={{...tdStyleR, color:pctColor(pct(C.divisionTotals.mndA, C.divisionTotals.mndT))}}>
                      {C.divisionTotals.mndT > 0 ? pct(C.divisionTotals.mndA, C.divisionTotals.mndT).toFixed(0)+"%" : "—"}
                    </td>
                    <td style={{...tdStyleR, borderLeft:"1px solid #e5e7eb"}}>{fmt(C.divisionTotals.pndT)}</td>
                    <td style={tdStyleR}>{fmt(C.divisionTotals.pndA)}</td>
                    <td style={{...tdStyleR, color:pctColor(pct(C.divisionTotals.pndA, C.divisionTotals.pndT))}}>
                      {C.divisionTotals.pndT > 0 ? pct(C.divisionTotals.pndA, C.divisionTotals.pndT).toFixed(0)+"%" : "—"}
                    </td>
                    <td style={{...tdStyle, borderLeft:"1px solid #e5e7eb"}}><Bar2 pct={overallPct} /></td>
                  </tr>
                  {(() => {
                    // Reconcile the KPI-tracked total to the full Ethical Sales headline.
                    // The gap = Ethical sales on non-KPI products (or rows whose seller
                    // name didn't match an SR) — counted in Sales, no KPI column to sit in.
                    const other = Math.round(C.salesActualFull - C.totalActual);
                    if (Math.abs(other) < 1) return null;
                    return (
                      <>
                        <tr style={{background:"#fffdf5"}}>
                          <td style={{...tdStyle, color:"#92400e", fontStyle:"italic"}}>
                            ↳ Other Ethical (non-KPI products / unmatched rep)
                          </td>
                          <td style={tdStyleR}>—</td>
                          <td style={{...tdStyleR, color:"#92400e"}}>{fmt(other)}</td>
                          <td style={tdStyleR} colSpan={9}></td>
                        </tr>
                        <tr style={{background:"#f0fdf4", fontWeight:700}}>
                          <td style={{...tdStyle, color:"#166534"}}>= SALES (full Ethical, ties to headline)</td>
                          <td style={tdStyleR}>{fmt(C.totalTarget)}</td>
                          <td style={{...tdStyleR, color:"#166534"}}>{fmt(C.salesActualFull)}</td>
                          <td style={tdStyleR} colSpan={9}></td>
                        </tr>
                      </>
                    );
                  })()}
                </tbody>
              </table>
            </div>
            <div style={{fontSize:10, color:"#9ca3af", marginTop:6}}>
              MND = ENS PWD · ENS RPB · GLU PWD · GLU RPB · PRO &nbsp;·&nbsp; PND = SM · SIM · STC · PED PWD · PED RPB
              <br />TOTAL = the 10 tracked KPI sub-brands. "Other Ethical" = sales on non-KPI products or rows whose rep name didn't match — counted in the Sales headline but with no KPI column. TOTAL + Other = Sales.
            </div>
          </Panel>

          {(() => {
            // FLM × coverage KPIs (Shop Around, Active 1M, NU, L&L) — Target/Actual/%.
            // Rule: no target (Tgt = 0) → hide the Actual (and don't count it in the
            // FLM/total), since we don't track that KPI for that SR this month.
            const kc = (t, a) => {
              const p = pct(a, t);
              return (<React.Fragment>
                <td style={{...tdStyleR, borderLeft:"1px solid #e5e7eb"}}>{fmt(t)}</td>
                <td style={tdStyleR}>{t > 0 ? fmt(a) : "—"}</td>
                <td style={{...tdStyleR, color:pctColor(p), fontWeight:600}}>{t > 0 ? p.toFixed(0)+"%" : "—"}</td>
              </React.Fragment>);
            };
            const scMap = {}; C.srScorecards.forEach(c => { scMap[c.code] = c; });
            const srCov = (s) => {
              const c = scMap[s.code] || {};
              return {
                code: s.code, name: s.name,
                shopT: c.shopTarget || 0, shopA: c.shopActual || 0,
                a1T: active1TM.bySr[s.code] || 0, a1A: active1AM.bySr[s.code] || 0,
                nuT: c.leadTarget || 0, nuA: c.leadActual || 0,
                llT: llTM.bySr[s.code] || 0, llA: llAM.bySr[s.code] || 0,
              };
            };
            const KKEYS = [["shopT","shopA"],["a1T","a1A"],["nuT","nuA"],["llT","llA"]];
            const rows = (flmList).map(f => {
              const srs = RAW.srs
                .filter(s => s.flm === f && (srMatch(s.code)))
                .map(srCov);
              const row = { flm: f, srs: srs };
              // FLM target = sum of SR targets; FLM actual = sum of SR actuals that HAVE a target.
              KKEYS.forEach(([tk, ak]) => {
                let t = 0, a = 0;
                srs.forEach(s => { t += s[tk]; if (s[tk] > 0) a += s[ak]; });
                row[tk] = t; row[ak] = a;
              });
              return row;
            });
            const tot = rows.reduce((s, r) => ({
              shopT:s.shopT+r.shopT, shopA:s.shopA+r.shopA, a1T:s.a1T+r.a1T, a1A:s.a1A+r.a1A,
              nuT:s.nuT+r.nuT, nuA:s.nuA+r.nuA, llT:s.llT+r.llT, llA:s.llA+r.llA,
            }), {shopT:0,shopA:0,a1T:0,a1A:0,nuT:0,nuA:0,llT:0,llA:0});
            const grpTh = { ...thStyleR, borderLeft:"1px solid #e5e7eb", textAlign:"center" };
            const pctStr = (t, a) => t > 0 ? pct(a, t).toFixed(0) + "%" : "";
            const actNum = (t, a) => t > 0 ? Math.round(a) : "";
            const exportCov = () => {
              const out = [];
              rows.forEach(r => r.srs.forEach(s => {
                out.push({
                  "FLM": r.flm, "SR Code": s.code, "SR Name": s.name,
                  "Shop Around Tgt": Math.round(s.shopT), "Shop Around Act": actNum(s.shopT, s.shopA), "Shop Around %": pctStr(s.shopT, s.shopA),
                  "Active 1M Tgt": Math.round(s.a1T), "Active 1M Act": actNum(s.a1T, s.a1A), "Active 1M %": pctStr(s.a1T, s.a1A),
                  "NU Tgt": Math.round(s.nuT), "NU Act": actNum(s.nuT, s.nuA), "NU %": pctStr(s.nuT, s.nuA),
                  "L&L Tgt": Math.round(s.llT), "L&L Act": actNum(s.llT, s.llA), "L&L %": pctStr(s.llT, s.llA),
                });
              }));
              out.push({
                "FLM": "TOTAL", "SR Code": "", "SR Name": "",
                "Shop Around Tgt": Math.round(tot.shopT), "Shop Around Act": Math.round(tot.shopA), "Shop Around %": pctStr(tot.shopT, tot.shopA),
                "Active 1M Tgt": Math.round(tot.a1T), "Active 1M Act": Math.round(tot.a1A), "Active 1M %": pctStr(tot.a1T, tot.a1A),
                "NU Tgt": Math.round(tot.nuT), "NU Act": Math.round(tot.nuA), "NU %": pctStr(tot.nuT, tot.nuA),
                "L&L Tgt": Math.round(tot.llT), "L&L Act": Math.round(tot.llA), "L&L %": pctStr(tot.llT, tot.llA),
              });
              exportToExcel(out, `CoverageKPIs_${MONTH_NAMES[month-1]}-${String(year).slice(2)}.xlsx`, "Coverage KPIs");
            };
            return (
              <Panel title={"FLM × Coverage KPIs — Shop Around · Active 1M · NU · L&L (" + MONTH_NAMES[month-1] + "-" + String(year).slice(2) + ")"}
                action={<ExportBtn onClick={exportCov} />}>
                <div style={{overflowX:"auto"}}>
                  <table style={tblStyle}>
                    <thead>
                      <tr style={{background:"#f9fafb"}}>
                        <th style={thStyle} rowSpan={2}>FLM</th>
                        <th style={grpTh} colSpan={3}>Shop Around</th>
                        <th style={grpTh} colSpan={3}>Active 1M</th>
                        <th style={grpTh} colSpan={3}>NU</th>
                        <th style={grpTh} colSpan={3}>L&amp;L</th>
                      </tr>
                      <tr style={{background:"#f9fafb"}}>
                        {[0,1,2,3].map(i => (
                          <React.Fragment key={i}>
                            <th style={{...thStyleR, borderLeft:"1px solid #e5e7eb", fontSize:10}}>Tgt</th>
                            <th style={{...thStyleR, fontSize:10}}>Act</th>
                            <th style={{...thStyleR, fontSize:10}}>%</th>
                          </React.Fragment>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map(r => {
                        const ek = "cov_" + r.flm;
                        return (
                        <React.Fragment key={r.flm}>
                          <tr style={{borderTop:"1px solid #e5e7eb", background:"#fafbfc", cursor:"pointer", fontWeight:600}}
                            onClick={() => setExpanded(e => ({ ...e, [ek]: !e[ek] }))}>
                            <td style={tdStyle}>
                              <span style={{display:"inline-block", width:14, color:"#6b7280"}}>{expanded[ek] ? "▾" : "▸"}</span>
                              {r.flm}
                              <span style={{fontSize:10, color:"#9ca3af", marginLeft:6, fontWeight:400}}>{r.srs.length} SRs</span>
                            </td>
                            {kc(r.shopT, r.shopA)}
                            {kc(r.a1T, r.a1A)}
                            {kc(r.nuT, r.nuA)}
                            {kc(r.llT, r.llA)}
                          </tr>
                          {expanded[ek] && r.srs.map(s => (
                            <tr key={s.code} style={{borderTop:"1px solid #f3f4f6", background:"#fff"}}>
                              <td style={{...tdStyle, paddingLeft:32, fontSize:11}}>
                                <span style={{color:"#9ca3af", fontFamily:"monospace", fontSize:10, marginRight:6}}>{s.code}</span>
                                {s.name}
                              </td>
                              {kc(s.shopT, s.shopA)}
                              {kc(s.a1T, s.a1A)}
                              {kc(s.nuT, s.nuA)}
                              {kc(s.llT, s.llA)}
                            </tr>
                          ))}
                        </React.Fragment>
                        );
                      })}
                      <tr style={{borderTop:"2px solid #d1d5db", background:"#f9fafb", fontWeight:700}}>
                        <td style={tdStyle}>TOTAL</td>
                        {kc(tot.shopT, tot.shopA)}
                        {kc(tot.a1T, tot.a1A)}
                        {kc(tot.nuT, tot.nuA)}
                        {kc(tot.llT, tot.llA)}
                      </tr>
                    </tbody>
                  </table>
                </div>
              </Panel>
            );
          })()}

          <CoverageRanking
            title="FLM Coverage Rating (Shop Around · Active 3-mo · New Listing · NU)"
            rows={C.flmCoverage.map(r => ({
              key: r.flm, name: r.flm, sub: null,
              kpis: [
                { label: "Shop", actual: r.shopA, target: r.shopT },
                { label: "Active", actual: r.activeA, target: r.activeT },
                { label: "New", actual: r.newA, target: r.newT },
                { label: "NU", actual: r.leadA, target: r.leadT },
              ],
            }))}
            onExport={() => {
              const rows = C.flmCoverage.map(r => ({
                "FLM": r.flm,
                "Shop Target": Math.round(r.shopT),
                "Shop Actual": Math.round(r.shopA),
                "Shop %": r.shopT > 0 ? ((r.shopA/r.shopT)*100).toFixed(0) + "%" : "—",
                "Active Target": r.activeT,
                "Active Actual": r.activeA,
                "Active %": r.activeT > 0 ? ((r.activeA/r.activeT)*100).toFixed(0) + "%" : "—",
                "New Target": r.newT,
                "New Actual": r.newA,
                "New %": r.newT > 0 ? ((r.newA/r.newT)*100).toFixed(0) + "%" : "—",
                "NU Target": r.leadT,
                "NU Actual": r.leadA,
                "NU %": r.leadT > 0 ? ((r.leadA/r.leadT)*100).toFixed(0) + "%" : "—",
              }));
              exportToExcel(rows, `FLMCoverage_${MONTH_NAMES[month-1]}${year}.xlsx`, "FLM Coverage");
            }}
          />

          <CoverageRanking
            title="SR Coverage Rating (Shop Around · Active 3-mo · New Listing · NU)"
            rows={C.srScorecards.map(r => ({
              key: r.code, name: r.name, sub: r.flm,
              kpis: [
                { label: "Shop", actual: r.shopActual, target: r.shopTarget },
                { label: "Active", actual: r.activeActual, target: r.activeTarget },
                { label: "New", actual: r.newActual, target: r.newTarget },
                { label: "NU", actual: r.leadActual, target: r.leadTarget },
              ],
            }))}
            onExport={() => {
              const rows = C.srScorecards.map(r => ({
                "SR Code": r.code,
                "SR Name": r.name,
                "FLM": r.flm,
                "Shop Target": Math.round(r.shopTarget),
                "Shop Actual": Math.round(r.shopActual),
                "Shop %": r.shopTarget > 0 ? ((r.shopActual/r.shopTarget)*100).toFixed(0) + "%" : "—",
                "Active Target": r.activeTarget,
                "Active Actual": r.activeActual,
                "Active %": r.activeTarget > 0 ? ((r.activeActual/r.activeTarget)*100).toFixed(0) + "%" : "—",
                "New Target": r.newTarget,
                "New Actual": r.newActual,
                "New %": r.newTarget > 0 ? ((r.newActual/r.newTarget)*100).toFixed(0) + "%" : "—",
                "NU Target": r.leadTarget,
                "NU Actual": r.leadActual,
                "NU %": r.leadTarget > 0 ? ((r.leadActual/r.leadTarget)*100).toFixed(0) + "%" : "—",
              }));
              exportToExcel(rows, `SRCoverage_${MONTH_NAMES[month-1]}${year}.xlsx`, "SR Coverage");
            }}
          />
        </>
      )}

      {/* ============ SR SCORECARDS ============ */}
      {tab === "scorecards" && (() => {
        // Compute rankings
        const ranked = [...C.srScorecards]
          .filter(sr => sr.totalT > 0)
          .map(sr => {
            const kpiResults = KPI_DEFS.map(k => ({
              k: k.key, label: k.label, color: k.color, ...sr.kpis[k.key],
            })).filter(r => r.target > 0);
            const goldKPIs = kpiResults.filter(r => r.pct >= 100).length;
            const silverKPIs = kpiResults.filter(r => r.pct >= 90 && r.pct < 100).length;
            const bronzeKPIs = kpiResults.filter(r => r.pct >= 80 && r.pct < 90).length;
            const lowKPIs = kpiResults.filter(r => r.pct < 50).length;
            return { ...sr, kpiResults, goldKPIs, silverKPIs, bronzeKPIs, lowKPIs };
          })
          .sort((a, b) => b.totalPct - a.totalPct);

        ranked.forEach((sr, i) => {
          sr.rank = i + 1;
          if (sr.totalPct >= 100) sr.medal = "gold";
          else if (sr.totalPct >= 90) sr.medal = "silver";
          else if (sr.totalPct >= 80) sr.medal = "bronze";
          else sr.medal = null;
        });

        const goldCount = ranked.filter(s => s.medal === "gold").length;
        const silverCount = ranked.filter(s => s.medal === "silver").length;
        const bronzeCount = ranked.filter(s => s.medal === "bronze").length;
        const noMedalCount = ranked.filter(s => !s.medal).length;
        const top1 = ranked[0];
        const bot1 = [...ranked].reverse()[0];
        const teamAvg = ranked.reduce((s, r) => s + r.totalPct, 0) / Math.max(1, ranked.length);

        const medalEmoji = (m) => m === "gold" ? "🥇" : m === "silver" ? "🥈" : m === "bronze" ? "🥉" : "";

        // Cell renderer for KPI: shows actual / target on top, % below — all color-coded
        const KpiCell = ({ value, target, actual }) => {
          if (!target && !actual) return (
            <td style={{...tdStyleR, color:"#d1d5db", fontSize:9, padding:"4px 4px"}}>—</td>
          );
          const p = pct(actual, target);
          return (
            <td style={{
              ...tdStyleR, padding:"4px 4px",
              background: target > 0 ? heatBg(p) : "#f3f4f6",
              borderRight:"1px solid #fff",
              lineHeight: 1.15,
            }}>
              <div style={{fontSize:9, fontWeight:600, color: target > 0 ? heatFg(p) : "#6b7280"}}>
                {fmt(actual)}<span style={{opacity:.55, fontWeight:400}}>/{fmt(target)}</span>
              </div>
              <div style={{fontSize:10, fontWeight:700, color: target > 0 ? heatFg(p) : "#9ca3af"}}>
                {target > 0 ? p.toFixed(0) + "%" : "—"}
              </div>
            </td>
          );
        };

        return (
          <>
            {/* === COMPACT INSIGHT CARDS === */}
            <div style={{display:"grid", gridTemplateColumns:"repeat(4, 1fr)", gap:6, marginBottom:10}}>
              <div style={{background:"linear-gradient(135deg, #fef3c7 0%, #fde68a 100%)",
                border:"1px solid #f59e0b", borderRadius:6, padding:"6px 10px"}}>
                <div style={{fontSize:8, color:"#92400e", fontWeight:700, textTransform:"uppercase"}}>
                  🏆 #1 Overall
                </div>
                {top1 && (
                  <div style={{display:"flex", justifyContent:"space-between", alignItems:"baseline", marginTop:2}}>
                    <div>
                      <div style={{fontSize:11, fontWeight:700, color:"#78350f"}}>{top1.name}</div>
                      <div style={{fontSize:9, color:"#92400e"}}>{top1.flm}</div>
                    </div>
                    <div style={{fontSize:14, fontWeight:700, color:"#78350f"}}>{top1.totalPct.toFixed(0)}%</div>
                  </div>
                )}
              </div>

              <div style={{background:"#fff", border:"1px solid #e5e7eb", borderRadius:6, padding:"6px 10px"}}>
                <div style={{fontSize:8, color:"#6b7280", fontWeight:700, textTransform:"uppercase"}}>
                  📊 Medals
                </div>
                <div style={{display:"flex", justifyContent:"space-around", marginTop:4, gap:4}}>
                  <span style={{fontSize:11, fontWeight:700}}>🥇{goldCount}</span>
                  <span style={{fontSize:11, fontWeight:700, color:"#6b7280"}}>🥈{silverCount}</span>
                  <span style={{fontSize:11, fontWeight:700, color:"#a16207"}}>🥉{bronzeCount}</span>
                  <span style={{fontSize:11, fontWeight:700, color:"#dc2626"}}>⚠{noMedalCount}</span>
                </div>
              </div>

              <div style={{background:"linear-gradient(135deg, #fee2e2 0%, #fecaca 100%)",
                border:"1px solid #dc2626", borderRadius:6, padding:"6px 10px"}}>
                <div style={{fontSize:8, color:"#991b1b", fontWeight:700, textTransform:"uppercase"}}>
                  ⚠️ Needs Attention
                </div>
                {bot1 && (
                  <div style={{display:"flex", justifyContent:"space-between", alignItems:"baseline", marginTop:2}}>
                    <div>
                      <div style={{fontSize:11, fontWeight:700, color:"#7f1d1d"}}>{bot1.name}</div>
                      <div style={{fontSize:9, color:"#991b1b"}}>{bot1.flm}</div>
                    </div>
                    <div style={{fontSize:14, fontWeight:700, color:"#7f1d1d"}}>{bot1.totalPct.toFixed(0)}%</div>
                  </div>
                )}
              </div>

              <div style={{background:"#fff", border:"1px solid #e5e7eb", borderRadius:6, padding:"6px 10px"}}>
                <div style={{fontSize:8, color:"#6b7280", fontWeight:700, textTransform:"uppercase"}}>
                  📈 Team Avg
                </div>
                <div style={{display:"flex", justifyContent:"space-between", alignItems:"baseline", marginTop:2}}>
                  <div style={{fontSize:9, color:"#6b7280"}}>
                    {ranked.filter(r => r.totalPct >= 100).length} of {ranked.length} hit target
                  </div>
                  <div style={{fontSize:14, fontWeight:700, color:pctColor(teamAvg)}}>{teamAvg.toFixed(0)}%</div>
                </div>
              </div>
            </div>

            {/* === ONE BIG TABLE === */}
            <Panel title={"📊 SR Performance Table — All KPIs · " + MONTH_NAMES[month-1] + "-" + String(year).slice(2)}
              action={<ExportBtn onClick={() => {
                const rows = ranked.map(sr => {
                  const row = {
                    "Rank": sr.rank,
                    "SR Code": sr.code,
                    "SR Name": sr.name,
                    "FLM": sr.flm,
                    "Total Target": Math.round(sr.totalT),
                    "Total Actual": Math.round(sr.totalA),
                    "Total %": sr.totalT > 0 ? sr.totalPct.toFixed(1) + "%" : "—",
                  };
                  KPI_DEFS.forEach(k => {
                    const d = sr.kpis[k.key];
                    row[k.label + " Tgt"] = Math.round(d.target);
                    row[k.label + " Act"] = Math.round(d.actual);
                    row[k.label + " %"] = d.target > 0 ? ((d.actual/d.target)*100).toFixed(0) + "%" : "—";
                  });
                  row["Shop Tgt"] = Math.round(sr.shopTarget);
                  row["Shop Act"] = Math.round(sr.shopActual);
                  row["Active Tgt"] = sr.activeTarget;
                  row["Active Act"] = sr.activeActual;
                  row["New Tgt"] = sr.newTarget;
                  row["New Act"] = sr.newActual;
                  row["NU Tgt"] = sr.leadTarget;
                  row["NU Act"] = sr.leadActual;
                  return row;
                });
                exportToExcel(rows, `SRScorecards_${MONTH_NAMES[month-1]}${year}.xlsx`, "SR Scorecards");
              }} />}>
              <div style={{overflowX:"auto", marginLeft:-4, marginRight:-4}}>
                <table style={{...tblStyle, fontSize:10, borderCollapse:"separate", borderSpacing:0}}>
                  <thead>
                    {/* Group header row */}
                    <tr style={{background:"#f3f4f6"}}>
                      <th style={{...thStyle, position:"sticky", left:0, background:"#f3f4f6", zIndex:2,
                        borderRight:"2px solid #d1d5db"}} colSpan={2}>SR</th>
                      <th style={{...thStyle, textAlign:"center", borderRight:"2px solid #d1d5db",
                        background:"#fef3c7", color:"#92400e"}} colSpan={2}>TOTAL</th>
                      <th style={{...thStyle, textAlign:"center", borderRight:"2px solid #d1d5db",
                        background:"#dbeafe", color:"#1e40af"}} colSpan={10}>10 KPIs (Actual / Target / %)</th>
                      <th style={{...thStyle, textAlign:"center",
                        background:"#fce7f3", color:"#9d174d"}} colSpan={4}>COVERAGE</th>
                    </tr>
                    {/* Detail header row */}
                    <tr style={{background:"#fafafa"}}>
                      <th style={{...thStyle, position:"sticky", left:0, background:"#fafafa", zIndex:2,
                        width:55, borderBottom:"2px solid #d1d5db"}}>Rank</th>
                      <th style={{...thStyle, position:"sticky", left:55, background:"#fafafa", zIndex:2,
                        minWidth:130, borderRight:"2px solid #d1d5db", borderBottom:"2px solid #d1d5db"}}>Sales Rep</th>
                      <th style={{...thStyleR, fontSize:9, borderBottom:"2px solid #d1d5db"}}>Sales</th>
                      <th style={{...thStyleR, fontSize:9, borderRight:"2px solid #d1d5db", borderBottom:"2px solid #d1d5db"}}>%</th>
                      {KPI_DEFS.map(k => (
                        <th key={k.key} style={{...thStyleR, fontSize:8, color:k.color, padding:"6px 4px", borderBottom:"2px solid #d1d5db"}}>
                          {k.label.replace(" ", String.fromCharCode(8203))}
                        </th>
                      ))}
                      <th style={{...thStyleR, fontSize:9, color:"#0ea5e9", borderLeft:"2px solid #d1d5db", borderBottom:"2px solid #d1d5db", padding:"6px 6px"}}>SHOP<br/><span style={{fontSize:7, fontWeight:400, color:"#9ca3af"}}>act/tgt</span></th>
                      <th style={{...thStyleR, fontSize:9, color:"#8b5cf6", borderBottom:"2px solid #d1d5db", padding:"6px 6px"}}>ACT<br/><span style={{fontSize:7, fontWeight:400, color:"#9ca3af"}}>act/tgt</span></th>
                      <th style={{...thStyleR, fontSize:9, color:"#ec4899", borderBottom:"2px solid #d1d5db", padding:"6px 6px"}}>NEW<br/><span style={{fontSize:7, fontWeight:400, color:"#9ca3af"}}>act/tgt</span></th>
                      <th style={{...thStyleR, fontSize:9, color:"#f59e0b", borderBottom:"2px solid #d1d5db", padding:"6px 6px"}}>NU<br/><span style={{fontSize:7, fontWeight:400, color:"#9ca3af"}}>act/tgt</span></th>
                    </tr>
                  </thead>
                  <tbody>
                    {ranked.map((sr, idx) => {
                      const rowBg = sr.medal === "gold" ? "#fffbeb" : (idx % 2 === 0 ? "#fff" : "#fafafa");
                      const shopP = pct(sr.shopActual, sr.shopTarget);
                      const activeP = pct(sr.activeActual, sr.activeTarget);
                      const newP = pct(sr.newActual, sr.newTarget);
                      const leadP = pct(sr.leadActual, sr.leadTarget);
                      return (
                        <tr key={sr.code} style={{borderTop:"1px solid #f3f4f6"}}>
                          <td style={{
                            ...tdStyle, position:"sticky", left:0, background:rowBg, zIndex:1,
                            textAlign:"center", padding:"4px 6px",
                          }}>
                            <div style={{display:"flex", flexDirection:"column", alignItems:"center"}}>
                              <span style={{fontSize:14}}>{medalEmoji(sr.medal)}</span>
                              <span style={{fontSize:10, fontWeight:700, color:sr.rank <= 3 ? "#111827" : "#6b7280"}}>
                                #{sr.rank}
                              </span>
                            </div>
                          </td>
                          <td style={{
                            ...tdStyle, position:"sticky", left:55, background:rowBg, zIndex:1,
                            borderRight:"2px solid #d1d5db", padding:"6px 8px",
                          }}>
                            <div style={{fontSize:11, fontWeight:600}}>{sr.name}</div>
                            <div style={{fontSize:9, color:"#6b7280"}}>{sr.flm}</div>
                          </td>
                          {/* TOTAL */}
                          <td style={{...tdStyleR, fontSize:10, fontWeight:600}}>
                            {fmt(sr.totalA)}<span style={{color:"#9ca3af", fontSize:9}}>/{fmt(sr.totalT)}</span>
                          </td>
                          <td style={{
                            ...tdStyleR, fontSize:11, fontWeight:700, color:pctColor(sr.totalPct),
                            borderRight:"2px solid #d1d5db",
                          }}>
                            {sr.totalPct.toFixed(0)}%
                          </td>
                          {/* 10 KPIs */}
                          {KPI_DEFS.map(k => {
                            const d = sr.kpis[k.key];
                            return (
                              <KpiCell key={k.key} value={d.pct} target={d.target} actual={d.actual} />
                            );
                          })}
                          {/* COVERAGE — show actual/target/% */}
                          <td style={{
                            ...tdStyleR, padding:"4px 4px", borderLeft:"2px solid #d1d5db",
                            background: sr.shopTarget > 0 ? heatBg(shopP) : "#f3f4f6",
                            lineHeight:1.15,
                          }}>
                            <div style={{fontSize:9, fontWeight:600, color: sr.shopTarget > 0 ? heatFg(shopP) : "#6b7280"}}>
                              {fmt(sr.shopActual)}<span style={{opacity:.55, fontWeight:400}}>/{fmt(sr.shopTarget)}</span>
                            </div>
                            <div style={{fontSize:10, fontWeight:700, color: sr.shopTarget > 0 ? heatFg(shopP) : "#9ca3af"}}>
                              {sr.shopTarget > 0 ? shopP.toFixed(0) + "%" : "—"}
                            </div>
                          </td>
                          <td style={{
                            ...tdStyleR, padding:"4px 4px",
                            background: sr.activeTarget > 0 ? heatBg(activeP) : "#f3f4f6",
                            lineHeight:1.15,
                          }}>
                            <div style={{fontSize:9, fontWeight:600, color: sr.activeTarget > 0 ? heatFg(activeP) : "#6b7280"}}>
                              {sr.activeActual}<span style={{opacity:.55, fontWeight:400}}>/{sr.activeTarget}</span>
                            </div>
                            <div style={{fontSize:10, fontWeight:700, color: sr.activeTarget > 0 ? heatFg(activeP) : "#9ca3af"}}>
                              {sr.activeTarget > 0 ? activeP.toFixed(0) + "%" : "—"}
                            </div>
                          </td>
                          <td style={{
                            ...tdStyleR, padding:"4px 4px",
                            background: sr.newTarget > 0 ? heatBg(newP) : "#f3f4f6",
                            lineHeight:1.15,
                          }}>
                            <div style={{fontSize:9, fontWeight:600, color: sr.newTarget > 0 ? heatFg(newP) : "#6b7280"}}>
                              {sr.newActual}<span style={{opacity:.55, fontWeight:400}}>/{sr.newTarget}</span>
                            </div>
                            <div style={{fontSize:10, fontWeight:700, color: sr.newTarget > 0 ? heatFg(newP) : "#9ca3af"}}>
                              {sr.newTarget > 0 ? newP.toFixed(0) + "%" : "—"}
                            </div>
                          </td>
                          <td style={{
                            ...tdStyleR, padding:"4px 4px",
                            background: sr.leadTarget > 0 ? heatBg(leadP) : "#f3f4f6",
                            lineHeight:1.15,
                          }}>
                            <div style={{fontSize:9, fontWeight:600, color: sr.leadTarget > 0 ? heatFg(leadP) : "#6b7280"}}>
                              {sr.leadActual || 0}<span style={{opacity:.55, fontWeight:400}}>/{sr.leadTarget}</span>
                            </div>
                            <div style={{fontSize:10, fontWeight:700, color: sr.leadTarget > 0 ? heatFg(leadP) : "#9ca3af"}}>
                              {sr.leadTarget > 0 ? leadP.toFixed(0) + "%" : "—"}
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <div style={{fontSize:9, color:"#9ca3af", marginTop:8}}>
                Each cell: <strong>actual / target</strong> on top, <strong>%</strong> below. Color legend: 
                <span style={{background:"#a7f3d0", padding:"1px 5px", marginLeft:4, borderRadius:2}}>≥120%</span>
                <span style={{background:"#d1fae5", padding:"1px 5px", marginLeft:2, borderRadius:2}}>100-120%</span>
                <span style={{background:"#fef3c7", padding:"1px 5px", marginLeft:2, borderRadius:2}}>80-100%</span>
                <span style={{background:"#fee2e2", padding:"1px 5px", marginLeft:2, borderRadius:2}}>50-80%</span>
                <span style={{background:"#fecaca", padding:"1px 5px", marginLeft:2, borderRadius:2}}>{"<50%"}</span>
                <span style={{marginLeft:8}}>· Coverage = Shop / Active 3-mo / New Listing / NU</span>
              </div>
            </Panel>
          </>
        );
      })()}

      {/* ============ KPI TAB ============ */}
      {tab.startsWith("kpi:") && (() => {
        const kpiKey = tab.slice(4);
        const k = C.kpiTotals.find(x => x.key === kpiKey);
        if (!k) return null;

        // SR rows for this KPI
        const srRows = C.srScorecards
          .map(sr => ({ code: sr.code, name: sr.name, flm: sr.flm,
            ...sr.kpis[kpiKey] }))
          .filter(r => r.target > 0 || r.actual > 0)
          .sort((a, b) => b.target - a.target);

        // Customer details for this KPI in selected month
        const custDetails = (RAW.kpiCustByMonth[month] || [])
          .filter(r => r.k === kpiKey)
          .filter(r => allFlm || (() => {
            const sr = RAW.srs.find(s => s.code === r.sr);
            return sr && sflmMatch(r.flm);
          })())
          .filter(r => allSr || srMatch(r.sr))
          .filter(r => custMatch(r.c))
          .map(r => {
            const cust = RAW.customers.find(c => c.c === r.c);
            const sr = RAW.srs.find(s => s.code === r.sr);
            return {
              ...r,
              cn: cust ? cust.n : "Cust " + r.c,
              srName: sr ? sr.name : "—",
              flm: sr ? sr.flm : "—",
            };
          })
          .sort((a, b) => b.v - a.v);

        return (
          <>
            <div style={{display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:8, marginBottom:10}}>
              <Metric label={k.label + " · Actual / Target"} value={fmt(k.actual)}
                actual={fmt(k.target)} pct={k.pct} accent={k.color}
                secondaryLabel="tgt" />
              <Metric label="Variance" value={fmt(k.actual - k.target)}
                actual={k.actual >= k.target ? "Above" : "Below"} pct={null}
                accent={k.actual >= k.target ? "#059669" : "#dc2626"} />
              <Metric label="Active SRs" value={String(srRows.filter(r => r.actual > 0).length)}
                actual={"of " + srRows.length} pct={null} accent="#6366f1" />
              <Metric label="Customers buying" value={String(custDetails.length)}
                actual="in selected scope" pct={null} accent="#8b5cf6" />
            </div>

            {/* Top 3 / Bottom 3 for this KPI */}
            {(() => {
              const ranked = [...srRows].filter(r => r.target > 0).sort((a, b) => b.pct - a.pct);
              const top3 = ranked.slice(0, 3);
              const bot3 = [...ranked].reverse().slice(0, 3);
              if (ranked.length === 0) return null;
              return (
                <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:10}}>
                  <Panel title={"🏆 Top 3 — " + k.label}>
                    {top3.map((r, i) => {
                      const medal = r.pct >= 100 ? "🥇" : r.pct >= 90 ? "🥈" : r.pct >= 80 ? "🥉" : "";
                      return (
                        <div key={r.code} style={{
                          display:"flex", alignItems:"center", gap:8, padding:"6px 0",
                          borderBottom: i < 2 ? "1px solid #f3f4f6" : "none",
                        }}>
                          <div style={{fontSize:14, width:20}}>{medal || ("#" + (i+1))}</div>
                          <div style={{flex:1}}>
                            <div style={{fontSize:12, fontWeight:600, color:"#111827"}}>{r.name}</div>
                            <div style={{fontSize:10, color:"#6b7280"}}>{r.flm}</div>
                          </div>
                          <div style={{textAlign:"right"}}>
                            <div style={{fontSize:13, fontWeight:700, color:pctColor(r.pct)}}>
                              {r.pct.toFixed(0)}%
                            </div>
                            <div style={{fontSize:10, color:"#6b7280"}}>
                              {fmt(r.actual)} / {fmt(r.target)}
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </Panel>
                  <Panel title={"⚠️ Needs Help — " + k.label}>
                    {bot3.map((r, i) => (
                      <div key={r.code} style={{
                        display:"flex", alignItems:"center", gap:8, padding:"6px 0",
                        borderBottom: i < 2 ? "1px solid #f3f4f6" : "none",
                      }}>
                        <div style={{fontSize:14, width:20, color:"#dc2626"}}>⚠</div>
                        <div style={{flex:1}}>
                          <div style={{fontSize:12, fontWeight:600, color:"#111827"}}>{r.name}</div>
                          <div style={{fontSize:10, color:"#6b7280"}}>{r.flm}</div>
                        </div>
                        <div style={{textAlign:"right"}}>
                          <div style={{fontSize:13, fontWeight:700, color:pctColor(r.pct)}}>
                            {r.pct.toFixed(0)}%
                          </div>
                          <div style={{fontSize:10, color:"#6b7280"}}>
                            {fmt(r.actual)} / {fmt(r.target)}
                          </div>
                        </div>
                      </div>
                    ))}
                  </Panel>
                </div>
              );
            })()}

            <Panel title={k.label + " — FLM × SR Matrix (expandable)"}
              action={<ExportBtn onClick={() => {
                const rows = srRows.map(r => ({
                  "FLM": r.flm,
                  "SR Code": r.code,
                  "SR Name": r.name,
                  "Target": Math.round(r.target),
                  "Actual": Math.round(r.actual),
                  "Variance": Math.round(r.actual - r.target),
                  "% Achievement": r.target > 0 ? ((r.actual/r.target)*100).toFixed(1) + "%" : "—",
                }));
                exportToExcel(rows, `${k.label.replace(/\s/g, "")}_SRMatrix_${MONTH_NAMES[month-1]}${year}.xlsx`, k.label + " SR");
              }} />}>
              {(flmList).map(f => {
                const flmSrs = srRows.filter(r => r.flm === f);
                if (flmSrs.length === 0) return null;
                const flmT = flmSrs.reduce((s, r) => s + r.target, 0);
                const flmA = flmSrs.reduce((s, r) => s + r.actual, 0);
                return (
                  <div key={f} style={{marginBottom:12}}>
                    <div style={{
                      display:"flex", justifyContent:"space-between", alignItems:"center",
                      cursor:"pointer", padding:"6px 8px", background:"#fafbfc",
                      borderRadius:4, fontSize:12, fontWeight:600,
                    }} onClick={() => setExpanded(e => ({ ...e, ["k_"+f]: !e["k_"+f] }))}>
                      <span>
                        <span style={{color:"#6b7280", marginRight:6}}>
                          {expanded["k_"+f] !== false ? "▾" : "▸"}
                        </span>
                        {f}
                        <span style={{color:"#9ca3af", fontWeight:400, marginLeft:8, fontSize:11}}>
                          {flmSrs.length} SRs
                        </span>
                      </span>
                      <span style={{fontSize:11, color:"#6b7280", fontWeight:500}}>
                        Tgt {fmt(flmT)} · Act {fmt(flmA)} ·
                        <span style={{color:pctColor(pct(flmA, flmT)), marginLeft:4, fontWeight:600}}>
                          {pct(flmA, flmT).toFixed(0)}%
                        </span>
                      </span>
                    </div>
                    {expanded["k_"+f] !== false && (
                      <table style={tblStyle}>
                        <thead><tr style={{background:"#f9fafb"}}>
                          <th style={thStyle}>SR Code</th>
                          <th style={thStyle}>Sales Rep</th>
                          <th style={thStyleR}>Target</th>
                          <th style={thStyleR}>Actual</th>
                          <th style={thStyleR}>Var</th>
                          <th style={thStyleR}>%</th>
                          <th style={thStyle}>Progress</th>
                        </tr></thead>
                        <tbody>
                          {flmSrs.map(r => (
                            <tr key={r.code} style={{borderTop:"1px solid #f3f4f6"}}>
                              <td style={{...tdStyle, fontFamily:"monospace", color:"#9ca3af", fontSize:10}}>{r.code}</td>
                              <td style={tdStyle}>{r.name}</td>
                              <td style={tdStyleR}>{fmt(r.target)}</td>
                              <td style={tdStyleR}>{fmt(r.actual)}</td>
                              <td style={{...tdStyleR, color: r.actual >= r.target ? "#059669" : "#dc2626"}}>
                                {fmt(r.variance)}
                              </td>
                              <td style={{...tdStyleR, color:pctColor(r.pct), fontWeight:600}}>
                                {r.pct.toFixed(0)}%
                              </td>
                              <td style={tdStyle}><Bar2 pct={r.pct} /></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>
                );
              })}
            </Panel>

            <Panel title={k.label + " — Customer Details (" + custDetails.length + " customers)"}
              action={<ExportBtn onClick={() => {
                const rows = custDetails.map(c => ({
                  "FLM": c.flm,
                  "SR Name": c.srName,
                  "Customer Code": c.c,
                  "Customer Name": c.cn,
                  "Sales (Actual)": Math.round(c.v),
                }));
                exportToExcel(rows, `${k.label.replace(/\s/g, "")}_Customers_${MONTH_NAMES[month-1]}${year}.xlsx`, k.label);
              }} />}>
              <div style={{maxHeight:420, overflowY:"auto"}}>
                <table style={tblStyle}>
                  <thead style={{position:"sticky", top:0, background:"#f9fafb", zIndex:1}}>
                    <tr>
                      <th style={thStyle}>Customer</th>
                      <th style={thStyle}>SR</th>
                      <th style={thStyle}>FLM</th>
                      <th style={thStyleR}>Sales</th>
                    </tr>
                  </thead>
                  <tbody>
                    {custDetails.map((c, i) => (
                      <tr key={i} style={{borderTop:"1px solid #f3f4f6"}}>
                        <td style={tdStyle}>
                          {c.cn}
                          <span style={{color:"#9ca3af", fontSize:9, marginLeft:6, fontFamily:"monospace"}}>{c.c}</span>
                        </td>
                        <td style={{...tdStyle, fontSize:11}}>{c.srName}</td>
                        <td style={{...tdStyle, color:"#6b7280", fontSize:11}}>{c.flm}</td>
                        <td style={tdStyleR}>{fmt(c.v)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Panel>
          </>
        );
      })()}

      {/* ============ SHOP AROUND ============ */}
      {tab === "shop" && (() => {
        const items = C.shopItems;
        const flmList = allFlm ? (RAW.flms || []) : flmSel;

        // Per-FLM and per-SR rollups. Rule (same as the Summary Coverage matrix):
        // an SR's actual only counts when that SR has a Shop Around target — so a
        // no-target SR (or whole FLM) shows zero actual and is excluded from totals.
        const flmRows = flmList.map(f => {
          const fis = items.filter(x => x.f === f);
          const srCodes = [...new Set(fis.map(x => x.sr))];
          const target = fis.reduce((s, x) => s + x.t, 0);
          const actual = fis.reduce((s, x) => s + (x.t > 0 ? x.v : 0), 0);  // gate: only targeted customers count
          const actualRaw = fis.reduce((s, x) => s + x.v, 0);
          return { flm: f, srs: srCodes, target, actual, actualRaw, customers: fis };
        }).filter(r => r.target > 0 || r.actualRaw > 0);
        const totT = flmRows.reduce((s, r) => s + r.target, 0);
        const totA = flmRows.reduce((s, r) => s + r.actual, 0);

        return (
          <>
            <div style={{display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8, marginBottom:10}}>
              <Metric label="Shop Around · Actual / Target" value={fmt(totA)}
                actual={fmt(totT)} pct={pct(totA, totT)} accent="#0ea5e9"
                secondaryLabel="tgt" />
              <Metric label="Customers in scope" value={String(items.filter(x => !x.isNew).length)}
                actual={items.filter(x => x.isNew).length > 0 ? items.filter(x => x.isNew).length + " new prospects" : "specific accounts"}
                pct={null} accent="#8b5cf6" />
              <Metric label="Logic" value="Customer-level"
                actual="Sales to specific accounts" pct={null} accent="#ec4899" />
            </div>

            <Panel title="Shop Around — FLM Matrix (expand to see SRs and customers)"
              action={<ExportBtn onClick={() => ytdCustomerExport(RAW, year, flmSel, srSel,
                m => (RAW.shopByMonth[m] || []).filter(x => !x.isNew).map(x => ({ sr: x.sr, c: x.c, cn: x.cn, v: x.v })),
                "ShopAround_byCustomer", "Shop Around YTD")} />}>
              <table style={tblStyle}>
                <thead><tr style={{background:"#f9fafb"}}>
                  <th style={thStyle}>FLM / SR / Customer</th>
                  <th style={thStyleR}>Target</th>
                  <th style={thStyleR}>Actual</th>
                  <th style={thStyleR}>%</th>
                  <th style={thStyle}>Progress</th>
                </tr></thead>
                <tbody>
                  {flmRows.map(f => (
                    <React.Fragment key={f.flm}>
                      <tr style={{borderTop:"1px solid #e5e7eb", background:"#fafbfc",
                        cursor:"pointer", fontWeight:600}}
                        onClick={() => setExpanded(e => ({ ...e, ["s_"+f.flm]: !e["s_"+f.flm] }))}>
                        <td style={tdStyle}>
                          <span style={{color:"#6b7280", marginRight:6}}>
                            {expanded["s_"+f.flm] ? "▾" : "▸"}
                          </span>
                          {f.flm}
                          <span style={{fontSize:10, color:"#9ca3af", marginLeft:6, fontWeight:400}}>
                            {f.srs.length} SRs · {f.customers.filter(x => !x.isNew).length} customers
                            {f.customers.filter(x => x.isNew).length > 0 && (
                              <span style={{color:"#f59e0b", marginLeft:4}}>
                                + {f.customers.filter(x => x.isNew).length} new prospect{f.customers.filter(x => x.isNew).length>1?"s":""}
                              </span>
                            )}
                          </span>
                        </td>
                        <td style={tdStyleR}>{fmt(f.target)}</td>
                        <td style={tdStyleR}>{fmt(f.actual)}</td>
                        <td style={{...tdStyleR, color:pctColor(pct(f.actual, f.target)), fontWeight:700}}>
                          {f.target > 0 ? pct(f.actual, f.target).toFixed(1)+"%" : "—"}
                        </td>
                        <td style={tdStyle}><Bar2 pct={pct(f.actual, f.target)} /></td>
                      </tr>
                      {expanded["s_"+f.flm] && f.srs.map(srCode => {
                        const sr = RAW.srs.find(s => s.code === srCode);
                        const srItems = f.customers.filter(x => x.sr === srCode);
                        const srT = srItems.reduce((s, x) => s + x.t, 0);
                        const srA = srItems.reduce((s, x) => s + (x.t > 0 ? x.v : 0), 0);  // gate: only targeted customers count
                        return (
                          <React.Fragment key={srCode}>
                            <tr style={{borderTop:"1px solid #f3f4f6", cursor:"pointer"}}
                              onClick={() => setExpanded(e => ({ ...e, ["s_"+f.flm+"_"+srCode]: !e["s_"+f.flm+"_"+srCode] }))}>
                              <td style={{...tdStyle, paddingLeft:30, fontSize:11.5}}>
                                <span style={{color:"#9ca3af", marginRight:6}}>
                                  {expanded["s_"+f.flm+"_"+srCode] ? "▾" : "▸"}
                                </span>
                                <span style={{fontFamily:"monospace", color:"#9ca3af", fontSize:10, marginRight:6}}>{srCode}</span>
                                <strong>{sr?.name || srCode}</strong>
                                <span style={{fontSize:10, color:"#9ca3af", marginLeft:6}}>
                                  {srItems.filter(x => !x.isNew).length} customers
                                  {srItems.filter(x => x.isNew).length > 0 && (
                                    <span style={{color:"#f59e0b", marginLeft:4}}>
                                      + {srItems.filter(x => x.isNew).reduce((s, x) => s + x.t, 0)} new prospect tgt
                                    </span>
                                  )}
                                </span>
                              </td>
                              <td style={tdStyleR}>{fmt(srT)}</td>
                              <td style={tdStyleR}>{fmt(srA)}</td>
                              <td style={{...tdStyleR, color:pctColor(pct(srA, srT)), fontWeight:600}}>
                                {srT > 0 ? pct(srA, srT).toFixed(1)+"%" : "—"}
                              </td>
                              <td style={tdStyle}><Bar2 pct={pct(srA, srT)} /></td>
                            </tr>
                            {expanded["s_"+f.flm+"_"+srCode] && srItems
                              .filter(c => !c.isNew)
                              .sort((a, b) => b.t - a.t)
                              .map(c => (
                              <tr key={c.c} style={{borderTop:"1px solid #f3f4f6", background:"#fafbfc"}}>
                                <td style={{...tdStyle, paddingLeft:54, fontSize:11}}>
                                  {c.cn}
                                  <span style={{color:"#9ca3af", fontSize:9, marginLeft:6, fontFamily:"monospace"}}>{c.c}</span>
                                </td>
                                <td style={tdStyleR}>{fmt(c.t)}</td>
                                <td style={tdStyleR}>{fmt(c.t > 0 ? c.v : 0)}</td>
                                <td style={{...tdStyleR, color:pctColor(pct(c.v, c.t)), fontWeight:600}}>
                                  {c.t > 0 ? pct(c.v, c.t).toFixed(0) + "%" : "—"}
                                </td>
                                <td style={tdStyle}><Bar2 pct={c.t > 0 ? pct(c.v, c.t) : 0} /></td>
                              </tr>
                            ))}
                          </React.Fragment>
                        );
                      })}
                    </React.Fragment>
                  ))}
                  <tr style={{borderTop:"2px solid #d1d5db", background:"#f9fafb", fontWeight:700}}>
                    <td style={tdStyle}>TOTAL</td>
                    <td style={tdStyleR}>{fmt(totT)}</td>
                    <td style={tdStyleR}>{fmt(totA)}</td>
                    <td style={{...tdStyleR, color:pctColor(pct(totA, totT))}}>
                      {pct(totA, totT).toFixed(1)}%
                    </td>
                    <td style={tdStyle}><Bar2 pct={pct(totA, totT)} /></td>
                  </tr>
                </tbody>
              </table>
            </Panel>
          </>
        );
      })()}

      {/* ============ ACTIVE CUSTOMERS ============ */}
      {tab === "active" && (() => {
        const am = RAW.activeByMonth[month] || {};
        const at = RAW.activeTargetByMonth[month] || {};
        const flmList = allFlm ? (RAW.flms || []) : flmSel;
        const flmRows = flmList.map(f => ({
          flm: f, target: at.byFlm?.[f] || 0, actual: am.byFlm?.[f] || 0,
          srs: RAW.srs.filter(s => s.flm === f),
        }));
        const totT = flmRows.reduce((s, r) => s + r.target, 0);
        const totA = (allFlm && allSr && allCust) ? am.total
          : flmRows.reduce((s, r) => s + r.actual, 0);
        const unassigned = am.byFlm?.Unassigned || 0;

        return (
          <>
            <div style={{display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8, marginBottom:10}}>
              <Metric label="Active 3-mo · Actual / Target" value={fmt(totA)}
                actual={fmt(totT)} pct={pct(totA, totT)} accent="#8b5cf6"
                secondaryLabel="tgt" info={ACTIVE_3M_INFO} />
              <Metric label="Logic" value="3-mo rolling"
                actual={MONTH_NAMES[Math.max(0, month-3)] + " + " + MONTH_NAMES[Math.max(0, month-2)] + " + " + MONTH_NAMES[month-1]} pct={null} accent="#0ea5e9" />
              <Metric label="Unassigned" value={String(unassigned)}
                actual="customers without FLM" pct={null} accent="#6b7280" />
            </div>

            <Panel title="Active Customers — FLM × SR Matrix (expandable)"
              action={<ExportBtn onClick={() => {
                const rows = [];
                flmRows.forEach(f => {
                  f.srs.forEach(s => {
                    const tt = at.bySr?.[s.code] || 0;
                    const aa = am.bySr?.[s.code] || 0;
                    if (tt === 0 && aa === 0) return;
                    rows.push({
                      "FLM": f.flm,
                      "SR Code": s.code,
                      "SR Name": s.name,
                      "Target": tt,
                      "Actual": aa,
                      "% Achievement": tt > 0 ? ((aa/tt)*100).toFixed(0) + "%" : "—",
                    });
                  });
                });
                exportToExcel(rows, `ActiveCust_Matrix_${MONTH_NAMES[month-1]}${year}.xlsx`, "Active Matrix");
              }} />}>
              <table style={tblStyle}>
                <thead><tr style={{background:"#f9fafb"}}>
                  <th style={thStyle}>FLM / SR</th>
                  <th style={thStyleR}>Target</th>
                  <th style={thStyleR}>Actual</th>
                  <th style={thStyleR}>%</th>
                  <th style={thStyle}>Progress</th>
                </tr></thead>
                <tbody>
                  {flmRows.map(f => {
                    const p = pct(f.actual, f.target);
                    return (
                      <React.Fragment key={f.flm}>
                        <tr style={{borderTop:"1px solid #e5e7eb", background:"#fafbfc",
                          cursor:"pointer", fontWeight:600}}
                          onClick={() => setExpanded(e => ({ ...e, ["a_"+f.flm]: !e["a_"+f.flm] }))}>
                          <td style={tdStyle}>
                            <span style={{color:"#6b7280", marginRight:6}}>
                              {expanded["a_"+f.flm] ? "▾" : "▸"}
                            </span>
                            {f.flm}
                            <span style={{fontSize:10, color:"#9ca3af", marginLeft:6, fontWeight:400}}>
                              {f.srs.length} SRs
                            </span>
                          </td>
                          <td style={tdStyleR}>{fmt(f.target)}</td>
                          <td style={tdStyleR}>{fmt(f.actual)}</td>
                          <td style={{...tdStyleR, color:pctColor(p), fontWeight:700}}>
                            {f.target > 0 ? p.toFixed(0) + "%" : "—"}
                          </td>
                          <td style={tdStyle}><Bar2 pct={p} /></td>
                        </tr>
                        {expanded["a_"+f.flm] && f.srs.map(s => {
                          const tt = at.bySr?.[s.code] || 0;
                          const aa = am.bySr?.[s.code] || 0;
                          if (tt === 0 && aa === 0) return null;
                          const sp = pct(aa, tt);
                          return (
                            <tr key={s.code} style={{borderTop:"1px solid #f3f4f6"}}>
                              <td style={{...tdStyle, paddingLeft:32, fontSize:11.5}}>
                                <span style={{fontFamily:"monospace", color:"#9ca3af", fontSize:10, marginRight:6}}>{s.code}</span>
                                {s.name}
                              </td>
                              <td style={tdStyleR}>{fmt(tt)}</td>
                              <td style={tdStyleR}>{fmt(aa)}</td>
                              <td style={{...tdStyleR, color:pctColor(sp), fontWeight:600}}>
                                {tt > 0 ? sp.toFixed(0) + "%" : "—"}
                              </td>
                              <td style={tdStyle}><Bar2 pct={sp} /></td>
                            </tr>
                          );
                        })}
                      </React.Fragment>
                    );
                  })}
                  {unassigned > 0 && allFlm && (
                    <tr style={{borderTop:"1px solid #f3f4f6", color:"#9ca3af", fontStyle:"italic"}}>
                      <td style={tdStyle}>Unassigned (no FLM)</td>
                      <td style={tdStyleR}>—</td>
                      <td style={tdStyleR}>{fmt(unassigned)}</td>
                      <td style={tdStyleR}>—</td>
                      <td style={tdStyle}>—</td>
                    </tr>
                  )}
                  <tr style={{borderTop:"2px solid #d1d5db", background:"#f9fafb", fontWeight:700}}>
                    <td style={tdStyle}>TOTAL</td>
                    <td style={tdStyleR}>{fmt(totT)}</td>
                    <td style={tdStyleR}>{fmt(totA + (allFlm ? unassigned : 0))}</td>
                    <td style={{...tdStyleR, color:pctColor(pct(totA + (allFlm ? unassigned : 0), totT))}}>
                      {pct(totA + (allFlm ? unassigned : 0), totT).toFixed(0)}%
                    </td>
                    <td style={tdStyle}><Bar2 pct={pct(totA + (allFlm ? unassigned : 0), totT)} /></td>
                  </tr>
                </tbody>
              </table>
            </Panel>

            <Panel title={"Active Customer List (" + (am.customers || []).length + " customers active in 3-mo window)"}
              action={<ExportBtn onClick={() => {
                const SHARED = RAW.sharedCustomers || {};
                const custSr = {}; (RAW.customers || []).forEach(c => { custSr[c.c] = c.sr; });
                ytdCustomerExport(RAW, year, flmSel, srSel, m => {
                  const recs = [];
                  ((RAW.activeByMonth[m] || {}).customers || []).forEach(c => {
                    let srs = SHARED[c] ? SHARED[c].map(r => r.sr) : (custSr[c] ? [custSr[c]] : []);
                    uniqArr(srs).forEach(sr => recs.push({ sr, c, v: 1 }));
                  });
                  return recs;
                }, "ActiveCustomers_byCustomer", "Active YTD");
              }} />}>
              <div style={{maxHeight:380, overflowY:"auto"}}>
                <table style={tblStyle}>
                  <thead style={{position:"sticky", top:0, background:"#f9fafb", zIndex:1}}>
                    <tr>
                      <th style={thStyle}>Customer Code</th>
                      <th style={thStyle}>Customer Name</th>
                      <th style={thStyle}>SR</th>
                      <th style={thStyle}>FLM</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(am.customers || [])
                      .map(cc => {
                        const cust = RAW.customers.find(x => x.c === cc);
                        const sr = cust && cust.sr ? RAW.srs.find(s => s.code === cust.sr) : null;
                        return { c: cc, n: cust?.n, f: cust?.f, sr };
                      })
                      .filter(c => flmMatch(c.f))
                      .filter(c => srMatch(c.sr && c.sr.code))
                      .filter(c => custMatch(c.c))
                      .slice(0, 200)
                      .map(c => (
                        <tr key={c.c} style={{borderTop:"1px solid #f3f4f6"}}>
                          <td style={{...tdStyle, fontFamily:"monospace", color:"#6b7280", fontSize:10}}>{c.c}</td>
                          <td style={tdStyle}>{c.n || "—"}</td>
                          <td style={{...tdStyle, fontSize:11}}>{c.sr?.name || "—"}</td>
                          <td style={{...tdStyle, color:"#6b7280", fontSize:11}}>{c.f || "Unassigned"}</td>
                        </tr>
                      ))}
                  </tbody>
                </table>
              </div>
            </Panel>

            {(() => {
              if (!RAW.customerMonthly) return (
                <RedeployNotice title="Customer Purchase Trend by SR" feature="The customer purchase trend" />
              );
              const maxP = RAW.customerMonthlyMaxPeriod || (year * 100 + month);
              const periods = [];
              { let y = 2025, mm = 1; while (y * 100 + mm <= maxP) { periods.push(y * 100 + mm); mm++; if (mm > 12) { mm = 1; y++; } } }
              const pLabel = (p) => MONTH_NAMES[(p % 100) - 1] + "-" + String(Math.floor(p / 100)).slice(2);
              // Group the SAME active customers the matrix counts, using the SAME
              // attribution (shared customers → each of their SRs), so the per-SR
              // counts here exactly match the Active matrix's Actual.
              const cmByCode = {};
              (RAW.customerMonthly || []).forEach(cu => { cmByCode[cu.c] = cu; });
              const custByCode = {};
              (RAW.customers || []).forEach(c => { custByCode[c.c] = c; });
              const SHARED = RAW.sharedCustomers || {};
              const bySR = {};
              (am.customers || []).forEach(c => {
                // Use the SAME attribution the backend used for the Active matrix
                // (who actually sold + shared + master), so the tabs stay consistent.
                // Fall back to the old client-side logic only pre-redeploy.
                let srs = (am.custSrs && am.custSrs[c]) ? am.custSrs[c] : null;
                if (!srs) {
                  srs = [];
                  if (SHARED[c]) srs = SHARED[c].map(r => r.sr);
                  else if (custByCode[c] && custByCode[c].sr) srs = [custByCode[c].sr];
                  srs = srs.filter((v, i) => srs.indexOf(v) === i); // once per SR
                }
                srs.forEach(sr => {
                  const cu = cmByCode[c] || { c: c, n: (custByCode[c] && custByCode[c].n) || ("Customer " + c), p: {} };
                  (bySR[sr] = bySR[sr] || []).push(cu);
                });
              });
              const srList = RAW.srs.filter(s =>
                (flmMatch(s.flm)) && (srMatch(s.code)));
              const srWithCust = srList.filter(s => (bySR[s.code] || []).length);

              const stHead = { position: "sticky", top: 0, zIndex: 2, background: "#f9fafb" };
              const stName = { position: "sticky", left: 0, zIndex: 1, background: "#fff", minWidth: 150, maxWidth: 150 };
              const stCorner = { ...stHead, left: 0, zIndex: 3, minWidth: 150, maxWidth: 150 };

              const exportTrend = () => {
                const rows = [];
                srWithCust.forEach(s => (bySR[s.code] || []).forEach(cu => {
                  const o = { "SR": s.name, "Manager": s.flm, "Customer Code": cu.c, "Customer Name": cu.n };
                  periods.forEach(p => { o[pLabel(p)] = (cu.p[p] || 0) > 0 ? 1 : 0; });
                  rows.push(o);
                }));
                exportToExcel(rows, `CustomerTrend_bySR_to_${pLabel(maxP)}.xlsx`, "Customer Trend");
              };

              return (
                <Panel title="Customer Purchase Trend by SR (1 = bought · 0 = none/return) — click an SR to expand"
                  action={<ExportBtn onClick={exportTrend} />}>
                  <div style={{overflow:"auto", maxHeight:"70vh", border:"1px solid #f3f4f6", borderRadius:6}}>
                    <table style={{...tblStyle, fontSize:11}}>
                      <thead>
                        <tr>
                          <th style={{...stCorner, ...thStyle, textAlign:"left"}}>SR / Customer</th>
                          {periods.map(p => <th key={p} style={{...stHead, ...thStyleR, whiteSpace:"nowrap"}}>{pLabel(p)}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {srWithCust.length === 0 && (
                          <tr><td style={tdStyle} colSpan={periods.length + 1}>No SR customers in the current filters.</td></tr>
                        )}
                        {srWithCust.map(s => {
                          const custs = bySR[s.code] || [];
                          const open = expanded["trend_" + s.code];
                          return (
                            <React.Fragment key={s.code}>
                              <tr style={{borderTop:"1px solid #e5e7eb", background:"#fafbfc", cursor:"pointer", fontWeight:600}}
                                onClick={() => setExpanded(e => ({ ...e, ["trend_" + s.code]: !e["trend_" + s.code] }))}>
                                <td style={{...stCorner, ...tdStyle, background:"#fafbfc"}}>
                                  <span style={{color:"#6b7280", marginRight:6}}>{open ? "▾" : "▸"}</span>
                                  {s.name}
                                  <span style={{fontSize:10, color:"#9ca3af", marginLeft:6, fontWeight:400}}>{custs.length} active · {s.flm}</span>
                                </td>
                                {periods.map(p => {
                                  const cnt = custs.reduce((a, cu) => a + ((cu.p[p] || 0) > 0 ? 1 : 0), 0);
                                  return <td key={p} style={{...tdStyleR, color:"#6b7280"}}>{cnt}</td>;
                                })}
                              </tr>
                              {open && custs.map(cu => (
                                <tr key={cu.c} style={{borderTop:"1px solid #f3f4f6"}}>
                                  <td style={{...stName, ...tdStyle, paddingLeft:22}}>
                                    <div style={{whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis", maxWidth:128}}>{cu.n}</div>
                                    <div style={{fontSize:9, color:"#9ca3af", fontFamily:"monospace"}}>{cu.c}</div>
                                  </td>
                                  {periods.map(p => {
                                    const one = (cu.p[p] || 0) > 0;
                                    return <td key={p} style={{...tdStyleR, fontWeight:700,
                                      background: one ? "#dcfce7" : "#fef0c7", color: one ? "#15803d" : "#b45309"}}>{one ? 1 : 0}</td>;
                                  })}
                                </tr>
                              ))}
                            </React.Fragment>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                  <div style={{fontSize:10, color:"#9ca3af", marginTop:6}}>
                    Shows the SAME active customers as the Active matrix (3-mo window, shared customers counted under each of their SRs) — so the "active" count per SR matches the matrix Actual. Expand an SR for each customer's 1/0 monthly trend (1 = purchased, 0 = no purchase or return).
                  </div>
                </Panel>
              );
            })()}
          </>
        );
      })()}

      {/* ============ NEW LISTING ============ */}
      {tab === "new" && (() => {
        const nm = RAW.newByMonth[month] || {};
        const nt = RAW.newTargetByMonth[month] || {};
        const flmList = allFlm ? (RAW.flms || []) : flmSel;
        const flmRows = flmList.map(f => ({
          flm: f, target: nt.byFlm?.[f] || 0, actual: nm.byFlm?.[f] || 0,
          srs: RAW.srs.filter(s => s.flm === f),
        }));
        const totT = flmRows.reduce((s, r) => s + r.target, 0);
        const totA = flmRows.reduce((s, r) => s + r.actual, 0);

        const items = (nm.items || [])
          .filter(p => flmMatch(p.f))
          .filter(p => srMatch(p.sr))
          .filter(p => custMatch(p.c));

        return (
          <>
            <div style={{display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8, marginBottom:10}}>
              <Metric label="New Listing · Actual / Target" value={fmt(totA)}
                actual={fmt(totT)} pct={pct(totA, totT)} accent="#ec4899"
                secondaryLabel="tgt" />
              <Metric label="New listings found" value={String(items.length)}
                actual="customer × KPI pairs" pct={null} accent="#8b5cf6" />
              <Metric label="Logic" value="12-mo rolling"
                actual="not bought in prior 12 months" pct={null} accent="#f59e0b" />
            </div>

            <Panel title="New Listing — FLM × SR Matrix (expandable)"
              action={<ExportBtn onClick={() => {
                const rows = [];
                flmRows.forEach(f => {
                  f.srs.forEach(s => {
                    const tt = nt.bySr?.[s.code] || 0;
                    const aa = nm.bySr?.[s.code] || 0;
                    if (tt === 0 && aa === 0) return;
                    rows.push({
                      "FLM": f.flm,
                      "SR Code": s.code,
                      "SR Name": s.name,
                      "Target": tt,
                      "Actual": aa,
                      "% Achievement": tt > 0 ? ((aa/tt)*100).toFixed(0) + "%" : "—",
                    });
                  });
                });
                exportToExcel(rows, `NewListing_Matrix_${MONTH_NAMES[month-1]}${year}.xlsx`, "New Matrix");
              }} />}>
              <table style={tblStyle}>
                <thead><tr style={{background:"#f9fafb"}}>
                  <th style={thStyle}>FLM / SR</th>
                  <th style={thStyleR}>Target</th>
                  <th style={thStyleR}>Actual</th>
                  <th style={thStyleR}>%</th>
                  <th style={thStyle}>Progress</th>
                </tr></thead>
                <tbody>
                  {flmRows.map(f => {
                    const p = pct(f.actual, f.target);
                    return (
                      <React.Fragment key={f.flm}>
                        <tr style={{borderTop:"1px solid #e5e7eb", background:"#fafbfc",
                          cursor:"pointer", fontWeight:600}}
                          onClick={() => setExpanded(e => ({ ...e, ["n_"+f.flm]: !e["n_"+f.flm] }))}>
                          <td style={tdStyle}>
                            <span style={{color:"#6b7280", marginRight:6}}>
                              {expanded["n_"+f.flm] ? "▾" : "▸"}
                            </span>
                            {f.flm}
                          </td>
                          <td style={tdStyleR}>{fmt(f.target)}</td>
                          <td style={tdStyleR}>{fmt(f.actual)}</td>
                          <td style={{...tdStyleR, color:pctColor(p), fontWeight:700}}>
                            {f.target > 0 ? p.toFixed(0) + "%" : "—"}
                          </td>
                          <td style={tdStyle}><Bar2 pct={p} /></td>
                        </tr>
                        {expanded["n_"+f.flm] && f.srs.map(s => {
                          const tt = nt.bySr?.[s.code] || 0;
                          const aa = nm.bySr?.[s.code] || 0;
                          if (tt === 0 && aa === 0) return null;
                          const sp = pct(aa, tt);
                          return (
                            <tr key={s.code} style={{borderTop:"1px solid #f3f4f6"}}>
                              <td style={{...tdStyle, paddingLeft:32, fontSize:11.5}}>
                                <span style={{fontFamily:"monospace", color:"#9ca3af", fontSize:10, marginRight:6}}>{s.code}</span>
                                {s.name}
                              </td>
                              <td style={tdStyleR}>{fmt(tt)}</td>
                              <td style={tdStyleR}>{fmt(aa)}</td>
                              <td style={{...tdStyleR, color:pctColor(sp), fontWeight:600}}>
                                {tt > 0 ? sp.toFixed(0) + "%" : "—"}
                              </td>
                              <td style={tdStyle}><Bar2 pct={sp} /></td>
                            </tr>
                          );
                        })}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </Panel>

            <Panel title={"New Listing items (" + items.length + ")"}
              action={<ExportBtn onClick={() => ytdCustomerExport(RAW, year, flmSel, srSel, m => {
                const its = (RAW.newByMonth[m] || {}).items || [];
                const recs = [];
                its.forEach(p => {
                  const srs = uniqArr(p.srs && p.srs.length ? p.srs : (p.sr ? [p.sr] : []));
                  srs.forEach(sr => recs.push({ sr, c: p.c, cn: p.n, v: 1 }));
                });
                return recs;
              }, "NewListing_byCustomer", "New Listing YTD")} />}>
              <div style={{maxHeight:380, overflowY:"auto"}}>
                <table style={tblStyle}>
                  <thead style={{position:"sticky", top:0, background:"#f9fafb", zIndex:1}}>
                    <tr>
                      <th style={thStyle}>Customer</th>
                      <th style={thStyle}>SR</th>
                      <th style={thStyle}>FLM</th>
                      <th style={thStyle}>KPI</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((p, i) => {
                      const sr = p.sr ? RAW.srs.find(s => s.code === p.sr) : null;
                      return (
                        <tr key={i} style={{borderTop:"1px solid #f3f4f6"}}>
                          <td style={tdStyle}>
                            {p.n || "Cust " + p.c}
                            <span style={{color:"#9ca3af", fontSize:9, marginLeft:6, fontFamily:"monospace"}}>{p.c}</span>
                          </td>
                          <td style={{...tdStyle, fontSize:11}}>{sr?.name || "—"}</td>
                          <td style={{...tdStyle, color:"#6b7280", fontSize:11}}>{p.f || "Unassigned"}</td>
                          <td style={tdStyle}>
                            <span style={{
                              background:"#eef2ff", padding:"2px 6px", borderRadius:3,
                              fontSize:10, color:"#4338ca", fontWeight:600,
                            }}>{p.k}</span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </Panel>
          </>
        );
      })()}

      {/* ============ LEADS ============ */}
      {tab === "leads" && (() => {
        const leadTM = (RAW.leadTargetByMonth && RAW.leadTargetByMonth[month]) || { bySr: {}, byFlm: {} };
        const leadAM = (RAW.leadActualByMonth && RAW.leadActualByMonth[month]) || { bySr: {}, byFlm: {} };
        return (
          <>
            <div style={{display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8, marginBottom:10}}>
              <Metric label="NU · Actual / Target" value={fmt(leadActualTotal)}
                actual={fmt(leadTotal)} pct={pct(leadActualTotal, leadTotal)} accent="#f59e0b"
                secondaryLabel="tgt" />
              <Metric label="SRs with target" value={String(Object.keys(leadTM.bySr || {}).filter(k => leadTM.bySr[k] > 0).length)}
                actual={(RAW.leads || []).length + " entries"} pct={null} accent="#0ea5e9" />
              <Metric label="Source" value="Actual-NU tab"
                actual="auto-extends per month" pct={null} accent="#6b7280" />
            </div>
            <Panel title="NU — FLM × SR (expandable)"
              action={<ExportBtn onClick={() => {
                const latest = latestDataMonth(RAW);
                const months = []; for (let m = 1; m <= latest; m++) months.push(m);
                const srs = RAW.srs.filter(s => (flmMatch(s.flm)) && (srMatch(s.code)));
                const out = srs.map(s => {
                  const o = { "FLM": s.flm, "SR Code": s.code, "SR Name": s.name };
                  let ytdA = 0, ytdT = 0;
                  months.forEach(m => {
                    const am = (RAW.leadActualByMonth && RAW.leadActualByMonth[m]) ? (RAW.leadActualByMonth[m].bySr[s.code] || 0) : 0;
                    const tm = (RAW.leadTargetByMonth && RAW.leadTargetByMonth[m]) ? (RAW.leadTargetByMonth[m].bySr[s.code] || 0) : 0;
                    o[MONTH_NAMES[m - 1] + "-" + String(year).slice(2)] = Math.round(am);
                    ytdA += am; ytdT += tm;
                  });
                  o["YTD Actual"] = Math.round(ytdA); o["YTD Target"] = Math.round(ytdT);
                  o["YTD %"] = ytdT > 0 ? ((ytdA / ytdT) * 100).toFixed(0) + "%" : "—";
                  return o;
                }).filter(o => o["YTD Actual"] !== 0 || o["YTD Target"] !== 0);
                exportToExcel(out, `NU_YTD_Jan-${MONTH_NAMES[latest-1]}-${year}.xlsx`, "NU YTD");
              }} />}>
              <table style={tblStyle}>
                <thead><tr style={{background:"#f9fafb"}}>
                  <th style={thStyle}>FLM / SR</th>
                  <th style={thStyleR}>Target</th>
                  <th style={thStyleR}>Actual</th>
                  <th style={thStyleR}>%</th>
                  <th style={thStyle}>Progress</th>
                </tr></thead>
                <tbody>
                  {(flmList).map(f => {
                    const tt = leadTM.byFlm[f] || 0;
                    const ta = leadAM.byFlm[f] || 0;
                    const tp = pct(ta, tt);
                    const flmSrs = RAW.srs.filter(s => s.flm === f);
                    return (
                      <React.Fragment key={f}>
                        <tr style={{borderTop:"1px solid #e5e7eb", background:"#fafbfc",
                          cursor:"pointer", fontWeight:600}}
                          onClick={() => setExpanded(e => ({ ...e, ["l_"+f]: !e["l_"+f] }))}>
                          <td style={tdStyle}>
                            <span style={{color:"#6b7280", marginRight:6}}>
                              {expanded["l_"+f] ? "▾" : "▸"}
                            </span>
                            {f}
                            <span style={{fontSize:10, color:"#9ca3af", marginLeft:6, fontWeight:400}}>
                              {flmSrs.length} SRs
                            </span>
                          </td>
                          <td style={tdStyleR}>{tt}</td>
                          <td style={tdStyleR}>{ta}</td>
                          <td style={{...tdStyleR, color:pctColor(tp), fontWeight:700}}>
                            {tt > 0 ? tp.toFixed(0) + "%" : "—"}
                          </td>
                          <td style={tdStyle}><Bar2 pct={tp} /></td>
                        </tr>
                        {expanded["l_"+f] && flmSrs.map(s => {
                          const t = leadTM.bySr[s.code] || 0;
                          const a = leadAM.bySr[s.code] || 0;
                          if (t === 0 && a === 0) return null;
                          const sp = pct(a, t);
                          return (
                            <tr key={s.code} style={{borderTop:"1px solid #f3f4f6"}}>
                              <td style={{...tdStyle, paddingLeft:32, fontSize:11.5}}>
                                <span style={{fontFamily:"monospace", color:"#9ca3af", fontSize:10, marginRight:6}}>{s.code}</span>
                                {s.name}
                              </td>
                              <td style={tdStyleR}>{t}</td>
                              <td style={tdStyleR}>{a}</td>
                              <td style={{...tdStyleR, color:pctColor(sp), fontWeight:600}}>
                                {t > 0 ? sp.toFixed(0) + "%" : "—"}
                              </td>
                              <td style={tdStyle}><Bar2 pct={sp} /></td>
                            </tr>
                          );
                        })}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </Panel>
          </>
        );
      })()}

      {tab === "active1" && (() => {
        const buildYtd = (actualByMonth, targetByMonth, fname, sheet) => {
          const latest = latestDataMonth(RAW);
          const months = []; for (let m = 1; m <= latest; m++) months.push(m);
          const srs = RAW.srs.filter(s => (flmMatch(s.flm)) && (srMatch(s.code)));
          const out = srs.map(s => {
            const o = { "FLM": s.flm, "SR Code": s.code, "SR Name": s.name };
            let ytdA = 0, ytdT = 0;
            months.forEach(m => {
              const am = (RAW[actualByMonth] && RAW[actualByMonth][m]) ? (RAW[actualByMonth][m].bySr[s.code] || 0) : 0;
              const tm = (RAW[targetByMonth] && RAW[targetByMonth][m]) ? (RAW[targetByMonth][m].bySr[s.code] || 0) : 0;
              o[MONTH_NAMES[m-1] + "-" + String(year).slice(2)] = Math.round(am);
              ytdA += am; ytdT += tm;
            });
            o["YTD Actual"] = Math.round(ytdA); o["YTD Target"] = Math.round(ytdT);
            return o;
          }).filter(o => o["YTD Actual"] !== 0 || o["YTD Target"] !== 0);
          exportToExcel(out, fname + "_Jan-" + MONTH_NAMES[latest-1] + "-" + year + ".xlsx", sheet);
        };
        return (
          <>
            <div style={{display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8, marginBottom:10}}>
              <Metric label="Active 1-mo · Actual / Target" value={fmt(active1Total)}
                actual={fmt(active1TotalT)} pct={pct(active1Total, active1TotalT)} accent="#7c3aed"
                secondaryLabel="tgt" />
              <Metric label="Window" value="1 month"
                actual="Net-positive buyers this month" pct={null} accent="#0ea5e9" />
              <Metric label="Logic" value="Same as 3-mo"
                actual="1-month rolling window" pct={null} accent="#6b7280" />
            </div>
            {(() => {
              // Customer lookup — type a code to see exactly how this month treats it:
              // active?, credited to which SR(s)?, assigned in Shared_Customers?, unassigned?
              const code = Number(String(custLookup).trim());
              const srByCode = {}; (RAW.srs || []).forEach(s => { srByCode[s.code] = s; });
              const nm = (c) => (srByCode[c] && srByCode[c].name) || ("SR " + c);
              let report = null;
              if (code) {
                const custObj = (RAW.customers || []).find(c => c.c === code)
                  || (RAW.customerMonthly || []).find(c => c.c === code);
                const activeCodes = active1AM.customers || [];
                const isActive = activeCodes.indexOf(code) >= 0;
                const creditedTo = (active1AM.custSrs && active1AM.custSrs[code]) || [];
                const asg = (RAW.custAssign && RAW.custAssign[code]) || null;
                const un = (active1AM.unassigned || []).find(u => u.c === code) || null;
                const sharedRows = (RAW.sharedCustomers && RAW.sharedCustomers[code]) || [];
                report = { custObj, isActive, creditedTo, asg, un, sharedRows };
              }
              return (
                <Panel title="🔎 Customer lookup — how is this outlet counted this month?">
                  <div style={{display:"flex", gap:8, alignItems:"center", marginBottom:8}}>
                    <input value={custLookup} onChange={e => setCustLookup(e.target.value)}
                      placeholder="Enter Customer Code (e.g. 281033726)"
                      style={{...selectStyle, minWidth:260, padding:"5px 9px"}} />
                    <span style={{fontSize:11, color:"#6b7280"}}>Showing for {MONTH_NAMES[month-1]}-{String(year).slice(2)}</span>
                  </div>
                  {!code ? (
                    <div style={{fontSize:12, color:"#9ca3af"}}>Type a customer code to see whether it's active, who it's credited to, and its Shared_Customers assignment.</div>
                  ) : !report.custObj && !report.isActive ? (
                    <div style={{fontSize:12, color:"#b45309"}}>No customer found for code {code}. Check the code.</div>
                  ) : (
                    <div style={{fontSize:12.5, lineHeight:1.7}}>
                      <div><b>{(report.custObj && (report.custObj.n)) || ("Customer " + code)}</b> <span style={{fontFamily:"monospace", color:"#9ca3af"}}>{code}</span>
                        {report.custObj && report.custObj.f ? <span style={{color:"#6b7280"}}> · {report.custObj.f}</span> : null}</div>
                      <div>Active this month: {report.isActive
                        ? <b style={{color:"#16a34a"}}>YES</b>
                        : <b style={{color:"#dc2626"}}>NO</b>}
                        {report.un ? <span style={{color:"#6b7280"}}> · PND {fmt(report.un.pnd)} · MND {fmt(report.un.mnd)}</span> : null}</div>
                      <div>Credited to: {report.creditedTo.length
                        ? <b style={{color:"#7c3aed"}}>{report.creditedTo.map(nm).join(", ")}</b>
                        : <span style={{color:"#b45309"}}>nobody (Unassigned bucket)</span>}</div>
                      <div>In Shared_Customers: {report.sharedRows.length
                        ? <span style={{color:"#16a34a"}}>{report.sharedRows.map(r => nm(r.sr) + " (" + (r.cat || "?") + ")").join(", ")}</span>
                        : <span style={{color:"#b45309"}}>NO — falls back to the booking rep</span>}</div>
                      {report.un ? (
                        <div style={{marginTop:6, padding:"6px 8px", background:"#fef3c7", borderRadius:4, color:"#92400e"}}>
                          This outlet is in the <b>Unassigned</b> list (has PND/MND sales, no Shared_Customers row).
                          <div style={{marginTop:4}}>
                            PND booked by: {report.un.pnd > 0 ? ((report.un.pndBy || []).length ? <b>{(report.un.pndBy).map(nm).join(", ")}</b> : <b style={{color:"#dc2626"}}>⚠ seller name not recognized</b>) : "—"}
                            {" · "}MND booked by: {report.un.mnd > 0 ? ((report.un.mndBy || []).length ? <b>{(report.un.mndBy).map(nm).join(", ")}</b> : <b style={{color:"#dc2626"}}>⚠ seller name not recognized</b>) : "—"}
                          </div>
                        </div>
                      ) : null}
                    </div>
                  )}
                </Panel>
              );
            })()}
            <Panel title="Active 1-Month — FLM × SR (expandable)"
              action={<ExportBtn onClick={() => buildYtd("active1ByMonth", "activeTarget1ByMonth", "Active1M_YTD", "Active 1M YTD")} />}>
              <KpiMatrix TM={active1TM} AM={active1AM} flmList={flmList} RAW={RAW}
                expanded={expanded} setExpanded={setExpanded} keyPrefix="a1" />
            </Panel>
            {(() => {
              // Which outlets make up each SR's number — click an SR to see the names.
              const custSrs = active1AM.custSrs || {};
              const custByCode = {}; (RAW.customers || []).forEach(c => { custByCode[c.c] = c; });
              const bySR = {};
              Object.keys(custSrs).forEach(cc => (custSrs[cc] || []).forEach(sr => { (bySR[sr] = bySR[sr] || []).push(Number(cc)); }));
              const srs = RAW.srs
                .filter(s => (flmMatch(s.flm)) && (srMatch(s.code)))
                .filter(s => (bySR[s.code] || []).length);
              return (
                <Panel title={"Active 1-Month — outlets per SR (" + MONTH_NAMES[month-1] + "-" + String(year).slice(2) + ") — click an SR"}>
                  <table style={tblStyle}>
                    <thead><tr style={{background:"#f9fafb"}}>
                      <th style={thStyle}>SR / Outlet</th><th style={thStyleR}>Active outlets</th>
                    </tr></thead>
                    <tbody>
                      {srs.map(s => {
                        const list = (bySR[s.code] || []).slice().sort((a,b) => a-b);
                        const ek = "a1c_" + s.code;
                        return (
                          <React.Fragment key={s.code}>
                            <tr style={{borderTop:"1px solid #e5e7eb", background:"#fafbfc", cursor:"pointer", fontWeight:600}}
                              onClick={() => setExpanded(e => ({ ...e, [ek]: !e[ek] }))}>
                              <td style={tdStyle}>
                                <span style={{color:"#6b7280", marginRight:6}}>{expanded[ek] ? "▾" : "▸"}</span>
                                <span style={{fontFamily:"monospace", color:"#9ca3af", fontSize:10, marginRight:6}}>{s.code}</span>
                                {s.name}
                                <span style={{fontSize:10, color:"#9ca3af", marginLeft:6, fontWeight:400}}>{s.flm}</span>
                              </td>
                              <td style={tdStyleR}>{list.length}</td>
                            </tr>
                            {expanded[ek] && list.map(cc => (
                              <tr key={cc} style={{borderTop:"1px solid #f3f4f6"}}>
                                <td style={{...tdStyle, paddingLeft:32, fontSize:11.5}}>
                                  {(custByCode[cc] && custByCode[cc].n) || ("Customer " + cc)}
                                  <span style={{fontFamily:"monospace", color:"#9ca3af", fontSize:10, marginLeft:6}}>{cc}</span>
                                </td>
                                <td style={tdStyleR}></td>
                              </tr>
                            ))}
                          </React.Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                </Panel>
              );
            })()}
            {(() => {
              // Outlets that are active (had PND/MND sales this month) but are NOT
              // in Shared_Customers, so no rep is formally assigned. Shows the
              // FLM's PND/MND reps (from the Team column) so you know who to add.
              const rawList = active1AM.unassigned || [];
              const srByCode = {}; (RAW.srs || []).forEach(s => { srByCode[s.code] = s; });
              const flmCatReps = RAW.flmCatReps || {};
              let list = rawList.filter(r => allFlm || flmMatch(r.f));
              list = list.slice().sort((a, b) => (b.pnd + b.mnd) - (a.pnd + a.mnd));
              const repNames = (arr) => (arr || []).map(x => x.name).join(", ") || "—";
              return (
                <Panel title={"⚠️ Unassigned active outlets (" + list.length + ") — had PND/MND sales but not in Shared_Customers"}>
                  <div style={{fontSize:11, color:"#6b7280", marginBottom:8}}>
                    These outlets are active but no rep is formally assigned, so they currently fall back to whoever booked the sale.
                    To credit the right rep, add a row in <b>Shared_Customers</b> (Customer Code · SR Code · Category · Weight · SR Name · Customer Name)
                    using the FLM's PND/MND rep shown here.
                  </div>
                  {list.length === 0 ? (
                    <div style={{fontSize:12, color:"#16a34a", padding:"6px 2px"}}>✓ Every active outlet in scope is assigned in Shared_Customers.</div>
                  ) : (
                  <div style={{overflowX:"auto"}}>
                  <table style={tblStyle}>
                    <thead><tr style={{background:"#f9fafb"}}>
                      <th style={thStyle}>Outlet</th>
                      <th style={thStyle}>FLM</th>
                      <th style={thStyleR}>PND</th>
                      <th style={thStyleR}>MND</th>
                      <th style={thStyle}>Counts for now</th>
                      <th style={thStyle}>Assign to (PND rep)</th>
                      <th style={thStyle}>Assign to (MND rep)</th>
                    </tr></thead>
                    <tbody>
                      {list.map(r => {
                        const cat = flmCatReps[r.f] || { PND: [], MND: [] };
                        const now = r.sr ? ((srByCode[r.sr] && srByCode[r.sr].name) || ("SR " + r.sr)) : "—";
                        return (
                          <tr key={r.c} style={{borderTop:"1px solid #f3f4f6"}}>
                            <td style={tdStyle}>
                              {r.n}
                              <span style={{fontFamily:"monospace", color:"#9ca3af", fontSize:10, marginLeft:6}}>{r.c}</span>
                            </td>
                            <td style={{...tdStyle, fontSize:11, color:"#6b7280"}}>{r.f || "—"}</td>
                            <td style={{...tdStyleR, color: r.pnd > 0 ? "#111827" : "#d1d5db"}}>{r.pnd > 0 ? fmt(r.pnd) : "—"}</td>
                            <td style={{...tdStyleR, color: r.mnd > 0 ? "#111827" : "#d1d5db"}}>{r.mnd > 0 ? fmt(r.mnd) : "—"}</td>
                            <td style={{...tdStyle, fontSize:11, color:"#6b7280"}}>{now}</td>
                            <td style={{...tdStyle, fontSize:11, color: r.pnd > 0 ? "#7c3aed" : "#d1d5db"}}>{r.pnd > 0 ? repNames(cat.PND) : "—"}</td>
                            <td style={{...tdStyle, fontSize:11, color: r.mnd > 0 ? "#0d9488" : "#d1d5db"}}>{r.mnd > 0 ? repNames(cat.MND) : "—"}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  </div>
                  )}
                </Panel>
              );
            })()}
          </>
        );
      })()}

      {tab === "ll" && (() => {
        const buildYtd = () => {
          const latest = latestDataMonth(RAW);
          const months = []; for (let m = 1; m <= latest; m++) months.push(m);
          const srs = RAW.srs.filter(s => (flmMatch(s.flm)) && (srMatch(s.code)));
          const out = srs.map(s => {
            const o = { "FLM": s.flm, "SR Code": s.code, "SR Name": s.name };
            let ytdA = 0, ytdT = 0;
            months.forEach(m => {
              const am = (RAW.llActualByMonth && RAW.llActualByMonth[m]) ? (RAW.llActualByMonth[m].bySr[s.code] || 0) : 0;
              const tm = (RAW.llTargetByMonth && RAW.llTargetByMonth[m]) ? (RAW.llTargetByMonth[m].bySr[s.code] || 0) : 0;
              o[MONTH_NAMES[m-1] + "-" + String(year).slice(2)] = Math.round(am);
              ytdA += am; ytdT += tm;
            });
            o["YTD Actual"] = Math.round(ytdA); o["YTD Target"] = Math.round(ytdT);
            return o;
          }).filter(o => o["YTD Actual"] !== 0 || o["YTD Target"] !== 0);
          exportToExcel(out, `LL_YTD_Jan-${MONTH_NAMES[latest-1]}-${year}.xlsx`, "L&L YTD");
        };
        return (
          <>
            <div style={{display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8, marginBottom:10}}>
              <Metric label="L&L · Actual / Target" value={fmt(llActualTotal)}
                actual={fmt(llTotalT)} pct={pct(llActualTotal, llTotalT)} accent="#0d9488"
                secondaryLabel="tgt" />
              <Metric label="SRs with target" value={String(Object.keys(llTM.bySr || {}).filter(k => llTM.bySr[k] > 0).length)}
                actual="from Target-L&L" pct={null} accent="#0ea5e9" />
              <Metric label="Source" value="Actual-L&L tab"
                actual="enter monthly actuals" pct={null} accent="#6b7280" />
            </div>
            <Panel title="L&L — FLM × SR (expandable)"
              action={<ExportBtn onClick={buildYtd} />}>
              <KpiMatrix TM={llTM} AM={llAM} flmList={flmList} RAW={RAW}
                expanded={expanded} setExpanded={setExpanded} keyPrefix="ll" />
            </Panel>
          </>
        );
      })()}

      {tab === "trend" && (() => {
        try {
          const TREND_KPIS = [
            { id: "sales",  label: "Total Sales",      color: "#2563eb" },
            { id: "active", label: "Active 3-mo Cust", color: "#8b5cf6" },
            { id: "shop",   label: "Shop Around",      color: "#0ea5e9" },
            { id: "new",    label: "New Listing",      color: "#ec4899" },
            { id: "nu",     label: "NU",               color: "#f59e0b" },
          ];

          const flmList = allFlm ? (RAW.flms || []) : flmSel;

          const computeMonth = (kpiId, m) => {
            let actual = 0, target = 0;
            try {
              if (kpiId === "sales") {
                (RAW.actuals || []).forEach(r => {
                  if (r.m !== m) return;
                  if (!flmMatch(r.f)) return;
                  actual += r.v || 0;
                });
                (RAW.targets || []).forEach(r => {
                  if (r.m !== m) return;
                  if (!flmMatch(r.f)) return;
                  target += r.t || 0;
                });
              } else if (kpiId === "active") {
                // activeByMonth[m] = {total, byFlm, bySr, customers}
                const ab = (RAW.activeByMonth && RAW.activeByMonth[m]) || {};
                if (!allFlm) {
                  actual = flmList.reduce((s2, f2) => s2 + ((ab.byFlm && ab.byFlm[f2]) || 0), 0);
                } else {
                  actual = ab.total || 0;
                }
                const at = RAW.activeTargetByMonth && RAW.activeTargetByMonth[m];
                if (at) {
                  if (!allFlm) {
                    target = flmList.reduce((s2, f2) => s2 + ((at.byFlm && at.byFlm[f2]) || 0), 0);
                  } else {
                    Object.values(at.bySr || {}).forEach(v => target += v);
                  }
                }
              } else if (kpiId === "shop") {
                const sb = (RAW.shopByMonth && RAW.shopByMonth[m]) || [];
                sb.forEach(r => {
                  if (!flmMatch(r.f)) return;
                  actual += r.v || 0;
                  target += r.t || 0;
                });
              } else if (kpiId === "new") {
                // newByMonth[m] = {items, byFlm, bySr}
                const nb = (RAW.newByMonth && RAW.newByMonth[m]) || {};
                if (!allFlm) {
                  actual = flmList.reduce((s2, f2) => s2 + ((nb.byFlm && nb.byFlm[f2]) || 0), 0);
                } else {
                  actual = (nb.items && nb.items.length) || 0;
                }
                const nt = RAW.newTargetByMonth && RAW.newTargetByMonth[m];
                if (nt) {
                  if (!allFlm) {
                    target = flmList.reduce((s2, f2) => s2 + ((nt.byFlm && nt.byFlm[f2]) || 0), 0);
                  } else {
                    Object.values(nt.bySr || {}).forEach(v => target += v);
                  }
                }
              } else if (kpiId === "nu") {
                const lt = RAW.leadTargetByMonth && RAW.leadTargetByMonth[m];
                const la = RAW.leadActualByMonth && RAW.leadActualByMonth[m];
                if (lt) {
                  if (!allFlm) {
                    target = flmList.reduce((s2, f2) => s2 + ((lt.byFlm && lt.byFlm[f2]) || 0), 0);
                  } else {
                    Object.values(lt.bySr || {}).forEach(v => target += v);
                  }
                }
                if (la) {
                  if (!allFlm) {
                    actual = flmList.reduce((s2, f2) => s2 + ((la.byFlm && la.byFlm[f2]) || 0), 0);
                  } else {
                    Object.values(la.bySr || {}).forEach(v => actual += v);
                  }
                }
              }
            } catch (e) {
              // swallow per-month errors
            }
            return { actual: Math.round(actual), target: Math.round(target) };
          };

          // Build chart data per KPI
          const buildKpiData = (kpiId) => {
            return MONTH_NAMES.map((mLabel, idx) => {
              const m = idx + 1;
              const { actual, target } = computeMonth(kpiId, m);
              return { month: mLabel, Actual: actual, Target: target };
            });
          };

          const fmtVal = (v) => {
            if (!v) return "0";
            return Math.round(v).toLocaleString();
          };

          const trendDirection = (data) => {
            const filtered = data.filter(d => d.Actual > 0);
            if (filtered.length < 2) return null;
            const first = filtered[0].Actual;
            const last = filtered[filtered.length - 1].Actual;
            if (first === 0) return last > 0 ? "↗ improving" : "→ stable";
            const change = (last - first) / first;
            if (change > 0.05) return "↗ improving";
            if (change < -0.05) return "↘ declining";
            return "→ stable";
          };

          return (
            <>
            <Panel title="📈 Trend Analysis — 2026 Year-to-Date"
              action={<ExportBtn onClick={() => {
                const rows = [];
                TREND_KPIS.forEach(kpi => {
                  MONTH_NAMES.forEach((mLabel, idx) => {
                    const { actual, target } = computeMonth(kpi.id, idx + 1);
                    if (actual === 0 && target === 0) return;
                    rows.push({
                      "KPI": kpi.label,
                      "Month": mLabel + " " + year,
                      "Actual": actual,
                      "Target": target,
                      "%": target > 0 ? ((actual/target)*100).toFixed(0) + "%" : "—",
                    });
                  });
                });
                exportToExcel(rows, `Trend_${year}_${allFlm ? "All" : flm.replace(/\s/g,"")}.xlsx`, "Trend");
              }} />}>
              <div style={{fontSize:11, color:"#6b7280", marginBottom:14}}>
                5 KPIs across 2026. FLM filter applies (currently: <strong>{allFlm ? "All" : flmSel.join(", ")}</strong>).
                Months with no data show 0. Trend = direction of last few months with data.
              </div>

              {TREND_KPIS.map(kpi => {
                const data = buildKpiData(kpi.id);
                const hasData = data.some(d => d.Actual > 0 || d.Target > 0);
                if (!hasData) return null;

                const trend = trendDirection(data);
                const trendColor = trend && trend.includes("improving") ? "#059669"
                  : trend && trend.includes("declining") ? "#dc2626" : "#6b7280";

                // Totals
                const totalA = data.reduce((s, d) => s + d.Actual, 0);
                const totalT = data.reduce((s, d) => s + d.Target, 0);

                return (
                  <div key={kpi.id} style={{
                    marginBottom: 18, padding: 14,
                    border: "1px solid #e5e7eb", borderRadius: 8, background: "#fff"
                  }}>
                    {/* Header */}
                    <div style={{display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:10}}>
                      <div>
                        <div style={{fontSize:14, fontWeight:700, color:kpi.color}}>{kpi.label}</div>
                        <div style={{fontSize:10, color:"#9ca3af", marginTop:2}}>
                          YTD: <strong>{fmtVal(totalA)}</strong> actual / <strong>{fmtVal(totalT)}</strong> target
                          {totalT > 0 && (
                            <span style={{marginLeft:6, color: pctColor(pct(totalA, totalT))}}>
                              · {pct(totalA, totalT).toFixed(0)}%
                            </span>
                          )}
                        </div>
                      </div>
                      {trend && (
                        <div style={{
                          fontSize:11, fontWeight:700, color:trendColor,
                          padding:"4px 10px", background:"#f9fafb", borderRadius:6,
                          border:"1px solid " + trendColor + "33",
                        }}>
                          {trend}
                        </div>
                      )}
                    </div>

                    {/* Bar chart with target + actual */}
                    <div style={{height:240}}>
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={data} margin={{top:25, right:15, left:0, bottom:5}}
                          barCategoryGap="20%" barGap={3}>
                          <CartesianGrid strokeDasharray="3 3" stroke="#f3f4f6" />
                          <XAxis dataKey="month" tick={{fontSize:10, fill:"#6b7280"}} />
                          <YAxis tick={{fontSize:10, fill:"#6b7280"}} tickFormatter={fmtVal} />
                          <Tooltip formatter={(v) => fmtVal(v)} />
                          <Legend wrapperStyle={{fontSize:11}} />
                          <Bar dataKey="Target" fill="#d1d5db" name="Target" maxBarSize={28}>
                            <LabelList dataKey="Target" position="top" formatter={fmtVal}
                              style={{fontSize:9, fill:"#9ca3af"}} />
                          </Bar>
                          <Bar dataKey="Actual" fill={kpi.color} name="Actual" maxBarSize={28}>
                            <LabelList dataKey="Actual" position="top" formatter={fmtVal}
                              style={{fontSize:9, fill:"#374151", fontWeight:600}} />
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </div>

                    {/* Monthly summary table */}
                    <div style={{marginTop:14}}>
                      <div style={{fontSize:11, fontWeight:600, color:"#374151", marginBottom:6}}>
                        Monthly Detail
                      </div>
                      <table style={{...tblStyle, fontSize:10}}>
                        <thead>
                          <tr style={{background:"#f9fafb"}}>
                            <th style={thStyle}>Month</th>
                            <th style={thStyleR}>Actual</th>
                            <th style={thStyleR}>Target</th>
                            <th style={thStyleR}>%</th>
                            <th style={thStyleR}>vs Prior</th>
                          </tr>
                        </thead>
                        <tbody>
                          {data.map((d, i) => {
                            if (d.Actual === 0 && d.Target === 0) return null;
                            const p = pct(d.Actual, d.Target);
                            const prior = i > 0 ? data[i-1].Actual : 0;
                            const change = prior > 0 ? ((d.Actual - prior) / prior * 100) : null;
                            return (
                              <tr key={d.month} style={{borderTop:"1px solid #f3f4f6"}}>
                                <td style={tdStyle}>{d.month} 2026</td>
                                <td style={{...tdStyleR, fontWeight:600}}>{fmtVal(d.Actual)}</td>
                                <td style={{...tdStyleR, color:"#6b7280"}}>
                                  {d.Target > 0 ? fmtVal(d.Target) : "—"}
                                </td>
                                <td style={{...tdStyleR, fontWeight:700,
                                  color: d.Target > 0 ? pctColor(p) : "#9ca3af"}}>
                                  {d.Target > 0 ? p.toFixed(0) + "%" : "—"}
                                </td>
                                <td style={{...tdStyleR,
                                  color: change > 0 ? "#059669" : change < 0 ? "#dc2626" : "#6b7280"}}>
                                  {change !== null ? (change > 0 ? "+" : "") + change.toFixed(0) + "%" : "—"}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                );
              })}
            </Panel>

            {/* ============= HEATMAP SECTIONS ============= */}
            {(() => {
              const HEATMAP_KPIS = [
                { id: "sales",  label: "Total Sales",      color: "#2563eb" },
                { id: "active", label: "Active 3-mo Cust", color: "#8b5cf6" },
                { id: "shop",   label: "Shop Around",      color: "#0ea5e9" },
                { id: "new",    label: "New Listing",      color: "#ec4899" },
                { id: "nu",     label: "NU",               color: "#f59e0b" },
              ];

              const fmtVal = (v) => {
                if (!v) return "0";
                return Math.round(v).toLocaleString();
              };

              // Compute actual/target for one KPI, one month, one scope
              const compute = (kpiId, m, srCode, flmName) => {
                let actual = 0, target = 0;
                try {
                  if (kpiId === "sales") {
                    (RAW.actuals || []).forEach(r => {
                      if (r.m !== m) return;
                      if (srCode != null && r.sr !== srCode) return;
                      if (flmName != null && r.f !== flmName) return;
                      actual += r.v || 0;
                    });
                    (RAW.targets || []).forEach(r => {
                      if (r.m !== m) return;
                      if (srCode != null && r.sr !== srCode) return;
                      if (flmName != null && r.f !== flmName) return;
                      target += r.t || 0;
                    });
                  } else if (kpiId === "active") {
                    // activeByMonth[m] = {total, byFlm, bySr, customers}
                    const ab = (RAW.activeByMonth && RAW.activeByMonth[m]) || {};
                    if (srCode != null) {
                      actual = (ab.bySr && ab.bySr[srCode]) || 0;
                    } else if (flmName != null) {
                      actual = (ab.byFlm && ab.byFlm[flmName]) || 0;
                    } else {
                      actual = ab.total || 0;
                    }
                    const at = RAW.activeTargetByMonth && RAW.activeTargetByMonth[m];
                    if (at) {
                      if (srCode != null) {
                        target = (at.bySr && at.bySr[srCode]) || 0;
                      } else if (flmName != null) {
                        target = (at.byFlm && at.byFlm[flmName]) || 0;
                      } else {
                        Object.values(at.bySr || {}).forEach(v => target += v);
                      }
                    }
                  } else if (kpiId === "shop") {
                    const sb = (RAW.shopByMonth && RAW.shopByMonth[m]) || [];
                    sb.forEach(r => {
                      if (srCode != null && r.sr !== srCode) return;
                      if (flmName != null && r.f !== flmName) return;
                      actual += r.v || 0;
                      target += r.t || 0;
                    });
                  } else if (kpiId === "new") {
                    // newByMonth[m] = {items, byFlm, bySr}
                    const nb = (RAW.newByMonth && RAW.newByMonth[m]) || {};
                    if (srCode != null) {
                      actual = (nb.bySr && nb.bySr[srCode]) || 0;
                    } else if (flmName != null) {
                      actual = (nb.byFlm && nb.byFlm[flmName]) || 0;
                    } else {
                      // Total = sum of all customers across SRs (or count of items)
                      actual = (nb.items && nb.items.length) || 0;
                    }
                    const nt = RAW.newTargetByMonth && RAW.newTargetByMonth[m];
                    if (nt) {
                      if (srCode != null) {
                        target = (nt.bySr && nt.bySr[srCode]) || 0;
                      } else if (flmName != null) {
                        target = (nt.byFlm && nt.byFlm[flmName]) || 0;
                      } else {
                        Object.values(nt.bySr || {}).forEach(v => target += v);
                      }
                    }
                  } else if (kpiId === "nu") {
                    const lt = RAW.leadTargetByMonth && RAW.leadTargetByMonth[m];
                    const la = RAW.leadActualByMonth && RAW.leadActualByMonth[m];
                    if (lt) {
                      if (srCode != null) {
                        target = (lt.bySr && lt.bySr[srCode]) || 0;
                      } else if (flmName != null) {
                        target = (lt.byFlm && lt.byFlm[flmName]) || 0;
                      } else {
                        Object.values(lt.bySr || {}).forEach(v => target += v);
                      }
                    }
                    if (la) {
                      if (srCode != null) {
                        actual = (la.bySr && la.bySr[srCode]) || 0;
                      } else if (flmName != null) {
                        actual = (la.byFlm && la.byFlm[flmName]) || 0;
                      } else {
                        Object.values(la.bySr || {}).forEach(v => actual += v);
                      }
                    }
                  }
                } catch (e) {}
                return { actual: Math.round(actual), target: Math.round(target) };
              };

              const Cell = ({ actual, target }) => {
                if (actual === 0 && target === 0) {
                  return <td style={{...tdStyleR, padding:"4px 4px", color:"#d1d5db", fontSize:9, background:"#fafafa"}}>—</td>;
                }
                const p = pct(actual, target);
                const bg = target > 0 ? heatBg(p) : "#f3f4f6";
                const fg = target > 0 ? heatFg(p) : "#6b7280";
                return (
                  <td style={{
                    ...tdStyleR, padding:"4px 4px",
                    background: bg, color: fg,
                    borderRight:"1px solid #fff",
                    lineHeight: 1.15,
                  }}>
                    <div style={{fontSize:9, fontWeight:600}}>
                      {fmtVal(actual)}<span style={{opacity:.55, fontWeight:400}}>/{fmtVal(target)}</span>
                    </div>
                    <div style={{fontSize:10, fontWeight:700}}>
                      {target > 0 ? p.toFixed(0) + "%" : "—"}
                    </div>
                  </td>
                );
              };

              return HEATMAP_KPIS.map(kpi => {
                // Filter FLMs by current filter
                const flmList = allFlm ? (RAW.flms || []) : flmSel;

                // Build FLM-level rows
                const flmRows = flmList.map(f => {
                  const cells = MONTH_NAMES.map((mLabel, idx) => {
                    return { mLabel, ...compute(kpi.id, idx + 1, null, f) };
                  });
                  // Get SRs under this FLM (only those with data for this KPI)
                  const flmSrs = (RAW.srs || [])
                    .filter(s => s.flm === f)
                    .map(s => {
                      const srCells = MONTH_NAMES.map((mLabel, idx) => {
                        return { mLabel, ...compute(kpi.id, idx + 1, s.code, null) };
                      });
                      return { sr: s, cells: srCells };
                    })
                    .filter(r => r.cells.some(c => c.actual > 0 || c.target > 0));
                  return { flm: f, cells, srs: flmSrs };
                });

                return (
                  <Panel key={"heatmap-" + kpi.id} title={"🗺️ " + kpi.label + " — Heatmap by Month (click ▸ to expand SRs)"}
                    action={<ExportBtn onClick={() => {
                      const rows = [];
                      flmRows.forEach(r => {
                        // FLM totals row
                        const flmRow = { "Level": "FLM", "FLM": r.flm, "SR Code": "", "SR Name": "" };
                        r.cells.forEach((c, i) => {
                          flmRow[MONTH_NAMES[i] + " Act"] = c.actual;
                          flmRow[MONTH_NAMES[i] + " Tgt"] = c.target;
                        });
                        rows.push(flmRow);
                        // SR rows
                        r.srs.forEach(srRow => {
                          const sr_row = { "Level": "SR", "FLM": r.flm, "SR Code": srRow.sr.code, "SR Name": srRow.sr.name };
                          srRow.cells.forEach((c, i) => {
                            sr_row[MONTH_NAMES[i] + " Act"] = c.actual;
                            sr_row[MONTH_NAMES[i] + " Tgt"] = c.target;
                          });
                          rows.push(sr_row);
                        });
                      });
                      exportToExcel(rows, `Heatmap_${kpi.label.replace(/\s/g,"")}_${year}.xlsx`, kpi.label);
                    }} />}>
                    <div style={{overflowX:"auto"}}>
                      <table style={{...tblStyle, fontSize:10, borderCollapse:"separate", borderSpacing:0}}>
                        <thead>
                          <tr style={{background:"#f9fafb"}}>
                            <th style={{...thStyle, position:"sticky", left:0, background:"#f9fafb", zIndex:1, minWidth:170}}>FLM / SR</th>
                            {MONTH_NAMES.map(m => (
                              <th key={m} style={{...thStyleR, fontSize:9, minWidth:70}}>{m}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {flmRows.map(r => (
                            <React.Fragment key={r.flm}>
                              {/* FLM row (clickable) */}
                              <tr
                                onClick={() => setExpanded(e => ({ ...e, ["hm_" + kpi.id + "_" + r.flm]: !e["hm_" + kpi.id + "_" + r.flm] }))}
                                style={{borderTop:"2px solid #e5e7eb", background:"#fafbfc", cursor:"pointer"}}
                              >
                                <td style={{...tdStyle, position:"sticky", left:0, background:"#fafbfc", zIndex:1, fontWeight:700, fontSize:11}}>
                                  <span style={{color:"#6b7280", marginRight:6}}>
                                    {expanded["hm_" + kpi.id + "_" + r.flm] ? "▾" : "▸"}
                                  </span>
                                  {r.flm}
                                  <span style={{fontSize:9, color:"#9ca3af", marginLeft:6, fontWeight:400}}>
                                    {r.srs.length} SRs
                                  </span>
                                </td>
                                {r.cells.map((c, i) => (
                                  <Cell key={i} actual={c.actual} target={c.target} />
                                ))}
                              </tr>
                              {/* SR rows (only when FLM expanded) */}
                              {expanded["hm_" + kpi.id + "_" + r.flm] && r.srs.map(srRow => (
                                <tr key={srRow.sr.code} style={{borderTop:"1px solid #f3f4f6"}}>
                                  <td style={{...tdStyle, position:"sticky", left:0, background:"#fff", zIndex:1, padding:"4px 8px 4px 32px", fontSize:11}}>
                                    <span style={{fontFamily:"monospace", color:"#9ca3af", fontSize:9, marginRight:6}}>{srRow.sr.code}</span>
                                    {srRow.sr.name}
                                  </td>
                                  {srRow.cells.map((c, i) => (
                                    <Cell key={i} actual={c.actual} target={c.target} />
                                  ))}
                                </tr>
                              ))}
                            </React.Fragment>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div style={{fontSize:9, color:"#9ca3af", marginTop:8}}>
                      Each cell: actual/target (top), % (bottom). Click any FLM row to expand its SRs.
                    </div>
                  </Panel>
                );
              });
            })()}
            </>
          );
        } catch (e) {
          return (
            <Panel title="📈 Trend Analysis — Error">
              <div style={{padding:20, color:"#dc2626"}}>
                <div style={{fontWeight:700, marginBottom:6}}>Could not render Trend tab</div>
                <div style={{fontSize:11, fontFamily:"monospace"}}>{e.message}</div>
              </div>
            </Panel>
          );
        }
      })()}

      {/* ============ CUSTOMER SALES (monthly matrix + lapsed analysis) ============ */}
      {tab === "custsales" && (() => {
        if (!RAW.customerMonthly) return (
          <RedeployNotice title="📅 Customer Sales" feature="The Customer Sales matrix" />
        );
        const maxP = RAW.customerMonthlyMaxPeriod || (year * 100 + month);
        const periods = [];
        { let y = 2025, mm = 1; while (y * 100 + mm <= maxP) { periods.push(y * 100 + mm); mm++; if (mm > 12) { mm = 1; y++; } } }
        const pLabel = (p) => MONTH_NAMES[(p % 100) - 1] + "-" + String(Math.floor(p / 100)).slice(2);

        // Group the months by calendar year (auto-extends as new months/years arrive).
        const yearGroups = [];
        periods.forEach(p => {
          const y = Math.floor(p / 100);
          let g = yearGroups.find(x => x.year === y);
          if (!g) { g = { year: y, periods: [] }; yearGroups.push(g); }
          g.periods.push(p);
        });

        const rows = (RAW.customerMonthly || []).map(cu => {
          const srObj = cu.sr ? RAW.srs.find(s => s.code === cu.sr) : null;
          const srName = srObj ? srObj.name : (cu.sr ? ("SR " + cu.sr) : "—");
          const yr = {};
          yearGroups.forEach(g => {
            const sum = g.periods.reduce((s, p) => s + (cu.p[p] || 0), 0);
            yr[g.year] = { sum, avg: sum / g.periods.length };
          });
          let gap = 0;
          for (let i = periods.length - 1; i >= 0; i--) { if ((cu.p[periods[i]] || 0) > 0) break; gap++; }
          const total = periods.reduce((s, p) => s + (cu.p[p] || 0), 0);
          return { ...cu, srName, yr, gap, total };
        });

        let scoped = rows;
        if (!allFlm) scoped = scoped.filter(r => flmMatch(r.f));
        // Match the customer's master SR, any rep who actually SOLD to it, OR any
        // assigned PND/MND rep — so filtering by a rep shows every outlet they
        // touched in the data (not just where they're the master rep).
        if (!allSr) scoped = scoped.filter(r => srMatch(r.sr) || (r.srs || []).some(srMatch) || (r.asg || []).some(srMatch));
        if (custSearch.trim()) {
          const q = custSearch.trim().toLowerCase();
          scoped = scoped.filter(r => String(r.c).includes(q) || String(r.n || "").toLowerCase().includes(q));
        }
        const cnt3 = scoped.filter(r => r.gap >= 3).length;
        const cnt6 = scoped.filter(r => r.gap >= 6).length;
        const cnt12 = scoped.filter(r => r.gap >= 12).length;

        let view = scoped;
        if (lapseFilter === "3") view = view.filter(r => r.gap >= 3);
        else if (lapseFilter === "6") view = view.filter(r => r.gap >= 6);
        else if (lapseFilter === "12") view = view.filter(r => r.gap >= 12);

        const sortVal = (r, key) => {
          if (key === "name") return String(r.n || "").toLowerCase();
          if (key === "gap") return r.gap;
          if (key === "total") return r.total;
          if (key[0] === "p") return r.p[Number(key.slice(1))] || 0;
          if (key.slice(0, 2) === "yt") return (r.yr[Number(key.slice(2))] || {}).sum || 0;
          if (key.slice(0, 2) === "ya") return (r.yr[Number(key.slice(2))] || {}).avg || 0;
          return 0;
        };
        const sdir = custSortDir === "asc" ? 1 : -1;
        view = [...view].sort((a, b) => {
          const va = sortVal(a, custSortKey), vb = sortVal(b, custSortKey);
          if (typeof va === "string") return va.localeCompare(vb) * sdir;
          return (va - vb) * sdir;
        });
        const RENDER_CAP = 5000; // effectively show every customer (Export still gives the full list)
        const shown = view.slice(0, RENDER_CAP);
        const truncated = view.length > RENDER_CAP;
        const sortBy = (key, defDir) => {
          if (custSortKey === key) setCustSortDir(d => d === "asc" ? "desc" : "asc");
          else { setCustSortKey(key); setCustSortDir(defDir || "desc"); }
        };
        const caret = (key) => custSortKey === key ? (custSortDir === "asc" ? " ▲" : " ▼") : "";

        const stickyHead = { position: "sticky", top: 0, zIndex: 2, background: "#f9fafb" };
        const stickyName = { position: "sticky", left: 0, zIndex: 1, background: "#fff", minWidth: 160, maxWidth: 160 };
        const stickyCorner = { ...stickyHead, left: 0, zIndex: 3, minWidth: 160, maxWidth: 160 };
        const sumColTh = { ...stickyHead, ...thStyleR, borderLeft: "2px solid #cbd5e1", whiteSpace: "nowrap", background: "#eef2ff" };
        const sumColTd = { ...tdStyleR, borderLeft: "2px solid #e5e7eb", fontWeight: 700, background: "#f8fafc", whiteSpace: "nowrap" };
        const colCount = 3 + yearGroups.reduce((s, g) => s + g.periods.length + 2, 0) + 1;

        const exportRows = () => {
          const out = view.map(r => {
            const o = { "Customer Code": r.c, "Customer Name": r.n, "SR": r.srName, "Manager (FLM)": r.f || "—" };
            yearGroups.forEach(g => {
              g.periods.forEach(p => { o[pLabel(p)] = Math.round(r.p[p] || 0); });
              o[g.year + " Total"] = Math.round(r.yr[g.year].sum);
              o[g.year + " Avg"] = Math.round(r.yr[g.year].avg);
            });
            o["Months Since Last Purchase"] = r.gap;
            return o;
          });
          exportToExcel(out, `CustomerSales_to_${pLabel(maxP)}.xlsx`, "Customer Sales");
        };

        const cell = (v) => {
          if (v === 0) return { ...tdStyleR, background: "#fef0c7", color: "#dc2626", fontWeight: 600 };
          if (v < 0) return { ...tdStyleR, color: "#b45309" };
          return tdStyleR;
        };

        return (
          <>
            <div style={{display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:8, marginBottom:10}}>
              <Metric label="Customers (in scope)" value={String(scoped.length)} accent="#2563eb" />
              <Metric label="No Purchase 3M+" value={String(cnt3)} accent="#f59e0b"
                info="Customers with zero sales in each of the last 3 consecutive months (up to the latest month in the data)." />
              <Metric label="No Purchase 6M+" value={String(cnt6)} accent="#ea580c" />
              <Metric label="No Purchase 12M+" value={String(cnt12)} accent="#dc2626" />
            </div>

            <Panel title={"Customer Monthly Sales (Jan-25 → " + pLabel(maxP) + ") · per-year Total & Avg · 0 in red"}
              action={
                <div style={{marginLeft:"auto", display:"flex", gap:6, alignItems:"center", flexWrap:"wrap"}}>
                  <input value={custSearch} onChange={e => setCustSearch(e.target.value)}
                    placeholder="🔍 Search customer / code…"
                    style={{...selectStyle, minWidth:180, padding:"4px 8px"}} />
                  <select value={lapseFilter} onChange={e => setLapseFilter(e.target.value)} style={selectStyle}>
                    <option value="all">All customers</option>
                    <option value="3">Not purchased 3M+</option>
                    <option value="6">Not purchased 6M+</option>
                    <option value="12">Not purchased 12M+</option>
                  </select>
                  <ExportBtn onClick={exportRows} />
                </div>
              }>
              <div style={{fontSize:10, color:"#6b7280", marginBottom:6}}>
                {truncated
                  ? "Showing top " + RENDER_CAP + " of " + view.length.toLocaleString() + " — use search / sort to narrow, or Export for the full list."
                  : "Showing " + view.length.toLocaleString() + " customers"} · click a column header to sort · filters apply · header frozen.
              </div>
              <div style={{overflow:"auto", maxHeight:"70vh", border:"1px solid #f3f4f6", borderRadius:6}}>
                <table style={{...tblStyle, fontSize:11}}>
                  <thead>
                    <tr>
                      <th onClick={() => sortBy("name", "asc")} style={{...stickyCorner, ...thStyle, textAlign:"left", cursor:"pointer"}}>Customer{caret("name")}</th>
                      <th style={{...stickyHead, ...thStyle}}>SR</th>
                      <th style={{...stickyHead, ...thStyle}}>Manager</th>
                      {yearGroups.map(g => (
                        <React.Fragment key={g.year}>
                          {g.periods.map(p => (
                            <th key={p} onClick={() => sortBy("p"+p)} style={{...stickyHead, ...thStyleR, whiteSpace:"nowrap", cursor:"pointer"}}>{pLabel(p)}{caret("p"+p)}</th>
                          ))}
                          <th onClick={() => sortBy("yt"+g.year)} style={{...sumColTh, cursor:"pointer"}}>{g.year} Total{caret("yt"+g.year)}</th>
                          <th onClick={() => sortBy("ya"+g.year)} style={{...sumColTh, borderLeft:"1px solid #c7d2fe", cursor:"pointer"}}>{g.year} Avg{caret("ya"+g.year)}</th>
                        </React.Fragment>
                      ))}
                      <th onClick={() => sortBy("gap")} style={{...stickyHead, ...thStyleR, borderLeft:"2px solid #cbd5e1", cursor:"pointer"}}>Gap{caret("gap")}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {view.length === 0 && (
                      <tr><td style={tdStyle} colSpan={colCount}>No customers match the current filters.</td></tr>
                    )}
                    {shown.map(r => (
                      <tr key={r.c} style={{borderTop:"1px solid #f3f4f6"}}>
                        <td style={{...stickyName, ...tdStyle}}>
                          <div style={{fontWeight:600, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis", maxWidth:152}}>{r.n}</div>
                          <div style={{fontSize:9, color:"#9ca3af", fontFamily:"monospace"}}>{r.c}</div>
                        </td>
                        <td style={tdStyle}>{r.srName}</td>
                        <td style={tdStyle}>{r.f || "—"}</td>
                        {yearGroups.map(g => (
                          <React.Fragment key={g.year}>
                            {g.periods.map(p => {
                              const v = r.p[p] || 0;
                              return <td key={p} style={cell(v)}>{fmt(v)}</td>;
                            })}
                            <td style={sumColTd}>{fmt(r.yr[g.year].sum)}</td>
                            <td style={{...sumColTd, borderLeft:"1px solid #eef2ff", color:"#4338ca"}}>{fmt(r.yr[g.year].avg)}</td>
                          </React.Fragment>
                        ))}
                        <td style={{...tdStyleR, borderLeft:"2px solid #e5e7eb", fontWeight:700,
                          color: r.gap >= 12 ? "#dc2626" : r.gap >= 6 ? "#ea580c" : r.gap >= 3 ? "#d97706" : "#059669"}}>
                          {r.gap}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Panel>
          </>
        );
      })()}

      {/* ============ ACCESS LOG (admin only) ============ */}
      {tab === "access" && (() => {
        if (user.role !== "Admin") return null;
        if (accessErr && /unknown action|not found/i.test(accessErr)) return (
          <RedeployNotice title="🔑 Access Log" feature="The Access Log" />
        );
        if (accessErr) return (
          <Panel title="🔑 Access Log — Error">
            <div style={{padding:20, color:"#dc2626", fontSize:12}}>{accessErr}</div>
          </Panel>
        );
        if (accessLog === null) return (
          <Panel title="🔑 Access Log">
            <div style={{padding:20, color:"#6b7280", fontSize:12}}>Loading access log…</div>
          </Panel>
        );

        const rows = accessLog;
        const dayObj = (ts) => { const d = new Date(ts); return isNaN(d.getTime()) ? null : d; };
        const dayKey = (ts) => { const d = dayObj(ts); return d ? d.toLocaleDateString() : String(ts).slice(0,10); };
        const fmtTs = (ts) => { const d = dayObj(ts); return d ? d.toLocaleString() : String(ts); };
        const succ = rows.filter(r => String(r.status).toLowerCase() === "success");
        const todayStr = new Date().toDateString();

        const byUserMap = {};
        succ.forEach(r => {
          const key = r.name || ("Code " + r.code);
          if (!byUserMap[key]) byUserMap[key] = { name: key, role: r.role, flm: r.flm, count: 0, last: null };
          byUserMap[key].count++;
          const d = dayObj(r.ts);
          if (d && (!byUserMap[key].last || d > byUserMap[key].last)) byUserMap[key].last = d;
        });
        const byUser = Object.values(byUserMap).sort((a,b) => b.count - a.count);

        const byDayMap = {};
        succ.forEach(r => {
          const k = dayKey(r.ts);
          if (!byDayMap[k]) byDayMap[k] = { day: k, count: 0, users: new Set(), sort: dayObj(r.ts) };
          byDayMap[k].count++;
          byDayMap[k].users.add(r.name || r.code);
        });
        const byDay = Object.values(byDayMap)
          .map(d => ({ day: d.day, count: d.count, users: d.users.size, sort: d.sort }))
          .sort((a,b) => (b.sort ? b.sort.getTime() : 0) - (a.sort ? a.sort.getTime() : 0));

        const totalLogins = succ.length;
        const uniqueUsers = byUser.length;
        const loginsToday = succ.filter(r => { const d = dayObj(r.ts); return d && d.toDateString() === todayStr; }).length;
        const failed = rows.length - succ.length;
        const recent = [...rows]
          .sort((a,b) => ((dayObj(b.ts)?dayObj(b.ts).getTime():0) - (dayObj(a.ts)?dayObj(a.ts).getTime():0)))
          .slice(0, 100);

        return (
          <>
            <div style={{display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:8, marginBottom:10}}>
              <Metric label="Total Logins" value={String(totalLogins)} accent="#2563eb" />
              <Metric label="Unique Users" value={String(uniqueUsers)} accent="#10b981" />
              <Metric label="Logins Today" value={String(loginsToday)} accent="#f59e0b" />
              <Metric label="Failed Attempts" value={String(failed)} accent="#dc2626" />
            </div>

            <Panel title="Logins by Day"
              action={<ExportBtn onClick={() => exportToExcel(
                byDay.map(d => ({ "Date": d.day, "Logins": d.count, "Unique Users": d.users })),
                "AccessLog_ByDay.xlsx", "By Day")} />}>
              <div style={{overflowX:"auto"}}>
                <table style={tblStyle}>
                  <thead><tr style={{background:"#f9fafb"}}>
                    <th style={thStyle}>Date</th>
                    <th style={thStyleR}>Logins</th>
                    <th style={thStyleR}>Unique Users</th>
                  </tr></thead>
                  <tbody>
                    {byDay.length === 0 && <tr><td style={tdStyle} colSpan={3}>No logins recorded yet.</td></tr>}
                    {byDay.map(d => (
                      <tr key={d.day} style={{borderTop:"1px solid #f3f4f6"}}>
                        <td style={tdStyle}>{d.day}</td>
                        <td style={tdStyleR}>{d.count}</td>
                        <td style={tdStyleR}>{d.users}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Panel>

            <Panel title="By User"
              action={<ExportBtn onClick={() => exportToExcel(
                byUser.map(u => ({ "User": u.name, "Role": u.role, "FLM": u.flm, "Logins": u.count, "Last Login": u.last ? u.last.toLocaleString() : "" })),
                "AccessLog_ByUser.xlsx", "By User")} />}>
              <div style={{overflowX:"auto"}}>
                <table style={tblStyle}>
                  <thead><tr style={{background:"#f9fafb"}}>
                    <th style={thStyle}>User</th>
                    <th style={thStyle}>Role</th>
                    <th style={thStyle}>FLM</th>
                    <th style={thStyleR}>Logins</th>
                    <th style={thStyle}>Last Login</th>
                  </tr></thead>
                  <tbody>
                    {byUser.length === 0 && <tr><td style={tdStyle} colSpan={5}>No logins recorded yet.</td></tr>}
                    {byUser.map(u => (
                      <tr key={u.name} style={{borderTop:"1px solid #f3f4f6"}}>
                        <td style={tdStyle}><strong>{u.name}</strong></td>
                        <td style={tdStyle}>{u.role || "—"}</td>
                        <td style={tdStyle}>{u.flm || "—"}</td>
                        <td style={tdStyleR}>{u.count}</td>
                        <td style={tdStyle}>{u.last ? u.last.toLocaleString() : "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Panel>

            <Panel title="Recent Activity (last 100)"
              action={<ExportBtn onClick={() => exportToExcel(
                [...rows].sort((a,b) => ((dayObj(b.ts)?dayObj(b.ts).getTime():0) - (dayObj(a.ts)?dayObj(a.ts).getTime():0)))
                  .map(r => ({ "Time": fmtTs(r.ts), "Name": r.name, "Role": r.role, "FLM": r.flm, "Code": r.code, "Status": r.status })),
                "AccessLog_Full.xlsx", "Access Log")} />}>
              <div style={{overflowX:"auto"}}>
                <table style={tblStyle}>
                  <thead><tr style={{background:"#f9fafb"}}>
                    <th style={thStyle}>Time</th>
                    <th style={thStyle}>Name</th>
                    <th style={thStyle}>Role</th>
                    <th style={thStyle}>FLM</th>
                    <th style={thStyle}>Status</th>
                  </tr></thead>
                  <tbody>
                    {recent.length === 0 && <tr><td style={tdStyle} colSpan={5}>No activity yet.</td></tr>}
                    {recent.map((r, i) => (
                      <tr key={i} style={{borderTop:"1px solid #f3f4f6"}}>
                        <td style={tdStyle}>{fmtTs(r.ts)}</td>
                        <td style={tdStyle}>{r.name || "—"}</td>
                        <td style={tdStyle}>{r.role || "—"}</td>
                        <td style={tdStyle}>{r.flm || "—"}</td>
                        <td style={{...tdStyle, color: String(r.status).toLowerCase()==="success" ? "#059669" : "#dc2626", fontWeight:600}}>{r.status}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Panel>
          </>
        );
      })()}

      <div style={{marginTop:14, fontSize:10, color:"#9ca3af", textAlign:"center"}}>
        Source: Daily Sales · Target Set (Ethical) · Material Code Lookup ·
        {RAW.srs.length} SRs · {RAW.flms.length} FLMs · {RAW.customers.length} customers · auto-refresh 5s
      </div>

      <style>{`@keyframes pulse { 0%,100% {opacity:1} 50% {opacity:.4} }`}</style>
    </div>
  );
}

// === Components ===
function FilterField({ label, children }) {
  return (
    <div>
      <div style={{fontSize:9, color:"#6b7280", textTransform:"uppercase",
        letterSpacing:.5, fontWeight:600, marginBottom:3}}>{label}</div>
      {children}
    </div>
  );
}

// Checkbox dropdown for multi-select filters. `selected` is an array; [] = All.
function MultiSelect({ options, selected, onChange, disabled, allLabel, searchable }) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState("");
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return;
    const h = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, [open]);
  const isSel = (v) => selected.indexOf(v) >= 0;
  const toggle = (v) => { onChange(isSel(v) ? selected.filter(x => x !== v) : selected.concat([v])); };
  const shown = (searchable && q)
    ? options.filter(o => String(o.label).toLowerCase().indexOf(q.toLowerCase()) >= 0)
    : options;
  const summary = selected.length === 0 ? allLabel
    : selected.length === 1 ? ((options.find(o => o.value === selected[0]) || {}).label || String(selected[0]))
    : selected.length + " selected";
  return (
    <div ref={ref} style={{ position: "relative" }}>
      <button type="button" disabled={disabled} onClick={() => !disabled && setOpen(o => !o)}
        style={{ ...selectStyle, textAlign: "left", background: disabled ? "#f3f4f6" : "#fff",
          cursor: disabled ? "not-allowed" : "pointer", color: disabled ? "#6b7280" : "#111827",
          whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
        {summary}<span style={{ float: "right", color: "#9ca3af" }}>▾</span>
      </button>
      {open && !disabled && (
        <div style={{ position: "absolute", zIndex: 30, top: "calc(100% + 2px)", left: 0, right: 0,
          minWidth: 190, maxHeight: 300, overflowY: "auto", background: "#fff", border: "1px solid #d1d5db",
          borderRadius: 8, boxShadow: "0 8px 24px rgba(0,0,0,.12)", padding: 4 }}>
          <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 6px" }}>
            <span style={{ fontSize: 11, color: "#2563eb", cursor: "pointer" }} onClick={() => onChange([])}>Clear (All)</span>
            <span style={{ fontSize: 11, color: "#6b7280", cursor: "pointer" }} onClick={() => onChange(options.map(o => o.value))}>Select all</span>
          </div>
          {searchable && (
            <input value={q} onChange={e => setQ(e.target.value)} placeholder="Search…"
              style={{ width: "100%", boxSizing: "border-box", fontSize: 11, padding: "4px 6px", margin: "2px 0",
                border: "1px solid #e5e7eb", borderRadius: 6 }} />
          )}
          {shown.slice(0, 500).map(o => (
            <label key={o.value} style={{ display: "flex", alignItems: "center", gap: 6, padding: "4px 6px",
              fontSize: 12, cursor: "pointer" }}>
              <input type="checkbox" checked={isSel(o.value)} onChange={() => toggle(o.value)} />
              <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{o.label}</span>
            </label>
          ))}
          {shown.length > 500 && (
            <div style={{ padding: "6px", fontSize: 10.5, color: "#6b7280", background: "#f9fafb" }}>
              Showing 500 of {shown.length.toLocaleString()} — type above to find any customer.
            </div>
          )}
          {shown.length === 0 && <div style={{ padding: 8, fontSize: 11, color: "#9ca3af" }}>No matches</div>}
        </div>
      )}
    </div>
  );
}

function TabBtn({ label, v, cur, on, color, small }) {
  const active = cur === v;
  return (
    <button onClick={() => on(v)} style={{
      background: active ? "#fff" : "transparent", border:"none",
      padding: small ? "8px 8px" : "9px 14px",
      fontSize: small ? 10.5 : 11.5,
      cursor:"pointer", fontWeight: active ? 600 : 500,
      color: active ? (color || "#2563eb") : "#6b7280",
      borderBottom: active ? "2px solid " + (color || "#2563eb") : "2px solid transparent",
      marginBottom: -1,
    }}>{label}</button>
  );
}

function Sep() {
  return <div style={{borderLeft:"1px solid #e5e7eb", margin:"0 4px", height:22, alignSelf:"center"}} />;
}

const ACTIVE_3M_INFO = "Active 3-mo = distinct customers with at least one Ethical purchase in the trailing 3-month window (the selected month plus the 2 months before it; it rolls into the previous year for early months — e.g. Jan counts Nov + Dec + Jan). Each active customer is credited to their SR(s); a shared customer counts for every SR that shares them, so per-SR totals can add up to more than the distinct customer count. Target comes from the 'Target-Active Cus 3 Months' sheet.";

// Small clickable info icon with a tap/click tooltip (works on mobile too).
function InfoDot({ text }) {
  const [open, setOpen] = useState(false);
  return (
    <span style={{position:"relative", display:"inline-block"}}>
      <span
        onClick={(e) => { e.stopPropagation(); setOpen(o => !o); }}
        style={{
          display:"inline-flex", alignItems:"center", justifyContent:"center",
          width:14, height:14, borderRadius:"50%", background:"#e0e7ff",
          color:"#4338ca", fontSize:10, fontWeight:700, cursor:"pointer",
          marginLeft:5, verticalAlign:"middle", fontStyle:"italic",
        }}
        title="Click for details"
      >i</span>
      {open && (
        <>
          <span onClick={(e) => { e.stopPropagation(); setOpen(false); }}
            style={{position:"fixed", inset:0, zIndex:40}} />
          <span style={{
            position:"absolute", top:"135%", left:0, zIndex:50,
            background:"#111827", color:"#fff", fontSize:10.5, fontWeight:400,
            lineHeight:1.55, padding:"9px 11px", borderRadius:6, width:265,
            boxShadow:"0 4px 14px rgba(0,0,0,.25)", textAlign:"left",
            textTransform:"none", letterSpacing:0, whiteSpace:"normal",
          }}>{text}</span>
        </>
      )}
    </span>
  );
}

function Metric({ label, value, actual, pct, accent, sub, secondaryLabel, info }) {
  return (
    <div style={{
      background:"#fff", border:"1px solid #e5e7eb",
      borderTop:"3px solid " + accent, borderRadius:8, padding:"10px 12px",
    }}>
      <div style={{fontSize:9, color:"#6b7280", textTransform:"uppercase", letterSpacing:.5, fontWeight:600}}>
        {label}{info && <InfoDot text={info} />}
      </div>
      <div style={{display:"flex", alignItems:"baseline", gap:4, marginTop:3, flexWrap:"wrap"}}>
        <div style={{fontSize:18, fontWeight:700, color:"#111827"}}>{value}</div>
        {actual !== undefined && actual !== null && actual !== "—" && (
          <div style={{fontSize:11, color:"#6b7280", fontWeight:500}}>
            {secondaryLabel ? secondaryLabel + " " : "/ "}{actual}
          </div>
        )}
      </div>
      {pct !== null && pct !== undefined && (
        <div style={{fontSize:11, fontWeight:600, marginTop:2, color: pctColor(pct)}}>
          {pct.toFixed(1)}%
        </div>
      )}
      {sub && (
        <div style={{fontSize:9, color:"#9ca3af", marginTop:2}}>{sub}</div>
      )}
    </div>
  );
}

function Panel({ title, children, action }) {
  return (
    <div style={{
      background:"#fff", border:"1px solid #e5e7eb", borderRadius:8,
      padding:12, marginBottom:10,
    }}>
      <div style={{
        fontSize:12, fontWeight:600, color:"#111827", marginBottom:10,
        display:"flex", alignItems:"center", gap:8, flexWrap:"wrap",
      }}>
        <span style={{flex:"1 1 auto", minWidth:0}}>{title}</span>
        {action}
      </div>
      {children}
    </div>
  );
}

// Ranked coverage scorecard: medal cards for the top 3 + a full ranked table.
// rows: [{ key, name, sub, kpis: [{ label, actual, target }] }]
function CoverageRanking({ title, rows, onExport }) {
  const scored = rows.map(r => {
    const withT = r.kpis.filter(k => k.target > 0);
    const hits = withT.filter(k => k.actual >= k.target).length;
    const score = withT.length
      ? withT.reduce((s, k) => s + Math.min((k.actual / k.target) * 100, 100), 0) / withT.length
      : 0;
    return { ...r, hits, misses: withT.length - hits, hitsTotal: withT.length, score };
  }).sort((a, b) => b.score - a.score);
  scored.forEach((r, i) => { r.rank = i + 1; });

  const tierOf = (s) => s >= 90 ? "Gold" : s >= 75 ? "Silver" : s >= 60 ? "Bronze" : "—";
  const rankColor = (rank) => rank === 1 ? "#f59e0b" : rank === 2 ? "#9ca3af" : rank === 3 ? "#c2410c" : "#9ca3af";
  const top3 = scored.slice(0, 3);

  const TierBadge = ({ tier }) => (
    <span style={{
      background: tier === "Gold" ? "#fef3c7" : tier === "Silver" ? "#f3f4f6" : tier === "Bronze" ? "#fff7ed" : "#f9fafb",
      color: tier === "Gold" ? "#92400e" : tier === "Silver" ? "#4b5563" : tier === "Bronze" ? "#9a3412" : "#9ca3af",
      border: "1px solid " + (tier === "Gold" ? "#fcd34d" : tier === "Silver" ? "#d1d5db" : tier === "Bronze" ? "#fdba74" : "#e5e7eb"),
      borderRadius: 6, padding: "2px 10px", fontSize: 11, fontWeight: 700, whiteSpace: "nowrap",
    }}>{tier}</span>
  );

  const Tile = ({ label, actual, target }) => {
    const has = target > 0;
    const p = has ? (actual / target) * 100 : 0;
    const col = !has ? "#9ca3af" : actual >= target ? "#059669" : "#dc2626";
    return (
      <div style={{textAlign:"center"}}>
        <div style={{fontSize:10, color:"#6b7280", fontWeight:600}}>{label}</div>
        <div style={{fontSize:16, fontWeight:700, color:col}}>{has ? p.toFixed(0) + "%" : "—"}</div>
        <div style={{fontSize:9, color:"#9ca3af"}}>{fmt(actual)} / {fmt(target)}</div>
        <div style={{fontSize:10, fontWeight:600, color:col}}>
          {has ? ((actual - target >= 0 ? "+" : "") + fmt(actual - target)) : ""}
        </div>
      </div>
    );
  };

  return (
    <Panel title={title} action={<ExportBtn onClick={onExport} />}>
      <div style={{display:"grid", gridTemplateColumns:"1fr", gap:8, marginBottom:10}}>
        {top3.map(r => (
          <div key={r.key} style={{
            border:"1px solid #e5e7eb", borderTop:"4px solid " + rankColor(r.rank),
            borderRadius:8, padding:"10px 14px", background:"#fff",
          }}>
            <div style={{display:"flex", alignItems:"center", gap:10}}>
              <span style={{fontSize:22, fontWeight:800, color:rankColor(r.rank)}}>#{r.rank}</span>
              <div style={{flex:1, minWidth:0}}>
                <div style={{fontSize:14, fontWeight:700, color:"#111827"}}>{r.name}</div>
                {r.sub && <div style={{fontSize:11, color:"#6b7280"}}>{r.sub}</div>}
              </div>
              <TierBadge tier={tierOf(r.score)} />
            </div>
            <div style={{display:"flex", alignItems:"baseline", gap:10, margin:"6px 0 10px"}}>
              <span style={{fontSize:11, color:"#6b7280", fontWeight:600}}>SCORE</span>
              <span style={{fontSize:26, fontWeight:800, color:"#111827"}}>{r.score.toFixed(1)}</span>
              <span style={{marginLeft:"auto", fontSize:12, fontWeight:700}}>
                <span style={{color:"#059669"}}>● {r.hits}</span>{" "}
                <span style={{color:"#dc2626"}}>● {r.misses}</span>
              </span>
            </div>
            <div style={{display:"grid", gridTemplateColumns:"repeat(" + r.kpis.length + ",1fr)", gap:6}}>
              {r.kpis.map(k => <Tile key={k.label} {...k} />)}
            </div>
          </div>
        ))}
      </div>

      <div style={{overflowX:"auto"}}>
        <table style={tblStyle}>
          <thead><tr style={{background:"#f9fafb"}}>
            <th style={thStyleR}>#</th>
            <th style={thStyle}>Name</th>
            {scored[0] && scored[0].kpis.map(k => <th key={k.label} style={thStyleR}>{k.label} %</th>)}
            <th style={{...thStyleR, borderLeft:"1px solid #e5e7eb"}}>Hits</th>
            <th style={thStyleR}>Score</th>
            <th style={thStyle}>Tier</th>
          </tr></thead>
          <tbody>
            {scored.map(r => (
              <tr key={r.key} style={{borderTop:"1px solid #f3f4f6", background: r.rank <= 3 ? "#fffdf5" : "#fff"}}>
                <td style={{...tdStyleR, fontWeight:700, color:rankColor(r.rank)}}>{r.rank}</td>
                <td style={tdStyle}>
                  <strong>{r.name}</strong>
                  {r.sub && <span style={{color:"#9ca3af", fontSize:10}}> · {r.sub}</span>}
                </td>
                {r.kpis.map(k => {
                  const has = k.target > 0;
                  const p = has ? (k.actual / k.target) * 100 : 0;
                  return (
                    <td key={k.label} style={{...tdStyleR, color: has ? pctColor(p) : "#d1d5db", fontWeight:600}}>
                      {has ? p.toFixed(0) + "%" : "—"}
                    </td>
                  );
                })}
                <td style={{...tdStyleR, borderLeft:"1px solid #e5e7eb"}}>{r.hits}/{r.hitsTotal}</td>
                <td style={{...tdStyleR, fontWeight:700}}>{r.score.toFixed(1)}</td>
                <td style={tdStyle}><TierBadge tier={tierOf(r.score)} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Panel>
  );
}

// Shown when a tab needs backend code that hasn't been redeployed yet.
function RedeployNotice({ title, feature }) {
  return (
    <Panel title={title}>
      <div style={{padding:"16px 18px", background:"#fffbeb", border:"1px solid #fcd34d", borderRadius:8, color:"#92400e", fontSize:12.5, lineHeight:1.6}}>
        <div style={{fontWeight:700, marginBottom:6}}>⏳ Backend update needed</div>
        <div>
          {feature} needs the latest Apps Script backend, which hasn't been deployed yet.
          The dashboard front-end is up to date, but the Google Apps Script must be
          re-deployed for this data to load.
        </div>
        <div style={{marginTop:8, fontSize:11, color:"#a16207"}}>
          In the Apps Script editor: paste the latest <code>APPS_SCRIPT_Code.gs</code>, then
          <strong> Deploy → Manage deployments → New version → Deploy</strong>, then open
          <code> …/exec?action=clearCache</code> once and reload.
        </div>
      </div>
    </Panel>
  );
}

function Bar2({ pct }) {
  const v = Math.min(pct, 150);
  return (
    <div style={{background:"#f3f4f6", borderRadius:3, height:6, overflow:"hidden", minWidth:60, marginTop:4}}>
      <div style={{
        width: ((v/150)*100) + "%", height:"100%",
        background: pctColor(pct), transition:"width .4s ease",
      }} />
    </div>
  );
}

// Reusable FLM → SR target-vs-actual matrix (expand an FLM to see its SRs).
// TM/AM are { bySr, byFlm } maps for the selected month.
function KpiMatrix({ TM, AM, flmList, RAW, expanded, setExpanded, keyPrefix }) {
  const tm = TM || { bySr:{}, byFlm:{} }, am = AM || { bySr:{}, byFlm:{} };
  return (
    <table style={tblStyle}>
      <thead><tr style={{background:"#f9fafb"}}>
        <th style={thStyle}>FLM / SR</th>
        <th style={thStyleR}>Target</th>
        <th style={thStyleR}>Actual</th>
        <th style={thStyleR}>%</th>
        <th style={thStyle}>Progress</th>
      </tr></thead>
      <tbody>
        {(flmList).map(f => {
          const tt = (tm.byFlm && tm.byFlm[f]) || 0;
          const ta = (am.byFlm && am.byFlm[f]) || 0;
          const tp = pct(ta, tt);
          const flmSrs = RAW.srs.filter(s => s.flm === f);
          const ek = keyPrefix + "_" + f;
          return (
            <React.Fragment key={f}>
              <tr style={{borderTop:"1px solid #e5e7eb", background:"#fafbfc", cursor:"pointer", fontWeight:600}}
                onClick={() => setExpanded(e => ({ ...e, [ek]: !e[ek] }))}>
                <td style={tdStyle}>
                  <span style={{color:"#6b7280", marginRight:6}}>{expanded[ek] ? "▾" : "▸"}</span>
                  {f}
                  <span style={{fontSize:10, color:"#9ca3af", marginLeft:6, fontWeight:400}}>{flmSrs.length} SRs</span>
                </td>
                <td style={tdStyleR}>{tt}</td>
                <td style={tdStyleR}>{ta}</td>
                <td style={{...tdStyleR, color:pctColor(tp), fontWeight:700}}>{tt > 0 ? tp.toFixed(0) + "%" : "—"}</td>
                <td style={tdStyle}><Bar2 pct={tp} /></td>
              </tr>
              {expanded[ek] && flmSrs.map(s => {
                const t = (tm.bySr && tm.bySr[s.code]) || 0;
                const a = (am.bySr && am.bySr[s.code]) || 0;
                if (t === 0 && a === 0) return null;
                const sp = pct(a, t);
                return (
                  <tr key={s.code} style={{borderTop:"1px solid #f3f4f6"}}>
                    <td style={{...tdStyle, paddingLeft:32, fontSize:11.5}}>
                      <span style={{fontFamily:"monospace", color:"#9ca3af", fontSize:10, marginRight:6}}>{s.code}</span>
                      {s.name}
                    </td>
                    <td style={tdStyleR}>{t}</td>
                    <td style={tdStyleR}>{a}</td>
                    <td style={{...tdStyleR, color:pctColor(sp), fontWeight:600}}>{t > 0 ? sp.toFixed(0) + "%" : "—"}</td>
                    <td style={tdStyle}><Bar2 pct={sp} /></td>
                  </tr>
                );
              })}
            </React.Fragment>
          );
        })}
      </tbody>
    </table>
  );
}

function SRScorecard({ sr, onKpiClick }) {
  return (
    <div style={{
      background:"#fff", border:"1px solid #e5e7eb", borderRadius:8,
      padding:12, marginBottom:8,
    }}>
      <div style={{
        display:"flex", justifyContent:"space-between", alignItems:"center",
        marginBottom:8, paddingBottom:8, borderBottom:"1px solid #f3f4f6",
      }}>
        <div>
          <div style={{fontSize:13, fontWeight:700, color:"#111827"}}>{sr.name}</div>
          <div style={{fontSize:10, color:"#6b7280", marginTop:1}}>
            <span style={{fontFamily:"monospace"}}>{sr.code}</span> ·
            FLM <strong style={{color:"#4338ca"}}>{sr.flm}</strong>
          </div>
        </div>
        <div style={{display:"flex", gap:14, alignItems:"center"}}>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:9, color:"#9ca3af", textTransform:"uppercase"}}>Total</div>
            <div style={{fontSize:13, fontWeight:600}}>
              {fmt(sr.totalA)} <span style={{color:"#9ca3af", fontWeight:400}}>/ {fmt(sr.totalT)}</span>
            </div>
          </div>
          <div style={{
            background: heatBg(sr.totalPct), color: heatFg(sr.totalPct),
            padding:"4px 10px", borderRadius:6, fontWeight:700, fontSize:12,
            minWidth:50, textAlign:"center",
          }}>
            {sr.totalPct.toFixed(0)}%
          </div>
        </div>
      </div>

      <div style={{display:"grid", gridTemplateColumns:"repeat(7, 1fr)", gap:5, marginBottom:8}}>
        {SUB_BRANDS.map(b => {
          const sb = sr.subBrands[b.name];
          if (sb.target === 0 && sb.actual === 0) return (
            <div key={b.name} style={{
              background:"#fafafa", borderRadius:4, padding:"5px 7px",
              fontSize:9, color:"#9ca3af", textAlign:"center",
            }}>{b.name}<div>—</div></div>
          );
          return (
            <div key={b.name} style={{
              background: heatBg(sb.pct), borderRadius:4, padding:"5px 7px",
              borderLeft:"3px solid " + b.color,
            }}>
              <div style={{fontSize:9, color:b.color, fontWeight:600, textTransform:"uppercase"}}>
                {b.name}
              </div>
              <div style={{fontSize:10.5, fontWeight:600, color:"#111827", marginTop:1}}>
                {fmt(sb.actual)} <span style={{color:"#6b7280", fontSize:9}}>/ {fmt(sb.target)}</span>
              </div>
              <div style={{fontSize:9, fontWeight:700, color:heatFg(sb.pct)}}>
                {sb.pct > 0 ? sb.pct.toFixed(0) + "%" : "—"}
              </div>
            </div>
          );
        })}
      </div>

      <div style={{display:"grid", gridTemplateColumns:"repeat(10, 1fr)", gap:3, marginBottom:8}}>
        {KPI_DEFS.map(k => {
          const d = sr.kpis[k.key];
          return (
            <div key={k.key}
              onClick={() => onKpiClick(k.key)}
              style={{
                background:"#fafafa", borderRadius:4, padding:"4px 5px",
                cursor:"pointer", borderTop:"2px solid " + k.color,
              }}>
              <div style={{fontSize:9, fontWeight:600, color:"#374151"}}>{k.label}</div>
              <div style={{fontSize:10, color:"#111827", marginTop:1, fontWeight:600}}>
                {fmt(d.actual)}
              </div>
              <div style={{fontSize:8, color:"#9ca3af"}}>
                / {fmt(d.target)}
              </div>
              <div style={{fontSize:8, fontWeight:700, marginTop:1, color: pctColor(d.pct)}}>
                {d.pct > 0 ? d.pct.toFixed(0) + "%" : "—"}
              </div>
            </div>
          );
        })}
      </div>

      <div style={{
        display:"flex", gap:14, paddingTop:6,
        borderTop:"1px solid #f3f4f6", fontSize:10, color:"#6b7280", flexWrap:"wrap",
      }}>
        <div>Shop: <strong style={{color:"#111827"}}>{fmt(sr.shopActual)}</strong>
          <span style={{color:"#9ca3af"}}> / {fmt(sr.shopTarget)}</span>
          {sr.shopTarget > 0 && (
            <span style={{color:pctColor(pct(sr.shopActual, sr.shopTarget)), fontWeight:600, marginLeft:3}}>
              ({pct(sr.shopActual, sr.shopTarget).toFixed(0)}%)
            </span>
          )}
        </div>
        <div>Active: <strong style={{color:"#111827"}}>{sr.activeActual}</strong>
          <span style={{color:"#9ca3af"}}> / {sr.activeTarget}</span>
        </div>
        <div>New: <strong style={{color:"#111827"}}>{sr.newActual}</strong>
          <span style={{color:"#9ca3af"}}> / {sr.newTarget}</span>
        </div>
        <div>NU: <strong style={{color:"#111827"}}>{sr.leadActual}</strong>
          <span style={{color:"#9ca3af"}}> / {sr.leadTarget}</span>
          {sr.leadTarget > 0 && (
            <span style={{color:pctColor(pct(sr.leadActual, sr.leadTarget)), fontWeight:600, marginLeft:3}}>
              ({pct(sr.leadActual, sr.leadTarget).toFixed(0)}%)
            </span>
          )}
        </div>
      </div>
    </div>
  );
}

const selectStyle = {
  background:"#fff", color:"#111827", border:"1px solid #d1d5db",
  borderRadius:6, padding:"5px 9px", fontSize:11.5, cursor:"pointer", width:"100%",
};
const btnStyle = {
  border:"1px solid #d1d5db", borderRadius:6, padding:"5px 11px",
  fontSize:11, cursor:"pointer", fontWeight:500,
};
const tooltipStyle = {
  background:"#fff", border:"1px solid #e5e7eb", borderRadius:6, fontSize:11,
};
const tblStyle = { width:"100%", borderCollapse:"collapse", fontSize:11.5 };
const thStyle = {
  textAlign:"left", padding:"7px 9px", color:"#6b7280",
  fontWeight:600, fontSize:10, textTransform:"uppercase", letterSpacing:.5,
};
const thStyleR = { ...thStyle, textAlign:"right" };
const tdStyle = { padding:"6px 9px", color:"#111827" };
const tdStyleR = { ...tdStyle, textAlign:"right", fontVariantNumeric:"tabular-nums" };
