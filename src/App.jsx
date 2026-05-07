import React, { useState, useMemo, useEffect } from "react";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, LabelList
} from "recharts";

// === CONFIGURE THIS LINE ===
// Replace with your Apps Script Web App URL (ends in /exec)
const API_URL = "https://script.google.com/macros/s/AKfycbyXkZNFHARUMpbJ1i47BV5Dwaq1E-0cgOrn7CKIcRW8PMJuzywqH5WQIXWfQ7JKXiei/exec";

// Refresh interval for live data (seconds)
const REFRESH_INTERVAL = 60;

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

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const fmt = (n) => {
  if (n == null || isNaN(n)) return "—";
  if (n === 0) return "0";
  if (Math.abs(n) >= 1_000_000) return (n / 1_000_000).toFixed(1) + "M";
  if (Math.abs(n) >= 1_000) return (n / 1_000).toFixed(1) + "K";
  return Math.round(n).toLocaleString();
};
const pct = (a, t) => (t > 0 ? (a / t) * 100 : 0);
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
  useEffect(() => {
    if (!user) return;
    let cancelled = false;
    
    const fetchData = async () => {
      setLoading(true);
      try {
        const res = await fetch(API_URL + "?action=data");
        if (!res.ok) throw new Error("HTTP " + res.status);
        const data = await res.json();
        if (cancelled) return;
        setLoading(false);
        if (data && data.error) {
          setError(data.error);
        } else {
          setRaw(data);
          setError(null);
        }
      } catch (err) {
        if (cancelled) return;
        setLoading(false);
        setError(err.message || "Failed to load data");
      }
    };
    
    fetchData();
    const id = setInterval(fetchData, REFRESH_INTERVAL * 1000);
    return () => { cancelled = true; clearInterval(id); };
  }, [user]);

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

  return <Dashboard user={user} raw={raw} onLogout={() => setUser(null)} />;
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
        onLogin(data.user);
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

function Dashboard({ user, raw, onLogout }) {
  const RAW = raw;
  const [tick, setTick] = useState(0);
  const [auto, setAuto] = useState(true);
  const [tab, setTab] = useState("summary");

  // Filters
  const [year, setYear] = useState(2026);
  const [month, setMonth] = useState(4);
  const [flm, setFlm] = useState(user.role === "FLM" ? user.flm : "All");
  const [srFilter, setSrFilter] = useState("All");
  const [custFilter, setCustFilter] = useState("All");
  const [custSearch, setCustSearch] = useState("");

  // Expanded FLM rows
  const [expanded, setExpanded] = useState({});

  const [lastRefresh, setLastRefresh] = useState(new Date());

  useEffect(() => {
    if (!auto) return;
    const id = setInterval(() => {
      setTick(t => t + 1);
      setLastRefresh(new Date());
    }, 5000);
    return () => clearInterval(id);
  }, [auto]);

  // Reset SR filter when FLM changes
  useEffect(() => {
    if (flm === "All") setSrFilter("All");
    else if (srFilter !== "All") {
      const sr = RAW.srs.find(s => s.code === Number(srFilter));
      if (!sr || sr.flm !== flm) setSrFilter("All");
    }
  }, [flm]);

  const C = useMemo(() => {
    // === Filtered SR set ===
    let srs = RAW.srs;
    if (flm !== "All") srs = srs.filter(s => s.flm === flm);
    if (srFilter !== "All") srs = srs.filter(s => s.code === Number(srFilter));
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
      KPI_DEFS.forEach(k => {
        const t = targetMap.get(s.code + "|" + k.key) || 0;
        const a = actualMap.get(s.code + "|" + k.key) || 0;
        card.kpis[k.key] = { target: t, actual: a, pct: pct(a, t), variance: a - t };
        card.totalT += t; card.totalA += a;
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
      card.shopActual = shopItems.reduce((s2, x) => s2 + x.v, 0);
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

    // === FLM-level rollup with SR list nested ===
    const flmList = (flm === "All" ? RAW.flms : [flm]);
    const flmRollup = flmList.map(f => {
      const flmSrs = RAW.srs.filter(s => s.flm === f &&
        (srFilter === "All" || s.code === Number(srFilter)));
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
      (flm === "All" || x.f === flm) &&
      (srFilter === "All" || x.sr === Number(srFilter)) &&
      (custFilter === "All" || x.c === Number(custFilter))
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

    return {
      srs, srScorecards, kpiTotals, subBrandTotals,
      totalTarget, totalActual,
      flmRollup, flmCoverage, shopItems,
    };
  }, [tick, year, month, flm, srFilter, custFilter]);

  const overallPct = pct(C.totalActual, C.totalTarget);
  const activeTotal = (flm === "All" && srFilter === "All")
    ? RAW.activeByMonth[month]?.total
    : C.flmCoverage.reduce((s, r) => s + r.activeA, 0);
  const activeTotalT = C.flmCoverage.reduce((s, r) => s + r.activeT, 0);
  const newTotal = C.flmCoverage.reduce((s, r) => s + r.newA, 0);
  const newTotalT = C.flmCoverage.reduce((s, r) => s + r.newT, 0);
  const shopTotal = C.shopItems.reduce((s, x) => s + x.v, 0);
  const shopTotalT = C.shopItems.reduce((s, x) => s + x.t, 0);
  const leadTotal = C.flmCoverage.reduce((s, r) => s + r.leadT, 0);
  const leadActualTotal = C.flmCoverage.reduce((s, r) => s + r.leadA, 0);

  // SR options (filtered by FLM)
  const srOptions = (flm === "All" ? RAW.srs : RAW.srs.filter(s => s.flm === flm))
    .sort((a, b) => a.name.localeCompare(b.name));

  // Customer search filter
  const filteredCustomers = useMemo(() => {
    let custs = RAW.customers || [];
    if (flm !== "All") custs = custs.filter(c => c.f === flm);
    if (srFilter !== "All") custs = custs.filter(c => c.sr === Number(srFilter));
    if (custSearch) {
      const q = custSearch.toLowerCase();
      custs = custs.filter(c => String(c.c).includes(q) || (c.n || "").toLowerCase().includes(q));
    }
    return custs.slice(0, 100); // cap for performance
  }, [flm, srFilter, custSearch]);

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
          <button onClick={() => { setTick(t => t + 1); setLastRefresh(new Date()); }}
            style={{ ...btnStyle, background: "#2563eb", borderColor: "#2563eb", color: "#fff" }}>
            ↻ Refresh
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
            <option value={2026}>2026</option>
          </select>
        </FilterField>
        <FilterField label="Month">
          <select value={month} onChange={e => setMonth(Number(e.target.value))} style={selectStyle}>
            {MONTH_NAMES.map((m, i) =>
              <option key={i} value={i+1}>{m}-{String(year).slice(2)}</option>)}
          </select>
        </FilterField>
        <FilterField label="FLM">
          <select value={flm}
            onChange={e => setFlm(e.target.value)}
            disabled={user.role === "FLM"}
            style={{
              ...selectStyle,
              background: user.role === "FLM" ? "#f3f4f6" : "#fff",
              cursor: user.role === "FLM" ? "not-allowed" : "pointer",
              color: user.role === "FLM" ? "#6b7280" : "#111827",
            }}>
            {user.role === "Admin" && <option>All</option>}
            {RAW.flms
              .filter(f => user.role === "Admin" || f === user.flm)
              .map(f => <option key={f}>{f}</option>)}
          </select>
        </FilterField>
        <FilterField label="SR">
          <select value={srFilter} onChange={e => setSrFilter(e.target.value)} style={selectStyle}>
            <option value="All">All</option>
            {srOptions.map(s =>
              <option key={s.code} value={s.code}>{s.name}</option>)}
          </select>
        </FilterField>
        <FilterField label="Customer">
          <select value={custFilter} onChange={e => setCustFilter(e.target.value)} style={selectStyle}>
            <option value="All">All ({filteredCustomers.length})</option>
            {filteredCustomers.map(c =>
              <option key={c.c} value={c.c}>{c.c} — {(c.n || "").substring(0, 30)}</option>)}
          </select>
        </FilterField>
      </div>

      {/* === SUMMARY METRICS — 6 cards === */}
      <div style={{display:"grid", gridTemplateColumns:"repeat(6, 1fr)", gap:8, marginBottom:12}}>
        <Metric label="Sales · Actual / Target" value={fmt(C.totalActual)}
          actual={fmt(C.totalTarget)} pct={overallPct} accent="#2563eb"
          secondaryLabel="tgt" />
        <Metric label="Active 3-mo · Actual / Target" value={fmt(activeTotal)}
          actual={fmt(activeTotalT)} pct={pct(activeTotal, activeTotalT)} accent="#8b5cf6"
          secondaryLabel="tgt" />
        <Metric label="Shop Around · Actual / Target" value={fmt(shopTotal)}
          actual={fmt(shopTotalT)} pct={pct(shopTotal, shopTotalT)} accent="#0ea5e9"
          secondaryLabel="tgt" />
        <Metric label="New Listing · Actual / Target" value={fmt(newTotal)}
          actual={fmt(newTotalT)} pct={pct(newTotal, newTotalT)} accent="#ec4899"
          secondaryLabel="tgt" />
        <Metric label="NU · Actual / Target" value={fmt(leadActualTotal)}
          actual={fmt(leadTotal)} pct={pct(leadActualTotal, leadTotal)} accent="#f59e0b"
          secondaryLabel="tgt" />
        <Metric label="Variance" value={fmt(C.totalActual - C.totalTarget)}
          actual={overallPct.toFixed(1) + "%"} pct={null}
          accent={C.totalActual >= C.totalTarget ? "#059669" : "#dc2626"}
          sub={C.totalActual >= C.totalTarget ? "Above target" : "Below target"} />
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
        <TabBtn label="New Listing" v="new" cur={tab} on={setTab} />
        <TabBtn label="NU" v="leads" cur={tab} on={setTab} />
          <TabBtn label="📈 Trend" v="trend" cur={tab} on={setTab} />
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

          <Panel title="FLM × KPI Matrix (click ▸ to expand SR detail)">
            <div style={{overflowX:"auto"}}>
              <table style={tblStyle}>
                <thead><tr style={{background:"#f9fafb"}}>
                  <th style={thStyle}>FLM / SR</th>
                  <th style={thStyleR}>Target</th>
                  <th style={thStyleR}>Actual</th>
                  <th style={thStyleR}>Var</th>
                  <th style={thStyleR}>%</th>
                  <th style={thStyle}>Progress</th>
                </tr></thead>
                <tbody>
                  {C.flmRollup.map(f => (
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
                        <td style={tdStyle}><Bar2 pct={f.pct} /></td>
                      </tr>
                      {expanded[f.flm] && f.srs.map(sr => {
                        const card = C.srScorecards.find(c => c.code === sr.code);
                        if (!card) return null;
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
                            <td style={tdStyle}><Bar2 pct={card.totalPct} /></td>
                          </tr>
                        );
                      })}
                    </React.Fragment>
                  ))}
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
                    <td style={tdStyle}><Bar2 pct={overallPct} /></td>
                  </tr>
                </tbody>
              </table>
            </div>
          </Panel>

          <Panel title="FLM Coverage Metrics (Shop Around · Active 3-mo · New Listing · NU)">
            <div style={{overflowX:"auto"}}>
              <table style={tblStyle}>
                <thead><tr style={{background:"#f9fafb"}}>
                  <th style={thStyle}>FLM</th>
                  <th style={thStyleR}>Shop Tgt</th>
                  <th style={thStyleR}>Shop Act</th>
                  <th style={thStyleR}>%</th>
                  <th style={{...thStyleR, borderLeft:"1px solid #e5e7eb"}}>Active Tgt</th>
                  <th style={thStyleR}>Active Act</th>
                  <th style={thStyleR}>%</th>
                  <th style={{...thStyleR, borderLeft:"1px solid #e5e7eb"}}>New Tgt</th>
                  <th style={thStyleR}>New Act</th>
                  <th style={thStyleR}>%</th>
                  <th style={{...thStyleR, borderLeft:"1px solid #e5e7eb"}}>NU Tgt</th>
                  <th style={thStyleR}>NU Act</th>
                  <th style={thStyleR}>%</th>
                </tr></thead>
                <tbody>
                  {C.flmCoverage.map(r => {
                    const sP = pct(r.shopA, r.shopT);
                    const aP = pct(r.activeA, r.activeT);
                    const nP = pct(r.newA, r.newT);
                    const lP = pct(r.leadA, r.leadT);
                    return (
                      <tr key={r.flm} style={{borderTop:"1px solid #f3f4f6"}}>
                        <td style={tdStyle}><strong>{r.flm}</strong></td>
                        <td style={tdStyleR}>{fmt(r.shopT)}</td>
                        <td style={tdStyleR}>{fmt(r.shopA)}</td>
                        <td style={{...tdStyleR, color:pctColor(sP), fontWeight:600}}>
                          {r.shopT > 0 ? sP.toFixed(0) + "%" : "—"}
                        </td>
                        <td style={{...tdStyleR, borderLeft:"1px solid #e5e7eb"}}>{r.activeT}</td>
                        <td style={tdStyleR}>{r.activeA}</td>
                        <td style={{...tdStyleR, color:pctColor(aP), fontWeight:600}}>
                          {r.activeT > 0 ? aP.toFixed(0) + "%" : "—"}
                        </td>
                        <td style={{...tdStyleR, borderLeft:"1px solid #e5e7eb"}}>{r.newT}</td>
                        <td style={tdStyleR}>{r.newA}</td>
                        <td style={{...tdStyleR, color:pctColor(nP), fontWeight:600}}>
                          {r.newT > 0 ? nP.toFixed(0) + "%" : "—"}
                        </td>
                        <td style={{...tdStyleR, borderLeft:"1px solid #e5e7eb"}}>{r.leadT}</td>
                        <td style={tdStyleR}>{r.leadA}</td>
                        <td style={{...tdStyleR, color:pctColor(lP), fontWeight:600}}>
                          {r.leadT > 0 ? lP.toFixed(0) + "%" : "—"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </Panel>
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
            <Panel title={"📊 SR Performance Table — All KPIs · " + MONTH_NAMES[month-1] + "-" + String(year).slice(2)}>
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
          .filter(r => flm === "All" || (() => {
            const sr = RAW.srs.find(s => s.code === r.sr);
            return sr && sr.flm === flm;
          })())
          .filter(r => srFilter === "All" || r.sr === Number(srFilter))
          .filter(r => custFilter === "All" || r.c === Number(custFilter))
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

            <Panel title={k.label + " — FLM × SR Matrix (expandable)"}>
              {(flm === "All" ? RAW.flms : [flm]).map(f => {
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

            <Panel title={k.label + " — Customer Details (" + custDetails.length + " customers)"}>
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
        const totT = items.reduce((s, x) => s + x.t, 0);
        const totA = items.reduce((s, x) => s + x.v, 0);
        const flmList = (flm === "All" ? RAW.flms : [flm]);

        // Per-FLM and per-SR rollups
        const flmRows = flmList.map(f => {
          const fis = items.filter(x => x.f === f);
          return {
            flm: f, srs: [...new Set(fis.map(x => x.sr))],
            target: fis.reduce((s, x) => s + x.t, 0),
            actual: fis.reduce((s, x) => s + x.v, 0),
            customers: fis,
          };
        }).filter(r => r.target > 0 || r.actual > 0);

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

            <Panel title="Shop Around — FLM Matrix (expand to see SRs and customers)">
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
                          {pct(f.actual, f.target).toFixed(1)}%
                        </td>
                        <td style={tdStyle}><Bar2 pct={pct(f.actual, f.target)} /></td>
                      </tr>
                      {expanded["s_"+f.flm] && f.srs.map(srCode => {
                        const sr = RAW.srs.find(s => s.code === srCode);
                        const srItems = f.customers.filter(x => x.sr === srCode);
                        const srT = srItems.reduce((s, x) => s + x.t, 0);
                        const srA = srItems.reduce((s, x) => s + x.v, 0);
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
                                {pct(srA, srT).toFixed(1)}%
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
                                <td style={tdStyleR}>{fmt(c.v)}</td>
                                <td style={{...tdStyleR, color:pctColor(pct(c.v, c.t)), fontWeight:600}}>
                                  {c.t > 0 ? pct(c.v, c.t).toFixed(0) + "%" : "—"}
                                </td>
                                <td style={tdStyle}><Bar2 pct={pct(c.v, c.t)} /></td>
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
        const flmList = (flm === "All" ? RAW.flms : [flm]);
        const flmRows = flmList.map(f => ({
          flm: f, target: at.byFlm?.[f] || 0, actual: am.byFlm?.[f] || 0,
          srs: RAW.srs.filter(s => s.flm === f),
        }));
        const totT = flmRows.reduce((s, r) => s + r.target, 0);
        const totA = (flm === "All" && srFilter === "All" && custFilter === "All") ? am.total
          : flmRows.reduce((s, r) => s + r.actual, 0);
        const unassigned = am.byFlm?.Unassigned || 0;

        return (
          <>
            <div style={{display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8, marginBottom:10}}>
              <Metric label="Active 3-mo · Actual / Target" value={fmt(totA)}
                actual={fmt(totT)} pct={pct(totA, totT)} accent="#8b5cf6"
                secondaryLabel="tgt" />
              <Metric label="Logic" value="3-mo rolling"
                actual={MONTH_NAMES[Math.max(0, month-3)] + " + " + MONTH_NAMES[Math.max(0, month-2)] + " + " + MONTH_NAMES[month-1]} pct={null} accent="#0ea5e9" />
              <Metric label="Unassigned" value={String(unassigned)}
                actual="customers without FLM" pct={null} accent="#6b7280" />
            </div>

            <Panel title="Active Customers — FLM × SR Matrix (expandable)">
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
                          <td style={tdStyleR}>{f.target}</td>
                          <td style={tdStyleR}>{f.actual}</td>
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
                              <td style={tdStyleR}>{tt}</td>
                              <td style={tdStyleR}>{aa}</td>
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
                  {unassigned > 0 && flm === "All" && (
                    <tr style={{borderTop:"1px solid #f3f4f6", color:"#9ca3af", fontStyle:"italic"}}>
                      <td style={tdStyle}>Unassigned (no FLM)</td>
                      <td style={tdStyleR}>—</td>
                      <td style={tdStyleR}>{unassigned}</td>
                      <td style={tdStyleR}>—</td>
                      <td style={tdStyle}>—</td>
                    </tr>
                  )}
                  <tr style={{borderTop:"2px solid #d1d5db", background:"#f9fafb", fontWeight:700}}>
                    <td style={tdStyle}>TOTAL</td>
                    <td style={tdStyleR}>{totT}</td>
                    <td style={tdStyleR}>{totA + (flm === "All" ? unassigned : 0)}</td>
                    <td style={{...tdStyleR, color:pctColor(pct(totA + (flm === "All" ? unassigned : 0), totT))}}>
                      {pct(totA + (flm === "All" ? unassigned : 0), totT).toFixed(0)}%
                    </td>
                    <td style={tdStyle}><Bar2 pct={pct(totA + (flm === "All" ? unassigned : 0), totT)} /></td>
                  </tr>
                </tbody>
              </table>
            </Panel>

            <Panel title={"Active Customer List (" + (am.customers || []).length + " customers active in 3-mo window)"}>
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
                      .filter(c => flm === "All" || c.f === flm)
                      .filter(c => srFilter === "All" || (c.sr && c.sr.code === Number(srFilter)))
                      .filter(c => custFilter === "All" || c.c === Number(custFilter))
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
          </>
        );
      })()}

      {/* ============ NEW LISTING ============ */}
      {tab === "new" && (() => {
        const nm = RAW.newByMonth[month] || {};
        const nt = RAW.newTargetByMonth[month] || {};
        const flmList = (flm === "All" ? RAW.flms : [flm]);
        const flmRows = flmList.map(f => ({
          flm: f, target: nt.byFlm?.[f] || 0, actual: nm.byFlm?.[f] || 0,
          srs: RAW.srs.filter(s => s.flm === f),
        }));
        const totT = flmRows.reduce((s, r) => s + r.target, 0);
        const totA = flmRows.reduce((s, r) => s + r.actual, 0);

        const items = (nm.items || [])
          .filter(p => flm === "All" || p.f === flm)
          .filter(p => srFilter === "All" || p.sr === Number(srFilter))
          .filter(p => custFilter === "All" || p.c === Number(custFilter));

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

            <Panel title="New Listing — FLM × SR Matrix (expandable)">
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
                          <td style={tdStyleR}>{f.target}</td>
                          <td style={tdStyleR}>{f.actual}</td>
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
                              <td style={tdStyleR}>{tt}</td>
                              <td style={tdStyleR}>{aa}</td>
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

            <Panel title={"New Listing items (" + items.length + ")"}>
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
            <Panel title="NU — FLM × SR (expandable)">
              <table style={tblStyle}>
                <thead><tr style={{background:"#f9fafb"}}>
                  <th style={thStyle}>FLM / SR</th>
                  <th style={thStyleR}>Target</th>
                  <th style={thStyleR}>Actual</th>
                  <th style={thStyleR}>%</th>
                  <th style={thStyle}>Progress</th>
                </tr></thead>
                <tbody>
                  {(flm === "All" ? RAW.flms : [flm]).map(f => {
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

      {tab === "trend" && (() => {
        // Compute monthly data per KPI
        // KPIs to show: Sales (total $), Active 3-mo (count), Shop Around ($), New Listing (count), NU (count)
        const TREND_KPIS = [
          { id: "sales",   label: "Total Sales",      color: "#2563eb", fmt: (v) => fmt(v) },
          { id: "active",  label: "Active 3-mo Cust", color: "#8b5cf6", fmt: (v) => v.toString() },
          { id: "shop",    label: "Shop Around",      color: "#0ea5e9", fmt: (v) => fmt(v) },
          { id: "new",     label: "New Listing",      color: "#ec4899", fmt: (v) => v.toString() },
          { id: "nu",      label: "NU",               color: "#f59e0b", fmt: (v) => v.toString() },
        ];

        // For each month, compute actual + target for each KPI
        // Also break down per FLM and per SR
        const flmList = flm === "All" ? RAW.flms : [flm];
        const srPool = sr === "All" 
          ? RAW.srs.filter(s => flmList.includes(s.flm))
          : RAW.srs.filter(s => s.code === Number(sr));

        // Build data structures for each KPI
        const compute = (kpiId, level) => {
          // level: "topline" | "flm" | "sr"
          // returns: array of {month, monthLabel, [seriesName]: actual, [seriesName + " tgt"]: target}
          const data = MONTH_NAMES.map((mLabel, idx) => {
            const m = idx + 1;
            const row = { month: mLabel };
            
            const computeForScope = (srCodes, flmNames) => {
              let actual = 0, target = 0;
              
              if (kpiId === "sales") {
                // Sum all KPI sales for these SRs/FLMs in this month
                RAW.actuals.forEach(r => {
                  if (r.m !== m) return;
                  if (srCodes && !srCodes.includes(r.sr)) return;
                  if (flmNames && !flmNames.includes(r.f)) return;
                  actual += r.v;
                });
                RAW.targets.forEach(r => {
                  if (r.m !== m) return;
                  if (srCodes && !srCodes.includes(r.sr)) return;
                  if (flmNames && !flmNames.includes(r.f)) return;
                  target += r.t;
                });
              } else if (kpiId === "active") {
                const ab = RAW.activeByMonth?.[m] || [];
                const at = RAW.activeTargetByMonth?.[m];
                ab.forEach(r => {
                  if (srCodes && !srCodes.includes(r.sr)) return;
                  if (flmNames && !flmNames.includes(r.f)) return;
                  actual += 1;
                });
                if (at) {
                  if (srCodes) {
                    srCodes.forEach(s => target += (at.bySr[s] || 0));
                  } else if (flmNames) {
                    flmNames.forEach(f => target += (at.byFlm[f] || 0));
                  } else {
                    Object.values(at.bySr).forEach(v => target += v);
                  }
                }
              } else if (kpiId === "shop") {
                const sb = RAW.shopByMonth?.[m] || [];
                sb.forEach(r => {
                  if (srCodes && !srCodes.includes(r.sr)) return;
                  if (flmNames && !flmNames.includes(r.f)) return;
                  actual += r.v || 0;
                  target += r.t || 0;
                });
              } else if (kpiId === "new") {
                const nb = RAW.newByMonth?.[m] || [];
                const nt = RAW.newTargetByMonth?.[m];
                nb.forEach(r => {
                  if (srCodes && !srCodes.includes(r.sr)) return;
                  if (flmNames && !flmNames.includes(r.f)) return;
                  actual += 1;
                });
                if (nt) {
                  if (srCodes) {
                    srCodes.forEach(s => target += (nt.bySr[s] || 0));
                  } else if (flmNames) {
                    flmNames.forEach(f => target += (nt.byFlm[f] || 0));
                  } else {
                    Object.values(nt.bySr).forEach(v => target += v);
                  }
                }
              } else if (kpiId === "nu") {
                const lt = RAW.leadTargetByMonth?.[m];
                const la = RAW.leadActualByMonth?.[m];
                if (lt) {
                  if (srCodes) {
                    srCodes.forEach(s => target += (lt.bySr[s] || 0));
                  } else if (flmNames) {
                    flmNames.forEach(f => target += (lt.byFlm[f] || 0));
                  } else {
                    Object.values(lt.bySr).forEach(v => target += v);
                  }
                }
                if (la) {
                  if (srCodes) {
                    srCodes.forEach(s => actual += (la.bySr[s] || 0));
                  } else if (flmNames) {
                    flmNames.forEach(f => actual += (la.byFlm[f] || 0));
                  } else {
                    Object.values(la.bySr).forEach(v => actual += v);
                  }
                }
              }
              return { actual, target };
            };

            if (level === "topline") {
              const { actual, target } = computeForScope(null, flmList);
              row["Actual"] = actual;
              row["Target"] = target;
            } else if (level === "flm") {
              flmList.forEach(f => {
                const { actual, target } = computeForScope(null, [f]);
                row[f] = actual;
                row[f + " tgt"] = target;
              });
            } else if (level === "sr") {
              srPool.forEach(s => {
                const { actual } = computeForScope([s.code], null);
                row[s.name] = actual;
              });
            }
            return row;
          });
          return data;
        };

        // Compute trend direction (last 3 months % change)
        const trendDir = (data, key) => {
          const vals = data.map(d => d[key] || 0).filter((_, i, a) => i >= a.length - 3);
          if (vals.length < 2) return null;
          const first = vals[0];
          const last = vals[vals.length - 1];
          if (first === 0) return last > 0 ? "↗" : "—";
          const change = (last - first) / first;
          if (change > 0.05) return "↗ improving";
          if (change < -0.05) return "↘ declining";
          return "→ stable";
        };

        return (
          <>
            <Panel title="📈 Trend Analysis — Year-to-Date 2026">
              <div style={{fontSize:11, color:"#6b7280", marginBottom:10}}>
                Showing 5 KPIs across 3 levels (Topline → FLM → SR). Filters above apply.
                Number labels shown on each data point. Months with no data omitted.
              </div>
              
              {TREND_KPIS.map(kpi => {
                const toplineData = compute(kpi.id, "topline");
                const flmData = compute(kpi.id, "flm");
                
                // Filter out months with both 0 actual and 0 target
                const hasData = toplineData.some(d => (d.Actual || 0) > 0 || (d.Target || 0) > 0);
                if (!hasData) return null;
                
                const trend = trendDir(toplineData, "Actual");
                const trendColor = trend?.includes("improving") ? "#059669" 
                  : trend?.includes("declining") ? "#dc2626" : "#6b7280";
                
                return (
                  <div key={kpi.id} style={{marginBottom:24, padding:12, border:"1px solid #e5e7eb", borderRadius:8, background:"#fff"}}>
                    <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10}}>
                      <div>
                        <div style={{fontSize:14, fontWeight:700, color:kpi.color}}>{kpi.label}</div>
                        <div style={{fontSize:10, color:"#9ca3af"}}>Topline · Actual vs Target by month</div>
                      </div>
                      {trend && (
                        <div style={{fontSize:12, fontWeight:600, color:trendColor}}>
                          {trend}
                        </div>
                      )}
                    </div>
                    
                    {/* TOPLINE chart */}
                    <ResponsiveContainer width="100%" height={220}>
                      <BarChart data={toplineData} margin={{top:20, right:20, left:0, bottom:5}}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#f3f4f6" />
                        <XAxis dataKey="month" tick={{fontSize:10, fill:"#6b7280"}} />
                        <YAxis tick={{fontSize:10, fill:"#6b7280"}} tickFormatter={(v) => v >= 1000 ? (v/1000).toFixed(0)+"K" : v} />
                        <Tooltip />
                        <Legend wrapperStyle={{fontSize:11}} />
                        <Bar dataKey="Target" fill="#e5e7eb" name="Target">
                          <LabelList dataKey="Target" position="top" style={{fontSize:9, fill:"#9ca3af"}} 
                            formatter={(v) => v ? (v >= 1000 ? (v/1000).toFixed(0)+"K" : v) : ""} />
                        </Bar>
                        <Bar dataKey="Actual" fill={kpi.color} name="Actual">
                          <LabelList dataKey="Actual" position="top" style={{fontSize:9, fill:kpi.color, fontWeight:700}} 
                            formatter={(v) => v ? (v >= 1000 ? (v/1000).toFixed(0)+"K" : v) : ""} />
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>

                    {/* FLM breakdown chart */}
                    {flm === "All" && flmList.length > 1 && (
                      <>
                        <div style={{fontSize:11, fontWeight:600, color:"#374151", marginTop:14, marginBottom:6}}>
                          Per FLM — Actual only
                        </div>
                        <ResponsiveContainer width="100%" height={200}>
                          <LineChart data={flmData} margin={{top:15, right:20, left:0, bottom:5}}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#f3f4f6" />
                            <XAxis dataKey="month" tick={{fontSize:10, fill:"#6b7280"}} />
                            <YAxis tick={{fontSize:10, fill:"#6b7280"}} tickFormatter={(v) => v >= 1000 ? (v/1000).toFixed(0)+"K" : v} />
                            <Tooltip />
                            <Legend wrapperStyle={{fontSize:10}} />
                            {flmList.map((f, i) => {
                              const colors = ["#2563eb","#dc2626","#059669","#f59e0b","#8b5cf6"];
                              return (
                                <Line key={f} type="monotone" dataKey={f} stroke={colors[i % 5]} strokeWidth={2}
                                  dot={{r:3}}>
                                  <LabelList dataKey={f} position="top" style={{fontSize:8, fill:colors[i % 5], fontWeight:600}} 
                                    formatter={(v) => v ? (v >= 1000 ? (v/1000).toFixed(0)+"K" : v) : ""} />
                                </Line>
                              );
                            })}
                          </LineChart>
                        </ResponsiveContainer>
                      </>
                    )}

                    {/* Summary table */}
                    <div style={{marginTop:14, fontSize:11, fontWeight:600, color:"#374151", marginBottom:6}}>
                      Monthly Summary
                    </div>
                    <table style={{...tblStyle, fontSize:10}}>
                      <thead><tr style={{background:"#f9fafb"}}>
                        <th style={thStyle}>Month</th>
                        <th style={thStyleR}>Actual</th>
                        <th style={thStyleR}>Target</th>
                        <th style={thStyleR}>%</th>
                        <th style={thStyleR}>vs prior</th>
                      </tr></thead>
                      <tbody>
                        {toplineData.map((d, i) => {
                          const a = d.Actual || 0;
                          const t = d.Target || 0;
                          if (a === 0 && t === 0) return null;
                          const p = pct(a, t);
                          const prior = i > 0 ? (toplineData[i-1].Actual || 0) : 0;
                          const change = prior > 0 ? ((a - prior) / prior * 100) : null;
                          return (
                            <tr key={d.month} style={{borderTop:"1px solid #f3f4f6"}}>
                              <td style={tdStyle}>{d.month} 2026</td>
                              <td style={{...tdStyleR, fontWeight:600}}>{kpi.fmt(a)}</td>
                              <td style={{...tdStyleR, color:"#6b7280"}}>{t > 0 ? kpi.fmt(t) : "—"}</td>
                              <td style={{...tdStyleR, fontWeight:700, color:t > 0 ? pctColor(p) : "#9ca3af"}}>
                                {t > 0 ? p.toFixed(0) + "%" : "—"}
                              </td>
                              <td style={{...tdStyleR, fontSize:10, color: change > 0 ? "#059669" : change < 0 ? "#dc2626" : "#6b7280"}}>
                                {change !== null ? (change > 0 ? "+" : "") + change.toFixed(0) + "%" : "—"}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                );
              })}
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

function Metric({ label, value, actual, pct, accent, sub, secondaryLabel }) {
  return (
    <div style={{
      background:"#fff", border:"1px solid #e5e7eb",
      borderTop:"3px solid " + accent, borderRadius:8, padding:"10px 12px",
    }}>
      <div style={{fontSize:9, color:"#6b7280", textTransform:"uppercase", letterSpacing:.5, fontWeight:600}}>{label}</div>
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

function Panel({ title, children }) {
  return (
    <div style={{
      background:"#fff", border:"1px solid #e5e7eb", borderRadius:8,
      padding:12, marginBottom:10,
    }}>
      <div style={{fontSize:12, fontWeight:600, color:"#111827", marginBottom:10}}>{title}</div>
      {children}
    </div>
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
