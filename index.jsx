import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  AreaChart, Area, ScatterChart, Scatter,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer
} from "recharts";

// ── Palette ──────────────────────────────────────────────────────────────────
const P = {
  bg: "#0a0d14",
  surface: "#111520",
  card: "#151b28",
  border: "#1e2a3d",
  accent: "#00d4ff",
  accent2: "#7c3aed",
  accent3: "#10b981",
  warn: "#f59e0b",
  danger: "#ef4444",
  text: "#e2e8f0",
  muted: "#64748b",
  chartColors: ["#00d4ff","#7c3aed","#10b981","#f59e0b","#ef4444","#ec4899","#06b6d4","#84cc16"],
};

// ── Helpers ───────────────────────────────────────────────────────────────────
function isNumericColumn(data, key) {
  const vals = data.slice(0, 20).map(r => r[key]).filter(v => v !== null && v !== undefined && v !== "");
  if (!vals.length) return false;
  return vals.filter(v => !isNaN(Number(v))).length / vals.length > 0.7;
}

function getColumnStats(data, key) {
  const vals = data.map(r => Number(r[key])).filter(v => !isNaN(v));
  if (!vals.length) return null;
  const sum = vals.reduce((a, b) => a + b, 0);
  const mean = sum / vals.length;
  const sorted = [...vals].sort((a, b) => a - b);
  const median = sorted[Math.floor(sorted.length / 2)];
  const min = sorted[0];
  const max = sorted[sorted.length - 1];
  const variance = vals.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / vals.length;
  const std = Math.sqrt(variance);
  return { sum, mean, median, min, max, std, count: vals.length };
}

function getCategoryFrequency(data, key) {
  const freq = {};
  data.forEach(r => {
    const v = String(r[key] ?? "");
    if (v && v !== "undefined") freq[v] = (freq[v] || 0) + 1;
  });
  return Object.entries(freq)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 12)
    .map(([name, value]) => ({ name: name.length > 14 ? name.slice(0, 13) + "…" : name, value }));
}

function fmt(n) {
  if (n === null || n === undefined) return "—";
  const num = Number(n);
  if (isNaN(num)) return String(n);
  if (Math.abs(num) >= 1e9) return (num / 1e9).toFixed(2) + "B";
  if (Math.abs(num) >= 1e6) return (num / 1e6).toFixed(2) + "M";
  if (Math.abs(num) >= 1e3) return (num / 1e3).toFixed(1) + "K";
  return num % 1 === 0 ? num.toLocaleString() : num.toFixed(2);
}

// ── Components ────────────────────────────────────────────────────────────────
const GlowDot = ({ color = P.accent }) => (
  <span style={{
    display: "inline-block", width: 8, height: 8, borderRadius: "50%",
    background: color, boxShadow: `0 0 8px ${color}`, marginRight: 6,
  }} />
);

const Badge = ({ children, color = P.accent }) => (
  <span style={{
    display: "inline-flex", alignItems: "center", gap: 4,
    padding: "2px 10px", borderRadius: 99, fontSize: 11, fontWeight: 600,
    background: color + "22", color, border: `1px solid ${color}44`,
    fontFamily: "'DM Mono', monospace", letterSpacing: "0.04em",
  }}>{children}</span>
);

const StatCard = ({ label, value, sub, color = P.accent, icon }) => (
  <div style={{
    background: P.card, border: `1px solid ${P.border}`, borderRadius: 16,
    padding: "20px 22px", position: "relative", overflow: "hidden",
    transition: "transform 0.2s, box-shadow 0.2s",
  }}
    onMouseEnter={e => { e.currentTarget.style.transform = "translateY(-2px)"; e.currentTarget.style.boxShadow = `0 8px 32px ${color}22`; }}
    onMouseLeave={e => { e.currentTarget.style.transform = ""; e.currentTarget.style.boxShadow = ""; }}
  >
    <div style={{ position: "absolute", top: 0, right: 0, width: 80, height: 80, borderRadius: "0 16px 0 80px", background: color + "0d" }} />
    <div style={{ fontSize: 24, marginBottom: 4 }}>{icon}</div>
    <div style={{ fontSize: 13, color: P.muted, marginBottom: 4, fontFamily: "'DM Mono', monospace" }}>{label}</div>
    <div style={{ fontSize: 28, fontWeight: 700, color, fontFamily: "'Space Grotesk', sans-serif", lineHeight: 1.1 }}>{value}</div>
    {sub && <div style={{ fontSize: 11, color: P.muted, marginTop: 4 }}>{sub}</div>}
  </div>
);

const ChartCard = ({ title, subtitle, children, span = 1 }) => (
  <div style={{
    background: P.card, border: `1px solid ${P.border}`, borderRadius: 16,
    padding: "20px 22px", gridColumn: span === 2 ? "span 2" : undefined,
  }}>
    <div style={{ marginBottom: 4, fontWeight: 700, color: P.text, fontSize: 15, fontFamily: "'Space Grotesk', sans-serif" }}>{title}</div>
    {subtitle && <div style={{ fontSize: 12, color: P.muted, marginBottom: 16 }}>{subtitle}</div>}
    <div style={{ marginTop: subtitle ? 0 : 12 }}>{children}</div>
  </div>
);

const InsightCard = ({ text, color = P.accent }) => (
  <div style={{
    display: "flex", alignItems: "flex-start", gap: 10,
    background: color + "0d", border: `1px solid ${color}33`,
    borderRadius: 12, padding: "12px 14px", fontSize: 13, color: P.text, lineHeight: 1.6,
  }}>
    <span style={{ color, fontSize: 16, marginTop: 1 }}>◆</span>
    <span>{text}</span>
  </div>
);

const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{
      background: P.surface, border: `1px solid ${P.border}`, borderRadius: 10,
      padding: "10px 14px", fontSize: 12, color: P.text,
    }}>
      <div style={{ color: P.muted, marginBottom: 4 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || P.accent }}>
          {p.name}: <strong>{fmt(p.value)}</strong>
        </div>
      ))}
    </div>
  );
};

// ── AI Insights Generator ─────────────────────────────────────────────────────
async function generateInsights(columns, numericStats, categoryCols, totalRows) {
  const summary = {
    totalRows,
    numericColumns: numericStats.map(s => ({
      name: s.key,
      mean: s.stats.mean.toFixed(2),
      min: s.stats.min,
      max: s.stats.max,
      std: s.stats.std.toFixed(2),
    })),
    categoricalColumns: categoryCols.map(c => ({
      name: c,
      topValues: "varies",
    })),
  };

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      system: `You are a data analyst. Given dataset statistics, return ONLY a JSON array of exactly 6 insight strings. No markdown, no backticks, no preamble. Each insight should be 1-2 sentences, actionable, and specific to the data. Format: ["insight1","insight2","insight3","insight4","insight5","insight6"]`,
      messages: [{ role: "user", content: `Analyze this dataset and give 6 key insights:\n${JSON.stringify(summary, null, 2)}` }],
    }),
  });
  const data = await res.json();
  const text = data.content?.find(b => b.type === "text")?.text || "[]";
  try {
    return JSON.parse(text.replace(/```json|```/g, "").trim());
  } catch {
    return ["Dataset loaded successfully.", "Explore the charts below for patterns.", "Check numeric distributions for outliers.", "Category breakdowns reveal frequency patterns.", "Use the table view for detailed row inspection.", "Scroll through charts for full analysis."];
  }
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function ExcelDashboard() {
  const [stage, setStage] = useState("upload"); // upload | analyzing | dashboard
  const [dragOver, setDragOver] = useState(false);
  const [fileName, setFileName] = useState("");
  const [rawData, setRawData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [numericCols, setNumericCols] = useState([]);
  const [catCols, setCatCols] = useState([]);
  const [insights, setInsights] = useState([]);
  const [activeTab, setActiveTab] = useState("overview");
  const [progress, setProgress] = useState(0);
  const [error, setError] = useState("");
  const fileRef = useRef();

  const processFile = useCallback(async (file) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx", "xls", "csv"].includes(ext)) {
      setError("Please upload an Excel (.xlsx, .xls) or CSV file.");
      return;
    }
    setError("");
    setFileName(file.name);
    setStage("analyzing");
    setProgress(10);

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        setProgress(30);
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        setProgress(55);

        if (!json.length) { setError("Sheet appears empty."); setStage("upload"); return; }
        const cols = Object.keys(json[0]);
        const numeric = cols.filter(c => isNumericColumn(json, c));
        const cat = cols.filter(c => !numeric.includes(c));

        setRawData(json);
        setColumns(cols);
        setNumericCols(numeric);
        setCatCols(cat);
        setProgress(75);

        const numStats = numeric.map(k => ({ key: k, stats: getColumnStats(json, k) })).filter(x => x.stats);
        const ins = await generateInsights(cols, numStats, cat, json.length);
        setInsights(ins);
        setProgress(100);
        setTimeout(() => setStage("dashboard"), 400);
      } catch (err) {
        setError("Could not parse file. Make sure it's a valid Excel/CSV.");
        setStage("upload");
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault(); setDragOver(false);
    processFile(e.dataTransfer.files[0]);
  }, [processFile]);

  // ── Computed chart data ────────────────────────────────────────────────────
  const numericStats = numericCols.map(k => ({ key: k, stats: getColumnStats(rawData, k) })).filter(x => x.stats);
  const catFreqs = catCols.slice(0, 4).map(c => ({ col: c, data: getCategoryFrequency(rawData, c) }));

  // line chart: first numeric col over row index (sample 60 pts)
  const lineData = (() => {
    if (!numericCols[0]) return [];
    const step = Math.max(1, Math.floor(rawData.length / 60));
    return rawData.filter((_, i) => i % step === 0).map((r, i) => ({
      index: i + 1,
      ...Object.fromEntries(numericCols.slice(0, 3).map(c => [c, Number(r[c]) || 0])),
    }));
  })();

  // scatter: first two numeric cols
  const scatterData = (() => {
    if (numericCols.length < 2) return [];
    const step = Math.max(1, Math.floor(rawData.length / 200));
    return rawData.filter((_, i) => i % step === 0).map(r => ({
      x: Number(r[numericCols[0]]) || 0,
      y: Number(r[numericCols[1]]) || 0,
    }));
  })();

  // ── STYLES ─────────────────────────────────────────────────────────────────
  const gs = {
    fontFamily: "'Space Grotesk', 'DM Mono', sans-serif",
    background: P.bg, minHeight: "100vh", color: P.text,
    padding: stage === "dashboard" ? "0" : "0",
  };

  // ── UPLOAD SCREEN ──────────────────────────────────────────────────────────
  if (stage === "upload") return (
    <div style={{ ...gs, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "100vh" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: ${P.surface}; }
        ::-webkit-scrollbar-thumb { background: ${P.border}; border-radius: 3px; }
        .tab-btn:hover { background: ${P.border} !important; }
        .upload-zone:hover { border-color: ${P.accent} !important; background: ${P.accent}08 !important; }
      `}</style>

      {/* Background grid */}
      <div style={{ position: "fixed", inset: 0, backgroundImage: `linear-gradient(${P.border} 1px, transparent 1px), linear-gradient(90deg, ${P.border} 1px, transparent 1px)`, backgroundSize: "40px 40px", opacity: 0.3, pointerEvents: "none" }} />

      <div style={{ position: "relative", textAlign: "center", maxWidth: 560, width: "90%", padding: "0 20px" }}>
        {/* Logo */}
        <div style={{ marginBottom: 8 }}>
          <span style={{ fontSize: 48, filter: `drop-shadow(0 0 20px ${P.accent})` }}>⬡</span>
        </div>
        <div style={{ fontSize: 13, letterSpacing: "0.25em", color: P.accent, fontFamily: "'DM Mono', monospace", marginBottom: 10, textTransform: "uppercase" }}>AI Dataset Intelligence</div>
        <h1 style={{ fontSize: 42, fontWeight: 700, margin: "0 0 12px", lineHeight: 1.1, background: `linear-gradient(135deg, ${P.text}, ${P.accent})`, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>
          DataLens
        </h1>
        <p style={{ color: P.muted, fontSize: 15, marginBottom: 40, lineHeight: 1.6 }}>
          Drop your Excel or CSV file. Get instant AI-powered charts, stats, and insights — no setup required.
        </p>

        {/* Drop Zone */}
        <div className="upload-zone" onDragOver={e => { e.preventDefault(); setDragOver(true); }} onDragLeave={() => setDragOver(false)} onDrop={onDrop}
          onClick={() => fileRef.current.click()}
          style={{
            border: `2px dashed ${dragOver ? P.accent : P.border}`,
            borderRadius: 20, padding: "48px 32px",
            background: dragOver ? P.accent + "0d" : P.card,
            cursor: "pointer", transition: "all 0.2s",
          }}
        >
          <div style={{ fontSize: 40, marginBottom: 12 }}>📊</div>
          <div style={{ fontWeight: 600, fontSize: 16, marginBottom: 6 }}>Drop your file here</div>
          <div style={{ color: P.muted, fontSize: 13, marginBottom: 20 }}>or click to browse</div>
          <div style={{ display: "flex", justifyContent: "center", gap: 8, flexWrap: "wrap" }}>
            {[".xlsx", ".xls", ".csv"].map(t => <Badge key={t} color={P.accent}>{t}</Badge>)}
          </div>
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={e => processFile(e.target.files[0])} />

        {error && <div style={{ marginTop: 16, color: P.danger, fontSize: 13 }}>⚠ {error}</div>}

        <div style={{ marginTop: 32, display: "flex", justifyContent: "center", gap: 24, color: P.muted, fontSize: 12 }}>
          {["AI Insights", "Auto Charts", "Full Stats", "No Signup"].map(f => (
            <div key={f} style={{ display: "flex", alignItems: "center", gap: 4 }}>
              <GlowDot color={P.accent3} />{f}
            </div>
          ))}
        </div>
      </div>
    </div>
  );

  // ── ANALYZING SCREEN ───────────────────────────────────────────────────────
  if (stage === "analyzing") return (
    <div style={{ ...gs, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "100vh" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600;700&family=DM+Mono:wght@400;500&display=swap'); * { box-sizing: border-box; } @keyframes spin { to { transform: rotate(360deg); } } @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }`}</style>
      <div style={{ position: "fixed", inset: 0, backgroundImage: `linear-gradient(${P.border} 1px, transparent 1px), linear-gradient(90deg, ${P.border} 1px, transparent 1px)`, backgroundSize: "40px 40px", opacity: 0.3, pointerEvents: "none" }} />
      <div style={{ position: "relative", textAlign: "center", maxWidth: 400 }}>
        <div style={{ width: 64, height: 64, border: `3px solid ${P.border}`, borderTop: `3px solid ${P.accent}`, borderRadius: "50%", animation: "spin 1s linear infinite", margin: "0 auto 24px" }} />
        <div style={{ fontWeight: 700, fontSize: 20, marginBottom: 8 }}>Analyzing <span style={{ color: P.accent }}>{fileName}</span></div>
        <div style={{ color: P.muted, fontSize: 13, marginBottom: 28, animation: "pulse 2s ease-in-out infinite" }}>
          {progress < 40 ? "Reading your data…" : progress < 70 ? "Detecting columns & types…" : "Generating AI insights…"}
        </div>
        <div style={{ background: P.card, borderRadius: 99, height: 6, overflow: "hidden", border: `1px solid ${P.border}` }}>
          <div style={{ height: "100%", width: `${progress}%`, background: `linear-gradient(90deg, ${P.accent2}, ${P.accent})`, borderRadius: 99, transition: "width 0.4s ease" }} />
        </div>
        <div style={{ marginTop: 10, fontSize: 12, color: P.muted, fontFamily: "'DM Mono', monospace" }}>{progress}%</div>
      </div>
    </div>
  );

  // ── DASHBOARD ──────────────────────────────────────────────────────────────
  const tabs = ["overview", "charts", "distributions", "table"];
  const tabLabels = { overview: "📊 Overview", charts: "📈 Charts", distributions: "🥧 Distributions", table: "🗃 Data Table" };

  return (
    <div style={gs}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { background: ${P.bg}; }
        ::-webkit-scrollbar { width: 5px; height: 5px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: ${P.border}; border-radius: 3px; }
        .tab-btn { transition: all 0.2s !important; }
        .tab-btn:hover { background: ${P.border} !important; }
        .row-hover:hover { background: ${P.border}22 !important; }
        @keyframes fadeUp { from { opacity:0; transform:translateY(12px); } to { opacity:1; transform:translateY(0); } }
        .fade-up { animation: fadeUp 0.4s ease forwards; }
      `}</style>

      {/* Top Bar */}
      <div style={{ background: P.surface, borderBottom: `1px solid ${P.border}`, padding: "0 24px", display: "flex", alignItems: "center", justifyContent: "space-between", height: 58, position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <span style={{ fontSize: 22, filter: `drop-shadow(0 0 8px ${P.accent})` }}>⬡</span>
          <span style={{ fontWeight: 700, fontSize: 16, fontFamily: "'Space Grotesk', sans-serif" }}>DataLens</span>
          <span style={{ color: P.border }}>|</span>
          <span style={{ fontSize: 13, color: P.muted, maxWidth: 240, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{fileName}</span>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <Badge color={P.accent3}>{rawData.length.toLocaleString()} rows</Badge>
          <Badge color={P.accent}>{columns.length} cols</Badge>
          <button onClick={() => { setStage("upload"); setRawData([]); }}
            style={{ background: "transparent", border: `1px solid ${P.border}`, color: P.muted, borderRadius: 8, padding: "5px 14px", cursor: "pointer", fontSize: 12, fontFamily: "'Space Grotesk', sans-serif" }}>
            ↩ New File
          </button>
        </div>
      </div>

      {/* Tab Bar */}
      <div style={{ background: P.surface, borderBottom: `1px solid ${P.border}`, padding: "0 24px", display: "flex", gap: 4 }}>
        {tabs.map(t => (
          <button key={t} className="tab-btn" onClick={() => setActiveTab(t)}
            style={{
              background: activeTab === t ? P.card : "transparent",
              border: "none", borderBottom: activeTab === t ? `2px solid ${P.accent}` : "2px solid transparent",
              color: activeTab === t ? P.text : P.muted, padding: "12px 16px",
              cursor: "pointer", fontSize: 13, fontFamily: "'Space Grotesk', sans-serif",
              fontWeight: activeTab === t ? 600 : 400, transition: "all 0.2s",
            }}>
            {tabLabels[t]}
          </button>
        ))}
      </div>

      <div style={{ padding: "24px", maxWidth: 1400, margin: "0 auto" }} className="fade-up">

        {/* ── OVERVIEW TAB ── */}
        {activeTab === "overview" && (
          <div>
            {/* Stat cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))", gap: 16, marginBottom: 24 }}>
              <StatCard label="Total Rows" value={rawData.length.toLocaleString()} icon="📋" color={P.accent} sub="records in dataset" />
              <StatCard label="Columns" value={columns.length} icon="📐" color={P.accent2} sub={`${numericCols.length} numeric, ${catCols.length} text`} />
              <StatCard label="Numeric Cols" value={numericCols.length} icon="🔢" color={P.accent3} sub="auto-detected" />
              <StatCard label="Text Cols" value={catCols.length} icon="🏷️" color={P.warn} sub="categorical" />
              {numericStats[0] && <StatCard label={`Avg ${numericStats[0].key}`} value={fmt(numericStats[0].stats.mean)} icon="📊" color={P.accent} sub={`max ${fmt(numericStats[0].stats.max)}`} />}
              {numericStats[1] && <StatCard label={`Avg ${numericStats[1].key}`} value={fmt(numericStats[1].stats.mean)} icon="📈" color={P.accent2} sub={`max ${fmt(numericStats[1].stats.max)}`} />}
            </div>

            {/* AI Insights */}
            <div style={{ marginBottom: 24 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
                <span style={{ fontSize: 18 }}>✦</span>
                <span style={{ fontWeight: 700, fontSize: 16 }}>AI Insights</span>
                <Badge color={P.accent2}>Powered by Claude</Badge>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 10 }}>
                {insights.map((ins, i) => (
                  <InsightCard key={i} text={ins} color={P.chartColors[i % P.chartColors.length]} />
                ))}
              </div>
            </div>

            {/* Column summary table */}
            <ChartCard title="Column Summary" subtitle="Auto-detected types and key statistics" span={2}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${P.border}` }}>
                      {["Column", "Type", "Min", "Max", "Mean", "Std Dev", "Count"].map(h => (
                        <th key={h} style={{ padding: "8px 12px", textAlign: "left", color: P.muted, fontWeight: 500, fontFamily: "'DM Mono', monospace", fontSize: 11 }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {columns.slice(0, 20).map((col, i) => {
                      const s = numericCols.includes(col) ? getColumnStats(rawData, col) : null;
                      return (
                        <tr key={col} className="row-hover" style={{ borderBottom: `1px solid ${P.border}44` }}>
                          <td style={{ padding: "8px 12px", fontWeight: 500 }}>{col}</td>
                          <td style={{ padding: "8px 12px" }}>
                            <Badge color={s ? P.accent : P.warn}>{s ? "numeric" : "text"}</Badge>
                          </td>
                          <td style={{ padding: "8px 12px", fontFamily: "'DM Mono', monospace", color: P.muted }}>{s ? fmt(s.min) : "—"}</td>
                          <td style={{ padding: "8px 12px", fontFamily: "'DM Mono', monospace", color: P.muted }}>{s ? fmt(s.max) : "—"}</td>
                          <td style={{ padding: "8px 12px", fontFamily: "'DM Mono', monospace" }}>{s ? fmt(s.mean) : "—"}</td>
                          <td style={{ padding: "8px 12px", fontFamily: "'DM Mono', monospace", color: P.muted }}>{s ? fmt(s.std) : "—"}</td>
                          <td style={{ padding: "8px 12px", fontFamily: "'DM Mono', monospace", color: P.muted }}>{s ? s.count : "—"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </ChartCard>
          </div>
        )}

        {/* ── CHARTS TAB ── */}
        {activeTab === "charts" && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(480px, 1fr))", gap: 20 }}>

            {/* Bar chart: top numeric cols */}
            {numericStats.length > 0 && (
              <ChartCard title="Column Averages" subtitle="Mean value per numeric column" span={2}>
                <ResponsiveContainer width="100%" height={260}>
                  <BarChart data={numericStats.map(s => ({ name: s.key.length > 16 ? s.key.slice(0, 15) + "…" : s.key, value: s.stats.mean }))} barSize={32}>
                    <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                    <XAxis dataKey="name" tick={{ fill: P.muted, fontSize: 11 }} />
                    <YAxis tickFormatter={fmt} tick={{ fill: P.muted, fontSize: 11 }} />
                    <Tooltip content={<CustomTooltip />} />
                    <Bar dataKey="value" name="Mean" radius={[6, 6, 0, 0]}>
                      {numericStats.map((_, i) => <Cell key={i} fill={P.chartColors[i % P.chartColors.length]} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
            )}

            {/* Line/Area chart */}
            {lineData.length > 1 && numericCols[0] && (
              <ChartCard title="Trend Over Records" subtitle={`${numericCols.slice(0, 3).join(", ")} sampled across rows`} span={2}>
                <ResponsiveContainer width="100%" height={240}>
                  <AreaChart data={lineData}>
                    <defs>
                      {numericCols.slice(0, 3).map((c, i) => (
                        <linearGradient key={c} id={`grad${i}`} x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor={P.chartColors[i]} stopOpacity={0.3} />
                          <stop offset="95%" stopColor={P.chartColors[i]} stopOpacity={0} />
                        </linearGradient>
                      ))}
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                    <XAxis dataKey="index" tick={{ fill: P.muted, fontSize: 10 }} label={{ value: "Row Index", position: "insideBottom", offset: -2, fill: P.muted, fontSize: 11 }} />
                    <YAxis tickFormatter={fmt} tick={{ fill: P.muted, fontSize: 10 }} />
                    <Tooltip content={<CustomTooltip />} />
                    <Legend wrapperStyle={{ fontSize: 12, color: P.muted }} />
                    {numericCols.slice(0, 3).map((c, i) => (
                      <Area key={c} type="monotone" dataKey={c} stroke={P.chartColors[i]} fill={`url(#grad${i})`} strokeWidth={2} dot={false} />
                    ))}
                  </AreaChart>
                </ResponsiveContainer>
              </ChartCard>
            )}

            {/* Scatter */}
            {scatterData.length > 0 && numericCols.length >= 2 && (
              <ChartCard title={`${numericCols[0]} vs ${numericCols[1]}`} subtitle="Scatter plot — spot correlations and clusters">
                <ResponsiveContainer width="100%" height={240}>
                  <ScatterChart>
                    <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                    <XAxis dataKey="x" name={numericCols[0]} tickFormatter={fmt} tick={{ fill: P.muted, fontSize: 10 }} label={{ value: numericCols[0], position: "insideBottom", offset: -2, fill: P.muted, fontSize: 11 }} />
                    <YAxis dataKey="y" name={numericCols[1]} tickFormatter={fmt} tick={{ fill: P.muted, fontSize: 10 }} />
                    <Tooltip cursor={{ strokeDasharray: "3 3" }} content={<CustomTooltip />} />
                    <Scatter data={scatterData} fill={P.accent} fillOpacity={0.6} />
                  </ScatterChart>
                </ResponsiveContainer>
              </ChartCard>
            )}

            {/* Min/Max/Mean grouped bar */}
            {numericStats.length > 0 && (
              <ChartCard title="Min / Mean / Max" subtitle="Range overview per numeric column">
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={numericStats.slice(0, 6).map(s => ({ name: s.key.length > 12 ? s.key.slice(0, 11) + "…" : s.key, min: s.stats.min, mean: Number(s.stats.mean.toFixed(2)), max: s.stats.max }))}>
                    <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                    <XAxis dataKey="name" tick={{ fill: P.muted, fontSize: 10 }} />
                    <YAxis tickFormatter={fmt} tick={{ fill: P.muted, fontSize: 10 }} />
                    <Tooltip content={<CustomTooltip />} />
                    <Legend wrapperStyle={{ fontSize: 11, color: P.muted }} />
                    <Bar dataKey="min" fill={P.accent2} radius={[4, 4, 0, 0]} barSize={14} />
                    <Bar dataKey="mean" fill={P.accent} radius={[4, 4, 0, 0]} barSize={14} />
                    <Bar dataKey="max" fill={P.accent3} radius={[4, 4, 0, 0]} barSize={14} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
            )}
          </div>
        )}

        {/* ── DISTRIBUTIONS TAB ── */}
        {activeTab === "distributions" && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(380px, 1fr))", gap: 20 }}>
            {catFreqs.map(({ col, data }, ci) => (
              <ChartCard key={col} title={col} subtitle={`Top ${data.length} values by frequency`}>
                <ResponsiveContainer width="100%" height={220}>
                  <PieChart>
                    <Pie data={data} cx="50%" cy="50%" outerRadius={80} dataKey="value" nameKey="name" label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`} labelLine={false} fontSize={10}>
                      {data.map((_, i) => <Cell key={i} fill={P.chartColors[i % P.chartColors.length]} />)}
                    </Pie>
                    <Tooltip content={<CustomTooltip />} />
                  </PieChart>
                </ResponsiveContainer>
                {/* bar version below pie */}
                <div style={{ marginTop: 8 }}>
                  {data.slice(0, 6).map((d, i) => (
                    <div key={d.name} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 5 }}>
                      <div style={{ width: 10, height: 10, borderRadius: 2, background: P.chartColors[i % P.chartColors.length], flexShrink: 0 }} />
                      <div style={{ flex: 1, fontSize: 11, color: P.muted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{d.name}</div>
                      <div style={{ width: `${(d.value / data[0].value) * 120}px`, height: 6, borderRadius: 3, background: P.chartColors[i % P.chartColors.length] + "88", transition: "width 0.4s", flexShrink: 0 }} />
                      <div style={{ fontSize: 11, fontFamily: "'DM Mono', monospace", color: P.text, minWidth: 28, textAlign: "right" }}>{d.value}</div>
                    </div>
                  ))}
                </div>
              </ChartCard>
            ))}
            {numericStats.slice(0, 4).map((s, i) => (
              <ChartCard key={s.key} title={s.key} subtitle="Numeric distribution summary">
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
                  {[["Min", s.stats.min], ["Max", s.stats.max], ["Mean", s.stats.mean], ["Median", s.stats.median], ["Std Dev", s.stats.std], ["Count", s.stats.count]].map(([lbl, val]) => (
                    <div key={lbl} style={{ background: P.surface, borderRadius: 10, padding: "10px 12px", border: `1px solid ${P.border}` }}>
                      <div style={{ fontSize: 10, color: P.muted, fontFamily: "'DM Mono', monospace", marginBottom: 2 }}>{lbl}</div>
                      <div style={{ fontSize: 18, fontWeight: 700, color: P.chartColors[i % P.chartColors.length] }}>{fmt(val)}</div>
                    </div>
                  ))}
                </div>
                {/* range bar */}
                <div style={{ background: P.surface, borderRadius: 8, padding: "10px 12px", border: `1px solid ${P.border}` }}>
                  <div style={{ fontSize: 11, color: P.muted, marginBottom: 6 }}>Range distribution</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 10, color: P.muted }}>
                    <span style={{ fontFamily: "'DM Mono', monospace" }}>{fmt(s.stats.min)}</span>
                    <div style={{ flex: 1, height: 8, background: P.border, borderRadius: 4, position: "relative", overflow: "hidden" }}>
                      <div style={{
                        position: "absolute", height: "100%", borderRadius: 4,
                        left: `${((s.stats.mean - s.std - s.stats.min) / (s.stats.max - s.stats.min)) * 100}%`,
                        width: `${(s.stats.std * 2 / (s.stats.max - s.stats.min)) * 100}%`,
                        background: `linear-gradient(90deg, ${P.chartColors[i % P.chartColors.length]}66, ${P.chartColors[i % P.chartColors.length]})`,
                      }} />
                      <div style={{
                        position: "absolute", height: "100%", width: 3, background: P.chartColors[i % P.chartColors.length],
                        left: `${((s.stats.mean - s.stats.min) / (s.stats.max - s.stats.min)) * 100}%`,
                      }} />
                    </div>
                    <span style={{ fontFamily: "'DM Mono', monospace" }}>{fmt(s.stats.max)}</span>
                  </div>
                  <div style={{ textAlign: "center", fontSize: 10, color: P.muted, marginTop: 3 }}>▲ mean = {fmt(s.stats.mean)}</div>
                </div>
              </ChartCard>
            ))}
          </div>
        )}

        {/* ── TABLE TAB ── */}
        {activeTab === "table" && (
          <div>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
              <div style={{ fontSize: 13, color: P.muted }}>Showing first 200 rows of {rawData.length.toLocaleString()} · {columns.length} columns</div>
              <Badge color={P.accent3}>Scroll to explore →</Badge>
            </div>
            <div style={{ background: P.card, border: `1px solid ${P.border}`, borderRadius: 16, overflow: "hidden" }}>
              <div style={{ overflowX: "auto", maxHeight: "70vh", overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ position: "sticky", top: 0, background: P.surface, zIndex: 10 }}>
                    <tr>
                      <th style={{ padding: "10px 14px", borderBottom: `1px solid ${P.border}`, textAlign: "left", color: P.muted, fontFamily: "'DM Mono', monospace", fontSize: 10, minWidth: 40 }}>#</th>
                      {columns.map(col => (
                        <th key={col} style={{ padding: "10px 14px", borderBottom: `1px solid ${P.border}`, textAlign: "left", color: P.muted, fontFamily: "'DM Mono', monospace", fontSize: 10, minWidth: 100, whiteSpace: "nowrap" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                            <span style={{ width: 6, height: 6, borderRadius: "50%", background: numericCols.includes(col) ? P.accent : P.warn, display: "inline-block" }} />
                            {col}
                          </div>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rawData.slice(0, 200).map((row, ri) => (
                      <tr key={ri} className="row-hover" style={{ borderBottom: `1px solid ${P.border}22` }}>
                        <td style={{ padding: "7px 14px", color: P.border, fontFamily: "'DM Mono', monospace" }}>{ri + 1}</td>
                        {columns.map(col => (
                          <td key={col} style={{ padding: "7px 14px", color: numericCols.includes(col) ? P.accent : P.text, fontFamily: numericCols.includes(col) ? "'DM Mono', monospace" : "inherit", whiteSpace: "nowrap", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis" }}>
                            {String(row[col] ?? "")}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
