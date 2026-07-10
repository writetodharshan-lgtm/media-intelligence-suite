import { useState, useRef, useEffect } from "react";
import Head from "next/head";

// ── PPTX TEXT EXTRACTOR ───────────────────────────────────────────────────────
async function extractPptxText(file) {
  if (!window.JSZip) {
    await new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
  }
  const JSZip = window.JSZip;
  const zip = await JSZip.loadAsync(file);
  const slideFiles = Object.keys(zip.files)
    .filter(n => n.match(/^ppt\/slides\/slide\d+\.xml$/))
    .sort((a, b) => parseInt(a.match(/\d+/)?.[0]) - parseInt(b.match(/\d+/)?.[0]));
  const texts = await Promise.all(
    slideFiles.map(async (name, i) => {
      const xml = await zip.files[name].async("string");
      const text = xml
        .replace(/<a:br\/>/g, "\n").replace(/<\/a:p>/g, "\n").replace(/<[^>]+>/g, "")
        .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
        .replace(/[ \t]+/g, " ").replace(/\n{3,}/g, "\n\n").trim();
      return `--- Slide ${i + 1} ---\n${text}`;
    })
  );
  return texts.join("\n\n");
}

// ── TOOLS CONFIG (logic preserved verbatim) ───────────────────────────────────
const TOOLKIT_TOOLS = [
  {
    id: "01", label: "Summary Drafter", name: "Summary Generator",
    description: "Turn raw coverage metrics into a board-ready executive summary.",
    color: "#8B5CF6", glow: "rgba(139,92,246,0.4)", icon: "◈",
    buttonLabel: "Generate Summary",
    inputs: [{ key: "data", label: "Paste your coverage metrics below", rows: 12,
      placeholder: "Example:\n— Total articles: 877\n— Total reach: 1.77B\n— Sentiment: 94% Positive\n— Top outlet: Associated Press (220M reach)" }],
    system: "You are a senior media intelligence analyst. Write a concise, accurate executive summary from the pre-calculated media metrics provided.\n\nRules:\n1. Use ONLY the numbers and information present in the data\n2. Do not recalculate, estimate or change any figures\n3. Keep the narrative company-focused\n4. Do not overemphasise competitors\n5. Use a neutral, executive tone throughout\n6. Cover all key metrics present: volume, reach, engagement, MIS (if present), sentiment, date range, top outlets\n7. Mention regional breakdown if present\n8. Keep the summary concise, no padding\n9. Write in plain flowing paragraphs only. No markdown, no headers, no bullet points, no bold, no asterisks. Pure prose only.",
    buildPrompt: (v) => "Here are the pre-calculated media metrics. Write an executive summary based on these exact figures:\n\n" + v.data
  },
  {
    id: "02", label: "Summary Checker", name: "Summary Checker",
    description: "Cross-check a written summary against the underlying numbers.",
    color: "#10B981", glow: "rgba(16,185,129,0.4)", icon: "◉",
    buttonLabel: "Run Quality Check",
    inputs: [
      { key: "datasheet", label: "Raw Datasheet", rows: 6, placeholder: "Paste the raw media datasheet here..." },
      { key: "summary", label: "Analyst Summary", rows: 6, placeholder: "Paste the analyst written summary here..." }
    ],
    system: "You are a quality reviewer for media intelligence reports.\n\nCheck the summary against the metrics for:\n1. Numerical accuracy - are all figures correct?\n2. Unsupported interpretations - any claims not in the data?\n3. Competitor focus - is it excessive or distracting?\n4. Overstated headlines - are claims proportionate?\n5. Tone - is it neutral and executive throughout?\n6. Prioritisation - are the most important insights featured prominently?\n\nReturn feedback in this format:\nOVERALL OBSERVATIONS\nSLIDE-LEVEL FEEDBACK\nKEY IMPROVEMENT AREAS",
    buildPrompt: (v) => "(1) Pre-calculated media metrics:\n" + v.datasheet + "\n\n(2) Analyst summary:\n" + v.summary
  },
  {
    id: "03", label: "PPT Validator", name: "PPT Validator",
    description: "Catch mismatches between your datasheet and deck before delivery.",
    color: "#F43F5E", glow: "rgba(244,63,94,0.4)", icon: "◎",
    buttonLabel: "Validate Data",
    inputs: [
      { key: "datasheet", label: "Raw Datasheet", rows: 6, placeholder: "Paste the raw media datasheet here..." },
      { key: "ppt", label: "PPT / Written Summary", rows: 6, placeholder: "Paste the PPT or written summary content here..." }
    ],
    system: "You are a data validation specialist for media reports.\n\nCheck for mismatches in: Reach, Engagement, MIS, Posts/Article count, Sentiment, Regional breakdown, Outlet tiers, Rankings.\nAuthor ranking protocol: Volume first > Sentiment second > Reach third\n\nOutput: List each metric with CORRECT or ISSUE DETECTED. Describe any issue clearly and concisely.",
    buildPrompt: (v) => "(1) Pre-calculated media metrics:\n" + v.datasheet + "\n\n(2) PPT/Summary to validate:\n" + v.ppt
  },
  {
    id: "04", label: "Key Insights Extractor", name: "Key Insights Extractor",
    description: "Surface themes, messages and patterns across 20–100 articles.",
    color: "#F59E0B", glow: "rgba(245,158,11,0.4)", icon: "◇",
    buttonLabel: "Extract Insights",
    inputs: [{ key: "articles", label: "Article Batch", rows: 14,
      placeholder: "Paste 20-100 article headlines, summaries or full text here." }],
    system: "You are a media content analyst.\n\nAnalyse the full set and extract:\n1. TOP THEMES - The 3-5 most dominant topics across all articles.\n2. KEY MESSAGES - The core narratives being communicated.\n3. BRAND MENTIONS - Which brands appear most and in what context.\n4. NARRATIVE PATTERNS - Any recurring story angles, framing or tone.\n\nKeep all insights concise and factual. No commentary or opinions.",
    buildPrompt: (v) => "Here is the article data to analyse:\n\n" + v.articles
  }
];

// ── TOOL ICONS (line SVGs, colored via currentColor) ──────────────────────────
function ToolIcon({ id, size = 18 }) {
  const p = { width: size, height: size, viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", strokeWidth: 1.6, strokeLinecap: "round", strokeLinejoin: "round" };
  if (id === "01") return (<svg {...p}><path d="M14 3v4a1 1 0 0 0 1 1h4" /><path d="M5 3h9l5 5v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z" /><path d="M8 12h8M8 16h6" /></svg>);
  if (id === "02") return (<svg {...p}><path d="M12 3l7 3v5c0 4.5-3 7.6-7 9-4-1.4-7-4.5-7-9V6l7-3z" /><path d="M9 12l2 2 4-4" /></svg>);
  if (id === "03") return (<svg {...p}><circle cx="11" cy="11" r="6" /><path d="M20 20l-3.4-3.4" /><path d="M8.6 11l1.7 1.7 3.1-3.3" /></svg>);
  return (<svg {...p}><path d="M12 3l1.7 4.6L18.4 9.3 13.7 11 12 15.6 10.3 11 5.6 9.3 10.3 7.6 12 3z" /><path d="M18 14.2l.6 1.7 1.7.6-1.7.6-.6 1.7-.6-1.7-1.7-.6 1.7-.6.6-1.7z" /></svg>);
}

// ── THEME ─────────────────────────────────────────────────────────────────────
function getTheme(dark) {
  if (dark) return {
    dark: true,
    bg: "#0B0D12", navBg: "rgba(11,13,18,0.72)",
    surface: "#141824", surfaceAlt: "#0F131C", surfaceInput: "#0E121A",
    border: "rgba(255,255,255,0.08)", borderStrong: "rgba(255,255,255,0.14)",
    text: "#EDEFF3", textMid: "#9AA3B2", textDim: "#6B7484", textFaint: "#3C4353",
    cardShadow: "0 10px 40px -24px rgba(0,0,0,0.8)",
    errText: "#FCA5A5", errBg: "rgba(244,63,94,0.08)", errBorder: "rgba(244,63,94,0.28)",
    glowA: "rgba(139,92,246,0.16)", glowB: "rgba(16,185,129,0.10)",
    heroEyebrow: "#9AA3B2",
  };
  return {
    dark: false,
    bg: "#F2F4F8", navBg: "rgba(242,244,248,0.75)",
    surface: "#FFFFFF", surfaceAlt: "#F7F9FC", surfaceInput: "#F5F7FA",
    border: "rgba(15,23,42,0.09)", borderStrong: "rgba(15,23,42,0.16)",
    text: "#0F1420", textMid: "#49525F", textDim: "#6B7482", textFaint: "#AAB1BD",
    cardShadow: "0 12px 34px -22px rgba(15,23,42,0.32)",
    errText: "#B91C3B", errBg: "rgba(244,63,94,0.06)", errBorder: "rgba(244,63,94,0.30)",
    glowA: "rgba(139,92,246,0.10)", glowB: "rgba(16,185,129,0.07)",
    heroEyebrow: "#6B7482",
  };
}

// ── UTILS (logic preserved verbatim) ──────────────────────────────────────────
async function callClaude(system, userPrompt) {
  const res = await fetch("/api/claude", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ system, prompt: userPrompt })
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "API error");
  return data.text || "";
}

function useTypewriter(text, speed = 12) {
  const [displayed, setDisplayed] = useState("");
  const [done, setDone] = useState(false);
  useEffect(() => {
    if (!text) { setDisplayed(""); setDone(false); return; }
    setDisplayed(""); setDone(false);
    let i = 0;
    const iv = setInterval(() => {
      i += speed;
      if (i >= text.length) { setDisplayed(text); setDone(true); clearInterval(iv); }
      else setDisplayed(text.slice(0, i));
    }, 16);
    return () => clearInterval(iv);
  }, [text]);
  return { displayed, done };
}

function Spinner({ color = "#fff", size = 14 }) {
  return <div style={{ width: size, height: size, border: "2px solid rgba(128,128,128,0.25)", borderTop: `2px solid ${color}`, borderRadius: "50%", animation: "spin 0.7s linear infinite", flexShrink: 0 }} />;
}

function CopyBtn({ text, color, t }) {
  const [copied, setCopied] = useState(false);
  const copy = () => { navigator.clipboard.writeText(text); setCopied(true); setTimeout(() => setCopied(false), 2000); };
  return (
    <button onClick={copy} className="mis-chip" style={{
      background: copied ? color : "transparent",
      border: `1px solid ${copied ? color : t.border}`,
      color: copied ? "#fff" : t.textMid,
      borderRadius: 8, padding: "5px 13px", fontSize: 11, cursor: "pointer",
      fontFamily: "var(--mono)", letterSpacing: "0.02em", display: "inline-flex", alignItems: "center", gap: 6,
    }}>
      {copied ? "Copied" : "Copy"}
    </button>
  );
}

// ── TOOLKIT ───────────────────────────────────────────────────────────────────
function ToolkitSection({ t }) {
  const [activeTab, setActiveTab] = useState(0);
  const [values, setValues] = useState({});
  const [loading, setLoading] = useState(false);
  const [refining, setRefining] = useState("");
  const [customPrompt, setCustomPrompt] = useState("");
  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const resultRef = useRef(null);
  const { displayed, done } = useTypewriter(result || "");
  const tool = TOOLKIT_TOOLS[activeTab];

  const switchTab = (i) => { setActiveTab(i); setValues({}); setResult(null); setError(""); setRefining(""); setCustomPrompt(""); };

  const run = async () => {
    const missing = tool.inputs.find(inp => !values[inp.key]?.trim());
    if (missing) { setError("Please fill in: " + missing.label); return; }
    setLoading(true); setResult(null); setError(""); setRefining(""); setCustomPrompt("");
    try {
      const text = await callClaude(tool.system, tool.buildPrompt(values));
      setResult(text);
      setTimeout(() => resultRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }), 100);
    } catch (e) { setError(e.message); }
    setLoading(false);
  };

  const refine = async (label, instruction) => {
    if (!result) return;
    setRefining(label);
    try {
      const text = await callClaude(
        "You are a senior media intelligence analyst. Rewrite the summary below according to the instruction. Output ONLY the rewritten text — no preamble, no explanation, no markdown, no headers, no bold, no asterisks. Pure prose only.",
        "INSTRUCTION: " + instruction + "\n\nORIGINAL SUMMARY:\n" + result
      );
      setResult(text);
    } catch (e) { setError(e.message); }
    setRefining("");
  };

  const REFINE_PRESETS = [
    { label: "Single Paragraph", icon: "¶", instruction: "Condense this entire summary into a single concise paragraph. Keep all key metrics and figures but remove any repetition." },
    { label: "Rephrase", icon: "↻", instruction: "Rephrase this summary using completely different wording while preserving all data points, figures, and insights exactly as they are." },
    { label: "More Formal", icon: "◆", instruction: "Rewrite this summary in a more formal, boardroom-ready executive tone. Keep all data points intact." },
    { label: "Simplify", icon: "○", instruction: "Simplify the language of this summary so it's easy to understand for a non-expert reader. Keep all numbers and data intact but use plain, straightforward language." },
    { label: "Shorter", icon: "↓", instruction: "Cut this summary down to roughly half its current length. Keep only the most critical metrics and insights." },
  ];

  const TENSE_ACTIONS = [
    { label: "Past Tense", icon: "⏪", instruction: "Rewrite this summary entirely in past tense. Change ALL verbs to past tense (e.g. 'generates' → 'generated', 'shows' → 'showed'). Keep all data, figures, and insights exactly the same." },
    { label: "Present Tense", icon: "⏩", instruction: "Rewrite this summary entirely in present tense. Change ALL verbs to present tense (e.g. 'generated' → 'generates', 'showed' → 'shows'). Keep all data, figures, and insights exactly the same." },
  ];

  const isSummaryTool = tool.id === "01";

  const makeRefineBtn = (action) => (
    <button key={action.label} onClick={() => refine(action.label, action.instruction)}
      disabled={!!refining} className="mis-chip"
      style={{
        background: refining === action.label ? tool.color + "1A" : t.surface,
        border: `1px solid ${refining === action.label ? tool.color : t.border}`,
        borderRadius: 999, padding: "7px 14px", cursor: refining ? "not-allowed" : "pointer",
        display: "inline-flex", alignItems: "center", gap: 7,
        opacity: refining && refining !== action.label ? 0.4 : 1,
        color: refining === action.label ? tool.color : t.textMid,
      }}>
      {refining === action.label
        ? <Spinner color={tool.color} size={12} />
        : <span style={{ fontSize: 11, color: tool.color, lineHeight: 1 }}>{action.icon}</span>}
      <span style={{ fontSize: 11.5, fontFamily: "var(--mono)", letterSpacing: "0.01em" }}>
        {refining === action.label ? "Refining…" : action.label}
      </span>
    </button>
  );

  const inputStyle = {
    width: "100%", background: t.surfaceInput, border: `1px solid ${t.border}`,
    borderRadius: 12, padding: "13px 15px", color: t.text, fontSize: 13,
    lineHeight: 1.7, fontFamily: "var(--mono)", resize: "vertical",
  };
  const labelStyle = { fontSize: 10, fontFamily: "var(--mono)", color: tool.color, letterSpacing: "0.14em", textTransform: "uppercase", display: "block", marginBottom: 8, fontWeight: 500 };

  return (
    <div style={{ "--accent": tool.color }}>
      {/* Tool selector */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(230px, 1fr))", gap: 12, marginBottom: 34 }}>
        {TOOLKIT_TOOLS.map((tl, i) => {
          const on = activeTab === i;
          return (
            <button key={tl.id} onClick={() => switchTab(i)} className="mis-card"
              aria-pressed={on}
              style={{
                "--accent": tl.color,
                position: "relative", overflow: "hidden", textAlign: "left", cursor: "pointer",
                background: on ? tl.color + "12" : t.surface,
                border: `1px solid ${on ? tl.color : t.border}`,
                borderRadius: 16, padding: "16px 17px",
                boxShadow: on ? `0 14px 40px -22px ${tl.color}` : t.cardShadow,
              }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
                <span style={{
                  display: "inline-flex", alignItems: "center", justifyContent: "center",
                  width: 34, height: 34, borderRadius: 10,
                  background: on ? tl.color + "22" : (t.dark ? "rgba(255,255,255,0.04)" : "rgba(15,23,42,0.04)"),
                  color: on ? tl.color : t.textDim, transition: "all .18s ease",
                }}>
                  <ToolIcon id={tl.id} size={18} />
                </span>
                <span style={{ display: "flex", alignItems: "center", gap: 9 }}>
                  {on && <span className="mis-pulse" style={{ width: 6, height: 6, borderRadius: "50%", background: tl.color }} />}
                  <span style={{ fontSize: 10, fontFamily: "var(--mono)", color: on ? tl.color : t.textFaint, letterSpacing: "0.14em" }}>{tl.id}</span>
                </span>
              </div>
              <div style={{ fontSize: 14, fontWeight: 600, color: on ? t.text : t.textMid, marginBottom: 4, letterSpacing: "-0.01em" }}>{tl.name}</div>
              <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", gap: 10 }}>
                <div style={{ fontSize: 12, color: on ? t.textMid : t.textDim, lineHeight: 1.55 }}>{tl.description}</div>
                <span className="mis-arrow" style={{ color: tl.color, flexShrink: 0, marginBottom: 1 }} aria-hidden="true">
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M5 12h14M13 6l6 6-6 6" /></svg>
                </span>
              </div>
            </button>
          );
        })}
      </div>

      {/* Working panel */}
      <div style={{ background: t.surface, border: `1px solid ${t.border}`, borderRadius: 20, boxShadow: t.cardShadow, overflow: "hidden" }}>
        <div className="mis-signal" style={{ "--accent": tool.color }} />
        <div style={{ padding: "22px 22px 24px" }}>
          {/* Tool header */}
          <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 22 }}>
            <div style={{
              width: 42, height: 42, borderRadius: 12, flexShrink: 0, color: "#fff",
              background: `linear-gradient(140deg, ${tool.color}, ${tool.color}CC)`,
              display: "flex", alignItems: "center", justifyContent: "center",
              boxShadow: `0 8px 22px -10px ${tool.color}`,
            }}>
              <ToolIcon id={tool.id} size={21} />
            </div>
            <div>
              <div style={{ fontSize: 16, fontWeight: 600, color: t.text, letterSpacing: "-0.01em" }}>{tool.name}</div>
              <div style={{ fontSize: 12.5, color: t.textDim, marginTop: 2 }}>{tool.description}</div>
            </div>
          </div>

          {/* Inputs (rendering logic preserved) */}
          {tool.inputs.map((inp) => {
            if ((tool.id === "02" || tool.id === "03") && inp.key === "datasheet") return (
              <div key={inp.key} style={{ marginBottom: 16 }}>
                <label style={labelStyle}>{inp.label}</label>
                <label className="mis-drop" style={{
                  display: "flex", alignItems: "center", gap: 12, background: t.surfaceInput,
                  border: `1px dashed ${t.borderStrong}`, borderRadius: 12, padding: "14px 16px",
                  cursor: "pointer", marginBottom: 8,
                }}>
                  <span style={{ display: "inline-flex", color: tool.color }}>
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round"><path d="M12 16V4M7 9l5-5 5 5" /><path d="M20 16v2a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-2" /></svg>
                  </span>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600, color: t.text }}>Upload CSV datasheet</div>
                    <div style={{ fontSize: 11.5, color: t.textDim, marginTop: 2 }}>{values[inp.key] ? "Loaded · " + values[inp.key].length.toLocaleString() + " characters" : "Drop a .csv file or click to browse"}</div>
                  </div>
                  <input type="file" accept=".csv,.txt" style={{ display: "none" }} onChange={e => { const f = e.target.files[0]; if (!f) return; const r = new FileReader(); r.onload = ev => setValues(v => ({ ...v, [inp.key]: ev.target.result })); r.readAsText(f); }} />
                </label>
                <textarea className="mis-input" rows={inp.rows} value={values[inp.key] || ""} placeholder={inp.placeholder + "\n\n(Or paste data here instead of uploading)"}
                  onChange={e => setValues(v => ({ ...v, [inp.key]: e.target.value }))} style={inputStyle} />
              </div>
            );

            if (tool.id === "03" && inp.key === "ppt") return (
              <div key={inp.key} style={{ marginBottom: 16 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <label style={{ ...labelStyle, marginBottom: 0 }}>{inp.label}</label>
                  <label className="mis-chip" style={{ display: "inline-flex", alignItems: "center", gap: 7, background: tool.color + "16", border: `1px solid ${tool.color}55`, borderRadius: 999, padding: "6px 13px", cursor: "pointer", fontSize: 11.5, color: tool.color, fontFamily: "var(--mono)" }}>
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M12 16V4M7 9l5-5 5 5" /></svg>
                    Upload deck / file
                    <input type="file" accept=".ppt,.pptx,.txt,.csv,.md" style={{ display: "none" }}
                      onChange={async e => {
                        const f = e.target.files[0]; if (!f) return;
                        if (f.name.endsWith(".pptx") || f.name.endsWith(".ppt")) {
                          try { const text = await extractPptxText(f); setValues(v => ({ ...v, [inp.key]: text })); }
                          catch (err) { setError("Could not read PPTX: " + err.message); }
                        } else {
                          const r = new FileReader(); r.onload = ev => setValues(v => ({ ...v, [inp.key]: ev.target.result })); r.readAsText(f);
                        }
                      }} />
                  </label>
                </div>
                {values[inp.key] && <div style={{ fontSize: 11, color: tool.color, fontFamily: "var(--mono)", marginBottom: 8, padding: "4px 11px", background: tool.color + "14", borderRadius: 999, display: "inline-block" }}>Loaded · {values[inp.key].length.toLocaleString()} characters</div>}
                <textarea className="mis-input" rows={inp.rows} value={values[inp.key] || ""} placeholder={inp.placeholder}
                  onChange={e => setValues(v => ({ ...v, [inp.key]: e.target.value }))} style={inputStyle} />
              </div>
            );

            if (tool.id === "04" && inp.key === "articles") return (
              <div key={inp.key} style={{ marginBottom: 16 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <label style={{ ...labelStyle, marginBottom: 0 }}>{inp.label}</label>
                  <label className="mis-chip" style={{ display: "inline-flex", alignItems: "center", gap: 7, background: tool.color + "16", border: `1px solid ${tool.color}55`, borderRadius: 999, padding: "6px 13px", cursor: "pointer", fontSize: 11.5, color: tool.color, fontFamily: "var(--mono)" }}>
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M12 16V4M7 9l5-5 5 5" /></svg>
                    Upload CSV
                    <input type="file" accept=".csv,.txt" style={{ display: "none" }} onChange={e => { const f = e.target.files[0]; if (!f) return; const r = new FileReader(); r.onload = ev => setValues(v => ({ ...v, [inp.key]: ev.target.result })); r.readAsText(f); }} />
                  </label>
                </div>
                <textarea className="mis-input" rows={inp.rows} value={values[inp.key] || ""} placeholder={inp.placeholder}
                  onChange={e => setValues(v => ({ ...v, [inp.key]: e.target.value }))} style={inputStyle} />
              </div>
            );

            return (
              <div key={inp.key} style={{ marginBottom: 16 }}>
                <label style={labelStyle}>{inp.label}</label>
                <textarea className="mis-input" rows={inp.rows} value={values[inp.key] || ""} placeholder={inp.placeholder}
                  onChange={e => setValues(v => ({ ...v, [inp.key]: e.target.value }))} style={inputStyle} />
              </div>
            );
          })}

          {error && (
            <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "10px 14px", borderRadius: 10, background: t.errBg, border: `1px solid ${t.errBorder}`, color: t.errText, fontSize: 12.5, marginBottom: 16 }}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="9" /><path d="M12 8v5M12 16.5v.01" /></svg>
              {error}
            </div>
          )}

          <button onClick={run} disabled={loading} className="mis-cta"
            style={{
              "--glow": tool.glow, width: "100%", padding: "14px", borderRadius: 13, border: "none",
              background: loading ? (t.dark ? "rgba(255,255,255,0.05)" : "rgba(15,23,42,0.05)") : `linear-gradient(135deg, ${tool.color}, ${tool.color}CC)`,
              color: loading ? t.textDim : "#fff", fontWeight: 600, fontSize: 14, letterSpacing: "0.01em",
              cursor: loading ? "not-allowed" : "pointer",
              display: "flex", alignItems: "center", justifyContent: "center", gap: 10,
              boxShadow: loading ? "none" : `0 10px 30px -12px ${tool.glow}`,
            }}>
            {loading
              ? <><Spinner color={tool.color} /><span style={{ color: tool.color }}>Analysing with Claude…</span></>
              : <><ToolIcon id={tool.id} size={17} />{tool.buttonLabel}</>}
          </button>
        </div>

        {/* Output */}
        {result && (
          <div ref={resultRef} style={{ borderTop: `1px solid ${t.border}`, background: t.surfaceAlt, padding: "20px 22px 24px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ width: 6, height: 6, borderRadius: "50%", background: tool.color }} />
                <span style={{ fontSize: 10, fontFamily: "var(--mono)", color: t.textDim, letterSpacing: "0.18em" }}>OUTPUT</span>
              </div>
              <CopyBtn text={result} color={tool.color} t={t} />
            </div>
            <div style={{ borderRadius: 14, border: `1px solid ${t.border}`, background: t.surface, padding: "18px 20px", whiteSpace: "pre-wrap", fontSize: 12.5, lineHeight: 1.9, color: t.textMid, fontFamily: "var(--mono)" }}>
              {displayed}{!done && <span style={{ color: tool.color }}>▍</span>}
            </div>

            {isSummaryTool && done && (
              <div style={{ marginTop: 18 }}>
                <div style={{ fontSize: 10, fontFamily: "var(--mono)", color: t.textFaint, letterSpacing: "0.16em", marginBottom: 11 }}>REFINE OUTPUT</div>
                <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
                  <input type="text" className="mis-input" value={customPrompt}
                    onChange={e => setCustomPrompt(e.target.value)}
                    onKeyDown={e => { if (e.key === "Enter" && customPrompt.trim() && !refining) refine("Custom", customPrompt.trim()); }}
                    placeholder="Type any instruction… e.g. 'Tone down the sentiment language'"
                    style={{ flex: 1, background: t.surfaceInput, border: `1px solid ${t.border}`, borderRadius: 10, padding: "10px 14px", color: t.text, fontSize: 12.5, fontFamily: "var(--mono)", outline: "none" }} />
                  <button onClick={() => { if (customPrompt.trim()) refine("Custom", customPrompt.trim()); }}
                    disabled={!customPrompt.trim() || !!refining} className="mis-cta"
                    style={{
                      "--glow": tool.glow,
                      background: !customPrompt.trim() || refining ? (t.dark ? "rgba(255,255,255,0.05)" : "rgba(15,23,42,0.05)") : `linear-gradient(135deg, ${tool.color}, ${tool.color}CC)`,
                      color: !customPrompt.trim() || refining ? t.textDim : "#fff",
                      border: "none", borderRadius: 10, padding: "10px 18px", fontSize: 12, fontFamily: "var(--mono)",
                      cursor: !customPrompt.trim() || refining ? "not-allowed" : "pointer", whiteSpace: "nowrap",
                      display: "flex", alignItems: "center", gap: 7,
                    }}>
                    {refining === "Custom" ? <><Spinner color="#fff" size={12} /> Refining…</> : "Refine"}
                  </button>
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 10 }}>
                  {TENSE_ACTIONS.map(makeRefineBtn)}
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                  {REFINE_PRESETS.map(makeRefineBtn)}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ── APP ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [dark, setDark] = useState(true);
  const t = getTheme(dark);

  return (
    <div style={{ background: t.bg, minHeight: "100vh", color: t.text, fontFamily: "var(--sans)", transition: "background 0.3s, color 0.3s", position: "relative" }}>
      <Head>
        <title>Media Intelligence Suite</title>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link rel="preconnect" href="https://fonts.gstatic.com" crossOrigin="anonymous" />
        <link href="https://fonts.googleapis.com/css2?family=Newsreader:ital,opsz,wght@0,6..72,400;0,6..72,500;1,6..72,400;1,6..72,500&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet" />
      </Head>
      <style>{`
        :root {
          --sans: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
          --serif: 'Newsreader', Georgia, serif;
          --mono: 'JetBrains Mono', ui-monospace, monospace;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes misRise { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: translateY(0); } }
        @keyframes misShimmer { 0% { background-position: 200% 0; } 100% { background-position: -200% 0; } }
        @keyframes misPulseKf { 0%,100% { opacity: 1; } 50% { opacity: 0.4; } }

        .mis-rise { animation: misRise 0.5s cubic-bezier(0.22,1,0.36,1) both; }
        .mis-signal { height: 2px; width: 100%; background: linear-gradient(90deg, transparent, var(--accent), transparent); background-size: 200% 100%; animation: misShimmer 3.4s linear infinite; }
        .mis-pulse { animation: misPulseKf 1.8s ease-in-out infinite; }

        .mis-card { transition: transform .18s ease, border-color .18s ease, box-shadow .18s ease, background .18s ease; }
        .mis-card:hover { transform: translateY(-3px); border-color: var(--accent); box-shadow: 0 16px 42px -22px var(--accent); }
        .mis-card:hover .mis-arrow { transform: translateX(4px); opacity: 1; }
        .mis-arrow { opacity: .55; transition: transform .18s ease, opacity .18s ease; }

        .mis-cta { transition: transform .16s ease, filter .16s ease, box-shadow .16s ease; }
        .mis-cta:not(:disabled):hover { transform: translateY(-1px); filter: brightness(1.07); box-shadow: 0 16px 38px -12px var(--glow) !important; }
        .mis-cta:not(:disabled):active { transform: translateY(0); }

        .mis-chip { transition: border-color .15s ease, color .15s ease, background .15s ease, transform .15s ease; }
        .mis-chip:not(:disabled):hover { transform: translateY(-1px); }

        .mis-drop { transition: border-color .18s ease, background .18s ease; }
        .mis-drop:hover { border-color: var(--accent); }

        .mis-input { transition: border-color .16s ease; }
        .mis-input:focus { border-color: var(--accent) !important; }

        .mis-toggle { transition: background .2s ease, border-color .2s ease; }
        .mis-toggle:hover { border-color: var(--accent, rgba(139,92,246,0.5)); }

        button:focus-visible, textarea:focus-visible, input:focus-visible { outline: 2px solid var(--accent, #8B5CF6); outline-offset: 2px; }
        textarea::placeholder, input::placeholder { color: currentColor; opacity: .38; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: rgba(128,128,128,0.28); border-radius: 4px; }

        @media (prefers-reduced-motion: reduce) {
          .mis-rise, .mis-signal, .mis-pulse { animation: none !important; }
          .mis-signal { background: var(--accent); opacity: .55; }
          .mis-card:hover { transform: none; }
        }
      `}</style>

      {/* Ambient glows */}
      <div aria-hidden="true" style={{ position: "fixed", inset: 0, pointerEvents: "none", background: `radial-gradient(70% 45% at 15% -5%, ${t.glowA}, transparent 60%), radial-gradient(60% 40% at 100% 0%, ${t.glowB}, transparent 55%)` }} />

      {/* NAV */}
      <div style={{ position: "sticky", top: 0, zIndex: 50, background: t.navBg, backdropFilter: "blur(14px)", WebkitBackdropFilter: "blur(14px)", borderBottom: `1px solid ${t.border}` }}>
        <div style={{ maxWidth: 940, margin: "0 auto", display: "flex", alignItems: "center", gap: 12, height: 58, padding: "0 22px" }}>
          <div style={{ width: 30, height: 30, borderRadius: 9, background: "linear-gradient(140deg, #8B5CF6, #6366F1)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, boxShadow: "0 6px 16px -8px #6366F1" }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M4 19V7M9 19V4M14 19v-8M19 19v-5" /></svg>
          </div>
          <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.15 }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: t.text }}>Media Intelligence Suite</span>
            <span style={{ fontSize: 10.5, color: t.textDim, fontFamily: "var(--mono)", letterSpacing: "0.04em" }}>Analyst workbench</span>
          </div>
          <div style={{ flex: 1 }} />
          <button onClick={() => setDark(d => !d)} className="mis-toggle" aria-label={dark ? "Switch to light mode" : "Switch to dark mode"}
            style={{ display: "flex", alignItems: "center", gap: 8, background: t.surface, border: `1px solid ${t.border}`, borderRadius: 999, padding: "7px 13px", cursor: "pointer", color: t.textMid, flexShrink: 0 }}>
            {dark
              ? <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="4.5" /><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4" /></svg>
              : <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z" /></svg>}
            <span style={{ fontSize: 11, fontFamily: "var(--mono)" }}>{dark ? "Light" : "Dark"}</span>
          </button>
        </div>
      </div>

      {/* MAIN */}
      <div style={{ position: "relative", maxWidth: 940, margin: "0 auto", padding: "56px 22px 72px" }}>
        <div className="mis-rise" style={{ marginBottom: 42 }}>
          <div style={{ fontSize: 11, fontFamily: "var(--mono)", color: t.heroEyebrow, letterSpacing: "0.24em", marginBottom: 18, textTransform: "uppercase" }}>
            Media intelligence · Four instruments
          </div>
          <h1 style={{ fontFamily: "var(--serif)", fontWeight: 500, lineHeight: 1.04, letterSpacing: "-0.015em", fontSize: "clamp(38px, 6.5vw, 62px)", color: t.text, marginBottom: 20 }}>
            Read the coverage.<br />
            Write the <em style={{ fontStyle: "italic", color: "#8B5CF6" }}>story</em>.
          </h1>
          <p style={{ fontSize: 15, color: t.textMid, lineHeight: 1.65, maxWidth: 540 }}>
            Draft executive summaries, quality-check analyst reports, validate decks against source data, and surface the themes that move a story — all in one workspace.
          </p>
        </div>

        <div className="mis-rise" style={{ animationDelay: "0.06s" }}>
          <ToolkitSection t={t} />
        </div>
      </div>

      <div style={{ position: "relative", textAlign: "center", padding: "24px", fontSize: 10.5, color: t.textFaint, fontFamily: "var(--mono)", letterSpacing: "0.04em", borderTop: `1px solid ${t.border}` }}>
        Media Intelligence Suite · Powered by Claude
      </div>
    </div>
  );
}
