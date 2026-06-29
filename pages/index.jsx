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

// ── TOOLS CONFIG ──────────────────────────────────────────────────────────────
const TOOLKIT_TOOLS = [
  {
    id: "01", label: "Summary Drafter", name: "Summary Generator",
    description: "Transform raw media data into a polished executive summary instantly.",
    color: "#8B5CF6", glow: "rgba(139,92,246,0.4)", icon: "◈",
    buttonLabel: "Generate Summary",
    inputs: [{ key: "data", label: "Paste your coverage metrics below", rows: 12,
      placeholder: "Example:\n— Total articles: 877\n— Total reach: 1.77B\n— Sentiment: 94% Positive\n— Top outlet: Associated Press (220M reach)" }],
    system: "You are a senior media intelligence analyst. Write a concise, accurate executive summary from the pre-calculated media metrics provided.\n\nRules:\n1. Use ONLY the numbers and information present in the data\n2. Do not recalculate, estimate or change any figures\n3. Keep the narrative company-focused\n4. Do not overemphasise competitors\n5. Use a neutral, executive tone throughout\n6. Cover all key metrics present: volume, reach, engagement, MIS (if present), sentiment, date range, top outlets\n7. Mention regional breakdown if present\n8. Keep the summary concise, no padding\n9. Write in plain flowing paragraphs only. No markdown, no headers, no bullet points, no bold, no asterisks. Pure prose only.",
    buildPrompt: (v) => "Here are the pre-calculated media metrics. Write an executive summary based on these exact figures:\n\n" + v.data
  },
  {
    id: "02", label: "Summary Checker", name: "Summary Checker",
    description: "Verify analyst summaries against raw data for accuracy and tone.",
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
    description: "Catch mismatches between your datasheet and PPT before client delivery.",
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
    description: "Extract themes, messages and patterns from 20-100 articles automatically.",
    color: "#F59E0B", glow: "rgba(245,158,11,0.4)", icon: "◇",
    buttonLabel: "Extract Insights",
    inputs: [{ key: "articles", label: "Article Batch", rows: 14,
      placeholder: "Paste 20-100 article headlines, summaries or full text here." }],
    system: "You are a media content analyst.\n\nAnalyse the full set and extract:\n1. TOP THEMES - The 3-5 most dominant topics across all articles.\n2. KEY MESSAGES - The core narratives being communicated.\n3. BRAND MENTIONS - Which brands appear most and in what context.\n4. NARRATIVE PATTERNS - Any recurring story angles, framing or tone.\n\nKeep all insights concise and factual. No commentary or opinions.",
    buildPrompt: (v) => "Here is the article data to analyse:\n\n" + v.articles
  }
];

// ── THEME ─────────────────────────────────────────────────────────────────────
function getTheme(dark) {
  if (dark) return {
    dark: true,
    bg: "#060810", bgNav: "rgba(6,8,16,0.97)", bgSub: "#080C14",
    bgCard: "#13131a", bgInput: "#0A0F1A",
    border: "rgba(255,255,255,0.05)", borderCard: "#1c1c28", borderInput: "#111827",
    text: "#F9FAFB", textMid: "#94A3B8", textDim: "#64748B", textFaint: "#334155",
    textGhost: "#1E3A5F", textFoot: "#1a2030",
    heroTitle: "#F1F5F9",
    gold: "#c8a96e", goldDim: "#8a8070", goldMid: "#a09888",
    outputBg: "#080B12",
    toggleBg: "#1a2030", infoBoxBg: "rgba(200,169,110,0.06)",
    infoBoxBorder: "rgba(200,169,110,0.18)", infoText: "#a09888",
    errBg: "#1a0f0f", errBorder: "#4a2020",
  };
  return {
    dark: false,
    bg: "#F8F9FB", bgNav: "rgba(255,255,255,0.97)", bgSub: "#EEF0F5",
    bgCard: "#FFFFFF", bgInput: "#F4F5F8",
    border: "rgba(0,0,0,0.07)", borderCard: "#E2E8F0", borderInput: "#DDE1EA",
    text: "#0F172A", textMid: "#334155", textDim: "#64748B", textFaint: "#94A3B8",
    textGhost: "#CBD5E1", textFoot: "#CBD5E1",
    heroTitle: "#0F172A",
    gold: "#96710A", goldDim: "#B08030", goldMid: "#7A5A10",
    outputBg: "#F0F4F8",
    toggleBg: "#E2E8F0", infoBoxBg: "rgba(150,113,10,0.05)",
    infoBoxBorder: "rgba(150,113,10,0.2)", infoText: "#6A4F10",
    errBg: "#FEF2F2", errBorder: "#FCA5A5",
  };
}

// ── UTILS ─────────────────────────────────────────────────────────────────────
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

function Spinner({ color = "#fff" }) {
  return <div style={{ width: 13, height: 13, border: "2px solid rgba(128,128,128,0.2)", borderTop: `2px solid ${color}`, borderRadius: "50%", animation: "spin 0.7s linear infinite", flexShrink: 0 }} />;
}

function CopyBtn({ text, color, t }) {
  const [copied, setCopied] = useState(false);
  const copy = () => { navigator.clipboard.writeText(text); setCopied(true); setTimeout(() => setCopied(false), 2000); };
  return (
    <button onClick={copy} style={{ background: copied ? color : "transparent", border: `1px solid ${copied ? color : t.borderCard}`, color: copied ? (t.dark ? "#0a0a0f" : "#fff") : t.textMid, borderRadius: 6, padding: "4px 12px", fontSize: 10, cursor: "pointer", fontFamily: "monospace" }}>
      {copied ? "✓ Copied" : "Copy"}
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
      disabled={!!refining}
      style={{
        background: refining === action.label ? tool.color + "20" : t.bgCard,
        border: `1px solid ${refining === action.label ? tool.color + "55" : t.borderCard}`,
        borderRadius: 6, padding: "7px 14px", cursor: refining ? "not-allowed" : "pointer",
        display: "flex", alignItems: "center", gap: 6, transition: "all 0.2s",
        opacity: refining && refining !== action.label ? 0.4 : 1,
      }}>
      {refining === action.label ? (
        <Spinner color={tool.color} />
      ) : (
        <span style={{ fontSize: 11, color: tool.color }}>{action.icon}</span>
      )}
      <span style={{ fontSize: 10, fontFamily: "monospace", color: refining === action.label ? tool.color : t.textMid }}>
        {refining === action.label ? "Refining..." : action.label}
      </span>
    </button>
  );

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 10, marginBottom: 24 }}>
        {TOOLKIT_TOOLS.map((tl, i) => (
          <button key={tl.id} onClick={() => switchTab(i)} style={{
            background: activeTab === i ? tl.color + "12" : t.bgCard,
            border: `1px solid ${activeTab === i ? tl.color + "55" : t.borderCard}`,
            borderRadius: 12, padding: "14px 16px", cursor: "pointer", textAlign: "left", transition: "all 0.2s"
          }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 5 }}>
              <span style={{ fontSize: 15, color: activeTab === i ? tl.color : t.textFaint }}>{tl.icon}</span>
              <span style={{ fontSize: 9, fontFamily: "monospace", color: activeTab === i ? tl.color : t.textFaint, letterSpacing: "0.12em" }}>{tl.id}</span>
            </div>
            <div style={{ fontSize: 12, fontWeight: 700, color: activeTab === i ? t.heroTitle : t.textDim, marginBottom: 3 }}>{tl.label}</div>
            <div style={{ fontSize: 10, color: activeTab === i ? t.textMid : t.textFaint, lineHeight: 1.5 }}>{tl.description}</div>
          </button>
        ))}
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 18, padding: "12px 16px", borderRadius: 10, background: tool.color + "0C", border: `1px solid ${tool.color}25` }}>
        <div style={{ width: 32, height: 32, borderRadius: 8, background: tool.color, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, flexShrink: 0 }}>{tool.icon}</div>
        <div>
          <div style={{ fontSize: 13, fontWeight: 700, color: t.heroTitle, marginBottom: 1 }}>{tool.name}</div>
          <div style={{ fontSize: 11, color: t.textDim }}>{tool.description}</div>
        </div>
      </div>

      {tool.inputs.map((inp) => {
        if ((tool.id === "02" || tool.id === "03") && inp.key === "datasheet") return (
          <div key={inp.key} style={{ marginBottom: 12 }}>
            <label style={{ fontSize: 9, fontFamily: "monospace", color: tool.color, letterSpacing: "0.12em", display: "block", marginBottom: 6 }}>{inp.label.toUpperCase()}</label>
            <label style={{ display: "flex", alignItems: "center", gap: 10, background: t.bgInput, border: `1px solid ${t.borderInput}`, borderRadius: 8, padding: "14px 16px", cursor: "pointer", marginBottom: 6 }}>
              <span style={{ fontSize: 20 }}>📎</span>
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, color: tool.color }}>Upload CSV Datasheet</div>
                <div style={{ fontSize: 11, color: t.textDim, marginTop: 2 }}>{values[inp.key] ? "✓ File loaded (" + values[inp.key].length.toLocaleString() + " chars)" : "Click to upload a .csv file"}</div>
              </div>
              <input type="file" accept=".csv,.txt" style={{ display: "none" }} onChange={e => { const f = e.target.files[0]; if (!f) return; const r = new FileReader(); r.onload = ev => setValues(v => ({ ...v, [inp.key]: ev.target.result })); r.readAsText(f); }} />
            </label>
            <textarea rows={inp.rows} value={values[inp.key] || ""} placeholder={inp.placeholder + "\n\n(Or paste data here instead of uploading)"}
              onChange={e => setValues(v => ({ ...v, [inp.key]: e.target.value }))}
              style={{ width: "100%", background: t.bgInput, border: `1px solid ${t.borderInput}`, borderRadius: 8, padding: "11px 13px", color: t.textMid, fontSize: 12, lineHeight: 1.7, fontFamily: "monospace", resize: "vertical" }} />
          </div>
        );

        if (tool.id === "03" && inp.key === "ppt") return (
          <div key={inp.key} style={{ marginBottom: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <label style={{ fontSize: 9, fontFamily: "monospace", color: tool.color, letterSpacing: "0.12em" }}>{inp.label.toUpperCase()}</label>
              <label style={{ display: "flex", alignItems: "center", gap: 6, background: tool.color + "18", border: `1px solid ${tool.color}55`, borderRadius: 6, padding: "5px 12px", cursor: "pointer", fontSize: 11, color: tool.color, fontFamily: "monospace" }}>
                Upload PPT/File
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
            {values[inp.key] && <div style={{ fontSize: 10, color: tool.color, fontFamily: "monospace", marginBottom: 6, padding: "4px 10px", background: tool.color + "12", borderRadius: 4, display: "inline-block" }}>File loaded ({values[inp.key].length.toLocaleString()} chars)</div>}
            <textarea rows={inp.rows} value={values[inp.key] || ""} placeholder={inp.placeholder}
              onChange={e => setValues(v => ({ ...v, [inp.key]: e.target.value }))}
              style={{ width: "100%", background: t.bgInput, border: `1px solid ${t.borderInput}`, borderRadius: 8, padding: "11px 13px", color: t.textMid, fontSize: 12, lineHeight: 1.7, fontFamily: "monospace", resize: "vertical" }} />
          </div>
        );

        if (tool.id === "04" && inp.key === "articles") return (
          <div key={inp.key} style={{ marginBottom: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <label style={{ fontSize: 9, fontFamily: "monospace", color: tool.color, letterSpacing: "0.12em" }}>{inp.label.toUpperCase()}</label>
              <label style={{ display: "flex", alignItems: "center", gap: 6, background: tool.color + "18", border: `1px solid ${tool.color}55`, borderRadius: 6, padding: "5px 12px", cursor: "pointer", fontSize: 11, color: tool.color, fontFamily: "monospace" }}>
                Upload CSV
                <input type="file" accept=".csv,.txt" style={{ display: "none" }} onChange={e => { const f = e.target.files[0]; if (!f) return; const r = new FileReader(); r.onload = ev => setValues(v => ({ ...v, [inp.key]: ev.target.result })); r.readAsText(f); }} />
              </label>
            </div>
            <textarea rows={inp.rows} value={values[inp.key] || ""} placeholder={inp.placeholder}
              onChange={e => setValues(v => ({ ...v, [inp.key]: e.target.value }))}
              style={{ width: "100%", background: t.bgInput, border: `1px solid ${t.borderInput}`, borderRadius: 8, padding: "11px 13px", color: t.textMid, fontSize: 12, lineHeight: 1.7, fontFamily: "monospace", resize: "vertical" }} />
          </div>
        );

        return (
          <div key={inp.key} style={{ marginBottom: 12 }}>
            <label style={{ fontSize: 9, fontFamily: "monospace", color: tool.color, letterSpacing: "0.12em", display: "block", marginBottom: 6 }}>{inp.label.toUpperCase()}</label>
            <textarea rows={inp.rows} value={values[inp.key] || ""} placeholder={inp.placeholder}
              onChange={e => setValues(v => ({ ...v, [inp.key]: e.target.value }))}
              style={{ width: "100%", background: t.bgInput, border: `1px solid ${t.borderInput}`, borderRadius: 8, padding: "11px 13px", color: t.textMid, fontSize: 12, lineHeight: 1.7, fontFamily: "monospace", resize: "vertical" }} />
          </div>
        );
      })}

      {error && <div style={{ padding: "8px 12px", borderRadius: 6, background: "rgba(244,63,94,0.06)", border: "1px solid rgba(244,63,94,0.2)", color: "#F87171", fontSize: 11, fontFamily: "monospace", marginBottom: 12 }}>⚠ {error}</div>}

      <button onClick={run} disabled={loading} style={{
        width: "100%", padding: "13px", borderRadius: 10, border: "none",
        background: loading ? t.toggleBg : `linear-gradient(135deg, ${tool.color}, ${tool.color}BB)`,
        color: loading ? t.textDim : "#fff", fontWeight: 700, fontSize: 13,
        cursor: loading ? "not-allowed" : "pointer",
        display: "flex", alignItems: "center", justifyContent: "center", gap: 8, marginBottom: 24,
        boxShadow: loading ? "none" : `0 6px 22px ${tool.glow}`
      }}>
        {loading ? <><Spinner color={tool.color} /><span style={{ color: tool.color }}>Analysing with Claude...</span></> : tool.icon + "  " + tool.buttonLabel}
      </button>

      {result && (
        <div ref={resultRef}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <div style={{ width: 5, height: 5, borderRadius: "50%", background: tool.color }} />
              <span style={{ fontSize: 9, fontFamily: "monospace", color: t.textDim, letterSpacing: "0.15em" }}>OUTPUT</span>
            </div>
            <CopyBtn text={result} color={tool.color} t={t} />
          </div>
          <div style={{ borderRadius: 12, border: `1px solid ${tool.color}22`, background: t.outputBg, padding: "18px 20px", whiteSpace: "pre-wrap", fontSize: 12, lineHeight: 1.85, color: t.textMid, fontFamily: "monospace" }}>
            {displayed}{!done && <span style={{ color: tool.color }}>|</span>}
          </div>

          {isSummaryTool && done && (
            <div style={{ marginTop: 14 }}>
              <div style={{ fontSize: 9, fontFamily: "monospace", color: t.textFaint, letterSpacing: "0.12em", marginBottom: 8 }}>REFINE OUTPUT</div>

              <div style={{ display: "flex", gap: 6, marginBottom: 10 }}>
                <input
                  type="text"
                  value={customPrompt}
                  onChange={e => setCustomPrompt(e.target.value)}
                  onKeyDown={e => { if (e.key === "Enter" && customPrompt.trim() && !refining) refine("Custom", customPrompt.trim()); }}
                  placeholder="Type any instruction... e.g. 'Tone down the sentiment language'"
                  style={{ flex: 1, background: t.bgInput, border: `1px solid ${t.borderInput}`, borderRadius: 6, padding: "8px 12px", color: t.textMid, fontSize: 11, fontFamily: "monospace", outline: "none" }}
                />
                <button
                  onClick={() => { if (customPrompt.trim()) refine("Custom", customPrompt.trim()); }}
                  disabled={!customPrompt.trim() || !!refining}
                  style={{
                    background: !customPrompt.trim() || refining ? t.toggleBg : tool.color,
                    color: !customPrompt.trim() || refining ? t.textDim : "#fff",
                    border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 10, fontFamily: "monospace",
                    cursor: !customPrompt.trim() || refining ? "not-allowed" : "pointer", whiteSpace: "nowrap",
                    display: "flex", alignItems: "center", gap: 6,
                  }}>
                  {refining === "Custom" ? <><Spinner color="#fff" /> Refining...</> : "Refine"}
                </button>
              </div>

              <div style={{ display: "flex", gap: 6, marginBottom: 10 }}>
                {TENSE_ACTIONS.map(makeRefineBtn)}
              </div>

              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {REFINE_PRESETS.map(makeRefineBtn)}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ── APP ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [dark, setDark] = useState(true);
  const t = getTheme(dark);

  return (
    <div style={{ background: t.bg, minHeight: "100vh", color: t.text, fontFamily: "'Syne', sans-serif", transition: "background 0.3s, color 0.3s" }}>
      <Head>
        <title>Media Intelligence Suite</title>
        <link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet" />
      </Head>
      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes slideUp { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        textarea, input { outline: none; transition: border-color 0.2s; }
        textarea:focus { border-color: rgba(128,128,128,0.35) !important; }
        input:focus { border-color: rgba(128,128,128,0.35) !important; }
        input::placeholder, textarea::placeholder { opacity: 0.4; }
        ::-webkit-scrollbar { width: 5px; }
        ::-webkit-scrollbar-thumb { background: rgba(128,128,128,0.2); border-radius: 3px; }
      `}</style>

      {/* ── NAV ── */}
      <div style={{ position: "sticky", top: 0, zIndex: 50, background: t.bgNav, backdropFilter: "blur(12px)", borderBottom: `1px solid ${t.border}`, padding: "0 20px" }}>
        <div style={{ maxWidth: 860, margin: "0 auto", display: "flex", alignItems: "center", gap: 10, height: 52 }}>
          <div style={{ width: 28, height: 28, borderRadius: 7, background: "linear-gradient(135deg, #8B5CF6, #6366F1)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 800, color: "#fff", flexShrink: 0 }}>M</div>
          <span style={{ fontSize: 11, fontWeight: 600, color: t.textDim, whiteSpace: "nowrap" }}>Media Intelligence Suite</span>
          <div style={{ flex: 1 }} />
          <button onClick={() => setDark(d => !d)} title={dark ? "Light mode" : "Dark mode"}
            style={{ display: "flex", alignItems: "center", gap: 6, background: t.toggleBg, border: `1px solid ${t.border}`, borderRadius: 20, padding: "5px 12px", cursor: "pointer", transition: "background 0.2s", flexShrink: 0 }}>
            <span style={{ fontSize: 13, lineHeight: 1 }}>{dark ? "☀️" : "🌙"}</span>
            <span style={{ fontSize: 9, color: t.textDim, fontFamily: "monospace", letterSpacing: "0.06em" }}>{dark ? "Light" : "Dark"}</span>
          </button>
        </div>
      </div>

      {/* ── MAIN ── */}
      <div style={{ maxWidth: 860, margin: "0 auto", padding: "36px 20px 64px", animation: "slideUp 0.3s ease" }}>
        <div style={{ marginBottom: 30 }}>
          <div style={{ fontSize: 9, fontFamily: "monospace", color: t.textDim, letterSpacing: "0.15em", marginBottom: 10, textTransform: "uppercase" }}>4 Specialist Tools</div>
          <h1 style={{ fontSize: "clamp(26px, 4.5vw, 38px)", fontWeight: 800, lineHeight: 1.05, letterSpacing: "-0.03em", marginBottom: 10 }}>
            <span style={{ color: t.heroTitle }}>Media Intelligence</span><br />
            <span style={{ background: "linear-gradient(90deg, #8B5CF6, #6366F1, #10B981)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", backgroundClip: "text" }}>at your fingertips</span>
          </h1>
          <p style={{ fontSize: 13, color: t.textDim, lineHeight: 1.65, maxWidth: 460 }}>Generate summaries, check quality, validate data, and extract insights from your media coverage.</p>
        </div>

        <ToolkitSection t={t} />
      </div>

      <div style={{ textAlign: "center", padding: "20px", fontSize: 9, color: t.textFoot, fontFamily: "monospace" }}>
        Media Intelligence Suite — Powered by Claude
      </div>
    </div>
  );
}
