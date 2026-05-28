import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ── Google Fonts ──────────────────────────────────────────────────────────────
const fontLink = document.createElement("link");
fontLink.rel = "stylesheet";
fontLink.href = "https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500;600&display=swap";
document.head.appendChild(fontLink);

// ── Constants ─────────────────────────────────────────────────────────────────
const STORAGE_KEY = "macro_tracker_v1";
const DEFAULT_TARGETS = { calories: 2800, protein: 220, carbs: 280, fat: 80 };
const MACRO_COLORS = { calories: "#f97316", protein: "#4ade80", carbs: "#60a5fa", fat: "#facc15" };
const todayKey = () => new Date().toISOString().slice(0, 10);

// ── API helper — reads key from Vite env in production ────────────────────────
const API_HEADERS = {
  "Content-Type": "application/json",
  "x-api-key": import.meta.env.VITE_ANTHROPIC_API_KEY,
  "anthropic-version": "2023-06-01",
  "anthropic-dangerous-direct-browser-access": "true",
};

function loadData() { try { return JSON.parse(localStorage.getItem(STORAGE_KEY)) || {}; } catch { return {}; } }
function saveData(data) { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); }
function sumMacros(entries) {
  return entries.reduce((acc, e) => ({
    calories: acc.calories + (e.calories || 0), protein: acc.protein + (e.protein || 0),
    carbs: acc.carbs + (e.carbs || 0), fat: acc.fat + (e.fat || 0),
  }), { calories: 0, protein: 0, carbs: 0, fat: 0 });
}

// ── Excel export ──────────────────────────────────────────────────────────────
function exportToExcel(allData, targets) {
  const days = Object.keys(allData).filter(k => !k.startsWith("__")).sort();

  // Sheet 1: Daily Summary
  const summaryRows = [
    ["Date", "Calories", "Cal Target", "Cal %", "Protein (g)", "Pro Target", "Carbs (g)", "Carb Target", "Fat (g)", "Fat Target", "Meals Logged"],
  ];
  days.forEach(day => {
    const entries = allData[day]?.entries || [];
    const t = sumMacros(entries);
    summaryRows.push([
      day,
      Math.round(t.calories), targets.calories, Math.round((t.calories / targets.calories) * 100) + "%",
      Math.round(t.protein), targets.protein,
      Math.round(t.carbs), targets.carbs,
      Math.round(t.fat), targets.fat,
      entries.length,
    ]);
  });

  // Sheet 2: All meals detail
  const mealRows = [
    ["Date", "Meal Name", "Calories", "Protein (g)", "Carbs (g)", "Fat (g)", "Source"],
  ];
  days.forEach(day => {
    (allData[day]?.entries || []).forEach(e => {
      mealRows.push([
        day, e.name,
        Math.round(e.calories || 0), Math.round(e.protein || 0),
        Math.round(e.carbs || 0), Math.round(e.fat || 0),
        e.source || "manual",
      ]);
    });
  });

  const wb = XLSX.utils.book_new();
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
  const wsMeals = XLSX.utils.aoa_to_sheet(mealRows);

  // Column widths
  wsSummary["!cols"] = [{ wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 12 }];
  wsMeals["!cols"] = [{ wch: 12 }, { wch: 30 }, { wch: 10 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }];

  XLSX.utils.book_append_sheet(wb, wsSummary, "Daily Summary");
  XLSX.utils.book_append_sheet(wb, wsMeals, "All Meals");

  const filename = `macro-history-${todayKey()}.xlsx`;
  XLSX.writeFile(wb, filename);
}

// ── Shared AI macro estimator ─────────────────────────────────────────────────
async function estimateMacrosFromText(description) {
  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: API_HEADERS,
    body: JSON.stringify({
      model: "claude-sonnet-4-5",
      max_tokens: 1000,
      system: `You are a nutrition expert. Given a food description, return ONLY a JSON object:
{"name":"concise food name","calories":number,"protein":number,"carbs":number,"fat":number,"confidence":"low|medium|high","note":"one-line note"}
No markdown, no backticks, no preamble. Raw JSON only.`,
      messages: [{ role: "user", content: `Estimate macros for: ${description}` }]
    })
  });
  const data = await response.json();
  if (data.error) throw new Error(data.error.message);
  const text = data.content?.find(b => b.type === "text")?.text || "";
  return JSON.parse(text.replace(/```json|```/g, "").trim());
}

// ── Ring ──────────────────────────────────────────────────────────────────────
function Ring({ value, max, color, size = 72, stroke = 6, label, sub }) {
  const r = (size - stroke * 2) / 2;
  const circ = 2 * Math.PI * r;
  const dash = circ * Math.min(value / max, 1);
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
      <svg width={size} height={size} style={{ transform: "rotate(-90deg)" }}>
        <circle cx={size/2} cy={size/2} r={r} fill="none" stroke="rgba(255,255,255,0.07)" strokeWidth={stroke} />
        <circle cx={size/2} cy={size/2} r={r} fill="none" stroke={color} strokeWidth={stroke}
          strokeDasharray={`${dash} ${circ}`} strokeLinecap="round"
          style={{ transition: "stroke-dasharray 0.5s ease" }} />
      </svg>
      <div style={{ textAlign: "center", marginTop: -size/2-2, marginBottom: size/2-10, position: "relative", zIndex: 1 }}>
        <div style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 18, color, lineHeight: 1 }}>{value}</div>
        <div style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 9, color: "rgba(255,255,255,0.4)", textTransform: "uppercase", letterSpacing: 1 }}>{sub}</div>
      </div>
      <div style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 11, color: "rgba(255,255,255,0.5)", textTransform: "uppercase", letterSpacing: 1 }}>{label}</div>
    </div>
  );
}

// ── MacroBar ──────────────────────────────────────────────────────────────────
function MacroBar({ label, value, max, color }) {
  const over = value > max;
  return (
    <div style={{ marginBottom: 10 }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
        <span style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 12, color: "rgba(255,255,255,0.5)", textTransform: "uppercase", letterSpacing: 1 }}>{label}</span>
        <span style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 14, color: over ? "#f87171" : color }}>
          {Math.round(value)}<span style={{ color: "rgba(255,255,255,0.25)", fontSize: 11 }}>/{max}g</span>
        </span>
      </div>
      <div style={{ height: 4, background: "rgba(255,255,255,0.07)", borderRadius: 2, overflow: "hidden" }}>
        <div style={{ height: "100%", width: `${Math.min((value/max)*100,100)}%`, background: over ? "#f87171" : color, borderRadius: 2, transition: "width 0.4s ease" }} />
      </div>
    </div>
  );
}

// ── EntryCard ─────────────────────────────────────────────────────────────────
function EntryCard({ entry, onDelete }) {
  const icon = entry.source === "voice" ? "🎤" : entry.source === "text" ? "💬" : "✏️";
  return (
    <div style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.08)", borderRadius: 12, padding: "12px 16px", display: "flex", alignItems: "center", gap: 12, animation: "slideIn 0.25s ease" }}>
      {entry.image
        ? <img src={entry.image} alt="" style={{ width: 48, height: 48, borderRadius: 8, objectFit: "cover", flexShrink: 0 }} />
        : <div style={{ width: 36, height: 36, borderRadius: 8, background: "rgba(255,255,255,0.05)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>{icon}</div>
      }
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 14, fontWeight: 500, color: "#fff", marginBottom: 3, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{entry.name}</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          {[{k:"calories",l:"kcal",c:MACRO_COLORS.calories},{k:"protein",l:"P",c:MACRO_COLORS.protein},{k:"carbs",l:"C",c:MACRO_COLORS.carbs},{k:"fat",l:"F",c:MACRO_COLORS.fat}].map(({k,l,c}) => (
            <span key={k} style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 11, color: c }}>
              {Math.round(entry[k]||0)}{l!=="kcal"?"g":""} <span style={{ color: "rgba(255,255,255,0.3)" }}>{l}</span>
            </span>
          ))}
        </div>
      </div>
      <button onClick={() => onDelete(entry.id)}
        style={{ background: "none", border: "none", cursor: "pointer", color: "rgba(255,255,255,0.2)", fontSize: 18, padding: 4, lineHeight: 1, flexShrink: 0, transition: "color 0.2s" }}
        onMouseEnter={e => e.target.style.color = "#f87171"}
        onMouseLeave={e => e.target.style.color = "rgba(255,255,255,0.2)"}>×</button>
    </div>
  );
}

// ── Text/Voice Modal ──────────────────────────────────────────────────────────
function TextVoiceModal({ onResult, onClose }) {
  const [text, setText] = useState("");
  const [listening, setListening] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [pulse, setPulse] = useState(false);
  const recognitionRef = useRef(null);
  const hasRecognition = typeof window !== "undefined" && ("SpeechRecognition" in window || "webkitSpeechRecognition" in window);

  function startListening() {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) { setError("Speech recognition not supported in this browser."); return; }
    const rec = new SR();
    rec.continuous = false; rec.interimResults = true; rec.lang = "en-US";
    rec.onstart = () => { setListening(true); setPulse(true); setError(""); };
    rec.onresult = e => setText(Array.from(e.results).map(r => r[0].transcript).join(""));
    rec.onerror = e => { setError("Mic error: " + e.error); setListening(false); setPulse(false); };
    rec.onend = () => { setListening(false); setPulse(false); };
    recognitionRef.current = rec;
    rec.start();
  }

  function stopListening() { recognitionRef.current?.stop(); setListening(false); setPulse(false); }

  async function analyse() {
    if (!text.trim()) return;
    setLoading(true); setError("");
    try { onResult(await estimateMacrosFromText(text.trim())); }
    catch(e) { setError(e.message || "Could not estimate macros. Try rephrasing or be more specific."); }
    setLoading(false);
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)", zIndex: 100, display: "flex", alignItems: "flex-end", justifyContent: "center", backdropFilter: "blur(4px)" }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{ background: "#141414", border: "1px solid rgba(255,255,255,0.12)", borderRadius: "20px 20px 0 0", padding: 24, width: "100%", maxWidth: 480, animation: "slideUp 0.3s ease" }}>
        <div style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 24, color: "#fff", marginBottom: 6 }}>DESCRIBE YOUR MEAL</div>
        <div style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 13, color: "rgba(255,255,255,0.35)", marginBottom: 20 }}>Type or speak naturally — e.g. "3 scrambled eggs, 2 toast, glass of OJ"</div>
        <div style={{ display: "flex", gap: 10, marginBottom: 14, alignItems: "flex-start" }}>
          <textarea value={text} onChange={e => setText(e.target.value)}
            placeholder="e.g. large chicken burrito with rice and beans…" rows={3}
            style={{ flex: 1, background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.12)", borderRadius: 12, padding: "12px 14px", color: "#fff", fontSize: 14, fontFamily: "'DM Sans',sans-serif", outline: "none", resize: "none", lineHeight: 1.5 }} />
          {hasRecognition && (
            <button onClick={listening ? stopListening : startListening}
              style={{ width: 52, height: 52, borderRadius: "50%", border: "none", cursor: "pointer", flexShrink: 0, marginTop: 4,
                background: listening ? "rgba(239,68,68,0.2)" : "rgba(255,255,255,0.08)",
                color: listening ? "#f87171" : "rgba(255,255,255,0.6)", fontSize: 22,
                display: "flex", alignItems: "center", justifyContent: "center",
                boxShadow: listening ? `0 0 0 ${pulse?"8px":"3px"} rgba(239,68,68,0.25)` : "none",
                transition: "all 0.3s ease" }}>
              {listening ? "⏹" : "🎤"}
            </button>
          )}
        </div>
        {listening && (
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14, padding: "8px 12px", background: "rgba(239,68,68,0.08)", border: "1px solid rgba(239,68,68,0.2)", borderRadius: 8 }}>
            <span style={{ width: 8, height: 8, borderRadius: "50%", background: "#f87171", display: "inline-block", animation: "micPulse 1s ease-in-out infinite" }} />
            <span style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 13, color: "#f87171" }}>Listening… speak your meal</span>
          </div>
        )}
        {error && <div style={{ color: "#f87171", fontSize: 13, marginBottom: 12, fontFamily: "'DM Sans',sans-serif" }}>{error}</div>}
        <div style={{ marginBottom: 16 }}>
          <div style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 11, color: "rgba(255,255,255,0.25)", textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>Quick examples</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
            {["200g chicken breast + rice", "protein shake 40g whey", "Big Mac meal", "2 eggs on toast", "Greek yogurt + banana"].map(ex => (
              <button key={ex} onClick={() => setText(ex)} style={{ background: "rgba(255,255,255,0.05)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 20, padding: "5px 12px", color: "rgba(255,255,255,0.5)", fontFamily: "'DM Sans',sans-serif", fontSize: 12, cursor: "pointer" }}>{ex}</button>
            ))}
          </div>
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={onClose} style={{ flex: 1, padding: "13px", background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 12, color: "rgba(255,255,255,0.6)", fontFamily: "'Bebas Neue',sans-serif", fontSize: 16, cursor: "pointer" }}>CANCEL</button>
          <button onClick={analyse} disabled={loading || !text.trim()}
            style={{ flex: 2, padding: "13px", background: (!text.trim()||loading) ? "rgba(96,165,250,0.3)" : "#60a5fa", border: "none", borderRadius: 12, color: "#000", fontFamily: "'Bebas Neue',sans-serif", fontSize: 18, cursor: (!text.trim()||loading) ? "not-allowed" : "pointer", letterSpacing: 1, transition: "background 0.2s" }}>
            {loading ? "ESTIMATING…" : "ESTIMATE MACROS"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── AI Photo Modal ────────────────────────────────────────────────────────────
function AIModal({ imageData, onResult, onClose }) {
  const [note, setNote] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function analyze() {
    setLoading(true); setError("");
    try {
      const prompt = note ? `Analyze this food photo. Context: "${note}".` : "Analyze this food photo.";
      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: API_HEADERS,
        body: JSON.stringify({
          model: "claude-sonnet-4-5", max_tokens: 1000,
          system: `You are a nutrition expert. Return ONLY JSON: {"name":"...","calories":0,"protein":0,"carbs":0,"fat":0,"confidence":"low|medium|high","note":"..."}. No markdown, no backticks.`,
          messages: [{ role: "user", content: [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: imageData.split(",")[1] } },
            { type: "text", text: prompt }
          ]}]
        })
      });
      const data = await response.json();
      const text = data.content?.find(b => b.type === "text")?.text || "";
      onResult(JSON.parse(text.replace(/```json|```/g, "").trim()), imageData);
    } catch(e) { setError(e.message || "Could not analyze image. Try again or enter manually."); }
    setLoading(false);
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)", zIndex: 100, display: "flex", alignItems: "flex-end", justifyContent: "center", backdropFilter: "blur(4px)" }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{ background: "#141414", border: "1px solid rgba(255,255,255,0.12)", borderRadius: "20px 20px 0 0", padding: 24, width: "100%", maxWidth: 480, animation: "slideUp 0.3s ease" }}>
        <div style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 24, color: "#fff", marginBottom: 16 }}>AI FOOD ANALYSIS</div>
        <img src={imageData} alt="" style={{ width: "100%", maxHeight: 200, objectFit: "cover", borderRadius: 12, marginBottom: 16 }} />
        <input value={note} onChange={e => setNote(e.target.value)} placeholder="Add context (e.g. 'large portion', '200g chicken')…"
          style={{ width: "100%", background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 10, padding: "10px 14px", color: "#fff", fontSize: 14, fontFamily: "'DM Sans',sans-serif", outline: "none", boxSizing: "border-box", marginBottom: 14 }} />
        {error && <div style={{ color: "#f87171", fontSize: 13, marginBottom: 12, fontFamily: "'DM Sans',sans-serif" }}>{error}</div>}
        <button onClick={analyze} disabled={loading}
          style={{ width: "100%", padding: "14px", background: loading ? "rgba(249,115,22,0.4)" : "#f97316", border: "none", borderRadius: 12, color: "#000", fontFamily: "'Bebas Neue',sans-serif", fontSize: 18, cursor: loading ? "not-allowed" : "pointer", letterSpacing: 1 }}>
          {loading ? "ANALYSING…" : "ANALYSE MEAL"}
        </button>
      </div>
    </div>
  );
}

// ── Confirm / Manual Entry Modal ──────────────────────────────────────────────
function ManualModal({ prefill, image, source, onSave, onClose }) {
  const [form, setForm] = useState({ name: prefill?.name||"", calories: prefill?.calories||"", protein: prefill?.protein||"", carbs: prefill?.carbs||"", fat: prefill?.fat||"" });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  function save() {
    if (!form.name || !form.calories) return;
    onSave({ id: Date.now(), name: form.name, calories: parseFloat(form.calories)||0, protein: parseFloat(form.protein)||0, carbs: parseFloat(form.carbs)||0, fat: parseFloat(form.fat)||0, image: image||null, source: source||"manual" });
    onClose();
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)", zIndex: 110, display: "flex", alignItems: "flex-end", justifyContent: "center", backdropFilter: "blur(4px)" }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{ background: "#141414", border: "1px solid rgba(255,255,255,0.12)", borderRadius: "20px 20px 0 0", padding: 24, width: "100%", maxWidth: 480, animation: "slideUp 0.3s ease", maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 24, color: "#fff", marginBottom: 16 }}>{prefill ? "CONFIRM MACROS" : "ADD FOOD"}</div>
        {prefill?.note && (
          <div style={{ background: "rgba(96,165,250,0.08)", border: "1px solid rgba(96,165,250,0.2)", borderRadius: 8, padding: "8px 12px", marginBottom: 14, fontFamily: "'DM Sans',sans-serif", fontSize: 12, color: "rgba(255,255,255,0.6)" }}>
            🤖 {prefill.note} <span style={{ color: "rgba(255,255,255,0.3)" }}>({prefill.confidence} confidence)</span>
          </div>
        )}
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 11, color: "rgba(255,255,255,0.4)", textTransform: "uppercase", letterSpacing: 1, marginBottom: 4 }}>Food Name</div>
          <input value={form.name} onChange={e => set("name", e.target.value)} placeholder="e.g. Chicken & Rice"
            style={{ width: "100%", background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 10, padding: "10px 14px", color: "#fff", fontSize: 14, fontFamily: "'DM Sans',sans-serif", outline: "none", boxSizing: "border-box" }} />
        </div>
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 11, color: "rgba(255,255,255,0.4)", textTransform: "uppercase", letterSpacing: 1, marginBottom: 4 }}>Calories (kcal)</div>
          <input value={form.calories} onChange={e => set("calories", e.target.value)} type="number" placeholder="0"
            style={{ width: "100%", background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 10, padding: "10px 14px", color: "#fff", fontSize: 14, fontFamily: "'DM Sans',sans-serif", outline: "none", boxSizing: "border-box" }} />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 16 }}>
          {[["Protein","protein"],["Carbs","carbs"],["Fat","fat"]].map(([l,k]) => (
            <div key={k}>
              <div style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 11, color: "rgba(255,255,255,0.4)", textTransform: "uppercase", letterSpacing: 1, marginBottom: 4 }}>{l} (g)</div>
              <input value={form[k]} onChange={e => set(k, e.target.value)} type="number" placeholder="0"
                style={{ width: "100%", background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 10, padding: "10px 10px", color: "#fff", fontSize: 14, fontFamily: "'DM Sans',sans-serif", outline: "none", boxSizing: "border-box" }} />
            </div>
          ))}
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={onClose} style={{ flex: 1, padding: "13px", background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 12, color: "rgba(255,255,255,0.6)", fontFamily: "'Bebas Neue',sans-serif", fontSize: 16, cursor: "pointer" }}>CANCEL</button>
          <button onClick={save} style={{ flex: 2, padding: "13px", background: "#4ade80", border: "none", borderRadius: 12, color: "#000", fontFamily: "'Bebas Neue',sans-serif", fontSize: 18, cursor: "pointer", letterSpacing: 1 }}>LOG MEAL</button>
        </div>
      </div>
    </div>
  );
}

// ── Targets Modal ─────────────────────────────────────────────────────────────
function TargetsModal({ targets, onSave, onClose }) {
  const [form, setForm] = useState({ ...targets });
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)", zIndex: 100, display: "flex", alignItems: "flex-end", justifyContent: "center", backdropFilter: "blur(4px)" }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{ background: "#141414", border: "1px solid rgba(255,255,255,0.12)", borderRadius: "20px 20px 0 0", padding: 24, width: "100%", maxWidth: 480, animation: "slideUp 0.3s ease" }}>
        <div style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 24, color: "#fff", marginBottom: 16 }}>DAILY TARGETS</div>
        {[{k:"calories",l:"Calories (kcal)",c:MACRO_COLORS.calories},{k:"protein",l:"Protein (g)",c:MACRO_COLORS.protein},{k:"carbs",l:"Carbs (g)",c:MACRO_COLORS.carbs},{k:"fat",l:"Fat (g)",c:MACRO_COLORS.fat}].map(({k,l,c}) => (
          <div key={k} style={{ marginBottom: 14 }}>
            <div style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 11, color: c, textTransform: "uppercase", letterSpacing: 1, marginBottom: 4 }}>{l}</div>
            <input type="number" value={form[k]} onChange={e => setForm(f => ({ ...f, [k]: parseInt(e.target.value)||0 }))}
              style={{ width: "100%", background: "rgba(255,255,255,0.06)", border: `1px solid ${c}30`, borderRadius: 10, padding: "10px 14px", color: "#fff", fontSize: 16, fontFamily: "'Bebas Neue',sans-serif", outline: "none", boxSizing: "border-box" }} />
          </div>
        ))}
        <button onClick={() => { onSave(form); onClose(); }}
          style={{ width: "100%", padding: "14px", background: "#f97316", border: "none", borderRadius: 12, color: "#000", fontFamily: "'Bebas Neue',sans-serif", fontSize: 18, cursor: "pointer", letterSpacing: 1 }}>SAVE TARGETS</button>
      </div>
    </div>
  );
}

// ── History View ──────────────────────────────────────────────────────────────
function HistoryView({ allData, targets, onExport }) {
  const days = Object.keys(allData).filter(k => !k.startsWith("__")).sort((a,b) => b.localeCompare(a)).slice(0, 30);
  return (
    <div style={{ padding: "0 16px 100px" }}>
      {/* Export button */}
      <button onClick={onExport}
        style={{ width: "100%", marginBottom: 20, padding: "13px", background: "rgba(74,222,128,0.1)", border: "1px solid rgba(74,222,128,0.25)", borderRadius: 12, color: "#4ade80", fontFamily: "'Bebas Neue',sans-serif", fontSize: 16, letterSpacing: 1, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
        <span>⬇</span> EXPORT TO EXCEL
      </button>

      {days.length === 0
        ? <div style={{ textAlign: "center", padding: 60, color: "rgba(255,255,255,0.3)", fontFamily: "'DM Sans',sans-serif" }}>No history yet. Start logging meals!</div>
        : days.map(day => {
          const entries = allData[day]?.entries || [];
          const t = sumMacros(entries);
          const isToday = day === todayKey();
          return (
            <div key={day} style={{ marginBottom: 16 }}>
              <div style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 13, color: "rgba(255,255,255,0.4)", letterSpacing: 2, marginBottom: 8 }}>
                {isToday ? "TODAY" : new Date(day + "T12:00:00").toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric" }).toUpperCase()}
              </div>
              <div style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.08)", borderRadius: 12, padding: "14px 16px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                  <span style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 22, color: MACRO_COLORS.calories }}>
                    {Math.round(t.calories)} <span style={{ fontSize: 13, color: "rgba(255,255,255,0.3)" }}>/ {targets.calories} kcal</span>
                  </span>
                  <span style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 12, color: "rgba(255,255,255,0.3)" }}>{entries.length} meal{entries.length !== 1 ? "s" : ""}</span>
                </div>
                <div style={{ display: "flex", gap: 16 }}>
                  {[{k:"protein",l:"Protein"},{k:"carbs",l:"Carbs"},{k:"fat",l:"Fat"}].map(({k,l}) => (
                    <div key={k}>
                      <div style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 15, color: MACRO_COLORS[k] }}>{Math.round(t[k])}g</div>
                      <div style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 10, color: "rgba(255,255,255,0.3)", textTransform: "uppercase", letterSpacing: 1 }}>{l}</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          );
        })
      }
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function MacroTracker() {
  const [tab, setTab] = useState("today");
  const [allData, setAllData] = useState(() => loadData());
  const [targets, setTargets] = useState(() => { const d = loadData(); return d.__targets || DEFAULT_TARGETS; });

  const [showAI, setShowAI] = useState(false);
  const [showTextVoice, setShowTextVoice] = useState(false);
  const [showManual, setShowManual] = useState(false);
  const [showTargets, setShowTargets] = useState(false);
  const [pendingImage, setPendingImage] = useState(null);
  const [aiFill, setAiFill] = useState(null);
  const [pendingSource, setPendingSource] = useState("manual");
  const [showAddMenu, setShowAddMenu] = useState(false);

  const fileRef = useRef();
  const cameraRef = useRef();

  const today = todayKey();
  const todayEntries = allData[today]?.entries || [];
  const totals = sumMacros(todayEntries);

  function persist(newAll) { setAllData(newAll); saveData({ ...newAll, __targets: targets }); }
  function persistTargets(t) { setTargets(t); saveData({ ...allData, __targets: t }); }
  function addEntry(entry) { persist({ ...allData, [today]: { entries: [...todayEntries, entry] } }); }
  function deleteEntry(id) { persist({ ...allData, [today]: { entries: todayEntries.filter(e => e.id !== id) } }); }

  function handleImage(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => { setPendingImage(e.target.result); setShowAI(true); setShowAddMenu(false); };
    reader.readAsDataURL(file);
  }

  function handleAIPhotoResult(result, image) { setAiFill({ ...result, image }); setPendingSource("photo"); setShowAI(false); setShowManual(true); }
  function handleTextVoiceResult(result) { setAiFill(result); setPendingSource("text"); setShowTextVoice(false); setShowManual(true); }

  const remaining = { calories: targets.calories - totals.calories, protein: targets.protein - totals.protein, carbs: targets.carbs - totals.carbs, fat: targets.fat - totals.fat };
  const calPct = Math.round((totals.calories / targets.calories) * 100);

  useEffect(() => {
    const style = document.createElement("style");
    style.textContent = `
      * { box-sizing:border-box; -webkit-tap-highlight-color:transparent; }
      @keyframes slideIn  { from{opacity:0;transform:translateY(8px)} to{opacity:1;transform:none} }
      @keyframes slideUp  { from{transform:translateY(100%)} to{transform:none} }
      @keyframes fadeIn   { from{opacity:0} to{opacity:1} }
      @keyframes micPulse { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:0.5;transform:scale(1.3)} }
      input[type=number]::-webkit-outer-spin-button,input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none;}
      input::placeholder,textarea::placeholder{color:rgba(255,255,255,0.2)!important;}
      ::-webkit-scrollbar{width:0;}
      textarea{resize:none;}
    `;
    document.head.appendChild(style);
  }, []);

  return (
    <div style={{ background: "#0c0c0c", minHeight: "100vh", color: "#fff", fontFamily: "'DM Sans',sans-serif", maxWidth: 480, margin: "0 auto", position: "relative" }}>

      {/* Header */}
      <div style={{ padding: "20px 20px 0", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <div style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 32, letterSpacing: 2, lineHeight: 1 }}>MACRO</div>
          <div style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 32, letterSpacing: 2, color: "#f97316", lineHeight: 1 }}>TRACKER</div>
        </div>
        <button onClick={() => setShowTargets(true)}
          style={{ background: "rgba(255,255,255,0.06)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 10, padding: "8px 14px", color: "rgba(255,255,255,0.6)", fontFamily: "'DM Sans',sans-serif", fontSize: 12, cursor: "pointer" }}>
          ⚙ Targets
        </button>
      </div>

      {/* Tabs */}
      <div style={{ display: "flex", padding: "16px 20px 0", gap: 8 }}>
        {["today", "history"].map(t => (
          <button key={t} onClick={() => setTab(t)}
            style={{ flex: 1, padding: "10px", borderRadius: 10, background: tab===t ? "rgba(249,115,22,0.15)" : "rgba(255,255,255,0.04)", border: tab===t ? "1px solid rgba(249,115,22,0.3)" : "1px solid rgba(255,255,255,0.08)", color: tab===t ? "#f97316" : "rgba(255,255,255,0.4)", fontFamily: "'Bebas Neue',sans-serif", fontSize: 15, letterSpacing: 1, cursor: "pointer" }}>
            {t.toUpperCase()}
          </button>
        ))}
      </div>

      {tab === "history" ? (
        <div style={{ marginTop: 20 }}>
          <HistoryView allData={allData} targets={targets} onExport={() => exportToExcel(allData, targets)} />
        </div>
      ) : (
        <>
          {/* Calorie ring */}
          <div style={{ padding: "24px 20px 0", textAlign: "center" }}>
            <div style={{ position: "relative", display: "inline-block" }}>
              <svg width={160} height={160} style={{ transform: "rotate(-90deg)" }}>
                <circle cx={80} cy={80} r={68} fill="none" stroke="rgba(255,255,255,0.05)" strokeWidth={10} />
                <circle cx={80} cy={80} r={68} fill="none" stroke={totals.calories > targets.calories ? "#f87171" : "#f97316"} strokeWidth={10}
                  strokeDasharray={`${Math.min(totals.calories/targets.calories,1)*2*Math.PI*68} ${2*Math.PI*68}`}
                  strokeLinecap="round" style={{ transition: "stroke-dasharray 0.6s ease" }} />
              </svg>
              <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
                <div style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 38, color: "#f97316", lineHeight: 1 }}>{Math.round(totals.calories)}</div>
                <div style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 11, color: "rgba(255,255,255,0.4)", textTransform: "uppercase", letterSpacing: 1 }}>of {targets.calories}</div>
                <div style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 12, color: "rgba(255,255,255,0.25)", marginTop: 2 }}>{calPct}%</div>
              </div>
            </div>
            <div style={{ marginTop: 8 }}>
              {remaining.calories >= 0
                ? <span style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 13, color: "rgba(255,255,255,0.4)" }}><span style={{ color: "#f97316", fontWeight: 600 }}>{Math.round(remaining.calories)}</span> kcal remaining</span>
                : <span style={{ fontFamily: "'DM Sans',sans-serif", fontSize: 13, color: "#f87171" }}>{Math.abs(Math.round(remaining.calories))} kcal over target</span>
              }
            </div>
          </div>

          {/* Macro rings */}
          <div style={{ display: "flex", justifyContent: "space-around", padding: "20px 20px 0" }}>
            <Ring value={Math.round(totals.protein)} max={targets.protein} color={MACRO_COLORS.protein} label="Protein" sub={`/${targets.protein}g`} />
            <Ring value={Math.round(totals.carbs)} max={targets.carbs} color={MACRO_COLORS.carbs} label="Carbs" sub={`/${targets.carbs}g`} />
            <Ring value={Math.round(totals.fat)} max={targets.fat} color={MACRO_COLORS.fat} label="Fat" sub={`/${targets.fat}g`} />
          </div>

          {/* Macro bars */}
          <div style={{ padding: "20px 20px 0" }}>
            <MacroBar label="Protein" value={totals.protein} max={targets.protein} color={MACRO_COLORS.protein} />
            <MacroBar label="Carbs" value={totals.carbs} max={targets.carbs} color={MACRO_COLORS.carbs} />
            <MacroBar label="Fat" value={totals.fat} max={targets.fat} color={MACRO_COLORS.fat} />
          </div>

          {/* Log */}
          <div style={{ padding: "20px 20px 100px" }}>
            <div style={{ fontFamily: "'Bebas Neue',sans-serif", fontSize: 16, letterSpacing: 2, color: "rgba(255,255,255,0.4)", marginBottom: 12 }}>
              TODAY'S LOG — {todayEntries.length} MEAL{todayEntries.length !== 1 ? "S" : ""}
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {todayEntries.length === 0 && <div style={{ textAlign: "center", padding: "32px 0", color: "rgba(255,255,255,0.2)", fontSize: 14 }}>No meals logged yet. Tap + to add one.</div>}
              {todayEntries.map(e => <EntryCard key={e.id} entry={e} onDelete={deleteEntry} />)}
            </div>
          </div>
        </>
      )}

      {/* FAB */}
      <div style={{ position: "fixed", bottom: 28, right: 20, zIndex: 50 }}>
        {showAddMenu && (
          <div style={{ position: "absolute", bottom: 64, right: 0, display: "flex", flexDirection: "column", gap: 10, alignItems: "flex-end", animation: "fadeIn 0.2s ease" }}>
            {[
              { label: "📷 Take Photo",    action: () => { cameraRef.current.click(); setShowAddMenu(false); } },
              { label: "🖼 Upload Photo",   action: () => { fileRef.current.click(); setShowAddMenu(false); } },
              { label: "🎤 Speak / Type",  action: () => { setShowTextVoice(true); setShowAddMenu(false); } },
              { label: "✏️ Manual Entry",   action: () => { setAiFill(null); setPendingImage(null); setShowManual(true); setShowAddMenu(false); } },
            ].map(({ label, action }) => (
              <button key={label} onClick={action}
                style={{ background: "#1a1a1a", border: "1px solid rgba(255,255,255,0.15)", borderRadius: 24, padding: "10px 18px", color: "#fff", fontFamily: "'DM Sans',sans-serif", fontSize: 14, cursor: "pointer", whiteSpace: "nowrap", boxShadow: "0 4px 20px rgba(0,0,0,0.5)" }}>
                {label}
              </button>
            ))}
          </div>
        )}
        <button onClick={() => setShowAddMenu(m => !m)}
          style={{ width: 56, height: 56, borderRadius: "50%", background: showAddMenu ? "#fff" : "#f97316", border: "none", cursor: "pointer", fontSize: 26, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 4px 24px rgba(249,115,22,0.4)", transition: "background 0.2s,transform 0.2s", transform: showAddMenu ? "rotate(45deg)" : "none", color: "#000" }}>+</button>
      </div>

      {/* Hidden file inputs */}
      <input ref={fileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={e => handleImage(e.target.files[0])} />
      <input ref={cameraRef} type="file" accept="image/*" capture="environment" style={{ display: "none" }} onChange={e => handleImage(e.target.files[0])} />

      {/* Modals */}
      {showAI && pendingImage && <AIModal imageData={pendingImage} onResult={handleAIPhotoResult} onClose={() => { setShowAI(false); setPendingImage(null); }} />}
      {showTextVoice && <TextVoiceModal onResult={handleTextVoiceResult} onClose={() => setShowTextVoice(false)} />}
      {showManual && <ManualModal prefill={aiFill} image={aiFill?.image || pendingImage} source={pendingSource} onSave={addEntry} onClose={() => { setShowManual(false); setAiFill(null); setPendingImage(null); }} />}
      {showTargets && <TargetsModal targets={targets} onSave={persistTargets} onClose={() => setShowTargets(false)} />}
      {showAddMenu && <div style={{ position: "fixed", inset: 0, zIndex: 40 }} onClick={() => setShowAddMenu(false)} />}
    </div>
  );
}