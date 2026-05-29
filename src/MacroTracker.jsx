import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

const fontLink = document.createElement("link");
fontLink.rel = "stylesheet";
fontLink.href = "https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500;600&display=swap";
document.head.appendChild(fontLink);

const STORAGE_KEY = "macro_tracker_v1";
const DEFAULT_TARGETS_WORKOUT = { calories: 2800, protein: 220, carbs: 280, fat: 80 };
const DEFAULT_TARGETS_REST    = { calories: 2300, protein: 220, carbs: 180, fat: 80 };
const MACRO_COLORS = { calories: "#f97316", protein: "#4ade80", carbs: "#60a5fa", fat: "#facc15" };
const todayKey = () => new Date().toISOString().slice(0, 10);

const API_HEADERS = {
  "Content-Type": "application/json",
  "x-api-key": import.meta.env.VITE_ANTHROPIC_API_KEY,
  "anthropic-version": "2023-06-01",
  "anthropic-dangerous-direct-browser-access": "true",
};

function migrateData(raw) {
  // Migrate any day that has entries as a plain array (old format) to new format
  const migrated = { ...raw };
  Object.keys(migrated).forEach(k => {
    if (k.startsWith("__")) return;
    const day = migrated[k];
    // Old format: { entries: [...] } — missing dayType
    if (day && day.entries && !day.dayType) {
      migrated[k] = { entries: day.entries, dayType: "workout" };
    }
    // Very old format: direct array
    if (Array.isArray(day)) {
      migrated[k] = { entries: day, dayType: "workout" };
    }
  });
  return migrated;
}
function loadData() {
  try {
    const raw = JSON.parse(localStorage.getItem(STORAGE_KEY)) || {};
    const migrated = stripImages(migrateData(raw));
    localStorage.setItem(STORAGE_KEY, JSON.stringify(migrated));
    return migrated;
  } catch { return {}; }
}
function saveData(data) {
  const json = JSON.stringify(data);
  localStorage.setItem(STORAGE_KEY, json);
  // Always keep a rolling backup of the last known good save
  localStorage.setItem(STORAGE_KEY + "_backup", json);
}

function stripImages(data) {
  const cleaned = {};
  Object.keys(data).forEach(k => {
    if (k.startsWith("__")) { cleaned[k] = data[k]; return; }
    const day = data[k];
    if (day?.entries) {
      cleaned[k] = { ...day, entries: day.entries.map(e => ({ ...e, image: null })) };
    } else {
      cleaned[k] = day;
    }
  });
  return cleaned;
}

function backupData(allData, targetsWorkout, targetsRest) {
  const payload = stripImages({ ...allData, __targetsWorkout: targetsWorkout, __targetsRest: targetsRest, __exportedAt: new Date().toISOString() });
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `macro-backup-${todayKey()}.json`;
  a.click();
}

function restoreData(file, onSuccess, onError) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const parsed = JSON.parse(e.target.result);
      // Validate it looks like macro data
      const hasData = Object.keys(parsed).some(k => !k.startsWith("__"));
      if (!hasData) throw new Error("File doesn't look like a macro backup");
      const migrated = migrateData(stripImages(parsed));
      localStorage.setItem(STORAGE_KEY, JSON.stringify(migrated));
      onSuccess(migrated);
    } catch(err) { onError(err.message); }
  };
  reader.readAsText(file);
}

function sumMacros(entries) {
  return entries.reduce((acc, e) => ({
    calories: acc.calories + (e.calories || 0), protein: acc.protein + (e.protein || 0),
    carbs: acc.carbs + (e.carbs || 0), fat: acc.fat + (e.fat || 0),
  }), { calories: 0, protein: 0, carbs: 0, fat: 0 });
}

function exportToExcel(allData, targetsWorkout, targetsRest) {
  const days = Object.keys(allData).filter(k => !k.startsWith("__")).sort();
  const summaryRows = [["Date", "Day Type", "Calories", "Cal Target", "Cal %", "Protein (g)", "Pro Target", "Carbs (g)", "Carb Target", "Fat (g)", "Fat Target", "Meals"]];
  days.forEach(day => {
    const entries = allData[day]?.entries || [];
    const dayType = allData[day]?.dayType || "workout";
    const tgt = dayType === "rest" ? targetsRest : targetsWorkout;
    const t = sumMacros(entries);
    summaryRows.push([day, dayType, Math.round(t.calories), tgt.calories, Math.round((t.calories/tgt.calories)*100)+"%", Math.round(t.protein), tgt.protein, Math.round(t.carbs), tgt.carbs, Math.round(t.fat), tgt.fat, entries.length]);
  });
  const mealRows = [["Date", "Day Type", "Meal Name", "Calories", "Protein (g)", "Carbs (g)", "Fat (g)", "Source"]];
  days.forEach(day => {
    const dayType = allData[day]?.dayType || "workout";
    (allData[day]?.entries || []).forEach(e => {
      mealRows.push([day, dayType, e.name, Math.round(e.calories||0), Math.round(e.protein||0), Math.round(e.carbs||0), Math.round(e.fat||0), e.source||"manual"]);
    });
  });
  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.aoa_to_sheet(summaryRows);
  const ws2 = XLSX.utils.aoa_to_sheet(mealRows);
  ws1["!cols"] = [{wch:12},{wch:10},{wch:10},{wch:10},{wch:8},{wch:12},{wch:10},{wch:10},{wch:10},{wch:10},{wch:10},{wch:8}];
  ws2["!cols"] = [{wch:12},{wch:10},{wch:30},{wch:10},{wch:12},{wch:10},{wch:10},{wch:10}];
  XLSX.utils.book_append_sheet(wb, ws1, "Daily Summary");
  XLSX.utils.book_append_sheet(wb, ws2, "All Meals");
  XLSX.writeFile(wb, `macro-history-${todayKey()}.xlsx`);
}

async function estimateMacrosFromText(description) {
  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST", headers: API_HEADERS,
    body: JSON.stringify({
      model: "claude-sonnet-4-5", max_tokens: 1000,
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

function Ring({ value, max, color, size=88, stroke=6, label, sub, over=false }) {
  // Ring fills as you approach target — remaining fills the arc
  const pct = Math.min(value / max, 1);
  const r=(size-stroke*2)/2, circ=2*Math.PI*r, dash=circ*pct;
  const displayColor = over ? "#f87171" : color;
  return (
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:6}}>
      <div style={{position:"relative",width:size,height:size}}>
        <svg width={size} height={size} style={{transform:"rotate(-90deg)"}}>
          <circle cx={size/2} cy={size/2} r={r} fill="none" stroke="rgba(255,255,255,0.07)" strokeWidth={stroke}/>
          <circle cx={size/2} cy={size/2} r={r} fill="none" stroke={displayColor} strokeWidth={stroke} strokeDasharray={`${dash} ${circ}`} strokeLinecap="round" style={{transition:"stroke-dasharray 0.5s ease"}}/>
        </svg>
        <div style={{position:"absolute",inset:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:displayColor,lineHeight:1}}>{over ? "0" : value}</div>
          <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:9,color:"rgba(255,255,255,0.4)",textTransform:"uppercase",letterSpacing:0.5}}>{sub}</div>
        </div>
      </div>
      <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:"rgba(255,255,255,0.5)",textTransform:"uppercase",letterSpacing:1}}>{label}</div>
    </div>
  );
}

function MacroBar({ label, value, max, color }) {
  const over=value>max;
  return (
    <div style={{marginBottom:10}}>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
        <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:12,color:"rgba(255,255,255,0.5)",textTransform:"uppercase",letterSpacing:1}}>{label}</span>
        <span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:14,color:over?"#f87171":color}}>{Math.round(value)}<span style={{color:"rgba(255,255,255,0.25)",fontSize:11}}>/{max}g</span></span>
      </div>
      <div style={{height:4,background:"rgba(255,255,255,0.07)",borderRadius:2,overflow:"hidden"}}>
        <div style={{height:"100%",width:`${Math.min((value/max)*100,100)}%`,background:over?"#f87171":color,borderRadius:2,transition:"width 0.4s ease"}}/>
      </div>
    </div>
  );
}

function EntryCard({ entry, onDelete, onEdit }) {
  const icon=entry.source==="voice"?"🎤":entry.source==="text"?"💬":entry.source==="photo"?"📷":"✏️";
  const timeStr = entry.timestamp ? new Date(entry.timestamp).toLocaleTimeString("en-US",{hour:"numeric",minute:"2-digit",hour12:true}) : "";
  return (
    <div style={{background:"rgba(255,255,255,0.04)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:12,padding:"12px 16px",display:"flex",alignItems:"center",gap:12,animation:"slideIn 0.25s ease"}}>
      <div style={{width:36,height:36,borderRadius:8,background:"rgba(255,255,255,0.05)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>{icon}</div>
      <div style={{flex:1,minWidth:0}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:2}}>
          <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:14,fontWeight:500,color:"#fff",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",flex:1,marginRight:8}}>{entry.name}</div>
          {timeStr&&<div style={{fontFamily:"'DM Sans',sans-serif",fontSize:10,color:"rgba(255,255,255,0.3)",flexShrink:0}}>{timeStr}</div>}
        </div>
        <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
          {[{k:"calories",l:"kcal",c:MACRO_COLORS.calories},{k:"protein",l:"P",c:MACRO_COLORS.protein},{k:"carbs",l:"C",c:MACRO_COLORS.carbs},{k:"fat",l:"F",c:MACRO_COLORS.fat}].map(({k,l,c})=>(
            <span key={k} style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:c}}>{Math.round(entry[k]||0)}{l!=="kcal"?"g":""} <span style={{color:"rgba(255,255,255,0.3)"}}>{l}</span></span>
          ))}
        </div>
      </div>
      <div style={{display:"flex",gap:4,flexShrink:0}}>
        <button onClick={()=>onEdit(entry)} style={{background:"none",border:"none",cursor:"pointer",color:"rgba(255,255,255,0.2)",fontSize:14,padding:4,lineHeight:1,transition:"color 0.2s"}} onMouseEnter={e=>e.target.style.color="#60a5fa"} onMouseLeave={e=>e.target.style.color="rgba(255,255,255,0.2)"}>✎</button>
        <button onClick={()=>onDelete(entry.id)} style={{background:"none",border:"none",cursor:"pointer",color:"rgba(255,255,255,0.2)",fontSize:18,padding:4,lineHeight:1,transition:"color 0.2s"}} onMouseEnter={e=>e.target.style.color="#f87171"} onMouseLeave={e=>e.target.style.color="rgba(255,255,255,0.2)"}>×</button>
      </div>
    </div>
  );
}

function TextVoiceModal({ onResult, onClose }) {
  const [text,setText]=useState(""); const [listening,setListening]=useState(false); const [loading,setLoading]=useState(false); const [error,setError]=useState(""); const [pulse,setPulse]=useState(false);
  const recognitionRef=useRef(null);
  const hasRecognition=typeof window!=="undefined"&&("SpeechRecognition"in window||"webkitSpeechRecognition"in window);
  function startListening(){const SR=window.SpeechRecognition||window.webkitSpeechRecognition;if(!SR){setError("Speech recognition not supported.");return;}const rec=new SR();rec.continuous=false;rec.interimResults=true;rec.lang="en-US";rec.onstart=()=>{setListening(true);setPulse(true);setError("");};rec.onresult=e=>setText(Array.from(e.results).map(r=>r[0].transcript).join(""));rec.onerror=e=>{setError("Mic error: "+e.error);setListening(false);setPulse(false);};rec.onend=()=>{setListening(false);setPulse(false);};recognitionRef.current=rec;rec.start();}
  function stopListening(){recognitionRef.current?.stop();setListening(false);setPulse(false);}
  async function analyse(){if(!text.trim())return;setLoading(true);setError("");try{onResult(await estimateMacrosFromText(text.trim()));}catch(e){setError(e.message||"Could not estimate macros.");}setLoading(false);}
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",zIndex:100,display:"flex",alignItems:"flex-end",justifyContent:"center",backdropFilter:"blur(4px)"}} onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#141414",border:"1px solid rgba(255,255,255,0.12)",borderRadius:"20px 20px 0 0",padding:24,width:"100%",maxWidth:480,animation:"slideUp 0.3s ease"}}>
        <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:24,color:"#fff",marginBottom:6}}>DESCRIBE YOUR MEAL</div>
        <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:13,color:"rgba(255,255,255,0.35)",marginBottom:20}}>Type or speak naturally — e.g. "3 scrambled eggs, 2 toast, OJ"</div>
        <div style={{display:"flex",gap:10,marginBottom:14,alignItems:"flex-start"}}>
          <textarea value={text} onChange={e=>setText(e.target.value)} placeholder="e.g. large chicken burrito with rice and beans…" rows={3} style={{flex:1,background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.12)",borderRadius:12,padding:"12px 14px",color:"#fff",fontSize:14,fontFamily:"'DM Sans',sans-serif",outline:"none",resize:"none",lineHeight:1.5}}/>
          {hasRecognition&&<button onClick={listening?stopListening:startListening} style={{width:52,height:52,borderRadius:"50%",border:"none",cursor:"pointer",flexShrink:0,marginTop:4,background:listening?"rgba(239,68,68,0.2)":"rgba(255,255,255,0.08)",color:listening?"#f87171":"rgba(255,255,255,0.6)",fontSize:22,display:"flex",alignItems:"center",justifyContent:"center",boxShadow:listening?`0 0 0 ${pulse?"8px":"3px"} rgba(239,68,68,0.25)`:"none",transition:"all 0.3s ease"}}>{listening?"⏹":"🎤"}</button>}
        </div>
        {listening&&<div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14,padding:"8px 12px",background:"rgba(239,68,68,0.08)",border:"1px solid rgba(239,68,68,0.2)",borderRadius:8}}><span style={{width:8,height:8,borderRadius:"50%",background:"#f87171",display:"inline-block",animation:"micPulse 1s ease-in-out infinite"}}/><span style={{fontFamily:"'DM Sans',sans-serif",fontSize:13,color:"#f87171"}}>Listening… speak your meal</span></div>}
        {error&&<div style={{color:"#f87171",fontSize:13,marginBottom:12,fontFamily:"'DM Sans',sans-serif"}}>{error}</div>}
        <div style={{marginBottom:16}}>
          <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:"rgba(255,255,255,0.25)",textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>Quick examples</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
            {["200g chicken breast + rice","protein shake 40g whey","Big Mac meal","2 eggs on toast","Greek yogurt + banana"].map(ex=>(
              <button key={ex} onClick={()=>setText(ex)} style={{background:"rgba(255,255,255,0.05)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:20,padding:"5px 12px",color:"rgba(255,255,255,0.5)",fontFamily:"'DM Sans',sans-serif",fontSize:12,cursor:"pointer"}}>{ex}</button>
            ))}
          </div>
        </div>
        <div style={{display:"flex",gap:10}}>
          <button onClick={onClose} style={{flex:1,padding:"13px",background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:12,color:"rgba(255,255,255,0.6)",fontFamily:"'Bebas Neue',sans-serif",fontSize:16,cursor:"pointer"}}>CANCEL</button>
          <button onClick={analyse} disabled={loading||!text.trim()} style={{flex:2,padding:"13px",background:(!text.trim()||loading)?"rgba(96,165,250,0.3)":"#60a5fa",border:"none",borderRadius:12,color:"#000",fontFamily:"'Bebas Neue',sans-serif",fontSize:18,cursor:(!text.trim()||loading)?"not-allowed":"pointer",letterSpacing:1}}>{loading?"ESTIMATING…":"ESTIMATE MACROS"}</button>
        </div>
      </div>
    </div>
  );
}

function AIModal({ imageData, onResult, onClose }) {
  const [note,setNote]=useState(""); const [loading,setLoading]=useState(false); const [error,setError]=useState("");
  async function analyze(){
    setLoading(true);setError("");
    try{
      const prompt=note?`Analyze this food photo. Context: "${note}".`:"Analyze this food photo.";
      const response=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:API_HEADERS,body:JSON.stringify({model:"claude-sonnet-4-5",max_tokens:300,system:`You are a nutrition expert. Return ONLY a single JSON object: {"name":"food name","calories":0,"protein":0,"carbs":0,"fat":0,"confidence":"low|medium|high","note":"brief note"}`,messages:[{role:"user",content:[{type:"image",source:{type:"base64",media_type:"image/jpeg",data:imageData.split(",")[1]}},{type:"text",text:prompt}]}]})});
      const data=await response.json();
      if(data.error)throw new Error(data.error.message);
      const rawText=data.content?.find(b=>b.type==="text")?.text||"";
      if(!rawText)throw new Error("Empty response — stop_reason: "+data.stop_reason);
      onResult(JSON.parse(rawText.replace(/```json|```/g,"").trim()),imageData);
    }catch(e){setError(e.message||"Could not analyze image.");}
    setLoading(false);
  }
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",zIndex:100,display:"flex",alignItems:"flex-end",justifyContent:"center",backdropFilter:"blur(4px)"}} onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#141414",border:"1px solid rgba(255,255,255,0.12)",borderRadius:"20px 20px 0 0",padding:24,width:"100%",maxWidth:480,animation:"slideUp 0.3s ease"}}>
        <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:24,color:"#fff",marginBottom:16}}>AI FOOD ANALYSIS</div>
        <img src={imageData} alt="" style={{width:"100%",maxHeight:200,objectFit:"cover",borderRadius:12,marginBottom:16}}/>
        <input value={note} onChange={e=>setNote(e.target.value)} placeholder="Add context (e.g. 'large portion', '200g chicken')…" style={{width:"100%",background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:10,padding:"10px 14px",color:"#fff",fontSize:14,fontFamily:"'DM Sans',sans-serif",outline:"none",boxSizing:"border-box",marginBottom:14}}/>
        {error&&<div style={{color:"#f87171",fontSize:13,marginBottom:12,fontFamily:"'DM Sans',sans-serif"}}>{error}</div>}
        <button onClick={analyze} disabled={loading} style={{width:"100%",padding:"14px",background:loading?"rgba(249,115,22,0.4)":"#f97316",border:"none",borderRadius:12,color:"#000",fontFamily:"'Bebas Neue',sans-serif",fontSize:18,cursor:loading?"not-allowed":"pointer",letterSpacing:1}}>{loading?"ANALYSING…":"ANALYSE MEAL"}</button>
      </div>
    </div>
  );
}

function calcCalories(protein, carbs, fat) {
  return Math.round((parseFloat(protein)||0)*4 + (parseFloat(carbs)||0)*4 + (parseFloat(fat)||0)*9);
}

function ManualModal({ prefill, image, source, onSave, onClose, editMode=false }) {
  const [name, setName] = useState(prefill?.name||"");
  const [protein, setProtein] = useState(String(prefill?.protein||""));
  const [carbs, setCarbs] = useState(String(prefill?.carbs||""));
  const [fat, setFat] = useState(String(prefill?.fat||""));
  const [calManual, setCalManual] = useState("");
  const [calOverride, setCalOverride] = useState(false);

  const autoCalories = calcCalories(protein, carbs, fat);
  const displayCalories = calOverride ? calManual : (autoCalories > 0 ? autoCalories : "");

  function handleProtein(v) { setProtein(v); setCalOverride(false); }
  function handleCarbs(v)   { setCarbs(v);   setCalOverride(false); }
  function handleFat(v)     { setFat(v);     setCalOverride(false); }
  function handleCalories(v){ setCalManual(v); setCalOverride(true); }
  function resetToAuto()    { setCalOverride(false); setCalManual(""); }

  function save() {
    if (!name) return;
    const calories = calOverride ? (parseFloat(calManual)||0) : autoCalories;
    onSave({ id: Date.now(), timestamp: new Date().toISOString(), name, calories, protein: parseFloat(protein)||0, carbs: parseFloat(carbs)||0, fat: parseFloat(fat)||0, image: null, source: source||"manual" });
    setTimeout(() => onClose(), 50);
  }

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",zIndex:110,display:"flex",alignItems:"flex-end",justifyContent:"center",backdropFilter:"blur(4px)"}}>
      <div style={{background:"#141414",border:"1px solid rgba(255,255,255,0.12)",borderRadius:"20px 20px 0 0",padding:24,width:"100%",maxWidth:480,animation:"slideUp 0.3s ease",maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:24,color:"#fff",marginBottom:16}}>{editMode?"EDIT MEAL":prefill?"CONFIRM MACROS":"ADD FOOD"}</div>
        {prefill?.note&&<div style={{background:"rgba(96,165,250,0.08)",border:"1px solid rgba(96,165,250,0.2)",borderRadius:8,padding:"8px 12px",marginBottom:14,fontFamily:"'DM Sans',sans-serif",fontSize:12,color:"rgba(255,255,255,0.6)"}}>🤖 {prefill.note} <span style={{color:"rgba(255,255,255,0.3)"}}>({prefill.confidence} confidence)</span></div>}

        {/* Food name */}
        <div style={{marginBottom:12}}>
          <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:"rgba(255,255,255,0.4)",textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>Food Name</div>
          <input value={name} onChange={e=>setName(e.target.value)} placeholder="e.g. Chicken & Rice"
            style={{width:"100%",background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:10,padding:"10px 14px",color:"#fff",fontSize:14,fontFamily:"'DM Sans',sans-serif",outline:"none",boxSizing:"border-box"}}/>
        </div>

        {/* Macros — editing any of these auto-updates calories */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:12}}>
          {[["Protein","protein",protein,handleProtein,MACRO_COLORS.protein],["Carbs","carbs",carbs,handleCarbs,MACRO_COLORS.carbs],["Fat","fat",fat,handleFat,MACRO_COLORS.fat]].map(([l,k,val,handler,c])=>(
            <div key={k}>
              <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:c,textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>{l} (g)</div>
              <input value={val} onChange={e=>handler(e.target.value)} type="number" placeholder="0"
                style={{width:"100%",background:"rgba(255,255,255,0.06)",border:`1px solid ${c}30`,borderRadius:10,padding:"10px",color:"#fff",fontSize:14,fontFamily:"'DM Sans',sans-serif",outline:"none",boxSizing:"border-box"}}/>
            </div>
          ))}
        </div>

        {/* Calories — auto or manual */}
        <div style={{marginBottom:16}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
            <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:MACRO_COLORS.calories,textTransform:"uppercase",letterSpacing:1}}>Calories (kcal)</div>
            {!calOverride && autoCalories>0 && <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:10,color:"rgba(255,255,255,0.3)"}}>auto-calculated</div>}
            {calOverride && <button onClick={resetToAuto} style={{background:"none",border:"none",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontSize:10,color:"#60a5fa",padding:0}}>reset to auto</button>}
          </div>
          <input value={displayCalories} onChange={e=>handleCalories(e.target.value)} type="number" placeholder="0"
            style={{width:"100%",background:calOverride?"rgba(255,255,255,0.08)":"rgba(249,115,22,0.06)",border:`1px solid ${calOverride?"rgba(255,255,255,0.15)":"rgba(249,115,22,0.3)"}`,borderRadius:10,padding:"10px 14px",color:calOverride?"#fff":MACRO_COLORS.calories,fontSize:16,fontFamily:"'Bebas Neue',sans-serif",outline:"none",boxSizing:"border-box",transition:"all 0.2s"}}/>
        </div>

        <div style={{display:"flex",gap:10}}>
          <button onClick={onClose} style={{flex:1,padding:"13px",background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:12,color:"rgba(255,255,255,0.6)",fontFamily:"'Bebas Neue',sans-serif",fontSize:16,cursor:"pointer"}}>CANCEL</button>
          <button onClick={save} style={{flex:2,padding:"13px",background:"#4ade80",border:"none",borderRadius:12,color:"#000",fontFamily:"'Bebas Neue',sans-serif",fontSize:18,cursor:"pointer",letterSpacing:1}}>{editMode?"SAVE CHANGES":"LOG MEAL"}</button>
        </div>
      </div>
    </div>
  );
}

// ── Targets Modal — edits both workout and rest targets ───────────────────────
function TargetsModal({ targetsWorkout, targetsRest, onSave, onClose }) {
  const [tab, setTab] = useState("workout");
  const [formW, setFormW] = useState({...targetsWorkout});
  const [formR, setFormR] = useState({...targetsRest});
  const form = tab === "workout" ? formW : formR;
  const setForm = tab === "workout" ? setFormW : setFormR;

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",zIndex:100,display:"flex",alignItems:"flex-end",justifyContent:"center",backdropFilter:"blur(4px)"}} onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#141414",border:"1px solid rgba(255,255,255,0.12)",borderRadius:"20px 20px 0 0",padding:24,width:"100%",maxWidth:480,animation:"slideUp 0.3s ease"}}>
        <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:24,color:"#fff",marginBottom:16}}>DAILY TARGETS</div>

        {/* Tab toggle */}
        <div style={{display:"flex",gap:8,marginBottom:20,background:"rgba(255,255,255,0.05)",borderRadius:12,padding:4}}>
          {[["workout","💪 Training"],["rest","😴 Rest"]].map(([v,l])=>(
            <button key={v} onClick={()=>setTab(v)} style={{flex:1,padding:"10px",borderRadius:9,border:"none",cursor:"pointer",fontFamily:"'Bebas Neue',sans-serif",fontSize:15,letterSpacing:1,transition:"all 0.2s",background:tab===v?"#f97316":"transparent",color:tab===v?"#000":"rgba(255,255,255,0.4)"}}>{l}</button>
          ))}
        </div>

        {[{k:"calories",l:"Calories (kcal)",c:MACRO_COLORS.calories},{k:"protein",l:"Protein (g)",c:MACRO_COLORS.protein},{k:"carbs",l:"Carbs (g)",c:MACRO_COLORS.carbs},{k:"fat",l:"Fat (g)",c:MACRO_COLORS.fat}].map(({k,l,c})=>(
          <div key={k} style={{marginBottom:14}}>
            <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:c,textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>{l}</div>
            <input type="number" value={form[k]} onChange={e=>setForm(f=>({...f,[k]:parseInt(e.target.value)||0}))} style={{width:"100%",background:"rgba(255,255,255,0.06)",border:`1px solid ${c}30`,borderRadius:10,padding:"10px 14px",color:"#fff",fontSize:16,fontFamily:"'Bebas Neue',sans-serif",outline:"none",boxSizing:"border-box"}}/>
          </div>
        ))}
        <button onClick={()=>{onSave(formW,formR);onClose();}} style={{width:"100%",padding:"14px",background:"#f97316",border:"none",borderRadius:12,color:"#000",fontFamily:"'Bebas Neue',sans-serif",fontSize:18,cursor:"pointer",letterSpacing:1}}>SAVE TARGETS</button>
      </div>
    </div>
  );
}

function HistoryView({ allData, targetsWorkout, targetsRest, onExport, onBackup, onShowRestore }) {
  const days=Object.keys(allData).filter(k=>!k.startsWith("__")).sort((a,b)=>b.localeCompare(a));
  const sevenDaysAgo = new Date(); sevenDaysAgo.setDate(sevenDaysAgo.getDate()-7);
  const cutoff = sevenDaysAgo.toISOString().slice(0,10);

  return (
    <div style={{padding:"0 16px 100px"}}>
      <button onClick={onExport} style={{width:"100%",marginBottom:10,padding:"13px",background:"rgba(74,222,128,0.1)",border:"1px solid rgba(74,222,128,0.25)",borderRadius:12,color:"#4ade80",fontFamily:"'Bebas Neue',sans-serif",fontSize:16,letterSpacing:1,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>⬇ EXPORT TO EXCEL</button>
      <div style={{display:"flex",gap:8,marginBottom:20}}>
        <button onClick={onBackup} style={{flex:1,padding:"11px",background:"rgba(249,115,22,0.08)",border:"1px solid rgba(249,115,22,0.2)",borderRadius:12,color:"#f97316",fontFamily:"'Bebas Neue',sans-serif",fontSize:14,letterSpacing:1,cursor:"pointer"}}>💾 BACKUP JSON</button>
        <button onClick={onShowRestore} style={{flex:1,padding:"11px",background:"rgba(96,165,250,0.08)",border:"1px solid rgba(96,165,250,0.2)",borderRadius:12,color:"#60a5fa",fontFamily:"'Bebas Neue',sans-serif",fontSize:14,letterSpacing:1,cursor:"pointer"}}>📂 RESTORE</button>
      </div>

      {days.length===0 && <div style={{textAlign:"center",padding:60,color:"rgba(255,255,255,0.3)",fontFamily:"'DM Sans',sans-serif"}}>No history yet. Start logging!</div>}

      {days.map(day=>{
        const entries=allData[day]?.entries||[];
        const dayType=allData[day]?.dayType||"workout";
        const tgt=dayType==="rest"?targetsRest:targetsWorkout;
        const t=sumMacros(entries);
        const isToday=day===todayKey();
        const isDetailed = day >= cutoff;
        const dateLabel = isToday?"TODAY":new Date(day+"T12:00:00").toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"}).toUpperCase();

        return (
          <div key={day} style={{marginBottom:16}}>
            {/* Date header */}
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:13,color:"rgba(255,255,255,0.4)",letterSpacing:2,marginBottom:8,display:"flex",alignItems:"center",gap:8}}>
              <span>{dateLabel}</span>
              <span style={{fontSize:10,padding:"2px 8px",borderRadius:20,background:dayType==="rest"?"rgba(96,165,250,0.15)":"rgba(249,115,22,0.15)",color:dayType==="rest"?"#60a5fa":"#f97316",border:`1px solid ${dayType==="rest"?"rgba(96,165,250,0.3)":"rgba(249,115,22,0.3)"}`}}>{dayType==="rest"?"REST":"TRAINING"}</span>
            </div>

            {/* Summary bar */}
            <div style={{background:"rgba(255,255,255,0.04)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:isDetailed?12:"12px 12px 0 0",padding:"12px 16px"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:isDetailed&&entries.length>0?8:0}}>
                <span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:22,color:MACRO_COLORS.calories}}>{Math.round(t.calories)} <span style={{fontSize:13,color:"rgba(255,255,255,0.3)"}}>/ {tgt.calories} kcal</span></span>
                <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:12,color:"rgba(255,255,255,0.3)"}}>{entries.length} meal{entries.length!==1?"s":""}</span>
              </div>
              <div style={{display:"flex",gap:16}}>
                {[{k:"protein",l:"Protein"},{k:"carbs",l:"Carbs"},{k:"fat",l:"Fat"}].map(({k,l})=>(
                  <div key={k}><div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:15,color:MACRO_COLORS[k]}}>{Math.round(t[k])}g</div><div style={{fontFamily:"'DM Sans',sans-serif",fontSize:10,color:"rgba(255,255,255,0.3)",textTransform:"uppercase",letterSpacing:1}}>{l}</div></div>
                ))}
              </div>
            </div>

            {/* Meal detail — only last 7 days */}
            {isDetailed && entries.length>0 && (
              <div style={{background:"rgba(255,255,255,0.02)",border:"1px solid rgba(255,255,255,0.06)",borderTop:"none",borderRadius:"0 0 12px 12px",padding:"8px 12px",display:"flex",flexDirection:"column",gap:6}}>
                {entries.map(e=>{
                  const timeStr = e.timestamp ? new Date(e.timestamp).toLocaleTimeString("en-US",{hour:"numeric",minute:"2-digit",hour12:true}) : "";
                  const icon=e.source==="voice"?"🎤":e.source==="text"?"💬":e.source==="photo"?"📷":"✏️";
                  return (
                    <div key={e.id} style={{display:"flex",alignItems:"flex-start",gap:8,padding:"6px 0",borderBottom:"1px solid rgba(255,255,255,0.04)"}}>
                      <span style={{fontSize:12,flexShrink:0,marginTop:2}}>{icon}</span>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3}}>
                          <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:12,color:"rgba(255,255,255,0.7)",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",flex:1,marginRight:8}}>{e.name}</span>
                          {timeStr&&<span style={{fontFamily:"'DM Sans',sans-serif",fontSize:10,color:"rgba(255,255,255,0.25)",flexShrink:0}}>{timeStr}</span>}
                        </div>
                        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                          <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:MACRO_COLORS.calories}}>{Math.round(e.calories||0)} <span style={{color:"rgba(255,255,255,0.3)"}}>kcal</span></span>
                          <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:MACRO_COLORS.protein}}>{Math.round(e.protein||0)}g <span style={{color:"rgba(255,255,255,0.3)"}}>P</span></span>
                          <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:MACRO_COLORS.carbs}}>{Math.round(e.carbs||0)}g <span style={{color:"rgba(255,255,255,0.3)"}}>C</span></span>
                          <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:MACRO_COLORS.fat}}>{Math.round(e.fat||0)}g <span style={{color:"rgba(255,255,255,0.3)"}}>F</span></span>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function RestoreModal({ onRestore, onClose }) {
  const [error, setError] = useState("");
  const [success, setSuccess] = useState(false);
  const fileRef = useRef();
  function handleFile(file) {
    if (!file) return;
    restoreData(file,
      (data) => { setSuccess(true); setTimeout(() => { onRestore(data); onClose(); }, 1200); },
      (msg) => setError(msg)
    );
  }
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",zIndex:100,display:"flex",alignItems:"flex-end",justifyContent:"center",backdropFilter:"blur(4px)"}} onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#141414",border:"1px solid rgba(255,255,255,0.12)",borderRadius:"20px 20px 0 0",padding:24,width:"100%",maxWidth:480,animation:"slideUp 0.3s ease"}}>
        <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:24,color:"#fff",marginBottom:8}}>RESTORE BACKUP</div>
        <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:13,color:"rgba(255,255,255,0.4)",marginBottom:24}}>Select a .json backup file to restore your data. Existing data will be preserved and merged.</div>
        {success && <div style={{background:"rgba(74,222,128,0.1)",border:"1px solid rgba(74,222,128,0.3)",borderRadius:10,padding:"12px 16px",marginBottom:16,fontFamily:"'DM Sans',sans-serif",fontSize:14,color:"#4ade80"}}>✓ Backup restored successfully!</div>}
        {error && <div style={{background:"rgba(248,113,113,0.1)",border:"1px solid rgba(248,113,113,0.3)",borderRadius:10,padding:"12px 16px",marginBottom:16,fontFamily:"'DM Sans',sans-serif",fontSize:13,color:"#f87171"}}>{error}</div>}
        <input ref={fileRef} type="file" accept=".json" style={{display:"none"}} onChange={e=>handleFile(e.target.files[0])}/>
        <button onClick={()=>fileRef.current.click()} style={{width:"100%",padding:"14px",background:"rgba(96,165,250,0.1)",border:"1px solid rgba(96,165,250,0.3)",borderRadius:12,color:"#60a5fa",fontFamily:"'Bebas Neue',sans-serif",fontSize:18,cursor:"pointer",letterSpacing:1,marginBottom:10}}>📂 SELECT BACKUP FILE</button>
        <button onClick={onClose} style={{width:"100%",padding:"13px",background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:12,color:"rgba(255,255,255,0.6)",fontFamily:"'Bebas Neue',sans-serif",fontSize:16,cursor:"pointer"}}>CANCEL</button>
      </div>
    </div>
  );
}

export default function MacroTracker() {
  const [tab, setTab] = useState("today");
  const [allData, setAllData] = useState(()=>loadData());
  const [targetsWorkout, setTargetsWorkout] = useState(()=>{const d=loadData();return d.__targetsWorkout||DEFAULT_TARGETS_WORKOUT;});
  const [targetsRest, setTargetsRest] = useState(()=>{const d=loadData();return d.__targetsRest||DEFAULT_TARGETS_REST;});

  const [showAI,setShowAI]=useState(false); const [showTextVoice,setShowTextVoice]=useState(false);
  const [showManual,setShowManual]=useState(false); const [showTargets,setShowTargets]=useState(false);
  const [showRestore,setShowRestore]=useState(false);
  const [editingEntry,setEditingEntry]=useState(null);
  const [pendingImage,setPendingImage]=useState(null); const [aiFill,setAiFill]=useState(null);
  const [pendingSource,setPendingSource]=useState("manual"); const [showAddMenu,setShowAddMenu]=useState(false);
  const fileRef=useRef(); const cameraRef=useRef();

  const today=todayKey();
  const todayData=allData[today]||{entries:[],dayType:"workout"};
  const todayEntries=todayData.entries||[];
  const dayType=todayData.dayType||"workout";
  const targets=dayType==="rest"?targetsRest:targetsWorkout;
  const totals=sumMacros(todayEntries);

  function persist(newAll){setAllData(newAll);saveData({...newAll,__targetsWorkout:targetsWorkout,__targetsRest:targetsRest});}

  function setDayType(type){
    const updated={...allData,[today]:{...todayData,dayType:type}};
    persist(updated);
  }

  function persistTargets(w,r){
    setTargetsWorkout(w);setTargetsRest(r);
    saveData({...allData,__targetsWorkout:w,__targetsRest:r});
  }
  function handleRestore(restoredData) {
    const merged = { ...restoredData, ...allData };
    const w = restoredData.__targetsWorkout || targetsWorkout;
    const r = restoredData.__targetsRest || targetsRest;
    setTargetsWorkout(w); setTargetsRest(r);
    setAllData(merged);
    saveData({ ...merged, __targetsWorkout: w, __targetsRest: r });
  }
  function addEntry(entry){persist({...allData,[today]:{...todayData,entries:[...todayEntries,entry]}});}
  function deleteEntry(id){persist({...allData,[today]:{...todayData,entries:todayEntries.filter(e=>e.id!==id)}});}
  function updateEntry(updated){persist({...allData,[today]:{...todayData,entries:todayEntries.map(e=>e.id===updated.id?updated:e)}});}
  function startEdit(entry){setEditingEntry(entry);}

  function handleImage(file){if(!file)return;const r=new FileReader();r.onload=e=>{setPendingImage(e.target.result);setShowAI(true);setShowAddMenu(false);};r.readAsDataURL(file);}
  function handleAIPhotoResult(result,image){setAiFill({...result,image});setPendingSource("photo");setShowAI(false);setShowManual(true);}
  function handleTextVoiceResult(result){setAiFill(result);setPendingSource("text");setShowTextVoice(false);setShowManual(true);}

  const remaining={calories:targets.calories-totals.calories,protein:targets.protein-totals.protein,carbs:targets.carbs-totals.carbs,fat:targets.fat-totals.fat};
  const calPct=Math.round((totals.calories/targets.calories)*100);

  useEffect(()=>{
    const s=document.createElement("style");
    s.textContent=`*{box-sizing:border-box;-webkit-tap-highlight-color:transparent;}@keyframes slideIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}@keyframes slideUp{from{transform:translateY(100%)}to{transform:none}}@keyframes fadeIn{from{opacity:0}to{opacity:1}}@keyframes micPulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:0.5;transform:scale(1.3)}}input[type=number]::-webkit-outer-spin-button,input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none;}input::placeholder,textarea::placeholder{color:rgba(255,255,255,0.2)!important;}::-webkit-scrollbar{width:0;}textarea{resize:none;}`;
    document.head.appendChild(s);
  },[]);

  return (
    <div style={{background:"#0c0c0c",minHeight:"100vh",color:"#fff",fontFamily:"'DM Sans',sans-serif",maxWidth:480,margin:"0 auto",position:"relative"}}>

      {/* Header */}
      <div style={{padding:"20px 20px 0",display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:32,letterSpacing:2,lineHeight:1}}>MACRO</div>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:32,letterSpacing:2,color:"#f97316",lineHeight:1}}>TRACKER</div>
        </div>
        <button onClick={()=>setShowTargets(true)} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:10,padding:"8px 14px",color:"rgba(255,255,255,0.6)",fontFamily:"'DM Sans',sans-serif",fontSize:12,cursor:"pointer"}}>⚙ Targets</button>
      </div>

      {/* Day type toggle */}
      <div style={{padding:"14px 20px 0"}}>
        <div style={{display:"flex",gap:8,background:"rgba(255,255,255,0.05)",borderRadius:12,padding:4}}>
          {[["workout","💪 Training Day"],["rest","😴 Rest Day"]].map(([v,l])=>(
            <button key={v} onClick={()=>setDayType(v)} style={{flex:1,padding:"10px",borderRadius:9,border:"none",cursor:"pointer",fontFamily:"'Bebas Neue',sans-serif",fontSize:14,letterSpacing:1,transition:"all 0.2s",background:dayType===v?(v==="rest"?"#60a5fa":"#f97316"):"transparent",color:dayType===v?"#000":"rgba(255,255,255,0.4)"}}>{l}</button>
          ))}
        </div>
      </div>

      {/* Tabs */}
      <div style={{display:"flex",padding:"12px 20px 0",gap:8}}>
        {["today","history"].map(t=>(
          <button key={t} onClick={()=>setTab(t)} style={{flex:1,padding:"10px",borderRadius:10,background:tab===t?"rgba(249,115,22,0.15)":"rgba(255,255,255,0.04)",border:tab===t?"1px solid rgba(249,115,22,0.3)":"1px solid rgba(255,255,255,0.08)",color:tab===t?"#f97316":"rgba(255,255,255,0.4)",fontFamily:"'Bebas Neue',sans-serif",fontSize:15,letterSpacing:1,cursor:"pointer"}}>{t.toUpperCase()}</button>
        ))}
      </div>

      {tab==="history"?(
        <div style={{marginTop:20}}><HistoryView allData={allData} targetsWorkout={targetsWorkout} targetsRest={targetsRest} onExport={()=>exportToExcel(allData,targetsWorkout,targetsRest)} onBackup={()=>backupData(allData,targetsWorkout,targetsRest)} onShowRestore={()=>setShowRestore(true)}/></div>
      ):(
        <>
          {/* Calorie ring */}
          <div style={{padding:"20px 20px 0",textAlign:"center"}}>
            <div style={{position:"relative",display:"inline-block"}}>
              <svg width={190} height={190} style={{transform:"rotate(-90deg)"}}>
                <circle cx={95} cy={95} r={80} fill="none" stroke="rgba(255,255,255,0.05)" strokeWidth={10}/>
                <circle cx={95} cy={95} r={80} fill="none" stroke={totals.calories>targets.calories?"#f87171":dayType==="rest"?"#60a5fa":"#f97316"} strokeWidth={10} strokeDasharray={`${Math.min(totals.calories/targets.calories,1)*2*Math.PI*80} ${2*Math.PI*80}`} strokeLinecap="round" style={{transition:"stroke-dasharray 0.6s ease"}}/>
              </svg>
              <div style={{position:"absolute",inset:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
                <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:"rgba(255,255,255,0.4)",textTransform:"uppercase",letterSpacing:1,marginBottom:2}}>{remaining.calories>=0?"remaining":"over target"}</div>
                <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:42,color:remaining.calories>=0?(dayType==="rest"?"#60a5fa":"#f97316"):"#f87171",lineHeight:1}}>{Math.abs(Math.round(remaining.calories))}</div>
                <div style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:"rgba(255,255,255,0.3)",marginTop:2}}>{Math.round(totals.calories)} / {targets.calories} kcal</div>
              </div>
            </div>
          </div>

          {/* Macro rings */}
          <div style={{display:"flex",justifyContent:"space-around",padding:"20px 20px 0"}}>
            <Ring value={Math.max(0,Math.round(remaining.protein))} max={targets.protein} color={MACRO_COLORS.protein} label="Protein" sub={`${Math.round(totals.protein)}/${targets.protein}g`} over={remaining.protein<0}/>
            <Ring value={Math.max(0,Math.round(remaining.carbs))} max={targets.carbs} color={MACRO_COLORS.carbs} label="Carbs" sub={`${Math.round(totals.carbs)}/${targets.carbs}g`} over={remaining.carbs<0}/>
            <Ring value={Math.max(0,Math.round(remaining.fat))} max={targets.fat} color={MACRO_COLORS.fat} label="Fat" sub={`${Math.round(totals.fat)}/${targets.fat}g`} over={remaining.fat<0}/>
          </div>

          {/* Macro bars */}
          <div style={{padding:"20px 20px 0"}}>
            <MacroBar label="Protein" value={totals.protein} max={targets.protein} color={MACRO_COLORS.protein}/>
            <MacroBar label="Carbs" value={totals.carbs} max={targets.carbs} color={MACRO_COLORS.carbs}/>
            <MacroBar label="Fat" value={totals.fat} max={targets.fat} color={MACRO_COLORS.fat}/>
          </div>

          {/* Log */}
          <div style={{padding:"20px 20px 100px"}}>
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:16,letterSpacing:2,color:"rgba(255,255,255,0.4)",marginBottom:12}}>TODAY'S LOG — {todayEntries.length} MEAL{todayEntries.length!==1?"S":""}</div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {todayEntries.length===0&&<div style={{textAlign:"center",padding:"32px 0",color:"rgba(255,255,255,0.2)",fontSize:14}}>No meals logged yet. Tap + to add one.</div>}
              {todayEntries.map(e=><EntryCard key={e.id} entry={e} onDelete={deleteEntry} onEdit={startEdit}/>)}
            </div>
          </div>
        </>
      )}

      {/* FAB */}
      <div style={{position:"fixed",bottom:28,right:20,zIndex:50}}>
        {showAddMenu&&(
          <div style={{position:"absolute",bottom:64,right:0,display:"flex",flexDirection:"column",gap:10,alignItems:"flex-end",animation:"fadeIn 0.2s ease"}}>
            {[
              {label:"📷 Take Photo",action:()=>{cameraRef.current.click();setShowAddMenu(false);}},
              {label:"🖼 Upload Photo",action:()=>{fileRef.current.click();setShowAddMenu(false);}},
              {label:"🎤 Speak / Type",action:()=>{setShowTextVoice(true);setShowAddMenu(false);}},
              {label:"✏️ Manual Entry",action:()=>{setAiFill(null);setPendingImage(null);setShowManual(true);setShowAddMenu(false);}},
            ].map(({label,action})=>(
              <button key={label} onClick={action} style={{background:"#1a1a1a",border:"1px solid rgba(255,255,255,0.15)",borderRadius:24,padding:"10px 18px",color:"#fff",fontFamily:"'DM Sans',sans-serif",fontSize:14,cursor:"pointer",whiteSpace:"nowrap",boxShadow:"0 4px 20px rgba(0,0,0,0.5)"}}>{label}</button>
            ))}
          </div>
        )}
        <button onClick={()=>setShowAddMenu(m=>!m)} style={{width:56,height:56,borderRadius:"50%",background:showAddMenu?"#fff":"#f97316",border:"none",cursor:"pointer",fontSize:26,display:"flex",alignItems:"center",justifyContent:"center",boxShadow:"0 4px 24px rgba(249,115,22,0.4)",transition:"background 0.2s,transform 0.2s",transform:showAddMenu?"rotate(45deg)":"none",color:"#000"}}>+</button>
      </div>

      <input ref={fileRef} type="file" accept="image/*" style={{display:"none"}} onChange={e=>handleImage(e.target.files[0])}/>
      <input ref={cameraRef} type="file" accept="image/*" capture="environment" style={{display:"none"}} onChange={e=>handleImage(e.target.files[0])}/>

      {editingEntry&&<ManualModal
        prefill={editingEntry}
        image={editingEntry.image}
        source={editingEntry.source}
        editMode={true}
        onSave={updated=>{updateEntry({...editingEntry,...updated,id:editingEntry.id});setEditingEntry(null);}}
        onClose={()=>setEditingEntry(null)}
      />}
      {showRestore&&<RestoreModal onRestore={handleRestore} onClose={()=>setShowRestore(false)}/>}
      {showAI&&pendingImage&&<AIModal imageData={pendingImage} onResult={handleAIPhotoResult} onClose={()=>{setShowAI(false);setPendingImage(null);}}/>}
      {showTextVoice&&<TextVoiceModal onResult={handleTextVoiceResult} onClose={()=>setShowTextVoice(false)}/>}
      {showManual&&<ManualModal prefill={aiFill} image={aiFill?.image||pendingImage} source={pendingSource} onSave={entry=>{addEntry(entry);setShowManual(false);setAiFill(null);setPendingImage(null);}} onClose={()=>{setShowManual(false);setAiFill(null);setPendingImage(null);}}/>}
      {showTargets&&<TargetsModal targetsWorkout={targetsWorkout} targetsRest={targetsRest} onSave={persistTargets} onClose={()=>setShowTargets(false)}/>}
      {showAddMenu&&<div style={{position:"fixed",inset:0,zIndex:40}} onClick={()=>setShowAddMenu(false)}/>}
    </div>
  );
}
