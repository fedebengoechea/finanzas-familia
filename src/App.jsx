import { useState, useEffect, useRef } from "react";

const USERS = {
  yo:     { name: "Yo",     avatar: "👨", color: "#25D366" },
  esposa: { name: "Esposa", avatar: "👩", color: "#FF6B9D" },
};

const EXPENSE_CATS = { supermercado:"🛒", comida:"🍽️", transporte:"🚗", salud:"💊", educación:"📚", entretenimiento:"🎬", ropa:"👗", servicios:"💡", hogar:"🏠", otros:"📦" };
const INCOME_CATS  = { sueldo:"💼", alquiler:"🏘️", inversión:"📈", freelance:"💻", bono:"🎁", venta:"🏷️", otros_ingreso:"💵" };
const ALL_ICONS    = { ...EXPENSE_CATS, ...INCOME_CATS };

const APPS_SCRIPT = `function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Pestaña Movimientos
  let sheet = ss.getSheetByName("Movimientos");
  if (!sheet) {
    sheet = ss.insertSheet("Movimientos");
    sheet.appendRow([
      "Fecha","Hora","Usuario","Tipo",
      "Categoría","Descripción","Monto","Moneda"
    ]);
  }

  const d = JSON.parse(e.postData.contents);
  sheet.appendRow([
    d.fecha, d.hora, d.usuario, d.tipo,
    d.categoria, d.descripcion,
    d.monto, d.moneda
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}`;

// ── API helpers ───────────────────────────────────────────────────────────────
async function callClaude(messages, system) {
  const res = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 1000, system, messages }),
  });
  return (await res.json()).content?.[0]?.text || "";
}

async function parseMessage(text) {
  const system = `Analizá mensajes financieros en español rioplatense. Respondé SOLO con JSON sin markdown.

Si describe un GASTO:
{"type":"expense","amount":número,"currency":"ARS"|"USD","category":"supermercado|comida|transporte|salud|educación|entretenimiento|ropa|servicios|hogar|otros","description":"texto corto","reply":"confirmación 1 oración con emoji"}

Si describe un INGRESO (sueldo, cobré, me pagaron, alquiler, dividendos, venta, etc.):
{"type":"income","amount":número,"currency":"ARS"|"USD","category":"sueldo|alquiler|inversión|freelance|bono|venta|otros_ingreso","description":"texto corto","reply":"confirmación 1 oración con emoji"}

Si NO es movimiento:
{"type":"none","reply":"respuesta amigable breve"}`;
  try {
    return JSON.parse((await callClaude([{ role:"user", content:text }], system)).replace(/```json|```/g,"").trim());
  } catch {
    return { type:"none", reply:"No pude entender ese mensaje 😅" };
  }
}

async function generateReport(entries) {
  if (!entries.length) return "Todavía no hay movimientos registrados.";
  const system = `Sos un asistente financiero familiar. Analizá ingresos y gastos, y hacé un reporte en español rioplatense con: balance neto, total ingresos, total gastos, desglose por categoría, quién gastó/cobró más, tasa de ahorro, y 2-3 consejos concretos. Usá emojis. Máx 400 palabras.`;
  return callClaude([{ role:"user", content: JSON.stringify(entries) }], system);
}

async function sendToSheets(url, entry, userName) {
  if (!url) return false;
  try {
    await fetch(url, {
      method: "POST", mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        fecha:      new Date(entry.ts).toLocaleDateString("es-AR"),
        hora:       new Date(entry.ts).toLocaleTimeString("es-AR",{hour:"2-digit",minute:"2-digit"}),
        usuario:    userName,
        tipo:       entry.type === "income" ? "Ingreso" : "Gasto",
        categoria:  entry.category,
        descripcion:entry.description,
        monto:      entry.amount,
        moneda:     entry.currency,
      }),
    });
    return true;
  } catch { return false; }
}

// ── Formatters ────────────────────────────────────────────────────────────────
const fmt     = (n, c) => c === "USD" ? `USD ${Number(n).toFixed(2)}` : `$${Number(n).toLocaleString("es-AR")}`;
const fmtTime = ts => new Date(ts).toLocaleTimeString("es-AR",{hour:"2-digit",minute:"2-digit"});
const fmtDate = ts => new Date(ts).toLocaleDateString("es-AR",{weekday:"short",day:"numeric",month:"short"});

// ── Setup Screen ──────────────────────────────────────────────────────────────
function SetupScreen({ onDone }) {
  const [url, setUrl] = useState("");
  const [copied, setCopied] = useState(false);
  function copy() { navigator.clipboard.writeText(APPS_SCRIPT).then(()=>{ setCopied(true); setTimeout(()=>setCopied(false),2000); }); }
  const steps = [
    { n:"1", icon:"📄", title:"Creá una planilla", body: <> Andá a <a href="https://sheets.new" target="_blank" rel="noreferrer" style={{color:"#075E54",fontWeight:700}}>sheets.new</a> para crear una Google Sheet nueva. </> },
    { n:"2", icon:"📝", title:"Abrí Apps Script",  body: <> En la planilla, andá a <strong>Extensiones → Apps Script</strong>. Borrá el código existente y pegá el siguiente: </> },
    { n:"3", icon:"🚀", title:"Publicá como app web", body: <> Hacé clic en <strong>Implementar → Nueva implementación</strong>. Tipo: <em>App web</em>. Ejecutar como: <em>Yo</em>. Acceso: <em>Cualquier usuario</em>. Copiá la URL. </> },
    { n:"4", icon:"🔗", title:"Pegá la URL abajo",  body: <>Eso es todo. Ingresos y gastos irán a la pestaña "Movimientos" de tu planilla.</> },
  ];
  return (
    <div style={{flex:1,overflowY:"auto",padding:20,background:"#ECE5DD",display:"flex",flexDirection:"column",gap:12}}>
      <div style={{textAlign:"center"}}><div style={{fontSize:44,marginBottom:4}}>📊</div><div style={{fontWeight:800,fontSize:19,color:"#075E54"}}>Conectar Google Sheets</div><div style={{fontSize:13,color:"#777",marginTop:3}}>Ingresos y gastos se guardan automáticamente</div></div>
      {steps.map(s=>(
        <div key={s.n} style={{background:"#fff",borderRadius:14,padding:"13px 15px",boxShadow:"0 1px 4px rgba(0,0,0,0.08)"}}>
          <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:5}}>
            <div style={{width:26,height:26,borderRadius:"50%",background:"#075E54",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:12,flexShrink:0}}>{s.n}</div>
            <div style={{fontWeight:800,fontSize:14,color:"#222"}}>{s.icon} {s.title}</div>
          </div>
          <div style={{fontSize:13,color:"#555",lineHeight:1.6,paddingLeft:35}}>{s.body}</div>
          {s.n==="2"&&(
            <div style={{marginTop:9,paddingLeft:35}}>
              <div style={{background:"#1e1e2e",borderRadius:10,padding:12,position:"relative"}}>
                <pre style={{color:"#cdd6f4",fontSize:11,lineHeight:1.6,margin:0,whiteSpace:"pre-wrap",overflowX:"auto"}}>{APPS_SCRIPT}</pre>
                <button onClick={copy} style={{position:"absolute",top:8,right:8,background:copied?"#25D366":"rgba(255,255,255,0.15)",color:"#fff",border:"none",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer",fontFamily:"inherit",fontWeight:700,transition:"background 0.2s"}}>{copied?"✓ Copiado":"Copiar"}</button>
              </div>
            </div>
          )}
        </div>
      ))}
      <div style={{background:"#fff",borderRadius:14,padding:15,boxShadow:"0 1px 4px rgba(0,0,0,0.08)"}}>
        <label style={{fontSize:13,fontWeight:700,color:"#075E54",display:"block",marginBottom:7}}>🔗 URL del Apps Script</label>
        <input value={url} onChange={e=>setUrl(e.target.value)} placeholder="https://script.google.com/macros/s/..." style={{width:"100%",padding:"10px 12px",borderRadius:10,border:"1.5px solid #ddd",fontSize:12,fontFamily:"monospace",outline:"none",marginBottom:10}}/>
        <button onClick={()=>onDone(url.trim())} style={{width:"100%",padding:12,background:url.trim()?"linear-gradient(135deg,#075E54,#25D366)":"#ccc",color:"#fff",border:"none",borderRadius:12,cursor:url.trim()?"pointer":"default",fontFamily:"inherit",fontWeight:800,fontSize:14}}>✓ Conectar planilla</button>
        <button onClick={()=>onDone("")} style={{width:"100%",marginTop:8,padding:9,background:"transparent",color:"#999",border:"none",cursor:"pointer",fontFamily:"inherit",fontSize:13}}>Saltar (solo guardar en el celu)</button>
      </div>
    </div>
  );
}

// ── Balance bar ───────────────────────────────────────────────────────────────
function BalanceBar({ entries }) {
  const arsIncome  = entries.filter(e=>e.type==="income"&&e.currency==="ARS").reduce((s,e)=>s+e.amount,0);
  const arsExpense = entries.filter(e=>e.type==="expense"&&e.currency==="ARS").reduce((s,e)=>s+e.amount,0);
  const balance    = arsIncome - arsExpense;
  const pct        = arsIncome > 0 ? Math.min(100, Math.round((arsExpense/arsIncome)*100)) : 0;
  const color      = balance >= 0 ? "#25D366" : "#FF5252";

  return (
    <div style={{background:"rgba(255,255,255,0.12)",borderRadius:10,padding:"8px 12px",marginBottom:8}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5}}>
        <div style={{display:"flex",gap:14}}>
          <span style={{fontSize:11,color:"rgba(255,255,255,0.8)"}}>⬆️ <strong style={{color:"#afffaf"}}>{fmt(arsIncome,"ARS")}</strong></span>
          <span style={{fontSize:11,color:"rgba(255,255,255,0.8)"}}>⬇️ <strong style={{color:"#ffb3b3"}}>{fmt(arsExpense,"ARS")}</strong></span>
        </div>
        <span style={{fontSize:12,fontWeight:800,color:balance>=0?"#afffaf":"#ffb3b3"}}>{balance>=0?"":"−"}{fmt(Math.abs(balance),"ARS")}</span>
      </div>
      {arsIncome>0&&(
        <div style={{height:4,background:"rgba(255,255,255,0.15)",borderRadius:2}}>
          <div style={{height:"100%",background:pct>90?"#FF5252":pct>70?"#FFC107":"#25D366",borderRadius:2,width:`${pct}%`,transition:"width 0.4s"}}/>
        </div>
      )}
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [ready,        setReady]        = useState(false);
  const [showSetup,    setShowSetup]    = useState(false);
  const [scriptUrl,    setScriptUrl]    = useState("");
  const [currentUser,  setCurrentUser]  = useState("yo");
  const [messages,     setMessages]     = useState([{ id:1, type:"system", text:"👋 ¡Hola! Registrá gastos e ingresos en lenguaje natural.\n\n💸 *Gastos:* \"gasté 2500 en el super\", \"taxi 800\"\n💰 *Ingresos:* \"cobré el sueldo 150000\", \"me pagaron el alquiler USD 400\"", ts:Date.now() }]);
  const [input,        setInput]        = useState("");
  const [entries,      setEntries]      = useState([]); // unified list
  const [loading,      setLoading]      = useState(false);
  const [view,         setView]         = useState("chat");
  const [histFilter,   setHistFilter]   = useState("all"); // all | income | expense
  const [report,       setReport]       = useState("");
  const [reportLoading,setReportLoading]= useState(false);
  const [syncStatus,   setSyncStatus]   = useState(null);
  const endRef = useRef(null);

  useEffect(()=>{
    const configured = localStorage.getItem("fam2_configured");
    if (!configured) { setShowSetup(true); setReady(true); return; }
    setScriptUrl(localStorage.getItem("fam2_script_url")||"");
    const e = localStorage.getItem("fam2_entries");  if (e) setEntries(JSON.parse(e));
    const m = localStorage.getItem("fam2_messages"); if (m) setMessages(JSON.parse(m));
    setReady(true);
  },[]);

  useEffect(()=>{ endRef.current?.scrollIntoView({behavior:"smooth"}); },[messages]);

  function saveEntries(e) { setEntries(e); localStorage.setItem("fam2_entries",  JSON.stringify(e)); }
  function saveMsgs(m)    { setMessages(m); localStorage.setItem("fam2_messages", JSON.stringify(m)); }

  function handleSetupDone(url) {
    localStorage.setItem("fam2_configured","1");
    localStorage.setItem("fam2_script_url", url||"");
    setScriptUrl(url||""); setShowSetup(false);
  }

  async function handleSend() {
    if (!input.trim()||loading) return;
    const txt = input.trim(); setInput(""); setLoading(true);
    const userMsg = { id:Date.now(), type:"user", user:currentUser, text:txt, ts:Date.now() };
    const newMsgs = [...messages, userMsg]; saveMsgs(newMsgs);
    try {
      const result = await parseMessage(txt);
      let newEntries = entries;
      if (result.type==="income"||result.type==="expense") {
        const entry = { id:Date.now(), type:result.type, amount:result.amount, currency:result.currency||"ARS", category:result.category||"otros", description:result.description, user:currentUser, ts:Date.now(), synced:false };
        newEntries = [...entries, entry]; saveEntries(newEntries);
        if (scriptUrl) {
          setSyncStatus("syncing");
          const ok = await sendToSheets(scriptUrl, entry, USERS[currentUser].name);
          if (ok) { const s=newEntries.map(e=>e.id===entry.id?{...e,synced:true}:e); saveEntries(s); setSyncStatus("ok"); setTimeout(()=>setSyncStatus(null),3000); }
          else setSyncStatus(null);
        }
      }
      const isIncome = result.type==="income";
      const botMsg = { id:Date.now()+1, type:"bot", text:result.reply,
        card: (result.type==="income"||result.type==="expense") ? { entryType:result.type, amount:result.amount, currency:result.currency||"ARS", category:result.category, description:result.description } : null,
        ts:Date.now() };
      saveMsgs([...newMsgs, botMsg]);
    } catch { saveMsgs([...newMsgs,{id:Date.now()+1,type:"bot",text:"Ups, algo salió mal 😅",ts:Date.now()}]); }
    setLoading(false);
  }

  async function handleReport() {
    setView("report"); setReportLoading(true); setReport("");
    setReport(await generateReport(entries)); setReportLoading(false);
  }

  // ── Derived stats ─────────────────────────────────────────────────────────
  const arsIncome   = entries.filter(e=>e.type==="income" &&e.currency==="ARS").reduce((s,e)=>s+e.amount,0);
  const arsExpense  = entries.filter(e=>e.type==="expense"&&e.currency==="ARS").reduce((s,e)=>s+e.amount,0);
  const usdIncome   = entries.filter(e=>e.type==="income" &&e.currency==="USD").reduce((s,e)=>s+e.amount,0);
  const usdExpense  = entries.filter(e=>e.type==="expense"&&e.currency==="USD").reduce((s,e)=>s+e.amount,0);

  const byCat = (type) => entries.filter(e=>e.type===type&&e.currency==="ARS").reduce((a,e)=>{ a[e.category]=(a[e.category]||0)+e.amount; return a; },{});
  const expCats = byCat("expense"); const incCats = byCat("income");

  const filteredEntries = histFilter==="all" ? entries : entries.filter(e=>e.type===histFilter);

  if (!ready) return null;

  return (
    <div style={{fontFamily:"'Nunito',sans-serif",height:"100vh",display:"flex",flexDirection:"column",background:"#ECE5DD",maxWidth:480,margin:"0 auto",overflow:"hidden",position:"relative"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:4px} ::-webkit-scrollbar-thumb{background:#ccc;border-radius:2px}
        .pop{animation:pop 0.2s ease}
        @keyframes pop{from{opacity:0;transform:scale(0.93) translateY(5px)}to{opacity:1;transform:scale(1) translateY(0)}}
        @keyframes dot{0%,80%,100%{transform:scale(0.7)}40%{transform:scale(1.1)}}
        textarea{resize:none} button{transition:opacity 0.15s,transform 0.1s} button:active{transform:scale(0.95)}
      `}</style>

      {/* ── Header ─────────────────────────────────────────────────────────── */}
      <div style={{background:"linear-gradient(135deg,#075E54,#128C7E)",padding:"12px 14px 8px",boxShadow:"0 2px 8px rgba(0,0,0,0.25)",flexShrink:0}}>
        {/* Row 1: logo + users */}
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
          <div style={{display:"flex",alignItems:"center",gap:9}}>
            <div style={{width:38,height:38,borderRadius:"50%",background:"rgba(255,255,255,0.15)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:19}}>💰</div>
            <div>
              <div style={{color:"#fff",fontWeight:800,fontSize:15}}>FinanzasFamilia</div>
              <div style={{display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>
                <span style={{color:"rgba(255,255,255,0.75)",fontSize:10}}>{entries.length} movimientos</span>
                {scriptUrl
                  ? <span style={{fontSize:10,padding:"1px 6px",borderRadius:8,background:syncStatus==="syncing"?"rgba(255,200,0,0.3)":syncStatus==="ok"?"rgba(37,211,102,0.3)":"rgba(255,255,255,0.12)",color:syncStatus==="ok"?"#afffaf":"rgba(255,255,255,0.75)",fontWeight:700}}>
                      {syncStatus==="syncing"?"⏳ Guardando...":syncStatus==="ok"?"✓ Sheets":"📊 Sheets"}
                    </span>
                  : <span onClick={()=>setShowSetup(true)} style={{fontSize:10,padding:"1px 6px",borderRadius:8,background:"rgba(255,150,0,0.25)",color:"#ffe0a0",fontWeight:700,cursor:"pointer"}}>⚠️ Sin Sheets</span>
                }
              </div>
            </div>
          </div>
          <div style={{display:"flex",gap:5}}>
            {Object.entries(USERS).map(([k,u])=>(
              <div key={k} onClick={()=>setCurrentUser(k)} style={{padding:"4px 9px",borderRadius:20,fontSize:12,fontWeight:700,background:currentUser===k?u.color:"rgba(255,255,255,0.14)",color:"#fff",border:`2px solid ${currentUser===k?"rgba(255,255,255,0.5)":"transparent"}`,cursor:"pointer"}}>
                {u.avatar} {u.name}
              </div>
            ))}
          </div>
        </div>
        {/* Balance bar */}
        <BalanceBar entries={entries}/>
        {/* Tabs */}
        <div style={{display:"flex",gap:3}}>
          {[["chat","💬"],["history","📋"],["report","📊"],["settings","⚙️"]].map(([v,l])=>(
            <button key={v} onClick={()=>v==="report"?handleReport():setView(v)} style={{flex:v==="settings"?0:1,padding:"5px 3px",borderRadius:7,border:"none",cursor:"pointer",background:view===v?"rgba(255,255,255,0.25)":"transparent",color:view===v?"#fff":"rgba(255,255,255,0.6)",fontWeight:view===v?800:600,fontSize:12,fontFamily:"inherit"}}>
              {l}
            </button>
          ))}
        </div>
      </div>

      {/* ── Setup overlay ──────────────────────────────────────────────────── */}
      {showSetup&&(
        <div style={{position:"absolute",inset:0,zIndex:100,display:"flex",flexDirection:"column",background:"#ECE5DD"}}>
          <div style={{background:"linear-gradient(135deg,#075E54,#128C7E)",padding:"12px 14px",display:"flex",alignItems:"center",gap:9,flexShrink:0}}>
            <button onClick={()=>setShowSetup(false)} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",borderRadius:8,padding:"4px 10px",cursor:"pointer",fontFamily:"inherit",fontWeight:700,fontSize:13}}>← Volver</button>
            <span style={{color:"#fff",fontWeight:800,fontSize:15}}>Conectar Google Sheets</span>
          </div>
          <SetupScreen onDone={handleSetupDone}/>
        </div>
      )}

      {/* ── CHAT ───────────────────────────────────────────────────────────── */}
      {view==="chat"&&(<>
        <div style={{flex:1,overflowY:"auto",padding:"12px 10px",display:"flex",flexDirection:"column",gap:6}}>
          {messages.map(msg=>{
            if (msg.type==="system") return (
              <div key={msg.id} style={{textAlign:"center",margin:"4px 0"}}>
                <span style={{background:"rgba(0,0,0,0.1)",color:"#555",fontSize:12,padding:"7px 12px",borderRadius:10,display:"inline-block",maxWidth:"92%",lineHeight:1.6,textAlign:"left",whiteSpace:"pre-line"}}
                  dangerouslySetInnerHTML={{__html:msg.text.replace(/\*(.*?)\*/g,"<strong>$1</strong>")}}/>
              </div>
            );
            const isUser = msg.type==="user";
            const u = isUser ? USERS[msg.user] : null;
            return (
              <div key={msg.id} className="pop" style={{display:"flex",flexDirection:"column",alignItems:isUser?"flex-end":"flex-start"}}>
                {isUser&&<div style={{fontSize:11,color:u.color,fontWeight:700,marginBottom:2,marginRight:4}}>{u.avatar} {u.name}</div>}
                <div style={{maxWidth:"78%",padding:"8px 12px",borderRadius:isUser?"18px 4px 18px 18px":"4px 18px 18px 18px",background:isUser?(msg.user==="yo"?"#DCF8C6":"#FFD6EA"):"#fff",boxShadow:"0 1px 3px rgba(0,0,0,0.1)",fontSize:14,lineHeight:1.5,color:"#333"}}>
                  {msg.text}
                  {msg.card&&(()=>{
                    const isInc = msg.card.entryType==="income";
                    return (
                      <div style={{marginTop:8,padding:"7px 10px",background:isInc?"rgba(37,211,102,0.12)":"rgba(18,140,126,0.1)",borderRadius:9,borderLeft:`3px solid ${isInc?"#25D366":"#128C7E"}`}}>
                        <div style={{display:"flex",alignItems:"center",gap:6}}>
                          <span style={{fontSize:13,fontWeight:800,color:isInc?"#1a7a40":"#075E54"}}>
                            {isInc?"⬆️":"⬇️"} {ALL_ICONS[msg.card.category]||"📦"} {fmt(msg.card.amount,msg.card.currency)}
                          </span>
                          <span style={{fontSize:10,fontWeight:700,padding:"1px 7px",borderRadius:8,background:isInc?"rgba(37,211,102,0.2)":"rgba(255,100,100,0.15)",color:isInc?"#1a7a40":"#c0392b"}}>{isInc?"INGRESO":"GASTO"}</span>
                        </div>
                        <div style={{fontSize:12,color:"#666",textTransform:"capitalize",marginTop:2}}>{msg.card.category.replace("_ingreso","")} · {msg.card.description}</div>
                      </div>
                    );
                  })()}
                  <div style={{fontSize:10,color:"#999",textAlign:"right",marginTop:4}}>{fmtTime(msg.ts)}</div>
                </div>
              </div>
            );
          })}
          {loading&&(
            <div style={{display:"flex"}}>
              <div style={{background:"#fff",borderRadius:"4px 18px 18px 18px",padding:"10px 14px",boxShadow:"0 1px 3px rgba(0,0,0,0.1)",display:"flex",gap:4}}>
                {[0,1,2].map(i=><div key={i} style={{width:7,height:7,borderRadius:"50%",background:"#128C7E",animation:`dot 0.8s ${i*0.15}s infinite`,opacity:0.7}}/>)}
              </div>
            </div>
          )}
          <div ref={endRef}/>
        </div>
        <div style={{padding:"7px 10px 12px",background:"#F0F0F0",display:"flex",gap:8,alignItems:"flex-end",flexShrink:0}}>
          <div style={{flex:1,background:"#fff",borderRadius:24,padding:"8px 14px",boxShadow:"0 1px 3px rgba(0,0,0,0.1)"}}>
            <textarea value={input} onChange={e=>setInput(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();handleSend();}}} placeholder={`${USERS[currentUser].avatar} Gasto o ingreso...`} rows={1} style={{width:"100%",border:"none",outline:"none",fontSize:14,fontFamily:"inherit",background:"transparent",color:"#333",lineHeight:1.4}}/>
          </div>
          <button onClick={handleSend} disabled={loading||!input.trim()} style={{width:44,height:44,borderRadius:"50%",background:input.trim()?"#25D366":"#ccc",border:"none",cursor:input.trim()?"pointer":"default",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,boxShadow:"0 2px 6px rgba(0,0,0,0.2)"}}>➤</button>
        </div>
      </>)}

      {/* ── HISTORY ────────────────────────────────────────────────────────── */}
      {view==="history"&&(
        <div style={{flex:1,overflowY:"auto",padding:12}}>
          {/* Summary cards */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:7,marginBottom:10}}>
            {[
              {label:"💰 Ingresos",  val:fmt(arsIncome,"ARS"),  sub:usdIncome>0?`+ USD ${usdIncome}`:"",  color:"#25D366", bg:"#f0fff4"},
              {label:"💸 Gastos",    val:fmt(arsExpense,"ARS"), sub:usdExpense>0?`+ USD ${usdExpense}`:"", color:"#FF5252", bg:"#fff5f5"},
              {label:"📊 Balance",   val:fmt(Math.abs(arsIncome-arsExpense),"ARS"), sub:(arsIncome-arsExpense)>=0?"Superávit":"Déficit", color:(arsIncome-arsExpense)>=0?"#25D366":"#FF5252", bg:(arsIncome-arsExpense)>=0?"#f0fff4":"#fff5f5"},
              {label:"📈 Ahorro %",  val: arsIncome>0?`${Math.max(0,Math.round(((arsIncome-arsExpense)/arsIncome)*100))}%`:"—", sub:"del ingreso", color:"#128C7E", bg:"#f0fffe"},
            ].map(c=>(
              <div key={c.label} style={{background:c.bg,borderRadius:13,padding:"11px 13px",boxShadow:"0 1px 4px rgba(0,0,0,0.07)",borderLeft:`3px solid ${c.color}`}}>
                <div style={{fontSize:11,color:"#888",fontWeight:600,marginBottom:2}}>{c.label}</div>
                <div style={{fontSize:16,fontWeight:800,color:"#222"}}>{c.val}</div>
                {c.sub&&<div style={{fontSize:11,color:c.color,fontWeight:600}}>{c.sub}</div>}
              </div>
            ))}
          </div>

          {/* Per-user */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:7,marginBottom:10}}>
            {Object.entries(USERS).map(([k,u])=>{
              const inc = entries.filter(e=>e.user===k&&e.type==="income" &&e.currency==="ARS").reduce((s,e)=>s+e.amount,0);
              const exp = entries.filter(e=>e.user===k&&e.type==="expense"&&e.currency==="ARS").reduce((s,e)=>s+e.amount,0);
              return (
                <div key={k} style={{background:"#fff",borderRadius:13,padding:"11px 13px",boxShadow:"0 1px 4px rgba(0,0,0,0.07)",borderLeft:`3px solid ${u.color}`}}>
                  <div style={{fontSize:12,color:"#888",fontWeight:700,marginBottom:3}}>{u.avatar} {u.name}</div>
                  <div style={{fontSize:11,color:"#4caf50"}}>⬆️ {fmt(inc,"ARS")}</div>
                  <div style={{fontSize:11,color:"#f44336"}}>⬇️ {fmt(exp,"ARS")}</div>
                </div>
              );
            })}
          </div>

          {/* Category breakdowns */}
          {(Object.keys(incCats).length>0||Object.keys(expCats).length>0)&&(
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:7,marginBottom:10}}>
              {[["income","💰 Ingresos",incCats,arsIncome,"#25D366"],["expense","💸 Gastos",expCats,arsExpense,"#FF5252"]].map(([type,title,cats,total,color])=>(
                Object.keys(cats).length>0&&(
                  <div key={type} style={{background:"#fff",borderRadius:13,padding:"11px 13px",boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
                    <div style={{fontWeight:800,fontSize:12,color:"#333",marginBottom:7}}>{title}</div>
                    {Object.entries(cats).sort((a,b)=>b[1]-a[1]).map(([cat,amt])=>{
                      const pct = total>0?Math.round((amt/total)*100):0;
                      return (
                        <div key={cat} style={{marginBottom:5}}>
                          <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:1}}>
                            <span>{ALL_ICONS[cat]||"📦"} {cat.replace("_ingreso","")}</span>
                            <span style={{fontWeight:700}}>{pct}%</span>
                          </div>
                          <div style={{height:3,background:"#eee",borderRadius:2}}><div style={{height:"100%",background:color,borderRadius:2,width:`${pct}%`}}/></div>
                        </div>
                      );
                    })}
                  </div>
                )
              ))}
            </div>
          )}

          {/* Filter tabs */}
          <div style={{display:"flex",gap:5,marginBottom:8}}>
            {[["all","Todos"],["income","Ingresos"],["expense","Gastos"]].map(([f,l])=>(
              <button key={f} onClick={()=>setHistFilter(f)} style={{flex:1,padding:"6px 4px",borderRadius:8,border:"none",cursor:"pointer",background:histFilter===f?"#075E54":"#fff",color:histFilter===f?"#fff":"#555",fontWeight:700,fontSize:12,fontFamily:"inherit",boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
                {l}
              </button>
            ))}
          </div>

          {/* Entry list */}
          <div style={{display:"flex",flexDirection:"column",gap:6}}>
            {[...filteredEntries].reverse().map(e=>{
              const isInc = e.type==="income";
              return (
                <div key={e.id} style={{background:"#fff",borderRadius:12,padding:"10px 13px",boxShadow:"0 1px 4px rgba(0,0,0,0.07)",display:"flex",justifyContent:"space-between",alignItems:"center",borderLeft:`3px solid ${isInc?"#25D366":"#FF5252"}`}}>
                  <div style={{display:"flex",alignItems:"center",gap:9}}>
                    <span style={{fontSize:21}}>{ALL_ICONS[e.category]||"📦"}</span>
                    <div>
                      <div style={{fontWeight:700,fontSize:13,color:"#222"}}>{e.description}</div>
                      <div style={{fontSize:11,color:"#999",display:"flex",alignItems:"center",gap:4}}>
                        {USERS[e.user]?.avatar} {USERS[e.user]?.name} · {fmtDate(e.ts)}
                        {scriptUrl&&<span style={{color:e.synced?"#25D366":"#ddd",fontSize:10}}>{e.synced?"✓":"○"}</span>}
                      </div>
                    </div>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontWeight:800,fontSize:14,color:isInc?"#1a7a40":"#c0392b"}}>{isInc?"+":"−"}{fmt(e.amount,e.currency)}</div>
                    <div style={{fontSize:10,color:isInc?"#25D366":"#FF5252",fontWeight:700}}>{isInc?"INGRESO":"GASTO"}</div>
                  </div>
                </div>
              );
            })}
            {!filteredEntries.length&&<div style={{textAlign:"center",padding:40,color:"#aaa"}}><div style={{fontSize:40}}>📭</div><div style={{marginTop:8}}>No hay movimientos aún</div></div>}
          </div>
          {entries.length>0&&<button onClick={()=>saveEntries([])} style={{width:"100%",marginTop:12,padding:10,background:"#FFF0F0",color:"#e74c3c",border:"1px solid #FFCCCC",borderRadius:10,cursor:"pointer",fontFamily:"inherit",fontWeight:700,fontSize:13}}>🗑️ Borrar todos los movimientos</button>}
        </div>
      )}

      {/* ── REPORT ─────────────────────────────────────────────────────────── */}
      {view==="report"&&(
        <div style={{flex:1,overflowY:"auto",padding:14}}>
          <div style={{background:"#fff",borderRadius:16,padding:16,boxShadow:"0 1px 4px rgba(0,0,0,0.08)",minHeight:200}}>
            <div style={{fontWeight:800,fontSize:15,color:"#075E54",marginBottom:12}}>📊 Análisis financiero IA</div>
            {reportLoading
              ?<div style={{textAlign:"center",padding:40,color:"#888"}}><div style={{fontSize:30,marginBottom:8}}>🤔</div>Analizando finanzas...</div>
              :<div style={{fontSize:14,lineHeight:1.8,color:"#333",whiteSpace:"pre-wrap"}}>{report||"Generando análisis..."}</div>}
          </div>
          <button onClick={handleReport} style={{width:"100%",marginTop:12,padding:12,background:"linear-gradient(135deg,#075E54,#25D366)",color:"#fff",border:"none",borderRadius:12,cursor:"pointer",fontFamily:"inherit",fontWeight:800,fontSize:14}}>🔄 Regenerar</button>
        </div>
      )}

      {/* ── SETTINGS ───────────────────────────────────────────────────────── */}
      {view==="settings"&&(
        <div style={{flex:1,overflowY:"auto",padding:14,display:"flex",flexDirection:"column",gap:10}}>
          <div style={{background:"#fff",borderRadius:15,padding:15,boxShadow:"0 1px 4px rgba(0,0,0,0.08)"}}>
            <div style={{fontWeight:800,fontSize:14,color:"#075E54",marginBottom:10}}>📊 Google Sheets</div>
            {scriptUrl
              ?<><div style={{fontSize:13,color:"#555",marginBottom:9}}>✅ Conectado. Ingresos y gastos se guardan en la pestaña "Movimientos" de tu planilla.</div>
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={()=>setShowSetup(true)} style={{flex:1,padding:"8px",background:"#f0f0f0",color:"#333",border:"none",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontWeight:700,fontSize:13}}>Cambiar URL</button>
                    <button onClick={()=>{localStorage.setItem("fam2_script_url","");setScriptUrl("");}} style={{flex:1,padding:"8px",background:"#FFF0F0",color:"#e74c3c",border:"none",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontWeight:700,fontSize:13}}>Desconectar</button>
                  </div>
                </>
              :<><div style={{fontSize:13,color:"#888",marginBottom:9}}>Sin conectar. Los datos solo se guardan en este dispositivo.</div>
                  <button onClick={()=>setShowSetup(true)} style={{width:"100%",padding:10,background:"linear-gradient(135deg,#075E54,#25D366)",color:"#fff",border:"none",borderRadius:10,cursor:"pointer",fontFamily:"inherit",fontWeight:800,fontSize:13}}>🔗 Conectar Google Sheets</button>
                </>
            }
          </div>
          <div style={{background:"#fff",borderRadius:15,padding:15,boxShadow:"0 1px 4px rgba(0,0,0,0.08)"}}>
            <div style={{fontWeight:800,fontSize:14,color:"#075E54",marginBottom:9}}>📱 Instalar como app</div>
            <div style={{fontSize:13,color:"#555",lineHeight:1.8}}>
              <p><strong>Android (Chrome):</strong><br/>Tocá los ⋮ → "Agregar a pantalla de inicio"</p><br/>
              <p><strong>iPhone (Safari):</strong><br/>Tocá el ícono compartir → "Agregar a pantalla de inicio"</p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
