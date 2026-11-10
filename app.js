import { SHEETS_API_URL, firebaseConfig, EDITORES } from "./config.js";

/* ═════════════ Definición de campos (mismo orden que el Excel) ═════════════ */
const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const FIELDS = [
  // key, Excel header, group, type, extra
  ["nave","Nave","gId","text",{req:1,list:"dlNave"}],
  ["anio","Año","gId","number"],
  ["mes","Mes","gId","select",{opts:MESES}],
  ["cliente","Cliente","gId","text",{list:"dlCliente"}],
  ["recalada","Recalada","gId","text"],
  ["eslora","Eslora","gId","number",{step:"0.01"}],
  ["manga","Manga","gId","number",{step:"0.01"}],
  ["proceso","Proceso","gId","select",{opts:["Descarga","Embarque","Descarga / Embarque","Sin Carga","Otro"]}],
  ["equipo","Equipo","gId","text",{list:"dlEquipo"}],
  ["producto","Producto","gId","text",{list:"dlProducto"}],
  ["tipoCarga","Tipo de carga","gId","select",{opts:["Granel solido","Granel liquido","Carga General","Break Bulk","Otro"]}],
  ["sitio","Sitio","gId","select",{opts:["Norte","Sur"]}],
  ["operador","Operador","gId","text",{list:"dlOperador"}],
  ["tonEmb","Toneladas embarcadas","gTon","number",{step:"0.001"}],
  ["tonDesc","Tonelaje descargado","gTon","number",{step:"0.001"}],
  ["arribo","Arribo","gTie","datetime-local"],
  ["pob","POB","gTie","datetime-local"],
  ["primeraLinea","Primera Línea","gTie","datetime-local"],
  ["inicio","Inicio carguío","gTie","datetime-local"],
  ["termino","Término carguío","gTie","datetime-local"],
  ["ultimaLinea","Última Línea","gTie","datetime-local"],
  ["zarpe","Zarpe","gTie","datetime-local"],
  ["detNave","Detención Nave (hrs)","gDet","number",{step:"0.01"}],
  ["detMant","Detención Mantención equipos (hrs)","gDet","number",{step:"0.01"}],
  ["detOper","Detención Operaciones (hrs)","gDet","number",{step:"0.01"}],
  ["detClima","Detención mal tiempo/lluvia (hrs)","gDet","number",{step:"0.01"}],
  ["rateContratado","Rate Contratado (Tons/día)","gCos","number"],
  ["usdProg","USD/ton Programado","gCos","number",{step:"0.01"}],
  ["usdLogrado","USD/ton Logrado","gCos","number",{step:"0.01"}],
  ["personal","Personal Total (N°)","gCos","number"],
  ["costoCLP","Costo Total (CLP)","gCos","number"],
];
// Columnas calculadas del Excel (se usan como respaldo si faltan los tiempos)
const CALC_HEADERS = {
  usoMuelleH:"Uso de muelle (Horas)", tOpH:"Tiempo de operación (Horas)", rateAlc:"Rate alcanzado (Tons/día)",
  rateNeto:"Rate Neto (Tons/día)", cumpl:"% Cumpl. Comercial", tphBruto:"Tons/hr Bruto", tphNeto:"Tons/hr Neto"
};

/* ═════════════ Utilidades ═════════════ */
const $ = s => document.querySelector(s);
const num = v => (v===null||v===undefined||v===""||isNaN(+v)) ? null : +v;
const fmt = (v,d=0) => v==null||!isFinite(v) ? "—" : v.toLocaleString("es-CL",{minimumFractionDigits:d,maximumFractionDigits:d});
const fmtK = v => v==null ? "—" : Math.abs(v)>=1e6 ? fmt(v/1e6,2)+" M" : Math.abs(v)>=1e4 ? fmt(v/1e3,1)+" k" : fmt(v);
const hrs = (a,b) => (a&&b) ? (new Date(b)-new Date(a))/36e5 : null;
const sum = (arr,f) => arr.reduce((s,r)=>s+(f(r)||0),0);
const median = arr => { const a=arr.filter(v=>v!=null&&isFinite(v)).sort((x,y)=>x-y); if(!a.length) return null; const m=a.length>>1; return a.length%2?a[m]:(a[m-1]+a[m])/2; };
const avg = arr => { const a=arr.filter(v=>v!=null&&isFinite(v)); return a.length? a.reduce((s,v)=>s+v,0)/a.length : null; };
const groupBy = (arr,f) => arr.reduce((m,r)=>{const k=f(r); (m[k]=m[k]||[]).push(r); return m;},{});
const toast = msg => { const t=$("#toast"); t.textContent=msg; t.classList.add("show"); setTimeout(()=>t.classList.remove("show"),2600); };
const uid = () => "s"+Date.now().toString(36)+Math.random().toString(36).slice(2,7);
const C = {navy:"#1F3A6E",navy2:"#4A6DB0",orange:"#E8804A",orangeL:"#F4B48F",ok:"#2E9E6B",bad:"#D2463C",warn:"#E0A400",teal:"#2A9DB5",violet:"#7B6CC9",grey:"#B8C0CF"};
const PALETTE = [C.navy,C.orange,C.teal,C.violet,C.ok,C.warn,C.navy2,C.orangeL,"#8A9BB8","#C76A9C","#5BB3A0","#A0785A"];

/* ═════════════ Normalización ═════════════ */
function normalize(r){
  const o = {...r};
  const t = s => typeof s==="string" ? s.trim() : s;
  ["nave","cliente","proceso","equipo","producto","tipoCarga","sitio","operador","mes","recalada"].forEach(k=>o[k]=t(o[k])||null);
  if (o.operador && /^reloncav/i.test(o.operador)) o.operador="Reloncaví";
  if (o.producto && /^carga general$/i.test(o.producto)) o.producto="Carga General";
  if (o.proceso && /^descarga$/i.test(o.proceso)) o.proceso="Descarga";
  if (o.cliente) o.cliente=o.cliente.toUpperCase();
  if (o.tipoCarga && /^break ?bulk$/i.test(o.tipoCarga)) o.tipoCarga="Break Bulk";
  if (o.nave) o.nave=o.nave.toUpperCase();
  const ref = o.inicio||o.primeraLinea||o.arribo||o.zarpe;
  if (!o.mes && ref) o.mes = MESES[new Date(ref).getMonth()];
  if (!o.anio && ref) o.anio = new Date(ref).getFullYear();
  o.anio = num(o.anio);
  return o;
}
// Indicadores (misma lógica del Excel)
function derive(r){
  const ton = (num(r.tonDesc)||0)+(num(r.tonEmb)||0);
  const det = (num(r.detNave)||0)+(num(r.detMant)||0)+(num(r.detOper)||0)+(num(r.detClima)||0);
  let usoMuelleH = hrs(r.arribo, r.zarpe) ?? num(r._calc?.usoMuelleH);
  if (usoMuelleH!=null && (usoMuelleH<0 || usoMuelleH>24*60)) usoMuelleH = hrs(r.primeraLinea,r.zarpe); // arribos con fecha errónea
  const tOpH = hrs(r.inicio, r.termino) ?? num(r._calc?.tOpH);
  const tReal = tOpH!=null ? Math.max(tOpH-det,0) : null;
  const rateAlc = (ton && tOpH) ? ton/(tOpH/24) : num(r._calc?.rateAlc);
  const rateNeto = (ton && tReal) ? ton/(tReal/24) : num(r._calc?.rateNeto);
  const rc = num(r.rateContratado);
  const cumpl = (rateAlc && rc) ? rateAlc/rc : num(r._calc?.cumpl);
  const tphBruto = (ton && tOpH) ? ton/tOpH : num(r._calc?.tphBruto);
  const tphNeto = (ton && tReal) ? ton/tReal : num(r._calc?.tphNeto);
  return {...r, ton, det, usoMuelleH, tOpH, tReal, rateAlc, rateNeto, cumpl, tphBruto, tphNeto};
}
function fromExcelRow(h){ // h: objeto {header: valor}
  const r = {};
  FIELDS.forEach(([k,head])=>{ let v=h[head]; if (typeof v==="string" && /^\s*-+\s*$/.test(v)) v=null; r[k]=v??null; });
  r.mes = h["Mes (limpio)"] || r.mes; r.producto = h["Producto (limpio)"] || r.producto;
  r._calc = {}; Object.entries(CALC_HEADERS).forEach(([k,head])=>{ const v=num(h[head]); if(v!=null) r._calc[k]=v; });
  FIELDS.filter(f=>f[3]==="datetime-local").forEach(([k])=>{ const v=r[k]; if(v instanceof Date) r[k]=isoLocal(v); else if(typeof v==="string" && v) r[k]= /(Z|[+-]\d\d:?\d\d)$/.test(v) ? isoLocal(new Date(v)) : v.slice(0,16); });
  r.n = num(h["N°"]);
  return normalize(r);
}
const isoLocal = d => new Date(d.getTime()-d.getTimezoneOffset()*6e4).toISOString().slice(0,16);

/* ═════════════ Almacenamiento (Firebase o local) ═════════════ */
let DATA = [];            // registros crudos
let store;                // adaptador
let user = null;
const canEdit = () => store.mode==="local" || store.mode==="sheets" || (user && (!EDITORES.length || EDITORES.includes(user.email)));


/* ---------- Google Sheets (Apps Script) ---------- */
const CALC_OUT = {usoMuelleH:"Uso de muelle (Horas)",tOpH:"Tiempo de operación (Horas)",tReal:"Tiempo real de operación (hrs)",rateAlc:"Rate alcanzado (Tons/día)",rateNeto:"Rate Neto (Tons/día)",cumpl:"% Cumpl. Comercial",tphBruto:"Tons/hr Bruto",tphNeto:"Tons/hr Neto"};
function toSheetRow(r){
  const d=derive(normalize(r)); const o={ID:r.id};
  FIELDS.forEach(([k,h])=>o[h]=r[k]??null);
  Object.entries(CALC_OUT).forEach(([k,h])=>o[h]= d[k]!=null&&isFinite(d[k]) ? Math.round(d[k]*1000)/1000 : null);
  o["Actualizado"]=r.actualizado||new Date().toISOString().slice(0,16);
  return o;
}
function editKey(force){
  let k=null; try{ k=localStorage.getItem("htp_sheet_key"); }catch(e){}
  if(!k||force){ k=prompt("Clave de edición de la planilla HTP:"); if(!k) throw new Error("Se requiere la clave de edición"); try{ localStorage.setItem("htp_sheet_key",k); }catch(e){} }
  return k;
}
async function sheetPost(body, retry=true){
  const res=await fetch(SHEETS_API_URL,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({...body,key:editKey(false)})}).then(r=>r.json());
  if(!res.ok){ if(/clave/i.test(res.error)&&retry){ editKey(true); return sheetPost(body,false); } throw new Error(res.error); }
  return res;
}
async function sheetLoad(){
  const res=await fetch(SHEETS_API_URL+(SHEETS_API_URL.includes("?")?"&":"?")+"t="+Date.now()).then(r=>r.json());
  if(!res.ok) throw new Error(res.error);
  DATA=res.rows.filter(h=>h.Nave).map((h,i)=>({...fromExcelRow(h),id:String(h.ID||"g"+(i+1))}));
  if(!DATA.length) showBanner(`La planilla de Google Sheets está vacía. <button class="btn sm navy" id="btnSeed">Cargar los ${SEED.length} servicios del Excel inicial</button>`);
  else hideBanner();
  $("#fuente").textContent=`Google Sheets · actualizado ${new Date().toLocaleTimeString("es-CL",{hour:"2-digit",minute:"2-digit"})}`;
  refresh();
}

async function initStore(){
  if (SHEETS_API_URL){
    store = {
      mode:"sheets",
      save: async r => { await sheetPost({action:"save",row:toSheetRow(r)}); await sheetLoad(); },
      remove: async id => { await sheetPost({action:"delete",id}); await sheetLoad(); },
      bulk: async rows => { for(let i=0;i<rows.length;i+=50) await sheetPost({action:"bulk",rows:rows.slice(i,i+50).map(toSheetRow)}); await sheetLoad(); },
    };
    $("#modo").className="status cloud"; $("#modo").textContent="GOOGLE SHEETS · EN LÍNEA";
    $("#btnLogin").style.display=""; $("#btnLogin").textContent="↻ Actualizar";
    $("#btnLogin").onclick=()=>sheetLoad().then(()=>toast("Datos actualizados desde Google Sheets")).catch(e=>toast("Error: "+e.message));
    setInterval(()=>{ if(!document.hidden && !$("#modal").classList.contains("open")) sheetLoad().catch(()=>{}); },120000);
    document.addEventListener("visibilitychange",()=>{ if(!document.hidden) sheetLoad().catch(()=>{}); });
    try{ await sheetLoad(); }catch(e){ showBanner("No se pudo leer Google Sheets: "+e.message+". Revisa la URL en config.js y que la app web tenga acceso «Cualquier persona»."); DATA=SEED.map(r=>({...r})); refresh(); }
    return;
  }
  if (firebaseConfig){
    const V="10.12.2", base=`https://www.gstatic.com/firebasejs/${V}/`;
    const [{initializeApp},fs,au] = await Promise.all([import(base+"firebase-app.js"),import(base+"firebase-firestore.js"),import(base+"firebase-auth.js")]);
    const app = initializeApp(firebaseConfig);
    const db = fs.getFirestore(app); const auth = au.getAuth(app);
    const col = fs.collection(db,"servicios_naves");
    store = {
      mode:"cloud",
      save: async r => { const {id,...rest}=r; await fs.setDoc(fs.doc(col,id), JSON.parse(JSON.stringify(rest))); },
      remove: async id => fs.deleteDoc(fs.doc(col,id)),
      bulk: async rows => { for (let i=0;i<rows.length;i+=400){ const b=fs.writeBatch(db); rows.slice(i,i+400).forEach(r=>{const {id,...rest}=r; b.set(fs.doc(col,id),JSON.parse(JSON.stringify(rest)));}); await b.commit(); } },
      login: () => au.signInWithPopup(auth,new au.GoogleAuthProvider()),
      logout: () => au.signOut(auth),
    };
    au.onAuthStateChanged(auth,u=>{ user=u; paintAuth(); });
    fs.onSnapshot(col, snap => {
      DATA = snap.docs.map(d=>({id:d.id,...d.data()}));
      if (!DATA.length) showBanner(`La base de datos en Firebase está vacía. <button class="btn sm navy" id="btnSeed">Cargar los ${SEED.length} servicios del Excel inicial</button>`);
      else hideBanner();
      refresh();
    }, err => { showBanner("Error al leer Firebase: "+err.message); });
    $("#modo").className="status cloud"; $("#modo").textContent="FIREBASE · EN LÍNEA"; $("#fuente").textContent="Firebase Firestore";
    $("#btnLogin").style.display="";
  } else {
    const KEY="htp_servicios_naves_v1";
    let saved=null; try{ saved=JSON.parse(localStorage.getItem(KEY)); }catch(e){}
    DATA = saved && saved.length ? saved : SEED.map(r=>({...r}));
    const persist = () => { try{ localStorage.setItem(KEY,JSON.stringify(DATA)); }catch(e){} };
    store = {
      mode:"local",
      save: async r => { const i=DATA.findIndex(x=>x.id===r.id); i>=0?DATA[i]=r:DATA.push(r); persist(); refresh(); },
      remove: async id => { DATA=DATA.filter(x=>x.id!==id); persist(); refresh(); },
      bulk: async rows => { const m=new Map(DATA.map(r=>[r.id,r])); rows.forEach(r=>m.set(r.id,r)); DATA=[...m.values()]; persist(); refresh(); },
    };
    showBanner("Modo local: los servicios se guardan sólo en este navegador. Para compartirlos con el equipo, completa <b>firebase-config.js</b> (ver README).");
    refresh();
  }
}
function paintAuth(){
  const b=$("#btnLogin");
  if (store.mode!=="cloud") return;
  b.textContent = user ? `Salir (${user.email.split("@")[0]})` : "Iniciar sesión";
  b.onclick = () => user ? store.logout() : store.login().catch(e=>toast("No se pudo iniciar sesión: "+e.code));
}
const showBanner = html => { const b=$("#banner"); b.innerHTML=html; b.style.display="block"; const s=$("#btnSeed"); if(s) s.onclick=seedCloud; };
const hideBanner = () => $("#banner").style.display="none";
async function seedCloud(){
  if (!canEdit()) return toast("Inicia sesión con un correo autorizado para cargar datos");
  await store.bulk(SEED); toast(`${SEED.length} servicios cargados en Firebase`);
}

/* ═════════════ Filtros ═════════════ */
const FILTERS = {fAnio:"anio",fMes:"mes",fProceso:"proceso",fProducto:"producto",fCliente:"cliente",fSitio:"sitio"};
function fillFilters(rows){
  Object.entries(FILTERS).forEach(([id,k])=>{
    const sel=$("#"+id), cur=sel.value;
    let vals=[...new Set(rows.map(r=>r[k]).filter(v=>v!=null&&v!==""))];
    vals = k==="mes" ? MESES.filter(m=>vals.includes(m)) : vals.sort((a,b)=>k==="anio"?b-a:String(a).localeCompare(b));
    sel.innerHTML=`<option value="">Todos</option>`+vals.map(v=>`<option${String(v)===cur?" selected":""}>${v}</option>`).join("");
  });
  // datalists del formulario
  [["dlNave","nave"],["dlCliente","cliente"],["dlProducto","producto"],["dlEquipo","equipo"],["dlOperador","operador"]].forEach(([id,k])=>{
    $("#"+id).innerHTML=[...new Set(rows.map(r=>r[k]).filter(Boolean))].sort().map(v=>`<option value="${v}">`).join("");
  });
}
function filtered(){
  return VIEW.filter(r=>Object.entries(FILTERS).every(([id,k])=>{const v=$("#"+id).value; return !v || String(r[k])===v;}));
}

/* ═════════════ Render ═════════════ */
let VIEW=[]; const charts={};
function refresh(){
  VIEW = DATA.map(r=>derive(normalize(r)));
  fillFilters(VIEW);
  render();
}
function chart(id,cfg){
  if (charts[id]) charts[id].destroy();
  cfg.options = Object.assign({responsive:true,maintainAspectRatio:false,animation:{duration:350},
    plugins:{legend:{labels:{boxWidth:12,font:{family:"Montserrat",size:11}}},tooltip:{titleFont:{family:"Montserrat"},bodyFont:{family:"Montserrat"}}}}, cfg.options||{});
  charts[id]=new Chart(document.getElementById(id),cfg);
}
Chart.defaults.font.family="Montserrat, system-ui, -apple-system, Segoe UI, sans-serif"; Chart.defaults.color="#667089"; Chart.defaults.borderColor="#E8ECF3";
const tonTip = {callbacks:{label:c=>`${c.dataset.label||c.label}: ${fmt(c.parsed.y??c.parsed.x??c.parsed)} t`}};

function kpi(l,v,s="",cls=""){ return `<div class="kpi ${cls}"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`; }
function monthKeys(rows){ // orden cronológico año-mes
  const ks=[...new Set(rows.filter(r=>r.anio&&r.mes).map(r=>r.anio*100+MESES.indexOf(r.mes)))].sort((a,b)=>a-b);
  return ks.map(k=>({k,label:MESES[k%100].slice(0,3)+" "+String(Math.floor(k/100)).slice(2)}));
}
const mk = r => r.anio*100+MESES.indexOf(r.mes);

function render(){
  const rows=filtered();
  const desc=sum(rows,r=>num(r.tonDesc)), emb=sum(rows,r=>num(r.tonEmb));
  const naves=rows.filter(r=>r.ton>0||r.proceso);

  /* ---------- RESUMEN ---------- */
  const months=monthKeys(rows); const nM=months.length||1;
  // delta vs período anterior equivalente (mismo filtro sin año → año anterior)
  const anio=$("#fAnio").value; let deltaTxt="";
  if (anio){ const prev=VIEW.filter(r=>r.anio==anio-1 && Object.entries(FILTERS).every(([id,k])=>id==="fAnio"||!$("#"+id).value||String(r[k])===$("#"+id).value));
    const pt=sum(prev,r=>r.ton); if(pt){ const d=(desc+emb-pt)/pt; deltaTxt=`<span class="delta ${d>=0?"up":"dn"}">${d>=0?"▲":"▼"} ${fmt(Math.abs(d)*100,1)}%</span> vs ${anio-1}`; } }
  $("#kpisResumen").innerHTML =
    kpi("Total transferido",fmtK(desc+emb)+" t",deltaTxt||`${fmt(nM)} meses con operación`,"o")+
    kpi("Descargado",fmtK(desc)+" t",`${fmt(rows.filter(r=>num(r.tonDesc)>0).length)} naves`)+
    kpi("Embarcado",fmtK(emb)+" t",`${fmt(rows.filter(r=>num(r.tonEmb)>0).length)} naves`,"t")+
    kpi("Naves atendidas",fmt(rows.length),`${fmt(rows.length/nM,1)} naves / mes`)+
    kpi("Promedio por nave",fmtK((desc+emb)/(rows.filter(r=>r.ton>0).length||1))+" t","sólo naves con carga")+
    kpi("Cumplimiento comercial",fmt((median(rows.map(r=>r.cumpl))||0)*100,1)+"%","mediana rate alcanzado / contratado",(median(rows.map(r=>r.cumpl))||0)>=1?"g":"r");

  const gm=groupBy(rows.filter(r=>r.anio&&r.mes),mk);
  chart("chMes",{type:"bar",data:{labels:months.map(m=>m.label),datasets:[
    {label:"Descarga",data:months.map(m=>sum(gm[m.k]||[],r=>num(r.tonDesc))),backgroundColor:C.navy,stack:"t",borderRadius:3,order:2},
    {label:"Embarque",data:months.map(m=>sum(gm[m.k]||[],r=>num(r.tonEmb))),backgroundColor:C.orange,stack:"t",borderRadius:3,order:2},
    {label:"N° naves",type:"line",data:months.map(m=>(gm[m.k]||[]).length),yAxisID:"y2",borderColor:C.teal,backgroundColor:C.teal,tension:.3,pointRadius:3,order:1}]},
    options:{scales:{x:{stacked:true,grid:{display:false}},y:{stacked:true,ticks:{callback:v=>fmtK(v)}},y2:{position:"right",grid:{display:false},beginAtZero:true,ticks:{precision:0}}},
      plugins:{tooltip:{callbacks:{label:c=>c.dataset.yAxisID==="y2"?`Naves: ${c.parsed.y}`:`${c.dataset.label}: ${fmt(c.parsed.y)} t`}}}}});

  const gp=groupBy(rows.filter(r=>r.ton>0),r=>r.proceso||"Sin dato");
  const pk=Object.keys(gp).sort((a,b)=>sum(gp[b],r=>r.ton)-sum(gp[a],r=>r.ton));
  chart("chProceso",{type:"doughnut",data:{labels:pk,datasets:[{data:pk.map(k=>sum(gp[k],r=>r.ton)),backgroundColor:[C.navy,C.orange,C.teal,C.violet,C.grey],borderWidth:2,borderColor:"#fff"}]},
    options:{cutout:"62%",plugins:{legend:{position:"bottom"},tooltip:{callbacks:{label:c=>`${c.label}: ${fmt(c.parsed)} t (${fmt(c.parsed/(desc+emb)*100,1)}%)`}}}}});

  const gprod=groupBy(rows.filter(r=>r.ton>0),r=>r.producto||"Sin dato");
  const prodK=Object.keys(gprod).sort((a,b)=>sum(gprod[b],r=>r.ton)-sum(gprod[a],r=>r.ton)).slice(0,10);
  chart("chProducto",{type:"bar",data:{labels:prodK,datasets:[{label:"Toneladas",data:prodK.map(k=>sum(gprod[k],r=>r.ton)),backgroundColor:prodK.map((_,i)=>i===0?C.orange:C.navy),borderRadius:4}]},
    options:{indexAxis:"y",plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${fmt(c.parsed.x)} t · ${gprod[c.label].length} naves`}}},scales:{x:{ticks:{callback:v=>fmtK(v)}},y:{grid:{display:false}}}}});

  const ops=[...new Set(rows.map(r=>r.operador||"Sin dato"))];
  chart("chSitio",{type:"bar",data:{labels:["Norte","Sur"],datasets:ops.map((o,i)=>({label:o,data:["Norte","Sur"].map(s=>rows.filter(r=>r.sitio===s&&(r.operador||"Sin dato")===o).length),backgroundColor:[C.navy,C.orange,C.grey,C.teal][i%4],borderRadius:3}))},
    options:{scales:{x:{stacked:true,grid:{display:false}},y:{stacked:true,ticks:{precision:0}}}}});

  // YoY: ignora filtro de año
  const rowsNoYear=VIEW.filter(r=>Object.entries(FILTERS).every(([id,k])=>id==="fAnio"||!$("#"+id).value||String(r[k])===$("#"+id).value));
  const years=[...new Set(rowsNoYear.map(r=>r.anio).filter(Boolean))].sort();
  chart("chYoY",{type:"line",data:{labels:MESES.map(m=>m.slice(0,3)),datasets:years.map((y,i)=>{let acc=0;return{label:String(y),
      data:(()=>{const last=Math.max(...rowsNoYear.filter(r=>r.anio===y&&r.ton>0).map(r=>MESES.indexOf(r.mes)));return MESES.map((m,i)=>{acc+=sum(rowsNoYear.filter(r=>r.anio===y&&r.mes===m),r=>r.ton); return i<=last&&acc?acc:null;});})(),
      borderColor:[C.grey,C.navy,C.orange,C.teal][i%4],backgroundColor:[C.grey,C.navy,C.orange,C.teal][i%4],tension:.25,borderWidth:i===years.length-1?3:2,pointRadius:2,spanGaps:true};})},
    options:{scales:{y:{ticks:{callback:v=>fmtK(v)}},x:{grid:{display:false}}},plugins:{tooltip:tonTip}}});

  /* ---------- PRODUCTIVIDAD ---------- */
  const withRate=rows.filter(r=>r.rateAlc&&r.rateContratado);
  const ok=withRate.filter(r=>r.cumpl>=1).length;
  $("#kpisProd").innerHTML=
    kpi("Rate alcanzado prom.",fmt(avg(rows.map(r=>r.rateAlc)))+" t/d","tons por día de operación","o")+
    kpi("Rate neto prom.",fmt(avg(rows.map(r=>r.rateNeto)))+" t/d","descontando detenciones")+
    kpi("Tons/hr bruto (mediana)",fmt(median(rows.map(r=>r.tphBruto)),1),"rendimiento horario","t")+
    kpi("Naves sobre contrato",`${ok} / ${withRate.length}`,`${fmt(withRate.length?ok/withRate.length*100:0,0)}% cumplen rate`,"g")+
    kpi("Cumplimiento comercial",fmt((median(rows.map(r=>r.cumpl))||0)*100,1)+"%","mediana del período");

  const rr=withRate.slice().sort((a,b)=>(a.inicio||"").localeCompare(b.inicio||"")).slice(-25);
  chart("chRate",{type:"bar",data:{labels:rr.map(r=>r.nave),datasets:[
    {label:"Rate alcanzado",data:rr.map(r=>r.rateAlc),backgroundColor:rr.map(r=>r.cumpl>=1?C.navy:C.orange),borderRadius:3,order:2},
    {label:"Rate contratado",type:"line",data:rr.map(r=>num(r.rateContratado)),borderColor:C.bad,backgroundColor:C.bad,borderDash:[5,4],pointStyle:"line",pointRadius:14,pointBorderWidth:3,pointHoverRadius:16,showLine:false,order:1}]},
    options:{scales:{x:{grid:{display:false},ticks:{maxRotation:60,minRotation:45,font:{size:10}}},y:{ticks:{callback:v=>fmtK(v)}}},
      plugins:{tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmt(c.parsed.y)} t/d`,afterBody:it=>{const r=rr[it[0].dataIndex];return `Cumplimiento: ${fmt(r.cumpl*100,1)}% · ${r.producto||""}`;}}}}}});

  const gmc=groupBy(withRate,mk); const mc=monthKeys(withRate);
  chart("chCumpl",{type:"bar",data:{labels:mc.map(m=>m.label),datasets:[
    {label:"% cumplimiento",data:mc.map(m=>median(gmc[m.k].map(r=>r.cumpl))*100),backgroundColor:mc.map(m=>median(gmc[m.k].map(r=>r.cumpl))>=1?C.ok:C.orange),borderRadius:3},
    {label:"Meta 100%",type:"line",data:mc.map(()=>100),borderColor:C.navy,borderDash:[6,4],pointRadius:0,borderWidth:1.5}]},
    options:{scales:{y:{ticks:{callback:v=>v+"%"}},x:{grid:{display:false}}},plugins:{tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmt(c.parsed.y,1)}%`}}}}});

  const gtp=groupBy(rows.filter(r=>r.tphBruto),r=>r.producto||"Sin dato");
  const tpk=Object.keys(gtp).sort((a,b)=>median(gtp[b].map(r=>r.tphBruto))-median(gtp[a].map(r=>r.tphBruto))).slice(0,10);
  chart("chTph",{type:"bar",data:{labels:tpk,datasets:[{label:"Tons/hr",data:tpk.map(k=>median(gtp[k].map(r=>r.tphBruto))),backgroundColor:C.teal,borderRadius:4}]},
    options:{indexAxis:"y",plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${fmt(c.parsed.x,1)} t/hr · ${gtp[c.label].length} naves`}}},scales:{y:{grid:{display:false}}}}});

  const rankList=(el,arr,val,fmtv)=>{const mx=Math.max(...arr.map(val),1);
    $(el).innerHTML=arr.length?arr.map((r,i)=>`<li><span class="i">${i+1}</span><div><b>${r.nave}</b> <span style="color:var(--muted)">· ${r.producto||""} · ${r.mes||""} ${r.anio||""}</span><div class="bar" style="width:${Math.max(val(r)/mx*100,3)}%"></div></div><span class="val">${fmtv(r)}</span></li>`).join(""):`<li class="empty">Sin datos</li>`;};
  rankList("#rankTop",rows.filter(r=>r.tphBruto).sort((a,b)=>b.tphBruto-a.tphBruto).slice(0,10),r=>r.tphBruto,r=>fmt(r.tphBruto,0)+" t/h");
  rankList("#rankLow",withRate.filter(r=>r.cumpl<1).sort((a,b)=>a.cumpl-b.cumpl).slice(0,10),r=>1-r.cumpl,r=>fmt(r.cumpl*100,1)+"%");

  /* ---------- MUELLE ---------- */
  const um=rows.filter(r=>r.usoMuelleH>0);
  const detT={Nave:sum(rows,r=>num(r.detNave)),"Mantención equipos":sum(rows,r=>num(r.detMant)),Operaciones:sum(rows,r=>num(r.detOper)),"Mal tiempo / lluvia":sum(rows,r=>num(r.detClima))};
  const totDet=Object.values(detT).reduce((a,b)=>a+b,0), totOp=sum(rows,r=>r.tOpH), totUM=sum(um,r=>r.usoMuelleH);
  $("#kpisMuelle").innerHTML=
    kpi("Uso de muelle",fmt(totUM)+" h",`${fmt(totUM/24,1)} días`,"o")+
    kpi("Estadía promedio",fmt(avg(um.map(r=>r.usoMuelleH)),1)+" h","arribo → zarpe por nave")+
    kpi("Tiempo de operación",fmt(totOp)+" h",`${fmt(totUM?totOp/totUM*100:0,1)}% del uso de muelle`,"t")+
    kpi("Horas de detención",fmt(totDet,1)+" h",`${fmt(totOp?totDet/totOp*100:0,1)}% del tiempo operativo`,"r")+
    kpi("Ocupación Sitio Sur",fmt(sum(um.filter(r=>r.sitio==="Sur"),r=>r.usoMuelleH))+" h",`Norte: ${fmt(sum(um.filter(r=>r.sitio==="Norte"),r=>r.usoMuelleH))} h`);

  const gum=groupBy(um.filter(r=>r.anio&&r.mes),mk); const mum=monthKeys(um);
  chart("chMuelle",{type:"bar",data:{labels:mum.map(m=>m.label),datasets:["Sur","Norte"].map((s,i)=>({label:"Sitio "+s,data:mum.map(m=>sum((gum[m.k]||[]).filter(r=>r.sitio===s),r=>r.usoMuelleH)),backgroundColor:i?C.orange:C.navy,stack:"m",borderRadius:3}))},
    options:{scales:{x:{stacked:true,grid:{display:false}},y:{stacked:true,title:{display:true,text:"horas"}}},plugins:{tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmt(c.parsed.y,1)} h`}}}}});
  chart("chDet",{type:"doughnut",data:{labels:Object.keys(detT),datasets:[{data:Object.values(detT),backgroundColor:[C.navy,C.orange,C.teal,C.violet],borderColor:"#fff",borderWidth:2}]},
    options:{cutout:"60%",plugins:{legend:{position:"bottom"},tooltip:{callbacks:{label:c=>`${c.label}: ${fmt(c.parsed,1)} h`}}}}});
  const tm=um.filter(r=>r.tOpH).sort((a,b)=>(a.inicio||"").localeCompare(b.inicio||"")).slice(-25);
  chart("chTiempo",{type:"bar",data:{labels:tm.map(r=>r.nave),datasets:[
    {label:"Operación efectiva",data:tm.map(r=>r.tReal),backgroundColor:C.navy},
    {label:"Detenciones",data:tm.map(r=>r.det),backgroundColor:C.bad},
    {label:"Sin operación (espera, amarre, zarpe)",data:tm.map(r=>Math.max(r.usoMuelleH-r.tOpH,0)),backgroundColor:C.grey}]},
    options:{scales:{x:{stacked:true,grid:{display:false},ticks:{maxRotation:60,minRotation:45,font:{size:10}}},y:{stacked:true,title:{display:true,text:"horas"}}},plugins:{tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmt(c.parsed.y,1)} h`}}}}});

  /* ---------- CLIENTES ---------- */
  const gc=groupBy(rows.filter(r=>r.ton>0),r=>r.cliente||"Sin dato");
  const ck=Object.keys(gc).sort((a,b)=>sum(gc[b],r=>r.ton)-sum(gc[a],r=>r.ton));
  const top=ck[0];
  $("#kpisCli").innerHTML=
    kpi("Clientes activos",fmt(ck.length),"con carga transferida","o")+
    kpi("Principal cliente",top||"—",top?`${fmt(sum(gc[top],r=>r.ton)/(desc+emb)*100,1)}% del tonelaje`:"")+
    kpi("Productos distintos",fmt(new Set(rows.map(r=>r.producto).filter(Boolean)).size),"","t")+
    kpi("Eslora promedio",fmt(avg(rows.map(r=>num(r.eslora))),1)+" m",`Máx. ${fmt(Math.max(0,...rows.map(r=>num(r.eslora)||0)),1)} m`);
  const ckTop=ck.slice(0,8), resto=ck.slice(8);
  chart("chCliente",{type:"doughnut",data:{labels:[...ckTop,...(resto.length?["Otros"]:[])],datasets:[{data:[...ckTop.map(k=>sum(gc[k],r=>r.ton)),...(resto.length?[resto.reduce((s,k)=>s+sum(gc[k],r=>r.ton),0)]:[])],backgroundColor:PALETTE,borderColor:"#fff",borderWidth:2}]},
    options:{cutout:"55%",plugins:{legend:{position:"right"},tooltip:{callbacks:{label:c=>`${c.label}: ${fmt(c.parsed)} t (${fmt(c.parsed/(desc+emb)*100,1)}%)`}}}}});
  const gt=groupBy(rows,r=>r.tipoCarga||"Sin dato"); const tk=Object.keys(gt).sort((a,b)=>gt[b].length-gt[a].length);
  chart("chTipo",{type:"bar",data:{labels:tk,datasets:[{label:"Naves",data:tk.map(k=>gt[k].length),backgroundColor:tk.map((_,i)=>PALETTE[i]),borderRadius:4}]},
    options:{plugins:{legend:{display:false}},scales:{y:{ticks:{precision:0}},x:{grid:{display:false}}}}});
  // matriz
  const prods=[...new Set(rows.filter(r=>r.ton>0).map(r=>r.producto||"Sin dato"))].sort((a,b)=>sum(gprod[b]||[],r=>r.ton)-sum(gprod[a]||[],r=>r.ton));
  const maxCell=Math.max(1,...ck.flatMap(c=>prods.map(p=>sum(gc[c].filter(r=>(r.producto||"Sin dato")===p),r=>r.ton))));
  $("#tblMatriz").innerHTML=`<thead><tr><th>Cliente</th>${prods.map(p=>`<th>${p}</th>`).join("")}<th>Total</th></tr></thead><tbody>`+
    ck.map(c=>`<tr><td><b>${c}</b></td>${prods.map(p=>{const v=sum(gc[c].filter(r=>(r.producto||"Sin dato")===p),r=>r.ton);return `<td class="num" style="background:${v?`rgba(232,128,74,${.08+.6*v/maxCell})`:""}">${v?fmt(v):""}</td>`;}).join("")}<td class="num"><b>${fmt(sum(gc[c],r=>r.ton))}</b></td></tr>`).join("")+"</tbody>";

  renderTable(rows);
}

/* ═════════════ Tabla de servicios ═════════════ */
const COLS=[["nave","Nave"],["anio","Año"],["mes","Mes"],["cliente","Cliente"],["proceso","Proceso"],["producto","Producto"],["sitio","Sitio"],["equipo","Equipo"],
  ["ton","Toneladas",1],["inicio","Inicio carguío"],["termino","Término"],["usoMuelleH","Uso muelle (h)",1],["tOpH","T. operación (h)",1],["det","Detenciones (h)",1],
  ["rateContratado","Rate contr.",1],["rateAlc","Rate alcanzado",1],["cumpl","% Cumpl.",1],["tphBruto","Tons/hr",1]];
let sortK="inicio", sortDir=-1;
function renderTable(rows){
  const q=$("#buscar").value.trim().toLowerCase();
  let list=rows.filter(r=>!q||[r.nave,r.cliente,r.producto,r.equipo,r.operador,r.proceso].join(" ").toLowerCase().includes(q));
  list.sort((a,b)=>{const x=a[sortK],y=b[sortK]; if(x==null) return 1; if(y==null) return -1; return (typeof x==="number"?x-y:String(x).localeCompare(y))*sortDir;});
  $("#conteo").textContent=`${list.length} servicios`;
  const cell=(r,[k,,n])=>{let v=r[k];
    if(k==="proceso") return `<td><span class="pill ${/desc/i.test(v)?"desc":/emb/i.test(v)?"emb":"otro"}">${v||"—"}</span></td>`;
    if(k==="cumpl") return `<td class="num"><span class="cump ${v>=1?"ok":v>=.9?"mid":"bad"}">${v!=null?fmt(v*100,1)+"%":"—"}</span></td>`;
    if(k==="inicio"||k==="termino") return `<td>${v?new Date(v).toLocaleString("es-CL",{day:"2-digit",month:"2-digit",year:"2-digit",hour:"2-digit",minute:"2-digit"}):"—"}</td>`;
    if(n) return `<td class="num">${v!=null&&v!==0?fmt(+v,k==="tOpH"||k==="usoMuelleH"||k==="det"?1:0):"—"}</td>`;
    return `<td>${v??"—"}</td>`;};
  $("#tbl").innerHTML=`<thead><tr>${COLS.map(([k,l])=>`<th data-k="${k}">${l}${sortK===k?(sortDir>0?" ▲":" ▼"):""}</th>`).join("")}<th></th></tr></thead><tbody>`+
    (list.length?list.map(r=>`<tr data-id="${r.id}">${COLS.map(c=>cell(r,c)).join("")}<td><button class="btn sm" data-edit="${r.id}">Editar</button></td></tr>`).join(""):`<tr><td colspan="${COLS.length+1}" class="empty">Sin servicios para el filtro</td></tr>`)+"</tbody>";
}
$("#tbl").addEventListener("click",e=>{
  const th=e.target.closest("th[data-k]"); if(th){ const k=th.dataset.k; sortDir = sortK===k ? -sortDir : -1; sortK=k; renderTable(filtered()); return; }
  const tr=e.target.closest("tr[data-id]"); if(tr) openForm(DATA.find(r=>r.id===tr.dataset.id));
});
$("#buscar").addEventListener("input",()=>renderTable(filtered()));

/* ═════════════ Formulario de ingreso ═════════════ */
let editing=null;
function buildForm(){
  FIELDS.forEach(([k,label,g,type,x={}])=>{
    const id="i_"+k; let input;
    if(type==="select") input=`<select id="${id}" name="${k}"><option value=""></option>${x.opts.map(o=>`<option>${o}</option>`).join("")}</select>`;
    else input=`<input id="${id}" name="${k}" type="${type}"${x.step?` step="${x.step}"`:""}${x.list?` list="${x.list}"`:""}${x.req?" required":""}${type==="number"?' inputmode="decimal"':""}>`;
    $("#"+g).insertAdjacentHTML("beforeend",`<div class="f"><label for="${id}">${label}${x.req?" *":""}</label>${input}</div>`);
  });
  $("#form").addEventListener("input",paintCalc);
  // autocompletar año/mes según inicio de carguío
  $("#i_inicio").addEventListener("change",e=>{ const v=e.target.value; if(!v) return; const d=new Date(v);
    if(!$("#i_anio").value) $("#i_anio").value=d.getFullYear(); if(!$("#i_mes").value) $("#i_mes").value=MESES[d.getMonth()]; paintCalc(); });
}
function readForm(){
  const r={...(editing||{})};
  FIELDS.forEach(([k,,,type])=>{ const v=$("#i_"+k).value; r[k]= v==="" ? null : type==="number" ? +v : v; });
  return r;
}
function paintCalc(){
  const d=derive(normalize(readForm()));
  const items=[["Uso de muelle",d.usoMuelleH,"h",1],["Tiempo de operación",d.tOpH,"h",1],["Detenciones",d.det,"h",1],["Tiempo real operación",d.tReal,"h",1],
    ["Toneladas totales",d.ton,"t",0],["Rate alcanzado",d.rateAlc,"t/día",0],["Rate neto",d.rateNeto,"t/día",0],["% Cumpl. comercial",d.cumpl!=null?d.cumpl*100:null,"%",1],
    ["Tons/hr bruto",d.tphBruto,"t/h",1],["Tons/hr neto",d.tphNeto,"t/h",1]];
  $("#calc").innerHTML=items.map(([l,v,u,dd])=>`<div>${l}<b>${v!=null&&isFinite(v)&&v!==0?fmt(v,dd)+" "+u:"—"}</b></div>`).join("");
}
function openForm(r){
  if(!canEdit()) { toast(store.mode==="cloud"?"Inicia sesión con un correo autorizado para ingresar o editar":"Sin permisos"); if(store.mode==="cloud"&&!user) store.login().catch(()=>{}); return; }
  editing=r||null;
  $("#mTitle").textContent=r?`Editar servicio · ${r.nave}`:"Ingresar servicio de nave";
  $("#btnDel").style.display=r?"":"none";
  FIELDS.forEach(([k])=>{ $("#i_"+k).value = r && r[k]!=null ? r[k] : ""; });
  if(!r){ const now=new Date(); $("#i_anio").value=now.getFullYear(); $("#i_mes").value=MESES[now.getMonth()]; $("#i_operador").value="Reloncaví"; }
  paintCalc(); $("#modal").classList.add("open"); $("#i_nave").focus();
}
const closeForm=()=>{ $("#modal").classList.remove("open"); editing=null; };
$("#form").addEventListener("submit",async e=>{
  e.preventDefault();
  const r=normalize(readForm()); if(!r.id) r.id=uid();
  r.actualizado=new Date().toISOString(); if(user) r.actualizadoPor=user.email; r.actualizado=r.actualizado.slice(0,16);
  const btn=e.submitter||$("#form button[type=submit]"); btn.disabled=true; btn.textContent="Guardando…";
  try{ await store.save(r); toast(editing?"Servicio actualizado":"Servicio ingresado"); closeForm(); }
  catch(err){ toast("Error al guardar: "+(err.code||err.message)); }
  finally{ btn.disabled=false; btn.textContent="Guardar servicio"; }
});
$("#btnDel").onclick=async()=>{ if(!editing) return; if(!confirm(`¿Eliminar el servicio de ${editing.nave}?`)) return;
  try{ await store.remove(editing.id); toast("Servicio eliminado"); closeForm(); }catch(err){ toast("Error: "+(err.code||err.message)); } };
$("#mClose").onclick=$("#btnCancel").onclick=closeForm;
$("#modal").addEventListener("click",e=>{ if(e.target.id==="modal") closeForm(); });
document.addEventListener("keydown",e=>{ if(e.key==="Escape") closeForm(); });
$("#btnNuevo").onclick=$("#btnNuevo2").onclick=()=>openForm(null);

/* ═════════════ Importar / Exportar Excel ═════════════ */
$("#btnImport").onclick=()=>{ if(!canEdit()) return toast("Inicia sesión para importar"); $("#fileXlsx").click(); };
$("#fileXlsx").addEventListener("change",async e=>{
  const f=e.target.files[0]; if(!f) return;
  const wb=XLSX.read(await f.arrayBuffer(),{cellDates:true});
  const ws=wb.Sheets["Estadística Operacional Naves"]||wb.Sheets[wb.SheetNames[0]];
  const aoa=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
  const hi=aoa.findIndex(row=>row.includes("Nave")&&row.includes("Proceso")); if(hi<0) return toast("No se encontró la fila de encabezados (Nave, Proceso…)");
  const head=aoa[hi];
  const rows=aoa.slice(hi+1).filter(row=>row[head.indexOf("Nave")]).map(row=>{const h={}; head.forEach((k,i)=>{if(k) h[k]=row[i];}); return fromExcelRow(h);});
  // evita duplicados: misma nave + inicio (o arribo/zarpe)
  const keyOf=r=>[r.nave,r.inicio||r.arribo||r.zarpe||r.n].join("|");
  const existing=new Map(DATA.map(r=>[keyOf(normalize(r)),r.id]));
  rows.forEach(r=>r.id=existing.get(keyOf(r))||uid());
  const nuevos=rows.filter(r=>!existing.has(keyOf(r))).length;
  if(!confirm(`Se leyeron ${rows.length} servicios (${nuevos} nuevos, ${rows.length-nuevos} se actualizarán). ¿Continuar?`)) return;
  await store.bulk(rows); toast(`Importados ${rows.length} servicios`); e.target.value="";
});
$("#btnExport").onclick=()=>{
  const rows=filtered().map(r=>{const o={}; FIELDS.forEach(([k,h])=>o[h]=r[k]??"");
    Object.assign(o,{"Uso de muelle (Horas)":r.usoMuelleH,"Tiempo de operación (Horas)":r.tOpH,"Tiempo real de operación (hrs)":r.tReal,"Rate alcanzado (Tons/día)":r.rateAlc,"Rate Neto (Tons/día)":r.rateNeto,"% Cumpl. Comercial":r.cumpl,"Tons/hr Bruto":r.tphBruto,"Tons/hr Neto":r.tphNeto}); return o;});
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Estadística Naves");
  XLSX.writeFile(wb,`Estadisticas_naves_HTP_${new Date().toISOString().slice(0,10)}.xlsx`);
};

/* ═════════════ Navegación y arranque ═════════════ */
document.querySelectorAll(".tab").forEach(t=>t.onclick=()=>{
  document.querySelectorAll(".tab").forEach(x=>x.classList.toggle("active",x===t));
  document.querySelectorAll(".view").forEach(v=>v.classList.toggle("active",v.id==="v-"+t.dataset.v));
  $("#filters").style.display="";
  try{ localStorage.setItem("htp_tab",t.dataset.v); }catch(e){}
  Object.values(charts).forEach(c=>c.resize());
});
Object.keys(FILTERS).forEach(id=>$("#"+id).addEventListener("change",render));
$("#btnLimpiar").onclick=()=>{ Object.keys(FILTERS).forEach(id=>$("#"+id).value=""); render(); };

let SEED=[];
(async()=>{
  buildForm();
  const raw=await fetch("data.json").then(r=>r.json());
  SEED=raw.map((h,i)=>({...fromExcelRow(h),id:"x"+String(i+1).padStart(4,"0")}));
  await initStore();
  try{ const t=localStorage.getItem("htp_tab"); if(t) document.querySelector(`.tab[data-v="${t}"]`)?.click(); }catch(e){}
  const y=new Date().getFullYear(); if([...$("#fAnio").options].some(o=>o.value==y)){ $("#fAnio").value=y; render(); }
})();
