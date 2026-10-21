"use strict";
/* ============ NORMATIVA (Dec. 351/79 Anexo IV – Res. SRT 84/12) ============ */
const T1=[
 {id:"t1",min:100,max:null,txt:"Visión ocasional solamente (circulación, depósitos)"},
 {id:"t2",min:100,max:300,txt:"Tareas intermitentes ordinarias y fáciles, con contrastes fuertes"},
 {id:"t3",min:300,max:750,txt:"Tareas moderadamente críticas y prolongadas, con detalles medianos"},
 {id:"t4",min:750,max:1500,txt:"Tareas severas y prolongadas y de poco contraste"},
 {id:"t5",min:1500,max:3000,txt:"Tareas muy severas y prolongadas, con detalles minuciosos o muy poco contraste"},
 {id:"t6",min:3000,max:null,txt:"Tareas muy prolongadas y muy severas, detalles minuciosos (p. ej. relojería fina)"},
 {id:"t7",min:5000,max:10000,txt:"Tareas excepcionales, difíciles o importantes"}];
const T4=[[250,125],[500,250],[1000,300],[2500,500],[5000,600],[10000,700]];
const METODO="Método de la grilla o cuadrícula (Res. SRT 84/12). Se determina el índice del local K = (L·A) / [Hm·(L+A)] y el número mínimo de puntos (x+2)², ubicando la célula fotoeléctrica del instrumento sobre el plano de trabajo o a 0,80 m del solado en medios de circulación. En puestos con iluminación localizada se mide sobre el plano de la tarea.";

/* ============ ESTADO, GUARDADO Y SINCRONIZACIÓN ============ */
const CFG=window.APP_CONFIG||{};
const CLOUD=!!(CFG.supabaseUrl&&CFG.supabaseAnonKey&&window.supabase);
const sb=CLOUD?window.supabase.createClient(CFG.supabaseUrl,CFG.supabaseAnonKey,{auth:{persistSession:true,autoRefreshToken:true,detectSessionInUrl:true}}):null;
const newId=()=>(window.crypto&&crypto.randomUUID)?crypto.randomUUID():"xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g,c=>{const r=Math.random()*16|0;return(c==="x"?r:(r&3|8)).toString(16)});
const uid=()=>Math.random().toString(36).slice(2,10);
const today=()=>{const d=new Date();return new Date(d-d.getTimezoneOffset()*6e4).toISOString().slice(0,10)};
function newPoint(n,o={}){return Object.assign({id:uid(),n,hora:"",sector:"",puesto:"",tipoIlum:"Mixta",fuente:"Descarga",ilum:"General",
  req:"t3",reqMin:"",reqMax:"",metodo:"grilla",L:"",A:"",Hm:"",filas:3,cols:3,lect:[],egen:"",obs:"",prev:null},o)}
function valid(d){return !!(d&&d.meta&&(Array.isArray(d.puntos)||Array.isArray(d.sectores)||Array.isArray(d.tomas)))}
let view="list",curId=null,S=null,session=null,lib={},localOK=true,saveT=null;
const net={state:"idle",timer:null,busy:false};
function cacheKey(){return"lux_lib_"+(session?session.user.id:"local")}
function loadCache(){lib={};try{const j=JSON.parse(localStorage.getItem(cacheKey()));if(j&&typeof j==="object")lib=j}catch(e){}}
function writeCache(){try{localStorage.setItem(cacheKey(),JSON.stringify(lib));localOK=true}catch(e){localOK=false}}
function meds(){return Object.entries(lib).filter(([,x])=>!x.deleted&&valid(x.data)).map(([id,x])=>({id,...x}))
  .sort((a,b)=>(b.data.meta.fechaMed||"").localeCompare(a.data.meta.fechaMed||"")||b.updatedAt-a.updatedAt)}
function latest(){return Object.values(lib).filter(x=>!x.deleted&&valid(x.data)).sort((a,b)=>b.updatedAt-a.updatedAt)[0]}
function blankMeta(){const L=(latest()||{data:{meta:{}}}).data.meta;
  return{razon:"",cuit:"",direccion:"",localidad:"",provincia:"",cp:"",horarios:"",instrumento:L.instrumento||"Luxómetro",marca:L.marca||"",modelo:L.modelo||"",serie:L.serie||"",fechaCal:L.fechaCal||"",
    metodologia:METODO,fechaMed:today(),horaIni:"",horaFin:"",condiciones:"",cielo:"",temp:"",docCert:true,docCroquis:true,observaciones:"",
    profesional:L.profesional||"",matricula:L.matricula||"",evaluarMax:L.evaluarMax??true,prevFecha:""}}
function blankMed(){return{tipo:"ilum",updatedAt:Date.now(),meta:blankMeta(),puntos:[],conclusiones:"",recomendaciones:""}}
function setSync(){
  const el=document.getElementById("syncSt");if(!el)return;
  if(view==="auth"||view==="recover"){el.textContent="";return}
  if(!CLOUD){el.textContent=localOK?"📱 en este dispositivo":"⚠ sin guardar";return}
  const pend=Object.values(lib).some(x=>x.dirty);
  if(pend)el.textContent=navigator.onLine?"☁ sincronizando…":"📱 sin conexión";
  else el.textContent=net.state==="error"?"⚠ sin sincronizar":"☁ sincronizado";
}
function putMed(id,data,extra={}){lib[id]=Object.assign({data,updatedAt:data.updatedAt||Date.now(),dirty:true,deleted:false},extra);writeCache();setSync();schedulePush(800)}
function save(){
  if(!S||!curId)return;
  S.updatedAt=Date.now();lib[curId]={data:S,updatedAt:S.updatedAt,dirty:true,deleted:false};
  clearTimeout(saveT);saveT=setTimeout(()=>{writeCache();setSync();schedulePush(1500)},200);
}
function schedulePush(ms){if(!CLOUD||!session)return;clearTimeout(net.timer);net.timer=setTimeout(sync,ms)}
async function sync(){
  if(!CLOUD||!session||net.busy)return;
  if(!navigator.onLine){setSync();return}
  net.busy=true;
  try{
    const {data:rows,error}=await sb.from("mediciones").select("id,data,deleted,updated_at");
    if(error)throw error;
    let changedCur=false,changed=false;
    for(const r of rows||[]){
      const ru=(r.data&&r.data.updatedAt)||Date.parse(r.updated_at)||0,l=lib[r.id];
      if(!l||ru>l.updatedAt){lib[r.id]={data:r.data,updatedAt:ru,dirty:false,deleted:!!r.deleted};changed=true;if(r.id===curId)changedCur=true}
    }
    const dirty=Object.entries(lib).filter(([,x])=>x.dirty);
    if(dirty.length){
      const payload=dirty.map(([id,x])=>({id,empresa:(x.data&&x.data.meta&&x.data.meta.razon)||"",fecha:(x.data&&x.data.meta&&/^\d{4}-\d{2}-\d{2}$/.test(x.data.meta.fechaMed))?x.data.meta.fechaMed:null,
        data:x.data,deleted:!!x.deleted,updated_at:new Date(x.updatedAt).toISOString()}));
      const {error:e2}=await sb.from("mediciones").upsert(payload);
      if(e2)throw e2;
      dirty.forEach(([id,x])=>{if(lib[id]&&lib[id].updatedAt===x.updatedAt)lib[id].dirty=false});
    }
    writeCache();net.state="ok";
    if(changedCur&&view==="med"){if(lib[curId].deleted){view="list";curId=null;S=null;render()}else{S=lib[curId].data;if(!editing)render()}}
    else if(changed&&view==="list")render();
  }catch(e){net.state="error";console.warn("sync",e);schedulePush(30000)}
  finally{net.busy=false;setSync()}
}
window.addEventListener("online",()=>schedulePush(300));
window.addEventListener("offline",setSync);
document.addEventListener("visibilitychange",()=>{if(document.visibilityState==="visible")schedulePush(300)});
setInterval(()=>{if(document.visibilityState==="visible")schedulePush(0)},90000);

/* ============ ÚTILES ============ */
const $=(s,r=document)=>r.querySelector(s);
const $$=(s,r=document)=>[...r.querySelectorAll(s)];
const esc=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const num=v=>{if(v==null)return NaN;const s=String(v).trim().replace(/\s/g,"");if(!s)return NaN;return parseFloat(s.replace(/\.(?=\d{3}(\D|$))/g,"").replace(",","."))};
const fmt=(v,d=2)=>isFinite(v)?v.toLocaleString("es-AR",{minimumFractionDigits:0,maximumFractionDigits:d}):"–";
const fdate=iso=>{if(!iso)return"";const[y,m,d]=iso.split("-");return`${d}/${m}/${y}`};
const nowHM=()=>new Date().toTimeString().slice(0,5);
const buzz=()=>{try{navigator.vibrate&&navigator.vibrate(25)}catch(e){}};
function toast(msg){$$(".toast").forEach(x=>x.remove());const t=document.createElement("div");t.className="toast";t.setAttribute("role","status");t.textContent=msg;document.body.appendChild(t);setTimeout(()=>t.remove(),2600)}
function ask(msg,{ok="Aceptar",cancel="Cancelar",danger=false}={}){
  return new Promise(res=>{
    const s=document.createElement("div");s.className="scrim";
    s.innerHTML=`<div class="dlg" role="dialog" aria-modal="true"><p>${esc(msg)}</p><div class="row" style="justify-content:flex-end">${cancel?`<button class="btn" data-a="0">${esc(cancel)}</button>`:""}<button class="btn ${danger?"d":"p"}" data-a="1">${esc(ok)}</button></div></div>`;
    s.addEventListener("click",e=>{const b=e.target.closest("[data-a]");if(b||e.target===s){s.remove();res(!!(b&&b.dataset.a==="1"))}});
    document.body.appendChild(s);$("[data-a='1']",s).focus();
  });
}

/* ============ CÁLCULOS ============ */
function indiceK(p){const L=num(p.L),A=num(p.A),H=num(p.Hm);if(!(L>0&&A>0&&H>0))return null;return(L*A)/(H*(L+A))}
function xFromK(K){if(K===null)return null;if(K>=3)return 4;return Math.max(1,Math.ceil(K-1e-9))}
function reqOf(p){if(p.req==="custom"){const mn=num(p.reqMin),mx=num(p.reqMax);return{min:isFinite(mn)?mn:null,max:isFinite(mx)?mx:null}}const t=T1.find(t=>t.id===p.req)||T1[2];return{min:t.min,max:t.max}}
const reqStr=r=>r.min==null?"–":(r.max?`${r.min}-${r.max}`:`${r.min}`);
function t4ref(e){if(!(e>0))return null;let ref=null;for(const[l,g]of T4){if(e>=l)ref=g}if(ref===null)ref=T4[0][1]*e/T4[0][0];return ref}
function calc(p){
  const vals=(p.lect||[]).map(num).filter(isFinite);const n=vals.length,K=indiceK(p),x=xFromK(K),nmin=x?(x+2)**2:null,r=reqOf(p);
  const o={n,K,x,nmin,req:r,vals};if(!n)return o;
  o.emed=vals.reduce((a,b)=>a+b,0)/n;o.emin=Math.min(...vals);o.emax=Math.max(...vals);
  o.unifOK=o.emin>=o.emed/2;o.minOK=r.min==null?null:o.emed>=r.min;o.maxOK=((S?S.meta.evaluarMax:true)&&r.max)?o.emed<=r.max:true;
  o.puntosOK=p.metodo==="grilla"?(nmin?n>=nmin:null):true;
  if(p.ilum!=="General"){const g=num(p.egen);o.t4=t4ref(o.emed);o.egen=isFinite(g)?g:null;o.t4OK=(o.egen!=null&&o.t4!=null)?o.egen>=o.t4:null}
  o.cumple=o.minOK!==false&&o.maxOK&&(n<2||o.unifOK)&&o.t4OK!==false;return o;
}
function status(p){const c=calc(p);
  if(!c.n)return{cls:"na",txt:"Sin medir",c};if(c.minOK===false)return{cls:"bad",txt:"Bajo mínimo",c};
  if(!c.maxOK)return{cls:"warn",txt:"Supera máximo",c};if(c.n>=2&&!c.unifOK)return{cls:"bad",txt:"No uniforme",c};
  if(c.t4OK===false)return{cls:"bad",txt:"Gral. insuficiente",c};if(c.puntosOK===false)return{cls:"warn",txt:"Faltan lecturas",c};
  return{cls:"ok",txt:"Cumple",c}}

/* ============ NAVEGACIÓN ============ */
let tab="pts",editing=null,kp={cell:0,buf:""},wake=null;
$("#tabs").addEventListener("click",e=>{const b=e.target.closest("button");if(!b)return;editing=null;go(b.dataset.t)});
function go(t){tab=t;$$("#tabs button").forEach(b=>b.classList.toggle("on",b.dataset.t===t));render();scrollTo(0,0)}
async function wakeOn(){try{if(navigator.wakeLock&&!wake){wake=await navigator.wakeLock.request("screen");wake.addEventListener("release",()=>{wake=null})}}catch(e){}}
function wakeOff(){try{wake&&wake.release()}catch(e){}wake=null}
document.addEventListener("visibilitychange",()=>{if(document.visibilityState==="visible"&&editing)wakeOn()});
function render(){
  const m=$("#main"),inMed=view==="med"&&!!S;
  $("#tabs").hidden=!inMed;$("#hBack").hidden=!inMed;
  const mod=inMed&&S.tipo&&MODS[S.tipo];
  $("#hdrTitle").textContent=inMed?(S.meta.razon||"Estudio sin empresa"):"Estudios de Higiene y Seguridad";
  const TL=mod?mod.tabs:["Datos","Medir","Análisis","Exportar"];$$("#tabs button").forEach((b,i)=>{b.lastChild.textContent=TL[i]});
  if(view==="auth")m.innerHTML=vAuth();
  else if(view==="recover")m.innerHTML=vRecover();
  else if(!inMed)m.innerHTML=vList();
  else if(mod)m.innerHTML=mod.view(tab);
  else if(tab==="est")m.innerHTML=vEst();
  else if(tab==="pts")m.innerHTML=editing?vEditor(cur()):vPuntos();
  else if(tab==="con")m.innerHTML=vCon();
  else m.innerHTML=vExp();
  if(inMed&&tab==="pts"&&editing)wakeOn();else wakeOff();
  if(view==="auth"||view==="recover")bindAuth();else if(!inMed)bindList();else{bind();if(mod)mod.bind(tab)}
  bindTop();setSync();
}
const cur=()=>S&&S.puntos.find(p=>p.id===editing);

/* ============ CAMPOS ============ */
function getPath(path){return path.split(".").reduce((o,k)=>o==null?o:o[k],path.startsWith("p.")?{p:cur()}:S)}
function setPath(path,val){const parts=path.split(".");let o=path.startsWith("p.")?{p:cur()}:S;for(let i=0;i<parts.length-1;i++)o=o[parts[i]];o[parts.at(-1)]=val}
function fld(path,label,o={}){
  const v=getPath(path)??"",cls="f"+(o.req?" req":""),id="f_"+path.replace(/\./g,"_");
  const st=o.full?'style="grid-column:1/-1"':"";
  if(o.type==="textarea")return`<label class="${cls}" ${st}><span>${label}</span><textarea id="${id}" data-b="${path}" rows="${o.rows||3}">${esc(v)}</textarea></label>`;
  if(o.opts)return`<label class="${cls}" ${st}><span>${label}</span><select id="${id}" data-b="${path}">${o.opts.map(x=>{const[val,t]=Array.isArray(x)?x:[x,x];return`<option value="${esc(val)}" ${String(v)===String(val)?"selected":""}>${esc(t)}</option>`}).join("")}</select></label>`;
  return`<label class="${cls}" ${st}><span>${label}</span><input id="${id}" data-b="${path}" type="${o.type||"text"}" ${o.im?`inputmode="${o.im}"`:""} value="${esc(v)}" ${o.ph?`placeholder="${esc(o.ph)}"`:""} autocomplete="off"></label>`;
}
function dfld(path,label){return`<label class="f req"><span>${label}</span><input id="f_${path.replace(/\./g,"_")}" data-d="${path}" type="text" inputmode="numeric" autocomplete="off" maxlength="10" placeholder="dd/mm/aaaa" value="${esc(fdate(getPath(path)))}"></label>`}
function tfld(path,label){return`<label class="f req"><span>${label}</span><div class="row" style="flex-wrap:nowrap"><input id="f_${path.replace(/\./g,"_")}" data-b="${path}" type="time" value="${esc(getPath(path))}"><button class="btn sm" data-now="${path}">Ahora</button></div></label>`}
function parseDMY(t){const m=/^(\d{2})\/(\d{2})\/(\d{4})$/.exec(t);if(!m)return null;const d=+m[1],mo=+m[2],y=+m[3];const dt=new Date(y,mo-1,d);if(y<1990||y>2100||dt.getMonth()!==mo-1||dt.getDate()!==d)return null;return`${m[3]}-${m[2]}-${m[1]}`}
function segF(path,label,opts,val){return`<div style="display:flex;flex-direction:column;gap:4px;min-width:0"><span class="muted" style="font-weight:600">${label}</span><div class="seg">${opts.map(o=>`<button data-seg="${path}" data-v="${o}" class="${String(val).trim()===o?"on":""}">${o}</button>`).join("")}</div></div>`}

/* ============ VISTAS ============ */
function calVigencia(){
  const f=S.meta.fechaCal,fm=S.meta.fechaMed;
  if(!f)return`<div class="banner bad">Falta la fecha de calibración del luxómetro (obligatoria). El certificado vigente se adjunta al protocolo.</div>`;
  const d=(new Date(fm||Date.now())-new Date(f))/864e5;
  if(d<0)return`<div class="banner warn">La fecha de calibración es posterior a la de medición.</div>`;
  if(d>365)return`<div class="banner warn">El certificado tiene ${Math.floor(d/30)} meses a la fecha de medición. Verificá que la calibración esté vigente (habitualmente anual).</div>`;
  return`<div class="banner ok">Calibración con ${Math.floor(d/30)} meses a la fecha de medición.</div>`;
}
function vEst(){return`
 <div class="card"><h2>Medición de hoy</h2><div class="grid">
  ${dfld("meta.fechaMed","Fecha de la medición")}${tfld("meta.horaIni","Hora de inicio")}${tfld("meta.horaFin","Hora de finalización")}
  ${fld("meta.cielo","Cielo",{opts:["","Despejado","Parcialmente nublado","Nublado","Lluvioso","Medición nocturna"]})}${fld("meta.temp","Temperatura (°C)",{im:"decimal"})}
 </div>
 <div style="margin-top:10px">${fld("meta.condiciones","Condiciones atmosféricas (texto del protocolo)",{type:"textarea",rows:2})}</div>
 <button class="btn sm" id="genCond" style="margin-top:8px">Armar texto con cielo y temperatura</button></div>
 <div class="card"><h2>Instrumento</h2><div class="grid">
  ${fld("meta.instrumento","Instrumento",{req:1})}${fld("meta.marca","Marca",{req:1})}${fld("meta.modelo","Modelo",{req:1})}
  ${fld("meta.serie","N° de serie",{req:1,ph:"Ver etiqueta del equipo"})}${dfld("meta.fechaCal","Fecha de calibración")}
 </div><div id="calBox">${calVigencia()}</div></div>
 <details class="more"><summary>Establecimiento · ${esc(S.meta.razon)}</summary><div class="in"><div class="grid">
  ${fld("meta.razon","Razón social",{req:1})}${fld("meta.cuit","C.U.I.T.",{req:1,im:"numeric"})}${fld("meta.direccion","Dirección",{req:1})}
  ${fld("meta.localidad","Localidad",{req:1})}${fld("meta.provincia","Provincia",{req:1})}${fld("meta.cp","C.P.",{req:1,im:"numeric"})}
  ${fld("meta.horarios","Horarios / turnos habituales de trabajo",{req:1,full:1})}</div></div></details>
 <details class="more"><summary>Metodología y documentación</summary><div class="in">
  ${fld("meta.metodologia","Metodología utilizada",{type:"textarea",rows:5})}
  <label class="chk"><input type="checkbox" id="f_docCert" data-b="meta.docCert" ${S.meta.docCert?"checked":""}> Certificado de calibración</label>
  <label class="chk"><input type="checkbox" id="f_docCroquis" data-b="meta.docCroquis" ${S.meta.docCroquis?"checked":""}> Plano o croquis con ubicación de los puntos</label>
  ${fld("meta.observaciones","Observaciones",{type:"textarea",rows:2})}</div></details>
 <details class="more"><summary>Profesional y criterio</summary><div class="in"><div class="grid">
  ${fld("meta.profesional","Profesional",{req:1})}${fld("meta.matricula","Título y matrícula",{req:1})}</div>
  <label class="chk"><input type="checkbox" id="f_evalMax" data-b="meta.evaluarMax" ${S.meta.evaluarMax?"checked":""}> Observar valores que superan el máximo del rango de la Tabla 1 (deslumbramiento)</label>
  <p class="muted">El Anexo IV fija mínimos. En la medición anterior se informaron como no conformes los puntos sobre 750 lux; esta opción mantiene ese criterio.</p></div></details>`}
function vPuntos(){
  const st=S.puntos.map(status),cnt={ok:0,bad:0,warn:0,na:0};st.forEach(s=>cnt[s.cls]++);
  const done=S.puntos.length&&cnt.na===0,nc=checks().length;
  const next=S.puntos.find(p=>!calc(p).n);
  return`
 <div class="sumbar"><div><b>${S.puntos.length-cnt.na}/${S.puntos.length}</b><small>medidos</small></div><div><b style="color:var(--ok)">${cnt.ok}</b><small>cumplen</small></div><div><b style="color:var(--bad)">${cnt.bad}</b><small>no cumplen</small></div><div><b style="color:var(--warn)">${cnt.warn}</b><small>observados</small></div></div>
 ${next?`<button class="btn p wide" data-edit="${next.id}" style="margin-bottom:12px;min-height:54px">Medir punto ${next.n} · ${esc(next.sector||"sin sector")} ›</button>`:""}
 ${done?`<div class="card" style="border-color:var(--ok)"><div class="row"><span style="flex:1;font-weight:600">${nc?`Todo medido. Faltan ${nc} datos del protocolo.`:"Protocolo completo ✓"}</span><button class="btn p" id="xlsOrig2">⬇ Excel</button></div></div>`:""}
 <div class="plist">${S.puntos.map((p,i)=>{const s=st[i];return`
  <button class="pitem" data-edit="${p.id}"><span class="num">${p.n}</span>
   <span style="min-width:0"><span class="t" style="display:block">${esc(p.sector||"(sin sector)")}</span>
   <span class="s" style="display:block">${p.hora?p.hora+" h · ":""}req. ${reqStr(reqOf(p))} lx${p.prev?` · ant. ${fmt(p.prev.emed,0)}`:""}</span></span>
   <span class="v"><b>${s.c.n?fmt(s.c.emed,0):"–"}</b><span class="pill ${s.cls}">${s.txt}</span></span></button>`}).join("")}</div>
 <button class="btn wide" id="addP" style="margin-top:12px">＋ Agregar punto</button>
 <p class="muted" style="margin-top:12px">“ant.” es la E media de la medición del ${fdate(S.meta.prevFecha)||"año anterior"}.</p>`}
function cellsHTML(p,c){
  const sel=kp.cell;
  if(p.metodo==="grilla"){
    return`<div class="mgrid" style="grid-template-columns:repeat(${p.cols},minmax(0,1fr))">${Array.from({length:p.filas*p.cols},(_,i)=>{const v=num(p.lect[i]);const has=isFinite(v);
      return`<button class="cell ${has?"":"empty"} ${i===sel?"sel":""} ${has&&c.n>1&&v===c.emin?"min":""}" data-cell="${i}" aria-label="Fila ${Math.floor(i/p.cols)+1} columna ${i%p.cols+1}">${has?fmt(v,1):`${Math.floor(i/p.cols)+1}.${i%p.cols+1}`}</button>`}).join("")}</div>`;
  }
  const n=p.lect.filter(x=>String(x??"").trim()).length;
  return`<div class="mgrid" style="grid-template-columns:repeat(4,minmax(0,1fr))">${Array.from({length:n},(_,i)=>{const v=num(p.lect[i]);return`<button class="cell ${i===sel?"sel":""} ${c.n>1&&v===c.emin?"min":""}" data-cell="${i}">${fmt(v,1)}</button>`}).join("")}<button class="cell add ${sel>=n?"sel":""}" data-cell="${n}" aria-label="Nueva lectura">＋</button></div>`;
}
function cellLabel(p){return p.metodo==="grilla"?`Fila ${Math.floor(kp.cell/p.cols)+1} · Col ${kp.cell%p.cols+1}`:`Lectura ${kp.cell+1}`}
function vEditor(p){
  if(!p){editing=null;return vPuntos()}
  const c=calc(p),i=S.puntos.indexOf(p);
  return`
 <div class="ptop"><button class="navb" id="back" aria-label="Volver a la lista">☰</button>
  <div class="ttl"><b>${p.n} · ${esc(p.sector||"Punto sin sector")}</b><span>${esc(p.puesto||"")}${p.hora?" · "+p.hora+" h":""} · req. ${reqStr(c.req)} lx</span></div>
  <button class="navb" id="prevP" ${i===0?"disabled":""} aria-label="Punto anterior">‹</button><button class="navb" id="nextP" ${i===S.puntos.length-1?"disabled":""} aria-label="Punto siguiente">›</button></div>
 <div class="card" id="liveBox"></div>
 <div class="card">
  <div class="seg" style="margin-bottom:10px"><button data-met="grilla" class="${p.metodo==="grilla"?"on":""}">Grilla del local</button><button data-met="puntual" class="${p.metodo==="puntual"?"on":""}">Puntual (puesto)</button></div>
  ${p.metodo==="grilla"?`<div class="g3">${fld("p.L","Largo m",{im:"decimal"})}${fld("p.A","Ancho m",{im:"decimal"})}${fld("p.Hm","Hm m",{im:"decimal",ph:"lum.→plano"})}</div>
  <div id="kbox" class="banner info" style="margin-bottom:0"></div>`:""}
  <div id="cells">${cellsHTML(p,c)}</div>
  <div class="disp"><small id="cellLbl">${cellLabel(p)}</small><span><b id="dispv" class="${kp.buf?"":"ph"}">${kp.buf||"0"}</b> <small>lx</small></span></div>
  <div class="keys" id="keys">${["7","8","9","4","5","6","1","2","3",",","0","⌫"].map(k=>`<button class="key ${k===","||k==="⌫"?"fn":""}" data-k="${k}">${k}</button>`).join("")}</div>
  <button class="btn p wide savek" id="kSave">Guardar lectura ›</button>
  <div class="row" style="margin-top:8px"><button class="btn sm" id="kClr">Borrar celda</button>${p.metodo==="grilla"?`<span class="muted" style="flex:1;text-align:right">Grilla ${p.filas}×${p.cols}</span><button class="btn sm" id="gridCfg">Cambiar</button>`:""}</div>
  <div id="gridCfgBox" hidden style="margin-top:10px"><div class="row"><label class="f" style="flex:1"><span>Filas</span><input id="g_f" type="number" inputmode="numeric" min="1" max="12" value="${p.filas}"></label><label class="f" style="flex:1"><span>Columnas</span><input id="g_c" type="number" inputmode="numeric" min="1" max="12" value="${p.cols}"></label></div>
   <div class="row" style="margin-top:8px"><button class="btn sm" id="applyK">Usar (x+2)² según K</button><button class="btn sm p" id="applyG">Aplicar</button></div></div>
 </div>
 <button class="btn p wide" id="saveNext" style="margin-bottom:12px;min-height:52px">${i===S.puntos.length-1?"Terminar ✓":"Siguiente punto ›"}</button>
 <details class="more" ${p.sector?"":"open"}><summary>Datos del punto</summary><div class="in"><div class="grid">
  ${fld("p.n","N° de punto",{im:"numeric",req:1})}${tfld("p.hora","Hora")}${fld("p.sector","Sector",{req:1})}${fld("p.puesto","Sección / puesto / puesto tipo",{req:1})}</div>
  <div style="display:grid;gap:10px;margin-top:10px">
  ${segF("p.tipoIlum","Tipo de iluminación",["Natural","Artificial","Mixta"],p.tipoIlum)}
  ${segF("p.fuente","Tipo de fuente lumínica",["Incandescente","Descarga","Mixta"],p.fuente)}
  ${segF("p.ilum","Iluminación",["General","Localizada","Mixta"],p.ilum)}</div>
  <p class="muted">Para LED, indicá la tecnología en observaciones (la planilla oficial prevé incandescente / descarga / mixta).</p>
  ${fld("p.req","Valor requerido · Tabla 1 Anexo IV",{opts:[...T1.map(t=>[t.id,`${t.max?t.min+"–"+t.max:t.min} lx · ${t.txt}`]),["custom","Otro valor (definido por el profesional)"]]})}
  <div class="grid" ${p.req==="custom"?"":"hidden"} style="margin-top:10px">${fld("p.reqMin","Mínimo lx",{im:"decimal"})}${fld("p.reqMax","Máximo lx (opcional)",{im:"decimal"})}</div>
  <div ${p.ilum!=="General"?"":"hidden"} style="margin-top:10px">${fld("p.egen","Iluminación general del local (lx), relación localizada/general",{im:"decimal"})}</div>
  <div style="margin-top:10px">${fld("p.obs","Observaciones",{type:"textarea",rows:2})}</div>
  <div class="row" style="margin-top:10px"><button class="btn sm" id="dupP">Duplicar punto</button><button class="btn sm d" id="delP">Eliminar punto</button></div>
 </div></details>`}
function updLive(){
  const p=cur();if(!p)return;const c=calc(p),r=c.req,box=$("#liveBox");
  const kb=$("#kbox");if(kb)kb.innerHTML=c.K!=null?`K = <b>${fmt(c.K,2)}</b> → mínimo <b>${c.nmin} puntos</b> (${c.x+2}×${c.x+2})${(p.filas*p.cols<c.nmin)?` · <b>tu grilla tiene ${p.filas*p.cols}</b>`:""}`:"Cargá largo, ancho y altura de montaje para calcular K y los puntos mínimos.";
  if(!box)return;
  if(!c.n){box.innerHTML=`<div class="live"><div><small>E media</small><b>–</b></div><div><small>E mín</small><b>–</b></div><div><small>${p.prev?"Anterior":"Requerido"}</small><b>${p.prev?fmt(p.prev.emed,0):reqStr(r)}</b></div></div><p class="muted" style="margin:8px 0 0">Tocá una celda y cargá la lectura con el teclado.</p>`;return}
  const pills=[];
  pills.push(c.minOK===false?`<span class="pill bad">&lt; ${r.min} lx</span>`:`<span class="pill ok">≥ ${r.min} lx</span>`);
  if(S.meta.evaluarMax&&r.max)pills.push(c.maxOK?`<span class="pill ok">≤ ${r.max}</span>`:`<span class="pill warn">&gt; ${r.max} lx</span>`);
  if(c.n>=2)pills.push(c.unifOK?`<span class="pill ok">Uniforme</span>`:`<span class="pill bad">Emín &lt; Emed/2</span>`);
  if(p.metodo==="grilla"&&c.nmin)pills.push(`<span class="pill ${c.puntosOK?"ok":"warn"}">${c.n}/${c.nmin} lect.</span>`);
  if(c.t4!=null)pills.push(c.t4OK===null?`<span class="pill na">Gral. ≥ ${fmt(c.t4,0)}</span>`:`<span class="pill ${c.t4OK?"ok":"bad"}">Gral. ${c.t4OK?"≥":"&lt;"} ${fmt(c.t4,0)}</span>`);
  box.innerHTML=`<div class="live"><div><small>E media</small><b>${fmt(c.emed,1)}</b></div><div><small>E mín · Emed/2</small><b>${fmt(c.emin,0)}·${fmt(c.emed/2,0)}</b></div><div><small>${p.prev?"Anterior":"Requerido"}</small><b>${p.prev?fmt(p.prev.emed,0):reqStr(r)}</b></div></div><div class="verdict">${pills.join("")}</div>`;
}
function refreshCells(){const p=cur();if(!p)return;$("#cells").innerHTML=cellsHTML(p,calc(p));const d=$("#dispv");d.textContent=kp.buf||"0";d.classList.toggle("ph",!kp.buf);$("#cellLbl").textContent=cellLabel(p);updLive()}
function vCon(){const st=S.puntos.map(status);return`
 <div class="card"><h2>Resumen</h2><div class="tablewrap"><table class="sum"><tr><th>N°</th><th>Sector</th><th>E med</th><th>Uniform.</th><th>Req.</th><th>Estado</th></tr>
 ${S.puntos.map((p,i)=>{const s=st[i];return`<tr><td>${p.n}</td><td>${esc(p.sector)}</td><td>${s.c.n?fmt(s.c.emed):"–"}</td><td>${s.c.n>1?fmt(s.c.emin)+" ≥ "+fmt(s.c.emed/2):"–"}</td><td>${reqStr(reqOf(p))}</td><td><span class="pill ${s.cls}">${s.txt}</span></td></tr>`}).join("")}</table></div></div>
 <div class="card"><h2>Análisis de los datos y mejoras</h2>
  <button class="btn p wide" id="genCon" style="margin-bottom:10px">Generar conclusiones automáticamente</button>
  ${fld("conclusiones","Conclusiones",{type:"textarea",rows:8})}<div style="height:10px"></div>
  ${fld("recomendaciones","Recomendaciones para adecuar el nivel de iluminación a la legislación vigente",{type:"textarea",rows:9})}</div>`}
function checks(){const m=S.meta,o=[];
  [["razon","Razón social"],["cuit","CUIT"],["direccion","Dirección"],["localidad","Localidad"],["provincia","Provincia"],["cp","C.P."],["horarios","Horarios/turnos"],["marca","Marca del instrumento"],["modelo","Modelo del instrumento"],["serie","N° de serie del instrumento"],["fechaCal","Fecha de calibración"],["fechaMed","Fecha de medición"],["horaIni","Hora de inicio"],["horaFin","Hora de finalización"],["condiciones","Condiciones atmosféricas"],["metodologia","Metodología"],["profesional","Profesional"],["matricula","Matrícula"]].forEach(([k,t])=>{if(!String(m[k]??"").trim())o.push(t)});
  S.puntos.forEach(p=>{const c=calc(p),mi=[];if(!p.hora)mi.push("hora");if(!p.sector)mi.push("sector");if(!c.n)mi.push("lecturas");if(p.metodo==="grilla"&&c.nmin&&c.n<c.nmin)mi.push(`lecturas de grilla (${c.n}/${c.nmin})`);if(mi.length)o.push(`Punto ${p.n}: ${mi.join(", ")}`)});
  if(!S.conclusiones.trim())o.push("Conclusiones");if(!S.recomendaciones.trim())o.push("Recomendaciones");return o}
function vExp(){const ch=checks();return`
 <div class="card" style="${ch.length?"":"border-color:var(--ok)"}"><h2>Descargar protocolo</h2>
  <button class="btn p wide" id="xlsOrig" style="min-height:58px;font-size:17px">⬇ Descargar Excel${ch.length?"":" ✓"}</button>
  <p class="muted">Mismo formato que la medición anterior: Hoja1 con datos del establecimiento y de la medición, Hoja 2 con los puntos, conclusiones y recomendaciones.</p>
  ${ch.length?`<div class="banner warn">Faltan ${ch.length} datos:</div><ul class="chk-list">${ch.map(x=>`<li>${esc(x)}</li>`).join("")}</ul>`:`<div class="banner ok">Protocolo completo según la Res. SRT 84/12.</div>`}
  <hr class="s"><div class="row"><button class="btn sm" id="xls">Excel Res. 84/12 (3 hojas + lecturas)</button><button class="btn sm" id="prt">Imprimir / PDF</button></div></div>
 <div class="card"><h2>Respaldo y sincronización</h2>
  <p class="muted" style="margin-top:0">${CLOUD?"Cada cambio se guarda en el teléfono y se sube a tu cuenta cuando hay conexión. Desde la PC, ingresá con el mismo usuario y vas a ver esta medición.":"Modo local: los datos quedan solo en este dispositivo."}</p>
  <button class="btn sm" id="bkp">Descargar respaldo (.json)</button></div>
 <div class="card"><h2>Esta medición</h2>
  <div class="row"><button class="btn sm" id="repMed">Repetir medición (nueva, con estos valores como anteriores)</button><button class="btn sm d" id="delMed">Eliminar medición</button></div></div></div>`}

/* ============ EVENTOS ============ */
function rerender(){const y=scrollY;render();scrollTo(0,y)}
function bind(){
  $$("[data-b]").forEach(el=>{
    const ev=(el.tagName==="SELECT"||el.type==="checkbox"||el.type==="time")?"change":"input";
    el.addEventListener(ev,()=>{let v=el.type==="checkbox"?el.checked:el.value;const path=el.dataset.b;if(path==="p.n")v=parseInt(v)||"";setPath(path,v);save();
      if(S.tipo&&MODS[S.tipo])MODS[S.tipo].change(path,el);
      else if(["p.req","p.ilum","meta.evaluarMax"].includes(path))rerender();else if(path.startsWith("p."))updLive();
      if(path==="meta.razon")$("#hdrTitle").textContent=v});
  });
  $$("[data-d]").forEach(el=>el.addEventListener("input",e=>{
    let d=el.value.replace(/\D/g,"").slice(0,8);
    if(e.inputType&&e.inputType.startsWith("delete"))d=el.value.replace(/[^\d/]/g,"");
    else{if(d.length>4)d=d.slice(0,2)+"/"+d.slice(2,4)+"/"+d.slice(4);else if(d.length>2)d=d.slice(0,2)+"/"+d.slice(2)}
    el.value=d;const iso=parseDMY(el.value);el.style.borderColor=(el.value.length===10&&!iso)?"var(--bad)":"";
    if(iso||!el.value){setPath(el.dataset.d,iso||"");save();const cb=$("#calBox");if(cb)cb.innerHTML=calVigencia()}}));
  $$("[data-now]").forEach(b=>b.onclick=e=>{e.preventDefault();setPath(b.dataset.now,nowHM());save();rerender()});
  $$("[data-edit]").forEach(el=>el.onclick=()=>openPoint(el.dataset.edit));
  $$("[data-seg]").forEach(b=>b.onclick=()=>{setPath(b.dataset.seg,b.dataset.v);save();rerender()});
  $$("[data-met]").forEach(b=>b.onclick=()=>{const p=cur();p.metodo=b.dataset.met;kp={cell:firstEmpty(p),buf:""};save();rerender()});
  const on=(id,fn)=>{const el=$("#"+id);if(el)el.onclick=fn};
  // teclado de lecturas
  const cells=$("#cells");
  if(cells){
    cells.addEventListener("click",e=>{const b=e.target.closest("[data-cell]");if(!b)return;const p=cur();kp.cell=+b.dataset.cell;const v=p.lect[kp.cell];kp.buf=v!=null&&String(v).trim()?String(v).replace(".",","):"";refreshCells()});
    $("#keys").addEventListener("click",e=>{const b=e.target.closest("[data-k]");if(b)press(b.dataset.k)});
    on("kSave",commit);on("kClr",()=>{const p=cur();if(p.metodo==="puntual")p.lect.splice(kp.cell,1);else p.lect[kp.cell]="";kp.buf="";save();refreshCells()});
    on("gridCfg",()=>{const g=$("#gridCfgBox");g.hidden=!g.hidden});
    on("applyK",()=>{const c=calc(cur());if(!c.x){toast("Cargá largo, ancho y Hm primero");return}$("#g_f").value=c.x+2;$("#g_c").value=c.x+2});
    on("applyG",()=>resizeGrid(+$("#g_f").value,+$("#g_c").value));
    updLive();
  }
  on("back",()=>{editing=null;render();scrollTo(0,0)});on("prevP",()=>stepP(-1));on("nextP",()=>stepP(1));
  on("saveNext",()=>{const i=S.puntos.findIndex(p=>p.id===editing);if(i===S.puntos.length-1){editing=null;render();scrollTo(0,0);toast("Último punto cargado")}else stepP(1)});
  on("addP",addPoint);
  on("dupP",()=>{const p=cur();const nn=Math.max(0,...S.puntos.map(x=>+x.n||0))+1;const n=Object.assign(JSON.parse(JSON.stringify(p)),{id:uid(),n:nn,lect:[],hora:"",prev:null});S.puntos.splice(S.puntos.indexOf(p)+1,0,n);save();openPoint(n.id)});
  on("delP",async()=>{if(!await ask("¿Eliminar este punto y sus lecturas?",{ok:"Eliminar",danger:true}))return;S.puntos=S.puntos.filter(p=>p.id!==editing);editing=null;save();render()});
  on("genCond",()=>{const m=S.meta;let t=m.cielo==="Medición nocturna"?"Medición realizada en horario nocturno":(m.cielo?`Mediciones realizadas durante el día con cielo ${m.cielo.toLowerCase()}`:"Mediciones realizadas durante el día");if(String(m.temp).trim())t+=` y una temperatura de ${m.temp} °C`;m.condiciones=t+".";save();rerender()});
  on("genCon",async()=>{if((S.conclusiones.trim()||S.recomendaciones.trim())&&!await ask("¿Reemplazar el texto actual por uno nuevo generado con los resultados?",{ok:"Reemplazar"}))return;autoText();save();render()});
  on("xls",()=>doExport(exportXLSX));on("xlsOrig",dlOrig);on("xlsOrig2",dlOrig);on("prt",printProto);on("bkp",backup);
  on("repMed",()=>{const id=newId();putMed(id,repeatFrom(S));openMed(id);go("est");toast("Nueva medición creada con los valores anteriores para comparar")});
  on("delMed",async()=>{if(!await ask("¿Eliminar esta medición? No se puede deshacer.",{ok:"Eliminar",danger:true}))return;const x=lib[curId];x.deleted=true;x.dirty=true;x.updatedAt=Date.now();x.data.updatedAt=x.updatedAt;writeCache();schedulePush(300);view="list";curId=null;S=null;render()});
}
function firstEmpty(p){const n=p.metodo==="grilla"?p.filas*p.cols:p.lect.filter(x=>String(x??"").trim()).length+1;for(let i=0;i<n;i++)if(!String(p.lect[i]??"").trim())return i;return 0}
function openPoint(id){editing=id;tab="pts";const p=cur();kp={cell:firstEmpty(p),buf:""};render();scrollTo(0,0)}
function stepP(d){const i=S.puntos.findIndex(p=>p.id===editing),j=i+d;if(j<0||j>=S.puntos.length)return;openPoint(S.puntos[j].id)}
function addPoint(){const last=S.puntos.at(-1);const n=newPoint(Math.max(0,...S.puntos.map(x=>+x.n||0))+1,last?{tipoIlum:last.tipoIlum,fuente:last.fuente,ilum:last.ilum,req:last.req,reqMin:last.reqMin,reqMax:last.reqMax,metodo:last.metodo,Hm:last.Hm}:{});S.puntos.push(n);save();openPoint(n.id)}
function press(k){
  if(k==="⌫")kp.buf=kp.buf.slice(0,-1);
  else if(k===","){if(!kp.buf.includes(","))kp.buf=(kp.buf||"0")+","}
  else if(kp.buf.replace(",","").length<6)kp.buf=(kp.buf==="0"?"":kp.buf)+k;
  const d=$("#dispv");d.textContent=kp.buf||"0";d.classList.toggle("ph",!kp.buf);
}
function commit(){
  const p=cur();if(!p)return;
  if(!kp.buf){toast("Escribí el valor antes de guardar");return}
  const v=num(kp.buf);if(!isFinite(v)||v<0){toast("Valor no válido");return}
  p.lect[kp.cell]=String(v);if(!p.hora)p.hora=nowHM();
  if(!S.meta.horaIni)S.meta.horaIni=nowHM();
  buzz();save();
  const total=p.metodo==="grilla"?p.filas*p.cols:null;
  let nx=-1;
  if(total){for(let s=1;s<=total;s++){const j=(kp.cell+s)%total;if(!String(p.lect[j]??"").trim()){nx=j;break}}}
  else nx=p.lect.filter(x=>String(x??"").trim()).length;
  kp.buf="";
  if(nx===-1){kp.cell=Math.min(kp.cell,total-1);refreshCells();toast("Grilla completa ✓");}
  else{kp.cell=nx;refreshCells()}
  const hi=$("#f_p_hora");if(hi)hi.value=p.hora;
}
function resizeGrid(f,c){
  const p=cur();f=Math.min(12,Math.max(1,f||1));c=Math.min(12,Math.max(1,c||1));
  const rows=[];for(let r=0;r<p.filas;r++)rows.push(p.lect.slice(r*p.cols,r*p.cols+p.cols));
  const nl=[];for(let r=0;r<f;r++)for(let k=0;k<c;k++)nl.push((rows[r]||[])[k]??"");
  const lost=p.lect.filter(x=>String(x??"").trim()).length-nl.filter(x=>String(x??"").trim()).length;
  const apply=()=>{p.filas=f;p.cols=c;p.lect=nl;kp={cell:firstEmpty(p),buf:""};save();rerender()};
  if(lost>0)ask(`Se perderán ${lost} lecturas que quedan fuera de la grilla nueva.`,{ok:"Aplicar",danger:true}).then(ok=>ok&&apply());else apply();
}
document.addEventListener("keydown",e=>{
  if(!(tab==="pts"&&editing&&$("#keys")))return;const t=e.target;if(t&&(t.tagName==="INPUT"||t.tagName==="TEXTAREA"||t.tagName==="SELECT"))return;
  if(/^\d$/.test(e.key))press(e.key);else if(e.key===","||e.key===".")press(",");else if(e.key==="Backspace")press("⌫");else if(e.key==="Enter"){e.preventDefault();commit()}else return;e.preventDefault();
});
$("#themeBtn").onclick=()=>{const r=document.documentElement;const dark=r.dataset.theme?r.dataset.theme==="dark":matchMedia("(prefers-color-scheme:dark)").matches;r.dataset.theme=dark?"light":"dark";try{localStorage.setItem(KEY+"_th",r.dataset.theme)}catch(e){}};
try{const th=localStorage.getItem(KEY+"_th");if(th)document.documentElement.dataset.theme=th}catch(e){}

/* ============ TEXTO AUTOMÁTICO ============ */
function autoText(){
  const bajo=[],alto=[],unif=[],gen=[],ok=[],sin=[];
  S.puntos.forEach(p=>{const c=calc(p);if(!c.n){sin.push(p.n);return}if(c.minOK===false)bajo.push(p);if(!c.maxOK)alto.push(p);if(c.n>=2&&!c.unifOK)unif.push(p);if(c.t4OK===false)gen.push(p);if(c.cumple)ok.push(p.n)});
  const L=a=>a.map(p=>`N°${p.n} (${p.sector})`).join(", ");const con=[];
  con.push(`Se evaluaron ${S.puntos.length-sin.length} puntos de muestreo en el establecimiento ${S.meta.razon}, comparando la iluminancia media con los valores del Anexo IV del Decreto 351/79 y verificando la uniformidad (E mínima ≥ E media / 2).`);
  if(sin.length)con.push(`Quedaron sin medir los puntos N° ${sin.join(", ")}.`);
  if(ok.length)con.push(`Cumplen con la legislación vigente los puntos N° ${ok.join(", ")}.`);
  if(bajo.length)con.push(`No cumplen con el valor mínimo requerido: ${L(bajo)}.`);
  if(alto.length)con.push(`Superan el valor máximo del rango para la tarea visual: ${L(alto)}.`);
  if(unif.length)con.push(`No cumplen con la uniformidad de iluminancia: ${L(unif)}.`);
  if(gen.length)con.push(`La iluminación general es insuficiente respecto de la localizada (Anexo IV): ${L(gen)}.`);
  if(!bajo.length&&!alto.length&&!unif.length&&!gen.length&&ok.length)con.push("Los niveles de iluminación determinados cumplen con la legislación vigente.");
  const pc=S.puntos.filter(p=>p.prev&&calc(p).n).map(p=>{const c=calc(p),d=(c.emed-p.prev.emed)/p.prev.emed*100;return Math.abs(d)>=20?`N°${p.n} ${d>0?"+":""}${fmt(d,0)}%`:null}).filter(Boolean);
  if(pc.length)con.push(`Variaciones significativas respecto de la medición del ${fdate(S.meta.prevFecha)}: ${pc.join(", ")}.`);
  const rec=[];
  if(bajo.length)rec.push(`Puntos ${bajo.map(p=>"N°"+p.n).join(", ")}: incrementar el nivel de iluminación hasta alcanzar el mínimo requerido (reemplazo de lámparas agotadas por LED de mayor flujo, incorporación de luminarias o iluminación localizada en el plano de trabajo) y verificar mediante una nueva medición.`);
  if(alto.length)rec.push(`Puntos ${alto.map(p=>"N°"+p.n).join(", ")}: controlar el aporte de luz natural mediante el uso de las persianas/cortinas existentes y reubicar los puestos con pantalla de forma perpendicular a las ventanas, para evitar deslumbramientos y reflejos.`);
  if(unif.length)rec.push(`Puntos ${unif.map(p=>"N°"+p.n).join(", ")}: mejorar la distribución de las luminarias, reemplazar las fuera de servicio y evitar obstrucciones para lograr que la iluminancia mínima sea al menos la mitad de la media.`);
  if(gen.length)rec.push(`Puntos ${gen.map(p=>"N°"+p.n).join(", ")}: aumentar la iluminación general del local según la relación iluminación localizada / general del Anexo IV.`);
  rec.push("Mantener un programa de limpieza periódica de luminarias y superficies, y de reemplazo preventivo de lámparas.");
  rec.push("Realizar la medición de iluminación con frecuencia anual o ante modificaciones en los puestos, la distribución o el sistema de iluminación.");
  S.conclusiones=con.join("\n\n");S.recomendaciones=rec.map(r=>"• "+r).join("\n");
}

/* ============ EXPORTAR ============ */
/* Agrega a cada hoja: A4, ajustar al ancho de página y orientación (pages: "p"/"l" por hoja) */
function withPageSetup(wb,pages){
  const out=XLSX.write(wb,{bookType:"xlsx",type:"array"});
  try{
    const cfb=XLSX.CFB.read(new Uint8Array(out),{type:"array"}),enc=new TextEncoder(),dec=new TextDecoder();
    wb.SheetNames.forEach((n,i)=>{
      const e=XLSX.CFB.find(cfb,"/xl/worksheets/sheet"+(i+1)+".xml");if(!e)return;
      let x=dec.decode(e.content);const o=(pages&&pages[i])==="l"?"landscape":"portrait";
      const ps=`<pageSetup paperSize="9" orientation="${o}" fitToWidth="1" fitToHeight="0"/>`;
      if(!/<sheetPr/.test(x))x=x.replace(/(<worksheet[^>]*>)/,'$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
      if(/<pageMargins[^>]*\/>/.test(x))x=x.replace(/(<pageMargins[^>]*\/>)/,"$1"+ps);
      else if(/<ignoredErrors/.test(x))x=x.replace(/<ignoredErrors/,'<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>'+ps+"<ignoredErrors");
      else x=x.replace("</worksheet>",'<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>'+ps+"</worksheet>");
      const b=enc.encode(x);e.content=b;e.size=b.length;
    });
    return XLSX.CFB.write(cfb,{fileType:"zip",type:"array"});
  }catch(err){console.warn("pageSetup",err);return out}
}
function saveWorkbook(wb,filename,pages){
  try{const data=withPageSetup(wb,pages);
    const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([data],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),3000);
    toast("Excel descargado")}catch(e){console.warn(e);toast("No se pudo generar el archivo")}
}
function doExport(fn){if(typeof XLSX==="undefined"){toast("El generador de Excel no cargó. Revisá la conexión y volvé a abrir la página.");return}fn()}
async function dlOrig(){const ch=checks();if(ch.length&&!await ask(`Faltan ${ch.length} datos:\n• ${ch.slice(0,8).join("\n• ")}${ch.length>8?"\n…":""}\n\n¿Descargar igual?`,{ok:"Descargar igual"}))return;doExport(exportOriginal)}
function rowsHoja2(){return S.puntos.map(p=>{const c=calc(p);return[p.n,p.hora||"",p.sector,p.puesto,String(p.tipoIlum).trim(),p.fuente,p.ilum,c.n>1?`${fmt(c.emin)} ≥ ${fmt(c.emed/2)}`:(c.n?"–":""),c.n?Math.round(c.emed*100)/100:"",reqStr(reqOf(p))]})}
function exportXLSX(){
  const m=S.meta,B={style:"thin",color:{rgb:"000000"}},border={top:B,bottom:B,left:B,right:B};
  const sT={font:{bold:true,sz:13},alignment:{horizontal:"center",vertical:"center",wrapText:true},fill:{fgColor:{rgb:"D9D9D9"}},border};
  const sL={font:{bold:true},fill:{fgColor:{rgb:"F2F2F2"}},alignment:{vertical:"center",wrapText:true},border};
  const sV={alignment:{vertical:"center",wrapText:true},border},sH={font:{bold:true,sz:10},alignment:{horizontal:"center",vertical:"center",wrapText:true},fill:{fgColor:{rgb:"E6E6E6"}},border};
  const sC={alignment:{horizontal:"center",vertical:"center",wrapText:true},border},sBad={...sC,font:{bold:true,color:{rgb:"B00000"}}};
  const sheet=(aoa,styles,merges,cols,rowsH)=>{const ws=XLSX.utils.aoa_to_sheet(aoa);Object.entries(styles).forEach(([r,st])=>{if(!ws[r])ws[r]={t:"s",v:""};ws[r].s=st});ws["!merges"]=merges.map(x=>XLSX.utils.decode_range(x));ws["!cols"]=cols.map(w=>({wch:w}));if(rowsH)ws["!rows"]=rowsH;return ws};
  const fill=(st,rng,s)=>{const r=XLSX.utils.decode_range(rng);for(let R=r.s.r;R<=r.e.r;R++)for(let C=r.s.c;C<=r.e.c;C++)st[XLSX.utils.encode_cell({r:R,c:C})]=s};
  const firma=`Firma, aclaración y registro del profesional interviniente: ${m.profesional} – ${m.matricula}`;
  const h1=[["PROTOCOLO PARA MEDICIÓN DE ILUMINACIÓN EN EL AMBIENTE LABORAL","","",""],["Razón social:",m.razon,"",""],["Dirección:",m.direccion,"",""],["Localidad:",m.localidad,"",""],["Provincia:",m.provincia,"",""],["C.P.:",m.cp,"C.U.I.T.:",m.cuit],["Horarios / turnos habituales de trabajo:",m.horarios,"",""],["DATOS DE LA MEDICIÓN","","",""],
   ["Marca, modelo y N° de serie del instrumento utilizado:",`${m.instrumento} ${m.marca} ${m.modelo} – N° de serie: ${m.serie||"—"}`,"",""],["Fecha de calibración del instrumental utilizado en la medición:",fdate(m.fechaCal),"",""],["Metodología utilizada en la medición:",m.metodologia,"",""],
   ["Fecha de la medición:",fdate(m.fechaMed),"Hora de inicio:",m.horaIni],["","","Hora de finalización:",m.horaFin],["Condiciones atmosféricas:",m.condiciones,"",""],["DOCUMENTACIÓN QUE SE ADJUNTARÁ A LA MEDICIÓN","","",""],
   ["Certificado de calibración:",m.docCert?"SÍ":"NO","",""],["Plano o croquis del establecimiento:",m.docCroquis?"SÍ":"NO","",""],["Observaciones:",m.observaciones,"",""],[firma,"","",""]];
  const st1={};fill(st1,"A1:D1",sT);fill(st1,"A8:D8",sT);fill(st1,"A15:D15",sT);
  for(let r=2;r<=19;r++){if(r===8||r===15)continue;fill(st1,`A${r}:D${r}`,sV);st1["A"+r]=sL}st1.C6=sL;st1.C12=sL;st1.C13=sL;fill(st1,"A19:D19",{...sV,font:{italic:true}});
  const ws1=sheet(h1,st1,["A1:D1","A8:D8","A15:D15","A19:D19","A12:A13","B12:B13",...[2,3,4,5,7,9,10,11,14,16,17,18].map(r=>`B${r}:D${r}`)],[34,34,20,22],h1.map((_,i)=>({hpt:i===10?110:(i===0?30:22)})));
  const H=["Punto de muestreo","Hora","Sector","Sección / Puesto / Puesto tipo","Tipo de iluminación: Natural / Artificial / Mixta","Tipo de fuente lumínica: Incandescente / Descarga / Mixta","Iluminación: General / Localizada / Mixta","Valor de la uniformidad de iluminancia E mínima ≥ (E media)/2","Valor medido (Lux)","Valor requerido legalmente según Anexo IV Dec. 351/79"];
  const rows=rowsHoja2();
  const h2=[["PROTOCOLO PARA MEDICIÓN DE ILUMINACIÓN EN EL AMBIENTE LABORAL",...Array(9).fill("")],[`Razón social: ${m.razon}`,"","","","",`C.U.I.T.: ${m.cuit}`,"","","",""],[`Dirección: ${m.direccion}`,"","",`Localidad: ${m.localidad}`,"","",`C.P.: ${m.cp}`,"",`Provincia: ${m.provincia}`,""],["DATOS DE LA MEDICIÓN",...Array(9).fill("")],H,...rows,[],[firma,...Array(9).fill("")]];
  const st2={};fill(st2,"A1:J1",sT);fill(st2,"A4:J4",sT);fill(st2,"A2:J3",sL);fill(st2,"A5:J5",sH);
  rows.forEach((rw,i)=>{const R=6+i;fill(st2,`A${R}:J${R}`,sC);st2["C"+R]=sV;st2["D"+R]=sV;const s=status(S.puntos[i]);if(s.cls==="bad"||s.cls==="warn")st2["I"+R]=sBad});
  const lr=6+rows.length+1;fill(st2,`A${lr}:J${lr}`,{font:{italic:true}});
  const ws2=sheet(h2,st2,["A1:J1","A4:J4","A2:E2","F2:J2","A3:C3","D3:F3","G3:H3","I3:J3",`A${lr}:J${lr}`],[10,8,20,22,14,16,14,20,12,16],[{hpt:28},{hpt:20},{hpt:20},{hpt:22},{hpt:62}]);
  const h3=[["PROTOCOLO PARA MEDICIÓN DE ILUMINACIÓN EN EL AMBIENTE LABORAL","",""],[`Razón social: ${m.razon}`,"",`C.U.I.T.: ${m.cuit}`],[`Dirección: ${m.direccion} – ${m.localidad} – C.P. ${m.cp} – ${m.provincia}`,"",""],["ANÁLISIS DE LOS DATOS Y MEJORAS A REALIZAR","",""],["Conclusiones","",""],[S.conclusiones,"",""],["Recomendaciones para adecuar el nivel de iluminación a la legislación vigente","",""],[S.recomendaciones,"",""],[firma,"",""]];
  const st3={};fill(st3,"A1:C1",sT);fill(st3,"A4:C4",sT);fill(st3,"A2:C3",sL);fill(st3,"A5:C5",sH);fill(st3,"A7:C7",sH);fill(st3,"A6:C6",{...sV,alignment:{vertical:"top",wrapText:true}});fill(st3,"A8:C8",{...sV,alignment:{vertical:"top",wrapText:true}});fill(st3,"A9:C9",{font:{italic:true}});
  const lines=t=>String(t||"").split("\n").reduce((a,l)=>a+Math.max(1,Math.ceil(l.length/95)),0);
  const ws3=sheet(h3,st3,["A1:C1","A4:C4","A2:B2","A3:C3","A5:C5","A6:C6","A7:C7","A8:C8","A9:C9"],[50,30,30],[{hpt:28},{hpt:20},{hpt:20},{hpt:22},{hpt:20},{hpt:Math.max(60,lines(S.conclusiones)*15)},{hpt:20},{hpt:Math.max(80,lines(S.recomendaciones)*15)}]);
  const an=[["ANEXO – REGISTRO DE LECTURAS DE CAMPO"],[]];
  S.puntos.forEach(p=>{const c=calc(p);an.push([`Punto ${p.n} – ${p.sector} – ${p.puesto}`]);
    an.push([`Método: ${p.metodo==="grilla"?"Grilla":"Puntual"}`,p.metodo==="grilla"?`L=${p.L||"–"} m · A=${p.A||"–"} m · Hm=${p.Hm||"–"} m · K=${c.K!=null?fmt(c.K):"–"} · mín. ${c.nmin??"–"} puntos`:"",`E med=${fmt(c.emed)} · E mín=${fmt(c.emin)} · E máx=${fmt(c.emax)} lx`]);
    if(p.metodo==="grilla"){for(let r=0;r<p.filas;r++)an.push(Array.from({length:p.cols},(_,k)=>{const v=num(p.lect[r*p.cols+k]);return isFinite(v)?v:""}))}else an.push(c.vals);
    if(p.obs)an.push([`Obs.: ${p.obs}`]);an.push([])});
  const ws4=XLSX.utils.aoa_to_sheet(an);ws4["!cols"]=Array(12).fill({wch:12});if(ws4.A1)ws4.A1.s={font:{bold:true,sz:12}};
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws1,"Hoja 1 - Datos");XLSX.utils.book_append_sheet(wb,ws2,"Hoja 2 - Puntos");XLSX.utils.book_append_sheet(wb,ws3,"Hoja 3 - Análisis");XLSX.utils.book_append_sheet(wb,ws4,"Anexo lecturas");
  saveWorkbook(wb,`Protocolo_Iluminacion_8412_${(m.razon||"empresa").replace(/[^A-Za-z0-9]+/g,"_")}_${m.fechaMed||""}.xlsx`,["p","l","p","p"]);
}
function exportOriginal(){
  const m=S.meta,T={style:"thin",color:{rgb:"000000"}},F=(sz=10,b=false)=>({name:"Roboto",sz,bold:b});
  const enc=XLSX.utils.encode_cell,dec=XLSX.utils.decode_range,mk=()=>({cells:{},merges:[]});
  const put=(w,ref,v,st={})=>{const c=w.cells[ref]||(w.cells[ref]={});c.v=v;c.s=Object.assign(c.s||{},st)};
  const box=(w,rng)=>{const r=dec(rng);for(let R=r.s.r;R<=r.e.r;R++)for(let C=r.s.c;C<=r.e.c;C++){const ref=enc({r:R,c:C});const c=w.cells[ref]||(w.cells[ref]={v:""});c.s=c.s||{};const b=Object.assign({},c.s.border||{});if(R===r.s.r)b.top=T;if(R===r.e.r)b.bottom=T;if(C===r.s.c)b.left=T;if(C===r.e.c)b.right=T;c.s.border=b}};
  const grid=(w,rng)=>{const r=dec(rng);for(let R=r.s.r;R<=r.e.r;R++)for(let C=r.s.c;C<=r.e.c;C++)box(w,enc({r:R,c:C}))};
  const mg=(w,rng)=>w.merges.push(dec(rng));
  const build=(w,cols,rows,margins)=>{const ws={};let mr=0,mc=0;for(const[ref,c]of Object.entries(w.cells)){const a=XLSX.utils.decode_cell(ref);mr=Math.max(mr,a.r);mc=Math.max(mc,a.c);const v=c.v;const cell=typeof v==="number"?{t:"n",v}:{t:"s",v:String(v??"")};cell.s=Object.assign({font:F(10)},c.s);if(c.s&&c.s.numFmt)cell.z=c.s.numFmt;ws[ref]=cell}
    ws["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:mr,c:mc}});ws["!merges"]=w.merges;ws["!cols"]=cols.map(x=>({wch:Math.max(1,x-0.73)}));ws["!rows"]=rows;ws["!margins"]=margins;return ws};
  const L={horizontal:"left",vertical:"center"},Cc={horizontal:"center",vertical:"center"};
  const linesH=(t,ch,lh)=>Math.max(1,String(t||"").split("\n").reduce((a,l)=>a+Math.max(1,Math.ceil(l.length/ch)),0))*lh+6;
  const a=mk(),rh1={1:70,2:41,3:13.5,4:16.5,5:7.5,6:17.25,7:10.5,8:16.5,9:9.75,10:15,11:10.5,12:16.5,13:13.5,14:18.75,15:18.75,16:13.5,17:13.5,18:15,19:18,20:18,21:16.5,22:16.5,24:3.75,26:25,27:8.25,28:13.5,29:16.5,31:10.5,32:13.5,34:30,35:13.5,36:13.5,37:13.5,38:7.5,40:9.75,42:6.75};
  put(a,"B2","PROTOCOLO PARA MEDICIÓN DE ILUMINACIÓN EN EL AMBIENTE LABORAL",{font:F(10,true),alignment:Cc});mg(a,"B2:H2");box(a,"B2:H2");
  [["B4","RAZON  SOCIAL:","D4",m.razon,"B4:C4","D4:H4"],["B6","DIRECCIÓN:","D6",m.direccion,null,"D6:H6"],["B8","LOCALIDAD:","D8",m.localidad,null,"D8:H8"],["B10","PROVINCIA:","D10",m.provincia,null,null]].forEach(([l,lt,v,vt,m1,m2])=>{put(a,l,lt,{alignment:L});put(a,v,vt,{alignment:L});if(m1)mg(a,m1);if(m2)mg(a,m2);const r=l.slice(1);box(a,`B${r}:H${r}`)});
  put(a,"B12","C.P.:",{alignment:L});put(a,"D12",/^\d+$/.test(m.cp)?+m.cp:m.cp,{alignment:L});box(a,"B12:E12");put(a,"F12","CUIT:",{alignment:L});put(a,"G12",m.cuit,{alignment:L});mg(a,"G12:H12");box(a,"F12:H12");
  put(a,"B14","Horarios / Turnos habituales de trabajo ",{font:F(10,true),alignment:L});mg(a,"B14:H14");put(a,"B15",m.horarios,{alignment:L});mg(a,"B15:H15");box(a,"B14:H15");
  put(a,"B17","DATOS DE LA MEDICIÓN",{font:F(10,true),alignment:Cc});mg(a,"B17:H17");box(a,"B17:H17");
  put(a,"B18","Marca,modelo y Nº de serie del instrumento utilizado",{font:F(10,true),alignment:L});mg(a,"B18:H18");
  const vc={alignment:{vertical:"center"}};put(a,"B19","Instrumento:",vc);put(a,"C19",m.instrumento,vc);put(a,"D19","Marca:",vc);put(a,"E19",m.marca,vc);put(a,"F19","Modelo:",vc);put(a,"G19",m.modelo,vc);box(a,"B18:H19");
  put(a,"B20","Fecha de calibración",{font:F(10,true),...vc});put(a,"B21",fdate(m.fechaCal),{alignment:L});put(a,"B22",`Nº de serie: ${m.serie||""}`,{alignment:L});box(a,"B20:H22");
  put(a,"B23","Metodología utilizada en la medición",{font:F(10,true),...vc});
  const met=(m.metodologia===METODO)?"Colocando la célula fotoeléctrica del instrumento sobre el plano de trabajo o a 0,80 metros del solado en medios de circulación. Índice del local K = (L·A)/[Hm·(L+A)]; puntos mínimos (x+2)².":String(m.metodologia||"");
  put(a,"B25",/grilla|cuadr[ií]cula/i.test(String(m.metodologia))?"Método de la Grilla o Cuadrícula":"",{alignment:L});mg(a,"B25:H25");
  put(a,"B26",met,{alignment:{horizontal:"left",vertical:"center",wrapText:true}});mg(a,"B26:H26");rh1[26]=linesH(met,62,13);box(a,"B23:H27");
  put(a,"B29","Fecha de la medición",{font:F(10,true),alignment:Cc});mg(a,"B29:C29");put(a,"E29","Hora de inicio",{font:F(10,true),alignment:Cc});put(a,"G29","Hora de finalización",{font:F(10,true),alignment:Cc});mg(a,"G29:H29");
  put(a,"B30",fdate(m.fechaMed),{alignment:Cc});mg(a,"B30:C30");put(a,"E30",m.horaIni?m.horaIni+" hs":"",{alignment:Cc});put(a,"G30",m.horaFin?m.horaFin+" hs":"",{alignment:Cc});mg(a,"G30:H30");box(a,"B29:H31");
  put(a,"B33","Condiciones atmosféricas",{font:F(10,true),...vc});put(a,"B34",m.condiciones,{alignment:{horizontal:"left",vertical:"center",wrapText:true}});mg(a,"B34:H34");rh1[34]=Math.max(30,linesH(m.condiciones,62,13));box(a,"B33:H35");
  put(a,"B37","DOCUMENTACIÓN QUE SE ADJUNTARÁ A LA MEDICIÓN",{font:F(10,true),alignment:Cc});mg(a,"B37:H37");box(a,"B37:H37");
  const docs=[m.docCert&&"Certificado de calibración",m.docCroquis&&"Croquis del establecimiento"].filter(Boolean);
  put(a,"B39",docs[0]||"",{alignment:L});mg(a,"B39:H39");put(a,"B41",docs[1]||"",{alignment:L});mg(a,"B41:H41");box(a,"B38:H42");
  if(String(m.observaciones||"").trim()){put(a,"B44","Observaciones",{font:F(10,true)});put(a,"B45",m.observaciones,{alignment:{horizontal:"left",vertical:"top",wrapText:true}});mg(a,"B45:H45");rh1[45]=linesH(m.observaciones,62,13);box(a,"B44:H45")}
  const rows1=[];for(let i=1;i<=45;i++)rows1[i-1]=rh1[i]?{hpt:rh1[i]}:{};
  const ws1=build(a,[7.6,11.4,11.5,8.7,12.6,7.6,11.5,7,11.5],rows1,{left:0.78,right:0.75,top:0.66,bottom:1,header:0,footer:0});
  const b=mk(),f9=F(9),rh2={1:56,3:12.5,4:12.5,5:12.5,6:20,7:20,8:12.5,9:17.5,10:3.75,11:24.75,12:6.75,16:6.75};
  put(b,"A3","PROTOCOLO PARA MEDICIÓN DE LA ILUMINACIÓN EN EL AMBIENTE LABORAL",{font:F(10,true),alignment:{horizontal:"center"}});mg(b,"A3:J3");box(b,"A2:J4");
  put(b,"A6","RAZON SOCIAL:",{font:f9,alignment:L});put(b,"C6",m.razon,{font:f9,alignment:L});put(b,"F6","CUIT:",{font:f9,alignment:L});put(b,"G6",m.cuit,{font:f9,alignment:L});box(b,"A6:J6");
  put(b,"A7","DIRECCIÓN:",{font:f9,alignment:L});mg(b,"A7:B7");put(b,"C7",m.direccion,{font:f9,alignment:L});put(b,"E7","LOCALIDAD:",{font:f9,alignment:L});put(b,"F7",m.localidad,{font:f9,alignment:L});put(b,"G7",`C.P.: ${m.cp}`,{font:f9,alignment:L});put(b,"H7","PROVINCIA:",{font:f9,alignment:L});put(b,"I7",m.provincia,{font:f9,alignment:L});box(b,"A7:J7");
  put(b,"A9","DATOS DE LA MEDICIÓN",{alignment:Cc});mg(b,"A9:J9");box(b,"A9:J9");
  const hc={font:f9,alignment:{horizontal:"center",vertical:"center",wrapText:true}},lc={font:f9,alignment:{horizontal:"center"}};
  put(b,"A10","Punto de muestreo y Sector",hc);mg(b,"A10:B16");box(b,"A10:B16");put(b,"C10","Sección puesto",hc);mg(b,"C10:C16");box(b,"C10:C16");
  [["D","Tipo de iluminación",["Natural","Artificial","Mixta"]],["E","Tipo de fuente",["Incandescente","Descarga","Mixta"]],["F","Iluminación",["General","Localizada","Mixta"]]].forEach(([c,t,o])=>{put(b,c+"10",t,hc);mg(b,`${c}10:${c}12`);o.forEach((x,i)=>put(b,c+(13+i),x,lc));box(b,`${c}10:${c}16`)});
  put(b,"G10","Valor de la uniformidad",hc);mg(b,"G10:G12");put(b,"G13","Iluminación",lc);put(b,"G14","E min > E.med/2",hc);mg(b,"G14:G16");box(b,"G10:G16");
  put(b,"H10","Valor medido",hc);mg(b,"H10:H12");put(b,"H13","( Lux )",hc);mg(b,"H13:H16");box(b,"H10:H16");
  put(b,"I10","Valor requerido",hc);mg(b,"I10:I12");put(b,"I13","Según ",lc);put(b,"I14","Anexo  IV",lc);put(b,"I15","Dec. 351/79",lc);box(b,"I10:I16");
  put(b,"J10","Hora",hc);mg(b,"J10:J12");box(b,"J10:J16");
  const gray={fgColor:{rgb:"D9D9D9"},patternType:"solid"};let R=17;
  S.puntos.forEach(p=>{const c=calc(p),st=status(p),bad=(st.cls==="bad"||st.cls==="warn");const base={font:f9,alignment:{horizontal:"center",vertical:"center"}};if(bad)base.fill=gray;
    put(b,"A"+R,+p.n||p.n,base);put(b,"B"+R,p.sector,{...base,alignment:{vertical:"center",wrapText:true}});put(b,"C"+R,p.puesto,{...base,alignment:{horizontal:"center",vertical:"center",wrapText:true}});
    put(b,"D"+R,String(p.tipoIlum).trim(),base);put(b,"E"+R,p.fuente,base);put(b,"F"+R,p.ilum,base);put(b,"G"+R,c.n>1?`${fmt(c.emin)} ≥ ${fmt(c.emed/2)}`:"",base);
    if(c.n)put(b,"H"+R,Math.round(c.emed*100)/100,{...base,numFmt:"0.00"});else put(b,"H"+R,"",base);put(b,"I"+R,reqStr(reqOf(p)),base);put(b,"J"+R,p.hora||"",base);
    rh2[R]=Math.max(14,Math.max(Math.ceil(String(p.sector).length/14),Math.ceil(String(p.puesto).length/12))*11+3);R++});
  grid(b,`A17:J${R-1}`);
  const cR=R+1;rh2[R]=12.5;put(b,"A"+cR,"Conclusiones y Recomendaciones para adecuar la iluminación a la legislación vigente",{font:F(10,true),alignment:Cc});mg(b,`A${cR}:J${cR+1}`);box(b,`A${cR}:J${cR+1}`);rh2[cR]=13;rh2[cR+1]=13;
  const tR=cR+3;rh2[cR+2]=13;const wrapT={font:F(10),alignment:{horizontal:"left",vertical:"top",wrapText:true}};
  put(b,"A"+tR,S.conclusiones,wrapT);mg(b,`A${tR}:J${tR}`);rh2[tR]=linesH(S.conclusiones,95,13);put(b,"A"+(tR+1),S.recomendaciones,wrapT);mg(b,`A${tR+1}:J${tR+1}`);rh2[tR+1]=linesH(S.recomendaciones,95,13);box(b,`A${cR+2}:J${tR+1}`);
  const sR=tR+3;put(b,"F"+sR,"_______________________________",{alignment:{horizontal:"right"}});mg(b,`F${sR}:J${sR}`);put(b,"F"+(sR+1),m.profesional,{font:F(9),alignment:{horizontal:"right"}});mg(b,`F${sR+1}:J${sR+1}`);
  put(b,"D"+(sR+2),m.matricula,{font:F(8),alignment:{horizontal:"right"}});mg(b,`D${sR+2}:J${sR+2}`);rh2[sR+2]=12;rh2[tR+2]=8;rh2[sR]=12;
  const rows2=[];for(let i=1;i<=sR+2;i++)rows2[i-1]=rh2[i]?{hpt:rh2[i]}:{};
  const ws2=build(b,[2.7,12,12.5,8,11,9,11,8.5,10,6],rows2,{left:0.62,right:0.38,top:0.34,bottom:0.44,header:0.16,footer:0});
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws1,"Hoja1");XLSX.utils.book_append_sheet(wb,ws2,"Hoja 2");
  saveWorkbook(wb,`Protocolo_Iluminacion_${(m.razon||"empresa").replace(/[^A-Za-z0-9]+/g,"_")}_${m.fechaMed||""}.xlsx`);
}
function printProto(){
  const m=S.meta,sig=`<div class="sig"><div>${esc(m.profesional)}<br>${esc(m.matricula)}<br>Firma, aclaración y registro del profesional interviniente</div></div>`,r=(l,v)=>`<tr><td class="lbl">${l}</td><td colspan="3">${esc(v)}</td></tr>`;
  const h1=`<div class="ph"><table><tr><td class="ttl" colspan="4">PROTOCOLO PARA MEDICIÓN DE ILUMINACIÓN EN EL AMBIENTE LABORAL</td></tr>${r("Razón social:",m.razon)}${r("Dirección:",m.direccion)}${r("Localidad:",m.localidad)}${r("Provincia:",m.provincia)}<tr><td class="lbl">C.P.:</td><td>${esc(m.cp)}</td><td class="lbl">C.U.I.T.:</td><td>${esc(m.cuit)}</td></tr>${r("Horarios / turnos:",m.horarios)}<tr><td class="ttl" colspan="4">DATOS DE LA MEDICIÓN</td></tr>${r("Instrumento:",`${m.instrumento} ${m.marca} ${m.modelo} – N° de serie: ${m.serie||"—"}`)}${r("Fecha de calibración:",fdate(m.fechaCal))}${r("Metodología:",m.metodologia)}<tr><td class="lbl">Fecha de la medición:</td><td>${fdate(m.fechaMed)}</td><td class="lbl">Inicio / fin:</td><td>${esc(m.horaIni)} – ${esc(m.horaFin)}</td></tr>${r("Condiciones atmosféricas:",m.condiciones)}<tr><td class="ttl" colspan="4">DOCUMENTACIÓN QUE SE ADJUNTARÁ</td></tr>${r("Certificado de calibración:",m.docCert?"SÍ":"NO")}${r("Plano o croquis:",m.docCroquis?"SÍ":"NO")}${r("Observaciones:",m.observaciones)}</table>${sig}</div>`;
  const h2=`<div class="ph l"><table><tr><td class="ttl" colspan="10">PROTOCOLO PARA MEDICIÓN DE ILUMINACIÓN EN EL AMBIENTE LABORAL</td></tr><tr><td colspan="6"><b>Razón social:</b> ${esc(m.razon)}</td><td colspan="4"><b>C.U.I.T.:</b> ${esc(m.cuit)}</td></tr><tr><th>Punto</th><th>Hora</th><th>Sector</th><th>Sección / Puesto</th><th>Tipo de iluminación</th><th>Tipo de fuente</th><th>Iluminación</th><th>Uniformidad E mín ≥ E med/2</th><th>Valor medido (Lux)</th><th>Valor requerido Anexo IV</th></tr>${rowsHoja2().map((rw,i)=>{const s=status(S.puntos[i]);return`<tr>${rw.map((v,k)=>`<td class="${k===2||k===3?"":"c"} ${k===8&&(s.cls==="bad"||s.cls==="warn")?"bad":""}">${esc(k===8&&v!==""?fmt(v):v)}</td>`).join("")}</tr>`}).join("")}</table>${sig}</div>`;
  const h3=`<div class="ph"><table><tr><td class="ttl">ANÁLISIS DE LOS DATOS Y MEJORAS A REALIZAR</td></tr><tr><th>Conclusiones</th></tr><tr><td style="white-space:pre-wrap;height:60mm;vertical-align:top">${esc(S.conclusiones)}</td></tr><tr><th>Recomendaciones</th></tr><tr><td style="white-space:pre-wrap;height:80mm;vertical-align:top">${esc(S.recomendaciones)}</td></tr></table>${sig}</div>`;
  $("#print").innerHTML=h1+h2+h3;setTimeout(()=>window.print(),50);
}
function backup(){
  const name=`respaldo_iluminacion_${(S.meta.razon||"").replace(/\W+/g,"_")}_${S.meta.fechaMed}.json`;
  const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([JSON.stringify(S,null,1)],{type:"application/json"}));a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),2000);
}
/* ============ INGRESO ============ */
const LOGO=`<span class="logo"><svg viewBox="0 0 24 24"><path d="M9 18h6M10 21h4M12 3a6 6 0 0 0-3.5 10.9c.6.5 1 1.2 1 2V16h5v-.1c0-.8.4-1.5 1-2A6 6 0 0 0 12 3z"/></svg></span>`;
function vAuth(){return`<div class="auth">
  <div class="brand">${LOGO}<div><b>Protocolo de Iluminación</b><span>Res. SRT 84/12 · Dec. 351/79 Anexo IV</span></div></div>
  <div class="card stack">
   <label class="f"><span>Email</span><input id="aEmail" type="email" autocomplete="email" inputmode="email" placeholder="nombre@empresa.com"></label>
   <label class="f"><span>Contraseña</span><input id="aPass" type="password" autocomplete="current-password" placeholder="Mínimo 6 caracteres"></label>
   <button class="btn p wide" id="aLogin">Ingresar</button>
   <button class="btn wide" id="aSignup">Crear cuenta nueva</button>
   <button class="linkbtn" id="aForgot">Olvidé mi contraseña</button>
   <div id="aMsg" role="status"></div>
  </div>
  <p class="muted">Cada usuario ve solo sus propias mediciones. Los datos se guardan en el teléfono mientras medís y se suben a la nube cuando hay conexión.</p></div>`}
function vRecover(){return`<div class="auth"><div class="brand">${LOGO}<div><b>Nueva contraseña</b><span>Elegí una contraseña para tu cuenta</span></div></div>
  <div class="card stack"><label class="f"><span>Nueva contraseña</span><input id="rPass" type="password" autocomplete="new-password"></label>
  <button class="btn p wide" id="rSave">Guardar contraseña</button><div id="aMsg" role="status"></div></div></div>`}
function authErr(e){const m=String(e&&e.message||e||"");
  if(/Invalid login/i.test(m))return"Email o contraseña incorrectos.";
  if(/not confirmed/i.test(m))return"Todavía no confirmaste el email. Buscá el correo de confirmación (revisá spam).";
  if(/already registered|already been registered/i.test(m))return"Ese email ya tiene cuenta. Usá Ingresar.";
  if(/Password should be/i.test(m))return"La contraseña debe tener al menos 6 caracteres.";
  if(/fetch|network/i.test(m))return"Sin conexión. Para ingresar la primera vez necesitás internet.";
  if(/rate limit/i.test(m))return"Demasiados intentos. Esperá unos minutos.";
  return m||"No se pudo completar la operación.";}
function aMsg(t,cls="bad"){const el=$("#aMsg");if(el)el.innerHTML=t?`<div class="banner ${cls}">${esc(t)}</div>`:""}
function bindAuth(){
  const on=(id,fn)=>{const el=$("#"+id);if(el)el.onclick=fn};
  const creds=()=>({email:($("#aEmail").value||"").trim(),password:$("#aPass").value||""});
  on("aLogin",async()=>{const c=creds();if(!c.email||!c.password){aMsg("Completá email y contraseña.");return}aMsg("Ingresando…","info");
    const{error}=await sb.auth.signInWithPassword(c);if(error)aMsg(authErr(error))});
  on("aSignup",async()=>{const c=creds();if(!c.email||c.password.length<6){aMsg("Escribí tu email y una contraseña de al menos 6 caracteres.");return}aMsg("Creando cuenta…","info");
    const{data,error}=await sb.auth.signUp({...c,options:{emailRedirectTo:location.origin}});
    if(error){aMsg(authErr(error));return}if(!data.session)aMsg("Listo. Te enviamos un email para confirmar la cuenta; abrilo y después ingresá.","ok")});
  on("aForgot",async()=>{const email=($("#aEmail").value||"").trim();if(!email){aMsg("Escribí tu email arriba y volvé a tocar “Olvidé mi contraseña”.");return}
    const{error}=await sb.auth.resetPasswordForEmail(email,{redirectTo:location.origin});aMsg(error?authErr(error):"Te enviamos un email para elegir una nueva contraseña.",error?"bad":"ok")});
  on("rSave",async()=>{const p=$("#rPass").value||"";if(p.length<6){aMsg("Mínimo 6 caracteres.");return}const{error}=await sb.auth.updateUser({password:p});
    if(error){aMsg(authErr(error));return}view="list";render();toast("Contraseña actualizada")});
  const pw=$("#aPass");if(pw)pw.addEventListener("keydown",e=>{if(e.key==="Enter")$("#aLogin").click()});
}

/* ============ MIS MEDICIONES ============ */
function vList(){
  const L=meds();
  return`
 <div class="row" style="margin-bottom:10px"><button class="btn p" id="newMed" style="flex:1;min-height:52px">＋ Nuevo estudio</button></div>
 <div class="row" style="margin-bottom:14px">
  <label class="btn sm" style="flex:1">Importar protocolo de iluminación (.xlsx)<input type="file" id="impX" accept=".xlsx,.xls" hidden></label>
  <label class="btn sm">Importar respaldo<input type="file" id="impJ" accept=".json,application/json" hidden></label></div>
 ${L.length?`<div class="mlist">${L.map(x=>{const d=x.data;let tot,med,bad,warn;
   if(d.tipo&&MODS[d.tipo]){const su=MODS[d.tipo].summary(d);tot=su.tot;med=su.done;bad=su.bad;warn=su.warn}
   else{const st=d.puntos.map(status);tot=d.puntos.length;med=d.puntos.filter(p=>calc(p).n).length;bad=st.filter(s=>s.cls==="bad").length;warn=st.filter(s=>s.cls==="warn").length}
   const tipoTxt=d.tipo&&MODS[d.tipo]?MODS[d.tipo].nombre:"Iluminación",unidad=(d.tipo&&MODS[d.tipo]&&MODS[d.tipo].unidad)||"puntos";
   return`<div class="mcard"><div style="min-width:0"><div class="t">${esc(d.meta.razon||"Empresa sin nombre")}</div>
     <div class="s"><b style="color:var(--acc)">${tipoTxt}</b> · ${fdate(d.meta.fechaMed)||"sin fecha"} · ${med}/${tot} ${unidad}</div></div>
     <div class="row" style="justify-content:flex-end">${bad?`<span class="pill bad">${bad} no cumplen</span>`:""}${warn?`<span class="pill warn">${warn} obs.</span>`:""}${tot&&med===tot&&!bad&&!warn?`<span class="pill ok">Cumple</span>`:""}</div>
     <div class="bar"><i style="width:${tot?Math.round(med/tot*100):0}%"></i></div>
     <div class="acts"><button class="btn sm p" data-open="${x.id}">Abrir</button><button class="btn sm" data-rep="${x.id}">Repetir estudio</button><button class="btn sm d" data-del="${x.id}">Eliminar</button></div></div>`}).join("")}</div>`
 :`<div class="card empty"><b>Todavía no tenés estudios</b>Creá uno nuevo (iluminación, carga de fuego, puesta a tierra, ruido, vibraciones o estrés térmico), o importá el Excel del protocolo de iluminación del año anterior para comparar.</div>`}
 <hr class="s">
 ${CLOUD?`<div class="row"><span class="muted" style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis">Sesión: ${esc(session&&session.user.email)}</span><button class="btn sm" id="logout">Cerrar sesión</button></div>`
 :`<div class="banner warn">Modo local: los datos quedan solo en este dispositivo. Para guardarlos en la nube configurá Supabase (ver README).</div>`}`;
}
function openMed(id){if(!lib[id])return;curId=id;S=lib[id].data;Object.values(MODS).forEach(m=>m.reset());view="med";tab="pts";editing=null;$$("#tabs button").forEach(b=>b.classList.toggle("on",b.dataset.t==="pts"));render();scrollTo(0,0)}
function repeatFrom(src){
  if(src.tipo&&MODS[src.tipo])return MODS[src.tipo].repeat(src);
  const d=JSON.parse(JSON.stringify(src));const old=S;S=d;
  d.meta.prevFecha=d.meta.fechaMed;
  d.puntos.forEach(p=>{const c=calc(p);p.prev=c.n?{emed:c.emed,emin:c.emin}:p.prev;p.lect=[];p.hora="";p.egen="";p.id=uid()});
  S=old;
  Object.assign(d.meta,{fechaMed:today(),horaIni:"",horaFin:"",condiciones:"",cielo:"",temp:""});
  d.conclusiones="";d.recomendaciones="";d.updatedAt=Date.now();return d;
}
function bindList(){
  const on=(id,fn)=>{const el=$("#"+id);if(el)el.onclick=fn};
  on("newMed",async()=>{const r=await chooseNew();if(!r)return;const id=newId();const d=r.tipo==="ilum"?blankMed():MODS[r.tipo].nuevo();
    if(r.from&&lib[r.from]){const sm=lib[r.from].data.meta;["razon","cuit","direccion","localidad","provincia","cp","horarios"].forEach(k=>{d.meta[k]=sm[k]||""})}
    putMed(id,d);openMed(id);go(r.from?"pts":"est")});
  on("logout",async()=>{if(Object.values(lib).some(x=>x.dirty)&&!await ask("Hay cambios que todavía no se subieron a la nube. Si cerrás sesión ahora quedan guardados en este teléfono y se suben la próxima vez que ingreses.",{ok:"Cerrar sesión"}))return;await sb.auth.signOut()});
  $$("[data-open]").forEach(b=>b.onclick=()=>openMed(b.dataset.open));
  $$("[data-rep]").forEach(b=>b.onclick=()=>{const src=lib[b.dataset.rep];const id=newId();putMed(id,repeatFrom(src.data));openMed(id);go("est");toast("Nueva medición creada con los valores anteriores para comparar")});
  $$("[data-del]").forEach(b=>b.onclick=async()=>{const x=lib[b.dataset.del];if(!await ask(`¿Eliminar la medición de ${x.data.meta.razon||"esta empresa"} del ${fdate(x.data.meta.fechaMed)}?`,{ok:"Eliminar",danger:true}))return;
    x.deleted=true;x.dirty=true;x.updatedAt=Date.now();x.data.updatedAt=x.updatedAt;writeCache();schedulePush(300);render()});
  const ij=$("#impJ");if(ij)ij.onchange=()=>{const f=ij.files[0];if(!f)return;const r=new FileReader();r.onload=()=>{try{const d=JSON.parse(r.result);if(!valid(d))throw 0;d.updatedAt=Date.now();const id=newId();putMed(id,d);openMed(id);toast("Respaldo importado")}catch(e){toast("El archivo no es un respaldo válido")}};r.readAsText(f)};
  const ix=$("#impX");if(ix)ix.onchange=()=>{const f=ix.files[0];if(!f)return;const r=new FileReader();r.onload=()=>{try{const d=importProtocol(r.result);const id=newId();putMed(id,d);openMed(id);go("est");toast(`Importado: ${d.puntos.length} puntos`)}catch(e){console.warn(e);toast(e&&e.userMsg||"No pude leer ese Excel. ¿Es un protocolo de iluminación?")}};r.readAsArrayBuffer(f)};
}

function chooseNew(){
  return new Promise(res=>{
    const emp={};meds().forEach(x=>{const k=(x.data.meta.razon||"").trim();if(k&&!emp[k])emp[k]=x.id});
    const s=document.createElement("div");s.className="scrim";
    s.innerHTML=`<div class="dlg" role="dialog" aria-modal="true"><p style="font-weight:700;font-size:17px;margin-bottom:10px">Nuevo estudio</p>
     ${Object.keys(emp).length?`<label class="f" style="margin-bottom:12px"><span>Empresa</span><select id="nEmp"><option value="">Empresa nueva (cargar datos)</option>${Object.entries(emp).map(([k,id])=>`<option value="${id}">${esc(k)}</option>`).join("")}</select></label>`:""}
     <div class="stack">
      ${[["ilum","💡 Iluminación","Res. SRT 84/12"],["fuego","🔥 Carga de fuego","Dec. 351/79 Anexo VII"],["pat","⚡ Puesta a tierra","Res. SRT 900/15"],["ruido","🔊 Ruido","Res. SRT 85/12"],["vib","📳 Vibraciones","Res. 295/03"],["termo","🌡️ Termohigrometría / estrés térmico","Res. SRT 30/23"]].map(([t,n,r])=>`<button class="btn p wide" data-t="${t}" style="justify-content:flex-start;min-height:52px;text-align:left">${n}<span style="font-weight:400;opacity:.85;margin-left:auto;font-size:12.5px">${r}</span></button>`).join("")}
      <button class="btn wide" data-t="">Cancelar</button></div></div>`;
    s.addEventListener("click",e=>{const b=e.target.closest("[data-t]");if(b||e.target===s){const sel=$("#nEmp",s);s.remove();res(b&&b.dataset.t?{tipo:b.dataset.t,from:sel?sel.value:""}:null)}});
    document.body.appendChild(s);
  });
}

/* ============ IMPORTAR PROTOCOLO ANTERIOR (.xlsx) ============ */
function importProtocol(buf){
  const wb=XLSX.read(buf,{type:"array",cellDates:true});
  const sheets=wb.SheetNames.map(n=>XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true,defval:""}));
  const filled=v=>v!==""&&v!=null&&!(typeof v==="string"&&!v.trim());
  const isLabel=v=>typeof v==="string"&&/:\s*$/.test(v.trim());
  function find(re,order=["right","below"]){
    for(const a of sheets)for(let r=0;r<a.length;r++){const row=a[r];for(let c=0;c<row.length;c++){const v=row[c];if(typeof v!=="string"||!re.test(v))continue;
      const after=v.split(/:\s*/).slice(1).join(":").trim();if(after)return after;
      for(const o of order){
        if(o==="right"){for(let k=c+1;k<row.length;k++){if(filled(row[k])){if(!isLabel(row[k]))return row[k];break}}}
        if(o==="below"){for(let k=r+1;k<Math.min(a.length,r+3);k++){const w=(a[k]||[])[c];if(filled(w)){if(!isLabel(w))return w;break}}}
      }}}
    return"";
  }
  const toISO=v=>{if(v instanceof Date&&!isNaN(v))return new Date(v-v.getTimezoneOffset()*6e4).toISOString().slice(0,10);
    if(typeof v==="number"&&v>20000&&v<80000){const d=new Date(Math.round((v-25569)*864e5));return d.toISOString().slice(0,10)}
    const m=/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/.exec(String(v||""));if(m){const y=m[3].length===2?"20"+m[3]:m[3];return`${y}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`}return""};
  // tabla de puntos
  let tbl=null;
  for(const a of sheets){for(let r=0;r<a.length;r++){const c=a[r].findIndex(v=>typeof v==="string"&&/punto de muestreo/i.test(v));if(c===0){tbl={a,h:r};break}}if(tbl)break}
  if(!tbl){const e=new Error("sin tabla");e.userMsg="No encontré la tabla de puntos de muestreo en ese Excel.";throw e}
  const hdr=tbl.a[tbl.h].map(v=>String(v||"").toLowerCase());
  const lay84=/hora/.test(hdr[1]||"");
  const col=lay84?{n:0,hora:1,sector:2,puesto:3,tipo:4,fuente:5,ilum:6,unif:7,med:8,req:9}:{n:0,sector:1,puesto:2,tipo:3,fuente:4,ilum:5,unif:6,med:7,req:8};
  const cap=(v,opts,def)=>{const s=String(v||"").trim().toLowerCase();return opts.find(o=>o.toLowerCase()===s)||def};
  const puntos=[];let started=false;
  for(let r=tbl.h+1;r<tbl.a.length;r++){const row=tbl.a[r];const n=num(row[col.n]);
    if(isFinite(n)&&filled(row[col.sector])){started=true;
      const emed=num(row[col.med]),emin=num(String(row[col.unif]||"").split(/[≥>]/)[0]);
      const rq=String(row[col.req]||"").replace(/\s/g,"");const mm=/^(\d+)(?:-(\d+))?$/.exec(rq);let req="t3",reqMin="",reqMax="";
      if(mm){const mn=+mm[1],mx=mm[2]?+mm[2]:null;const t=T1.find(t=>t.min===mn&&(t.max||null)===mx);if(t)req=t.id;else{req="custom";reqMin=String(mn);reqMax=mx?String(mx):""}}
      puntos.push(newPoint(n,{sector:String(row[col.sector]).trim(),puesto:String(row[col.puesto]||"").trim(),
        tipoIlum:cap(row[col.tipo],["Natural","Artificial","Mixta"],"Mixta"),fuente:cap(row[col.fuente],["Incandescente","Descarga","Mixta"],"Descarga"),ilum:cap(row[col.ilum],["General","Localizada","Mixta"],"General"),
        req,reqMin,reqMax,prev:isFinite(emed)?{emed,emin:isFinite(emin)?emin:null}:null}));
    }else if(started&&!filled(row[col.n]))break;
  }
  if(!puntos.length){const e=new Error("sin puntos");e.userMsg="La tabla de puntos del Excel está vacía.";throw e}
  const meta=blankMeta();
  const s=v=>String(v??"").trim();
  meta.razon=s(find(/raz[oó]n\s*social/i));meta.cuit=s(find(/^c\.?u\.?i\.?t/i));meta.direccion=s(find(/^direcci[oó]n/i));
  meta.localidad=s(find(/^localidad/i));meta.provincia=s(find(/^provincia/i));meta.cp=s(find(/^c\.\s*p\./i)).replace(/\.0$/,"");
  meta.horarios=s(find(/horarios/i,["right","below"]));
  const marca=s(find(/^marca:?\s*$/i)),modelo=s(find(/^modelo:?\s*$/i));if(marca)meta.marca=marca;if(modelo)meta.modelo=modelo;
  const cal=toISO(find(/fecha de calibraci/i));if(cal)meta.fechaCal=cal;
  meta.prevFecha=toISO(find(/fecha de la medici/i,["below","right"]));
  return{updatedAt:Date.now(),meta,puntos,conclusiones:"",recomendaciones:""};
}

/* ============ ARRANQUE ============ */
function bindTop(){$("#hBack").onclick=()=>{editing=null;view="list";curId=null;S=null;wakeOff();render();scrollTo(0,0)}}
async function boot(){
  if(CLOUD){
    try{const{data}=await sb.auth.getSession();session=data.session}catch(e){session=null}
    sb.auth.onAuthStateChange((ev,s)=>{
      if(ev==="PASSWORD_RECOVERY"){session=s;view="recover";render();return}
      const was=session&&session.user.id,now=s&&s.user.id;session=s;
      if(was!==now){curId=null;S=null;editing=null;loadCache();view=s?"list":"auth";render();if(s)sync()}
    });
    if(!session){view="auth";render();return}
  }
  loadCache();view="list";render();sync();
}
if("serviceWorker" in navigator&&location.protocol==="https:")window.addEventListener("load",()=>navigator.serviceWorker.register("sw.js").catch(()=>{}));
boot();

