"use strict";
/* =====================================================================
   Estudios adicionales: CARGA DE FUEGO (Dec. 351/79 Anexo VII, Cap. 18)
   y PUESTA A TIERRA (Res. SRT 900/15 – AEA 90364-7-771).
   Usa los helpers globales de app.js ($, $$, esc, num, fmt, fdate, fld,
   dfld, tfld, segF, ask, toast, save, rerender, S, ...).
   ===================================================================== */

/* ---------- utilidades Excel comunes ---------- */
const XB={
  T:{style:"thin",color:{rgb:"000000"}},
  F:(sz=10,b=false)=>({name:"Arial",sz,bold:b}),
  mk:()=>({cells:{},merges:[]}),
  put(w,ref,v,st={}){const c=w.cells[ref]||(w.cells[ref]={});c.v=v;c.s=Object.assign(c.s||{},st)},
  box(w,rng){const r=XLSX.utils.decode_range(rng);for(let R=r.s.r;R<=r.e.r;R++)for(let C=r.s.c;C<=r.e.c;C++){const ref=XLSX.utils.encode_cell({r:R,c:C});const c=w.cells[ref]||(w.cells[ref]={v:""});c.s=c.s||{};const b=Object.assign({},c.s.border||{});if(R===r.s.r)b.top=XB.T;if(R===r.e.r)b.bottom=XB.T;if(C===r.s.c)b.left=XB.T;if(C===r.e.c)b.right=XB.T;c.s.border=b}},
  grid(w,rng){const r=XLSX.utils.decode_range(rng);for(let R=r.s.r;R<=r.e.r;R++)for(let C=r.s.c;C<=r.e.c;C++)XB.box(w,XLSX.utils.encode_cell({r:R,c:C}))},
  fill(w,rng,rgb){const r=XLSX.utils.decode_range(rng);for(let R=r.s.r;R<=r.e.r;R++)for(let C=r.s.c;C<=r.e.c;C++){const ref=XLSX.utils.encode_cell({r:R,c:C});const c=w.cells[ref]||(w.cells[ref]={v:""});c.s=Object.assign(c.s||{},{fill:{patternType:"solid",fgColor:{rgb}}})}},
  mg(w,rng){w.merges.push(XLSX.utils.decode_range(rng))},
  build(w,cols,rows,margins){const ws={};let mr=0,mc=0;
    for(const[ref,c]of Object.entries(w.cells)){const a=XLSX.utils.decode_cell(ref);mr=Math.max(mr,a.r);mc=Math.max(mc,a.c);const v=c.v;const cell=typeof v==="number"?{t:"n",v}:{t:"s",v:String(v??"")};cell.s=Object.assign({font:XB.F(10),alignment:{vertical:"center",wrapText:true}},c.s);if(c.s&&c.s.numFmt)cell.z=c.s.numFmt;ws[ref]=cell}
    ws["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:mr,c:Math.max(mc,cols.length-1)}});ws["!merges"]=w.merges;ws["!cols"]=cols.map(x=>({wch:Math.max(1,x-0.73)}));
    ws["!rows"]=rows||[];if(margins)ws["!margins"]=margins;return ws},
  lines:(t,ch)=>Math.max(1,String(t||"").split("\n").reduce((a,l)=>a+Math.max(1,Math.ceil(l.length/ch)),0)),
  fname:(pre)=>`${pre}_${(S.meta.razon||"empresa").replace(/[^A-Za-z0-9]+/g,"_")}_${S.meta.fechaMed||""}.xlsx`
};
const C_={horizontal:"center",vertical:"center",wrapText:true},L_={horizontal:"left",vertical:"center",wrapText:true};

/* ---------- bloque común: establecimiento ---------- */
function estBlock(open){return`<details class="more" ${open?"open":""}><summary>Establecimiento · ${esc(S.meta.razon||"sin datos")}</summary><div class="in"><div class="grid">
  ${fld("meta.razon","Razón social",{req:1})}${fld("meta.cuit","C.U.I.T.",{req:1,im:"numeric"})}${fld("meta.direccion","Dirección",{req:1})}
  ${fld("meta.localidad","Localidad",{req:1})}${fld("meta.provincia","Provincia",{req:1})}${fld("meta.cp","C.P.",{req:1,im:"numeric"})}
  ${fld("meta.horarios","Horarios / turnos habituales de trabajo",{full:1})}</div></div></details>`}
function profBlock(){return`<details class="more"><summary>Profesional</summary><div class="in"><div class="grid">
  ${fld("meta.profesional","Profesional",{req:1})}${fld("meta.matricula","Título y matrícula",{req:1})}</div></div></details>`}
function checksList(ch){return ch.length?`<div class="banner warn">Faltan ${ch.length} datos:</div><ul class="chk-list">${ch.map(x=>`<li>${esc(x)}</li>`).join("")}</ul>`:`<div class="banner ok">Estudio completo.</div>`}
function baseChecks(keys){const o=[];keys.forEach(([k,t])=>{if(!String(S.meta[k]??"").trim())o.push(t)});return o}
const EST_KEYS=[["razon","Razón social"],["cuit","CUIT"],["direccion","Dirección"],["localidad","Localidad"],["provincia","Provincia"],["cp","C.P."],["fechaMed","Fecha"],["profesional","Profesional"],["matricula","Matrícula"]];
function textCard(title,genLabel){return`<div class="card"><h2>${title}</h2>
  <button class="btn p wide" id="genTxt" style="margin-bottom:10px">${genLabel}</button>
  ${fld("conclusiones","Conclusiones",{type:"textarea",rows:8})}<div style="height:10px"></div>
  ${fld("recomendaciones","Recomendaciones",{type:"textarea",rows:7})}</div>`}
function expCard(ch,btnLabel,desc){return`<div class="card" style="${ch.length?"":"border-color:var(--ok)"}"><h2>Descargar</h2>
  <button class="btn p wide" id="dlX" style="min-height:58px;font-size:17px">⬇ ${btnLabel}${ch.length?"":" ✓"}</button>
  <p class="muted">${desc}</p>${checksList(ch)}</div>
  <div class="card"><h2>Este estudio</h2><div class="row"><button class="btn sm" id="bkp">Descargar respaldo (.json)</button><button class="btn sm" id="repMed">Repetir estudio (nuevo)</button><button class="btn sm d" id="delMed">Eliminar</button></div></div>`}
async function confirmDl(ch,fn){if(ch.length&&!await ask(`Faltan ${ch.length} datos:\n• ${ch.slice(0,8).join("\n• ")}${ch.length>8?"\n…":""}\n\n¿Descargar igual?`,{ok:"Descargar igual"}))return;doExport(fn)}

/* =====================================================================
   CARGA DE FUEGO
   ===================================================================== */
const PC_MADERA=4400; // kcal/kg (18,41 MJ/kg)
const RIESGOS=[[1,"Explosivo"],[2,"Inflamable"],[3,"Muy combustible"],[4,"Combustible"],[5,"Poco combustible"],[6,"Incombustible"],[7,"Refractario"]];
const ACTIVIDADES=["Residencial / Administrativo","Comercial / Industrial / Depósito","Espectáculos / Cultura"];
const NP_ACT={"Residencial / Administrativo":[1,2,5,6,7],"Espectáculos / Cultura":[1,2,5,6,7],"Comercial / Industrial / Depósito":[]};
const RANGOS=["Hasta 15 kg/m²","16 a 30 kg/m²","31 a 60 kg/m²","61 a 100 kg/m²","Más de 100 kg/m²"];
// Cuadro 2.2.1 (ventilación natural) y 2.2.2 (ventilación mecánica) – columnas Riesgo 1..5
const RF={natural:[["—","F 60","F 30","F 30","—"],["—","F 90","F 60","F 30","F 30"],["—","F 120","F 90","F 60","F 30"],["—","F 180","F 120","F 90","F 60"],["—","F 180","F 180","F 120","F 90"]],
          mecanica:[["—","NP","F 60","F 60","F 30"],["—","NP","F 90","F 60","F 60"],["—","NP","F 120","F 90","F 60"],["—","NP","F 180","F 120","F 90"],["—","NP","NP","F 180","F 120"]]};
// Tabla 1 (fuegos clase A) y Tabla 2 (clase B) – columnas Riesgo 1..5
const EXT_A=[["—","—","1A","1A","1A"],["—","—","2A","1A","1A"],["—","—","3A","2A","1A"],["—","—","6A","4A","3A"],["A determinar","A determinar","A determinar","A determinar","A determinar"]];
const EXT_B=[["—","6B","4B","—","—"],["—","8B","6B","—","—"],["—","10B","8B","—","—"],["—","20B","10B","—","—"],["A determinar","A determinar","A determinar","A determinar","A determinar"]];
// Poderes caloríficos de referencia (kcal/kg) – orientativos, editables en cada fila
const MATS=[["Madera",4400],["Papeles y otros",4000],["Cartón",4000],["Sillas (tapizado / espuma)",6000],["Sillas e informáticos",6000],["Equipos informáticos",6000],["Cortinas / textiles",5000],["Cartel",5000],["Algodón",4000],["Lana",5000],["Cuero",4500],["PVC",4300],["Poliestireno",9600],["Polietileno",11000],["Caucho",10000],["Espuma de poliuretano",6000],["Nafta",10500],["Gasoil",10200],["Aceite mineral",10000],["Alcohol etílico",6400]];
function newSector(o={}){return Object.assign({id:uid(),nombre:"",planta:"Planta baja",sup:"",riesgo:4,clase:"A",materiales:[{id:uid(),desc:"Madera",kg:"",pc:"4400"}],obs:"",prev:null},o)}
function newFuego(){const m=blankMeta();m.actividad="Residencial / Administrativo";m.ventilacion="natural";
  m.metodologia="Se adopta lo establecido en el Decreto 351/79, Anexo VII, Capítulo 18: suma de los pesos de los materiales combustibles por sus respectivos poderes caloríficos, considerando combustión completa y en ausencia de dispersiones, referida a la madera (4400 kcal/kg).";
  return{tipo:"fuego",updatedAt:Date.now(),meta:m,sectores:[],conclusiones:"",recomendaciones:""}}
function rangoIdx(q){return q<=15?0:q<=30?1:q<=60?2:q<=100?3:4}
function calcSector(s){
  const rows=(s.materiales||[]).map(m=>{const kg=num(m.kg),pc=num(m.pc);return{kg,pc,q:isFinite(kg)&&isFinite(pc)?kg*pc:NaN}});
  const Q=rows.reduce((a,r)=>a+(isFinite(r.q)?r.q:0),0),sup=num(s.sup),kgm=Q/PC_MADERA,qf=sup>0?kgm/sup:NaN;
  const r=+s.riesgo||4,ri=isFinite(qf)?rangoIdx(qf):null,vent=(S&&S.meta.ventilacion)||"natural";
  const o={rows,Q,kgm,sup,qf,ri,r};
  if(ri!=null&&r<=5){o.F=RF[vent][ri][r-1];o.extA=EXT_A[ri][r-1];o.extB=/B/.test(s.clase||"")?EXT_B[ri][r-1]:"—"}
  else if(ri!=null){o.F="—";o.extA="—";o.extB="—"}
  o.np=(NP_ACT[(S&&S.meta.actividad)]||[]).includes(r);
  return o}
function fuegoSummary(d){const old=S;S=d;const L=d.sectores.map(calcSector);S=old;const done=L.filter(c=>isFinite(c.qf)).length;
  return{tot:d.sectores.length,done,bad:L.filter(c=>c.np||c.F==="NP").length,warn:0,label:"Carga de fuego"}}
function fuegoChecks(){const o=baseChecks(EST_KEYS);if(!S.sectores.length)o.push("Al menos un sector");
  S.sectores.forEach((s,i)=>{const c=calcSector(s),mi=[];if(!s.nombre)mi.push("nombre");if(!(c.sup>0))mi.push("superficie");if(!(c.Q>0))mi.push("materiales");if(mi.length)o.push(`Sector ${i+1}: ${mi.join(", ")}`)});
  if(!S.conclusiones.trim())o.push("Conclusiones");return o}
let fEdit=null; // sector en edición
function vFuego(tab){
  if(tab==="est")return`
   <div class="card"><h2>Estudio de carga de fuego</h2><div class="grid">
    ${dfld("meta.fechaMed","Fecha del relevamiento")}
    ${fld("meta.actividad","Actividad predominante (Tabla 2.1)",{opts:ACTIVIDADES})}
    ${fld("meta.ventilacion","Ventilación de los sectores",{opts:[["natural","Natural · Cuadro 2.2.1"],["mecanica","Mecánica · Cuadro 2.2.2"]]})}</div>
    <p class="muted">Carga de fuego = Σ (Pi × Ki) / (4400 kcal/kg × S). La resistencia al fuego y los extintores salen del Anexo VII según riesgo y carga de fuego.</p></div>
   ${estBlock(!S.meta.razon)}
   <details class="more"><summary>Metodología y observaciones</summary><div class="in">${fld("meta.metodologia","Metodología",{type:"textarea",rows:4})}<div style="height:8px"></div>${fld("meta.observaciones","Observaciones",{type:"textarea",rows:2})}</div></details>
   ${profBlock()}`;
  if(tab==="pts")return fEdit!=null&&S.sectores[fEdit]?vSector(fEdit):vSectores();
  if(tab==="con"){const L=S.sectores.map(calcSector);const supT=L.reduce((a,c)=>a+(c.sup>0?c.sup:0),0);
    const plantas={};S.sectores.forEach((s,i)=>{const k=s.planta||"Sin planta";plantas[k]=(plantas[k]||0)+(L[i].sup>0?L[i].sup:0)});
    return`<div class="card"><h2>Resumen por sector</h2><div class="tablewrap"><table class="sum"><tr><th>Sector</th><th>Sup.</th><th>qf kg/m²</th><th>R</th><th>F</th><th>Ext.</th></tr>
     ${S.sectores.map((s,i)=>{const c=L[i];return`<tr><td>${esc(s.nombre)}<br><span class="muted">${esc(s.planta)}</span></td><td>${fmt(c.sup,2)}</td><td><b>${fmt(c.qf,1)}</b></td><td>${s.riesgo}</td><td>${c.F||"–"}</td><td>${c.extA||"–"}${c.extB&&c.extB!=="—"?" · "+c.extB:""}</td></tr>`}).join("")}</table></div>
     <hr class="s"><p class="muted" style="margin:0">Mínimo por planta (Anexo VII, 7.1.1): 1 extintor cada 200 m² o fracción, potencial mínimo 1A y 5BC.<br>${Object.entries(plantas).map(([k,v])=>`${esc(k)}: ${fmt(v,2)} m² → <b>${Math.max(1,Math.ceil(v/200))}</b> extintor(es)`).join("<br>")}<br>Superficie total relevada: ${fmt(supT,2)} m². Distancia máxima a recorrer hasta el extintor: 20 m (clase A) y 15 m (clase B), art. 176 Dec. 351/79.</p></div>
     ${textCard("Conclusión carga de fuego","Generar conclusión automáticamente")}`}
  return expCard(fuegoChecks(),"Descargar Excel carga de fuego","Una hoja con la tabla de cálculo de cada sector (como en el informe) y otra con el resumen: riesgo, resistencia al fuego exigible, extintores y conclusión.");
}
function vSectores(){
  const L=S.sectores.map(calcSector);
  return`${S.sectores.length?`<div class="plist">${S.sectores.map((s,i)=>{const c=L[i];return`
   <button class="pitem" data-fs="${i}"><span class="num">${i+1}</span><span style="min-width:0"><span class="t" style="display:block">${esc(s.nombre||"(sin nombre)")}</span>
   <span class="s" style="display:block">${esc(s.planta)} · ${fmt(c.sup,2)} m² · Riesgo ${s.riesgo} · Clase ${esc(s.clase)}</span></span>
   <span class="v"><b>${fmt(c.qf,1)}</b><span class="pill ${!isFinite(c.qf)?"na":c.np?"bad":"ok"}">${!isFinite(c.qf)?"Sin datos":c.np?"Riesgo NP":(c.F||"")}</span></span></button>`}).join("")}</div>`
   :`<div class="card empty"><b>Sin sectores</b>Agregá cada sector de incendio con su superficie y los materiales combustibles.</div>`}
   <button class="btn p wide" id="addS" style="margin-top:12px;min-height:52px">＋ Agregar sector</button>
   <p class="muted" style="margin-top:10px">El número grande es la carga de fuego en kg/m² de madera equivalente.</p>`}
function vSector(i){
  const s=S.sectores[i];
  return`<div class="ptop"><button class="navb" id="fBack" aria-label="Volver a sectores">☰</button>
   <div class="ttl"><b>${i+1} · ${esc(s.nombre||"Sector nuevo")}</b><span>${esc(s.planta)}</span></div>
   <button class="navb" id="fPrev" ${i===0?"disabled":""}>‹</button><button class="navb" id="fNext" ${i===S.sectores.length-1?"disabled":""}>›</button></div>
  <div class="card" id="fLive"></div>
  <div class="card"><div class="grid">
   ${fld(`sectores.${i}.nombre`,"Sector",{req:1})}${fld(`sectores.${i}.planta`,"Planta / nivel")}
   ${fld(`sectores.${i}.sup`,"Superficie (m²)",{req:1,im:"decimal"})}
   ${fld(`sectores.${i}.clase`,"Clase de fuego",{opts:["A","B","C","AC","AB","ABC","D","K"]})}</div>
   <div style="margin-top:10px">${fld(`sectores.${i}.riesgo`,"Riesgo del sector (Tabla 2.1)",{opts:RIESGOS.map(([n,t])=>[n,`Riesgo ${n} · ${t}`])})}</div></div>
  <div class="card"><h2>Materiales combustibles</h2>
   <div id="matList">${s.materiales.map((m,j)=>matRow(i,j,m)).join("")}</div>
   <button class="btn wide" id="addM" style="margin-top:8px">＋ Agregar material</button>
   <datalist id="matLib">${MATS.map(([n])=>`<option value="${esc(n)}">`).join("")}</datalist>
   <p class="muted">Al elegir un material de la lista se completa un poder calorífico de referencia; ajustalo si tenés el dato del material real.</p></div>
  <div class="card">${fld(`sectores.${i}.obs`,"Observaciones del sector",{type:"textarea",rows:2})}
   <div class="row" style="margin-top:10px"><button class="btn sm" id="dupS">Duplicar sector</button><button class="btn sm d" id="delS">Eliminar sector</button></div></div>
  <button class="btn p wide" id="fNext2" style="min-height:52px">${i===S.sectores.length-1?"Listo ✓":"Siguiente sector ›"}</button>`}
function matRow(i,j,m){const kg=num(m.kg),pc=num(m.pc);return`<div class="mrow" data-m="${j}" style="display:grid;grid-template-columns:minmax(0,1.6fr) minmax(0,.8fr) minmax(0,.9fr) 40px;gap:6px;align-items:end;margin-bottom:8px">
  <label class="f"><span>${j===0?"Material":""}</span><input data-b="sectores.${i}.materiales.${j}.desc" list="matLib" value="${esc(m.desc)}" autocomplete="off"></label>
  <label class="f"><span>${j===0?"kg":""}</span><input data-b="sectores.${i}.materiales.${j}.kg" inputmode="decimal" value="${esc(m.kg)}"></label>
  <label class="f"><span>${j===0?"kcal/kg":""}</span><input data-b="sectores.${i}.materiales.${j}.pc" inputmode="decimal" value="${esc(m.pc)}"></label>
  <button class="navb" data-delm="${j}" aria-label="Quitar material" style="width:40px;height:46px">✕</button>
  <span class="muted" style="grid-column:1/-1;margin-top:-4px" data-mq="${j}">${isFinite(kg)&&isFinite(pc)?fmt(kg*pc,0)+" kcal":""}</span></div>`}
function fLiveUpd(){const box=$("#fLive");if(!box||fEdit==null)return;const s=S.sectores[fEdit];const c=calcSector(s);
  s.materiales.forEach((m,j)=>{const el=$(`[data-mq="${j}"]`);if(el){const kg=num(m.kg),pc=num(m.pc);el.textContent=isFinite(kg)&&isFinite(pc)?fmt(kg*pc,0)+" kcal":""}});
  box.innerHTML=`<div class="live"><div><small>Calor total</small><b>${fmt(c.Q/1000,0)}</b><small>Mcal</small></div><div><small>Madera eq.</small><b>${fmt(c.kgm,0)}</b><small>kg</small></div><div><small>Carga de fuego</small><b>${fmt(c.qf,1)}</b><small>kg/m²</small></div></div>
   <div class="verdict">${isFinite(c.qf)?`<span class="pill na">${RANGOS[c.ri]}</span><span class="pill ${c.F==="NP"?"bad":"ok"}">Resistencia ${c.F}</span><span class="pill ok">Ext. A: ${c.extA}</span>${c.extB!=="—"?`<span class="pill ok">Ext. B: ${c.extB}</span>`:""}`:`<span class="pill na">Cargá superficie y materiales</span>`}
   ${c.np?`<span class="pill bad">Riesgo ${c.r} no permitido para ${esc(S.meta.actividad)}</span>`:""}${s.prev&&isFinite(s.prev.qf)?`<span class="pill na">Anterior ${fmt(s.prev.qf,1)}</span>`:""}</div>`}
function bindFuego(tab){
  const on=(id,fn)=>{const el=$("#"+id);if(el)el.onclick=fn};
  $$("[data-fs]").forEach(b=>b.onclick=()=>{fEdit=+b.dataset.fs;render();scrollTo(0,0)});
  on("addS",()=>{const last=S.sectores.at(-1);S.sectores.push(newSector(last?{planta:last.planta,riesgo:last.riesgo,clase:last.clase}:{}));fEdit=S.sectores.length-1;save();render();scrollTo(0,0)});
  on("fBack",()=>{fEdit=null;render()});on("fPrev",()=>{fEdit--;render();scrollTo(0,0)});on("fNext",()=>{fEdit++;render();scrollTo(0,0)});
  on("fNext2",()=>{if(fEdit<S.sectores.length-1){fEdit++}else fEdit=null;render();scrollTo(0,0)});
  on("addM",()=>{S.sectores[fEdit].materiales.push({id:uid(),desc:"",kg:"",pc:""});save();rerender();const ins=$$("#matList input[list]");ins.at(-1)&&ins.at(-1).focus()});
  $$("[data-delm]").forEach(b=>b.onclick=()=>{S.sectores[fEdit].materiales.splice(+b.dataset.delm,1);save();rerender()});
  on("dupS",()=>{const c=JSON.parse(JSON.stringify(S.sectores[fEdit]));c.id=uid();c.nombre+=" (copia)";c.prev=null;S.sectores.splice(fEdit+1,0,c);fEdit++;save();render();scrollTo(0,0)});
  on("delS",async()=>{if(!await ask("¿Eliminar este sector y sus materiales?",{ok:"Eliminar",danger:true}))return;S.sectores.splice(fEdit,1);fEdit=null;save();render()});
  on("genTxt",async()=>{if((S.conclusiones.trim()||S.recomendaciones.trim())&&!await ask("¿Reemplazar el texto actual?",{ok:"Reemplazar"}))return;fuegoText();save();render()});
  on("dlX",()=>confirmDl(fuegoChecks(),exportFuego));
  if(tab==="pts")fLiveUpd();
}
function fuegoChange(path,el){
  const m=/^sectores\.(\d+)\.materiales\.(\d+)\.desc$/.exec(path);
  if(m){const mat=MATS.find(([n])=>n.toLowerCase()===String(el.value).trim().toLowerCase());if(mat){const row=S.sectores[+m[1]].materiales[+m[2]];if(!String(row.pc).trim()||MATS.some(([,v])=>String(v)===String(row.pc))){row.pc=String(mat[1]);const pi=$(`[data-b="sectores.${m[1]}.materiales.${m[2]}.pc"]`);if(pi)pi.value=row.pc;save()}}}
  if(/^sectores\.\d+\.(riesgo|clase)$/.test(path))S.sectores[fEdit].riesgo=+S.sectores[fEdit].riesgo||4;
  if(/^meta\.(actividad|ventilacion)$/.test(path))return;
  fLiveUpd();
}
function fuegoText(){
  const L=S.sectores.map(calcSector),ok=L.filter(c=>isFinite(c.qf));
  const max=ok.length?Math.max(...ok.map(c=>c.qf)):NaN,Fs=[...new Set(ok.map(c=>c.F))],As=[...new Set(ok.map(c=>c.extA))];
  const riesgos=[...new Set(S.sectores.map(s=>s.riesgo))];const vent=S.meta.ventilacion==="mecanica"?"mecánica (Cuadro 2.2.2)":"natural (Cuadro 2.2.1)";
  const t=[];
  t.push(`Se relevaron ${S.sectores.length} sectores de incendio en el establecimiento ${S.meta.razon}, con actividad predominante ${S.meta.actividad.toLowerCase()}. El riesgo determinado según la Tabla 2.1 del Anexo VII del Decreto 351/79 es ${riesgos.map(r=>`Riesgo ${r} – ${RIESGOS[r-1][1]}`).join(", ")}.`);
  if(ok.length)t.push(`La carga de fuego calculada va de ${fmt(Math.min(...ok.map(c=>c.qf)),1)} a ${fmt(max,1)} kg/m² (sector de mayor carga: ${S.sectores[L.indexOf(ok.find(c=>c.qf===max))].nombre}). Con ventilación ${vent}, la resistencia al fuego exigible es ${Fs.join(" / ")}.`);
  t.push(`Según la Tabla 1 del Anexo VII corresponde un potencial extintor mínimo ${As.join(" / ")} para fuegos clase A${L.some(c=>c.extB&&c.extB!=="—")?`, y ${[...new Set(L.map(c=>c.extB).filter(x=>x&&x!=="—"))].join(" / ")} para fuegos clase B`:""}, con al menos un extintor cada 200 m² por planta y una distancia máxima a recorrer de 20 m (clase A) y 15 m (clase B).`);
  const np=S.sectores.filter((s,i)=>L[i].np||L[i].F==="NP");
  if(np.length)t.push(`Los sectores ${np.map(s=>s.nombre).join(", ")} presentan una combinación de riesgo y actividad o ventilación no permitida (NP) por el Anexo VII y requieren adecuación.`);
  S.conclusiones=t.join("\n\n");
  const r=["Mantener los extintores con carga y prueba hidráulica vigentes, señalizados, libres de obstrucciones y a una altura no mayor a 1,50 m.","Mantener actualizado el estudio de carga de fuego ante cambios de uso, mobiliario o almacenamiento de materiales combustibles."];
  if(max>30)r.unshift("Evaluar la reducción de la carga de fuego (orden, retiro de material combustible en desuso, almacenamiento en sectores separados).");
  S.recomendaciones=r.map(x=>"• "+x).join("\n");
}
function exportFuego(){
  const m=S.meta,a=XB.mk(),rows=[];let R=1;const hpt=(r,h)=>{rows[r-1]={hpt:h}};
  XB.put(a,"A1","ESTUDIO DE CARGA DE FUEGO",{font:XB.F(12,true),alignment:C_});XB.mg(a,"A1:D1");hpt(1,24);
  XB.put(a,"A2",`${m.razon} – ${m.direccion}, ${m.localidad} – CUIT ${m.cuit}`,{alignment:C_});XB.mg(a,"A2:D2");
  XB.put(a,"A3",`Decreto 351/79 – Anexo VII, Cap. 18 · Σ Pi×Ki / (Sup × 4400) = carga de fuego · Fecha: ${fdate(m.fechaMed)}`,{font:XB.F(9),alignment:C_});XB.mg(a,"A3:D3");
  R=5;let planta=null;
  S.sectores.forEach(s=>{const c=calcSector(s);
    if(s.planta!==planta){planta=s.planta;XB.put(a,"A"+R,"– "+String(planta||"").toUpperCase(),{font:XB.F(10,true)});XB.mg(a,`A${R}:D${R}`);R+=2}
    const g="D9D9D9";
    XB.put(a,"A"+R,"Sector",{font:XB.F(10,true),alignment:C_});XB.put(a,"B"+R,s.nombre,{font:XB.F(10,true),alignment:C_});XB.put(a,"C"+R,"Superficie del Sector (m²)",{font:XB.F(10,true),alignment:C_});XB.put(a,"D"+R,c.sup,{alignment:C_,numFmt:"0.00"});XB.fill(a,`A${R}:D${R}`,g);
    XB.put(a,"A"+(R+1),"Riesgo del Sector",{font:XB.F(10,true),alignment:C_});XB.put(a,"B"+(R+1),`Riesgo ${s.riesgo} – ${RIESGOS[s.riesgo-1][1]}`,{alignment:C_});XB.put(a,"C"+(R+1),"Clase de Fuego",{font:XB.F(10,true),alignment:C_});XB.put(a,"D"+(R+1),s.clase,{alignment:C_});
    ["Descripción material","Cantidad (kg)","Poder Calorífico (Kcal/kg)","Calor Generado (Kcal)"].forEach((h,k)=>XB.put(a,"ABCD"[k]+(R+2),h,{font:XB.F(10,true),alignment:C_}));XB.fill(a,`A${R+2}:D${R+2}`,"F2F2F2");
    let r=R+3;s.materiales.forEach((mt,j)=>{const x=c.rows[j];XB.put(a,"A"+r,String(mt.desc||"").toUpperCase(),{alignment:L_});XB.put(a,"B"+r,isFinite(x.kg)?x.kg:"",{alignment:C_});XB.put(a,"C"+r,isFinite(x.pc)?x.pc:"",{alignment:C_});XB.put(a,"D"+r,isFinite(x.q)?x.q:"",{alignment:C_,numFmt:"#,##0"});r++});
    [["TOTAL CALOR GENERADO",c.Q,"#,##0"],["Kg Madera Equivalente",Math.round(c.kgm),"0"],["Carga de Fuego (kg/m²)",isFinite(c.qf)?Math.round(c.qf*10)/10:"","0.0"],["Resistencia al fuego exigible",c.F||"",null],["Potencial extintor mínimo",`${c.extA||""}${c.extB&&c.extB!=="—"?" / "+c.extB:""}`,null]].forEach(([t,v,f])=>{XB.put(a,"C"+r,t,{font:XB.F(10,true),alignment:L_});XB.put(a,"D"+r,v,{font:XB.F(10,true),alignment:C_,...(f?{numFmt:f}:{})});r++});
    XB.grid(a,`A${R}:D${r-1}`);R=r+2});
  const ws1=XB.build(a,[30,16,30,22],rows,{left:0.6,right:0.4,top:0.6,bottom:0.6,header:0.3,footer:0.3});
  // Resumen
  const b=XB.mk(),r2=[];XB.put(b,"A1","RESUMEN – CARGA DE FUEGO, RESISTENCIA AL FUEGO Y EXTINTORES",{font:XB.F(12,true),alignment:C_});XB.mg(b,"A1:H1");r2[0]={hpt:24};
  XB.put(b,"A2",`${m.razon} · Actividad: ${m.actividad} · Ventilación: ${m.ventilacion==="mecanica"?"mecánica (Cuadro 2.2.2)":"natural (Cuadro 2.2.1)"}`,{alignment:C_});XB.mg(b,"A2:H2");
  ["Planta","Sector","Superficie (m²)","Clase de fuego","Riesgo","Carga de fuego (kg/m²)","Resistencia al fuego","Potencial extintor"].forEach((h,k)=>XB.put(b,"ABCDEFGH"[k]+"4",h,{font:XB.F(10,true),alignment:C_}));XB.fill(b,"A4:H4","D9D9D9");r2[3]={hpt:30};
  let rr=5;S.sectores.forEach(s=>{const c=calcSector(s);[s.planta,s.nombre,isFinite(c.sup)?c.sup:"",s.clase,`R${s.riesgo}`,isFinite(c.qf)?Math.round(c.qf*10)/10:"",c.F||"",`${c.extA||""}${c.extB&&c.extB!=="—"?" / "+c.extB:""}`].forEach((v,k)=>XB.put(b,"ABCDEFGH"[k]+rr,v,{alignment:C_}));if(c.np||c.F==="NP")XB.fill(b,`A${rr}:H${rr}`,"F4CCCC");rr++});
  XB.grid(b,`A4:H${rr-1}`);rr++;
  XB.put(b,"A"+rr,"Conclusión",{font:XB.F(10,true)});XB.mg(b,`A${rr}:H${rr}`);XB.fill(b,`A${rr}:H${rr}`,"F2F2F2");rr++;
  XB.put(b,"A"+rr,S.conclusiones,{alignment:{horizontal:"left",vertical:"top",wrapText:true}});XB.mg(b,`A${rr}:H${rr}`);r2[rr-1]={hpt:XB.lines(S.conclusiones,110)*13+8};rr++;
  XB.put(b,"A"+rr,"Recomendaciones",{font:XB.F(10,true)});XB.mg(b,`A${rr}:H${rr}`);XB.fill(b,`A${rr}:H${rr}`,"F2F2F2");rr++;
  XB.put(b,"A"+rr,S.recomendaciones,{alignment:{horizontal:"left",vertical:"top",wrapText:true}});XB.mg(b,`A${rr}:H${rr}`);r2[rr-1]={hpt:XB.lines(S.recomendaciones,110)*13+8};XB.box(b,`A${rr-3}:H${rr}`);rr+=2;
  XB.put(b,"E"+rr,`${m.profesional} – ${m.matricula}`,{font:XB.F(9),alignment:{horizontal:"right"}});XB.mg(b,`E${rr}:H${rr}`);
  const ws2=XB.build(b,[14,22,12,10,9,14,13,15],r2);
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws1,"Cálculo por sector");XLSX.utils.book_append_sheet(wb,ws2,"Resumen");
  saveWorkbook(wb,XB.fname("Carga_de_Fuego"),["p","l"]);
}
function repeatFuego(src){const d=JSON.parse(JSON.stringify(src));const old=S;S=d;d.sectores.forEach(s=>{const c=calcSector(s);s.prev={qf:c.qf};s.id=uid()});S=old;
  Object.assign(d.meta,{fechaMed:today()});d.conclusiones="";d.recomendaciones="";d.updatedAt=Date.now();return d}

/* =====================================================================
   PUESTA A TIERRA – Res. SRT 900/15
   ===================================================================== */
const PAT_TERRENO=["Lecho seco","Lecho húmedo","Arcilloso","Pantanoso","Lluvias recientes","Arenoso seco","Arenoso húmedo","Otro"];
const PAT_USO=["Toma de tierra de seguridad de masas","Toma de tierra del neutro de transformador","Protección de equipos electrónicos","Protección de equipos informáticos","Iluminación","Pararrayos","Otro"];
const PAT_ESQ=["TT","TN-S","TN-C","TN-C-S","IT"];
const METODO_PAT="La medición se llevó a cabo siguiendo los lineamientos de la Norma IRAM 2281 Parte II: desde el punto de toma de tierra o electrodo dispersor existente se separan los electrodos auxiliares del instrumento a 5 m y 10 m en línea recta (método de las tres picas). Se complementa con medición de lazo de impedancia en circuito energizado y verificación de continuidad de las masas.";
function newToma(n,o={}){return Object.assign({id:uid(),n,sector:"",terreno:"Lecho húmedo",uso:"Toma de tierra de seguridad de masas",esquema:"TT",valor:"",continuo:"SI",capacidad:"SI",proteccion:"DD",desconecta:"SI",obs:"",prev:null},o)}
function newDisy(o={}){return Object.assign({id:uid(),tablero:"",circuito:"",marca:"",polos:"Bipolar",In:"",IDn:"30",medio:"No dispara",t1:"",t5:"",corr:""},o)}
function newPat(){const m=blankMeta();m.instrumento="Telurímetro";m.marca="";m.modelo="";m.serie="";m.fechaCal="";m.metodologia=METODO_PAT;m.limite="40";m.probador="";
  return{tipo:"pat",updatedAt:Date.now(),meta:m,tomas:[],disy:[],conclusiones:"",recomendaciones:""}}
function calcToma(t){const v=num(t.valor),lim=num(S&&S.meta.limite)||40;const o={v,lim};
  o.valorOK=isFinite(v)?v<=lim:null;o.cumple=o.valorOK===null?null:(o.valorOK&&t.continuo==="SI"&&t.capacidad==="SI"&&t.desconecta==="SI");return o}
function calcDisy(d){const t1=num(d.t1),t5=num(d.t5),c=num(d.corr),I=num(d.IDn)||30;
  const o={medioOK:d.medio==="No dispara",t1OK:isFinite(t1)?t1<=300:null,t5OK:isFinite(t5)?t5<=40:null,corrOK:isFinite(c)?(c>I/2&&c<=I):null};
  const any=isFinite(t1)||isFinite(t5)||isFinite(c);o.pasa=any?(o.medioOK&&o.t1OK!==false&&o.t5OK!==false&&o.corrOK!==false):null;return o}
function patSummary(d){const old=S;S=d;const L=d.tomas.map(calcToma),D=d.disy.map(calcDisy);S=old;
  return{tot:d.tomas.length,done:L.filter(c=>c.cumple!==null).length,bad:L.filter(c=>c.cumple===false).length+D.filter(c=>c.pasa===false).length,warn:0,label:"Puesta a tierra"}}
function patChecks(){const o=baseChecks([...EST_KEYS,["marca","Marca del instrumento"],["modelo","Modelo del instrumento"],["serie","N° de serie del instrumento"],["fechaCal","Fecha de calibración"],["horaIni","Hora de inicio"],["horaFin","Hora de finalización"],["metodologia","Metodología"]]);
  if(!S.tomas.length)o.push("Al menos una toma de tierra");
  S.tomas.forEach(t=>{const mi=[];if(!t.sector)mi.push("sector");if(!isFinite(num(t.valor)))mi.push("valor medido");if(mi.length)o.push(`Toma ${t.n}: ${mi.join(", ")}`)});
  if(!S.conclusiones.trim())o.push("Conclusiones (40)");if(!S.recomendaciones.trim())o.push("Recomendaciones (41)");return o}
let pEdit=null,pSub="tomas";
function vPat(tab){
  if(tab==="est")return`
   <div class="card"><h2>Medición de puesta a tierra</h2><div class="grid">
    ${dfld("meta.fechaMed","(9) Fecha de la medición")}${tfld("meta.horaIni","(10) Hora de inicio")}${tfld("meta.horaFin","(11) Hora de finalización")}
    ${fld("meta.limite","Valor máximo de PAT admitido (Ω)",{im:"decimal"})}</div>
    <p class="muted">40 Ω es el máximo para diferenciales de hasta 30 mA (AEA 90364-7-771, Tabla 771.3.I, referida por la Res. SRT 900/15).</p></div>
   <div class="card"><h2>(7–8) Instrumento</h2><div class="grid">
    ${fld("meta.instrumento","Instrumento")}${fld("meta.marca","Marca",{req:1})}${fld("meta.modelo","Modelo",{req:1})}${fld("meta.serie","N° de serie",{req:1})}${dfld("meta.fechaCal","Fecha de calibración")}
    ${fld("meta.probador","Probador de diferenciales (marca, modelo, serie)",{full:1})}</div><div id="calBox">${calVigencia()}</div></div>
   ${estBlock(!S.meta.razon)}
   <details class="more"><summary>(12–15) Metodología, observaciones y documentación</summary><div class="in">
    ${fld("meta.metodologia","(12) Metodología utilizada",{type:"textarea",rows:5})}<div style="height:8px"></div>${fld("meta.observaciones","(13) Observaciones",{type:"textarea",rows:2})}
    <label class="chk"><input type="checkbox" data-b="meta.docCert" ${S.meta.docCert?"checked":""}> (14) Certificado de calibración</label>
    <label class="chk"><input type="checkbox" data-b="meta.docCroquis" ${S.meta.docCroquis?"checked":""}> (15) Plano o croquis con la ubicación de las tomas</label></div></details>
   ${profBlock()}`;
  if(tab==="pts"){
    if(pEdit!=null){const arr=pSub==="tomas"?S.tomas:S.disy;if(arr[pEdit])return pSub==="tomas"?vToma(pEdit):vDisy(pEdit);pEdit=null}
    const L=S.tomas.map(calcToma),D=S.disy.map(calcDisy);
    return`<div class="seg" style="margin-bottom:12px"><button data-psub="tomas" class="${pSub==="tomas"?"on":""}">Tomas de tierra (${S.tomas.length})</button><button data-psub="disy" class="${pSub==="disy"?"on":""}">Disyuntores (${S.disy.length})</button></div>
     ${pSub==="tomas"?`${S.tomas.length?`<div class="plist">${S.tomas.map((t,i)=>{const c=L[i];return`<button class="pitem" data-pe="${i}"><span class="num">${t.n}</span><span style="min-width:0"><span class="t" style="display:block">${esc(t.sector||"(sin sector)")}</span><span class="s" style="display:block">${esc(t.esquema)} · ${esc(t.uso)}</span></span><span class="v"><b>${isFinite(c.v)?fmt(c.v,2)+" Ω":"–"}</b><span class="pill ${c.cumple===null?"na":c.cumple?"ok":"bad"}">${c.cumple===null?"Sin medir":c.cumple?"Cumple":"No cumple"}</span></span></button>`}).join("")}</div>`:`<div class="card empty"><b>Sin tomas cargadas</b>Agregá cada toma de tierra con el número que figura en el croquis.</div>`}
       <button class="btn p wide" id="addT" style="margin-top:12px;min-height:52px">＋ Agregar toma de tierra</button>`
     :`${S.disy.length?`<div class="plist">${S.disy.map((d,i)=>{const c=D[i];return`<button class="pitem" data-pe="${i}"><span class="num">${i+1}</span><span style="min-width:0"><span class="t" style="display:block">${esc(d.tablero||"(sin tablero)")}</span><span class="s" style="display:block">${esc(d.circuito)} · ${esc(d.polos)} ${esc(d.In)}A · ${esc(d.IDn)} mA</span></span><span class="v"><b>${isFinite(num(d.t5))?fmt(num(d.t5),1)+" ms":"–"}</b><span class="pill ${c.pasa===null?"na":c.pasa?"ok":"bad"}">${c.pasa===null?"Sin probar":c.pasa?"Pasa":"No pasa"}</span></span></button>`}).join("")}</div>`:`<div class="card empty"><b>Sin disyuntores</b>Cargá la prueba de cada interruptor diferencial (½ IΔn, 1 IΔn, 5 IΔn y escalón).</div>`}
       <button class="btn p wide" id="addD" style="margin-top:12px;min-height:52px">＋ Agregar disyuntor</button>`}`}
  if(tab==="con"){const L=S.tomas.map(calcToma);return`<div class="card"><h2>Resumen</h2><div class="tablewrap"><table class="sum"><tr><th>N°</th><th>Sector</th><th>Ω</th><th>Cont.</th><th>Prot.</th><th>Estado</th></tr>
     ${S.tomas.map((t,i)=>{const c=L[i];return`<tr><td>${t.n}</td><td>${esc(t.sector)}</td><td>${isFinite(c.v)?fmt(c.v,2):"–"}</td><td>${t.continuo}</td><td>${t.proteccion}</td><td><span class="pill ${c.cumple===null?"na":c.cumple?"ok":"bad"}">${c.cumple===null?"–":c.cumple?"Cumple":"No cumple"}</span></td></tr>`}).join("")}</table></div></div>
     ${textCard("(40–41) Análisis de las mejoras a realizar","Generar conclusiones automáticamente")}`}
  return expCard(patChecks(),"Descargar protocolo Res. SRT 900/15","Cuatro hojas como el protocolo oficial: (1–15) datos de la medición, (16–33) tomas de tierra y continuidad, prueba de disyuntores, y (34–41) conclusiones y recomendaciones.");
}
function vToma(i){const t=S.tomas[i],p=`tomas.${i}.`;return`<div class="ptop"><button class="navb" id="pBack">☰</button><div class="ttl"><b>Toma ${t.n} · ${esc(t.sector||"nueva")}</b><span>${esc(t.uso)}</span></div>
  <button class="navb" id="pPrev" ${i===0?"disabled":""}>‹</button><button class="navb" id="pNext" ${i===S.tomas.length-1?"disabled":""}>›</button></div>
  <div class="card" id="pLive"></div>
  <div class="card"><div class="grid">${fld(p+"n","(22) N° de toma",{im:"numeric"})}${fld(p+"sector","(23) Sector",{req:1})}</div>
   <div style="margin-top:12px">${fld(p+"valor","(27) Valor medido (Ω)",{req:1,im:"decimal",ph:"ej. 0,65"})}</div></div>
  <div class="card"><div class="grid">${fld(p+"terreno","(24) Condición del terreno",{opts:PAT_TERRENO})}${fld(p+"uso","(25) Uso de la puesta a tierra",{opts:PAT_USO})}</div>
   <div style="display:grid;gap:10px;margin-top:12px">
   ${segF(p+"esquema","(26) Esquema de conexión a tierra",PAT_ESQ,t.esquema)}
   ${segF(p+"continuo","(29) ¿El circuito de PAT es continuo y permanente?",["SI","NO"],t.continuo)}
   ${segF(p+"capacidad","(30) ¿Tiene capacidad de carga para la corriente de falla y resistencia apropiada?",["SI","NO"],t.capacidad)}
   ${segF(p+"proteccion","(31) Protección contra contactos indirectos",["DD","IA","Fus"],t.proteccion)}
   ${segF(p+"desconecta","(32) ¿El dispositivo desconecta en forma automática en los tiempos de la AEA?",["SI","NO"],t.desconecta)}</div>
   <p class="muted">DD: interruptor diferencial · IA: interruptor automático · Fus: fusible.</p></div>
  <div class="card">${fld(p+"obs","Observaciones",{type:"textarea",rows:2})}<div class="row" style="margin-top:10px"><button class="btn sm d" id="delT">Eliminar toma</button></div></div>
  <button class="btn p wide" id="pNext2" style="min-height:52px">${i===S.tomas.length-1?"Listo ✓":"Siguiente toma ›"}</button>`}
function vDisy(i){const d=S.disy[i],p=`disy.${i}.`;return`<div class="ptop"><button class="navb" id="pBack">☰</button><div class="ttl"><b>Disyuntor ${i+1} · ${esc(d.tablero||"nuevo")}</b><span>${esc(d.circuito)}</span></div>
  <button class="navb" id="pPrev" ${i===0?"disabled":""}>‹</button><button class="navb" id="pNext" ${i===S.disy.length-1?"disabled":""}>›</button></div>
  <div class="card" id="pLive"></div>
  <div class="card"><div class="grid">${fld(p+"tablero","Tablero",{req:1})}${fld(p+"circuito","Circuito")}${fld(p+"marca","Marca")}${fld(p+"In","In (A)",{im:"numeric"})}${fld(p+"IDn","IΔn (mA)",{im:"numeric"})}</div>
   <div style="margin-top:10px">${segF(p+"polos","Modelo",["Bipolar","Tetrapolar"],d.polos)}</div></div>
  <div class="card"><h2>Ensayo · semiciclo positivo</h2>
   ${segF(p+"medio","½ IΔn (no debe disparar)",["No dispara","Dispara"],d.medio)}
   <div class="g3" style="margin-top:10px">${fld(p+"t1","1 × IΔn (ms) · ≤ 300",{im:"decimal"})}${fld(p+"t5","5 × IΔn (ms) · ≤ 40",{im:"decimal"})}${fld(p+"corr","Escalón corriente (mA)",{im:"decimal"})}</div>
   <p class="muted">El escalón de corriente debe disparar entre ½ IΔn y IΔn.</p></div>
  <div class="row" style="margin-bottom:12px"><button class="btn sm" id="dupD">Duplicar</button><button class="btn sm d" id="delD">Eliminar</button></div>
  <button class="btn p wide" id="pNext2" style="min-height:52px">${i===S.disy.length-1?"Listo ✓":"Siguiente disyuntor ›"}</button>`}
function pLiveUpd(){const box=$("#pLive");if(!box||pEdit==null)return;
  if(pSub==="tomas"){const t=S.tomas[pEdit],c=calcToma(t);box.innerHTML=`<div class="live"><div><small>Medido</small><b>${isFinite(c.v)?fmt(c.v,2):"–"}</b><small>Ω</small></div><div><small>Máximo</small><b>${fmt(c.lim,0)}</b><small>Ω</small></div><div><small>${t.prev&&isFinite(t.prev.v)?"Anterior":"Estado"}</small><b>${t.prev&&isFinite(t.prev.v)?fmt(t.prev.v,2):(c.cumple===null?"–":c.cumple?"✓":"✕")}</b></div></div>
    <div class="verdict">${c.valorOK===null?`<span class="pill na">Cargá el valor medido</span>`:`<span class="pill ${c.valorOK?"ok":"bad"}">${c.valorOK?"≤":">"} ${fmt(c.lim,0)} Ω</span>`}<span class="pill ${t.continuo==="SI"?"ok":"bad"}">Continuidad ${t.continuo}</span><span class="pill ${t.desconecta==="SI"?"ok":"bad"}">Desconexión ${t.desconecta}</span></div>`}
  else{const d=S.disy[pEdit],c=calcDisy(d);const pl=(ok,t)=>`<span class="pill ${ok===null?"na":ok?"ok":"bad"}">${t}</span>`;
    box.innerHTML=`<div class="verdict" style="margin-top:0">${pl(c.medioOK,"½ IΔn "+(c.medioOK?"no dispara":"dispara"))}${pl(c.t1OK,"1× ≤ 300 ms")}${pl(c.t5OK,"5× ≤ 40 ms")}${pl(c.corrOK,"Escalón")}</div><p style="margin:8px 0 0;font-weight:700">${c.pasa===null?"Sin ensayar":c.pasa?"Pasa ✓":"No pasa ✕"}</p>`}}
function bindPat(tab){
  const on=(id,fn)=>{const el=$("#"+id);if(el)el.onclick=fn};
  $$("[data-psub]").forEach(b=>b.onclick=()=>{pSub=b.dataset.psub;pEdit=null;render()});
  $$("[data-pe]").forEach(b=>b.onclick=()=>{pEdit=+b.dataset.pe;render();scrollTo(0,0)});
  on("addT",()=>{const last=S.tomas.at(-1);S.tomas.push(newToma(Math.max(0,...S.tomas.map(t=>+t.n||0))+1,last?{terreno:last.terreno,uso:last.uso,esquema:last.esquema,proteccion:last.proteccion}:{}));pEdit=S.tomas.length-1;save();render();scrollTo(0,0)});
  on("addD",()=>{const last=S.disy.at(-1);S.disy.push(newDisy(last?{tablero:last.tablero,marca:last.marca,polos:last.polos,IDn:last.IDn}:{}));pEdit=S.disy.length-1;save();render();scrollTo(0,0)});
  const arr=()=>pSub==="tomas"?S.tomas:S.disy;
  on("pBack",()=>{pEdit=null;render()});on("pPrev",()=>{pEdit--;render();scrollTo(0,0)});on("pNext",()=>{pEdit++;render();scrollTo(0,0)});
  on("pNext2",()=>{if(pEdit<arr().length-1)pEdit++;else pEdit=null;render();scrollTo(0,0)});
  on("delT",async()=>{if(!await ask("¿Eliminar esta toma?",{ok:"Eliminar",danger:true}))return;S.tomas.splice(pEdit,1);pEdit=null;save();render()});
  on("delD",async()=>{if(!await ask("¿Eliminar este disyuntor?",{ok:"Eliminar",danger:true}))return;S.disy.splice(pEdit,1);pEdit=null;save();render()});
  on("dupD",()=>{const c=JSON.parse(JSON.stringify(S.disy[pEdit]));c.id=uid();c.t1="";c.t5="";c.corr="";c.circuito="";S.disy.splice(pEdit+1,0,c);pEdit++;save();render();scrollTo(0,0)});
  on("genTxt",async()=>{if((S.conclusiones.trim()||S.recomendaciones.trim())&&!await ask("¿Reemplazar el texto actual?",{ok:"Reemplazar"}))return;patText();save();render()});
  on("dlX",()=>confirmDl(patChecks(),exportPat));
  if(tab==="pts")pLiveUpd();
}
function patChange(path){if(path==="tomas."+pEdit+".n")S.tomas[pEdit].n=parseInt(S.tomas[pEdit].n)||"";pLiveUpd()}
function patText(){
  const L=S.tomas.map(calcToma),D=S.disy.map(calcDisy),lim=num(S.meta.limite)||40,vals=L.filter(c=>isFinite(c.v)).map(c=>c.v);
  const bad=S.tomas.filter((t,i)=>L[i].cumple===false),dbad=S.disy.filter((d,i)=>D[i].pasa===false);const t=[];
  if(vals.length)t.push((`Se midieron ${vals.length} tomas de tierra con valores entre ${fmt(Math.min(...vals),2)} Ω y ${fmt(Math.max(...vals),2)} Ω. ${bad.length?"":`Todos los valores resultan inferiores al máximo de ${fmt(lim,0)} Ω establecido en la Reglamentación para la Ejecución de Instalaciones Eléctricas en Inmuebles de la AEA (90364-7-771, Tabla 771.3.I) y la Norma IRAM 2281 Parte III.`}`).trim());
  if(bad.length)t.push(`No cumplen las tomas ${bad.map(x=>`N°${x.n} (${x.sector})`).join(", ")}: valor de puesta a tierra superior a ${fmt(lim,0)} Ω y/o falta de continuidad o de desconexión automática.`);
  if(S.tomas.length&&S.tomas.every(x=>x.continuo==="SI"))t.push("Se verificó la continuidad del circuito de protección (tomacorrientes, masas y tableros) en conexión con la toma de tierra.");
  if(S.disy.length)t.push(`Se ensayaron ${S.disy.length} interruptores diferenciales${dbad.length?`; no superaron el ensayo: ${dbad.map(d=>`${d.tablero} – ${d.circuito}`).join(", ")}`:", comprobándose su correcto funcionamiento (no disparo a ½ IΔn, disparo a IΔn en menos de 300 ms y a 5 IΔn en menos de 40 ms)"}.`);
  S.conclusiones=t.join("\n\n");
  const r=[];if(bad.length)r.push("Adecuar las tomas que no cumplen (mejorar el electrodo dispersor, agregar jabalinas o tratar el terreno) y verificar la continuidad del conductor de protección hasta cada masa; repetir la medición luego de la adecuación.");
  if(dbad.length)r.push("Reemplazar los interruptores diferenciales que no superaron el ensayo por dispositivos de IΔn ≤ 30 mA.");
  r.push("Accionar mensualmente el botón de prueba de los interruptores diferenciales y repetir la medición de puesta a tierra con frecuencia anual.");
  S.recomendaciones=bad.length||dbad.length?r.map(x=>"• "+x).join("\n"):"No se realizan recomendaciones adicionales.\n• "+r.at(-1);
}
function exportPat(){
  const m=S.meta,hd=(w,cols,firstNum)=>{const n=firstNum;const last="ABCDEFGHIJK"[cols-1];
    XB.put(w,"A1","PROTOCOLO de MEDICIÓN de la PUESTA a TIERRA y CONTINUIDAD de las MASAS",{font:{...XB.F(12,true),color:{rgb:"FFFFFF"}},alignment:C_});XB.fill(w,`A1:${last}1`,"000000");XB.mg(w,`A1:${last}1`);
    return last};
  const B=XB.F(10,true);
  // Hoja 1 (1-15)
  const a=XB.mk(),r1=[];hd(a,4);r1[0]={hpt:30};
  [["(1) Razón Social:",m.razon],["(2) Dirección:",m.direccion],["(3) Localidad:",m.localidad],["(4) Provincia:",m.provincia]].forEach(([l,v],k)=>{const R=3+k;XB.put(a,"A"+R,l,{font:B});XB.put(a,"B"+R,v);XB.mg(a,`B${R}:D${R}`);XB.box(a,`A${R}:D${R}`)});
  XB.put(a,"A7","(5) C.P.A.:",{font:B});XB.put(a,"B7",m.cp);XB.put(a,"C7","(6) C.U.I.T. N°:",{font:B});XB.put(a,"D7",m.cuit);XB.grid(a,"A7:D7");
  XB.put(a,"A9","Datos para la medición",{font:B,alignment:C_});XB.mg(a,"A9:D9");XB.fill(a,"A9:D9","D9D9D9");
  XB.put(a,"A10","(7) Marca, modelo y número de serie del instrumento utilizado:",{font:B});XB.mg(a,"A10:D10");
  XB.put(a,"A11",`${m.instrumento} ${m.marca} ${m.modelo} – N° de serie: ${m.serie}${m.probador?` · Probador de diferenciales: ${m.probador}`:""}`,{alignment:C_});XB.mg(a,"A11:D11");
  XB.put(a,"A12","(8) Fecha de certificación de calibración del instrumento:",{font:B});XB.mg(a,"A12:C12");XB.put(a,"D12",fdate(m.fechaCal),{alignment:C_});
  XB.put(a,"A13","(9) Fecha de la medición",{font:B,alignment:C_});XB.put(a,"B13","",{});XB.mg(a,"A13:B13");XB.put(a,"C13","(10) Hora de inicio",{font:B,alignment:C_});XB.put(a,"D13","(11) Hora finalización",{font:B,alignment:C_});
  XB.put(a,"A14",fdate(m.fechaMed),{alignment:C_});XB.mg(a,"A14:B14");XB.put(a,"C14",m.horaIni?m.horaIni+" hs":"",{alignment:C_});XB.put(a,"D14",m.horaFin?m.horaFin+" hs":"",{alignment:C_});
  XB.put(a,"A15","(12) Metodología utilizada:",{font:B});XB.mg(a,"A15:D15");XB.put(a,"A16",m.metodologia,{alignment:{horizontal:"left",vertical:"top",wrapText:true}});XB.mg(a,"A16:D16");r1[15]={hpt:XB.lines(m.metodologia,85)*13+8};
  XB.grid(a,"A10:D16");XB.box(a,"A9:D16");
  XB.put(a,"A18",`(13) Observaciones: ${m.observaciones||""}`,{font:B,alignment:{horizontal:"left",vertical:"top",wrapText:true}});XB.mg(a,"A18:D18");r1[17]={hpt:Math.max(50,XB.lines(m.observaciones,85)*13+20)};XB.box(a,"A18:D18");
  XB.put(a,"A20","Documentación que se adjuntará a la medición",{font:B,alignment:C_});XB.mg(a,"A20:D20");XB.fill(a,"A20:D20","D9D9D9");
  XB.put(a,"A21",`(14) Certificado de calibración: ${m.docCert?"SI":"NO"}`,{font:B});XB.mg(a,"A21:D21");XB.put(a,"A22",`(15) Plano o croquis: ${m.docCroquis?"SI":"NO"}`,{font:B});XB.mg(a,"A22:D22");XB.grid(a,"A20:D22");
  XB.put(a,"C24","Hoja: 1/4",{alignment:{horizontal:"right"}});XB.mg(a,"C24:D24");XB.put(a,"C26",`${m.profesional} – ${m.matricula}`,{font:XB.F(9),alignment:{horizontal:"right",wrapText:true}});XB.mg(a,"C26:D26");r1[25]={hpt:26};
  const ws1=XB.build(a,[30,18,22,24],r1,{left:0.7,right:0.5,top:0.6,bottom:0.6,header:0.3,footer:0.3});
  // encabezado repetido 16-21
  const hdr2=(w,last,n0)=>{XB.put(w,"A2",`(${n0}) Razón social: ${m.razon}`,{font:B});XB.put(w,"A3",`(${n0+2}) Dirección: ${m.direccion}`,{font:B});
    XB.put(w,"F2",`(${n0+1}) C.U.I.T. N°: ${m.cuit}`,{font:B});XB.put(w,"E3",`(${n0+3}) Localidad: ${m.localidad}`,{font:B});XB.put(w,"G3",`(${n0+4}) C.P.A.: ${m.cp}`,{font:B});XB.put(w,"I3",`(${n0+5}) Provincia: ${m.provincia}`,{font:B});
    XB.mg(w,"A2:E2");XB.mg(w,`F2:${last}2`);XB.mg(w,"A3:D3");XB.mg(w,"E3:F3");XB.mg(w,"G3:H3");XB.mg(w,`I3:${last}3`);XB.grid(w,`A2:${last}3`)};
  // Hoja 2 (16-33)
  const b=XB.mk(),r2=[];hd(b,11);hdr2(b,"K",16);r2[0]={hpt:26};
  XB.put(b,"A5","Datos de la medición",{font:B,alignment:C_});XB.mg(b,"A5:K5");XB.fill(b,"A5:K5","D9D9D9");
  XB.put(b,"G6","Medición de la P.A.T.",{font:B,alignment:C_});XB.mg(b,"G6:H6");XB.put(b,"I6","Continuidad de las masas",{font:B,alignment:C_});XB.mg(b,"I6:K6");
  const H2=[["22","N° de toma de tierra"],["23","Sector"],["24","Descripción de la condición del terreno al momento de la medición"],["25","Uso de la puesta a tierra"],["26","Esquema de conexión a tierra utilizado (TT/TN-S/TN-C/TN-C-S/IT)"],["27","Valor obtenido en la medición en ohmios (Ω)"],["28","Cumple con la reglamentación AEA (SI/NO)"],["29","El circuito de P.A.T. es continuo y permanente (SI/NO)"],["30","El circuito de P.A.T. tiene capacidad de carga para conducir la corriente de falla y una resistencia apropiada (SI/NO)"],["31","Protección contra contactos indirectos: DD / IA / Fus"],["32","El dispositivo desconecta en forma automática la alimentación para lograr la protección contra los contactos indirectos (SI/NO)"]];
  const cl="ABCDEFGHIJK";
  H2.forEach(([n,t],k)=>{XB.put(b,cl[k]+"7",n,{font:XB.F(8,true),alignment:C_});XB.put(b,cl[k]+"8",t,{font:XB.F(8),alignment:C_})});r2[7]={hpt:110};
  let R=9;S.tomas.forEach(t=>{const c=calcToma(t);[+t.n||t.n,t.sector,t.terreno,t.uso,t.esquema,isFinite(c.v)?c.v:"",c.valorOK===null?"":c.valorOK?"SI":"NO",t.continuo,t.capacidad,t.proteccion,t.desconecta].forEach((v,k)=>XB.put(b,cl[k]+R,v,{font:XB.F(9,k<2),alignment:C_,...(k===5?{numFmt:"0.00"}:{})}));if(c.cumple===false)XB.fill(b,`A${R}:K${R}`,"F4CCCC");r2[R-1]={hpt:30};R++});
  XB.grid(b,`A6:K${R-1}`);
  XB.put(b,"A"+R,`(33) Información adicional: ${S.tomas.map(t=>t.obs?`Toma ${t.n}: ${t.obs}`:"").filter(Boolean).join(" · ")||"ver anexo conclusiones"}`,{font:B,alignment:{horizontal:"left",vertical:"top",wrapText:true}});XB.mg(b,`A${R}:K${R}`);r2[R-1]={hpt:40};XB.box(b,`A${R}:K${R}`);
  XB.put(b,"I"+(R+2),"Hoja: 2/4",{alignment:{horizontal:"right"}});XB.mg(b,`I${R+2}:K${R+2}`);
  const ws2=XB.build(b,[7,16,13,15,11,10,9,10,13,9,14],r2,{left:0.4,right:0.4,top:0.5,bottom:0.5,header:0.3,footer:0.3});
  // Hoja 3 disyuntores
  const c3=XB.mk(),r3=[];XB.put(c3,"A1","PROTOCOLO PRUEBA DE DISYUNTORES – ANEXO RES. SRT N° 900/15",{font:{...XB.F(12,true),color:{rgb:"FFFFFF"}},alignment:C_});XB.fill(c3,"A1:K1","000000");XB.mg(c3,"A1:K1");r3[0]={hpt:26};hdr2(c3,"K",16);
  XB.put(c3,"A5","Datos de los ensayos",{font:B,alignment:C_});XB.mg(c3,"A5:K5");XB.fill(c3,"A5:K5","D9D9D9");
  XB.put(c3,"C6","Condición",{font:B,alignment:C_});XB.put(c3,"D6","Interruptor diferencial",{font:B,alignment:C_});XB.mg(c3,"D6:G6");XB.put(c3,"H6","Semiciclo positivo",{font:B,alignment:C_});XB.mg(c3,"H6:K6");
  [["Tablero",""],["Circuito",""],["Pasa / No pasa",""],["Marca",""],["Modelo","Bi/tetrapolar"],["In","A"],["IΔn","mA"],["½ x","No debe disparar"],["1 x","< 300 ms"],["5 x","< 40 ms"],["Escalón corriente de prueba","mA"]].forEach(([t,u],k)=>{XB.put(c3,cl[k]+"7",t,{font:XB.F(9,true),alignment:C_});XB.put(c3,cl[k]+"8",u,{font:XB.F(8),alignment:C_})});r3[6]={hpt:30};
  R=9;S.disy.forEach(d=>{const c=calcDisy(d);[d.tablero,d.circuito,c.pasa===null?"":c.pasa?"Pasa":"No pasa",d.marca,d.polos,num(d.In)||d.In,num(d.IDn)||d.IDn,d.medio==="No dispara"?"Pasa":"Dispara",isFinite(num(d.t1))?num(d.t1):"",isFinite(num(d.t5))?num(d.t5):"",isFinite(num(d.corr))?num(d.corr):""].forEach((v,k)=>XB.put(c3,cl[k]+R,v,{font:XB.F(9),alignment:C_}));if(c.pasa===false)XB.fill(c3,`A${R}:K${R}`,"F4CCCC");r3[R-1]={hpt:28};R++});
  XB.grid(c3,`A6:K${Math.max(R-1,8)}`);
  const info=`Información adicional: ${m.probador?`Para la medición de valores se utilizó ${m.probador}. `:""}Se verificaron las características de los diferenciales tomando valores para 0,5 – 1 y 5 veces la corriente diferencial (IΔn) para analizar la velocidad de respuesta, y la prueba escalón de corriente para verificar el valor de la corriente diferencial de actuación.`;
  XB.put(c3,"A"+R,info,{font:XB.F(9),alignment:{horizontal:"left",vertical:"top",wrapText:true}});XB.mg(c3,`A${R}:K${R}`);r3[R-1]={hpt:XB.lines(info,140)*12+10};XB.box(c3,`A${R}:K${R}`);
  XB.put(c3,"I"+(R+2),"Hoja: 3/4",{alignment:{horizontal:"right"}});XB.mg(c3,`I${R+2}:K${R+2}`);
  const ws3=XB.build(c3,[14,16,10,10,12,7,7,11,9,9,12],r3,{left:0.4,right:0.4,top:0.5,bottom:0.5,header:0.3,footer:0.3});
  // Hoja 4 (34-41)
  const d4=XB.mk(),r4=[];hd(d4,11);hdr2(d4,"K",34);r4[0]={hpt:26};
  XB.put(d4,"A5","Análisis de las mejoras a realizar",{font:B,alignment:C_});XB.mg(d4,"A5:K5");XB.fill(d4,"A5:K5","D9D9D9");
  XB.put(d4,"A6","(40) Conclusiones",{font:B,alignment:C_});XB.mg(d4,"A6:E6");XB.put(d4,"F6","(41) Recomendaciones para la adecuación a la legislación vigente",{font:B,alignment:C_});XB.mg(d4,"F6:K6");
  XB.put(d4,"A7",S.conclusiones,{alignment:{horizontal:"left",vertical:"top",wrapText:true}});XB.mg(d4,"A7:E7");XB.put(d4,"F7",S.recomendaciones,{alignment:{horizontal:"left",vertical:"top",wrapText:true}});XB.mg(d4,"F7:K7");
  r4[6]={hpt:Math.max(200,Math.max(XB.lines(S.conclusiones,60),XB.lines(S.recomendaciones,60))*13+12)};XB.grid(d4,"A6:K6");XB.box(d4,"A7:E7");XB.box(d4,"F7:K7");
  XB.put(d4,"I9","Hoja: 4/4",{alignment:{horizontal:"right"}});XB.mg(d4,"I9:K9");XB.put(d4,"F11",`${m.profesional} – ${m.matricula}`,{font:XB.F(9),alignment:{horizontal:"right"}});XB.mg(d4,"F11:K11");
  const ws4=XB.build(d4,[7,16,13,15,11,10,9,10,13,9,14],r4,{left:0.4,right:0.4,top:0.5,bottom:0.5,header:0.3,footer:0.3});
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws1,"Hoja 1");XLSX.utils.book_append_sheet(wb,ws2,"Hoja 2 - Tomas");XLSX.utils.book_append_sheet(wb,ws3,"Hoja 3 - Disyuntores");XLSX.utils.book_append_sheet(wb,ws4,"Hoja 4 - Conclusiones");
  saveWorkbook(wb,XB.fname("Puesta_a_Tierra_900-15"),["p","l","l","l"]);
}
function repeatPat(src){const d=JSON.parse(JSON.stringify(src));d.tomas.forEach(t=>{t.prev={v:num(t.valor)};t.valor="";t.id=uid()});d.disy.forEach(x=>{x.t1="";x.t5="";x.corr="";x.medio="No dispara";x.id=uid()});
  Object.assign(d.meta,{fechaMed:today(),horaIni:"",horaFin:"",observaciones:""});d.conclusiones="";d.recomendaciones="";d.updatedAt=Date.now();return d}

/* ---------- registro de tipos ---------- */
const MODS={
  fuego:{nombre:"Carga de fuego",norma:"Dec. 351/79 Anexo VII",tabs:["Datos","Sectores","Análisis","Exportar"],nuevo:newFuego,view:vFuego,bind:bindFuego,change:fuegoChange,summary:fuegoSummary,repeat:repeatFuego,reset:()=>{fEdit=null}},
  pat:{nombre:"Puesta a tierra",norma:"Res. SRT 900/15",tabs:["Datos","Medir","Análisis","Exportar"],nuevo:newPat,view:vPat,bind:bindPat,change:patChange,summary:patSummary,repeat:repeatPat,reset:()=>{pEdit=null;pSub="tomas"}}
};
