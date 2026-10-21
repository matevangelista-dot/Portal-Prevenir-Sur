"use strict";
/* =====================================================================
   RUIDO (Res. SRT 85/12 – Res. MTEySS 295/03 Anexo V)
   VIBRACIONES mano-brazo y cuerpo entero (Res. MTEySS 295/03 Anexo V)
   TERMOHIGROMETRÍA Y ESTRÉS TÉRMICO (calor: Res. SRT 30/23 · frío: Res. MTEySS 295/03)
   Se apoya en los helpers de app.js y estudios.js (XB, estBlock, profBlock, ...).
   ===================================================================== */
const log10=Math.log10,sq=x=>x*x;
const nOrNull=v=>{const x=num(v);return isFinite(x)?x:null};

/* ---------- motor común: lista de puntos + editor ---------- */
function makeModule(cfg){
  const st={e:null};
  const arr=()=>S[cfg.arr];
  function list(){
    const A=arr();
    return`${A.length?`<div class="plist">${A.map((it,i)=>{const c=cfg.calc(it),s=cfg.status(c,it);return`
     <button class="pitem" data-ge="${i}"><span class="num">${esc(it.n)}</span><span style="min-width:0"><span class="t" style="display:block">${esc(cfg.title(it))}</span>
     <span class="s" style="display:block">${esc(cfg.sub(it))}</span></span><span class="v"><b>${cfg.value(it,c)}</b><span class="pill ${s.cls}">${s.txt}</span></span></button>`}).join("")}</div>`
     :`<div class="card empty"><b>${cfg.emptyT}</b>${cfg.emptyD}</div>`}
     <button class="btn p wide" id="gAdd" style="margin-top:12px;min-height:52px">＋ ${cfg.addLabel}</button>`}
  function editor(i){const it=arr()[i];return`<div class="ptop"><button class="navb" id="gBack" aria-label="Volver">☰</button>
    <div class="ttl"><b>${esc(it.n)} · ${esc(cfg.title(it))}</b><span>${esc(cfg.sub(it))}</span></div>
    <button class="navb" id="gPrev" ${i===0?"disabled":""}>‹</button><button class="navb" id="gNext" ${i===arr().length-1?"disabled":""}>›</button></div>
    <div class="card" id="gLive"></div>${cfg.fields(i,it)}
    <div class="row" style="margin-bottom:12px"><button class="btn sm" id="gDup">Duplicar</button><button class="btn sm d" id="gDel">Eliminar</button></div>
    <button class="btn p wide" id="gNext2" style="min-height:52px">${i===arr().length-1?"Listo ✓":"Siguiente ›"}</button>`}
  function live(){const box=$("#gLive");if(!box||st.e==null)return;const it=arr()[st.e];if(!it)return;box.innerHTML=cfg.live(it,cfg.calc(it))}
  function view(tab){
    if(tab==="est")return cfg.est();
    if(tab==="pts"){if(st.e!=null&&arr()[st.e])return editor(st.e);st.e=null;return(cfg.listHead?cfg.listHead():"")+list()}
    if(tab==="con")return cfg.summaryTable()+textCard("Análisis de los datos y mejoras a realizar","Generar conclusiones automáticamente");
    return expCard(cfg.checks(),cfg.dlLabel,cfg.dlDesc);
  }
  function bind(tab){
    const on=(id,fn)=>{const el=$("#"+id);if(el)el.onclick=fn};
    $$("[data-ge]").forEach(b=>b.onclick=()=>{st.e=+b.dataset.ge;render();scrollTo(0,0)});
    on("gAdd",()=>{const A=arr();A.push(cfg.newItem(Math.max(0,...A.map(x=>+x.n||0))+1,A.at(-1)));st.e=A.length-1;save();render();scrollTo(0,0)});
    on("gBack",()=>{st.e=null;render()});on("gPrev",()=>{st.e--;render();scrollTo(0,0)});on("gNext",()=>{st.e++;render();scrollTo(0,0)});
    on("gNext2",()=>{if(st.e<arr().length-1)st.e++;else st.e=null;render();scrollTo(0,0)});
    on("gDup",()=>{const A=arr(),c=JSON.parse(JSON.stringify(A[st.e]));c.id=uid();c.n=Math.max(0,...A.map(x=>+x.n||0))+1;c.prev=null;if(cfg.clearOnDup)cfg.clearOnDup(c);A.splice(st.e+1,0,c);st.e++;save();render();scrollTo(0,0)});
    on("gDel",async()=>{if(!await ask("¿Eliminar este registro?",{ok:"Eliminar",danger:true}))return;arr().splice(st.e,1);st.e=null;save();render()});
    on("genTxt",async()=>{if((S.conclusiones.trim()||S.recomendaciones.trim())&&!await ask("¿Reemplazar el texto actual?",{ok:"Reemplazar"}))return;cfg.autoText();save();render()});
    on("dlX",()=>confirmDl(cfg.checks(),cfg.exportFn));
    if(cfg.bindExtra)cfg.bindExtra(st);
    if(tab==="pts")live();
  }
  function change(path){if(cfg.onChange&&cfg.onChange(path,st))return;live()}
  function summary(d){const old=S;S=d;const L=d[cfg.arr].map(it=>({c:cfg.calc(it),it}));S=old;
    const ss=L.map(x=>cfg.status(x.c,x.it));return{tot:L.length,done:ss.filter(s=>s.cls!=="na").length,bad:ss.filter(s=>s.cls==="bad").length,warn:ss.filter(s=>s.cls==="warn").length}}
  function repeat(src){const d=JSON.parse(JSON.stringify(src));const old=S;S=d;d[cfg.arr].forEach(it=>{it.prev=cfg.prevOf(it,cfg.calc(it));cfg.clearOnRepeat(it);it.id=uid()});S=old;
    Object.assign(d.meta,{fechaMed:today(),horaIni:"",horaFin:"",observaciones:""});d.conclusiones="";d.recomendaciones="";d.updatedAt=Date.now();return d}
  return{nombre:cfg.nombre,norma:cfg.norma,unidad:cfg.unidad,tabs:cfg.tabs,nuevo:cfg.nuevo,view,bind,change,summary,repeat,reset:()=>{st.e=null},_st:st};
}
/* Excel genérico: Hoja 1 datos, Hoja 2 tabla, Hoja 3 conclusiones */
function exportStd({titulo,norma,datos,headers,rows,bad,widths,info,fname,numFirst}){
  const m=S.meta,B=XB.F(10,true),n=headers.length,last=XLSX.utils.encode_col(n-1);
  const head=(w,t)=>{XB.put(w,"A1",t,{font:{...XB.F(12,true),color:{rgb:"FFFFFF"}},alignment:C_});XB.fill(w,`A1:${last}1`,"000000");XB.mg(w,`A1:${last}1`)};
  const idf=(w,r,nums)=>{const h=Math.max(1,Math.floor(n/2));const c2=XLSX.utils.encode_col(h);
    XB.put(w,"A"+r,`${nums?`(${nums[0]}) `:""}Razón social: ${m.razon}`,{font:B});XB.mg(w,`A${r}:${XLSX.utils.encode_col(h-1)}${r}`);XB.put(w,c2+r,`${nums?`(${nums[1]}) `:""}C.U.I.T.: ${m.cuit}`,{font:B});XB.mg(w,`${c2}${r}:${last}${r}`);
    XB.put(w,"A"+(r+1),`${nums?`(${nums[2]}) `:""}Dirección: ${m.direccion} · ${nums?`(${nums[3]}) `:""}Localidad: ${m.localidad} · ${nums?`(${nums[4]}) `:""}C.P.: ${m.cp} · ${nums?`(${nums[5]}) `:""}Provincia: ${m.provincia}`,{font:B});XB.mg(w,`A${r+1}:${last}${r+1}`);XB.grid(w,`A${r}:${last}${r+1}`)};
  // Hoja 1
  const a=XB.mk(),r1=[];XB.put(a,"A1",titulo,{font:{...XB.F(12,true),color:{rgb:"FFFFFF"}},alignment:C_});XB.fill(a,"A1:D1","000000");XB.mg(a,"A1:D1");r1[0]={hpt:30};
  XB.put(a,"A2",norma,{font:XB.F(9),alignment:C_});XB.mg(a,"A2:D2");
  let R=4;datos.forEach(([l,v])=>{if(l===null){XB.put(a,"A"+R,v,{font:B,alignment:C_});XB.mg(a,`A${R}:D${R}`);XB.fill(a,`A${R}:D${R}`,"D9D9D9");XB.box(a,`A${R}:D${R}`);R++;return}
    XB.put(a,"A"+R,l,{font:B});XB.put(a,"B"+R,v,{alignment:{horizontal:"left",vertical:"top",wrapText:true}});XB.mg(a,`B${R}:D${R}`);XB.grid(a,`A${R}:D${R}`);const L=XB.lines(v,60);if(L>1)r1[R-1]={hpt:L*13+6};R++});
  XB.put(a,"C"+(R+1),`${m.profesional} – ${m.matricula}`,{font:XB.F(9),alignment:{horizontal:"right",wrapText:true}});XB.mg(a,`C${R+1}:D${R+1}`);r1[R]={hpt:26};
  const ws1=XB.build(a,[34,22,20,22],r1);
  // Hoja 2
  const b=XB.mk(),r2=[];head(b,titulo);r2[0]={hpt:26};idf(b,2,numFirst?numFirst.h2:null);
  XB.put(b,"A5","Datos de la medición",{font:B,alignment:C_});XB.mg(b,`A5:${last}5`);XB.fill(b,`A5:${last}5`,"D9D9D9");
  headers.forEach((h,k)=>XB.put(b,XLSX.utils.encode_col(k)+"6",h,{font:XB.F(8,true),alignment:C_}));XB.fill(b,`A6:${last}6`,"F2F2F2");r2[5]={hpt:66};
  R=7;rows.forEach((row,i)=>{row.forEach((v,k)=>XB.put(b,XLSX.utils.encode_col(k)+R,v??"",{font:XB.F(9,k<2),alignment:C_}));if(bad&&bad[i])XB.fill(b,`A${R}:${last}${R}`,"F4CCCC");r2[R-1]={hpt:30};R++});
  XB.grid(b,`A6:${last}${Math.max(R-1,6)}`);
  if(info){XB.put(b,"A"+R,info,{font:XB.F(9),alignment:{horizontal:"left",vertical:"top",wrapText:true}});XB.mg(b,`A${R}:${last}${R}`);r2[R-1]={hpt:XB.lines(info,140)*12+10};XB.box(b,`A${R}:${last}${R}`)}
  const ws2=XB.build(b,widths,r2,{left:0.4,right:0.4,top:0.5,bottom:0.5,header:0.3,footer:0.3});
  // Hoja 3
  const c=XB.mk(),r3=[];head(c,titulo);r3[0]={hpt:26};idf(c,2,numFirst?numFirst.h3:null);
  XB.put(c,"A5","Análisis de los datos y mejoras a realizar",{font:B,alignment:C_});XB.mg(c,`A5:${last}5`);XB.fill(c,`A5:${last}5`,"D9D9D9");
  const half=Math.floor(n/2),hc=XLSX.utils.encode_col(half-1),hc2=XLSX.utils.encode_col(half);
  XB.put(c,"A6",`${numFirst?`(${numFirst.c}) `:""}Conclusiones`,{font:B,alignment:C_});XB.mg(c,`A6:${hc}6`);XB.put(c,hc2+"6",`${numFirst?`(${numFirst.r}) `:""}Recomendaciones para adecuar a la legislación vigente`,{font:B,alignment:C_});XB.mg(c,`${hc2}6:${last}6`);
  XB.put(c,"A7",S.conclusiones,{alignment:{horizontal:"left",vertical:"top",wrapText:true}});XB.mg(c,`A7:${hc}7`);XB.put(c,hc2+"7",S.recomendaciones,{alignment:{horizontal:"left",vertical:"top",wrapText:true}});XB.mg(c,`${hc2}7:${last}7`);
  r3[6]={hpt:Math.max(200,Math.max(XB.lines(S.conclusiones,55),XB.lines(S.recomendaciones,55))*13+12)};XB.grid(c,`A6:${last}6`);XB.box(c,`A7:${hc}7`);XB.box(c,`${hc2}7:${last}7`);
  XB.put(c,hc2+"9",`${m.profesional} – ${m.matricula}`,{font:XB.F(9),alignment:{horizontal:"right"}});XB.mg(c,`${hc2}9:${last}9`);
  const ws3=XB.build(c,widths,r3,{left:0.4,right:0.4,top:0.5,bottom:0.5,header:0.3,footer:0.3});
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws1,"Hoja 1 - Datos");XLSX.utils.book_append_sheet(wb,ws2,"Hoja 2 - Mediciones");XLSX.utils.book_append_sheet(wb,ws3,"Hoja 3 - Conclusiones");
  saveWorkbook(wb,XB.fname(fname),["p","l","l"]);
}
function stdEst(titulo,nota,extra){return`<div class="card"><h2>${titulo}</h2><div class="grid">
   ${dfld("meta.fechaMed","Fecha de la medición")}${tfld("meta.horaIni","Hora de inicio")}${tfld("meta.horaFin","Hora de finalización")}${extra||""}</div>
   ${nota?`<p class="muted">${nota}</p>`:""}</div>
  <div class="card"><h2>Instrumento</h2><div class="grid">
   ${fld("meta.instrumento","Instrumento")}${fld("meta.marca","Marca",{req:1})}${fld("meta.modelo","Modelo",{req:1})}${fld("meta.serie","N° de serie",{req:1})}${dfld("meta.fechaCal","Fecha de calibración")}</div>
   <div id="calBox">${calVigencia()}</div></div>
  ${estBlock(!S.meta.razon)}
  <details class="more"><summary>Metodología, observaciones y documentación</summary><div class="in">
   ${fld("meta.metodologia","Metodología utilizada",{type:"textarea",rows:5})}<div style="height:8px"></div>${fld("meta.observaciones","Observaciones",{type:"textarea",rows:2})}
   <label class="chk"><input type="checkbox" data-b="meta.docCert" ${S.meta.docCert?"checked":""}> Certificado de calibración</label>
   <label class="chk"><input type="checkbox" data-b="meta.docCroquis" ${S.meta.docCroquis?"checked":""}> Plano o croquis con los puntos de medición</label></div></details>
  ${profBlock()}`}
const STD_KEYS=[...EST_KEYS,["marca","Marca del instrumento"],["modelo","Modelo del instrumento"],["serie","N° de serie del instrumento"],["fechaCal","Fecha de calibración"],["horaIni","Hora de inicio"],["horaFin","Hora de finalización"],["metodologia","Metodología"]];
function stdChecks(arrKey,label,need){const o=baseChecks(STD_KEYS);if(!S[arrKey].length)o.push(`Al menos un ${label}`);S[arrKey].forEach(it=>{const mi=need(it);if(mi.length)o.push(`${label[0].toUpperCase()+label.slice(1)} ${it.n}: ${mi.join(", ")}`)});
  if(!S.conclusiones.trim())o.push("Conclusiones");if(!S.recomendaciones.trim())o.push("Recomendaciones");return o}
function newMetaFor(instr,metodo,extra){const m=blankMeta();Object.assign(m,{instrumento:instr,marca:"",modelo:"",serie:"",fechaCal:"",metodologia:metodo},extra||{});return m}
const pill=(ok,t)=>`<span class="pill ${ok===null||ok===undefined?"na":ok?"ok":"bad"}">${t}</span>`;
const yesno=b=>b===null?"":b?"SI":"NO";

/* =====================================================================
   RUIDO
   ===================================================================== */
const TIPO_RUIDO=["Continuo o intermitente","De impulso o de impacto"];
const T_PERM=L=>8/Math.pow(2,(L-85)/3); // horas permitidas (85 dBA · 8 h · intercambio 3 dB)
const fmtH=h=>!isFinite(h)?"–":h>=1?fmt(h,2)+" h":h*60>=1?fmt(h*60,1)+" min":fmt(h*3600,2)+" s";
function calcRuido(p){
  const te=num(p.te),o={te,modo:p.modo};
  if(p.modo==="laeq"){const L=num(p.laeq);if(isFinite(L)&&te>0){o.L=L;o.D=te/T_PERM(L);o.Lmax=85+10*log10(8/te)}}
  else if(p.modo==="frac"){let D=0,any=false;(p.tramos||[]).forEach(t=>{const L=num(t.l),c=num(t.c);if(isFinite(L)&&c>0){any=true;if(L>=80)D+=c/T_PERM(L)}});if(any)o.D=D}
  else{const d=num(p.dosis),ti=num(p.ti);if(isFinite(d)){o.dMed=d;o.D=(ti>0&&te>0)?d/100*te/(ti/60):d/100}}
  if(o.D!=null&&o.D>0)o.L8=85+10*log10(o.D);
  const pk=num(p.pico);o.pico=isFinite(pk)?pk:null;o.picoOK=o.pico===null?null:o.pico<=140;
  o.cumple=o.D==null?(o.picoOK===null?null:o.picoOK):(o.D<=1.0001&&o.picoOK!==false);
  return o}
const ruido=makeModule({
  arr:"puntos",nombre:"Ruido",norma:"Res. SRT 85/12",unidad:"puntos",tabs:["Datos","Medir","Análisis","Exportar"],
  nuevo:()=>({tipo:"ruido",updatedAt:Date.now(),meta:newMetaFor("Decibelímetro integrador / dosímetro","Medición según Res. SRT 85/12 con decibelímetro integrador clase 2 (IRAM 4074 / IEC 804) en la posición del oído del trabajador, durante al menos un ciclo de trabajo, y/o dosímetro personal configurado con criterio 85 dBA para 8 h y tasa de intercambio de 3 dB (Res. MTEySS 295/03, Anexo V).",{jornada:"8",condNormales:"",condMedicion:""}),puntos:[],conclusiones:"",recomendaciones:""}),
  newItem:(n,l)=>({id:uid(),n,sector:"",puesto:"",te:l?l.te:"8",ti:"",tipoRuido:"Continuo o intermitente",pico:"",modo:l?l.modo:"laeq",laeq:"",dosis:"",tramos:[{l:"",c:""}],obs:"",prev:null}),
  title:p=>p.sector||"(sin sector)",sub:p=>`${p.puesto||"puesto"} · Te ${p.te||"–"} h · ${p.modo==="dosis"?"dosímetro":p.modo==="frac"?"por tramos":"LAeq"}`,
  calc:calcRuido,
  status:c=>c.cumple===null?{cls:"na",txt:"Sin medir"}:c.cumple?{cls:"ok",txt:"Cumple"}:{cls:"bad",txt:"No cumple"},
  value:(p,c)=>c.L8!=null?fmt(c.L8,1)+" dBA":"–",
  emptyT:"Sin puntos de medición",emptyD:"Agregá cada puesto o puesto tipo con su tiempo de exposición diaria.",addLabel:"Agregar punto de medición",
  est:()=>stdEst("Medición de ruido",`Límite: 85 dBA para 8 h diarias, tasa de intercambio 3 dB, nivel pico máximo 140 dBC (Res. MTEySS 295/03, Anexo V). Protocolo Res. SRT 85/12.`,
    `${fld("meta.jornada","(12) Duración de la jornada (h)",{im:"decimal"})}`)+`<div class="card"><h2>(13–14) Condiciones</h2>${fld("meta.condNormales","(13) Condiciones normales de trabajo y fuentes de ruido",{type:"textarea",rows:2})}<div style="height:8px"></div>${fld("meta.condMedicion","(14) Condiciones de trabajo durante la medición (si difieren)",{type:"textarea",rows:2})}</div>`,
  fields:(i,p)=>{const P=`puntos.${i}.`;return`<div class="card"><div class="grid">${fld(P+"n","(23) Punto",{im:"numeric"})}${fld(P+"sector","(24) Sector",{req:1})}${fld(P+"puesto","(25) Puesto / puesto tipo",{req:1})}
     ${fld(P+"te","(26) Tiempo de exposición diaria Te (h)",{req:1,im:"decimal"})}${fld(P+"ti","(27) Tiempo de integración (min)",{im:"decimal"})}</div>
     <div style="margin-top:10px">${segF(P+"tipoRuido","(28) Características del ruido",TIPO_RUIDO,p.tipoRuido)}</div>
     <div style="margin-top:10px">${fld(P+"pico","(29) Nivel pico LCpico (dBC) · máx. 140",{im:"decimal"})}</div></div>
    <div class="card"><h2>Resultado de la medición</h2>
     <div class="seg" style="margin-bottom:10px"><button data-rmodo="laeq" class="${p.modo==="laeq"?"on":""}">(30) LAeq,Te</button><button data-rmodo="frac" class="${p.modo==="frac"?"on":""}">(31) Por tramos</button><button data-rmodo="dosis" class="${p.modo==="dosis"?"on":""}">(32) Dosímetro</button></div>
     ${p.modo==="laeq"?fld(P+"laeq","Nivel sonoro continuo equivalente LAeq,Te (dBA)",{req:1,im:"decimal"})
      :p.modo==="dosis"?`${fld(P+"dosis","Dosis medida (%)",{req:1,im:"decimal"})}<p class="muted">Si el tiempo de integración (27) es menor que Te, la dosis se proyecta a la jornada: D = dosis × Te / Ti.</p>`
      :`<div id="trList">${p.tramos.map((t,j)=>`<div style="display:grid;grid-template-columns:1fr 1fr 40px;gap:6px;align-items:end;margin-bottom:8px">
         <label class="f"><span>${j===0?"Nivel LAeq (dBA)":""}</span><input data-b="${P}tramos.${j}.l" inputmode="decimal" value="${esc(t.l)}"></label>
         <label class="f"><span>${j===0?"Tiempo Ci (h)":""}</span><input data-b="${P}tramos.${j}.c" inputmode="decimal" value="${esc(t.c)}"></label>
         <button class="navb" data-deltr="${j}" style="width:40px;height:46px" aria-label="Quitar tramo">✕</button></div>`).join("")}</div>
       <button class="btn wide" id="addTr">＋ Agregar tramo</button><p class="muted">Se suman las fracciones Ci/Ti de los tramos ≥ 80 dBA (Ti = tiempo permitido para cada nivel).</p>`}</div>
    <div class="card">${fld(P+"obs","(34) Información adicional",{type:"textarea",rows:2})}</div>`},
  live:(p,c)=>`<div class="live"><div><small>Dosis diaria</small><b>${c.D!=null?fmt(c.D*100,0)+"%":"–"}</b></div><div><small>LAeq,8h</small><b>${c.L8!=null?fmt(c.L8,1):"–"}</b><small>dBA</small></div><div><small>${p.prev&&p.prev.L8!=null?"Anterior":"Máx. para Te"}</small><b>${p.prev&&p.prev.L8!=null?fmt(p.prev.L8,1):c.Lmax!=null?fmt(c.Lmax,1):(num(p.te)>0?fmt(85+10*log10(8/num(p.te)),1):"–")}</b><small>dBA</small></div></div>
    <div class="verdict">${c.D!=null?pill(c.D<=1.0001,c.D<=1.0001?"Dosis ≤ 100 %":"Dosis > 100 %"):pill(null,"Cargá el resultado")}${c.pico!==null?pill(c.picoOK,`Pico ${fmt(c.pico,1)} ${c.picoOK?"≤":">"} 140 dBC`):""}${c.L!=null?`<span class="pill na">Tiempo permitido a ${fmt(c.L,1)} dBA: ${fmtH(T_PERM(c.L))}</span>`:""}</div>`,
  bindExtra:st=>{$$("[data-rmodo]").forEach(b=>b.onclick=()=>{S.puntos[st.e].modo=b.dataset.rmodo;save();rerender()});
    const a=$("#addTr");if(a)a.onclick=()=>{S.puntos[st.e].tramos.push({l:"",c:""});save();rerender()};
    $$("[data-deltr]").forEach(b=>b.onclick=()=>{S.puntos[st.e].tramos.splice(+b.dataset.deltr,1);save();rerender()})},
  summaryTable:()=>`<div class="card"><h2>Resumen</h2><div class="tablewrap"><table class="sum"><tr><th>N°</th><th>Sector / puesto</th><th>Te</th><th>LAeq,8h</th><th>Dosis</th><th>Pico</th><th>Estado</th></tr>
    ${S.puntos.map(p=>{const c=calcRuido(p);return`<tr><td>${p.n}</td><td>${esc(p.sector)}<br><span class="muted">${esc(p.puesto)}</span></td><td>${esc(p.te)} h</td><td>${c.L8!=null?fmt(c.L8,1):"–"}</td><td>${c.D!=null?fmt(c.D*100,0)+"%":"–"}</td><td>${c.pico!==null?fmt(c.pico,1):"–"}</td><td>${pill(c.cumple,c.cumple===null?"–":c.cumple?"Cumple":"No cumple")}</td></tr>`}).join("")}</table></div></div>`,
  checks:()=>stdChecks("puntos","punto",p=>{const c=calcRuido(p),mi=[];if(!p.sector)mi.push("sector");if(!(num(p.te)>0))mi.push("tiempo de exposición");if(c.D==null)mi.push("resultado");return mi}),
  autoText:()=>{const L=S.puntos.map(p=>({p,c:calcRuido(p)})),bad=L.filter(x=>x.c.cumple===false),ok=L.filter(x=>x.c.cumple),v=L.filter(x=>x.c.L8!=null).map(x=>x.c.L8);const t=[];
    t.push(`Se evaluaron ${L.length} puntos de medición en el establecimiento ${S.meta.razon}${v.length?`, con niveles equivalentes diarios (LAeq,8h) entre ${fmt(Math.min(...v),1)} y ${fmt(Math.max(...v),1)} dBA`:""}, comparados con el valor límite de 85 dBA para 8 h y tasa de intercambio de 3 dB de la Res. MTEySS 295/03, Anexo V.`);
    if(ok.length)t.push(`Cumplen con la legislación vigente los puntos N° ${ok.map(x=>x.p.n).join(", ")}.`);
    if(bad.length)t.push(`No cumplen: ${bad.map(x=>`N°${x.p.n} (${x.p.sector} – ${x.p.puesto}${x.c.D!=null?`, dosis ${fmt(x.c.D*100,0)} %`:""}${x.c.picoOK===false?", pico > 140 dBC":""})`).join("; ")}.`);
    S.conclusiones=t.join("\n\n");
    const r=[];if(bad.length){r.push("Priorizar medidas de ingeniería sobre la fuente (encapsulado, mantenimiento, aislamiento antivibratorio) y sobre la transmisión (barreras, tratamiento acústico), y medidas administrativas de reducción del tiempo de exposición.");
      r.push(`Uso obligatorio de protección auditiva certificada con atenuación adecuada en los puestos N° ${bad.map(x=>x.p.n).join(", ")}, con capacitación en su uso y señalización del área.`);
      r.push("Incluir al personal expuesto en el examen periódico con audiometría y declarar el agente ruido en la nómina de expuestos de la ART.")}
    r.push("Repetir la medición con frecuencia anual o ante cambios en el proceso, las máquinas o los puestos.");
    S.recomendaciones=r.map(x=>"• "+x).join("\n")},
  exportFn:()=>{const m=S.meta;exportStd({titulo:"PROTOCOLO PARA LA MEDICIÓN DEL NIVEL DE RUIDO EN EL AMBIENTE LABORAL",norma:"Res. SRT 85/12 · Valores límite Res. MTEySS 295/03, Anexo V (85 dBA – 8 h – Q = 3 dB · 140 dBC pico)",fname:"Ruido_85-12",
    datos:[["(1) Razón social",m.razon],["(2) Dirección",m.direccion],["(3) Localidad",m.localidad],["(4) Provincia",m.provincia],["(5) C.P.",m.cp],["(6) C.U.I.T.",m.cuit],[null,"Datos de la medición"],
      ["(7) Marca, modelo y N° de serie del instrumento",`${m.instrumento} ${m.marca} ${m.modelo} – N° de serie: ${m.serie}`],["(8) Fecha de calibración",fdate(m.fechaCal)],["(9) Fecha de la medición",fdate(m.fechaMed)],["(10) Hora de inicio",m.horaIni],["(11) Hora de finalización",m.horaFin],
      ["(12) Duración de la jornada (h)",m.jornada],["(13) Condiciones normales de trabajo",m.condNormales],["(14) Condiciones durante la medición",m.condMedicion||"Iguales a las normales"],["Metodología",m.metodologia],["Observaciones",m.observaciones],[null,"Documentación que se adjuntará"],["(15) Certificado de calibración",m.docCert?"SI":"NO"],["(16) Plano o croquis",m.docCroquis?"SI":"NO"]],
    headers:["(23) Punto de medición","(24) Sector","(25) Puesto / puesto tipo","(26) Tiempo de exposición Te (h)","(27) Tiempo de integración (min)","(28) Características del ruido","(29) Nivel pico LCpico (dBC)","(30) LAeq,Te (dBA)","(31) Suma de fracciones","(32) Dosis (%)","LAeq,8h (dBA)","(33) Cumple SI/NO"],
    rows:S.puntos.map(p=>{const c=calcRuido(p);return[+p.n||p.n,p.sector,p.puesto,nOrNull(p.te),nOrNull(p.ti),p.tipoRuido,c.pico,p.modo==="laeq"?c.L:"",p.modo==="frac"&&c.D!=null?Math.round(c.D*1000)/1000:"",p.modo==="dosis"&&c.D!=null?Math.round(c.D*1000)/10:"",c.L8!=null?Math.round(c.L8*10)/10:"",yesno(c.cumple)]}),
    bad:S.puntos.map(p=>calcRuido(p).cumple===false),widths:[8,15,16,10,10,13,10,10,10,9,10,9],
    info:`(34) Información adicional: ${S.puntos.filter(p=>p.obs).map(p=>`Punto ${p.n}: ${p.obs}`).join(" · ")||"—"}${S.puntos.some(p=>p.modo==="frac")?" · Suma de fracciones: "+S.puntos.filter(p=>p.modo==="frac").map(p=>`Pto ${p.n}: `+p.tramos.filter(t=>num(t.l)>=80&&num(t.c)>0).map(t=>`${t.c} h / ${fmt(T_PERM(num(t.l)),2)} h (${t.l} dBA)`).join(" + ")).join(" · "):""}`,
    numFirst:{h2:[17,18,19,20,21,22],h3:[35,36,37,38,39,40],c:41,r:42}})},
  prevOf:(p,c)=>({L8:c.L8??null}),clearOnRepeat:p=>{p.laeq="";p.dosis="";p.pico="";p.tramos=[{l:"",c:""}]},clearOnDup:p=>{p.laeq="";p.dosis="";p.pico=""}
});

/* =====================================================================
   VIBRACIONES
   ===================================================================== */
// Mano-brazo – Res. 295/03: componente dominante, aceleración rms ponderada (m/s²) según duración total diaria
function limMB(t){return t>=4?4:t>=2?6:t>=1?8:12}
// Cuerpo entero – curvas de la Res. 295/03 (ISO 2631) en la banda de máxima sensibilidad: Z 4–8 Hz, X/Y 1–2 Hz
const CE_T=[24,16,8,4,2.5,1,25/60,16/60,1/60],CE_Z=[0.14,0.212,0.315,0.53,0.71,1.18,1.8,2.12,2.8],CE_XY=[0.1,0.15,0.224,0.355,0.5,0.85,1.25,1.5,2.0];
function limCE(t,tab){if(!(t>0))return null;if(t>=CE_T[0])return tab[0];if(t<=CE_T.at(-1))return tab.at(-1);
  for(let i=0;i<CE_T.length-1;i++){const a=CE_T[i],b=CE_T[i+1];if(t<=a&&t>=b){const f=(log10(t)-log10(a))/(log10(b)-log10(a));return Math.pow(10,log10(tab[i])+f*(log10(tab[i+1])-log10(tab[i])))}}return null}
function calcVib(p){const ax=num(p.ax),ay=num(p.ay),az=num(p.az),t=num(p.t),o={t};const have=[ax,ay,az].filter(isFinite);if(!have.length||!(t>0))return o;
  if(p.seg==="mb"){o.dom=Math.max(...have);o.lim=limMB(t);o.ahv=Math.sqrt(have.reduce((s,x)=>s+sq(x),0));o.A8=o.ahv*Math.sqrt(t/8);o.ok=o.dom<=o.lim}
  else{o.lz=limCE(t,CE_Z);o.lxy=limCE(t,CE_XY);o.zok=isFinite(az)?az<=o.lz:null;o.xok=isFinite(ax)?ax<=o.lxy:null;o.yok=isFinite(ay)?ay<=o.lxy:null;
    o.awt=Math.sqrt(sq(1.4*(isFinite(ax)?ax:0))+sq(1.4*(isFinite(ay)?ay:0))+sq(isFinite(az)?az:0));o.A8=Math.max(1.4*(ax||0),1.4*(ay||0),az||0)*Math.sqrt(t/8);
    o.ok=[o.zok,o.xok,o.yok].every(v=>v!==false)}
  const cf=num(p.fc);o.fc=isFinite(cf)?cf:null;o.fcOK=o.fc===null?null:o.fc<=6;return o}
function vibEval(p,c){if(c.ok==null)return{};if(p.seg==="mb")return{ev:c.dom,lim:c.lim};
  const z=num(p.az),xy=Math.max(num(p.ax)||0,num(p.ay)||0),rz=isFinite(z)?z/c.lz:-1,rxy=xy/c.lxy;return rz>=rxy?{ev:z,lim:c.lz,eje:"Z"}:{ev:xy,lim:c.lxy,eje:"X/Y"}}
const vib=makeModule({
  arr:"puntos",nombre:"Vibraciones",norma:"Res. MTEySS 295/03",unidad:"puntos",tabs:["Datos","Medir","Análisis","Exportar"],
  nuevo:()=>({tipo:"vib",updatedAt:Date.now(),meta:newMetaFor("Vibrómetro / analizador de vibraciones triaxial","Medición de la aceleración eficaz (rms) ponderada en frecuencia en los tres ejes ortogonales. Mano-brazo: acelerómetro en la zona de ingreso de la vibración a la mano, rango 5–1500 Hz; se evalúa la componente dominante según la duración total diaria. Cuerpo entero: acelerómetro en almohadilla bajo el asiento del operador con el equipo en funcionamiento, medición mínima de 1 minuto, rango 1–80 Hz; se compara cada eje con los valores límite de la Res. MTEySS 295/03 (Anexo V) en la banda de mayor sensibilidad (Z: 4–8 Hz; X/Y: 1–2 Hz)."),puntos:[],conclusiones:"",recomendaciones:""}),
  newItem:(n,l)=>({id:uid(),n,sector:"",puesto:"",equipo:"",seg:l?l.seg:"mb",t:"",ax:"",ay:"",az:"",fc:"",obs:"",prev:null}),
  title:p=>p.equipo||p.puesto||"(sin equipo)",sub:p=>`${p.seg==="mb"?"Mano-brazo":"Cuerpo entero"} · ${p.sector||"sector"} · ${p.t||"–"} h/día`,
  calc:calcVib,status:c=>c.ok==null?{cls:"na",txt:"Sin medir"}:c.ok?{cls:"ok",txt:"Cumple"}:{cls:"bad",txt:"No cumple"},
  value:(p,c)=>p.seg==="mb"?(c.dom!=null?fmt(c.dom,2):"–"):(c.awt!=null?fmt(c.awt,2):"–"),
  emptyT:"Sin mediciones",emptyD:"Agregá cada equipo o herramienta con el tipo de exposición (mano-brazo o cuerpo entero) y el tiempo diario.",addLabel:"Agregar medición",
  est:()=>stdEst("Medición de vibraciones","Mano-brazo: valor límite de la componente dominante 4 m/s² (4 a 8 h), 6 (2 a 4 h), 8 (1 a 2 h) y 12 m/s² (menos de 1 h). Cuerpo entero: curvas por eje y tiempo de exposición. Res. MTEySS 295/03, Anexo V."),
  fields:(i,p)=>{const P=`puntos.${i}.`;return`<div class="card"><div style="margin-bottom:10px">${segF(P+"seg","Tipo de exposición",["mb","ce"],p.seg).replace(">mb<",">Mano-brazo<").replace(">ce<",">Cuerpo entero<")}</div><div class="grid">
     ${fld(P+"n","N°",{im:"numeric"})}${fld(P+"sector","Sector",{req:1})}${fld(P+"puesto","Puesto")}${fld(P+"equipo",p.seg==="mb"?"Herramienta / equipo":"Vehículo / máquina",{req:1})}
     ${fld(P+"t","Duración total diaria de la exposición (h)",{req:1,im:"decimal"})}</div></div>
    <div class="card"><h2>Aceleración rms ponderada (m/s²)</h2><div class="g3">${fld(P+"ax","Eje X",{im:"decimal"})}${fld(P+"ay","Eje Y",{im:"decimal"})}${fld(P+"az","Eje Z",{im:"decimal"})}</div>
     ${p.seg==="ce"?`<div style="margin-top:10px">${fld(P+"fc","Factor de cresta (≤ 6)",{im:"decimal"})}</div>`:""}
     <p class="muted">${p.seg==="mb"?"Se evalúa la mayor de las tres componentes contra el valor límite según la duración diaria.":"Z se compara con el límite longitudinal (4–8 Hz) y X/Y con el transversal (1–2 Hz) para el tiempo de exposición."}</p></div>
    <div class="card">${fld(P+"obs","Observaciones",{type:"textarea",rows:2})}</div>`},
  live:(p,c)=>{if(c.ok==null)return`<p class="muted" style="margin:0">Cargá la duración diaria y al menos un eje.</p>`;
    return p.seg==="mb"?`<div class="live"><div><small>Comp. dominante</small><b>${fmt(c.dom,2)}</b><small>m/s²</small></div><div><small>Límite (${fmt(c.t,1)} h)</small><b>${fmt(c.lim,0)}</b><small>m/s²</small></div><div><small>A(8) ref.</small><b>${fmt(c.A8,2)}</b><small>m/s²</small></div></div><div class="verdict">${pill(c.ok,c.ok?"Cumple Res. 295/03":"Supera el valor límite")}${c.A8>5?pill(false,"A(8) > 5 (ref. Dir. 2002/44/CE)"):c.A8>2.5?`<span class="pill warn">A(8) > 2,5: nivel de acción (ref.)</span>`:""}</div>`
    :`<div class="live"><div><small>Z · límite</small><b>${isFinite(num(p.az))?fmt(num(p.az),2):"–"}·${fmt(c.lz,2)}</b></div><div><small>X/Y · límite</small><b>${fmt(Math.max(num(p.ax)||0,num(p.ay)||0),2)}·${fmt(c.lxy,2)}</b></div><div><small>Awt</small><b>${fmt(c.awt,2)}</b><small>m/s²</small></div></div>
     <div class="verdict">${pill(c.zok,"Eje Z")}${pill(c.xok,"Eje X")}${pill(c.yok,"Eje Y")}${c.fc!==null?pill(c.fcOK,`Factor de cresta ${fmt(c.fc,1)}`):""}${c.A8>1.15?pill(false,"A(8) > 1,15 (ref. UE)"):c.A8>0.5?`<span class="pill warn">A(8) > 0,5: nivel de acción (ref.)</span>`:""}</div>`},
  onChange:p=>{if(/\.seg$/.test(p)){rerender();return true}},
  summaryTable:()=>`<div class="card"><h2>Resumen</h2><div class="tablewrap"><table class="sum"><tr><th>N°</th><th>Equipo</th><th>Tipo</th><th>h/día</th><th>Valor</th><th>Límite</th><th>Estado</th></tr>
    ${S.puntos.map(p=>{const c=calcVib(p);return`<tr><td>${p.n}</td><td>${esc(p.equipo)}<br><span class="muted">${esc(p.sector)}</span></td><td>${p.seg==="mb"?"M-B":"C-E"}</td><td>${esc(p.t)}</td><td>${fmt(vibEval(p,c).ev,2)}</td><td>${fmt(vibEval(p,c).lim,2)}</td><td>${pill(c.ok,c.ok==null?"–":c.ok?"Cumple":"No cumple")}</td></tr>`}).join("")}</table></div></div>`,
  checks:()=>stdChecks("puntos","punto",p=>{const c=calcVib(p),mi=[];if(!p.equipo)mi.push("equipo");if(!(num(p.t)>0))mi.push("duración diaria");if(c.ok==null)mi.push("aceleraciones");return mi}),
  autoText:()=>{const L=S.puntos.map(p=>({p,c:calcVib(p)})),bad=L.filter(x=>x.c.ok===false),ok=L.filter(x=>x.c.ok);const t=[];
    t.push(`Se evaluaron ${L.length} exposiciones a vibraciones (${L.filter(x=>x.p.seg==="mb").length} mano-brazo y ${L.filter(x=>x.p.seg==="ce").length} cuerpo entero) en ${S.meta.razon}, comparando la aceleración rms ponderada en frecuencia con los valores límite de la Res. MTEySS 295/03, Anexo V, según el tiempo de exposición diario.`);
    if(ok.length)t.push(`Cumplen: ${ok.map(x=>`N°${x.p.n} (${x.p.equipo})`).join(", ")}.`);if(bad.length)t.push(`Superan el valor límite: ${bad.map(x=>`N°${x.p.n} (${x.p.equipo})`).join(", ")}.`);
    S.conclusiones=t.join("\n\n");const r=[];
    if(bad.length){r.push("Reducir la vibración en la fuente: mantenimiento y balanceo de herramientas y equipos, reemplazo por modelos de baja vibración, asientos con suspensión y mantenimiento de caminos y rodados.");r.push("Reducir el tiempo de exposición diario mediante rotación de tareas y pausas, y proveer guantes antivibratorios certificados en tareas mano-brazo.");r.push("Incluir al personal expuesto en los exámenes médicos periódicos correspondientes y declarar el agente en la nómina de expuestos de la ART.")}
    r.push("Repetir la evaluación ante cambios de equipos, herramientas o condiciones de trabajo.");S.recomendaciones=r.map(x=>"• "+x).join("\n")},
  exportFn:()=>{const m=S.meta;exportStd({titulo:"INFORME DE MEDICIÓN DE VIBRACIONES",norma:"Res. MTEySS 295/03, Anexo V – Vibración segmental (mano-brazo) y de cuerpo entero",fname:"Vibraciones",
    datos:[["Razón social",m.razon],["Dirección",m.direccion],["Localidad",m.localidad],["Provincia",m.provincia],["C.P.",m.cp],["C.U.I.T.",m.cuit],[null,"Datos de la medición"],["Instrumento (marca, modelo, N° de serie)",`${m.instrumento} ${m.marca} ${m.modelo} – N° de serie: ${m.serie}`],["Fecha de calibración",fdate(m.fechaCal)],["Fecha de la medición",fdate(m.fechaMed)],["Hora de inicio / finalización",`${m.horaIni} – ${m.horaFin}`],["Metodología",m.metodologia],["Observaciones",m.observaciones],[null,"Documentación que se adjuntará"],["Certificado de calibración",m.docCert?"SI":"NO"],["Plano o croquis",m.docCroquis?"SI":"NO"]],
    headers:["N°","Sector","Puesto","Equipo / herramienta","Tipo de exposición","Duración diaria (h)","ax (m/s²)","ay (m/s²)","az (m/s²)","Valor evaluado (m/s²)","Valor límite (m/s²)","Cumple SI/NO"],
    rows:S.puntos.map(p=>{const c=calcVib(p),e=vibEval(p,c);
      return[+p.n||p.n,p.sector,p.puesto,p.equipo,p.seg==="mb"?"Mano-brazo":"Cuerpo entero",nOrNull(p.t),nOrNull(p.ax),nOrNull(p.ay),nOrNull(p.az),e.ev!=null?Math.round(e.ev*100)/100:"",e.lim!=null?Math.round(e.lim*100)/100:"",yesno(c.ok??null)]}),
    bad:S.puntos.map(p=>calcVib(p).ok===false),widths:[6,14,14,18,12,10,9,9,9,11,11,9],
    info:"Mano-brazo: se evalúa la componente dominante de la aceleración rms ponderada (Res. 295/03: 4 m/s² de 4 a 8 h; 6 de 2 a 4 h; 8 de 1 a 2 h; 12 m/s² menos de 1 h). Cuerpo entero: se compara el eje Z con la curva longitudinal (4–8 Hz) y X/Y con la transversal (1–2 Hz) para el tiempo de exposición; se informa el eje más desfavorable. "+(S.puntos.filter(p=>p.obs).map(p=>`N°${p.n}: ${p.obs}`).join(" · "))})},
  prevOf:(p,c)=>({v:p.seg==="mb"?c.dom:c.awt}),clearOnRepeat:p=>{p.ax="";p.ay="";p.az="";p.fc=""},clearOnDup:p=>{p.ax="";p.ay="";p.az=""}
});

/* =====================================================================
   TERMOHIGROMETRÍA Y ESTRÉS TÉRMICO
   ===================================================================== */
const ROPA=[["0","Camisa de algodón y pantalón largo / mameluco de tela (0)"],["0.5","Mameluco SMS de polipropileno, una capa (+0,5)"],["1","Mameluco de poliolefina, una capa (+1)"],["3","Ropa tejida de doble capa (+3)"],["11","Mameluco con capucha barrera de vapor (+11)"]];
const TM_CAT=[["115","Reposo · 115 W"],["180","Liviano · 180 W"],["300","Moderado · 300 W"],["415","Pesado · 415 W"],["520","Muy pesado · 520 W"]];
const VLP=M=>56.7-11.5*log10(M),VLA=M=>59.9-14.1*log10(M);
const tgbhOf=(r,solar)=>{const bh=num(r.tbh),g=num(r.tg),bs=num(r.tbs);if(!isFinite(bh)||!isFinite(g))return NaN;return solar?0.7*bh+0.2*g+0.1*(isFinite(bs)?bs:g):0.7*bh+0.3*g};
// Temperatura equivalente por enfriamiento del viento (Siple-Passel), v en m/s
const teqFrio=(T,v)=>!(isFinite(T))?NaN:(!(v>1.8)?T:33-(33-T)*(0.474+0.454*Math.sqrt(v)-0.0454*v));
function calcTermo(p){const o={};const solar=p.solar==="SI";
  const R=(p.lect||[]).map(r=>({t:num(r.min),g:tgbhOf(r,solar),m:num(r.tm)})).filter(r=>r.t>0&&isFinite(r.g));
  const tt=R.reduce((s,r)=>s+r.t,0);
  if(tt>0){o.tgbh=R.reduce((s,r)=>s+r.g*r.t,0)/tt;const mm=R.filter(r=>r.m>0);o.tm=mm.length?mm.reduce((s,r)=>s+r.m*r.t,0)/mm.reduce((s,r)=>s+r.t,0):num(p.tmDef);
    o.ef=o.tgbh+(num(p.var)||0);o.vlp=VLP(o.tm);o.vla=VLA(o.tm);o.lim=p.aclim==="SI"?o.vlp:o.vla;o.ok=o.ef<=o.lim;o.sobreVLA=o.ef>o.vla;o.minutos=tt;
    const B=num(p.tgbhDesc);if(!o.ok&&isFinite(B)){const A=33.7;if(B<A){const ft=(A-B)/((o.ef-o.lim)+(A-B));o.ft=Math.max(0,Math.min(1,ft))}else o.ft=0}}
  const T=num((p.lect&&p.lect[0]&&p.lect[0].tbs)),v=num(p.va);o.tbs=isFinite(T)?T:null;o.hr=nOrNull(p.hr);o.va=isFinite(v)?v:null;
  if(o.tbs!==null&&o.tbs<16){o.teq=teqFrio(o.tbs,o.va||0)}
  return o}
function frioTxt(c){if(c.teq==null)return[];const r=[];const T=c.tbs,te=c.teq;
  if(te<=-32)r.push(["bad","Teq ≤ −32 °C: no se permite la exposición continua de la piel"]);else if(te<=-12)r.push(["bad","Teq ≤ −12 °C: trabajo bajo supervisión permanente (sistema de parejas)"]);else if(te<=-7)r.push(["warn","Teq ≤ −7 °C: refugios con calefacción próximos al puesto"]);
  if(T<4)r.push(["warn","T < 4 °C: ropa de protección aislante seca"]);
  if(T<=-17.5)r.push(["warn","T ≤ −17,5 °C: manoplas"]);else if(T<-7)r.push(["warn","T < −7 °C: guantes para trabajo moderado"]);else if(T<4)r.push(["warn","T < 4 °C: guantes para trabajo liviano"]);else if(T<16)r.push(["na","T < 16 °C: guantes en trabajo sedentario de precisión"]);
  return r}
const termo=makeModule({
  arr:"puntos",nombre:"Estrés térmico",norma:"Res. SRT 30/23 · Res. 295/03",unidad:"puestos",tabs:["Datos","Medir","Análisis","Exportar"],
  nuevo:()=>({tipo:"termo",updatedAt:Date.now(),meta:newMetaFor("Monitor de estrés térmico TGBH","Medición de temperatura de bulbo seco, bulbo húmedo natural y globo según UNE-EN ISO 7243:2017, a la altura del abdomen del trabajador, luego de la estabilización del instrumento. TGBH sin carga solar = 0,7·tbh + 0,3·tg; con carga solar = 0,7·tbh + 0,2·tg + 0,1·tbs. Se pondera en el tiempo durante una hora, se ajusta por ropa (VAR) y se compara con el valor límite (VLP, aclimatados) o el nivel de acción (VLA, no aclimatados) según la tasa metabólica (ISO 8996), conforme a la Res. SRT 30/2023. Estrés por frío según Res. MTEySS 295/03."),puntos:[],conclusiones:"",recomendaciones:""}),
  newItem:(n,l)=>({id:uid(),n,sector:"",puesto:"",tarea:"",solar:l?l.solar:"NO",aclim:l?l.aclim:"SI",var:l?l.var:"0",hr:"",va:"",tgbhDesc:"",tmDef:"300",lect:[{min:"60",tbs:"",tbh:"",tg:"",tm:l&&l.lect&&l.lect[0]?l.lect[0].tm:"300"}],obs:"",prev:null}),
  title:p=>p.sector||"(sin sector)",sub:p=>`${p.puesto||"puesto"} · ${p.aclim==="SI"?"aclimatado":"no aclimatado"}${p.solar==="SI"?" · con sol":""}`,
  calc:calcTermo,status:c=>c.ok==null?(c.teq!=null?{cls:"warn",txt:"Frío"}:{cls:"na",txt:"Sin medir"}):c.ok?{cls:"ok",txt:"Cumple"}:{cls:"bad",txt:"Supera límite"},
  value:(p,c)=>c.ef!=null?fmt(c.ef,1)+" °C":"–",
  emptyT:"Sin puestos medidos",emptyD:"Agregá cada puesto o grupo de exposición homogénea con sus lecturas de bulbo seco, bulbo húmedo y globo.",addLabel:"Agregar puesto",
  est:()=>stdEst("Termohigrometría y estrés térmico","Calor: Res. SRT 30/2023 (TGBH, VLP = 56,7 − 11,5·log TM para aclimatados; VLA = 59,9 − 14,1·log TM para no aclimatados). Frío: Res. MTEySS 295/03 (temperatura central ≥ 36 °C, temperatura equivalente por viento)."),
  fields:(i,p)=>{const P=`puntos.${i}.`;return`<div class="card"><div class="grid">${fld(P+"n","N°",{im:"numeric"})}${fld(P+"sector","Sector",{req:1})}${fld(P+"puesto","Puesto / grupo homogéneo")}${fld(P+"tarea","Tarea")}</div>
     <div style="display:grid;gap:10px;margin-top:10px">${segF(P+"solar","Carga solar",["NO","SI"],p.solar)}${segF(P+"aclim","Personal aclimatado",["SI","NO"],p.aclim)}</div>
     <div style="margin-top:10px">${fld(P+"var","Vestimenta (VAR)",{opts:ROPA})}</div></div>
    <div class="card"><h2>Lecturas en la hora más desfavorable</h2>
     ${p.lect.map((r,j)=>`<div style="border-top:${j?"1px solid var(--line)":"0"};padding-top:${j?"10px":"0"};margin-top:${j?"10px":"0"}">
       <div class="g3">${fld(`${P}lect.${j}.tbs`,"T bulbo seco °C",{im:"decimal"})}${fld(`${P}lect.${j}.tbh`,"T bulbo húmedo °C",{im:"decimal"})}${fld(`${P}lect.${j}.tg`,"T globo °C",{im:"decimal"})}</div>
       <div class="grid" style="margin-top:8px">${fld(`${P}lect.${j}.min`,"Minutos en la hora",{im:"numeric"})}${fld(`${P}lect.${j}.tm`,"Tasa metabólica",{opts:[...TM_CAT,...(TM_CAT.some(x=>x[0]===String(r.tm))?[]:[[String(r.tm),r.tm+" W"]])]})}</div>
       ${j?`<button class="btn sm d" data-dell="${j}" style="margin-top:8px">Quitar lectura</button>`:""}</div>`).join("")}
     <button class="btn wide" id="addL" style="margin-top:10px">＋ Agregar otra condición dentro de la hora</button>
     <p class="muted">Si el trabajador alterna tareas o lugares, cargá cada condición con sus minutos: TGBH y tasa metabólica se ponderan en la hora.</p></div>
    <div class="card"><h2>Termohigrometría</h2><div class="grid">${fld(P+"hr","Humedad relativa (%)",{im:"decimal"})}${fld(P+"va","Velocidad del aire (m/s)",{im:"decimal"})}${fld(P+"tgbhDesc","TGBH en zona de descanso (°C)",{im:"decimal",ph:"para régimen trabajo/descanso"})}</div></div>
    <div class="card">${fld(P+"obs","Observaciones",{type:"textarea",rows:2})}</div>`},
  live:(p,c)=>{const fr=frioTxt(c);let h="";
    if(c.ef!=null)h+=`<div class="live"><div><small>TGBH ef.</small><b>${fmt(c.ef,1)}</b><small>°C</small></div><div><small>${p.aclim==="SI"?"VLP":"VLA"} (${fmt(c.tm,0)} W)</small><b>${fmt(c.lim,1)}</b><small>°C</small></div><div><small>${p.prev&&p.prev.ef!=null?"Anterior":"HR · aire"}</small><b>${p.prev&&p.prev.ef!=null?fmt(p.prev.ef,1):(c.hr!=null?fmt(c.hr,0)+"%":"–")}</b><small>${p.prev&&p.prev.ef!=null?"°C":(c.va!=null?fmt(c.va,1)+" m/s":"")}</small></div></div>
      <div class="verdict">${pill(c.ok,c.ok?"Cumple Res. SRT 30/23":"Supera el valor límite")}${c.sobreVLA&&p.aclim==="SI"?`<span class="pill warn">Supera el VLA: declarar a la ART (ESOP 80001)</span>`:""}${c.ft!=null?`<span class="pill warn">Trabajo ${fmt(c.ft*60,0)} min / descanso ${fmt(60-c.ft*60,0)} min por hora</span>`:""}${c.minutos!==60?`<span class="pill na">Lecturas: ${fmt(c.minutos,0)} min de 60</span>`:""}</div>`;
    else h+=`<p class="muted" style="margin:0 0 6px">Cargá bulbo húmedo y globo para calcular el TGBH.</p>`;
    if(fr.length)h+=`<div class="verdict">${c.teq!=null?`<span class="pill na">Temp. equivalente por viento ${fmt(c.teq,1)} °C</span>`:""}${fr.map(([k,t])=>`<span class="pill ${k}">${t}</span>`).join("")}</div>`;
    return h},
  bindExtra:st=>{const a=$("#addL");if(a)a.onclick=()=>{const p=S.puntos[st.e],l=p.lect.at(-1);p.lect.push({min:"",tbs:"",tbh:"",tg:"",tm:l?l.tm:"300"});save();rerender()};
    $$("[data-dell]").forEach(b=>b.onclick=()=>{S.puntos[st.e].lect.splice(+b.dataset.dell,1);save();rerender()})},
  summaryTable:()=>`<div class="card"><h2>Resumen</h2><div class="tablewrap"><table class="sum"><tr><th>N°</th><th>Sector / puesto</th><th>TGBH ef.</th><th>TM</th><th>Límite</th><th>Estado</th></tr>
    ${S.puntos.map(p=>{const c=calcTermo(p);return`<tr><td>${p.n}</td><td>${esc(p.sector)}<br><span class="muted">${esc(p.puesto)}</span></td><td>${c.ef!=null?fmt(c.ef,1):"–"}</td><td>${c.tm?fmt(c.tm,0)+" W":"–"}</td><td>${c.lim!=null?fmt(c.lim,1)+(p.aclim==="SI"?" VLP":" VLA"):"–"}</td><td>${pill(c.ok,c.ok==null?"–":c.ok?"Cumple":"Supera")}</td></tr>`}).join("")}</table></div>
    <p class="muted">Valores de referencia de confort (no exigidos por la normativa argentina): temperatura 17–27 °C en trabajos sedentarios, humedad relativa 30–70 %, velocidad del aire ≤ 0,25 m/s.</p></div>`,
  checks:()=>stdChecks("puntos","puesto",p=>{const c=calcTermo(p),mi=[];if(!p.sector)mi.push("sector");if(c.ef==null&&c.teq==null)mi.push("lecturas");return mi}),
  autoText:()=>{const L=S.puntos.map(p=>({p,c:calcTermo(p)})),cal=L.filter(x=>x.c.ef!=null),bad=cal.filter(x=>!x.c.ok),fr=L.filter(x=>x.c.teq!=null);const t=[];
    if(cal.length)t.push(`Se evaluó la carga térmica en ${cal.length} puestos de ${S.meta.razon} mediante el índice TGBH (UNE-EN ISO 7243), ajustado por vestimenta y comparado con el valor límite según la tasa metabólica estimada, conforme a la Res. SRT 30/2023. Los valores de TGBH efectivo se encontraron entre ${fmt(Math.min(...cal.map(x=>x.c.ef)),1)} y ${fmt(Math.max(...cal.map(x=>x.c.ef)),1)} °C.`);
    const ok=cal.filter(x=>x.c.ok);if(ok.length)t.push(`No superan el valor límite: ${ok.map(x=>`N°${x.p.n} (${x.p.sector})`).join(", ")}.`);
    if(bad.length)t.push(`Superan el valor límite: ${bad.map(x=>`N°${x.p.n} (${x.p.sector}, TGBH ef. ${fmt(x.c.ef,1)} °C > ${fmt(x.c.lim,1)} °C${x.c.ft!=null?`; régimen sugerido ${fmt(x.c.ft*60,0)} min de trabajo por hora`:""})`).join("; ")}.`);
    if(fr.length)t.push(`En condiciones de frío (${fr.map(x=>`N°${x.p.n}: ${fmt(x.c.tbs,1)} °C, temperatura equivalente ${fmt(x.c.teq,1)} °C`).join("; ")}) se aplican las medidas de la Res. MTEySS 295/03 para estrés por frío.`);
    S.conclusiones=t.join("\n\n");const r=[];
    if(bad.length){r.push("Implementar controles de ingeniería (ventilación, extracción de calor, apantallamiento de fuentes radiantes, sombra en tareas al sol) y administrativos (régimen de trabajo/descanso en zona más fresca, rotación, reprogramación de tareas a horas de menor carga).");
      r.push("Asegurar hidratación con agua fresca cercana al puesto, plan de aclimatación (Res. SRT 30/23: 5 a 14 días) para personal nuevo o que retoma tras 7 días o más, y aptitud médica del personal expuesto.");r.push("Declarar a la ART al personal que supera el nivel de acción (VLA) bajo el agente carga térmica.")}
    if(fr.length)r.push("En exposición al frío: ropa de abrigo aislante seca, guantes según temperatura, refugios calefaccionados, pausas de recuperación y bebidas calientes.");
    r.push("Repetir la evaluación en la época del año y el horario más desfavorables.");S.recomendaciones=r.map(x=>"• "+x).join("\n")},
  exportFn:()=>{const m=S.meta;exportStd({titulo:"INFORME DE TERMOHIGROMETRÍA Y ESTRÉS TÉRMICO",norma:"Calor: Res. SRT 30/2023 (UNE-EN ISO 7243 · ISO 8996) · Frío: Res. MTEySS 295/03",fname:"Estres_Termico",
    datos:[["Razón social",m.razon],["Dirección",m.direccion],["Localidad",m.localidad],["Provincia",m.provincia],["C.P.",m.cp],["C.U.I.T.",m.cuit],[null,"Datos de la medición"],["Instrumento (marca, modelo, N° de serie)",`${m.instrumento} ${m.marca} ${m.modelo} – N° de serie: ${m.serie}`],["Fecha de calibración",fdate(m.fechaCal)],["Fecha de la medición",fdate(m.fechaMed)],["Hora de inicio / finalización",`${m.horaIni} – ${m.horaFin}`],["Metodología",m.metodologia],["Observaciones",m.observaciones],[null,"Documentación que se adjuntará"],["Certificado de calibración",m.docCert?"SI":"NO"],["Plano o croquis",m.docCroquis?"SI":"NO"]],
    headers:["N°","Sector","Puesto / tarea","T bulbo seco (°C)","T bulbo húmedo (°C)","T globo (°C)","HR (%)","Vel. aire (m/s)","Carga solar","TGBH (°C)","VAR ropa","TGBH ef. (°C)","Tasa metab. (W)","Aclimatado","Valor límite (°C)","Cumple SI/NO","Régimen trabajo / descanso","T equiv. frío (°C)"],
    rows:S.puntos.map(p=>{const c=calcTermo(p),l=p.lect[0]||{};const r1=v=>c.ef!=null&&isFinite(v)?Math.round(v*10)/10:"";
      return[+p.n||p.n,p.sector,[p.puesto,p.tarea].filter(Boolean).join(" – "),nOrNull(l.tbs),nOrNull(l.tbh),nOrNull(l.tg),c.hr,c.va,p.solar,r1(c.tgbh),num(p.var)||0,r1(c.ef),c.tm?Math.round(c.tm):"",p.aclim,c.lim!=null?`${fmt(c.lim,1)} (${p.aclim==="SI"?"VLP":"VLA"})`:"",yesno(c.ok??null),c.ft!=null?`${fmt(c.ft*60,0)}/${fmt(60-c.ft*60,0)} min`:(c.ok?"Continuo":""),c.teq!=null?Math.round(c.teq*10)/10:""]}),
    bad:S.puntos.map(p=>calcTermo(p).ok===false),widths:[5,12,14,8,8,8,6,7,7,7,6,8,8,8,10,7,11,8],
    info:"VLP = 56,7 − 11,5·log10(TM) (aclimatados); VLA = 59,9 − 14,1·log10(TM) (no aclimatados). TGBH ponderado en una hora y ajustado por vestimenta (VAR). Régimen de trabajo/descanso: ft = (A − B) / [(C − D) + (A − B)], con A = 33,7 °C, B = TGBH en descanso, C = TGBH ef. en el puesto, D = valor límite. Temperatura equivalente por viento según Res. MTEySS 295/03 (estrés por frío). Si el TGBH ponderado se basa en lecturas menores a 60 min, se indica en observaciones. "+S.puntos.filter(p=>p.obs).map(p=>`N°${p.n}: ${p.obs}`).join(" · ")})},
  prevOf:(p,c)=>({ef:c.ef??null}),clearOnRepeat:p=>{p.lect.forEach(r=>{r.tbs="";r.tbh="";r.tg=""});p.hr="";p.va="";p.tgbhDesc=""},clearOnDup:p=>{p.lect.forEach(r=>{r.tbs="";r.tbh="";r.tg=""})}
});

Object.assign(MODS,{ruido,vib,termo});
MODS.fuego.unidad="sectores";MODS.pat.unidad="tomas";
