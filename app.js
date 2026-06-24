/* ===========================================================
   Planning del Supervisor · Migraciones Entel
   Lógica 100% en el navegador (sin servidor).
   =========================================================== */

// ---- Mapeo de encabezados del CRM -> campos internos ----
const HEADER_MAP = {
  'idFicha_Venta':'id',
  'Fecha Venta':'fecha',
  'Hora_Venta':'hora',
  'DE_Campana_Netcall':'campana',
  'RV_Tipo_Ofrecimiento':'tipoOfrec',
  'RV_Tipo_Venta':'tipoVenta',
  'RV_Linea_Upselling':'lineaUpsell',
  'RV_Linea_Migrar':'lineaMigrar',
  'RV_Plan_Tarifario':'plan',
  'RV_Cargo_Fijo_Plan':'cargoFijo',
  'RV_Ganancia':'ganancia',
  'RV_Cant_Familia_Adicional':'cantFamilia',
  'RV_Tipo_Producto':'tipoProducto',
  'RV_Cant_Accesorios':'cantAcces',
  'DE_Monto_Disp_Finan_Equipos':'montoFinan',
  'Supervisor_Vendedor':'supervisor',
  'Documento_Vendedor':'docVendedor',
  'Nombre_Vendedor':'vendedor',
  'VBO_Estado_Venta_BO':'estado',
  'VBO_Sub_Estado_Venta_BO':'subEstado'
};

let ALL = [];          // todos los registros
let ADMIN = false;     // true cuando el perfil es Jefatura (puede desbloquear/editar)
let ROLE = null;       // 'sup' | 'jefe'
let SUPNAME = null;    // nombre del supervisor cuando ROLE==='sup'
let APPNAME = '';      // nombre del archivo / origen
let DASH_WIRED = false, ROLE_WIRED = false;
// crossRules: qué cuenta como cross-selling (configurable en la UI)
let STATE = { validEstados:null, sortKey:'ventas', sortDir:-1, pivotKey:'tot', pivotDir:-1 };

const nowStamp = ()=> new Date().toLocaleString('es-PE',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});

// ---------- Modo Jefatura (PIN) ----------
function jefaPin(){ return localStorage.getItem('jefaPin'); }
function toggleJefa(){
  if(ADMIN){ ADMIN=false; updateJefaUI(); render(); return; }
  let pin=jefaPin();
  if(!pin){
    const np=prompt('Configura el PIN de Jefatura (solo la primera vez).\nLo necesitarás para modificar objetivos bloqueados:');
    if(!np) return; localStorage.setItem('jefaPin',np); pin=np;
    alert('PIN de Jefatura configurado. Guárdalo: con él podrás editar objetivos y desbloquear registros.');
  }
  const entry=prompt('Ingresa el PIN de Jefatura:');
  if(entry===null) return;
  if(entry===pin){ ADMIN=true; updateJefaUI(); render(); }
  else alert('PIN incorrecto.');
}
function updateJefaUI(){
  const b=$('btnJefa'); if(!b) return;
  b.textContent = ADMIN?'🔓 Jefa activa · salir':'🔑 Modo Jefa';
  b.style.background = ADMIN?'var(--azul)':'#fff';
  b.style.color = ADMIN?'#fff':'var(--azul)';
}

// ---------- utilidades ----------
const $ = id => document.getElementById(id);
const num = v => { if(v==null) return 0; const n = parseFloat(String(v).replace(',','.').replace(/[^0-9.\-]/g,'')); return isNaN(n)?0:n; };
const clean = v => (v==null?'':String(v)).trim();
function hourOf(h){ // "18:49:54" -> {h:18,m:49,dec:18.83}
  const s = clean(h); const m = s.match(/(\d{1,2}):(\d{2})/);
  if(!m) return null;
  const H = +m[1], M = +m[2]; return {h:H,m:M,dec:H+M/60};
}
function titleCase(s){ s=clean(s); return s.replace(/\w\S*/g, t=>t.charAt(0).toUpperCase()+t.slice(1).toLowerCase()); }

// Clasificación por RV_Tipo_Ofrecimiento:
//   CROSS-SELLING            -> cross-selling
//   REGULAR / MULTILINEA / … -> migración (todo lo demás)
function isCross(r){ return clean(r.tipoOfrec).toUpperCase().indexOf('CROSS') >= 0; }
function isMig(r){ return !isCross(r); }

// =========================================================
//  CARGA DE ARCHIVO
// =========================================================
const drop = $('drop');
['dragover','dragenter'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('drag');}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('drag');}));
drop.addEventListener('drop',ev=>{ if(ev.dataTransfer.files.length) handleFile(ev.dataTransfer.files[0]); });
$('file').addEventListener('change',ev=>{ if(ev.target.files.length) handleFile(ev.target.files[0]); });

// demo
if(!window.__noSample && window.SAMPLE_ROWS){ $('sampleBox').classList.remove('hidden'); }
$('btnSample').addEventListener('click',()=>{
  ALL = window.SAMPLE_ROWS.map(normalizeRecord);
  afterLoad('Datos de ejemplo');
});
// auto-demo para pruebas: #demo (pantalla de perfil), #demojefe (Jefatura), #demosup (un supervisor)
if(location.hash.indexOf('#demo')===0 && window.SAMPLE_ROWS){
  window.addEventListener('load',()=>{
    ALL = window.SAMPLE_ROWS.map(normalizeRecord); afterLoad('Datos de ejemplo');
    if(location.hash==='#demojefe'){ ROLE='jefe'; SUPNAME=null; ADMIN=true; initDashboard(); }
    else if(location.hash==='#demosup'){ const sups=[...new Set(ALL.map(r=>r.supervisor).filter(Boolean))].sort(); ROLE='sup'; SUPNAME=sups[1]||sups[0]; ADMIN=false; initDashboard(); }
  });
}

function setStatus(html){ $('loadStatus').innerHTML = html; }

function handleFile(file){
  setStatus('<div class="spinner"></div><div class="muted">Leyendo <b>'+file.name+'</b> ('+(file.size/1048576).toFixed(1)+' MB). Si el archivo es grande puede tardar ~1 minuto…</div>');
  const reader = new FileReader();
  reader.onload = e=>{
    try{
      const data = new Uint8Array(e.target.result);
      // 'dense:true' permite leer archivos grandes (el modo normal deja la hoja sin cargar)
      let wb = XLSX.read(data,{type:'array',cellDates:false,raw:true,dense:true});
      let ws = wb.Sheets[wb.SheetNames[0]];
      if(!ws){
        setStatus('<span style="color:var(--bad)">No pude cargar la hoja de datos. El archivo puede ser demasiado grande para el navegador; intenta con el export de un día o de un rango más corto.</span>');
        return;
      }
      const rows = XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});
      if(!rows.length){ setStatus('<span style="color:var(--bad)">La hoja "'+wb.SheetNames[0]+'" no tiene filas de datos.</span>'); return; }
      const headers = rows[0].map(clean);
      // índice de cada campo
      const idx = {};
      headers.forEach((h,i)=>{ if(HEADER_MAP[h]) idx[HEADER_MAP[h]] = i; });
      if(idx.vendedor==null || idx.fecha==null){
        setStatus('<span style="color:var(--bad)">No encontré las columnas esperadas del CRM (Nombre_Vendedor / Fecha Venta). ¿Es el volcado correcto?</span>');
        return;
      }
      const recs = [];
      for(let i=1;i<rows.length;i++){
        const row = rows[i]; if(!row || row.length===0) continue;
        const rec = {};
        for(const f in idx) rec[f] = row[idx[f]];
        if(clean(rec.vendedor)==='' && clean(rec.fecha)==='') continue;
        recs.push(normalizeRecord(rec));
      }
      ALL = recs;
      afterLoad(file.name);
    }catch(err){
      console.error(err);
      setStatus('<span style="color:var(--bad)">No pude leer el archivo: '+err.message+'</span>');
    }
  };
  reader.onerror = ()=> setStatus('<span style="color:var(--bad)">Error leyendo el archivo.</span>');
  reader.readAsArrayBuffer(file);
}

function normalizeRecord(r){
  const ho = hourOf(r.hora);
  return {
    id: clean(r.id),
    fecha: clean(r.fecha).slice(0,10),
    horaStr: clean(r.hora).slice(0,8),
    h: ho? ho.h : null,
    hdec: ho? ho.dec : null,
    campana: clean(r.campana),
    tipoOfrec: clean(r.tipoOfrec),
    tipoVenta: clean(r.tipoVenta),
    lineaUpsell: clean(r.lineaUpsell),
    lineaMigrar: clean(r.lineaMigrar),
    plan: clean(r.plan),
    cargoFijo: num(r.cargoFijo),
    ganancia: num(r.ganancia),
    cantFamilia: num(r.cantFamilia),
    tipoProducto: clean(r.tipoProducto),
    cantAcces: num(r.cantAcces),
    montoFinan: num(r.montoFinan),
    docVendedor: clean(r.docVendedor),
    supervisor: titleCase(r.supervisor),
    vendedor: titleCase(r.vendedor),
    estado: clean(r.estado).toUpperCase(),
    subEstado: clean(r.subEstado)
  };
}

// =========================================================
//  DESPUÉS DE CARGAR -> preparar filtros y pintar
// =========================================================
function afterLoad(name){
  // Tomar SOLO la campaña MIGRACION REGULAR (DE_Campana_Netcall)
  ALL = ALL.filter(r=> clean(r.campana).toUpperCase()==='MIGRACION REGULAR');
  if(!ALL.length){
    setStatus('<span style="color:var(--bad)">El archivo no tiene filas de la campaña <b>MIGRACION REGULAR</b> (columna DE_Campana_Netcall). Revisa que sea el volcado correcto.</span>');
    return;
  }
  APPNAME = name;
  showRoleScreen();
}

// ---- Selección de perfil ----
function showRoleScreen(){
  const sups=[...new Set(ALL.map(r=>r.supervisor).filter(Boolean))].sort();
  $('roleSup').innerHTML='<option value="">— elige tu nombre —</option>'+sups.map(s=>`<option value="${s}">${s}</option>`).join('');
  $('loader').classList.add('hidden');
  $('dash').classList.add('hidden');
  $('rolescreen').classList.remove('hidden');
  $('roleMsg').textContent='';
  if(!ROLE_WIRED){
    $('enterSup').addEventListener('click',()=>{
      const v=$('roleSup').value;
      if(!v){ $('roleMsg').textContent='Selecciona tu nombre de supervisor.'; return; }
      ROLE='sup'; SUPNAME=v; ADMIN=false; initDashboard();
    });
    $('enterJefe').addEventListener('click',()=>{
      let pin=jefaPin();
      if(!pin){ const np=prompt('Configura el PIN de Jefatura (solo la primera vez). Guárdalo:'); if(!np) return; localStorage.setItem('jefaPin',np); pin=np; }
      const entry=prompt('Ingresa el PIN de Jefatura:');
      if(entry===null) return;
      if(entry!==pin){ $('roleMsg').textContent='PIN incorrecto.'; return; }
      ROLE='jefe'; SUPNAME=null; ADMIN=true; initDashboard();
    });
    ROLE_WIRED=true;
  }
}

function initDashboard(){
  $('rolescreen').classList.add('hidden');
  $('loader').classList.add('hidden');
  $('dash').classList.remove('hidden');
  const scope = ROLE==='jefe' ? '👑 Jefatura · todos los equipos' : '🧑‍💼 Supervisor · '+SUPNAME;
  $('hdrInfo').textContent = APPNAME + ' · ' + ALL.length.toLocaleString('es-PE') + ' registros · ' + scope;

  // Fechas — multi-selección
  STATE.fechas = null;
  buildFechaPanel();

  // Estados (válidos por defecto: los que contienen VALID)
  const estados = {};
  ALL.forEach(r=>{ if(r.estado) estados[r.estado]=(estados[r.estado]||0)+1; });
  const eb = $('estadosBox'); eb.innerHTML='';
  const validDefault = Object.keys(estados).filter(e=>e.indexOf('VALID')>=0);
  STATE.validEstados = new Set(validDefault.length?validDefault:Object.keys(estados));
  Object.entries(estados).sort((a,b)=>b[1]-a[1]).forEach(([e,c])=>{
    const lab=document.createElement('label');
    const ck=document.createElement('input'); ck.type='checkbox'; ck.checked=STATE.validEstados.has(e);
    ck.addEventListener('change',()=>{ ck.checked?STATE.validEstados.add(e):STATE.validEstados.delete(e); render(); });
    lab.appendChild(ck); lab.appendChild(document.createTextNode(' '+e+' ('+c.toLocaleString('es-PE')+')'));
    eb.appendChild(lab);
  });

  buildSupOptions();
  buildCorteOptions();

  if(!DASH_WIRED){
    $('fFechaBtn').addEventListener('click',e=>{ e.stopPropagation(); $('fFechaPanel').classList.toggle('hidden'); });
    document.addEventListener('click',e=>{ if(!$('fFechaDrop').contains(e.target)) $('fFechaPanel').classList.add('hidden'); });
    ['fSup','fCorte','fIni','fFin','fJornada','fMeta'].forEach(id=> $(id).addEventListener('change',()=>{ buildCorteOptions(); render(); }));
    $('btnReset').addEventListener('click',()=>{ location.reload(); });
    $('addTaskBtn').addEventListener('click',addTask);
    $('newTask').addEventListener('keydown',e=>{ if(e.key==='Enter') addTask(); });
    $('mAddBtn').addEventListener('click',addMonitoreo);
    $('btnJefa').addEventListener('click',showRoleScreen);   // ahora: "Cambiar perfil"
    DASH_WIRED=true;
  }
  updateProfileUI();
  render();
}

function updateProfileUI(){
  const b=$('btnJefa'); if(!b) return;
  b.textContent='🔄 Cambiar perfil';
  b.style.background='#fff'; b.style.color='var(--azul)';
}

// ---- Fechas multi-selección ----
function buildFechaPanel(){
  const cnt={}; ALL.forEach(r=>{ if(r.fecha) cnt[r.fecha]=(cnt[r.fecha]||0)+1; });
  const fechas=Object.keys(cnt).sort();
  if(!STATE.fechas || !STATE.fechas.size){
    const top=fechas.reduce((a,b)=> cnt[b]>=cnt[a]?b:a, fechas[0]);
    STATE.fechas=new Set([top]);
  }
  const panel=$('fFechaPanel');
  panel.innerHTML='<div class="ddtools"><a href="#" id="fTodas">Todas</a><a href="#" id="fUna">Solo la última</a></div>'+
    fechas.map(f=>`<label><input type="checkbox" value="${f}" ${STATE.fechas.has(f)?'checked':''}> ${f} (${cnt[f]})</label>`).join('');
  panel.querySelectorAll('input[type=checkbox]').forEach(ck=>ck.addEventListener('change',()=>{
    if(ck.checked) STATE.fechas.add(ck.value); else STATE.fechas.delete(ck.value);
    if(!STATE.fechas.size){ STATE.fechas.add(ck.value); ck.checked=true; }   // no permitir vacío
    afterFechaChange();
  }));
  $('fTodas').addEventListener('click',e=>{ e.preventDefault(); STATE.fechas=new Set(fechas); buildFechaPanel(); afterFechaChange(); });
  $('fUna').addEventListener('click',e=>{ e.preventDefault(); STATE.fechas=new Set([fechas[fechas.length-1]]); buildFechaPanel(); afterFechaChange(); });
  updateFechaBtn();
}
function afterFechaChange(){ updateFechaBtn(); buildSupOptions(); buildCorteOptions(); render(); }
function updateFechaBtn(){
  const n=STATE.fechas.size;
  $('fFechaBtn').textContent = n===1 ? [...STATE.fechas][0] : (n+' fechas seleccionadas');
}
function selFechas(){ return [...STATE.fechas]; }
function primaryFecha(){ return selFechas().sort().slice(-1)[0]; }

function buildSupOptions(){
  const sel = $('fSup');
  // Perfil Supervisor: fijo a su equipo, sin "Todos" ni otros supervisores
  if(ROLE==='sup'){
    sel.innerHTML=`<option value="${SUPNAME}">${SUPNAME}</option>`;
    sel.value=SUPNAME; sel.disabled=true;
    return;
  }
  sel.disabled=false;
  const fset=STATE.fechas;
  const sups = [...new Set(ALL.filter(r=>fset.has(r.fecha)).map(r=>r.supervisor).filter(Boolean))].sort();
  const prev = sel.value;
  sel.innerHTML='<option value="__ALL__">— Todos los supervisores —</option>';
  sups.forEach(s=>{ const o=document.createElement('option'); o.value=s; o.textContent=s; sel.appendChild(o); });
  if(prev && [...sel.options].some(o=>o.value===prev)) sel.value=prev;
}

function buildCorteOptions(){
  const ini = +$('fIni').value, fin=+$('fFin').value;
  const fset=STATE.fechas;
  const maxHp = Math.max(ini, ...ALL.filter(r=>fset.has(r.fecha)&&r.h!=null).map(r=>r.h));
  const top = Math.max(fin, isFinite(maxHp)?maxHp+1:fin);   // incluye horas extra si las hay
  const sel=$('fCorte'); const prev=sel.value; sel.innerHTML='';
  for(let h=ini+1;h<=top;h++){
    const o=document.createElement('option'); o.value=h;
    o.textContent=String(h).padStart(2,'0')+':00'+(h===fin?' · cierre':'')+(h>fin?' (extra)':'');
    sel.appendChild(o);
  }
  // por defecto: cubre hasta la última hora con ventas (incluye horas extra en histórico)
  if(prev && [...sel.options].some(o=>o.value===prev)){ sel.value=prev; }
  else sel.value = String(top);
}

// =========================================================
//  CÁLCULO PRINCIPAL
// =========================================================
function compute(){
  const fset=STATE.fechas, fecha=primaryFecha(), sup=$('fSup').value;
  const multiFecha = fset.size>1;
  const ini=+$('fIni').value, fin=+$('fFin').value, corte=+$('fCorte').value;
  const jornadaH = Math.max(1, +$('fJornada').value||8);   // horas laborables de la jornada base
  const meta=+$('fMeta').value||1;
  const elapsedFranja = Math.max(0.01, Math.min(corte,fin)-ini);  // horas de franja transcurridas
  const elapsedEff = Math.max(0.5, Math.min(elapsedFranja, jornadaH)); // tope a la jornada (8 h)

  const obj = getObjetivos(sup,fecha);   // objetivos por agente (bloqueables)

  // ¿la venta cae en la franja base (jornada) o en horas extra?
  const inFranja = r => r.h!=null && r.h>=ini && r.h<fin;

  let rows = ALL.filter(r=>fset.has(r.fecha));
  if(sup!=='__ALL__') rows = rows.filter(r=>r.supervisor===sup);

  const isValid = r=> STATE.validEstados.has(r.estado);
  const valid = rows.filter(isValid);
  const validUpTo = valid.filter(r=> r.hdec!=null && r.hdec < corte);

  // agrupar por agente: ventas (jornada base) vs extra (horas extra)
  const ag = {};
  valid.forEach(r=>{
    if(r.hdec==null || r.hdec >= corte) return;          // solo hasta la hora de corte
    const k=r.vendedor||'(sin nombre)';
    if(!ag[k]) ag[k]={vendedor:k,supervisor:r.supervisor,ventas:0,extra:0,tot:0,mig:0,cross:0,crossReg:0,crossPort:0,crossOtro:0,ganancia:0,ultima:null};
    const a=ag[k];
    if(inFranja(r)) a.ventas++; else a.extra++;           // base vs horas extra
    a.tot++;
    if(isCross(r)){                                       // CROSS-SELLING: desglosar por tipo de venta
      a.cross++;
      const tv=clean(r.tipoVenta).toUpperCase();
      if(tv.indexOf('PORTAB')>=0) a.crossPort++;
      else if(tv.indexOf('REGULAR')>=0) a.crossReg++;
      else a.crossOtro++;
    } else a.mig++;
    a.ganancia+=r.ganancia;
    if(a.ultima==null || r.hdec>a.ultima) a.ultima=r.hdec;
  });
  let agents = Object.values(ag);
  // ritmo y proyección — SIEMPRE sobre la jornada base (comparable entre agentes)
  agents.forEach(a=>{
    a.meta = (obj.locked && obj.metas[a.vendedor]!=null) ? obj.metas[a.vendedor] : meta;
    a.ritmo = a.ventas/elapsedEff;                         // SPH: ventas por hora efectiva (sobre 8 h, sin refrigerio)
    a.proy = Math.round(a.ventas * jornadaH / elapsedEff); // proyección a la jornada de 8 h ( = SPH × 8 )
    if(a.proy < a.ventas) a.proy = a.ventas;
    a.pct = a.meta? Math.round(a.ventas/a.meta*100) : 0;   // % meta sobre jornada base
    a.pctProy = a.meta? Math.round(a.proy/a.meta*100) : 0;
    a.cumple = a.ventas >= a.meta;
  });
  const teamMeta = agents.reduce((s,a)=>s+a.meta,0);
  // cuartiles por ventas de JORNADA (desc) -> comparación justa
  agents.sort((x,y)=> y.ventas-x.ventas || x.vendedor.localeCompare(y.vendedor));
  const n=agents.length;
  agents.forEach((a,i)=>{ a.q = n? Math.min(4, Math.floor(i*4/n)+1) : 1; });

  // ---- SPD y acumulado del período (equipo filtrado, todas las fechas) ----
  // SPD se mide sobre ventas de JORNADA base (no incluye horas extra) para que sea comparable.
  const teamValidAll = ALL.filter(r=> (sup==='__ALL__'||r.supervisor===sup) && isValid(r));
  const byDate={};
  teamValidAll.forEach(r=>{ const f=r.fecha||'(s/f)'; if(!byDate[f]) byDate[f]={base:0,extra:0,ags:new Set()}; if(inFranja(r)) byDate[f].base++; else byDate[f].extra++; if(r.vendedor) byDate[f].ags.add(r.vendedor); });
  let acumVentas=0, acumExtra=0, agentDays=0, dias=0;
  Object.values(byDate).forEach(d=>{ acumVentas+=d.base; acumExtra+=d.extra; agentDays+=d.ags.size; dias++; });
  const ventasDia = agents.reduce((s,a)=>s+a.ventas,0);
  const extraDia = agents.reduce((s,a)=>s+a.extra,0);
  const spdDia = n? ventasDia/n : 0;
  const spdAcum = agentDays? acumVentas/agentDays : 0;

  return {fecha,fset,multiFecha,sup,ini,fin,corte,jornadaH,meta,totalH:(fin-ini),elapsed:elapsedFranja,rows,valid,validUpTo,agents,n,obj,teamMeta,
          ventasDia,extraDia,spdDia,spdAcum,acumVentas,acumExtra,dias,agentDays};
}

// =========================================================
//  RENDER
// =========================================================
function render(){
  const C = compute();
  renderKpis(C); renderSPD(C); renderObjetivos(C); renderChart(C); renderQuartiles(C); renderTable(C); renderPivot(C); renderCross(C); renderMonitoreo(C); renderTasks(C);
  $('chipEquipo').textContent = C.sup==='__ALL__'?'Todos':C.sup;
}

function renderKpis(C){
  const ventas=C.agents.reduce((s,a)=>s+a.ventas,0);
  const extra=C.agents.reduce((s,a)=>s+a.extra,0);
  const mig=C.agents.reduce((s,a)=>s+a.mig,0);
  const cross=C.agents.reduce((s,a)=>s+a.cross,0);
  const proy=C.agents.reduce((s,a)=>s+a.proy,0);
  const tot=ventas+extra;
  const metaEq=C.teamMeta || (C.meta*C.n);
  const pctEq=metaEq?Math.round(ventas/metaEq*100):0;
  const pctProyEq=metaEq?Math.round(proy/metaEq*100):0;
  const col = p => p>=100?'var(--ok)':p>=70?'var(--warn)':'var(--bad)';
  const pc = n => tot?Math.round(n/tot*100):0;
  const fr = String(C.ini).padStart(2,'0')+'–'+String(C.fin).padStart(2,'0');
  const k=[
    {l:'Agentes activos',v:C.n,x:'<span class="muted">en el equipo</span>'},
    {l:'Ventas en jornada ('+fr+')',v:ventas.toLocaleString('es-PE'),x:'<span class="muted">base comparable</span>'},
    {l:'Horas extra (fuera de franja)',v:extra.toLocaleString('es-PE'),x:'<span class="muted">ventas adicionales · total '+tot.toLocaleString('es-PE')+'</span>'},
    {l:'Migraciones',v:mig.toLocaleString('es-PE'),x:'<span class="muted">'+pc(mig)+'% del total</span>'},
    {l:'Cross-selling',v:cross.toLocaleString('es-PE'),x:'<span class="muted">'+pc(cross)+'% del total</span>'},
    {l:'Avance vs meta equipo',v:pctEq+'%',x:'<span style="color:'+col(pctEq)+'">'+ventas+' / '+metaEq+'</span>'},
    {l:'Proyección cierre (jornada)',v:proy.toLocaleString('es-PE'),x:'<span style="color:'+col(pctProyEq)+'">'+pctProyEq+'% de la meta</span>'},
  ];
  $('kpis').innerHTML = k.map(o=>`<div class="kpi"><div class="l">${o.l}</div><div class="v">${o.v}</div><div class="x">${o.x}</div></div>`).join('');
}

function renderSPD(C){
  const ambito = C.sup==='__ALL__'?'todos los equipos':C.sup;
  const lblDia = C.multiFecha ? 'SPD de la selección' : 'SPD del día (jornada)';
  const k=[
    {l:lblDia,v:C.spdDia.toFixed(1),x:'<span class="muted">'+C.ventasDia+' base / '+C.n+' agentes</span>'},
    {l:'SPD acumulado (jornada)',v:C.spdAcum.toFixed(1),x:'<span class="muted">promedio en '+C.dias+' día(s)</span>'},
    {l:'Horas extra'+(C.multiFecha?'':' del día'),v:C.extraDia.toLocaleString('es-PE'),x:'<span class="muted">no entran al SPD</span>'},
    {l:'Acumulado período (base)',v:C.acumVentas.toLocaleString('es-PE'),x:'<span class="muted">+'+C.acumExtra.toLocaleString('es-PE')+' en horas extra</span>'},
  ];
  $('spdKpis').innerHTML = k.map(o=>`<div class="kpi"><div class="l">${o.l}</div><div class="v">${o.v}</div><div class="x">${o.x}</div></div>`).join('');
}

function renderChart(C){
  // ventas válidas por hora; incluye horas extra (antes de la franja o después del cierre)
  let minH=C.ini, maxH=C.fin-1;
  C.valid.forEach(r=>{ if(r.h!=null){ if(r.h<minH)minH=r.h; if(r.h>maxH)maxH=r.h; } });
  const buckets={}; for(let h=minH;h<=maxH;h++) buckets[h]=0;
  C.valid.forEach(r=>{ if(r.h!=null && r.h>=minH && r.h<=maxH) buckets[r.h]++; });
  const hours=Object.keys(buckets).map(Number).sort((a,b)=>a-b);
  const max=Math.max(1,...hours.map(h=>buckets[h]));
  const W=Math.max(560, hours.length*44), H=240, pad=28, bw=Math.min(34,(W-2*pad)/hours.length-8);
  let cum=0; const total=hours.reduce((s,h)=>s+buckets[h],0);
  let bars='',line='',pts=[];
  hours.forEach((h,i)=>{
    const x=pad+i*((W-2*pad)/hours.length);
    const bh=(buckets[h]/max)*(H-2*pad);
    const future = h>=C.corte;
    const afterClose = h>=C.fin;        // después de las 18:00 -> verde coral
    const beforeOpen = h<C.ini;         // antes de la franja -> morado
    const fill = future?'#cbd5e6':(afterClose?'var(--coral)':(beforeOpen?'var(--extra)':'var(--azul2)'));
    bars+=`<rect x="${x}" y="${H-pad-bh}" width="${bw}" height="${bh}" rx="3" fill="${fill}"></rect>`;
    bars+=`<text x="${x+bw/2}" y="${H-pad-bh-4}" font-size="10" text-anchor="middle" fill="#334">${buckets[h]||''}</text>`;
    bars+=`<text x="${x+bw/2}" y="${H-pad+12}" font-size="9" text-anchor="middle" fill="#889">${String(h).padStart(2,'0')}h</text>`;
    cum+=buckets[h];
    const cx=x+bw/2, cy=H-pad-(cum/Math.max(1,total))*(H-2*pad);
    pts.push([cx,cy]);
  });
  line='<polyline fill="none" stroke="#f59e0b" stroke-width="2" points="'+pts.map(p=>p.join(',')).join(' ')+'"></polyline>';
  const dots=pts.map(p=>`<circle cx="${p[0]}" cy="${p[1]}" r="2.5" fill="#f59e0b"></circle>`).join('');
  // línea de corte
  const corteIdx=hours.indexOf(C.corte);
  let corteLine='';
  if(corteIdx>=0){ const x=pad+corteIdx*((W-2*pad)/hours.length); corteLine=`<line x1="${x}" y1="${pad-6}" x2="${x}" y2="${H-pad}" stroke="#dc2626" stroke-dasharray="4 3" stroke-width="1.5"></line><text x="${x+3}" y="${pad-2}" font-size="9" fill="#dc2626">ahora</text>`; }
  $('chart').innerHTML=`<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet">${bars}${corteLine}${line}${dots}</svg>`;
}

function renderQuartiles(C){
  const colors={1:'var(--q1)',2:'var(--q2)',3:'var(--q3)',4:'var(--q4)'};
  const names={1:'Q1 · Top',2:'Q2',3:'Q3',4:'Q4 · Cola'};
  const totalV=C.agents.reduce((s,a)=>s+a.ventas,0)||1;
  let html='';
  for(let q=1;q<=4;q++){
    const list=C.agents.filter(a=>a.q===q);
    const vs=list.map(a=>a.ventas);
    const ventas=vs.reduce((s,v)=>s+v,0);
    const avg=list.length?ventas/list.length:0;
    const min=vs.length?Math.min(...vs):0, max=vs.length?Math.max(...vs):0;
    const share=Math.round(ventas/totalV*100);
    const nombres = list.map(a=>{
      const nm = a.vendedor.split(' ').slice(0,2).join(' ');
      const ex = a.extra? ` <span style="opacity:.85">(+${a.extra})</span>`:'';
      return `<li><span>${nm}</span><b>${a.ventas}${ex}</b></li>`;
    }).join('') || '<li class="muted2">— sin agentes —</li>';
    html+=`<div class="qcard" style="background:${colors[q]}">
      <div class="qhead">${names[q]}</div>
      <div class="qbig">${list.length}<span> agentes</span></div>
      <div class="qrow"><span>Promedio</span><b>${avg.toFixed(1)}</b></div>
      <div class="qrow"><span>Rango</span><b>${min}–${max}</b></div>
      <div class="qrow"><span>Ventas · %</span><b>${ventas} · ${share}%</b></div>
      <div class="qbar"><div style="width:${share}%"></div></div>
      <ul class="qnames">${nombres}</ul>
    </div>`;
  }
  $('quartiles').innerHTML=html;
}

const COLS=[
  {k:'q',t:'Q'},{k:'vendedor',t:'Agente'},{k:'ventas',t:'Jornada'},{k:'extra',t:'Extra'},{k:'mig',t:'Migrac.'},
  {k:'cross',t:'Cross'},{k:'ultima',t:'Últ. venta'},{k:'ritmo',t:'SPH'},
  {k:'proy',t:'Proy. cierre'},{k:'pct',t:'% Meta'}
];
function renderTable(C){
  $('thead').innerHTML=COLS.map(c=>`<th data-k="${c.k}">${c.t}</th>`).join('');
  $('thead').querySelectorAll('th').forEach(th=>th.addEventListener('click',()=>{
    const k=th.dataset.k; STATE.sortDir = (STATE.sortKey===k)? -STATE.sortDir : -1; STATE.sortKey=k; renderTable(C);
  }));
  const qcol={1:'var(--q1)',2:'var(--q2)',3:'var(--q3)',4:'var(--q4)'};
  let ag=[...C.agents];
  const k=STATE.sortKey;
  ag.sort((a,b)=>{ let x=a[k],y=b[k]; if(typeof x==='string'){return STATE.sortDir*x.localeCompare(y);} return STATE.sortDir*((x||0)-(y||0)); });
  const fmtH=d=> d==null?'—':String(Math.floor(d)).padStart(2,'0')+':'+String(Math.round((d%1)*60)).padStart(2,'0');
  const col=p=>p>=100?'var(--ok)':p>=70?'var(--warn)':'var(--bad)';
  $('tbody').innerHTML=ag.map(a=>{
    const pctW=Math.min(100,a.pct);
    return `<tr>
      <td><span class="tag" style="background:${qcol[a.q]}">Q${a.q}</span></td>
      <td>${a.vendedor}</td>
      <td><b>${a.ventas}</b></td>
      <td>${a.extra?('<span class="pill" style="color:var(--coral)">+'+a.extra+'</span>'):''}</td>
      <td>${a.mig||''}</td>
      <td>${a.cross?('<span class="pill" style="color:var(--azul)">'+a.cross+'</span>'):''}</td>
      <td>${fmtH(a.ultima)}</td>
      <td>${a.ritmo.toFixed(1)}</td>
      <td><b style="color:var(--azul)">${a.proy}</b></td>
      <td><div style="display:flex;align-items:center;gap:6px"><div class="bar-wrap"><div class="bar-fill" style="width:${pctW}%;background:${col(a.pct)}"></div></div><span style="color:${col(a.pct)};font-weight:700">${a.pct}%</span></div></td>
    </tr>`;
  }).join('');
}

function renderPivot(C){
  // tabla dinámica: filas = supervisor, columnas = hora, valor = cantidad de ventas válidas
  const sup={}, hoursSet=new Set();
  C.valid.forEach(r=>{ if(r.h==null) return; hoursSet.add(r.h); const s=r.supervisor||'(sin supervisor)'; (sup[s]=sup[s]||{})[r.h]=(sup[s][r.h]||0)+1; });
  const hours=[...hoursSet].sort((a,b)=>a-b);
  const hc = h => (h>=C.fin?'color:var(--coral)':(h<C.ini?'color:var(--extra)':''));
  // filas con total
  let filas=Object.keys(sup).map(s=>{
    let tot=0; hours.forEach(h=>tot+=(sup[s][h]||0));
    return {s, tot};
  });
  $('pivotInfo').textContent = filas.length+' supervisores · '+hours.length+' horas · '+C.fset.size+' fecha(s)';
  if(!filas.length){ $('pivotHead').innerHTML=''; $('pivotBody').innerHTML='<tr><td class="muted">Sin datos para la selección.</td></tr>'; return; }
  // ordenar
  const dir=STATE.pivotDir||-1;
  if(STATE.pivotKey==='sup') filas.sort((a,b)=> dir*a.s.localeCompare(b.s));
  else filas.sort((a,b)=> dir*(a.tot-b.tot) || a.s.localeCompare(b.s));   // por Total (default desc)
  const arrow = STATE.pivotKey==='sup' ? '' : (dir<0?' ▼':' ▲');
  $('pivotHead').innerHTML='<tr><th data-pk="sup" style="position:sticky;left:0;background:#fafbfe;cursor:pointer">Supervisor</th>'+
    hours.map(h=>`<th style="text-align:center;${hc(h)}">${String(h).padStart(2,'0')}h</th>`).join('')+
    `<th data-pk="tot" style="text-align:center;cursor:pointer" title="Clic para ordenar">Total${arrow}</th></tr>`;
  let body=''; const colTot={}; let grand=0;
  filas.forEach(f=>{
    const s=f.s;
    const cells=hours.map(h=>{ const v=sup[s][h]||0; colTot[h]=(colTot[h]||0)+v; return `<td style="text-align:center;${v&&hc(h)}">${v||''}</td>`; }).join('');
    grand+=f.tot;
    body+=`<tr><td style="position:sticky;left:0;background:#fff;font-weight:600">${s}</td>${cells}<td style="text-align:center;font-weight:700">${f.tot}</td></tr>`;
  });
  body+=`<tr style="background:#f1f5fd;font-weight:700"><td style="position:sticky;left:0;background:#f1f5fd">Total</td>`+
    hours.map(h=>`<td style="text-align:center;${hc(h)}">${colTot[h]||0}</td>`).join('')+
    `<td style="text-align:center">${grand}</td></tr>`;
  $('pivotBody').innerHTML=body;
  // ordenamiento por clic en encabezado
  $('pivotHead').querySelectorAll('th[data-pk]').forEach(th=>th.addEventListener('click',()=>{
    const k=th.dataset.pk;
    if(STATE.pivotKey===k) STATE.pivotDir=-(STATE.pivotDir||-1);
    else { STATE.pivotKey=k; STATE.pivotDir=(k==='sup'?1:-1); }
    renderPivot(C);
  }));
}

function renderCross(C){
  // respeta el filtro de supervisor: un supervisor -> su equipo; "Todos" -> todos los agentes con cross
  const cross=C.agents.filter(a=>a.cross>0).sort((a,b)=>b.cross-a.cross || a.vendedor.localeCompare(b.vendedor));
  const total=C.agents.reduce((s,a)=>s+a.cross,0);
  const totalVentas=C.agents.reduce((s,a)=>s+a.ventas,0);
  const showSup = C.sup==='__ALL__';
  $('crossKpi').innerHTML=`<span class="chip">${cross.length} agentes con cross-selling</span> &nbsp; <span class="muted">${total} ventas cruzadas (${totalVentas?Math.round(total/totalVentas*100):0}% del total)${showSup?' · todos los equipos':' · '+C.sup}</span>`;
  $('crossHead').innerHTML = '<th>Agente</th>'+(showSup?'<th>Supervisor</th>':'')+'<th>Cross</th><th>Regular</th><th>Portab.</th><th>Otro</th>';
  const colspan = showSup?6:5;
  $('crossBody').innerHTML = cross.length? cross.map(a=>{
      const reg=a.crossReg||0, por=a.crossPort||0, otr=a.crossOtro||0;
      return `<tr><td>${a.vendedor}</td>${showSup?('<td class="muted">'+(a.supervisor||'')+'</td>'):''}<td><b>${a.cross}</b></td><td>${reg||''}</td><td>${por||''}</td><td>${otr||''}</td></tr>`;
    }).join('')
    : `<tr><td colspan="${colspan}" class="muted">Nadie registró cross-selling en la selección.</td></tr>`;
}

// =========================================================
//  OBJETIVOS DEL DÍA POR AGENTE (bloqueables · localStorage)
// =========================================================
function objKey(sup,fecha){ return 'objetivos::'+sup+'::'+fecha; }
function getObjetivos(sup,fecha){
  if(sup==='__ALL__') return {locked:false,setAt:null,metas:{}};
  try{ return JSON.parse(localStorage.getItem(objKey(sup,fecha))||'null') || {locked:false,setAt:null,metas:{}}; }
  catch(e){ return {locked:false,setAt:null,metas:{}}; }
}
function saveObjetivos(sup,fecha,o){ localStorage.setItem(objKey(sup,fecha),JSON.stringify(o)); }
function rosterFor(sup){
  const set=new Set();
  ALL.forEach(r=>{ if(r.supervisor===sup && r.vendedor) set.add(r.vendedor); });
  return [...set].sort();
}
function renderObjetivos(C){
  const body=$('objBody'), state=$('objLockState');
  if(C.sup==='__ALL__'){
    state.innerHTML='';
    body.innerHTML='<p class="muted">Selecciona tu <b>supervisor</b> en los filtros de arriba para fijar y bloquear los objetivos del día por agente.</p>';
    return;
  }
  const roster=rosterFor(C.sup);
  const vmap={}; C.agents.forEach(a=>vmap[a.vendedor]=a);
  const obj=C.obj;
  const def=+$('fMeta').value||6;

  function editorHtml(saveLabel){
    const rows=roster.map(v=>{
      const val=obj.metas[v]!=null?obj.metas[v]:def;
      const ventas=vmap[v]?vmap[v].ventas:0;
      return `<tr><td>${v}</td><td><input type="number" min="0" data-ag="${encodeURIComponent(v)}" value="${val}" style="width:90px;padding:6px;border:1px solid #cbd5e6;border-radius:8px"></td><td class="muted">${ventas}</td></tr>`;
    }).join('');
    return '<div class="scroll" style="max-height:340px"><table><thead><tr><th>Agente</th><th>Objetivo del día</th><th>Ventas (al corte)</th></tr></thead><tbody>'+rows+'</tbody></table></div>'+
      '<div style="margin-top:10px;display:flex;gap:10px;align-items:center;flex-wrap:wrap"><button class="btn" id="lockBtn">'+saveLabel+'</button><span class="muted" style="font-size:12px">Queda registrado con fecha y hora.</span></div>';
  }
  function wireSave(askConfirm){
    $('lockBtn').addEventListener('click',()=>{
      if(askConfirm && !confirm('¿Bloquear los objetivos del día para '+C.sup+'?\nEl supervisor no podrá editarlos después (solo Jefatura).')) return;
      const metas={};
      body.querySelectorAll('input[data-ag]').forEach(inp=>{ metas[decodeURIComponent(inp.dataset.ag)]=+inp.value||0; });
      const prev=getObjetivos(C.sup,C.fecha);
      const rec={locked:true, metas, setAt: prev.setAt||nowStamp(), setBy: prev.setBy||'Supervisor'};
      if(ADMIN && prev.locked){ rec.editAt=nowStamp(); rec.editadoJefa=true; }
      saveObjetivos(C.sup,C.fecha,rec);
      render();
    });
  }

  if(obj.locked && !ADMIN){
    state.innerHTML='<span class="chip" style="background:#e7f6ec;color:var(--ok)">🔒 Bloqueado · '+obj.setAt+'</span>'+
      (obj.editadoJefa?' <span class="chip" style="background:#e8efff;color:var(--azul)">✏️ Editado por Jefa · '+(obj.editAt||'')+'</span>':'');
    const rows=roster.map(v=>{
      const m=obj.metas[v]!=null?obj.metas[v]:0;
      const ventas=vmap[v]?vmap[v].ventas:0;
      const pct=m?Math.round(ventas/m*100):0; const cumple=m&&ventas>=m;
      const col=cumple?'var(--ok)':(pct>=70?'var(--warn)':'var(--bad)');
      return `<tr><td>${v}</td><td><b>${m}</b></td><td>${ventas}</td>
        <td style="color:${col};font-weight:700">${pct}%</td>
        <td>${cumple?'<span style="color:var(--ok);font-weight:700">✔ Cumple</span>':'<span style="color:var(--bad);font-weight:700">✗ No</span>'}</td></tr>`;
    }).join('');
    body.innerHTML='<p class="muted" style="margin-top:0">Objetivos fijados a primera hora — <b>no editables</b> por el supervisor. Solo Jefatura puede modificarlos (Modo Jefa).</p>'+
      '<div class="scroll" style="max-height:340px"><table><thead><tr><th>Agente</th><th>Objetivo</th><th>Ventas (al corte)</th><th>% avance</th><th>Cumplimiento</th></tr></thead><tbody>'+rows+'</tbody></table></div>';
  } else if(obj.locked && ADMIN){
    state.innerHTML='<span class="chip" style="background:#e7f6ec;color:var(--ok)">🔒 Bloqueado</span> <span class="chip" style="background:var(--azul);color:#fff">Modo Jefa: puedes editar</span>';
    body.innerHTML='<p class="muted" style="margin-top:0">Estás en <b>Modo Jefa</b>. Puedes corregir el objetivo de un agente; el cambio queda registrado como <b>editado por Jefatura</b> con su hora.</p>'+editorHtml('💾 Guardar cambios (Jefa)');
    wireSave(false);
  } else {
    state.innerHTML='<span class="chip" style="background:#fff4e5;color:#b45309">⚠ Sin bloquear</span>';
    body.innerHTML='<p class="muted" style="margin-top:0">Ingresa el objetivo de cada agente según su <b>skill</b>. Al bloquear quedan fijos para el día y <b>no se podrán modificar</b> (salvo Jefatura).</p>'+editorHtml('🔒 Guardar y bloquear objetivos');
    wireSave(true);
  }
  if(C.multiFecha){
    body.insertAdjacentHTML('afterbegin','<div class="chip" style="background:#fff4e5;color:#b45309;margin-bottom:8px;display:inline-block">Objetivos/tareas aplican al día <b>'+C.fecha+'</b> (hay varias fechas seleccionadas)</div>');
  }
}

// =========================================================
//  MONITOREO DE CALIDAD (registros bloqueables · localStorage)
// =========================================================
function monKey(sup){ return 'monitoreos::'+sup; }
function getMonitoreos(sup){ try{ return JSON.parse(localStorage.getItem(monKey(sup))||'[]'); }catch(e){ return []; } }
function saveMonitoreos(sup,list){ localStorage.setItem(monKey(sup),JSON.stringify(list)); }
function dniMapFor(sup){
  const m={};
  ALL.forEach(r=>{ if((sup==='__ALL__'||r.supervisor===sup) && r.docVendedor) m[r.docVendedor]=r.vendedor; });
  return m;
}
let MON_C=null;
function renderMonitoreo(C){
  MON_C=C;
  const dmap=dniMapFor(C.sup);
  $('dlAsesores').innerHTML=Object.entries(dmap).map(([dni,nom])=>`<option value="${dni}">${nom}</option>`).join('');
  if(!$('mFecha').value) $('mFecha').value=C.fecha;
  const list=getMonitoreos(C.sup);
  $('monState').innerHTML = C.sup==='__ALL__' ? '' : '<span class="chip">'+list.length+' monitoreos</span>';
  if(C.sup==='__ALL__'){
    $('monBody').innerHTML='<tr><td colspan="10" class="muted">Selecciona tu supervisor para registrar y ver monitoreos.</td></tr>';
    return;
  }
  const sino = v => v==='Sí'
    ? '<span style="color:var(--ok);font-weight:700">Sí</span>'
    : (v==='No' ? '<span class="muted">No</span>' : '<span class="muted">—</span>');
  $('monBody').innerHTML = list.length? list.slice().reverse().map((m,i)=>{
    const nom = m.asesor || dmap[m.dniAsesor] || '';
    return `<tr><td>${list.length-i}</td><td>${m.dniAsesor}</td><td>${nom}</td><td>${m.fecha}</td><td>${m.nroLlamada}</td><td>${m.dniCliente}</td><td>${sino(m.migra)}</td><td>${sino(m.cross)}</td><td class="muted">${m.registradoEn}</td><td><span style="color:var(--ok);font-weight:700">🔒 Fijo</span></td></tr>`;
  }).join('') : '<tr><td colspan="10" class="muted">Aún no hay monitoreos registrados.</td></tr>';
}
function addMonitoreo(){
  if(!MON_C || MON_C.sup==='__ALL__'){ $('monMsg').textContent='Selecciona tu supervisor antes de registrar.'; return; }
  const dniAsesor=$('mDniAsesor').value.trim();
  const fecha=$('mFecha').value;
  const nroLlamada=$('mNroLlamada').value.trim();
  const dniCliente=$('mDniCliente').value.trim();
  const migra=$('mMigra').value, cross=$('mCross').value;
  if(!dniAsesor||!fecha||!nroLlamada||!dniCliente){ $('monMsg').textContent='Completa los 4 campos: DNI Asesor, Fecha, N° de llamada y DNI Cliente.'; return; }
  $('monMsg').textContent='';
  const dmap=dniMapFor(MON_C.sup);
  const now=new Date();
  const registradoEn=now.toLocaleString('es-PE',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
  const list=getMonitoreos(MON_C.sup);
  list.push({dniAsesor,asesor:dmap[dniAsesor]||'',fecha,nroLlamada,dniCliente,migra,cross,registradoEn});
  saveMonitoreos(MON_C.sup,list);
  $('mDniAsesor').value=''; $('mNroLlamada').value=''; $('mDniCliente').value='';
  renderMonitoreo(MON_C);
}

// =========================================================
//  TAREAS DEL SUPERVISOR (localStorage)
// =========================================================
function taskKey(C){ return 'tareas::'+C.sup+'::'+C.fecha; }
function getTasks(C){
  try{ return JSON.parse(localStorage.getItem(taskKey(C))||'null') || defaultTasks(); }
  catch(e){ return defaultTasks(); }
}
function defaultTasks(){
  return [
    {t:'Monitoreo de calidad: escuchar 5 llamadas del equipo',done:false},
    {t:'Feedback 1 a 1 con agentes del cuartil Q4',done:false},
    {t:'Revisar ventas observadas / caídas y gestionar',done:false},
    {t:'Reforzar speech de cross-selling en briefing',done:false},
  ];
}
function saveTasks(C,list){ localStorage.setItem(taskKey(C),JSON.stringify(list)); }
let CURRENT_C=null;
function renderTasks(C){
  CURRENT_C=C;
  const list=getTasks(C);
  const done=list.filter(t=>t.done).length;
  const pct=list.length?Math.round(done/list.length*100):0;
  $('taskProgress').innerHTML=`<div style="display:flex;align-items:center;gap:8px"><div class="bar-wrap" style="flex:1;height:10px"><div class="bar-fill" style="width:${pct}%;background:${pct===100?'var(--ok)':'var(--azul2)'}"></div></div><b>${done}/${list.length}</b></div>`;
  $('taskList').innerHTML=list.map((t,i)=>{
    const lock = t.done && !ADMIN;   // marcada y no Jefa -> no se puede cambiar
    return `<div class="task ${t.done?'done':''}">
     <input type="checkbox" ${t.done?'checked':''} ${lock?'disabled':''} data-i="${i}">
     <span class="t">${t.t}</span>
     <span class="meta">${t.done?('<span style="color:var(--ok);font-weight:700">✔ '+(t.doneAt||'')+'</span>'):''} ${ADMIN?'<a href="#" data-del="'+i+'" style="color:var(--bad);text-decoration:none" title="Eliminar (Jefa)">✕</a>':''}</span></div>`;
  }).join('');
  $('taskList').querySelectorAll('input[type=checkbox]').forEach(ck=>ck.addEventListener('change',()=>{
    const t=list[+ck.dataset.i];
    if(ck.checked){ t.done=true; t.doneAt=nowStamp(); }
    else if(ADMIN){ t.done=false; t.doneAt=null; }   // solo Jefa puede desmarcar
    saveTasks(C,list); renderTasks(C);
  }));
  $('taskList').querySelectorAll('a[data-del]').forEach(a=>a.addEventListener('click',e=>{
    e.preventDefault(); if(!ADMIN) return; list.splice(+a.dataset.del,1); saveTasks(C,list); renderTasks(C);
  }));
}
function addTask(){
  if(!CURRENT_C) return;
  const v=$('newTask').value.trim(); if(!v) return;
  const list=getTasks(CURRENT_C); list.push({t:v,done:false,doneAt:null}); saveTasks(CURRENT_C,list);
  $('newTask').value=''; renderTasks(CURRENT_C);
}
