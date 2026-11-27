// Registro de visitas de la página Semana Tec H-FARM 2027.
// La página manda eventos anónimos (sin nombres ni correos) y aquí se guardan en una Google Sheet privada.

const TZ = 'America/Mexico_City';

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const id = props.getProperty('SHEET_ID');
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('Semana Tec H-FARM - visitas');
    props.setProperty('SHEET_ID', ss.getId());
  }
  ss.setSpreadsheetTimeZone(TZ);
  let ev = ss.getSheetByName('Eventos');
  if (!ev) {
    ev = ss.getSheets()[0];
    ev.setName('Eventos');
    ev.getRange(1, 1, 1, 15).setValues([[
      'fecha_hora', 'evento', 'detalle', 'visitante', 'sesion', 'origen', 'campana',
      'dispositivo', 'ancho_px', 'idioma', 'zona_horaria', 'pagina', 'segundos', 'scroll_pct', 'navegador'
    ]]).setFontWeight('bold');
    ev.setFrozenRows(1);
    ev.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  buildResumen_(ss);
  Logger.log('Hoja: ' + ss.getUrl());
}

function buildResumen_(ss) {
  let r = ss.getSheetByName('Resumen');
  if (r) r.clear(); else r = ss.insertSheet('Resumen', 0);
  const E = 'Eventos!';
  const kpis = [
    ['Indicador', 'Valor'],
    ['Visitas (páginas vistas)', '=COUNTIF(' + E + 'B:B,"visita")'],
    ['Visitantes únicos', '=IFERROR(COUNTUNIQUE(FILTER(' + E + 'D2:D,' + E + 'B2:B="visita")),0)'],
    ['Visitas hoy', '=COUNTIFS(' + E + 'B:B,"visita",' + E + 'A:A,">="&TODAY())'],
    ['Visitas últimos 7 días', '=COUNTIFS(' + E + 'B:B,"visita",' + E + 'A:A,">="&TODAY()-6)'],
    ['Tiempo promedio en la página (min)', '=IFERROR(ROUND(AVERAGEIF(' + E + 'B:B,"salida",' + E + 'M:M)/60,1),0)'],
    ['Scroll promedio (% de la página)', '=IFERROR(ROUND(AVERAGEIF(' + E + 'B:B,"salida",' + E + 'N:N),0),0)'],
    ['Clics en «Regístrate»', '=COUNTIF(' + E + 'B:B,"clic_registro")'],
    ['Clics en «Sitio de pagos»', '=COUNTIF(' + E + 'B:B,"clic_pagos")'],
    ['Clics en comunidad de WhatsApp', '=COUNTIF(' + E + 'B:B,"clic_comunidad_whatsapp")'],
    ['Clics en WhatsApp de contacto', '=COUNTIF(' + E + 'B:B,"clic_whatsapp_contacto")'],
    ['Clics en correos', '=COUNTIF(' + E + 'B:B,"clic_correo")'],
    ['Reproducciones del video', '=COUNTIF(' + E + 'B:B,"video")'],
    ['Clics a YouTube', '=COUNTIF(' + E + 'B:B,"clic_youtube")'],
    ['% de visitantes que dio clic en «Regístrate»', '=IFERROR(ROUND(COUNTUNIQUE(FILTER(' + E + 'D2:D,' + E + 'B2:B="clic_registro"))/COUNTUNIQUE(FILTER(' + E + 'D2:D,' + E + 'B2:B="visita"))*100,1),0)'],
    ['% de visitantes que dio clic en pagos', '=IFERROR(ROUND(COUNTUNIQUE(FILTER(' + E + 'D2:D,' + E + 'B2:B="clic_pagos"))/COUNTUNIQUE(FILTER(' + E + 'D2:D,' + E + 'B2:B="visita"))*100,1),0)']
  ];
  r.getRange(1, 1, kpis.length, 2).setValues(kpis);
  r.getRange(1, 1, 1, 2).setFontWeight('bold');

  const q = (col, title, sql) => {
    r.getRange(1, col).setValue(title).setFontWeight('bold');
    r.getRange(2, col).setFormula(sql);
  };
  q(4, 'Visitas por día', '=IFERROR(QUERY(' + E + 'A2:D,"select toDate(A), count(B) where B=\'visita\' group by toDate(A) order by toDate(A) desc label toDate(A) \'Día\', count(B) \'Visitas\'",0),"Sin datos")');
  q(8, 'De dónde llegan', '=IFERROR(QUERY(' + E + 'B2:G,"select F, count(B) where B=\'visita\' group by F order by count(B) desc label F \'Origen\', count(B) \'Visitas\'",0),"Sin datos")');
  q(11, 'Campaña o liga (?v=)', '=IFERROR(QUERY(' + E + 'B2:G,"select G, count(B) where B=\'visita\' and G<>\'\' group by G order by count(B) desc label G \'Campaña\', count(B) \'Visitas\'",0),"Sin datos")');
  q(14, 'Dispositivo', '=IFERROR(QUERY(' + E + 'B2:H,"select H, count(B) where B=\'visita\' group by H order by count(B) desc label H \'Dispositivo\', count(B) \'Visitas\'",0),"Sin datos")');
  q(17, 'Secciones vistas', '=IFERROR(QUERY(' + E + 'B2:C,"select C, count(B) where B=\'seccion\' group by C order by count(B) desc label C \'Sección\', count(B) \'Veces\'",0),"Sin datos")');
  q(20, 'Días de la agenda consultados', '=IFERROR(QUERY(' + E + 'B2:C,"select C, count(B) where B=\'dia\' group by C order by count(B) desc label C \'Día\', count(B) \'Veces\'",0),"Sin datos")');
  q(23, 'Preguntas abiertas', '=IFERROR(QUERY(' + E + 'B2:C,"select C, count(B) where B=\'pregunta\' group by C order by count(B) desc label C \'Pregunta\', count(B) \'Veces\'",0),"Sin datos")');
  q(26, 'Perfiles, fotos y entregable', '=IFERROR(QUERY(' + E + 'B2:C,"select B, C, count(B) where B=\'perfil\' or B=\'foto\' or B=\'entregable\' group by B, C order by count(B) desc label B \'Tipo\', C \'Detalle\', count(B) \'Veces\'",0),"Sin datos")');
  q(30, 'Navegador', '=IFERROR(QUERY(' + E + 'B2:O,"select O, count(B) where B=\'visita\' group by O order by count(B) desc label O \'Navegador\', count(B) \'Visitas\'",0),"Sin datos")');
  r.setFrozenRows(1);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(8000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const events = Array.isArray(d.events) ? d.events.slice(0, 20) : [];
    if (!events.length) return ContentService.createTextOutput('vacío');
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
    const sh = ss.getSheetByName('Eventos');
    const now = new Date();
    const rows = events.map(ev => [
      now, clip_(ev.e, 40), clip_(ev.d, 120), clip_(d.v, 40), clip_(d.s, 40), clip_(d.r, 120), clip_(d.c, 60),
      clip_(d.dev, 20), num_(d.w), clip_(d.l, 20), clip_(d.tz, 40), clip_(d.p, 120), num_(ev.t), num_(ev.sc), clip_(d.b, 30)
    ]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const key = PropertiesService.getScriptProperties().getProperty('DASH_KEY');
  if (key && p.k === key) {
    return HtmlService.createHtmlOutput(dashHtml_(JSON.stringify(rows_()).replace(/</g, '\\u003c')))
      .setTitle('Visitas · Semana Tec H-FARM')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return ContentService.createTextOutput('Registro de visitas activo');
}

// Crea (una sola vez) la clave de la liga privada del dashboard.
function setupDashboard() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty('DASH_KEY');
  if (!k) { k = Utilities.getUuid().replace(/-/g, '').slice(0, 20); props.setProperty('DASH_KEY', k); }
  Logger.log('Clave del dashboard: ' + k);
}

function rows_() {
  const sh = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName('Eventos');
  const vals = sh.getDataRange().getValues();
  vals.shift();
  return vals.filter(r => r[1]).map(r => [
    Utilities.formatDate(r[0] instanceof Date ? r[0] : new Date(r[0]), TZ, "yyyy-MM-dd'T'HH:mm"),
    String(r[1]), String(r[2]), String(r[3]), String(r[4]), String(r[5]), String(r[6]), String(r[7]), String(r[14]),
    Number(r[12]) || 0, Number(r[13]) || 0
  ]);
}


function clip_(x, n) { return String(x == null ? '' : x).slice(0, n); }
function num_(x) { const v = Number(x); return isFinite(v) && x !== '' && x != null ? v : ''; }

function dashHtml_(dataJson) {
  return `<!doctype html><html lang="es"><head><meta charset="utf-8">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;700&family=Barlow+Condensed:wght@800&family=Newsreader:ital,wght@0,500;1,400&display=swap">
<style>
:root{--crema:#F1EEE8;--crema2:#E7E1D6;--tinta:#101216;--gris:#5E6168;--gris2:#3A3D43;--regla:#CFC9BD;--nar:#FF5B1A;--nart:#B83A0B}
*{box-sizing:border-box}body{margin:0;background:var(--crema);color:var(--tinta);font:16px/1.5 Archivo,Arial,sans-serif}
.wrap{max-width:1180px;margin:0 auto;padding:28px max(16px,env(safe-area-inset-left)) 60px}
.top{display:flex;align-items:center;gap:14px;flex-wrap:wrap;border-bottom:2px solid var(--tinta);padding-bottom:16px}
.sq{width:14px;height:14px;background:var(--nar);display:inline-block}
.eyebrow{font-size:13px;font-weight:700;letter-spacing:3px;text-transform:uppercase}
h1{font-family:'Barlow Condensed',sans-serif;font-weight:800;text-transform:uppercase;font-size:clamp(34px,5vw,56px);line-height:.95;margin:14px 0 4px}
h1 span{color:var(--nart)}
.meta{color:var(--gris);font-size:14px;margin:0}
.filters{display:flex;gap:8px;flex-wrap:wrap;margin:18px 0 6px}
.filters button{font:700 13px Archivo,Arial;letter-spacing:1px;text-transform:uppercase;border:2px solid var(--tinta);background:transparent;padding:8px 12px;cursor:pointer}
.filters button.on{background:var(--tinta);color:var(--crema)}
.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:18px;margin:22px 0}
.kpi{border-top:2px solid var(--tinta);padding-top:10px}.kpi.acc{border-top-color:var(--nar)}
.kpi b{display:block;font-family:Newsreader,Georgia,serif;font-weight:500;font-size:clamp(30px,4vw,46px);line-height:1}
.kpi span{font-size:13px;color:var(--gris);letter-spacing:1px;text-transform:uppercase;font-weight:700}
.grid{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:28px}
.card{border-top:2px solid var(--tinta);padding-top:12px;min-width:0}
.card h2{font-family:Newsreader,Georgia,serif;font-weight:500;font-size:24px;margin:0 0 4px}
.card p.sub{margin:0 0 12px;color:var(--gris);font-size:14px}
.full{grid-column:1/-1}
.bar{display:grid;grid-template-columns:minmax(0,190px) minmax(0,1fr) 64px;gap:10px;align-items:center;padding:5px 0;font-size:15px}
.bar .l{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bar .t{height:14px;background:var(--crema2)}.bar .t i{display:block;height:100%;background:var(--tinta)}
.bar.hot .t i{background:var(--nar)}
.bar .n{text-align:right;font-variant-numeric:tabular-nums;font-weight:700}
.empty{color:var(--gris);font-style:italic;font-family:Newsreader,Georgia,serif;font-size:18px}
table{width:100%;border-collapse:collapse;font-size:14px;font-variant-numeric:tabular-nums}
th,td{text-align:left;padding:7px 8px;border-bottom:1px solid var(--regla)}th{background:var(--tinta);color:var(--crema);font-size:12px;letter-spacing:1px;text-transform:uppercase}
.scroll{overflow-x:auto}
svg text{font-family:Archivo,Arial;font-size:11px;fill:var(--gris)}
.note{margin-top:30px;color:var(--gris);font-size:13px}
@media(max-width:820px){.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.grid{grid-template-columns:1fr}.bar{grid-template-columns:minmax(0,130px) minmax(0,1fr) 52px}}
</style></head><body><div class="wrap">
<div class="top"><span class="sq"></span><span class="eyebrow">Semana Tec Internacional · H-FARM 2027</span></div>
<h1>Visitas a la <span>página</span></h1>
<p class="meta" id="gen"></p>
<div class="filters" id="filters"><button data-d="1">Hoy</button><button data-d="7">7 días</button><button data-d="30">30 días</button><button data-d="0" class="on">Todo</button></div>
<div class="kpis" id="kpis"></div>
<div class="grid">
 <div class="card full"><h2>Visitas por día</h2><p class="sub">Barras: visitas · línea naranja: visitantes únicos</p><div id="daily"></div></div>
 <div class="card"><h2>De dónde llegan</h2><p class="sub">Canal de origen de cada visita</p><div id="canal"></div></div>
 <div class="card"><h2>Liga o campaña</h2><p class="sub">Lo que va después de ? en la liga, por ejemplo v=3</p><div id="camp"></div></div>
 <div class="card full"><h2>Qué ven de la página</h2><p class="sub">% de visitas que llegó a ver cada sección, en el orden de la página</p><div id="secc"></div></div>
 <div class="card"><h2>Hasta dónde bajan</h2><p class="sub">Qué tanto de la página recorrieron antes de irse</p><div id="scrl"></div></div>
 <div class="card"><h2>Qué hacen</h2><p class="sub">Clics en botones y ligas</p><div id="clics"></div></div>
 <div class="card"><h2>Días de la agenda que revisan</h2><div id="dias"></div></div>
 <div class="card"><h2>Preguntas que abren</h2><div id="preg"></div></div>
 <div class="card"><h2>Perfiles, fotos y entregable</h2><div id="otros"></div></div>
 <div class="card"><h2>Dispositivo y navegador</h2><div id="dev"></div><div id="nav" style="margin-top:14px"></div></div>
 <div class="card full"><h2>Últimas visitas</h2><div class="scroll"><table id="last"></table></div></div>
</div>
<p class="note">Datos anónimos: no se guardan nombres, correos ni teléfonos de quien visita. Se actualiza cada vez que abres o recargas esta página.</p>
</div>
<script>
var RAW = ${dataJson};
var SEC = [['inicio','Portada y video'],['hfarm','Qué es H-FARM y campus'],['venecia','Venecia y Burano'],['semana','Tu semana (agenda)'],['entregable','Lo que te llevas'],['materias','Qué materia acreditas'],['inversion','Inversión y pagos'],['compass','Compass y vuelo'],['preguntas','Preguntas frecuentes'],['registro','Siguiente paso y contactos'],['videos','Videos de YouTube']];
var CLK = [['clic_registro','Regístrate'],['clic_pagos','Sitio de pagos'],['clic_comunidad_whatsapp','Comunidad de WhatsApp'],['clic_whatsapp_contacto','WhatsApp de contacto'],['clic_correo','Correos'],['video','Play al video'],['clic_youtube','Videos de YouTube'],['calculadora','Calculadora de pagos'],['menu','Menú de arriba']];
function rows(){return RAW.map(function(r){return{t:r[0],day:r[0].slice(0,10),e:r[1],d:r[2],v:r[3],s:r[4],r:r[5],c:r[6],dev:r[7],b:r[8],sec:r[9],sc:r[10]};});}
function canal(x){var r=(x.r||'').toLowerCase();if(x.b==='WhatsApp'||r.indexOf('whatsapp')>-1)return'WhatsApp';if(x.b==='Facebook/Instagram'||/facebook|fb\\.|instagram/.test(r))return'Facebook / Instagram';if(r.indexOf('linkedin')>-1||r.indexOf('lnkd')>-1)return'LinkedIn';if(r.indexOf('google')>-1)return'Google';if(r.indexOf('t.co')>-1||r.indexOf('twitter')>-1||r==='x.com')return'X / Twitter';if(r.indexOf('claude.ai')>-1)return'Presentación (Claude)';if(r==='directo'||!r)return'Directo (liga copiada o QR)';if(r==='misma página')return'Misma página';return r;}
function today(){var d=new Date();return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);}
function since(n){var d=new Date();d.setDate(d.getDate()-(n-1));return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);}
function count(list,key){var m={};list.forEach(function(x){var k=key(x);if(k)m[k]=(m[k]||0)+1;});return Object.keys(m).map(function(k){return[k,m[k]];}).sort(function(a,b){return b[1]-a[1];});}
function uniq(list,f){var s={};list.forEach(function(x){s[f(x)]=1;});return Object.keys(s).length;}
function bars(el,pairs,opt){opt=opt||{};var box=document.getElementById(el);if(!pairs.length){box.innerHTML='<p class="empty">Todavía sin datos</p>';return;}var max=opt.max||Math.max.apply(null,pairs.map(function(p){return p[1];}));box.innerHTML=pairs.map(function(p,i){var w=max?Math.round(p[1]/max*100):0;var lab=opt.pct?(p[1]+'%'):p[1];return'<div class="bar'+(i===0&&!opt.flat?' hot':'')+'"><span class="l" title="'+esc(p[0])+'">'+esc(p[0])+'</span><span class="t"><i style="width:'+w+'%"></i></span><span class="n">'+lab+'</span></div>';}).join('');}
function esc(s){return String(s).replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function render(days){
 var all=rows();var cut=days?since(days):'';var R=all.filter(function(x){return!cut||x.day>=cut;});
 var V=R.filter(function(x){return x.e==='visita';});var S=R.filter(function(x){return x.e==='salida';});
 var nv=V.length,nu=uniq(V,function(x){return x.v;});
 var avgT=S.length?Math.round(S.reduce(function(a,x){return a+x.sec;},0)/S.length):0;
 var avgS=S.length?Math.round(S.reduce(function(a,x){return a+x.sc;},0)/S.length):0;
 var regU=uniq(R.filter(function(x){return x.e==='clic_registro';}),function(x){return x.v;});
 var payU=uniq(R.filter(function(x){return x.e==='clic_pagos';}),function(x){return x.v;});
 var k=[['Visitas',nv,1],['Visitantes únicos',nu,0],['Tiempo promedio',avgT>=60?(Math.floor(avgT/60)+' min '+(avgT%60)+' s'):(avgT+' s'),0],['Bajan en promedio',avgS+'%',0],['Dieron clic en Regístrate',regU+(nu?' · '+Math.round(regU/nu*100)+'%':''),1],['Dieron clic en pagos',payU+(nu?' · '+Math.round(payU/nu*100)+'%':''),0],['Visitas hoy',all.filter(function(x){return x.e==='visita'&&x.day===today();}).length,0],['Celular',nv?Math.round(V.filter(function(x){return x.dev==='celular';}).length/nv*100)+'%':'0%',0]];
 document.getElementById('kpis').innerHTML=k.map(function(x){return'<div class="kpi'+(x[2]?' acc':'')+'"><b>'+x[1]+'</b><span>'+x[0]+'</span></div>';}).join('');
 daily(all,days);
 bars('canal',count(V,canal));
 bars('camp',count(V,function(x){return x.c||'(sin liga especial)';}));
 var secV={};R.filter(function(x){return x.e==='seccion';}).forEach(function(x){secV[x.d]=secV[x.d]||{};secV[x.d][x.s]=1;});
 var ses=uniq(V,function(x){return x.s;});
 bars('secc',SEC.map(function(p){var n=secV[p[0]]?Object.keys(secV[p[0]]).length:0;return[p[1],ses?Math.min(100,Math.round(n/ses*100)):0];}).filter(function(){return nv>0;}),{pct:1,max:100,flat:1});
 var b=[['Menos de 25%',0],['25 a 50%',0],['50 a 75%',0],['75 a 99%',0],['Hasta el final',0]];S.forEach(function(x){var i=x.sc>=99?4:x.sc>=75?3:x.sc>=50?2:x.sc>=25?1:0;b[i][1]++;});
 bars('scrl',S.length?b:[],{flat:1});
 bars('clics',CLK.map(function(p){return[p[1],R.filter(function(x){return x.e===p[0];}).length];}).filter(function(p){return p[1]>0;}));
 bars('dias',count(R.filter(function(x){return x.e==='dia';}),function(x){return x.d.replace(/^(\\w{3}) (\\d)/,'$1 $2 · ');}));
 bars('preg',count(R.filter(function(x){return x.e==='pregunta';}),function(x){return x.d;}));
 bars('otros',count(R.filter(function(x){return x.e==='perfil'||x.e==='foto'||x.e==='entregable';}),function(x){return({perfil:'Perfil: ',foto:'Foto: ',entregable:'Entregable: '})[x.e]+x.d;}));
 bars('dev',count(V,function(x){return x.dev;}));bars('nav',count(V,function(x){return x.b;}));
 var last=V.slice().reverse().slice(0,25);var exitBy={};S.forEach(function(x){exitBy[x.s]=x;});
 document.getElementById('last').innerHTML='<tr><th>Fecha y hora</th><th>Origen</th><th>Liga</th><th>Dispositivo</th><th>Tiempo</th><th>Bajó</th></tr>'+(last.length?last.map(function(x){var o=exitBy[x.s];return'<tr><td>'+x.t.replace('T',' ')+'</td><td>'+esc(canal(x))+'</td><td>'+esc(x.c||'')+'</td><td>'+esc(x.dev)+'</td><td>'+(o?o.sec+' s':'')+'</td><td>'+(o?o.sc+'%':'')+'</td></tr>';}).join(''):'<tr><td colspan="6" class="empty">Todavía sin visitas</td></tr>');
}
function daily(all,days){var box=document.getElementById('daily');var n=days&&days>1?days:(days===1?1:30);var keys=[];for(var i=n-1;i>=0;i--){var d=new Date();d.setDate(d.getDate()-i);keys.push(d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2));}
 var vis={},uni={};all.forEach(function(x){if(x.e!=='visita')return;vis[x.day]=(vis[x.day]||0)+1;uni[x.day]=uni[x.day]||{};uni[x.day][x.v]=1;});
 var vals=keys.map(function(k){return vis[k]||0;}),u=keys.map(function(k){return uni[k]?Object.keys(uni[k]).length:0;});var max=Math.max(1,Math.max.apply(null,vals));
 var W=1100,H=200,pl=30,pb=24,bw=(W-pl)/keys.length;var s='<svg viewBox="0 0 '+W+' '+(H+pb)+'" width="100%" role="img" aria-label="Visitas por día">';
 (max>=2?[0,0.5,1]:[0,1]).forEach(function(f){var y=H-f*(H-10);s+='<line x1="'+pl+'" x2="'+W+'" y1="'+y+'" y2="'+y+'" stroke="#CFC9BD"/><text x="0" y="'+(y+4)+'">'+Math.round(f*max)+'</text>';});
 vals.forEach(function(v,i){var h=v/max*(H-10);s+='<rect x="'+(pl+i*bw+bw*0.15)+'" y="'+(H-h)+'" width="'+(bw*0.7)+'" height="'+h+'" fill="#101216"><title>'+keys[i]+': '+v+' visitas</title></rect>';if(keys.length<=31&&(i%Math.ceil(keys.length/10)===0||i===keys.length-1))s+='<text x="'+(pl+i*bw+bw/2)+'" y="'+(H+16)+'" text-anchor="middle">'+keys[i].slice(5).replace('-','/')+'</text>';});
 s+='<polyline fill="none" stroke="#FF5B1A" stroke-width="3" points="'+u.map(function(v,i){return(pl+i*bw+bw/2)+','+(H-v/max*(H-10));}).join(' ')+'"/></svg>';box.innerHTML=s;}
document.getElementById('gen').textContent='Datos al '+new Date().toLocaleString('es-MX',{dateStyle:'long',timeStyle:'short'})+' · '+(function(n){return n+(n===1?' visita registrada':' visitas registradas');})(RAW.filter(function(r){return r[1]==='visita';}).length)+' desde que se activó el contador';
document.getElementById('filters').addEventListener('click',function(ev){var b=ev.target.closest('button');if(!b)return;[].forEach.call(this.children,function(x){x.classList.toggle('on',x===b);});render(+b.dataset.d);});
render(0);
</script></body></html>`;
}
