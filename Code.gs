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

function doGet() {
  return ContentService.createTextOutput('Registro de visitas activo');
}

function clip_(x, n) { return String(x == null ? '' : x).slice(0, n); }
function num_(x) { const v = Number(x); return isFinite(v) && x !== '' && x != null ? v : ''; }
