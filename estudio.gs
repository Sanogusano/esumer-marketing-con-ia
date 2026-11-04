/**
 * ESTUDIO DE CONTENIDOS · Loop de agentes para droguerías
 *
 * Los agentes los escribes tú en la pestaña "Agentes". Este motor:
 * 1. Lee tus agentes y tu cerebro (Google Docs).
 * 2. Corre los agentes en orden. El Supervisor revisa y devuelve el trabajo al agente que corresponda.
 * 3. Guarda la historia del día y te la envía por correo.
 *
 * Instalación: Extensiones → Apps Script → borrar todo → pegar → guardar → recargar la hoja →
 * menú Estudio → Instalar.
 */

// ===================== CONFIGURACIÓN =====================

const H_CONFIG = 'Config';
const H_AGENTES = 'Agentes';
const H_HISTORIAS = 'Historias';
const H_REGISTRO = 'Registro';

const DOCS = ['Kit de marca', 'Productos y temporadas', 'Reglas del Regente'];
const PAUSA_MS = 5000; // el plan gratuito de Gemini permite pocas solicitudes por minuto

// Palabras que el Regente nunca deja pasar (regla fija, sin IA)
const PROHIBIDAS = ['amoxicilina', 'azitromicina', 'cefalexina', 'ciprofloxacino', 'antibiotico', 'losartan', 'metformina',
  'insulina', 'clonazepam', 'alprazolam', 'tramadol', 'codeina', 'sildenafil', 'anticonceptiv', 'corticoide', 'prednisolona',
  'cura', 'curar', 'elimina la', 'garantizado', 'garantizada', '100 %', '100%', 'milagro', 'recomendado por medicos'];

// ===================== MENÚ =====================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Estudio')
    .addItem('Instalar', 'instalar')
    .addSeparator()
    .addItem('Crear la historia de hoy', 'crearHistoriaAhora')
    .addSeparator()
    .addItem('Activar correo diario', 'activarCorreoDiario')
    .addItem('Detener correo diario', 'detenerCorreoDiario')
    .addToUi();
}

// Al pegar la clave o un enlace en Config, se actualiza el estado
function onEdit(e) {
  try { if (e && e.range.getSheet().getName() === H_CONFIG) actualizarEstado_(); } catch (err) {}
}

// ===================== INSTALACIÓN =====================

function instalar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const negro = h => h.setBackground('#16161D').setFontColor('#FFFFFF').setFontWeight('bold');

  if (!ss.getSheetByName(H_CONFIG)) {
    const h = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName(H_CONFIG) : ss.insertSheet(H_CONFIG);
    const filas = [
      ['Dato', 'Valor', 'Qué es'],
      ['Clave de Gemini', '', 'Pégala aquí. La sacas en aistudio.google.com/apikey'],
      ['Droguería', 'Mi droguería', 'El nombre de tu droguería'],
      ['Correo', Session.getEffectiveUser().getEmail(), 'A dónde llega la historia del día'],
      ['Hora del correo', 7, 'Hora del día (0 a 23) en que llega el correo'],
      ['Tema de hoy (opcional)', '', 'Si lo llenas, el Estratega trabaja sobre este tema'],
      ['Vueltas máximas', 3, 'Cuántas veces puede devolver el Supervisor antes de pasarlo a revisión humana'],
      ['Modelo', 'gemini-flash-latest', 'Modelo de Gemini'],
      ['Kit de marca', '', 'Enlace al Google Doc'],
      ['Productos y temporadas', '', 'Enlace al Google Doc'],
      ['Reglas del Regente', '', 'Enlace al Google Doc'],
      ['Estado', '', '']
    ];
    h.getRange(1, 1, filas.length, 3).setValues(filas);
    negro(h.getRange('A1:C1'));
    h.getRange('A2:A12').setFontWeight('bold');
    h.setColumnWidth(1, 200); h.setColumnWidth(2, 430); h.setColumnWidth(3, 470);
    h.setFrozenRows(1);
  }
  if (!ss.getSheetByName(H_AGENTES)) {
    const h = ss.insertSheet(H_AGENTES);
    negro(h.getRange(1, 1, 1, 6).setValues([['Orden', 'Agente', 'Tipo', 'Lee del cerebro', 'Instrucciones', 'Activo']]));
    h.setFrozenRows(1);
    h.setColumnWidth(2, 130); h.setColumnWidth(3, 110); h.setColumnWidth(4, 260); h.setColumnWidth(5, 640);
    h.getRange('E:E').setWrap(true);
    h.getRange('C2:C30').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Trabajador', 'Supervisor']).build());
    h.getRange('F2:F30').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Sí', 'No']).build());
  }
  if (!ss.getSheetByName(H_HISTORIAS)) {
    const h = ss.insertSheet(H_HISTORIAS);
    negro(h.getRange(1, 1, 1, 5).setValues([['Fecha', 'Estado', 'Vueltas', 'Historia', 'Correo']]));
    h.setFrozenRows(1); h.setColumnWidth(4, 700); h.getRange('D:D').setWrap(true);
  }
  if (!ss.getSheetByName(H_REGISTRO)) {
    const h = ss.insertSheet(H_REGISTRO);
    negro(h.getRange(1, 1, 1, 5).setValues([['Fecha', 'Vuelta', 'Agente', 'Resultado', 'Detalle']]));
    h.setFrozenRows(1); h.setColumnWidth(4, 220); h.setColumnWidth(5, 700); h.getRange('E:E').setWrap(true);
  }
  actualizarEstado_();
  ss.setActiveSheet(ss.getSheetByName(H_CONFIG));
  ui.alert('El Estudio quedó instalado. Pega tu clave de Gemini en la pestaña Config, celda B2.');
}

// ===================== DISPARADORES =====================

function crearHistoriaAhora() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Los agentes están trabajando. Mira la pestaña Registro.', 'Estudio', 10);
  const r = historiaDelDia();
  if (r) SpreadsheetApp.getUi().alert('Historia ' + r.estado.toLowerCase() + ' en ' + r.vueltas + ' vuelta(s). ' +
    (r.correo ? 'Te llegó al correo.' : 'No se pudo enviar el correo.') + ' Mírala en la pestaña Historias.');
}

function activarCorreoDiario() {
  detenerCorreoDiario(true);
  const hora = Math.min(23, Math.max(0, Number(config_()['Hora del correo']) || 7));
  ScriptApp.newTrigger('historiaDelDia').timeBased().everyDays(1).atHour(hora).create();
  actualizarEstado_();
  SpreadsheetApp.getUi().alert('Listo. Todos los días, entre las ' + hora + ' y las ' + (hora + 1) + ', te llega la historia del día.');
}

function detenerCorreoDiario(silencioso) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'historiaDelDia').forEach(t => ScriptApp.deleteTrigger(t));
  actualizarEstado_();
  if (silencioso !== true) SpreadsheetApp.getUi().alert('Correo diario detenido.');
}

// ===================== EL LOOP =====================

function historiaDelDia() {
  const cfg = config_();
  if (!cfg['Clave de Gemini']) { actualizarEstado_(); return null; }
  const { trabajadores, supervisor } = leerAgentes_();
  if (!trabajadores.length) throw new Error('No hay agentes Trabajador activos en la pestaña Agentes.');

  const cerebro = cerebro_(cfg);
  const fecha = new Date();
  const contexto = {
    droguería: cfg['Droguería'],
    fecha_de_hoy: Utilities.formatDate(fecha, Session.getScriptTimeZone(), "EEEE d 'de' MMMM 'de' yyyy"),
    es_quincena: [14, 15, 16, 29, 30, 31, 1].indexOf(fecha.getDate()) >= 0 ? 'Sí' : 'No',
    tema_sugerido_por_el_dueño: String(cfg['Tema de hoy (opcional)'] || ''),
    historias_recientes: historiasRecientes_(5)
  };
  const maxVueltas = Number(cfg['Vueltas máximas']) || 3;
  registrar_(0, 'Motor', 'INICIO', 'Agentes: ' + trabajadores.map(a => a.nombre).join(' → ') + (supervisor ? ' · Supervisor: ' + supervisor.nombre : ''));

  const salidas = {};
  let desde = 0, comentario = null, aprobado = false, vuelta = 1;

  while (vuelta <= maxVueltas) {
    for (let i = desde; i < trabajadores.length; i++) {
      const ag = trabajadores[i];
      const previas = trabajadores.slice(0, i).map(a => ({ agente: a.nombre, entrega: salidas[a.nombre] }));
      const fb = (i === desde && comentario) ? { tu_entrega_anterior: salidas[ag.nombre], el_supervisor_la_devolvio_por: comentario } : null;
      salidas[ag.nombre] = correrTrabajador_(ag, contexto, previas, fb, cerebro, cfg);
      registrar_(vuelta, ag.nombre, 'ENTREGA', salidas[ag.nombre]);
    }
    if (!supervisor) { aprobado = true; break; }

    let rev = reglaFija_(trabajadores, salidas);
    if (rev) {
      registrar_(vuelta, supervisor.nombre + ' (regla fija)', 'RECHAZADO → ' + rev.devolver, rev.motivo);
    } else {
      rev = correrSupervisor_(supervisor, trabajadores, salidas, contexto, cerebro, cfg);
      registrar_(vuelta, supervisor.nombre, rev.ok ? 'APROBADO' : 'RECHAZADO → ' + rev.devolver, rev.motivo);
    }
    if (rev.ok) { aprobado = true; break; }

    vuelta++;
    if (vuelta > maxVueltas) break;
    desde = Math.max(0, trabajadores.findIndex(a => norm_(a.nombre) === norm_(rev.devolver)));
    if (desde < 0) desde = trabajadores.length - 1;
    comentario = rev.motivo;
  }

  const vueltas = Math.min(vuelta, maxVueltas);
  const estado = aprobado ? 'Aprobada' : 'Revisión humana';
  if (!aprobado) registrar_(vueltas, 'Motor', 'REVISIÓN HUMANA', 'No pasó la revisión en ' + maxVueltas + ' vueltas.');

  const final = salidas[trabajadores[trabajadores.length - 1].nombre];
  let correo = false;
  try { enviarCorreo_(cfg, contexto, trabajadores, salidas, estado, vueltas); correo = true; }
  catch (err) { registrar_(vueltas, 'Motor', 'ERROR DE CORREO', err.message); }

  hoja_(H_HISTORIAS).appendRow([fecha, estado, vueltas, final, correo ? 'Enviado' : 'No enviado']);
  const f = hoja_(H_HISTORIAS).getLastRow();
  hoja_(H_HISTORIAS).getRange(f, 2).setFontColor(aprobado ? '#1F7A4D' : '#C23A18').setFontWeight('bold');
  return { estado: estado, vueltas: vueltas, correo: correo };
}

// ===================== AGENTES =====================

function correrTrabajador_(ag, contexto, previas, fb, cerebro, cfg) {
  const sistema = 'Eres el agente "' + ag.nombre + '" de un equipo de contenidos de una droguería de barrio en Colombia.\n' +
    'TUS INSTRUCCIONES:\n' + ag.instrucciones + '\n\n' +
    'REGLAS DEL EQUIPO: escribe en español colombiano, claro y corto. No uses formato Markdown con asteriscos. ' +
    'Nunca menciones medicamentos de venta con fórmula médica, dosis ni promesas de curar.\n\n' +
    'TU CEREBRO (documentos de la droguería):\n' + leerCerebro_(ag.lee, cerebro);
  const entrada = { contexto: contexto, trabajo_de_los_agentes_anteriores: previas };
  if (fb) entrada.correccion = fb;
  return gemini_(sistema, JSON.stringify(entrada), cfg, false);
}

function correrSupervisor_(sup, trabajadores, salidas, contexto, cerebro, cfg) {
  const nombres = trabajadores.map(a => a.nombre);
  const sistema = 'Eres el agente "' + sup.nombre + '", el supervisor del equipo de contenidos de una droguería.\n' +
    'TUS INSTRUCCIONES:\n' + sup.instrucciones + '\n\n' +
    'TU CEREBRO:\n' + leerCerebro_(sup.lee, cerebro) + '\n\n' +
    'Revisa el trabajo del equipo. Si todo cumple, apruébalo. Si algo falla, di qué y a qué agente se lo devuelves.\n' +
    'Los agentes del equipo son: ' + nombres.join(', ') + '.\n' +
    'Responde SOLO un JSON: {"veredicto": "APROBADO" o "RECHAZADO", "motivo": "una línea concreta", "devolver_a": "nombre exacto de un agente o vacío"}';
  const entrada = { contexto: contexto, trabajo_del_equipo: nombres.map(n => ({ agente: n, entrega: salidas[n] })) };
  const r = gemini_(sistema, JSON.stringify(entrada), cfg, true);
  const ok = String(r.veredicto || '').toUpperCase().indexOf('APROBADO') === 0;
  const dev = nombres.find(n => norm_(n) === norm_(r.devolver_a)) || nombres[nombres.length - 1];
  return { ok: ok, motivo: String(r.motivo || (ok ? 'Cumple.' : 'Sin motivo.')), devolver: dev };
}

/** Regla fija del supervisor: palabras que nunca pasan. */
function reglaFija_(trabajadores, salidas) {
  for (const ag of trabajadores) {
    const t = ' ' + norm_(salidas[ag.nombre]) + ' ';
    const p = PROHIBIDAS.find(w => new RegExp('(^|[^a-z0-9])' + norm_(w).replace(/[.*+?^${}()|[\]\\%]/g, '\\$&')).test(t));
    if (p) return { ok: false, devolver: ag.nombre, motivo: 'La entrega de ' + ag.nombre + ' menciona "' + p.trim() + '". En una droguería eso no se puede publicar. Cambia el tema o el producto por uno de venta libre.' };
    if (/\d+\s?(mg|ml|mcg|tabletas|capsulas|gotas)\b/.test(t)) return { ok: false, devolver: ag.nombre, motivo: 'La entrega de ' + ag.nombre + ' menciona una dosis. Eso no se publica.' };
  }
  return null;
}

// ===================== CORREO =====================

function enviarCorreo_(cfg, contexto, trabajadores, salidas, estado, vueltas) {
  const para = String(cfg['Correo'] || Session.getEffectiveUser().getEmail());
  const esc = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br>');
  const ultimo = trabajadores[trabajadores.length - 1].nombre;
  const color = estado === 'Aprobada' ? '#1F7A4D' : '#C23A18';
  let html = '<div style="font-family:Arial,sans-serif;max-width:640px;color:#16161D">' +
    '<p style="color:#C23A18;font-weight:bold;letter-spacing:1px;font-size:12px">ESTUDIO DE CONTENIDOS · ' + esc(cfg['Droguería']).toUpperCase() + '</p>' +
    '<h2 style="margin:0 0 6px">La historia de hoy</h2>' +
    '<p style="margin:0 0 18px;color:#55555F">' + esc(contexto.fecha_de_hoy) + ' · <b style="color:' + color + '">' + estado + '</b> en ' + vueltas + ' vuelta(s)</p>' +
    '<div style="background:#F6F5F1;border-left:4px solid #C23A18;padding:16px 20px;font-size:15px;line-height:1.5">' + esc(salidas[ultimo]) + '</div>' +
    '<p style="margin:22px 0 8px;font-weight:bold">Lo que hizo cada agente</p>';
  trabajadores.slice(0, -1).forEach(a => {
    html += '<p style="margin:0 0 4px;font-size:13px;color:#55555F"><b>' + esc(a.nombre) + '</b></p><p style="margin:0 0 14px;font-size:13px">' + esc(salidas[a.nombre]) + '</p>';
  });
  html += '<p style="font-size:12px;color:#8A8A93;margin-top:20px">Monta los textos en tu plantilla de Canva y publícala en tus historias.</p></div>';
  MailApp.sendEmail({ to: para, subject: 'Historia del día · ' + cfg['Droguería'] + ' · ' + estado, htmlBody: html });
}

// ===================== GEMINI =====================

let ultima_ = 0;

function gemini_(sistema, texto, cfg, json) {
  const espera = PAUSA_MS - (Date.now() - ultima_);
  if (espera > 0) Utilities.sleep(espera);
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(cfg['Modelo'] || 'gemini-flash-latest') + ':generateContent';
  const gen = { temperature: json ? 0.2 : 0.8 };
  if (json) gen.responseMimeType = 'application/json';
  const cuerpo = { systemInstruction: { parts: [{ text: sistema }] }, contents: [{ role: 'user', parts: [{ text: texto }] }], generationConfig: gen };
  for (let intento = 1; intento <= 3; intento++) {
    ultima_ = Date.now();
    const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': String(cfg['Clave de Gemini']).trim() }, payload: JSON.stringify(cuerpo) });
    const code = res.getResponseCode();
    if (code === 200) {
      const j = JSON.parse(res.getContentText());
      const t = ((((j.candidates || [])[0] || {}).content || {}).parts || []).map(p => p.text || '').join('').trim();
      if (!json) return t.replace(/\*\*/g, '');
      try { return JSON.parse(t.replace(/^```(json)?/i, '').replace(/```$/, '').trim()); }
      catch (e) { return { veredicto: 'RECHAZADO', motivo: 'El supervisor no respondió en el formato esperado.', devolver_a: '' }; }
    }
    if (code === 429 || code >= 500) { Utilities.sleep(15000 * intento); continue; }
    if (code === 404) throw new Error('El modelo "' + cfg['Modelo'] + '" no existe. Cámbialo en Config.');
    throw new Error('Gemini rechazó la solicitud (código ' + code + '). Revisa la clave en Config.');
  }
  throw new Error('Se agotaron las solicitudes gratuitas de Gemini por ahora. Espera un minuto.');
}

// ===================== UTILIDADES =====================

function hoja_(n) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
  if (!h) throw new Error('Falta la pestaña ' + n + '. Usa Estudio → Instalar.');
  return h;
}

function config_() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_CONFIG);
  const m = {};
  if (h) h.getRange(1, 1, Math.max(h.getLastRow(), 1), 2).getValues().forEach(r => { m[String(r[0]).trim()] = r[1]; });
  return m;
}

function leerAgentes_() {
  const h = hoja_(H_AGENTES);
  const n = h.getLastRow();
  const filas = n < 2 ? [] : h.getRange(2, 1, n - 1, 6).getValues();
  const agentes = filas
    .filter(f => String(f[1]).trim() && String(f[4]).trim() && norm_(f[5]) !== 'no')
    .map(f => ({ orden: Number(f[0]) || 99, nombre: String(f[1]).trim(), tipo: norm_(f[2]), lee: String(f[3]), instrucciones: String(f[4]).trim() }))
    .sort((a, b) => a.orden - b.orden);
  return { trabajadores: agentes.filter(a => a.tipo !== 'supervisor'), supervisor: agentes.find(a => a.tipo === 'supervisor') || null };
}

function cerebro_(cfg) {
  const out = {};
  DOCS.forEach(d => {
    const url = String(cfg[d] || '').trim();
    out[d] = '';
    if (!url) return;
    try { out[d] = DocumentApp.openByUrl(url).getBody().getText().trim(); }
    catch (e) { registrar_(0, 'Motor', 'AVISO', 'No pude abrir el documento "' + d + '". Revisa el enlace en Config.'); }
  });
  return out;
}

function leerCerebro_(lee, cerebro) {
  const pedidos = String(lee || '').split(/[,;\n]/).map(s => norm_(s)).filter(Boolean);
  const textos = DOCS.filter(d => pedidos.some(p => norm_(d).indexOf(p) >= 0 || p.indexOf(norm_(d)) >= 0))
    .map(d => '=== ' + d + ' ===\n' + (cerebro[d] || '(documento vacío o sin conectar)'));
  return textos.join('\n\n') || '(Este agente no lee documentos del cerebro.)';
}

function historiasRecientes_(n) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_HISTORIAS);
  if (!h || h.getLastRow() < 2) return [];
  const desde = Math.max(2, h.getLastRow() - n + 1);
  return h.getRange(desde, 4, h.getLastRow() - desde + 1, 1).getValues().map(r => String(r[0]).slice(0, 200));
}

function registrar_(vuelta, agente, resultado, detalle) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_REGISTRO);
  if (!h) return;
  h.appendRow([new Date(), vuelta, agente, resultado, String(detalle || '').slice(0, 3000)]);
  const f = h.getLastRow(), r = String(resultado);
  if (r.indexOf('APROBADO') === 0) h.getRange(f, 4).setFontColor('#1F7A4D').setFontWeight('bold');
  else if (/^(RECHAZADO|REVISIÓN|ERROR)/.test(r)) h.getRange(f, 4).setFontColor('#C23A18').setFontWeight('bold');
  SpreadsheetApp.flush();
}

function actualizarEstado_() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_CONFIG);
  if (!h) return;
  const cfg = config_();
  const docs = DOCS.filter(d => String(cfg[d] || '').trim()).length;
  let diario = false;
  try { diario = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'historiaDelDia'); } catch (e) {}
  const txt = !cfg['Clave de Gemini'] ? 'Falta la clave de Gemini en B2'
    : 'Estudio activo · ' + docs + ' de 3 documentos conectados · Correo diario: ' + (diario ? 'activado' : 'apagado');
  h.getRange('B12').setValue(txt).setFontWeight('bold').setFontColor(cfg['Clave de Gemini'] ? '#1F7A4D' : '#C23A18');
}

function norm_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim(); }
