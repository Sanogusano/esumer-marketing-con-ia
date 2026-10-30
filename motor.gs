/**
 * RETOMA · Motor de agentes para droguerías
 * Orquestador + 6 agentes: Clasificador, Estratega, Redactor, Auditor y Editor
 * (el agente de Diagnóstico se usa en el chat de Gemini).
 *
 * Instalación: Extensiones → Apps Script → borrar todo → pegar este código →
 * guardar → recargar la hoja → menú Retoma → Instalar.
 */

// ===================== CONFIGURACIÓN =====================

const HOJA_CLIENTES = 'Clientes';
const HOJA_CONFIG = 'Config';
const HOJA_REGISTRO = 'Registro';
const HOJA_PEDIDOS = 'Pedidos';
const HOJA_TABLERO = 'Tablero';

const COLS_BASE = ['ClienteID', 'Nombre', 'Barrio', 'Categoria', 'CicloDias', 'UltimaCompra', 'Compras12m', 'Valor12m',
  'CanalPreferido', 'Domicilio', 'TieneWhatsApp', 'TieneCorreo', 'UsaDescuentos', 'AutorizaContacto'];
const COLS_MOTOR = ['Atraso', 'Estado', 'Frecuencia', 'Valor', 'Campaña', 'Canal', 'Momento', 'Ciclo IA', 'Mensaje', 'Explicación'];
const C = {}; // índice (base 0) de cada columna
COLS_BASE.concat(COLS_MOTOR).forEach((n, i) => { C[n] = i; });
const TOTAL_COLS = COLS_BASE.length + COLS_MOTOR.length;

const DOCUMENTOS = ['Tono de voz', 'PQRS', 'Enfoque', 'Reglas de negocio', 'Cumplimiento'];
const CAMPANAS_ACTIVAS = ['Formulado que no volvió', 'Pediátrico', 'Cosmética', 'Descuento'];
const ESTADOS = ['Al día', 'Atrasado', 'Muy atrasado', 'Perdido'];
const FINALES = ['Aprobado', 'Revisión humana', 'No contactar', 'Sin campaña'];

const MAX_VUELTAS = 3;
const PAUSA_ENTRE_LLAMADAS_MS = 7000;   // el plan gratuito de Gemini permite pocas solicitudes por minuto
const PRESUPUESTO_MS = 4.5 * 60 * 1000; // Apps Script corta a los 6 minutos

const LIMITE_PALABRAS = { 'WhatsApp': 60, 'Llamada': 90, 'Correo': 95, 'Aviso en mostrador': 30 };

// ===================== MENÚ =====================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Retoma')
    .addItem('Instalar', 'instalar')
    .addSeparator()
    .addItem('Procesar clientes en espera', 'procesarAhora')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Procesar campaña (10 clientes)')
      .addItem('Formulado que no volvió', 'campanaFormulado')
      .addItem('Pediátrico', 'campanaPediatrico')
      .addItem('Cosmética', 'campanaCosmetica')
      .addItem('Descuento', 'campanaDescuento'))
    .addSeparator()
    .addItem('Actualizar cerebro', 'actualizarCerebro')
    .addItem('Actualizar tablero', 'actualizarTablero')
    .addToUi();
}

// ===================== INSTALACIÓN =====================

function instalar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. Hoja de clientes
  let hc = ss.getSheetByName(HOJA_CLIENTES);
  if (!hc) {
    hc = ss.getSheets().find(s => String(s.getRange(1, 1).getValue()).trim() === 'ClienteID');
    if (!hc) { ui.alert('No encuentro la base. Importa primero clientes-drogueria.csv (la celda A1 debe decir ClienteID).'); return; }
    hc.setName(HOJA_CLIENTES);
  }
  hc.getRange(1, COLS_BASE.length + 1, 1, COLS_MOTOR.length).setValues([COLS_MOTOR])
    .setBackground('#16161D').setFontColor('#FFFFFF').setFontWeight('bold');
  hc.getRange(1, 1, 1, COLS_BASE.length).setBackground('#ECEAE4').setFontWeight('bold');
  hc.setFrozenRows(1);
  hc.getRange(2, C['Valor12m'] + 1, Math.max(hc.getMaxRows() - 1, 1), 1).setNumberFormat('$#,##0');

  // 2. Config
  let hcfg = ss.getSheetByName(HOJA_CONFIG);
  if (!hcfg) {
    hcfg = ss.insertSheet(HOJA_CONFIG);
    const filas = [
      ['Dato', 'Valor', 'Ayuda'],
      ['Clave', '', 'Pega aquí tu clave de Gemini (aistudio.google.com/apikey).'],
      ['Modelo', 'gemini-flash-latest', 'Modelo de Gemini que usan los agentes.'],
      ['Fecha de corte', new Date(2026, 9, 31), 'Con esta fecha se calcula el atraso de cada cliente.'],
      ['Ciclo por defecto', 45, 'Días que se usan cuando el cliente no tiene ciclo.'],
      ['Tono de voz', '', 'Enlace al documento. Si está vacío se usa el cerebro de ejemplo.'],
      ['PQRS', '', 'Enlace al documento.'],
      ['Enfoque', '', 'Enlace al documento.'],
      ['Reglas de negocio', '', 'Enlace al documento.'],
      ['Cumplimiento', '', 'Enlace al documento.'],
      ['Estado', '', '']
    ];
    hcfg.getRange(1, 1, filas.length, 3).setValues(filas);
    hcfg.getRange('B4').setNumberFormat('dd/mm/yyyy');
    hcfg.getRange('A1:C1').setBackground('#16161D').setFontColor('#FFFFFF').setFontWeight('bold');
    hcfg.getRange('A2:A11').setFontWeight('bold');
    hcfg.setColumnWidth(1, 170); hcfg.setColumnWidth(2, 420); hcfg.setColumnWidth(3, 460);
    hcfg.setFrozenRows(1);
  }

  // 3. Registro y Pedidos
  if (!ss.getSheetByName(HOJA_REGISTRO)) {
    const hr = ss.insertSheet(HOJA_REGISTRO);
    hr.getRange(1, 1, 1, 7).setValues([['Fecha', 'ClienteID', 'Nombre', 'Vuelta', 'Agente', 'Resultado', 'Detalle']])
      .setBackground('#16161D').setFontColor('#FFFFFF').setFontWeight('bold');
    hr.setFrozenRows(1); hr.setColumnWidth(6, 260); hr.setColumnWidth(7, 620);
  }
  if (!ss.getSheetByName(HOJA_PEDIDOS)) {
    const hp = ss.insertSheet(HOJA_PEDIDOS);
    hp.getRange(1, 1, 1, 5).setValues([['Pedido', 'Estado', 'Documento', 'Cambio aplicado', 'Fecha']])
      .setBackground('#16161D').setFontColor('#FFFFFF').setFontWeight('bold');
    hp.setFrozenRows(1); hp.setColumnWidth(1, 460); hp.setColumnWidth(4, 460);
  }
  if (!ss.getSheetByName(HOJA_TABLERO)) ss.insertSheet(HOJA_TABLERO);

  // 4. Cálculos de la base: los clientes que ya existen quedan como "Base"
  recalcular_(true);
  actualizarTablero();

  // 5. Disparadores
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('alCambiar').forSpreadsheet(ss).onChange().create();

  PropertiesService.getDocumentProperties().setProperty('instalado', 'si');
  actualizarEstado_();
  ss.setActiveSheet(ss.getSheetByName(HOJA_CONFIG));
  ui.alert('Retoma quedó instalado. Ahora pega tu clave de Gemini en la pestaña Config, celda B2.');
}

// ===================== DISPARADORES =====================

function alCambiar(e) {
  if (PropertiesService.getDocumentProperties().getProperty('instalado') !== 'si') return;
  actualizarEstado_();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(500)) { marcarNuevos_(); return; } // ya hay un ciclo corriendo: él los recoge
  let nuevos = 0;
  try { nuevos = recalcular_(false); } finally { lock.releaseLock(); }
  if (nuevos > 0 || hayPendientes_()) procesarCola_();
}

/** Marca como "En espera" las filas nuevas sin tocar nada más. */
function marcarNuevos_() {
  const hc = hoja_(HOJA_CLIENTES);
  leerClientes_(hc).forEach((f, i) => {
    if (String(f[C['ClienteID']]).trim() && String(f[C['Categoria']]).trim() && !String(f[C['Ciclo IA']]).trim())
      hc.getRange(i + 2, C['Ciclo IA'] + 1).setValue('En espera');
  });
}

function continuar() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'continuar').forEach(t => ScriptApp.deleteTrigger(t));
  procesarCola_();
}

function procesarAhora() {
  marcarNuevos_();
  const n = procesarCola_();
  SpreadsheetApp.getActiveSpreadsheet().toast(n + ' cliente(s) procesado(s).', 'Retoma', 5);
}

function campanaFormulado() { campana_('Formulado que no volvió'); }
function campanaPediatrico() { campana_('Pediátrico'); }
function campanaCosmetica() { campana_('Cosmética'); }
function campanaDescuento() { campana_('Descuento'); }

function campana_(nombre) {
  const hc = hoja_(HOJA_CLIENTES);
  const datos = leerClientes_(hc);
  const elegidos = datos
    .map((f, i) => ({ f: f, fila: i + 2 }))
    .filter(x => x.f[C['Campaña']] === nombre && x.f[C['Ciclo IA']] === 'Base')
    .sort((a, b) => num_(b.f[C['Valor12m']]) - num_(a.f[C['Valor12m']]))
    .slice(0, 10);
  if (!elegidos.length) {
    SpreadsheetApp.getUi().alert('No quedan clientes sin procesar en la campaña ' + nombre + '.');
    return;
  }
  elegidos.forEach(x => hc.getRange(x.fila, C['Ciclo IA'] + 1).setValue('En espera'));
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast(elegidos.length + ' clientes en cola. Van saliendo uno por uno.', 'Retoma', 8);
  procesarCola_();
}

// ===================== CÁLCULOS (HERRAMIENTA, SIN IA) =====================

/** Calcula atraso, estado, frecuencia, valor, campaña y canal. Devuelve cuántos clientes nuevos encontró. */
function recalcular_(esInstalacion) {
  const hc = hoja_(HOJA_CLIENTES);
  const datos = leerClientes_(hc);
  if (!datos.length) return 0;
  const cfg = config_();
  let nuevos = 0;
  const marcar = [];
  const salida = datos.map((f, i) => {
    const fuera = f.slice(COLS_BASE.length, COLS_BASE.length + 6);
    if (!String(f[C['ClienteID']]).trim() || !String(f[C['Categoria']]).trim()) return fuera;
    const cli = cliente_(f);
    const calc = calcular(cli, cfg.corte, cfg.cicloDefecto);
    const ciclo = String(f[C['Ciclo IA']]).trim();
    fuera[0] = calc.atraso; fuera[1] = calc.estado; fuera[2] = calc.frecuencia; fuera[3] = calc.valor;
    if (!FINALES.includes(ciclo) && ciclo !== 'Procesando') { // lo que decidieron los agentes no se toca
      fuera[4] = calc.campana; fuera[5] = calc.canal;
    }
    if (!ciclo) { marcar.push(i + 2); if (!esInstalacion) nuevos++; }
    return fuera;
  });
  hc.getRange(2, COLS_BASE.length + 1, salida.length, 6).setValues(salida);
  if (esInstalacion && marcar.length === datos.length) {
    hc.getRange(2, C['Ciclo IA'] + 1, datos.length, 1).setValue('Base');
  } else {
    marcar.forEach(fila => hc.getRange(fila, C['Ciclo IA'] + 1).setValue(esInstalacion ? 'Base' : 'En espera'));
  }
  return nuevos;
}

function calcular(cli, corte, cicloDefecto) {
  const ultima = fecha_(cli.UltimaCompra);
  const dias = ultima ? Math.round((corte - ultima) / 86400000) : 9999;
  const ciclo = num_(cli.CicloDias) || cicloDefecto;
  const atraso = dias - ciclo;
  const estado = atraso <= 0 ? 'Al día' : atraso <= 30 ? 'Atrasado' : atraso <= 90 ? 'Muy atrasado' : 'Perdido';
  const compras = num_(cli.Compras12m), valor12 = num_(cli.Valor12m);
  const frecuencia = compras >= 10 ? 'Alta' : compras >= 4 ? 'Media' : 'Baja';
  const valor = valor12 >= 1500000 ? 'Alto' : valor12 >= 500000 ? 'Medio' : 'Bajo';
  const campana = campanaPorReglas(cli, estado);
  return { dias: dias, ciclo: ciclo, atraso: atraso, estado: estado, frecuencia: frecuencia, valor: valor,
    campana: campana, canal: canalPorReglas(cli, campana) };
}

function campanaPorReglas(cli, estado) {
  if (!si_(cli.AutorizaContacto)) return 'No contactar';
  if (estado === 'Al día') return 'Sin campaña';
  const cat = norm_(cli.Categoria);
  if (cat === 'formulado') return 'Formulado que no volvió';
  if (cat === 'pediatrico') return 'Pediátrico';
  if (cat === 'cosmetica') return 'Cosmética';
  if (si_(cli.UsaDescuentos)) return 'Descuento';
  return 'Sin campaña';
}

function canalPorReglas(cli, campana) {
  if (campana === 'No contactar') return 'No contactar';
  const pref = norm_(cli.CanalPreferido);
  if (pref === 'mostrador') return 'Aviso en mostrador';
  if (pref === 'whatsapp') return si_(cli.TieneWhatsApp) ? 'WhatsApp' : 'Llamada';
  if (pref === 'correo') return si_(cli.TieneCorreo) ? 'Correo' : 'Llamada';
  return 'Llamada';
}

/** Campañas que puede recibir un cliente según sus datos. Lo usa el Auditor. */
function campanasPermitidas(cli, estado) {
  if (!si_(cli.AutorizaContacto)) return ['No contactar'];
  if (estado === 'Al día') return ['Sin campaña'];
  const ok = ['Sin campaña'];
  const cat = norm_(cli.Categoria);
  if (cat === 'formulado') ok.push('Formulado que no volvió');
  if (cat === 'pediatrico') ok.push('Pediátrico');
  if (cat === 'cosmetica') ok.push('Cosmética');
  if (si_(cli.UsaDescuentos)) ok.push('Descuento');
  return ok;
}

function canalesPosibles(cli) {
  const ok = ['Llamada', 'Aviso en mostrador'];
  if (si_(cli.TieneWhatsApp)) ok.push('WhatsApp');
  if (si_(cli.TieneCorreo)) ok.push('Correo');
  return ok;
}

// ===================== ORQUESTADOR =====================

function procesarCola_() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) return 0;
  const inicio = Date.now();
  let procesados = 0;
  try {
    const cfg = config_();
    if (!cfg.clave) { actualizarEstado_(); return 0; }
    const hc = hoja_(HOJA_CLIENTES);
    const cerebro = cerebro_(cfg);
    while (Date.now() - inicio < PRESUPUESTO_MS) {
      const datos = leerClientes_(hc);
      const i = datos.findIndex(f => f[C['Ciclo IA']] === 'En espera');
      if (i < 0) break;
      const fila = i + 2;
      hc.getRange(fila, C['Ciclo IA'] + 1).setValue('Procesando');
      SpreadsheetApp.flush();
      try {
        const r = cicloCliente_(cliente_(datos[i]), cfg, cerebro);
        hc.getRange(fila, C['Campaña'] + 1, 1, 6).setValues([[r.campana, r.canal, r.momento, r.final, r.mensaje, r.explicacion]]);
      } catch (err) {
        hc.getRange(fila, C['Ciclo IA'] + 1, 1, 3).setValues([['Revisión humana', '', 'Error: ' + err.message]]);
        registrar_(datos[i][C['ClienteID']], datos[i][C['Nombre']], '-', 'Orquestador', 'ERROR', err.message);
        if (/clave|API key|403|400/i.test(err.message)) break;
      }
      procesados++;
      SpreadsheetApp.flush();
    }
    actualizarTablero();
  } finally {
    lock.releaseLock();
  }
  if (hayPendientes_()) programarContinuacion_();
  return procesados;
}

function cicloCliente_(cli, cfg, cerebro) {
  const calc = calcular(cli, cfg.corte, cfg.cicloDefecto);
  const id = cli.ClienteID, nom = cli.Nombre;
  registrar_(id, nom, 0, 'Orquestador', 'INICIO', 'Atraso ' + calc.atraso + ' días · ciclo ' + calc.ciclo + ' · ' + cli.Categoria);

  // 1. Clasificador
  const clas = agenteClasificador_(cli, calc, cfg);
  if (clas.estado !== calc.estado) {
    registrar_(id, nom, 0, 'Orquestador', 'CORRIGE', 'El Clasificador dijo "' + clas.estado + '"; según el criterio es "' + calc.estado + '".');
    clas.estado = calc.estado;
  }
  registrar_(id, nom, 0, 'Clasificador', clas.estado, clas.explicacion + (clas.alerta ? ' · Alerta: ' + clas.alerta : ''));

  // 2. Estratega
  let est = agenteEstratega_(cli, calc, clas, cerebro, cfg, null);
  registrar_(id, nom, 1, 'Estratega', est.campana + ' · ' + est.canal, est.justificacion + ' · Momento: ' + est.momento);

  let red = null, vuelta = 1, aprobado = false, motivo = '';
  while (vuelta <= MAX_VUELTAS) {
    if (est.campana === 'No contactar' || est.campana === 'Sin campaña') {
      const aud0 = auditarEstrategia_(cli, calc, est);
      if (aud0.ok) {
        registrar_(id, nom, vuelta, 'Auditor', 'APROBADO', 'No se envía mensaje: ' + est.campana + '.');
        return { campana: est.campana, canal: est.campana === 'No contactar' ? 'No contactar' : '—', momento: '—',
          final: est.campana, mensaje: '', explicacion: clas.explicacion };
      }
      registrar_(id, nom, vuelta, 'Auditor', 'RECHAZADO → Estratega', aud0.motivo);
      vuelta++;
      if (vuelta > MAX_VUELTAS) { motivo = aud0.motivo; break; }
      est = agenteEstratega_(cli, calc, clas, cerebro, cfg, aud0.motivo);
      registrar_(id, nom, vuelta, 'Estratega', est.campana + ' · ' + est.canal, est.justificacion);
      continue;
    }

    // 3. Redactor
    if (!red || red.devolver === 'Redactor') {
      red = agenteRedactor_(cli, calc, clas, est, cerebro, cfg, red ? red.motivo : null, red ? red.texto : null);
      registrar_(id, nom, vuelta, 'Redactor', est.canal, red.texto);
    }

    // 4. Auditor: primero las reglas fijas, luego la IA
    let aud = auditarEstrategia_(cli, calc, est);
    if (aud.ok) aud = auditarTexto_(red.texto, est.canal);
    if (aud.ok) aud = agenteAuditor_(cli, est, red.texto, cerebro, cfg);

    if (aud.ok) {
      registrar_(id, nom, vuelta, 'Auditor', 'APROBADO', aud.motivo || 'Cumple el cerebro.');
      aprobado = true;
      break;
    }
    registrar_(id, nom, vuelta, 'Auditor', 'RECHAZADO → ' + aud.devolver, aud.motivo);
    motivo = aud.motivo;
    vuelta++;
    if (vuelta > MAX_VUELTAS) break;
    if (aud.devolver === 'Estratega') {
      est = agenteEstratega_(cli, calc, clas, cerebro, cfg, aud.motivo);
      registrar_(id, nom, vuelta, 'Estratega', est.campana + ' · ' + est.canal, est.justificacion);
      red = null; // con una estrategia nueva se escribe de cero
    } else {
      red.devolver = 'Redactor'; red.motivo = aud.motivo;
    }
  }

  if (!aprobado) registrar_(id, nom, MAX_VUELTAS, 'Orquestador', 'REVISIÓN HUMANA', 'No pasó en ' + MAX_VUELTAS + ' vueltas. Último motivo: ' + motivo);
  return { campana: est.campana, canal: est.canal, momento: est.momento,
    final: aprobado ? 'Aprobado' : 'Revisión humana', mensaje: red ? red.texto : '', explicacion: clas.explicacion };
}

// ===================== AGENTES =====================

function agenteClasificador_(cli, calc, cfg) {
  const sistema = 'Eres el Agente Clasificador de una droguería de barrio en Medellín. ' +
    'Recibes la ficha de un cliente y unos cálculos ya hechos. No recalcules: interpreta. ' +
    'Criterio de estado según días de atraso: 0 o menos "Al día"; 1 a 30 "Atrasado"; 31 a 90 "Muy atrasado"; más de 90 "Perdido". ' +
    'Responde SOLO un JSON: {"estado": "...", "explicacion": "una línea en español sencillo para el dueño", "alerta": "algo raro en los datos o vacío"}. ' +
    'Nunca menciones medicamentos ni enfermedades.';
  const r = gemini_(sistema, { cliente: fichaSegura_(cli), calculos: calc }, cfg);
  return { estado: String(r.estado || ''), explicacion: String(r.explicacion || ''), alerta: String(r.alerta || '') };
}

function agenteEstratega_(cli, calc, clas, cerebro, cfg, correccion) {
  const sistema = 'Eres el Agente Estratega de una droguería de barrio. Decides campaña, canal y momento de contacto.\n' +
    'Campañas posibles: "Formulado que no volvió", "Pediátrico", "Cosmética", "Descuento", "No contactar", "Sin campaña".\n' +
    'Reglas fijas, en orden: si AutorizaContacto es "No" → "No contactar". Si está "Al día" → "Sin campaña". ' +
    'Formulado atrasado → "Formulado que no volvió"; Pediátrico atrasado → "Pediátrico"; Cosmética atrasada → "Cosmética"; ' +
    'si usa descuentos → "Descuento"; si no aplica ninguna → "Sin campaña".\n' +
    'Canales posibles: "WhatsApp" (solo si tiene WhatsApp), "Correo" (solo si tiene correo), "Llamada", "Aviso en mostrador", "No contactar". ' +
    'Respeta el canal preferido del cliente si es posible.\n' +
    'Usa estos documentos del cerebro de la droguería:\n\n' + cerebro['Enfoque'] + '\n\n' + cerebro['Reglas de negocio'] + '\n\n' +
    'Responde SOLO un JSON: {"campana": "...", "canal": "...", "momento": "día y franja horaria", "justificacion": "una línea"}.';
  const entrada = { cliente: fichaSegura_(cli), calculos: calc, clasificacion: clas };
  if (correccion) entrada.el_auditor_te_devolvio_esto = correccion;
  const r = gemini_(sistema, entrada, cfg);
  r.campana = canonico_(r.campana, CAMPANAS_ACTIVAS.concat(['No contactar', 'Sin campaña']));
  r.canal = canonico_(r.canal, ['WhatsApp', 'Correo', 'Llamada', 'Aviso en mostrador', 'No contactar']);
  return { campana: String(r.campana || ''), canal: String(r.canal || ''), momento: String(r.momento || ''), justificacion: String(r.justificacion || '') };
}

function agenteRedactor_(cli, calc, clas, est, cerebro, cfg, correccion, anterior) {
  const objetivos = {
    'Formulado que no volvió': 'recordarle que ya le toca lo de cada mes y ofrecerle domicilio',
    'Pediátrico': 'ayudar a la mamá o al papá a no quedarse sin lo del niño',
    'Cosmética': 'contarle una novedad de su tipo de producto',
    'Descuento': 'darle un beneficio concreto por volver, sin pasar el tope de las reglas'
  };
  const formatos = {
    'WhatsApp': 'un mensaje de WhatsApp de máximo 60 palabras',
    'Llamada': 'un guion de llamada de máximo 90 palabras con saludo, motivo, una pregunta y cierre',
    'Correo': 'un correo: primera línea "Asunto: ...", luego el cuerpo de máximo 80 palabras',
    'Aviso en mostrador': 'una frase de máximo 30 palabras para decirle cuando venga a la droguería'
  };
  const sistema = 'Eres el Agente Redactor de una droguería de barrio. Escribes como el dueño, no como una empresa.\n' +
    'Sigue al pie de la letra estos documentos del cerebro:\n\n' + cerebro['Tono de voz'] + '\n\n' + cerebro['PQRS'] + '\n\n' + cerebro['Reglas de negocio'] + '\n\n' +
    'Nunca nombres medicamentos, dosis, enfermedades ni condiciones de salud. Nunca recomiendes tratamientos. Frases cortas.\n' +
    'Responde SOLO un JSON: {"texto": "..."}.';
  const entrada = {
    cliente: { nombre: primerNombre_(cli.Nombre), nombre_completo: cli.Nombre, barrio: cli.Barrio, domicilio: cli.Domicilio },
    campana: est.campana, objetivo: objetivos[est.campana] || 'invitarlo a volver',
    entregable: formatos[est.canal] || formatos['WhatsApp'], momento: est.momento
  };
  if (correccion) { entrada.version_anterior = anterior; entrada.el_auditor_la_rechazo_por = correccion; }
  const r = gemini_(sistema, entrada, cfg);
  return { texto: String(r.texto || '').trim(), devolver: null, motivo: null };
}

function agenteAuditor_(cli, est, texto, cerebro, cfg) {
  const sistema = 'Eres el Agente Auditor de una droguería. Revisas un mensaje antes de que salga.\n' +
    'Documentos con las reglas que debes hacer cumplir:\n\n' + cerebro['Cumplimiento'] + '\n\n' + cerebro['Reglas de negocio'] + '\n\n' + cerebro['Tono de voz'] + '\n\n' +
    'Rechaza si: nombra medicamentos, dosis, enfermedades o condiciones de salud; recomienda tratamientos; promete algo que las reglas no permiten; ' +
    'usa frases de robot como "espero que te encuentres bien", "no dudes en" o "solución"; o no suena al tono de voz.\n' +
    'Si el problema es el texto, devuélvelo al "Redactor". Si el problema es la campaña o el canal, devuélvelo al "Estratega".\n' +
    'Responde SOLO un JSON: {"veredicto": "APROBADO" o "RECHAZADO", "motivo": "una línea concreta", "devolver_a": "Redactor" o "Estratega" o ""}.';
  const r = gemini_(sistema, { campana: est.campana, canal: est.canal, mensaje: texto }, cfg);
  const ok = String(r.veredicto || '').toUpperCase().indexOf('APROBADO') === 0;
  return { ok: ok, motivo: String(r.motivo || ''), devolver: r.devolver_a === 'Estratega' ? 'Estratega' : 'Redactor' };
}

/** Reglas fijas del Auditor sobre la estrategia (sin IA). */
function auditarEstrategia_(cli, calc, est) {
  const permitidas = campanasPermitidas(cli, calc.estado);
  if (permitidas.indexOf(est.campana) < 0)
    return { ok: false, devolver: 'Estratega', motivo: 'La campaña "' + est.campana + '" no corresponde a este cliente. Puede ser: ' + permitidas.join(', ') + '.' };
  if (est.campana === 'No contactar' || est.campana === 'Sin campaña') return { ok: true };
  if (canalesPosibles(cli).indexOf(est.canal) < 0)
    return { ok: false, devolver: 'Estratega', motivo: 'El canal "' + est.canal + '" no es posible para este cliente. Puede ser: ' + canalesPosibles(cli).join(', ') + '.' };
  return { ok: true };
}

/** Reglas fijas del Auditor sobre el texto (sin IA). */
function auditarTexto(texto, canal) {
  if (!texto) return { ok: false, devolver: 'Redactor', motivo: 'El mensaje llegó vacío.' };
  const palabras = texto.split(/\s+/).filter(Boolean).length;
  const limite = LIMITE_PALABRAS[canal] || 60;
  if (palabras > limite) return { ok: false, devolver: 'Redactor', motivo: 'Tiene ' + palabras + ' palabras y el máximo para ' + canal + ' es ' + limite + '.' };
  const t = ' ' + norm_(texto) + ' ';
  const salud = ['losartan', 'metformina', 'acetaminofen', 'ibuprofeno', 'insulina', 'omeprazol', 'enalapril', 'atorvastatina',
    'levotiroxina', 'amoxicilina', 'salbutamol', 'loratadina', 'naproxeno', 'diclofenaco', 'antibiotico', 'anticonceptiv',
    'presion', 'hipertension', 'diabetes', 'diabetic', 'colesterol', 'tiroides', 'asma', 'depresion', 'ansiedad', 'embarazo',
    'cancer', 'infeccion', 'enfermedad', 'tratamiento', 'medicamento', 'pastilla', 'jarabe'];
  const hallado = salud.find(w => t.indexOf(w) >= 0);
  if (hallado) return { ok: false, devolver: 'Redactor', motivo: 'Menciona un dato de salud ("' + hallado + '"). Está prohibido por Cumplimiento.' };
  if (/\d+\s?(mg|ml|mcg|gr|tabletas|capsulas|gotas)\b/.test(t)) return { ok: false, devolver: 'Redactor', motivo: 'Menciona una dosis.' };
  const robot = ['no dudes en', 'espero que te encuentres', 'espero que se encuentre', 'solucion', 'quedamos atentos'];
  const r = robot.find(w => t.indexOf(w) >= 0);
  if (r) return { ok: false, devolver: 'Redactor', motivo: 'Usa una frase de robot: "' + r + '".' };
  return { ok: true };
}
function auditarTexto_(texto, canal) { return auditarTexto(texto, canal); }

// ===================== EDITOR DEL CEREBRO =====================

function actualizarCerebro() {
  const cfg = config_();
  const ui = SpreadsheetApp.getUi();
  if (!cfg.clave) { ui.alert('Falta la clave de Gemini en Config.'); return; }
  const hp = hoja_(HOJA_PEDIDOS);
  const n = hp.getLastRow();
  if (n < 2) { ui.alert('Escribe un pedido en la pestaña Pedidos, columna A.'); return; }
  const filas = hp.getRange(2, 1, n - 1, 2).getValues();
  let hechos = 0;
  filas.forEach((f, k) => {
    const pedido = String(f[0]).trim();
    if (!pedido || String(f[1]).trim()) return;
    const fila = k + 2;
    const cerebro = cerebro_(cfg);
    const sistema = 'Eres el Agente Editor del cerebro de una droguería. El cerebro son 5 documentos en Markdown: ' + DOCUMENTOS.join(', ') + '.\n' +
      'Recibes un pedido del dueño. Decide qué documento cambia, aplica el cambio y devuelve el documento COMPLETO actualizado, ' +
      'conservando su estructura (# título, ## secciones, reglas con guiones) y todo lo que no cambia. ' +
      'Si el pedido contradice una regla, reemplázala. Nunca quites reglas de Cumplimiento sobre datos de salud ni autorización de contacto.\n' +
      'Responde SOLO un JSON: {"documento": "uno de los 5 nombres exactos", "texto_nuevo": "documento completo", "resumen": "qué cambiaste, en una línea"}.';
    try {
      const r = gemini_(sistema, { pedido: pedido, cerebro: cerebro }, cfg);
      const doc = String(r.documento || '');
      if (DOCUMENTOS.indexOf(doc) < 0) throw new Error('El Editor no identificó el documento.');
      const url = cfg.docs[doc];
      if (!url) {
        hp.getRange(fila, 2, 1, 4).setValues([['Sin documento', doc, 'Conecta el documento "' + doc + '" en Config y vuelve a intentar.', new Date()]]);
        return;
      }
      DocumentApp.openByUrl(url).getBody().setText(String(r.texto_nuevo || ''));
      hp.getRange(fila, 2, 1, 4).setValues([['Aplicado', doc, String(r.resumen || ''), new Date()]]);
      registrar_('-', '-', '-', 'Editor', 'Actualizó ' + doc, String(r.resumen || ''));
      hechos++;
    } catch (err) {
      hp.getRange(fila, 2, 1, 4).setValues([['Error', '', err.message, new Date()]]);
    }
  });
  ui.alert(hechos ? 'Listo: ' + hechos + ' cambio(s) aplicado(s). Mira la pestaña Pedidos.' : 'No se aplicó ningún cambio. Mira la pestaña Pedidos.');
}

// ===================== TABLERO =====================

function actualizarTablero() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ht = ss.getSheetByName(HOJA_TABLERO) || ss.insertSheet(HOJA_TABLERO);
  const datos = leerClientes_(hoja_(HOJA_CLIENTES)).filter(f => String(f[C['ClienteID']]).trim());
  const total = datos.length || 1;
  const contar = (col, val) => datos.filter(f => f[C[col]] === val).length;
  const sumaMes = filtro => Math.round(datos.filter(filtro).reduce((s, f) => s + num_(f[C['Valor12m']]), 0) / 12);

  const formulados = datos.filter(f => f[C['Campaña']] === 'Formulado que no volvió');
  const enRiesgo = sumaMes(f => f[C['Campaña']] === 'Formulado que no volvió');
  const potencial = Math.round(0.5 * sumaMes(f => CAMPANAS_ACTIVAS.indexOf(f[C['Campaña']]) >= 0));
  const autoriza = datos.filter(f => si_(f[C['AutorizaContacto']])).length;

  ht.clear();
  const v = [
    ['RETOMA · TABLERO', '', '', ''],
    ['Actualizado', new Date(), '', ''],
    ['', '', '', ''],
    ['Indicador', 'Valor', '', ''],
    ['Clientes en la base', datos.length, '', ''],
    ['Formulados que no volvieron', formulados.length, '', ''],
    ['Valor mensual en riesgo (formulados)', enRiesgo, '', ''],
    ['Potencial de recuperación al mes (50 %)', potencial, '', ''],
    ['Autorizan contacto', autoriza / total, '', ''],
    ['Mensajes aprobados', contar('Ciclo IA', 'Aprobado'), '', ''],
    ['En revisión humana', contar('Ciclo IA', 'Revisión humana'), '', ''],
    ['', '', '', ''],
    ['Estado de recompra', 'Clientes', 'Campaña', 'Clientes']
  ];
  const camps = CAMPANAS_ACTIVAS.concat(['No contactar', 'Sin campaña']);
  for (let k = 0; k < 6; k++) {
    v.push([ESTADOS[k] || '', ESTADOS[k] ? contar('Estado', ESTADOS[k]) : '', camps[k], contar('Campaña', camps[k])]);
  }
  v.push(['', '', '', '']);
  v.push(['Canal sugerido', 'Clientes', '', '']);
  ['WhatsApp', 'Llamada', 'Correo', 'Aviso en mostrador', 'No contactar'].forEach(c => v.push([c, contar('Canal', c), '', '']));
  ht.getRange(1, 1, v.length, 4).setValues(v);

  ht.getRange('A1').setFontSize(16).setFontWeight('bold');
  ht.getRange('B2').setNumberFormat('dd/mm/yyyy hh:mm');
  ht.getRange('A4:B4').setBackground('#16161D').setFontColor('#FFFFFF').setFontWeight('bold');
  ht.getRange('A13:D13').setBackground('#16161D').setFontColor('#FFFFFF').setFontWeight('bold');
  ht.getRange('A21:B21').setBackground('#16161D').setFontColor('#FFFFFF').setFontWeight('bold');
  ht.getRange('B7:B8').setNumberFormat('$#,##0');
  ht.getRange('B9').setNumberFormat('0%');
  ht.getRange('B5:B11').setFontWeight('bold').setFontSize(13);
  ht.setColumnWidth(1, 300); ht.setColumnWidth(2, 140); ht.setColumnWidth(3, 230); ht.setColumnWidth(4, 110);

  if (!ht.getCharts().length) {
    ht.insertChart(ht.newChart().asBarChart().addRange(ht.getRange('A13:B17'))
      .setPosition(1, 6, 0, 0).setOption('title', 'Clientes por estado de recompra')
      .setOption('legend', { position: 'none' }).setOption('colors', ['#C23A18']).setOption('width', 520).setOption('height', 280).build());
    ht.insertChart(ht.newChart().asBarChart().addRange(ht.getRange('C13:D19'))
      .setPosition(16, 6, 0, 0).setOption('title', 'Clientes por campaña')
      .setOption('legend', { position: 'none' }).setOption('colors', ['#16161D']).setOption('width', 520).setOption('height', 300).build());
  }
}

// ===================== GEMINI =====================

let ultimaLlamada_ = 0;

function gemini_(sistema, entrada, cfg) {
  const espera = PAUSA_ENTRE_LLAMADAS_MS - (Date.now() - ultimaLlamada_);
  if (espera > 0) Utilities.sleep(espera);
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(cfg.modelo) + ':generateContent';
  const cuerpo = {
    systemInstruction: { parts: [{ text: sistema }] },
    contents: [{ role: 'user', parts: [{ text: JSON.stringify(entrada) }] }],
    generationConfig: { temperature: 0.6, responseMimeType: 'application/json' }
  };
  for (let intento = 1; intento <= 3; intento++) {
    ultimaLlamada_ = Date.now();
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': cfg.clave }, payload: JSON.stringify(cuerpo)
    });
    const code = res.getResponseCode();
    if (code === 200) {
      const j = JSON.parse(res.getContentText());
      const txt = (((j.candidates || [])[0] || {}).content || {}).parts;
      const t = (txt || []).map(p => p.text || '').join('').replace(/^```(json)?/i, '').replace(/```$/, '').trim();
      try { return JSON.parse(t); } catch (e) { throw new Error('Gemini no devolvió un JSON válido.'); }
    }
    if (code === 429 || code >= 500) { Utilities.sleep(20000 * intento); continue; }
    if (code === 400 || code === 403) throw new Error('Gemini rechazó la clave o el modelo (código ' + code + '). Revisa Config.');
    if (code === 404) throw new Error('El modelo "' + cfg.modelo + '" no existe. Cámbialo en Config.');
    throw new Error('Gemini respondió con código ' + code + '.');
  }
  throw new Error('Se agotaron las solicitudes gratuitas por ahora (429). Espera un minuto.');
}

// ===================== CEREBRO DE EJEMPLO =====================

const CEREBRO_EJEMPLO = {
  'Tono de voz': '# Tono de voz\n## Para qué sirve\nQue los mensajes suenen al dueño de la Droguería San Rafael, en Belén.\n## Reglas\n- Tratamos de usted a los mayores de 50 y de tú a los jóvenes.\n- Saludamos por el primer nombre: "Doña Gloria", "Don Jorge", "Camila".\n- Frases cortas, como se habla en el mostrador.\n- Cerramos con una pregunta fácil de responder.\n## Ejemplos\n- Sí: "Doña Gloria, ¿cómo me le va? ¿Se lo alisto?"\n- Sí: "Quiubo Camila, te guardé lo tuyo."\n- No: "Estimado cliente, le informamos..."\n- No: "¡Hola! 😊 Esperamos que estés genial."\n## Nunca\n- Emojis de más de uno por mensaje.\n- "No dudes en", "solución", "quedamos atentos".',
  'PQRS': '# PQRS\n## Para qué sirve\nCómo respondemos quejas y reclamos.\n## Reglas\n- Si el cliente tuvo un problema, primero pedimos disculpas con su nombre.\n- Producto vencido o en mal estado: se cambia el mismo día, sin pedir factura.\n- Domicilio tarde: el siguiente domicilio va gratis.\n- Nunca discutimos por WhatsApp: si se complica, llama el dueño.\n## Ejemplos\n- Sí: "Don Jorge, qué pena con usted. Hoy mismo se lo cambiamos."\n- No: "Su caso fue radicado."\n## Nunca\n- Prometer reembolsos en efectivo.',
  'Enfoque': '# Enfoque\n## Para qué sirve\nA quién queremos recuperar primero.\n## Reglas\n- Prioridad 1: formulados atrasados, porque compran cada mes.\n- Prioridad 2: pediátricos, porque la familia entera compra donde compra la mamá.\n- Prioridad 3: cosmética y descuentos.\n- Nos diferencia: conocemos al cliente por el nombre y llevamos el domicilio rápido.\n- Formulados: martes a jueves de 9 a 11 a.m. Pediátrico: sábados en la mañana. Cosmética: viernes en la tarde.\n## Ejemplos\n- Sí: empezar por el formulado de más valor.\n- No: mandar descuentos a quien compra igual sin descuento.\n## Nunca\n- Competir por precio con las cadenas.',
  'Reglas de negocio': '# Reglas de negocio\n## Para qué sirve\nLo que la droguería puede ofrecer.\n## Reglas\n- Horario: lunes a sábado de 7 a.m. a 9 p.m.; domingos de 8 a.m. a 2 p.m.\n- Domicilio gratis en compras desde $40.000; si no, cuesta $4.000.\n- Barrios con domicilio: Belén, Laureles, La América y Guayabal.\n- Descuento máximo: 10 %. Solo para clientes que usan descuentos.\n- Afiliados a Coopidrogas: podemos encargar lo que no tengamos para el día siguiente.\n## Ejemplos\n- Sí: "Se lo llevo hoy mismo a la casa."\n- No: "Le damos 30 % de descuento."\n## Nunca\n- Prometer un producto agotado.',
  'Cumplimiento': '# Cumplimiento\n## Para qué sirve\nCuidar los datos de salud de los clientes (Ley 1581).\n## Reglas\n- Nunca nombrar medicamentos, dosis, enfermedades ni condiciones de salud.\n- Nunca recomendar tratamientos.\n- Decir "lo de cada mes" o "lo suyo" en vez del producto.\n- Solo escribir a clientes que autorizaron contacto.\n- No escribir antes de las 8 a.m. ni después de las 8 p.m.\n## Ejemplos\n- Sí: "Ya le toca lo de cada mes."\n- No: "Ya le toca su pastilla de la presión."\n## Nunca\n- Mencionar la fórmula médica de otra persona.'
};

// ===================== UTILIDADES =====================

function hoja_(nombre) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!h) throw new Error('No existe la pestaña ' + nombre + '. Usa Retoma → Instalar.');
  return h;
}

function leerClientes_(hc) {
  const n = hc.getLastRow();
  if (n < 2) return [];
  return hc.getRange(2, 1, n - 1, TOTAL_COLS).getValues();
}

function cliente_(fila) {
  const o = {};
  COLS_BASE.forEach(c => { o[c] = fila[C[c]]; });
  return o;
}

/** Lo que se le manda a la IA: sin datos que no necesita. */
function fichaSegura_(cli) {
  return { nombre: primerNombre_(cli.Nombre), barrio: cli.Barrio, categoria: cli.Categoria, ciclo_dias: cli.CicloDias,
    compras_12m: cli.Compras12m, valor_12m: cli.Valor12m, canal_preferido: cli.CanalPreferido, domicilio: cli.Domicilio,
    tiene_whatsapp: cli.TieneWhatsApp, tiene_correo: cli.TieneCorreo, usa_descuentos: cli.UsaDescuentos, autoriza_contacto: cli.AutorizaContacto };
}

function config_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(HOJA_CONFIG);
  const m = {};
  if (h) h.getRange(1, 1, Math.max(h.getLastRow(), 1), 2).getValues().forEach(r => { m[String(r[0]).trim()] = r[1]; });
  const docs = {};
  DOCUMENTOS.forEach(d => { docs[d] = String(m[d] || '').trim(); });
  return {
    clave: String(m['Clave'] || '').trim(),
    modelo: String(m['Modelo'] || 'gemini-flash-latest').trim(),
    corte: fecha_(m['Fecha de corte']) || new Date(2026, 9, 31),
    cicloDefecto: num_(m['Ciclo por defecto']) || 45,
    docs: docs
  };
}

function cerebro_(cfg) {
  const out = {};
  DOCUMENTOS.forEach(d => {
    out[d] = CEREBRO_EJEMPLO[d];
    if (cfg.docs[d]) {
      try {
        const t = DocumentApp.openByUrl(cfg.docs[d]).getBody().getText().trim();
        if (t) out[d] = t;
      } catch (e) {
        registrar_('-', '-', '-', 'Orquestador', 'AVISO', 'No pude abrir el documento "' + d + '". Uso el de ejemplo. Revisa el enlace en Config.');
      }
    }
  });
  return out;
}

function actualizarEstado_() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  if (!h) return;
  const cfg = config_();
  const conectados = DOCUMENTOS.filter(d => cfg.docs[d]).length;
  const estado = cfg.clave
    ? 'Retoma activo · ' + conectados + ' de 5 documentos conectados' + (conectados < 5 ? ' (los demás usan el ejemplo)' : '')
    : 'Falta la clave de Gemini en B2';
  h.getRange('B11').setValue(estado).setFontWeight('bold').setFontColor(cfg.clave ? '#1F7A4D' : '#C23A18');
}

function registrar_(id, nombre, vuelta, agente, resultado, detalle) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_REGISTRO);
  if (!h) return;
  h.appendRow([new Date(), id, nombre, vuelta, agente, resultado, String(detalle || '').slice(0, 1500)]);
  const f = h.getLastRow();
  const r = String(resultado);
  if (r.indexOf('APROBADO') === 0) h.getRange(f, 6).setFontColor('#1F7A4D').setFontWeight('bold');
  else if (r.indexOf('RECHAZADO') === 0 || r.indexOf('ERROR') === 0 || r.indexOf('REVISIÓN') === 0) h.getRange(f, 6).setFontColor('#C23A18').setFontWeight('bold');
}

function hayPendientes_() {
  return leerClientes_(hoja_(HOJA_CLIENTES)).some(f => f[C['Ciclo IA']] === 'En espera');
}

function programarContinuacion_() {
  const ya = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'continuar');
  if (!ya) ScriptApp.newTrigger('continuar').timeBased().after(60 * 1000).create();
}

function fecha_(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}

function num_(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v || '').replace(/[$\s.]/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function norm_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function si_(v) { const s = norm_(v); return s === 'si' || s === 'true' || s === 'x'; }

function canonico_(v, lista) {
  const n = norm_(v);
  return lista.find(x => norm_(x) === n) || lista.find(x => n && norm_(x).indexOf(n) === 0) || String(v || '');
}

function primerNombre_(n) { return String(n || '').trim().split(/\s+/)[0]; }
