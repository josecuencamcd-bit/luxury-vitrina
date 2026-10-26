/**
 * Sincronización automática: Inventario Subase → Supabase.
 * Vive en el Apps Script VINCULADO a la planilla "Inventario Subase"
 * (va en el MISMO proyecto que menu-borrado-inventario.gs, archivo aparte).
 *
 * Qué hace: cada 3 minutos (una vez instalado el disparador, ver
 * instalarDisparadorSync() al final) lee las 7 pestañas acordadas y las
 * manda a la función planilla-sync de Supabase:
 *   - STOCK TIENDA   → productos            (reemplaza por código)
 *   - OFERTA FLASH   → oferta_flash         (reemplaza por código)
 *   - Oferta Vigente → oferta_vigente       (reemplaza por código)
 *   - INVENTARIO     → inventario_snapshot  (reemplaza por código, solo
 *                       las columnas A,B,C,I,O,P,Q — las mismas que se
 *                       borran a mano)
 *   - CONFIG APP     → config_app           (reemplaza, fila única)
 *   - VENTAS         → ventas               (solo agrega filas nuevas)
 *   - traspaso       → traspasos            (solo agrega filas nuevas)
 *
 * VENTAS y traspaso son hojas que se van llenando y después se borran
 * (con el menú de borrado), así que no tienen un código único por fila
 * para "reemplazar" — en vez de eso, se guarda en Propiedades del Script
 * cuál fue la última fila ya mandada, y en cada corrida solo se manda lo
 * nuevo. Si la hoja tiene MENOS filas que la última vez (= alguien la
 * borró), el bookmark se reinicia solo.
 *
 * CONFIGURACIÓN NECESARIA (una sola vez):
 * 1. Extensiones → Apps Script → ⚙️ Configuración del proyecto →
 *    "Propiedades de secuencia de comandos" → agregar:
 *      SYNC_SECRET = (el mismo valor que pusiste en Supabase, pídemelo si
 *      no lo tienes a mano)
 * 2. Ejecutar UNA VEZ la función instalarDisparadorSync() (desde el
 *    editor: seleccionarla en el desplegable de funciones → ▶ Ejecutar).
 *    Eso deja el disparador de cada 3 minutos andando solo.
 */

var SUPABASE_URL_ = 'https://zlpwgkoajqmaprikdkan.supabase.co';
var SUPABASE_ANON_KEY_ = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InpscHdna29hanFtYXByaWtka2FuIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTI4NzkyNDQsImV4cCI6MjEwODQ1NTI0NH0.jwRcHCpRLAyWEE2iRbCQpooGU_zunqXtZoEu5CWaetY';

// ---------- Punto de entrada (lo llama el disparador cada 3 min) ----------

function sincronizarTodo() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var resultados = [];

  resultados.push(sincronizarCatalogo_(libro, 'STOCK TIENDA', 'productos', 3, 12));
  resultados.push(sincronizarCatalogo_(libro, 'OFERTA FLASH', 'oferta_flash', 3, 16));
  resultados.push(sincronizarCatalogo_(libro, 'Oferta Vigente', 'oferta_vigente', 2, 7));
  resultados.push(sincronizarCatalogo_(libro, 'INVENTARIO ', 'inventario_snapshot', 3, 18));
  resultados.push(sincronizarConfig_(libro));
  resultados.push(sincronizarIncremental_(libro, 'VENTAS', 'ventas', 2, 5));
  resultados.push(sincronizarIncremental_(libro, 'traspaso', 'traspasos', 2, 3));

  var fallos = resultados.filter(function (r) { return r && !r.ok; });
  if (fallos.length) {
    Logger.log('Sync con errores: ' + JSON.stringify(fallos));
  }
}

// ---------- Catálogos: se reemplaza por código en cada corrida ----------

function sincronizarCatalogo_(libro, nombreHoja, tabla, filaInicio, columnas) {
  try {
    var hoja = libro.getSheetByName(nombreHoja);
    if (!hoja) return { ok: false, tabla: tabla, error: 'No existe la hoja ' + nombreHoja };
    var ultima = hoja.getLastRow();
    var filas = ultima - filaInicio + 1;
    var datos = filas > 0 ? hoja.getRange(filaInicio, 1, filas, columnas).getValues() : [];
    return llamarSync_(tabla, { tabla: tabla, filas: datos });
  } catch (e) {
    return { ok: false, tabla: tabla, error: String(e) };
  }
}

// ---------- CONFIG APP: fila única, columna B ----------

function sincronizarConfig_(libro) {
  try {
    var hoja = libro.getSheetByName('CONFIG APP');
    if (!hoja) return { ok: false, tabla: 'config_app', error: 'No existe CONFIG APP' };
    var v = hoja.getRange('B2:B9').getValues();
    var valores = {
      dolar: v[0][0], multiplicador: v[1][0], descuentos: v[2][0], whatsapp: v[3][0],
      mensaje_tipo: v[4][0], redondeo: v[5][0], ars_por_usd: v[6][0], ofertas_activas: v[7][0]
    };
    return llamarSync_('config_app', { tabla: 'config_app', valores: valores });
  } catch (e) {
    return { ok: false, tabla: 'config_app', error: String(e) };
  }
}

// ---------- VENTAS / traspaso: solo se agregan filas nuevas ----------

function sincronizarIncremental_(libro, nombreHoja, tabla, filaInicio, columnas) {
  try {
    var hoja = libro.getSheetByName(nombreHoja);
    if (!hoja) return { ok: false, tabla: tabla, error: 'No existe la hoja ' + nombreHoja };
    var props = PropertiesService.getScriptProperties();
    var clave = 'bm_' + tabla;
    var marcador = Number(props.getProperty(clave) || 0);
    var ultima = hoja.getLastRow();

    // Si hay menos filas que el marcador, asumimos que se borró la hoja.
    if (ultima < filaInicio + marcador) marcador = 0;

    var desde = filaInicio + marcador;
    var filasNuevas = ultima - desde + 1;
    if (filasNuevas <= 0) return { ok: true, tabla: tabla, filas: 0 };

    var datos = hoja.getRange(desde, 1, filasNuevas, columnas).getValues();
    var r = llamarSync_(tabla, { tabla: tabla, filas: datos });
    if (r.ok) props.setProperty(clave, String(marcador + filasNuevas));
    return r;
  } catch (e) {
    return { ok: false, tabla: tabla, error: String(e) };
  }
}

// ---------- Llamada HTTP a la función planilla-sync ----------

function llamarSync_(tabla, payload) {
  var secreto = PropertiesService.getScriptProperties().getProperty('SYNC_SECRET');
  if (!secreto) return { ok: false, tabla: tabla, error: 'Falta SYNC_SECRET en Propiedades del Script' };
  var resp = UrlFetchApp.fetch(SUPABASE_URL_ + '/functions/v1/planilla-sync', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'apikey': SUPABASE_ANON_KEY_,
      'Authorization': 'Bearer ' + SUPABASE_ANON_KEY_,
      'x-sync-secret': secreto
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var codigo = resp.getResponseCode();
  if (codigo >= 200 && codigo < 300) return { ok: true, tabla: tabla };
  return { ok: false, tabla: tabla, error: 'HTTP ' + codigo + ': ' + resp.getContentText() };
}

// ---------- Instalación del disparador (ejecutar UNA SOLA VEZ a mano) ----------

function instalarDisparadorSync() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sincronizarTodo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sincronizarTodo').timeBased().everyMinutes(3).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('Sincronización automática instalada (cada 3 min).', 'Supabase', 5);
}
