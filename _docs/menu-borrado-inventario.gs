/**
 * Botones de borrado con confirmación para la planilla Inventario 2026.
 * Solo borra los datos que se pegan o se digitan; las fórmulas se mantienen.
 *
 * IMPORTANTE: este script vive en el Apps Script VINCULADO a la planilla
 * "Inventario 2026" (Extensiones → Apps Script desde la hoja), NO en el
 * proyecto standalone de app_carga.gs (el que corre la app web / WhatsApp).
 * onOpen() solo funciona en un script vinculado a la hoja — por eso van
 * en archivos separados.
 *
 * Columnas que se borran por hoja (26/10, confirmado con Jose Miguel):
 * - CARGA: solo columna A (Código) desde la fila 2.
 * - INVENTARIO: A,B,C,I,O,P,Q desde la fila 3 (la hoja tiene 2 filas de
 *   encabezado — fila 1 título, fila 2 nombres de columna — los datos
 *   empiezan en la 3, igual que FILA_INICIO_INV_ en app_carga.gs).
 * - VENTAS: A a E desde la fila 2 (antes se borraba hasta la Z).
 * - traspaso: A a D desde la fila 2 (antes se borraba hasta la Z).
 *
 * Nota sobre la API/app en vivo: mapaInventario_() en app_carga.gs lee la
 * hoja INVENTARIO en vivo (sin CacheService) y solo usa las columnas A-G,
 * así que borrar I/O/P/Q no le afecta y cualquier cambio en A-G se ve de
 * inmediato en la app sin necesidad de invalidar caché. El único catálogo
 * cacheado es "Oferta Vigente" (20s) y CONFIG APP (45s), ninguno de los
 * dos vive en esta planilla.
 */

// Crea el menú "Inventario" cada vez que se abre la planilla
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🗂 Inventario')
    .addItem('Borrar CARGA MASIVA', 'borrarCarga')
    .addItem('Borrar TRASPASO', 'borrarTraspaso')
    .addItem('Borrar VENTAS', 'borrarVentas')
    .addSeparator()
    .addItem('Borrar INVENTARIO', 'borrarInventario')
    .addItem('Borrar TODO (nuevo inventario)', 'borrarTodo')
    .addToUi();
}

// ---------- Botones ----------

function borrarCarga() {
  if (!confirmar_('¿Borrar la columna A (Código) de la hoja CARGA, desde la fila 2?')) return;
  limpiar_('CARGA', ['A2:A']);
  avisar_('CARGA borrada.');
}

function borrarTraspaso() {
  if (!confirmar_('¿Borrar los traspasos (columnas A a D) de la hoja traspaso?')) return;
  limpiar_('traspaso', ['A2:D']);
  avisar_('Traspasos borrados.');
}

function borrarVentas() {
  if (!confirmar_('¿Borrar las ventas (columnas A a E) de la hoja VENTAS?')) return;
  limpiar_('VENTAS', ['A2:E']);
  avisar_('Ventas borradas.');
}

function borrarInventario() {
  if (!confirmarConPalabra_('Se borrará Producto, Descripción, Saldo, Físico, Precio, Línea y Familia de INVENTARIO (desde la fila 3).')) return;
  limpiar_('INVENTARIO ', ['A3:C', 'I3:I', 'O3:Q']);
  avisar_('Inventario borrado. Ya puedes pegar el reporte nuevo en A, B y C.');
}

function borrarTodo() {
  if (!confirmarConPalabra_('Se borrará CARGA, traspaso, VENTAS e INVENTARIO completos.')) return;
  limpiar_('CARGA', ['A2:A']);
  limpiar_('traspaso', ['A2:D']);
  limpiar_('VENTAS', ['A2:E']);
  limpiar_('INVENTARIO ', ['A3:C', 'I3:I', 'O3:Q']);
  avisar_('Planilla lista para un inventario nuevo.');
}

// ---------- Funciones de apoyo ----------

// Borra solo el contenido (mantiene formato y validaciones)
function limpiar_(nombreHoja, rangos) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (!hoja) throw new Error('No encuentro la hoja "' + nombreHoja + '"');
  var ultima = Math.max(hoja.getMaxRows(), 3);
  rangos.forEach(function (r) {
    var partes = r.split(':');                       // "A3:C" -> ["A3","C"]
    var hasta = partes[1].replace(/\d+/g, '') + ultima;
    hoja.getRange(partes[0] + ':' + hasta).clearContent();
  });
}

// Confirmación simple Sí / No
function confirmar_(mensaje) {
  var ui = SpreadsheetApp.getUi();
  return ui.alert('Confirmar', mensaje + '\n\nEsta acción no se puede deshacer.',
                  ui.ButtonSet.YES_NO) === ui.Button.YES;
}

// Confirmación reforzada: hay que escribir BORRAR
function confirmarConPalabra_(mensaje) {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Confirmar borrado',
                    mensaje + '\n\nEscribe BORRAR para continuar:',
                    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return false;
  if (r.getResponseText().trim().toUpperCase() !== 'BORRAR') {
    ui.alert('No se borró nada (no escribiste BORRAR).');
    return false;
  }
  return true;
}

function avisar_(texto) {
  SpreadsheetApp.getActiveSpreadsheet().toast(texto, 'Inventario', 5);
}
