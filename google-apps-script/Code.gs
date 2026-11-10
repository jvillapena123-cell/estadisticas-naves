/**
 * HTP · Estadística Operacional de Naves — API sobre Google Sheets
 * Pegar en: Google Sheets → Extensiones → Apps Script → Code.gs
 * Implementar → Nueva implementación → Aplicación web
 *   Ejecutar como: Yo   ·   Quién tiene acceso: Cualquier persona
 * Copiar la URL que termina en /exec en firebase-config.js → SHEETS_API_URL
 */

const HOJA = 'Naves';                 // nombre de la pestaña con los datos
const CLAVE_EDICION = 'HTP2026';      // ← CAMBIA esta clave: se pide en la página para ingresar/editar
const TZ = 'America/Santiago';

function hoja_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA);
  if (!sh) throw new Error('No existe la pestaña "' + HOJA + '"');
  return sh;
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function encabezados_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}

/** GET → lectura; con ?action=… escritura. Con ?callback=fn responde JSONP (evita bloqueo CORS) */
function doGet(e) {
  const p = (e && e.parameter) || {};
  let res;
  if (p.action) {
    try { res = escribir_(Object.assign(JSON.parse(p.payload || '{}'), { action: p.action })); }
    catch (err) { res = { ok: false, error: String(err) }; }
  } else {
    res = leer_();
  }
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + JSON.stringify(res) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return out_(res);
}

function doPost(e) {
  try { return out_(escribir_(JSON.parse(e.postData.contents))); }
  catch (err) { return out_({ ok: false, error: String(err) }); }
}

const esFecha_ = v => Object.prototype.toString.call(v) === '[object Date]';

function leer_() {
  try {
    const sh = hoja_();
    const data = sh.getDataRange().getValues();
    const head = data.shift().map(String);
    const rows = data
      .filter(r => r.some(v => v !== '' && v !== null))
      .map(r => {
        const o = {};
        head.forEach((h, i) => {
          if (!h) return;
          let v = r[i];
          if (esFecha_(v)) v = Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm");
          o[h] = v === '' ? null : v;
        });
        return o;
      });
    return { ok: true, rows: rows, actualizado: new Date().toISOString() };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function escribir_(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (req.key !== CLAVE_EDICION) return { ok: false, error: 'Clave de edición incorrecta' };
    const sh = hoja_();
    let head = encabezados_(sh);
    const idCol = head.indexOf('ID');
    if (idCol < 0) throw new Error('Falta la columna "ID" en la fila 1');

    const ids = sh.getLastRow() > 1
      ? sh.getRange(2, idCol + 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]))
      : [];

    const toRow = obj => {
      Object.keys(obj).forEach(k => {
        if (head.indexOf(k) < 0) { sh.getRange(1, head.length + 1).setValue(k); head.push(k); }
      });
      return head.map(h => {
        let v = obj[h];
        if (v === null || v === undefined) return '';
        if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(v)) {
          return Utilities.parseDate(v, TZ, "yyyy-MM-dd'T'HH:mm");
        }
        return v;
      });
    };
    const upsert = obj => {
      const values = toRow(obj);
      const i = ids.indexOf(String(obj.ID));
      if (i >= 0) sh.getRange(i + 2, 1, 1, values.length).setValues([values]);
      else { sh.appendRow(values); ids.push(String(obj.ID)); }
    };

    if (req.action === 'save') upsert(req.row);
    else if (req.action === 'bulk') req.rows.forEach(upsert);
    else if (req.action === 'delete') {
      const i = ids.indexOf(String(req.id));
      if (i >= 0) sh.deleteRow(i + 2);
    } else throw new Error('Acción desconocida');

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
