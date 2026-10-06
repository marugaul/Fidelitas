/**
 * SC-250 · Backend de persistencia (Google Apps Script)
 *
 * PASOS (5 minutos):
 *  1. Crear una Google Sheet nueva (sheets.new) y ponerle nombre, ej. "SC250 Resultados".
 *  2. Menú Extensiones → Apps Script. Borrar el contenido y pegar TODO este archivo.
 *  3. Guardar (ícono de disco).
 *  4. Botón azul "Implementar" → "Nueva implementación".
 *     - Tipo: Aplicación web
 *     - Ejecutar como: Yo
 *     - Quién tiene acceso: Cualquier persona
 *     → Implementar. Autorizar permisos cuando lo pida.
 *  5. Copiar la "URL de la aplicación web" (termina en /exec)
 *     y pegarla en la constante API_URL de quiz.html y evaluacion.html.
 *
 * Cada envío queda como una fila en la hoja "quiz", "eval" o "poker".
 *
 * NOTA: si ya tenías una implementación, después de pegar este código hay que
 * hacer Implementar → Administrar implementaciones → editar → Versión: Nueva versión.
 */

const SHEET_ID = '1ROIASNdSa0pW75309gmxOeZ5coto0yeWLWAQM4fW4yY';

const SHEETS = {
  quiz: ['ts', 'fecha', 'nombre', 'correctas', 'total', 'puntaje', 'respuestas', 'json'],
  poker: ['ts', 'fecha', 'tipo', 'tarea', 'nombre', 'valor', 'texto', 'json'],
  eval: ['ts', 'fecha', 'evaluador', 'grupo', 'contenido', 'estructura', 'dominio', 'contexto', 'reflexion', 'total', 'json'],
};

function getSheet(app) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(app);
  if (!sh) {
    sh = ss.insertSheet(app);
    sh.appendRow(SHEETS[app]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const app = body.app;
    const r = body.record;
    if (!SHEETS[app] || !r) return json({ ok: false, error: 'bad request' });
    const fecha = new Date(r.ts || Date.now());
    const row = app === 'quiz'
      ? [r.ts, fecha, r.name, r.correct, r.total, r.score, JSON.stringify(r.answers), JSON.stringify(r)]
      : app === 'poker'
      ? [r.ts, fecha, r.type, r.task || '', r.name, r.value != null ? r.value : (r.final != null ? r.final : ''),
         r.title || '', JSON.stringify(r)]
      : [r.ts, fecha, r.rater, r.group, r.ratings.contenido, r.ratings.estructura, r.ratings.dominio,
         r.ratings.contexto, r.ratings.reflexion, r.total, JSON.stringify(r)];
    getSheet(app).appendRow(row);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const app = (e.parameter && e.parameter.app) || '';
  if (!SHEETS[app]) return json({ ok: false, error: 'app must be quiz, eval or poker' });
  const sh = getSheet(app);
  const last = sh.getLastRow();
  if (last < 2) return json([]);
  const jsonCol = SHEETS[app].length;
  const values = sh.getRange(2, jsonCol, last - 1, 1).getValues();
  const out = [];
  values.forEach(v => { try { if (v[0]) out.push(JSON.parse(v[0])); } catch (_) {} });
  return json(out);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
