/*  ══════════════════════════════════════════════════════════════
    PANEL PRINCIPAL — Apps Script de lectura + escritura + borrado
    ──────────────────────────────────────────────────────────────
    Guarda los trades del Panel Principal en una pestaña propia
    ("PanelPrincipal") de tu misma planilla, sin tocar tus hojas
    actuales. El dashboard le PIDE los datos (doGet), le MANDA
    trades nuevos y le pide BORRAR trades existentes (doPost).

    v4 — agrega una columna "ID" (invisible para el uso normal,
    solo sirve para identificar cada fila sin ambigüedad al
    borrar). Si la pestaña ya existía de una versión anterior sin
    esa columna, se migra sola la primera vez que corre: agrega la
    columna ID y le asigna un identificador único a cada fila que
    todavía no tenga uno. No reordena ni toca ninguna otra columna.

    El trade sigue viajando codificado en Base64 (puro ASCII) por
    el mismo motivo que antes: Apps Script corrompe tildes/ñ tanto
    en JSON crudo como en form-urlencoded (bug de la plataforma).
    ══════════════════════════════════════════════════════════════ */

var HOJA = 'PanelPrincipal';
var COLS = ['Fecha', 'Instrumento', 'Direccion', 'Resultado', 'Modelo', 'Notas', 'ID'];

function _hoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(COLS);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold');
  }
  _ensureIdColumn(sh);
  return sh;
}

// Migración: agrega la columna ID si falta y backfillea IDs de filas viejas.
function _ensureIdColumn(sh) {
  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  if (lastRow < 1) { sh.appendRow(COLS); return; }

  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim().toLowerCase(); });
  var idCol = head.indexOf('id') + 1; // 1-based; 0 si no existe

  if (!idCol) {
    idCol = lastCol + 1;
    sh.getRange(1, idCol).setValue('ID').setFontWeight('bold');
  }

  if (lastRow > 1) {
    var idRange = sh.getRange(2, idCol, lastRow - 1, 1);
    var ids = idRange.getValues();
    var changed = false;
    for (var i = 0; i < ids.length; i++) {
      if (!ids[i][0]) { ids[i][0] = Utilities.getUuid(); changed = true; }
    }
    if (changed) idRange.setValues(ids);
  }
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _fechaTxt(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2) + '-' + ('0' + v.getDate()).slice(-2);
  }
  return String(v);
}

// ── LECTURA: el dashboard pide todos los trades ──
function doGet(e) {
  var sh = _hoja();
  var values = sh.getDataRange().getValues();
  var out = [];
  if (values.length > 1) {
    var head = values[0].map(function (h) { return String(h).trim().toLowerCase(); });
    function col(name) { return head.indexOf(name); }
    var iFecha = col('fecha');
    var iInst  = col('instrumento');
    var iDir   = col('direccion') >= 0 ? col('direccion') : col('dirección');
    var iRes   = col('resultado');
    var iMod   = col('modelo');
    var iNot   = col('notas');
    var iId    = col('id');
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if ((iFecha < 0 || row[iFecha] === '') && (iRes < 0 || row[iRes] === '')) continue;
      out.push({
        id:          iId    >= 0 ? String(row[iId]) : '',
        fecha:       iFecha >= 0 ? _fechaTxt(row[iFecha]) : '',
        instrumento: iInst  >= 0 ? String(row[iInst]) : '',
        direccion:   iDir   >= 0 ? String(row[iDir]) : '',
        resultado:   iRes   >= 0 ? row[iRes] : '',
        modelo:      iMod   >= 0 ? String(row[iMod]) : '',
        notas:       iNot   >= 0 ? String(row[iNot]) : ''
      });
    }
  }
  return _json(out);
}

// ── ESCRITURA / BORRADO: el dashboard manda una acción (Base64 → UTF-8 → JSON) ──
function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) || '';
    var bytes = Utilities.base64Decode(raw);
    var jsonStr = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    var data = JSON.parse(jsonStr);
    var action = data.action || 'add';

    var sh = _hoja();
    var lastCol = sh.getLastColumn();
    var head = sh.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(function (h) { return String(h).trim().toLowerCase(); });

    if (action === 'delete') {
      var idCol = head.indexOf('id') + 1;
      if (!idCol || !data.id) return _json({ ok: false, error: 'Falta columna ID o id a borrar.' });
      var lastRow = sh.getLastRow();
      if (lastRow < 2) return _json({ ok: false, error: 'No hay filas.' });
      var ids = sh.getRange(2, idCol, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(data.id)) {
          sh.deleteRow(i + 2); // +2: fila 1 es header, ids es 0-based desde la fila 2
          return _json({ ok: true, deleted: true });
        }
      }
      return _json({ ok: false, error: 'No se encontró el trade (ID no coincide).' });
    }

    // action === 'add' (default)
    var row = head.map(function (key) {
      if (key === 'fecha')        return data.fecha || '';
      if (key === 'instrumento')  return data.instrumento || '';
      if (key === 'direccion' || key === 'dirección') return data.direccion || '';
      if (key === 'resultado')    return (data.resultado === '' || data.resultado == null) ? '' : Number(data.resultado);
      if (key === 'modelo')       return data.modelo || '';
      if (key === 'notas')        return data.notas || '';
      if (key === 'id')           return Utilities.getUuid();
      return '';
    });
    sh.appendRow(row);
    return _json({ ok: true });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}
