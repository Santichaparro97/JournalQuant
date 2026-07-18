/*  ══════════════════════════════════════════════════════════════
    JOURNAL QUANT — Apps Script unico para toda la web
    ──────────────────────────────────────────────────────────────
    Este script esta "atado" a tu planilla, asi que puede leer y
    escribir CUALQUIER pestana de ese mismo archivo. Le da servicio
    a dos partes de la web:

      1) PANEL PRINCIPAL  -> pestana propia "PanelPrincipal"
                             (registro en %, calendario)
      2) REGISTRO          -> tu hoja principal de trades
                             (la planilla de ~65 columnas)

    v5 — agrega las acciones del REGISTRO (listar pestanas, leer
    una pestana completa con sus formulas, editar una celda y
    agregar una fila) SIN tocar nada de lo que ya funcionaba del
    Panel Principal: si no se manda ningun "action", se comporta
    exactamente igual que la v4.

    PROTECCIONES DEL REGISTRO:
      · Respaldo automatico: antes de la primera escritura del dia
        sobre una hoja, se duplica en una pestana oculta llamada
        "_bk_<hoja>_<fecha>". Una sola por dia, no se acumula.
      · Las celdas con formula NO se pisan: si intentas editar una,
        el script la rechaza y avisa.
      · Al agregar una fila, las formulas de la fila anterior se
        copian hacia abajo, asi las columnas calculadas siguen
        calculando solas.

    COMO ACTUALIZARLO:
      1) Tu Sheet -> Extensiones -> Apps Script
      2) Ctrl+A, borrar, pegar este archivo completo
      3) Guardar (Ctrl+S)
      4) Implementar -> Administrar implementaciones -> lapiz
         -> Version: "Nueva version" -> Implementar
      La URL /exec NO cambia.
    ══════════════════════════════════════════════════════════════ */

var HOJA = 'PanelPrincipal';
var COLS = ['Fecha', 'Instrumento', 'Direccion', 'Resultado', 'Modelo', 'Notas', 'ID'];

/* ─────────── helpers generales ─────────── */

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

// Hoja por nombre, SIN tocarle nada (para el Registro).
function _sheetByName(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!name) return ss.getSheets()[0];
  return ss.getSheetByName(name);
}

/* ─────────── PANEL PRINCIPAL (igual que v4) ─────────── */

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

function _ensureIdColumn(sh) {
  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  if (lastRow < 1) { sh.appendRow(COLS); return; }

  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim().toLowerCase(); });
  var idCol = head.indexOf('id') + 1;

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

function _panelRead() {
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
  return out;
}

/* ─────────── REGISTRO (nuevo en v5) ─────────── */

// Respaldo automatico: una copia oculta por hoja y por dia.
function _backupOnce(sh) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tag = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || 'GMT', 'yyyy-MM-dd');
    var bkName = '_bk_' + sh.getName() + '_' + tag;
    if (bkName.length > 95) bkName = bkName.substring(0, 95);
    if (ss.getSheetByName(bkName)) return;      // ya hay respaldo de hoy
    var copy = sh.copyTo(ss);
    copy.setName(bkName);
    copy.hideSheet();
  } catch (err) {
    // El respaldo nunca debe impedir guardar: si falla, se sigue igual.
  }
}

function _listSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .filter(function (sh) {
      var n = sh.getName();
      return n.indexOf('_bk_') !== 0;          // los respaldos no se listan
    })
    .map(function (sh) {
      return { name: sh.getName(), rows: Math.max(0, sh.getLastRow() - 1), cols: sh.getLastColumn() };
    });
}

function _gridRead(name) {
  var sh = _sheetByName(name);
  if (!sh) return { ok: false, error: 'No existe la hoja: ' + name };

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { ok: true, sheet: sh.getName(), headers: [], rows: [], rowIndices: [], formulaCols: [] };
  }

  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h); });

  var n = lastRow - 1;
  var rows = [], rowIndices = [], formulaCols = [];
  for (var c = 0; c < lastCol; c++) formulaCols.push(false);

  if (n > 0) {
    var rng  = sh.getRange(2, 1, n, lastCol);
    var vals = rng.getValues();
    var fmls = rng.getFormulas();

    for (var c2 = 0; c2 < lastCol; c2++) {
      for (var r2 = 0; r2 < n; r2++) {
        if (fmls[r2][c2]) { formulaCols[c2] = true; break; }
      }
    }

    for (var r = 0; r < n; r++) {
      var out = [];
      var empty = true;
      for (var c3 = 0; c3 < lastCol; c3++) {
        var v = vals[r][c3];
        if (Object.prototype.toString.call(v) === '[object Date]') { out.push(_fechaTxt(v)); empty = false; }
        else if (typeof v === 'number' || typeof v === 'boolean') { out.push(v); empty = false; }
        else { var s = String(v); if (s !== '') empty = false; out.push(s); }
      }
      if (empty) continue;                      // saltea filas totalmente vacias
      rows.push(out);
      rowIndices.push(r + 2);                   // fila real en la planilla
    }
  }

  return {
    ok: true, sheet: sh.getName(),
    headers: headers, rows: rows, rowIndices: rowIndices, formulaCols: formulaCols
  };
}

function _coerce(value, type) {
  if (value === null || value === undefined || value === '') return '';
  if (type === 'num') {
    var n = Number(String(value).replace(',', '.'));
    return isNaN(n) ? value : n;
  }
  if (type === 'bool') {
    var s = String(value).trim().toLowerCase();
    if (s === 'true' || s === 'si' || s === 'sí' || s === 'x' || s === '1') return true;
    if (s === 'false' || s === 'no' || s === '0') return false;
    return value;
  }
  if (type === 'date') {
    var m = String(value).match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return value;
  }
  return value;
}

function _cellUpdate(data) {
  var sh = _sheetByName(data.sheet);
  if (!sh) return { ok: false, error: 'Hoja no encontrada: ' + data.sheet };

  var r = Number(data.row), c = Number(data.col);
  if (!(r > 1) || !(c > 0)) return { ok: false, error: 'Fila o columna invalida.' };

  var cell = sh.getRange(r, c);
  if (cell.getFormula()) {
    return { ok: false, error: 'Esa celda tiene una formula: no se edita desde la web para no romperla.' };
  }

  _backupOnce(sh);
  cell.setValue(_coerce(data.value, data.type));
  return { ok: true };
}

function _rowAdd(data) {
  var sh = _sheetByName(data.sheet);
  if (!sh) return { ok: false, error: 'Hoja no encontrada: ' + data.sheet };

  _backupOnce(sh);

  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  var vals  = data.values || {};
  var types = data.types  || {};

  var row = headers.map(function (h) {
    var v = vals[h];
    if (v === undefined || v === null) return '';
    return _coerce(v, types[h]);
  });

  sh.appendRow(row);
  var newRow = sh.getLastRow();

  // Propaga las formulas de la fila anterior hacia la nueva.
  if (newRow > 2) {
    var prev = sh.getRange(newRow - 1, 1, 1, lastCol);
    var fmls = prev.getFormulas()[0];
    for (var c = 0; c < lastCol; c++) {
      if (fmls[c]) prev.getCell(1, c + 1).copyTo(sh.getRange(newRow, c + 1));
    }
  }

  return { ok: true, row: newRow };
}

/* ─────────── entrypoints ─────────── */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'sheets') return _json({ ok: true, sheets: _listSheets() });
  if (action === 'grid')   return _json(_gridRead(e.parameter.sheet));
  return _json(_panelRead());                   // compatibilidad Panel Principal
}

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) || '';
    var bytes = Utilities.base64Decode(raw);
    var jsonStr = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    var data = JSON.parse(jsonStr);
    var action = String(data.action || 'add').toLowerCase();

    // ── Registro ──
    if (action === 'cellupdate') return _json(_cellUpdate(data));
    if (action === 'rowadd')     return _json(_rowAdd(data));

    // ── Panel Principal (igual que v4) ──
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
          sh.deleteRow(i + 2);
          return _json({ ok: true, deleted: true });
        }
      }
      return _json({ ok: false, error: 'No se encontro el trade (ID no coincide).' });
    }

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
