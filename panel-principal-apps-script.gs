/*  ══════════════════════════════════════════════════════════════
    PANEL PRINCIPAL — Apps Script de lectura + escritura
    ──────────────────────────────────────────────────────────────
    Guarda los trades del Panel Principal en una pestaña propia
    ("PanelPrincipal") de tu misma planilla, sin tocar tus hojas
    actuales. El dashboard le PIDE los datos (doGet) y le MANDA
    trades nuevos (doPost).

    v3 — el trade viaja codificado en Base64 (puro ASCII). Apps
    Script corrompe tildes/ñ con JSON crudo y también con
    application/x-www-form-urlencoded (bug de la plataforma, no
    del código). Base64 elimina el problema de raíz: no hay
    caracteres especiales en el cuerpo del POST, así que no hay
    nada que Apps Script pueda malinterpretar. Del lado del
    servidor se decodifica explícitamente como UTF-8.
    ══════════════════════════════════════════════════════════════ */

var HOJA = 'PanelPrincipal';
var COLS = ['Fecha', 'Instrumento', 'Direccion', 'Resultado', 'Modelo', 'Notas'];

function _hoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(COLS);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold');
  }
  return sh;
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
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if ((iFecha < 0 || row[iFecha] === '') && (iRes < 0 || row[iRes] === '')) continue;
      out.push({
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

// ── ESCRITURA: el dashboard manda un trade nuevo (Base64 → UTF-8 → JSON) ──
function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) || '';
    var bytes = Utilities.base64Decode(raw);
    var jsonStr = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    var data = JSON.parse(jsonStr);

    var sh = _hoja();
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(function (h) { return String(h).trim().toLowerCase(); });
    var row = head.map(function (key) {
      if (key === 'fecha')        return data.fecha || '';
      if (key === 'instrumento')  return data.instrumento || '';
      if (key === 'direccion' || key === 'dirección') return data.direccion || '';
      if (key === 'resultado')    return (data.resultado === '' || data.resultado == null) ? '' : Number(data.resultado);
      if (key === 'modelo')       return data.modelo || '';
      if (key === 'notas')        return data.notas || '';
      return '';
    });
    sh.appendRow(row);
    return _json({ ok: true });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}
