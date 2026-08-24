function doGet(e) {
  // Solo lectura: no se toma lock de escritura. Además, solo se lee la
  // columna A (las keys) de entrada — es la parte liviana. La columna B
  // (values) puede tener celdas grandes (fotos en base64 guardadas como
  // 'salephoto:...'), así que solo se leen las filas que realmente hacen
  // falta en vez de traer TODA la hoja en cada petición, por chica que sea.
  var sheet = getSheet_();
  var action = e.parameter.action;
  var lastRow = sheet.getLastRow();
  var keyCol = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];

  if (action === 'get') {
    var key = e.parameter.key;
    for (var i = 0; i < keyCol.length; i++) {
      if (keyCol[i][0] === key) {
        var value = sheet.getRange(i + 2, 2).getValue();
        return jsonOut_({key: key, value: value});
      }
    }
    return jsonOut_(null);
  } else if (action === 'list') {
    var prefix = e.parameter.prefix || '';
    var keys = [];
    for (var j = 0; j < keyCol.length; j++) {
      var kj = keyCol[j][0];
      if (kj && kj.toString().indexOf(prefix) === 0) keys.push(kj);
    }
    return jsonOut_({keys: keys});
  } else if (action === 'listValues') {
    // Igual que 'list', pero devuelve también el value de cada fila que
    // matchea el prefijo — solo se leen esas filas puntuales, no toda la hoja.
    var prefixV = e.parameter.prefix || '';
    var items = [];
    for (var m = 0; m < keyCol.length; m++) {
      var km = keyCol[m][0];
      if (km && km.toString().indexOf(prefixV) === 0) {
        var val = sheet.getRange(m + 2, 2).getValue();
        items.push({key: km, value: val});
      }
    }
    return jsonOut_({items: items});
  }
  return jsonOut_({error: 'accion desconocida'});
}

function doPost(e) {
  var sheet = getSheet_();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var key = body.key;
    var lastRow = sheet.getLastRow();
    var keyCol = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];

    if (action === 'set') {
      var value = body.value;
      for (var i = 0; i < keyCol.length; i++) {
        if (keyCol[i][0] === key) {
          sheet.getRange(i + 2, 2).setValue(value);
          return jsonOut_({key: key, value: value});
        }
      }
      sheet.appendRow([key, value]);
      return jsonOut_({key: key, value: value});

    } else if (action === 'delete') {
      for (var j = 0; j < keyCol.length; j++) {
        if (keyCol[j][0] === key) {
          sheet.deleteRow(j + 2);
          return jsonOut_({key: key, deleted: true});
        }
      }
      return jsonOut_(null);
    }
    return jsonOut_({error: 'accion desconocida'});
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('KV');
  if (!sheet) {
    sheet = ss.insertSheet('KV');
    sheet.appendRow(['key', 'value']);
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
