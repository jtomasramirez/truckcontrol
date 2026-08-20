function doGet(e) {
  // Solo lectura: no se toma lock de escritura (getDataRange().getValues()
  // ya es una lectura consistente por sí sola), así varias peticiones en
  // paralelo no se bloquean entre sí esperando el lock.
  var sheet = getSheet_();
  var action = e.parameter.action;
  var data = sheet.getDataRange().getValues();
  if (action === 'get') {
    var key = e.parameter.key;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) return jsonOut_({key: key, value: data[i][1]});
    }
    return jsonOut_(null);
  } else if (action === 'list') {
    var prefix = e.parameter.prefix || '';
    var keys = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().indexOf(prefix) === 0) keys.push(data[i][0]);
    }
    return jsonOut_({keys: keys});
  } else if (action === 'listValues') {
    // Igual que 'list', pero devuelve también el value de cada fila en una
    // sola respuesta — evita tener que pedir key por key (usado para meses,
    // que pueden tener hasta 31 filas).
    var prefixV = e.parameter.prefix || '';
    var items = [];
    for (var j = 1; j < data.length; j++) {
      if (data[j][0] && data[j][0].toString().indexOf(prefixV) === 0) {
        items.push({key: data[j][0], value: data[j][1]});
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
    var data = sheet.getDataRange().getValues();

    if (action === 'set') {
      var value = body.value;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(value);
          return jsonOut_({key: key, value: value});
        }
      }
      sheet.appendRow([key, value]);
      return jsonOut_({key: key, value: value});

    } else if (action === 'delete') {
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.deleteRow(i + 1);
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
