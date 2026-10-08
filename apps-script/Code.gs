/**
 * Точка входа для веб-приложения Google Apps Script
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('PERO Заказы')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function getTargetSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = ['Заказы', 'Pero', 'Лист1'];
  for (var i = 0; i < names.length; i++) {
    var sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  return ss.getSheets()[0];
}

function colIndexByHeader(headers, keys) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').toLowerCase().replace(/\n/g, ' ').trim();
    for (var k = 0; k < keys.length; k++) {
      if (h.indexOf(keys[k]) !== -1) return i;
    }
  }
  return -1;
}

function mapHeaders(headers) {
  return {
    id: colIndexByHeader(headers, ['№', 'заказ', 'номер']),
    date: colIndexByHeader(headers, ['дата']),
    status: colIndexByHeader(headers, ['статус']),
    model: colIndexByHeader(headers, ['модель']),
    size: colIndexByHeader(headers, ['размер']),
    kolodka: colIndexByHeader(headers, ['колодка']),
    verh: colIndexByHeader(headers, ['верх']),
    podklad: colIndexByHeader(headers, ['подклад']),
    podoshva: colIndexByHeader(headers, ['подошва']),
    payment: colIndexByHeader(headers, ['оплат']),
    note: colIndexByHeader(headers, ['примечание']),
    track: colIndexByHeader(headers, ['трек', 'трэк', 'штрихкод'])
  };
}

function cellVal(row, idx) {
  if (idx === -1) return '';
  var v = row[idx];
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  }
  return String(v);
}

function rowToOrder(row, m, status) {
  return {
    id: cellVal(row, m.id),
    date: cellVal(row, m.date),
    status: status || cellVal(row, m.status),
    model: cellVal(row, m.model),
    size: cellVal(row, m.size),
    kolodka: cellVal(row, m.kolodka),
    verh: cellVal(row, m.verh),
    podklad: cellVal(row, m.podklad),
    podoshva: cellVal(row, m.podoshva),
    payment: cellVal(row, m.payment) || 'Не оплачено',
    note: cellVal(row, m.note),
    track: cellVal(row, m.track)
  };
}

function getLastOrderStats(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { lastRowIndex: 1, nextOrderId: 1600 };

  var start = Math.max(2, lastRow - 1500);
  var n = lastRow - start + 1;
  var colA = sheet.getRange(start, 1, n, 1).getValues();
  var lastRealRow = 1, lastOrderId = 0;

  for (var i = colA.length - 1; i >= 0; i--) {
    var val = colA[i][0];
    if (val !== '' && val !== null && !isNaN(val)) {
      lastRealRow = start + i;
      lastOrderId = Number(val);
      break;
    }
  }
  return {
    lastRowIndex: lastRealRow,
    nextOrderId: lastOrderId > 0 ? lastOrderId + 1 : 1600
  };
}

function getOrCreateDictSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Справочники');

  if (!sheet) {
    sheet = ss.insertSheet('Справочники');
    var titles = ['Модель', 'Размер', 'Колодка', 'Подошва', 'Верх', 'Подклад'];
    sheet.getRange(1, 1, 1, 6).setValues([titles]);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#eef4fc');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getDropdownLists() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('pero_dropdown_lists');
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {}
  }

  var sheet = getOrCreateDictSheet();
  var lastRow = sheet.getLastRow();
  var lists = { model: [], size: [], kolodka: [], podoshva: [], verh: [], podklad: [] };

  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) lists.model.push(String(data[i][0]).trim());
      if (data[i][1]) lists.size.push(String(data[i][1]).trim());
      if (data[i][2]) lists.kolodka.push(String(data[i][2]).trim());
      if (data[i][3]) lists.podoshva.push(String(data[i][3]).trim());
      if (data[i][4]) lists.verh.push(String(data[i][4]).trim());
      if (data[i][5]) lists.podklad.push(String(data[i][5]).trim());
    }
  }

  cache.put('pero_dropdown_lists', JSON.stringify(lists), 21600);
  return lists;
}

function addDropdownItem(listType, newValue) {
  try {
    var sheet = getOrCreateDictSheet();
    var colMap = { 'model': 1, 'size': 2, 'kolodka': 3, 'podoshva': 4, 'verh': 5, 'podklad': 6 };
    var col = colMap[listType];
    if (!col) throw new Error('Неизвестный тип списка');

    var val = String(newValue || '').trim();
    if (!val) throw new Error('Значение не может быть пустым');

    var lastRow = sheet.getLastRow();
    var targetRow = 2;

    if (lastRow > 1) {
      var colData = sheet.getRange(2, col, lastRow - 1, 1).getValues();
      for (var i = 0; i < colData.length; i++) {
        var cur = String(colData[i][0] || '').trim();
        if (cur.toLowerCase() === val.toLowerCase()) {
          return { ok: true, value: cur, existed: true };
        }
        if (cur !== '') {
          targetRow = i + 3;
        }
      }
    }

    sheet.getRange(targetRow, col).setValue(val);
    CacheService.getScriptCache().remove('pero_dropdown_lists');
    return { ok: true, value: val, existed: false };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function getInitialData() {
  try {
    var sheet = getTargetSheet();
    var stats = getLastOrderStats(sheet);
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    var withoutTrack = [];
    var allOrders = [];
    var trackCreatedOrders = [];
    var last = stats.lastRowIndex;
    var start = Math.max(2, last - 600);
    var width = Math.min(sheet.getLastColumn(), 20);

    if (last > 1) {
      var data = sheet.getRange(1, 1, last, width).getValues();
      var m = mapHeaders(data[0]);
      if (m.id === -1) m.id = 0;
      var r, idVal, trackVal, statusVal;

      for (r = Math.max(1, start - 1); r < data.length; r++) {
        idVal = data[r][m.id];
        if (idVal === '' || idVal === null) continue;
        statusVal = m.status !== -1 ? String(data[r][m.status]).trim() : '';
        trackVal = m.track !== -1 ? String(data[r][m.track] || '').trim() : '';
        allOrders.push(idVal);

        if ((trackVal === '' || trackVal === null) && statusVal !== 'Отменен') {
          withoutTrack.push(idVal);
        }

        // Заказы, у которых трек создан и готов к печати / отправке
        if (statusVal === 'Трек создан' && trackVal !== '') {
          trackCreatedOrders.push(rowToOrder(data[r], m, statusVal));
        }
      }
    }

    return {
      ok: true,
      nextId: stats.nextOrderId,
      date: today,
      ordersList: withoutTrack.reverse(),
      allOrders: allOrders.reverse(),
      trackOrders: trackCreatedOrders,
      dropdowns: getDropdownLists(),
      statuses: ['Новый', 'В работе', 'Готовы', 'Трек создан', 'Отправлены'],
      paymentStatuses: ['Не оплачено', 'Оплачено']
    };
  } catch (e) {
    return {
      ok: false,
      error: String(e.message || e),
      nextId: '',
      date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy'),
      ordersList: [],
      allOrders: [],
      trackOrders: [],
      dropdowns: getDropdownLists(),
      statuses: ['Новый', 'В работе', 'Готовы', 'Трек создан', 'Отправлены'],
      paymentStatuses: ['Не оплачено', 'Оплачено']
    };
  }
}

function getOrderById(orderId) {
  try {
    var sheet = getTargetSheet();
    var stats = getLastOrderStats(sheet);
    var width = Math.min(sheet.getLastColumn(), 20);
    var lastRow = stats.lastRowIndex;
    var startRow = Math.max(2, lastRow - 350);
    var numRows = lastRow - startRow + 1;

    var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
    var m = mapHeaders(headers);
    if (m.id === -1) m.id = 0;

    var data = sheet.getRange(startRow, 1, numRows, width).getValues();

    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][m.id]) === String(orderId)) {
        return rowToOrder(data[i], m);
      }
    }

    if (startRow > 2) {
      var fullData = sheet.getRange(2, 1, startRow - 2, width).getValues();
      for (var j = fullData.length - 1; j >= 0; j--) {
        if (String(fullData[j][m.id]) === String(orderId)) {
          return rowToOrder(fullData[j], m);
        }
      }
    }

    throw new Error('Заказ №' + orderId + ' не найден');
  } catch (e) {
    throw new Error(e.message);
  }
}

function getOrdersByStatus(statusName) {
  var sheet = getTargetSheet();
  var stats = getLastOrderStats(sheet);
  var width = Math.min(sheet.getLastColumn(), 20);
  var lastRow = stats.lastRowIndex;
  var startRow = Math.max(2, lastRow - 350);
  var numRows = lastRow - startRow + 1;

  var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  var m = mapHeaders(headers);
  if (m.id === -1) m.id = 0;

  var data = sheet.getRange(startRow, 1, numRows, width).getValues();
  var orders = [];
  for (var i = 0; i < data.length; i++) {
    var idVal = data[i][m.id];
    if (idVal === '' || idVal === null) continue;
    var status = m.status !== -1 ? String(data[i][m.status]) : '';
    if (status !== statusName) continue;
    orders.push(rowToOrder(data[i], m, status));
  }
  return orders;
}

function getNewOrders() {
  try {
    var orders = getOrdersByStatus('Новый');
    return { ok: true, count: orders.length, orders: orders };
  } catch (e) {
    throw new Error(e.message);
  }
}

function getTrackCreatedOrders() {
  try {
    var all = getOrdersByStatus('Трек создан');
    var orders = [];
    for (var i = 0; i < all.length; i++) {
      if (String(all[i].track || '').trim()) orders.push(all[i]);
    }
    return { ok: true, count: orders.length, orders: orders };
  } catch (e) {
    throw new Error(e.message);
  }
}

function updateOrderStatus(orderId, status) {
  try {
    var sheet = getTargetSheet();
    var stats = getLastOrderStats(sheet);
    var width = Math.min(sheet.getLastColumn(), 20);
    var lastRow = stats.lastRowIndex;
    var startRow = Math.max(2, lastRow - 350);
    var numRows = lastRow - startRow + 1;

    var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
    var m = mapHeaders(headers);
    if (m.id === -1) m.id = 0;
    if (m.status === -1) throw new Error('Не найден столбец Статус');

    var data = sheet.getRange(startRow, 1, numRows, width).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][m.id]) === String(orderId)) {
        var realRowNum = startRow + i;
        sheet.getRange(realRowNum, m.status + 1).setValue(status);
        SpreadsheetApp.flush();
        return 'Статус заказа №' + orderId + ' изменён на «' + status + '»';
      }
    }
    throw new Error('Заказ №' + orderId + ' не найден');
  } catch (e) {
    throw new Error(e.message);
  }
}

function updateOrderPayment(orderId, paymentStatus) {
  try {
    var sheet = getTargetSheet();
    var stats = getLastOrderStats(sheet);
    var width = Math.min(sheet.getLastColumn(), 20);
    var lastRow = stats.lastRowIndex;
    var startRow = Math.max(2, lastRow - 350);
    var numRows = lastRow - startRow + 1;

    var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
    var m = mapHeaders(headers);
    if (m.id === -1) m.id = 0;
    if (m.payment === -1) throw new Error('Не найден столбец Оплата');

    var data = sheet.getRange(startRow, 1, numRows, width).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][m.id]) === String(orderId)) {
        var realRowNum = startRow + i;
        sheet.getRange(realRowNum, m.payment + 1).setValue(paymentStatus);
        SpreadsheetApp.flush();
        return 'Оплата заказа №' + orderId + ' изменена на «' + paymentStatus + '»';
      }
    }
    throw new Error('Заказ №' + orderId + ' не найден');
  } catch (e) {
    throw new Error(e.message);
  }
}

function saveOrder(data) {
  try {
    var sheet = getTargetSheet();
    var stats = getLastOrderStats(sheet);
    var targetRow = stats.lastRowIndex + 1;
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var newRow = headers.map(function () { return ''; });

    for (var idx = 0; idx < headers.length; idx++) {
      var h = String(headers[idx] || '').toLowerCase().replace(/\n/g, ' ').trim();
      if (h.indexOf('№') !== -1) newRow[idx] = Number(data.orderId);
      else if (h.indexOf('дата') !== -1) newRow[idx] = data.date;
      else if (h.indexOf('статус') !== -1) newRow[idx] = 'Новый';
      else if (h.indexOf('модель') !== -1) newRow[idx] = data.model;
      else if (h.indexOf('размер') !== -1) newRow[idx] = data.size;
      else if (h.indexOf('колодка') !== -1) newRow[idx] = data.kolodka;
      else if (h.indexOf('верх') !== -1) newRow[idx] = data.verh;
      else if (h.indexOf('подклад') !== -1) newRow[idx] = data.podklad;
      else if (h.indexOf('подошва') !== -1) newRow[idx] = data.podoshva;
      else if (h.indexOf('оплат') !== -1) newRow[idx] = 'Не оплачено';
      else if (h.indexOf('примечание') !== -1) newRow[idx] = data.note;
    }

    var range = sheet.getRange(targetRow, 1, 1, lastCol);
    range.setValues([newRow]);
    range.setBackground('#9900FF');
    range.setFontColor('#FFFFFF');
    SpreadsheetApp.flush();

    return 'Заказ №' + data.orderId + ' создан (строка ' + targetRow + ')';
  } catch (e) {
    throw new Error('Не удалось сохранить: ' + e.message);
  }
}

function saveSingleTrack(orderId, trackNumber) {
  try {
    var cleanTrack = String(trackNumber || '').trim();
    if (!cleanTrack) throw new Error('Трек-номер не может быть пустым');

    var sheet = getTargetSheet();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) throw new Error('Таблица пуста');

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var idIdx = colIndexByHeader(headers, ['№', 'заказ', 'номер']);
    var trackIdx = colIndexByHeader(headers, ['трек', 'трэк', 'штрихкод']);
    var statusIdx = colIndexByHeader(headers, ['статус']);

    if (idIdx === -1) idIdx = 0;
    if (trackIdx === -1) throw new Error('Не найден столбец "Трек-номер"');
    if (statusIdx === -1) throw new Error('Не найден столбец "Статус"');

    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    var data = dataRange.getValues();
    var targetRowIdx = -1;

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === String(orderId).trim()) {
        targetRowIdx = i;
        break;
      }
    }

    if (targetRowIdx === -1) throw new Error('Заказ №' + orderId + ' не найден');

    var realRowNum = targetRowIdx + 2;
    sheet.getRange(realRowNum, trackIdx + 1).setValue(cleanTrack);
    sheet.getRange(realRowNum, statusIdx + 1).setValue('Трек создан');
    SpreadsheetApp.flush();

    return { ok: true, message: 'Трек для заказа №' + orderId + ' сохранён, статус изменён на «Трек создан»' };
  } catch (e) {
    throw new Error(e.message);
  }
}

function saveMultipleTracks(updates) {
  try {
    var sheet = getTargetSheet();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { ok: true, message: 'Таблица пуста' };

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var idIdx = colIndexByHeader(headers, ['№', 'заказ', 'номер']);
    var trackIdx = colIndexByHeader(headers, ['трек', 'трэк', 'штрихкод']);
    var statusIdx = colIndexByHeader(headers, ['статус']);

    if (idIdx === -1) idIdx = 0;
    if (trackIdx === -1) throw new Error('Не найден столбец "Трек-номер"');
    if (statusIdx === -1) throw new Error('Не найден столбец "Статус"');

    var updateMap = {};
    updates.forEach(function(item) {
      if (item && item.id != null) {
        var cleanTrack = String(item.track || '').trim();
        if (cleanTrack) updateMap[String(item.id).trim()] = cleanTrack;
      }
    });

    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    var data = dataRange.getValues();
    var count = 0;

    for (var i = 0; i < data.length; i++) {
      var rowId = String(data[i][idIdx]).trim();
      if (updateMap.hasOwnProperty(rowId)) {
        data[i][trackIdx] = updateMap[rowId];
        data[i][statusIdx] = 'Трек создан';
        count++;
      }
    }

    if (count > 0) {
      dataRange.setValues(data);
      SpreadsheetApp.flush();
    }

    return { ok: true, message: 'Успешно сохранено треков и обновлен статус: ' + count };
  } catch (e) {
    throw new Error('Не удалось сохранить треки: ' + e.message);
  }
}

function markAllTrackOrdersAsSent() {
  try {
    var sheet = getTargetSheet();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { ok: true, count: 0, message: 'Таблица пуста' };

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var idIdx = colIndexByHeader(headers, ['№', 'заказ', 'номер']);
    var statusIdx = colIndexByHeader(headers, ['статус']);

    if (statusIdx === -1) throw new Error('Не найден столбец "Статус"');
    if (idIdx === -1) idIdx = 0;

    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    var data = dataRange.getValues();
    var updatedIds = [];

    for (var i = 0; i < data.length; i++) {
      var currentStatus = String(data[i][statusIdx] || '').trim();
      if (currentStatus === 'Трек создан') {
        data[i][statusIdx] = 'Отправлены';
        updatedIds.push(data[i][idIdx]);
      }
    }

    if (updatedIds.length > 0) {
      dataRange.setValues(data);
      SpreadsheetApp.flush();
    }

    return {
      ok: true,
      count: updatedIds.length,
      ids: updatedIds,
      message: 'Заказы (' + updatedIds.join(', ') + ') переведены в статус «Отправлены»'
    };
  } catch (e) {
    throw new Error('Ошибка при обновлении статусов: ' + e.message);
  }
}
