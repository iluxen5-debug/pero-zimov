/**
 * PERO Заказы — серверная часть (Google Apps Script)
 */

// Лист с актуальным прайсом. При смене прайса достаточно поменять название здесь.
var PRICE_SHEET_NAME = 'Прайс с 05.10.2026';

// Если название модели в Справочниках отличается от названия в прайсе —
// добавьте пару сюда: 'как в справочнике': 'как в прайсе' (регистр не важен).
var PRICE_ALIASES = {
  'кеды сетка': 'кеды с сеткой',
  'сникерсы микс': 'сникеры микс'
};

var STATUSES = ['Новый', 'В работе', 'Готовы', 'Трек создан', 'Отправлены'];
var PAYMENT_STATUSES = ['Не оплачено', 'Оплачено'];

/**
 * Точка входа для веб-приложения Google Apps Script
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('PERO Заказы')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Выполняет функцию под общей блокировкой скрипта, чтобы два человека
 * не записали данные одновременно (например, в одну и ту же строку).
 */
function withLock(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Таблица занята другим пользователем, попробуйте ещё раз через несколько секунд');
  }
  try {
    return fn();
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
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
  var m = {
    id: colIndexByHeader(headers, ['№', 'заказ', 'номер']),
    date: colIndexByHeader(headers, ['дата']),
    status: colIndexByHeader(headers, ['статус']),
    model: colIndexByHeader(headers, ['модель']),
    size: colIndexByHeader(headers, ['размер']),
    kolodka: colIndexByHeader(headers, ['колодка']),
    verh: colIndexByHeader(headers, ['верх']),
    podklad: colIndexByHeader(headers, ['подклад']),
    podoshva: colIndexByHeader(headers, ['подошва']),
    price: colIndexByHeader(headers, ['цена']),
    payment: colIndexByHeader(headers, ['оплат']),
    note: colIndexByHeader(headers, ['примечание']),
    track: colIndexByHeader(headers, ['трек', 'трэк', 'штрихкод'])
  };
  if (m.id === -1) m.id = 0;
  return m;
}

function getHeaderMap(sheet) {
  var width = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  return { headers: headers, m: mapHeaders(headers), width: width };
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
    price: cellVal(row, m.price),
    payment: cellVal(row, m.payment) || 'Не оплачено',
    note: cellVal(row, m.note),
    track: cellVal(row, m.track)
  };
}

/**
 * Ищет строку заказа по номеру во всём листе (снизу вверх — берётся самая свежая).
 * Возвращает номер строки в таблице или -1.
 */
function findOrderRow(sheet, idIdx, orderId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var target = String(orderId).trim();
  var ids = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]).trim() === target) return i + 2;
  }
  return -1;
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

// ============================================================
// ПРАЙС
// ============================================================

function normName(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[^a-zа-я0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

// Грубое «основание» названия: без предлогов и окончаний, чтобы
// «Кеды сетка» и «Кеды с сеткой» считались одним и тем же.
function stemName(s) {
  return normName(s).split(' ').filter(function (w) {
    return w.length > 1;
  }).map(function (w) {
    return w.replace(/(ой|ей|ая|яя|ый|ий|ое|ее|ые|ие|ами|ями|ов|ев|а|я|о|е|ы|и|у|ю)$/, '');
  }).join(' ');
}

/**
 * Читает лист прайса: колонка A — наименование, колонка B — цена.
 * Возвращает [{name, price}] только для строк с числовой ценой.
 */
function readPriceList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRICE_SHEET_NAME);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  var data = sheet.getRange(1, 1, lastRow, 2).getValues();
  var list = [];
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    var raw = data[i][1];
    var price = typeof raw === 'number' ? raw : Number(String(raw).replace(/[^\d.,]/g, '').replace(',', '.'));
    if (name && raw !== '' && !isNaN(price) && price > 0) {
      list.push({ name: name, price: price });
    }
  }
  return list;
}

function findPrice(model, priceList) {
  var key = normName(model);
  if (!key) return '';
  if (PRICE_ALIASES[key]) key = normName(PRICE_ALIASES[key]);

  var i;
  for (i = 0; i < priceList.length; i++) {
    if (normName(priceList[i].name) === key) return priceList[i].price;
  }
  var stem = stemName(key);
  for (i = 0; i < priceList.length; i++) {
    if (stemName(priceList[i].name) === stem) return priceList[i].price;
  }
  return '';
}

/**
 * Карта «модель из справочника → цена из прайса» для формы.
 */
function getPriceMap(models) {
  var list = readPriceList();
  var map = {};
  for (var i = 0; i < models.length; i++) {
    var p = findPrice(models[i], list);
    if (p !== '') map[models[i]] = p;
  }
  return map;
}

/**
 * Цена для одной модели (используется, когда модель только что добавили в справочник).
 */
function getPriceForModel(model) {
  return findPrice(model, readPriceList());
}

// ============================================================
// СПРАВОЧНИКИ
// ============================================================

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
    } catch (e) {}
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

  // 5 минут: правки, сделанные прямо в листе «Справочники», появятся в приложении быстро
  cache.put('pero_dropdown_lists', JSON.stringify(lists), 300);
  return lists;
}

function addDropdownItem(listType, newValue) {
  try {
    return withLock(function () {
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
    });
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
// ЗАГРУЗКА ДАННЫХ
// ============================================================

function getInitialData() {
  var dropdowns = getDropdownLists();
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
      var r, idVal, trackVal, statusVal;

      for (r = Math.max(1, start - 1); r < data.length; r++) {
        idVal = data[r][m.id];
        if (idVal === '' || idVal === null) continue;
        statusVal = m.status !== -1 ? String(data[r][m.status]).trim() : '';
        trackVal = m.track !== -1 ? String(data[r][m.track] || '').trim() : '';
        allOrders.push(idVal);

        if (trackVal === '' && statusVal !== 'Отменен') {
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
      dropdowns: dropdowns,
      prices: getPriceMap(dropdowns.model || []),
      statuses: STATUSES,
      paymentStatuses: PAYMENT_STATUSES
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
      dropdowns: dropdowns,
      prices: {},
      statuses: STATUSES,
      paymentStatuses: PAYMENT_STATUSES
    };
  }
}

function getOrderById(orderId) {
  var sheet = getTargetSheet();
  var hm = getHeaderMap(sheet);
  var row = findOrderRow(sheet, hm.m.id, orderId);
  if (row === -1) throw new Error('Заказ №' + orderId + ' не найден');
  var data = sheet.getRange(row, 1, 1, hm.width).getValues()[0];
  return rowToOrder(data, hm.m);
}

function getOrdersByStatus(statusName) {
  var sheet = getTargetSheet();
  var stats = getLastOrderStats(sheet);
  var width = Math.min(sheet.getLastColumn(), 20);
  var lastRow = stats.lastRowIndex;
  var startRow = Math.max(2, lastRow - 350);
  var numRows = lastRow - startRow + 1;
  if (numRows < 1) return [];

  var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  var m = mapHeaders(headers);

  var data = sheet.getRange(startRow, 1, numRows, width).getValues();
  var orders = [];
  for (var i = 0; i < data.length; i++) {
    var idVal = data[i][m.id];
    if (idVal === '' || idVal === null) continue;
    var status = m.status !== -1 ? String(data[i][m.status]).trim() : '';
    if (status !== statusName) continue;
    orders.push(rowToOrder(data[i], m, status));
  }
  return orders;
}

function getNewOrders() {
  var orders = getOrdersByStatus('Новый');
  return { ok: true, count: orders.length, orders: orders };
}

function getTrackCreatedOrders() {
  var all = getOrdersByStatus('Трек создан');
  var orders = [];
  for (var i = 0; i < all.length; i++) {
    if (String(all[i].track || '').trim()) orders.push(all[i]);
  }
  return { ok: true, count: orders.length, orders: orders };
}

// ============================================================
// ЗАПИСЬ (всё под блокировкой и только в нужные ячейки)
// ============================================================

function updateOrderStatus(orderId, status) {
  return withLock(function () {
    var sheet = getTargetSheet();
    var hm = getHeaderMap(sheet);
    if (hm.m.status === -1) throw new Error('Не найден столбец Статус');
    var row = findOrderRow(sheet, hm.m.id, orderId);
    if (row === -1) throw new Error('Заказ №' + orderId + ' не найден');
    sheet.getRange(row, hm.m.status + 1).setValue(status);
    return 'Статус заказа №' + orderId + ' изменён на «' + status + '»';
  });
}

function updateOrderPayment(orderId, paymentStatus) {
  return withLock(function () {
    var sheet = getTargetSheet();
    var hm = getHeaderMap(sheet);
    if (hm.m.payment === -1) throw new Error('Не найден столбец Оплата');
    var row = findOrderRow(sheet, hm.m.id, orderId);
    if (row === -1) throw new Error('Заказ №' + orderId + ' не найден');
    sheet.getRange(row, hm.m.payment + 1).setValue(paymentStatus);
    return 'Оплата заказа №' + orderId + ' изменена на «' + paymentStatus + '»';
  });
}

/**
 * Создаёт заказ. Номер выдаётся на сервере под блокировкой, поэтому
 * при одновременной работе нескольких человек номера не повторяются.
 * Возвращает { ok, id, nextId, message }.
 */
function saveOrder(data) {
  try {
    return withLock(function () {
      var sheet = getTargetSheet();
      var stats = getLastOrderStats(sheet);
      var orderId = stats.nextOrderId;
      var targetRow = stats.lastRowIndex + 1;
      var lastCol = sheet.getLastColumn();
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var newRow = headers.map(function () { return ''; });

      var price = data.price;
      if (price === '' || price === null || price === undefined) {
        price = getPriceForModel(data.model);
      }
      price = (price === '' || isNaN(Number(price))) ? '' : Number(price);

      for (var idx = 0; idx < headers.length; idx++) {
        var h = String(headers[idx] || '').toLowerCase().replace(/\n/g, ' ').trim();
        if (h.indexOf('№') !== -1) newRow[idx] = orderId;
        else if (h.indexOf('дата') !== -1) newRow[idx] = data.date;
        else if (h.indexOf('статус') !== -1) newRow[idx] = 'Новый';
        else if (h.indexOf('модель') !== -1) newRow[idx] = data.model;
        else if (h.indexOf('размер') !== -1) newRow[idx] = data.size;
        else if (h.indexOf('колодка') !== -1) newRow[idx] = data.kolodka;
        else if (h.indexOf('верх') !== -1) newRow[idx] = data.verh;
        else if (h.indexOf('подклад') !== -1) newRow[idx] = data.podklad;
        else if (h.indexOf('подошва') !== -1) newRow[idx] = data.podoshva;
        else if (h.indexOf('цена') !== -1) newRow[idx] = price;
        else if (h.indexOf('оплат') !== -1) newRow[idx] = 'Не оплачено';
        else if (h.indexOf('примечание') !== -1) newRow[idx] = data.note;
      }

      var range = sheet.getRange(targetRow, 1, 1, lastCol);
      range.setValues([newRow]);
      range.setBackground('#9900FF');
      range.setFontColor('#FFFFFF');

      var msg = 'Заказ №' + orderId + ' создан (строка ' + targetRow + ')';
      if (price === '') msg += '. Цена не найдена в прайсе — заполните в таблице';
      return { ok: true, id: orderId, nextId: orderId + 1, message: msg };
    });
  } catch (e) {
    throw new Error('Не удалось сохранить: ' + e.message);
  }
}

function saveSingleTrack(orderId, trackNumber) {
  var cleanTrack = String(trackNumber || '').trim();
  if (!cleanTrack) throw new Error('Трек-номер не может быть пустым');

  return withLock(function () {
    var sheet = getTargetSheet();
    var hm = getHeaderMap(sheet);
    if (hm.m.track === -1) throw new Error('Не найден столбец "Трек-номер"');
    if (hm.m.status === -1) throw new Error('Не найден столбец "Статус"');

    var row = findOrderRow(sheet, hm.m.id, orderId);
    if (row === -1) throw new Error('Заказ №' + orderId + ' не найден');

    sheet.getRange(row, hm.m.track + 1).setValue(cleanTrack);
    sheet.getRange(row, hm.m.status + 1).setValue('Трек создан');

    return { ok: true, message: 'Трек для заказа №' + orderId + ' сохранён, статус изменён на «Трек создан»' };
  });
}

function saveMultipleTracks(updates) {
  try {
    return withLock(function () {
      var sheet = getTargetSheet();
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return { ok: true, message: 'Таблица пуста' };

      var hm = getHeaderMap(sheet);
      if (hm.m.track === -1) throw new Error('Не найден столбец "Трек-номер"');
      if (hm.m.status === -1) throw new Error('Не найден столбец "Статус"');

      var updateMap = {};
      (updates || []).forEach(function (item) {
        if (item && item.id != null) {
          var cleanTrack = String(item.track || '').trim();
          if (cleanTrack) updateMap[String(item.id).trim()] = cleanTrack;
        }
      });

      // Читаем только колонку номеров, пишем только в нужные ячейки
      var ids = sheet.getRange(2, hm.m.id + 1, lastRow - 1, 1).getValues();
      var done = {};
      var statusCells = [];
      var count = 0;

      for (var i = ids.length - 1; i >= 0; i--) {
        var rowId = String(ids[i][0]).trim();
        if (updateMap.hasOwnProperty(rowId) && !done[rowId]) {
          var row = i + 2;
          sheet.getRange(row, hm.m.track + 1).setValue(updateMap[rowId]);
          statusCells.push(sheet.getRange(row, hm.m.status + 1).getA1Notation());
          done[rowId] = true;
          count++;
        }
      }

      if (statusCells.length) {
        sheet.getRangeList(statusCells).setValue('Трек создан');
      }

      return { ok: true, message: 'Успешно сохранено треков и обновлен статус: ' + count };
    });
  } catch (e) {
    throw new Error('Не удалось сохранить треки: ' + e.message);
  }
}

function markAllTrackOrdersAsSent() {
  try {
    return withLock(function () {
      var sheet = getTargetSheet();
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return { ok: true, count: 0, message: 'Таблица пуста' };

      var hm = getHeaderMap(sheet);
      if (hm.m.status === -1) throw new Error('Не найден столбец "Статус"');

      var statuses = sheet.getRange(2, hm.m.status + 1, lastRow - 1, 1).getValues();
      var ids = sheet.getRange(2, hm.m.id + 1, lastRow - 1, 1).getValues();
      var cells = [];
      var updatedIds = [];

      for (var i = 0; i < statuses.length; i++) {
        if (String(statuses[i][0] || '').trim() === 'Трек создан') {
          cells.push(sheet.getRange(i + 2, hm.m.status + 1).getA1Notation());
          updatedIds.push(ids[i][0]);
        }
      }

      if (cells.length) {
        sheet.getRangeList(cells).setValue('Отправлены');
      }

      return {
        ok: true,
        count: updatedIds.length,
        ids: updatedIds,
        message: 'Заказы (' + updatedIds.join(', ') + ') переведены в статус «Отправлены»'
      };
    });
  } catch (e) {
    throw new Error('Ошибка при обновлении статусов: ' + e.message);
  }
}
