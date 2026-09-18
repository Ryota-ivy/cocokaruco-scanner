const INVENTORY_FILTER_CONFIG = {
  sheetName: '商品在庫一覧',
  filterRow: 1,
  headerRow: 2,
  firstDataRow: 3,
  keywordColumn: 2,
  filterHeaders: [
    '種類',
    'ブランド名',
    'カラー',
    'サイズ',
    'メーカー',
    '在庫数',
    '在庫有無',
    'Instagram掲載',
    'EC掲載'
  ]
};

function onOpen() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  spreadsheet.getUi()
    .createMenu('商品検索・フィルター')
    .addItem('検索欄を設定・更新', 'setupInventoryFilters')
    .addItem('検索をクリア', 'clearInventoryFilters')
    .addToUi();

  const sheet = spreadsheet.getSheetByName(INVENTORY_FILTER_CONFIG.sheetName);
  if (sheet) {
    setupInventoryFilters_(sheet);
  }
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const config = INVENTORY_FILTER_CONFIG;

  if (sheet.getName() !== config.sheetName) return;
  if (e.range.getRow() !== config.filterRow) return;

  const headers = getInventoryHeaders_(sheet);
  const editedHeader = headers[e.range.getColumn() - 1];
  const isKeyword = e.range.getColumn() === config.keywordColumn;
  const isKeptColumnFilter = e.range.getColumn() <= 8 &&
    config.filterHeaders.includes(editedHeader);

  if (!isKeyword && !isKeptColumnFilter) return;

  applyInventoryFilters_(sheet, headers);
}

function setupInventoryFilters() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(INVENTORY_FILTER_CONFIG.sheetName);

  if (!sheet) {
    throw new Error('「商品在庫一覧」シートが見つかりません。');
  }

  setupInventoryFilters_(sheet);
  applyInventoryFilters_(sheet, getInventoryHeaders_(sheet));
}

function setupInventoryFilters_(sheet) {
  const config = INVENTORY_FILTER_CONFIG;
  const headers = getInventoryHeaders_(sheet);
  const lastRow = sheet.getLastRow();

  sheet.getRange(config.filterRow, config.keywordColumn)
    .setNote('商品番号・バーコード・商品名などを横断検索します。空欄にすると解除されます。')
    .setBackground('#fff2cc');

  // B1〜H1の検索・フィルターは残し、I1以降の追加フィルターだけ削除する。
  config.filterHeaders.forEach(function (headerName) {
    const columnIndex = headers.indexOf(headerName) + 1;
    if (!columnIndex || columnIndex === config.keywordColumn) return;

    const filterCell = sheet.getRange(config.filterRow, columnIndex);

    if (columnIndex <= 8) {
      const choices = getFilterChoices_(sheet, headerName, columnIndex, lastRow);
      filterCell
        .setDataValidation(
          SpreadsheetApp.newDataValidation()
            .requireValueInList(['すべて'].concat(choices), true)
            .setAllowInvalid(false)
            .build()
        )
        .setNote(headerName + 'で絞り込みます。')
        .setBackground('#d9ead3');
    } else {
      filterCell
        .clearContent()
        .clearDataValidations()
        .setNote(null)
        .setBackground(null);
    }
  });

  sheet.setFrozenRows(config.headerRow);
}

function getFilterChoices_(sheet, headerName, columnIndex, lastRow) {
  const config = INVENTORY_FILTER_CONFIG;

  if (headerName === '在庫数' || headerName === '在庫有無') {
    return ['在庫あり', '在庫なし'];
  }

  if (lastRow < config.firstDataRow) return [];

  const values = sheet
    .getRange(config.firstDataRow, columnIndex, lastRow - config.firstDataRow + 1, 1)
    .getDisplayValues()
    .map(function (row) { return String(row[0]).trim(); })
    .filter(function (value) { return value !== '' && value !== 'すべて'; });

  return Array.from(new Set(values)).sort(function (left, right) {
    return left.localeCompare(right, 'ja');
  });
}

function applyInventoryFilters_(sheet, headers) {
  const config = INVENTORY_FILTER_CONFIG;
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), headers.length);
  const maxRows = sheet.getMaxRows();

  if (maxRows >= config.firstDataRow) {
    sheet.showRows(config.firstDataRow, maxRows - config.firstDataRow + 1);
  }

  if (lastRow < config.firstDataRow || lastColumn === 0) return;

  const filterValues = sheet
    .getRange(config.filterRow, 1, 1, lastColumn)
    .getDisplayValues()[0];
  const data = sheet
    .getRange(config.firstDataRow, 1, lastRow - config.firstDataRow + 1, lastColumn)
    .getDisplayValues();
  const keyword = normalizeFilterValue_(filterValues[config.keywordColumn - 1]);
  const activeFilters = [];

  config.filterHeaders.forEach(function (headerName) {
    const columnIndex = headers.indexOf(headerName);
    if (columnIndex < 0 || columnIndex >= 8 || columnIndex === config.keywordColumn - 1) return;

    const value = String(filterValues[columnIndex] || '').trim();
    if (value && value !== 'すべて') {
      activeFilters.push({
        header: headerName,
        columnIndex: columnIndex,
        value: value
      });
    }
  });

  const rowsToHide = [];

  data.forEach(function (row, index) {
    const matchesKeyword = !keyword || row.some(function (cellValue) {
      return normalizeFilterValue_(cellValue).includes(keyword);
    });
    const matchesFilters = activeFilters.every(function (filter) {
      return matchesInventoryFilter_(row[filter.columnIndex], filter);
    });

    if (!matchesKeyword || !matchesFilters) {
      rowsToHide.push(config.firstDataRow + index);
    }
  });

  hideRowGroups_(sheet, rowsToHide);
}

function matchesInventoryFilter_(cellValue, filter) {
  const normalizedValue = normalizeFilterValue_(cellValue);

  if (filter.header === '在庫数') {
    const numericStock = Number(String(cellValue).replace(/,/g, '')) || 0;
    return filter.value === '在庫あり' ? numericStock > 0 : numericStock <= 0;
  }

  if (filter.header === '在庫有無') {
    if (normalizedValue === '在庫あり' || normalizedValue === '在庫なし') {
      return normalizedValue === normalizeFilterValue_(filter.value);
    }

    const numericStock = Number(String(cellValue).replace(/,/g, '')) || 0;
    return filter.value === '在庫あり' ? numericStock > 0 : numericStock <= 0;
  }

  return normalizedValue === normalizeFilterValue_(filter.value);
}

function hideRowGroups_(sheet, rowNumbers) {
  if (!rowNumbers.length) return;

  let startRow = rowNumbers[0];
  let previousRow = rowNumbers[0];

  for (let i = 1; i <= rowNumbers.length; i++) {
    const currentRow = rowNumbers[i];

    if (currentRow === previousRow + 1) {
      previousRow = currentRow;
      continue;
    }

    sheet.hideRows(startRow, previousRow - startRow + 1);
    startRow = currentRow;
    previousRow = currentRow;
  }
}

function clearInventoryFilters() {
  const config = INVENTORY_FILTER_CONFIG;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);

  if (!sheet) {
    throw new Error('「商品在庫一覧」シートが見つかりません。');
  }

  const headers = getInventoryHeaders_(sheet);
  const cellsToClear = [config.keywordColumn];

  config.filterHeaders.forEach(function (headerName) {
    const columnIndex = headers.indexOf(headerName) + 1;
    if (columnIndex && !cellsToClear.includes(columnIndex)) {
      cellsToClear.push(columnIndex);
    }
  });

  cellsToClear.forEach(function (columnIndex) {
    sheet.getRange(config.filterRow, columnIndex).clearContent();
  });

  const maxRows = sheet.getMaxRows();
  if (maxRows >= config.firstDataRow) {
    sheet.showRows(config.firstDataRow, maxRows - config.firstDataRow + 1);
  }
}

function getInventoryHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return [];

  return sheet
    .getRange(INVENTORY_FILTER_CONFIG.headerRow, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(function (header) { return String(header).trim(); });
}

function normalizeFilterValue_(value) {
  return String(value || '').trim().toLowerCase();
}

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');

  template.barcode = e.parameter.barcode || '';
  template.newBarcode = e.parameter.newBarcode || '';

  return template.evaluate()
    .setTitle('cocokaruco 在庫管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function registerProduct(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!data.barcode && data.generateBarcode) {
    data.barcode = generateInternalEan13Barcode_();
  }
  const sheet = ss.getSheetByName('商品在庫一覧');

  const lastRow = sheet.getLastRow();
  const nextRow = lastRow + 1;

  sheet.getRange(nextRow, 1, 1, 13).setValues([[
    data.productNumber || '',
    data.barcode || '',
    data.productName || '',
    data.type || '',
    data.brand || '',
    data.color || '',
    data.size || '',
    data.maker || '',
    data.arrivalDate || '',
    data.cost || '',
    data.price || '',
    data.stock || '',
    data.instagram || ''
  ]]);

  // 新しく追加された種類・ブランド・カラー・サイズ等を
  // フィルター候補へ即時反映する。
  setupInventoryFilters_(sheet);

  // 登録時にフィルターが掛かっている場合も、その条件を維持して再適用。
  applyInventoryFilters_(sheet, getInventoryHeaders_(sheet));

  return '登録しました！';
}

function checkBarcode(barcode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('商品在庫一覧');

  const lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    return { found: false };
  }

  const data = sheet
    .getRange(3, 2, lastRow - 2, 11)
    .getDisplayValues();

  const matches = data.filter(function(row) {
    return String(row[0]).trim() === String(barcode).trim();
  });

  if (matches.length === 0) {
    return { found: false };
  }

  const row = matches[0];

  return {
    found: true,
    count: matches.length,
    barcode: row[0],
    productName: row[1],
    type: row[2],
    brand: row[3],
    color: row[4],
    size: row[5],
    stock: row[10]
  };
}

function addStockByBarcode(barcode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('商品在庫一覧');

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return { success: false };

  const barcodes = sheet
    .getRange(3, 2, lastRow - 2, 1)
    .getDisplayValues();

  for (let i = 0; i < barcodes.length; i++) {
    if (String(barcodes[i][0]).trim() === String(barcode).trim()) {
      const targetRow = i + 3;
      const stockCell = sheet.getRange(targetRow, 12);
      const currentStock = Number(stockCell.getValue()) || 0;

      stockCell.setValue(currentStock + 1);

      return {
        success: true,
        newStock: currentStock + 1
      };
    }
  }

  return { success: false };
}


/**
 * 店内管理用バーコードを CK000001 形式で安全に採番する。
 * ScriptLock により同時登録時の重複を防止する。
 */
function generateInternalBarcode() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('商品在庫一覧');
    if (!sheet) throw new Error('「商品在庫一覧」シートが見つかりません。');

    const lastRow = sheet.getLastRow();
    const values = lastRow >= 3
      ? sheet.getRange(3, 2, lastRow - 2, 1).getDisplayValues().flat()
      : [];

    let maxNumber = 0;
    values.forEach(function(value) {
      const match = String(value).trim().match(/^CK(\d{6})$/i);
      if (match) maxNumber = Math.max(maxNumber, Number(match[1]));
    });

    const nextNumber = maxNumber + 1;
    if (nextNumber > 999999) throw new Error('自社バーコードの採番上限に達しました。');
    return 'CK' + String(nextNumber).padStart(6, '0');
  } finally {
    lock.releaseLock();
  }
}


function generateInternalEan13Barcode_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    let seq = Number(props.getProperty('INTERNAL_EAN_SEQ') || 0) + 1;
    if (seq > 9999999999) throw new Error('自社バーコード採番上限です。');
    const body12 = '20' + String(seq).padStart(10, '0');
    let sum = 0;
    for (let i = 0; i < 12; i++) sum += Number(body12[i]) * (i % 2 === 0 ? 1 : 3);
    const code = body12 + String((10 - (sum % 10)) % 10);
    props.setProperty('INTERNAL_EAN_SEQ', String(seq));
    return code;
  } finally {
    lock.releaseLock();
  }
}


function generateUniqueEan13Barcode() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('商品在庫一覧');
    if (!sheet) throw new Error('商品在庫一覧シートが見つかりません。');

    const lastRow = sheet.getLastRow();
    const used = new Set(
      lastRow >= 3
        ? sheet.getRange(3, 2, lastRow - 2, 1).getDisplayValues().flat().map(String)
        : []
    );

    const props = PropertiesService.getScriptProperties();
    let seq = Number(props.getProperty('INTERNAL_EAN_SEQ') || 0);

    for (let attempt = 0; attempt < 100000; attempt++) {
      seq++;
      if (seq > 9999999999) throw new Error('自社バーコード採番上限です。');

      const body12 = '20' + String(seq).padStart(10, '0');
      let sum = 0;
      for (let i = 0; i < 12; i++) {
        sum += Number(body12[i]) * (i % 2 === 0 ? 1 : 3);
      }
      const code = body12 + String((10 - (sum % 10)) % 10);

      if (!used.has(code)) {
        props.setProperty('INTERNAL_EAN_SEQ', String(seq));
        return code;
      }
    }
    throw new Error('未使用バーコードを発行できませんでした。');
  } finally {
    lock.releaseLock();
  }
}


/**
 * 販売用：バーコードから商品情報を取得する。
 */
function getProductForSale(barcode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('商品在庫一覧');
  if (!sheet) throw new Error('「商品在庫一覧」シートが見つかりません。');
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return { found: false };
  const data = sheet.getRange(3, 1, lastRow - 2, 12).getDisplayValues();
  const target = String(barcode || '').trim();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]).trim() === target) {
      return {
        found: true, row: i + 3, productNumber: data[i][0], barcode: data[i][1],
        productName: data[i][2], type: data[i][3], brand: data[i][4],
        color: data[i][5], size: data[i][6], maker: data[i][7],
        price: data[i][10], stock: Number(String(data[i][11]).replace(/,/g, '')) || 0
      };
    }
  }
  return { found: false };
}

/**
 * 指定数量を販売し、在庫を減算して「販売履歴」と「日別売上」に記録する。
 * ScriptLockで二重タップ・同時販売による在庫競合を防ぐ。
 */
function sellProductByBarcode(barcode, quantity) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('商品在庫一覧');
    if (!sheet) throw new Error('「商品在庫一覧」シートが見つかりません。');
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return { success: false, reason: 'not_found' };
    const barcodes = sheet.getRange(3, 2, lastRow - 2, 1).getDisplayValues();
    for (let i = 0; i < barcodes.length; i++) {
      if (String(barcodes[i][0]).trim() !== String(barcode || '').trim()) continue;
      const row = i + 3;
      const values = sheet.getRange(row, 1, 1, 12).getDisplayValues()[0];
      const stockCell = sheet.getRange(row, 12);
      const currentStock = Number(stockCell.getValue()) || 0;
      if (currentStock <= 0) return { success: false, reason: 'out_of_stock', stock: 0 };
      const saleQuantity = Math.floor(Number(quantity == null ? 1 : quantity));
      if (!isFinite(saleQuantity) || saleQuantity < 1) {
        return { success: false, reason: 'invalid_quantity', stock: currentStock };
      }
      if (saleQuantity > currentStock) {
        return { success: false, reason: 'out_of_stock', stock: currentStock };
      }
      const newStock = currentStock - saleQuantity;
      const soldAt = new Date();
      const salePrice = Number(String(values[10] || '').replace(/[^\d.-]/g, '')) || 0;
      const history = setupSalesHistorySheet_(ss);

      // 履歴と日別集計を先に記録し、両方が成功した販売だけ在庫へ反映する。
      history.appendRow([
        soldAt, values[0], values[1], values[2], values[4],
        values[5], values[6], salePrice, saleQuantity, newStock, salePrice * saleQuantity
      ]);
      history.getRange(history.getLastRow(), 1).setNumberFormat('yyyy/MM/dd HH:mm');
      history.getRange(history.getLastRow(), 8).setNumberFormat('¥#,##0');
      history.getRange(history.getLastRow(), 11).setNumberFormat('¥#,##0');

      rebuildDailySalesSummary_(ss, history);
      stockCell.setValue(newStock);

      return {
        success: true, newStock: newStock, quantity: saleQuantity, barcode: values[1],
        productName: values[2], brand: values[4], color: values[5],
        size: values[6], price: salePrice
      };
    }
    return { success: false, reason: 'not_found' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 販売履歴シートを既存データを残したまま最新の列構成へ整える。
 */
function setupSalesHistorySheet_(ss) {
  let sheet = ss.getSheetByName('販売履歴');
  if (!sheet) sheet = ss.insertSheet('販売履歴');

  const headers = [
    '販売日時', '商品番号', 'バーコード', '商品名', 'ブランド名', 'カラー',
    'サイズ', '販売価格', '数量', '販売後在庫', '売上金額'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('yyyy/MM/dd HH:mm');
  sheet.getRange('H:H').setNumberFormat('¥#,##0');
  sheet.getRange('K:K').setNumberFormat('¥#,##0');
  return sheet;
}

/**
 * 「販売履歴」全体から日ごとの点数と売上金額を再集計する。
 * 過去の履歴や後から修正された履歴も集計結果へ反映される。
 */
function rebuildDailySalesSummary_(ss, history) {
  let sheet = ss.getSheetByName('日別売上');
  if (!sheet) sheet = ss.insertSheet('日別売上');

  const headers = ['売上日', '販売点数', '売上合計', '最終販売日時'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  const timeZone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const summary = {};
  const historyLastRow = history.getLastRow();

  if (historyLastRow >= 2) {
    const rows = history.getRange(2, 1, historyLastRow - 1, 11).getValues();
    rows.forEach(function(row) {
      const soldAt = row[0];
      if (!(soldAt instanceof Date) || isNaN(soldAt.getTime())) return;

      const dateKey = Utilities.formatDate(soldAt, timeZone, 'yyyy/MM/dd');
      const quantity = Number(row[8]) || 0;
      const hasRecordedAmount = row[10] !== '' && row[10] !== null;
      const amount = hasRecordedAmount
        ? Number(row[10]) || 0
        : (Number(row[7]) || 0) * quantity;

      if (!summary[dateKey]) {
        summary[dateKey] = { quantity: 0, amount: 0, lastSoldAt: soldAt };
      }
      summary[dateKey].quantity += quantity;
      summary[dateKey].amount += amount;
      if (soldAt > summary[dateKey].lastSoldAt) summary[dateKey].lastSoldAt = soldAt;
    });
  }

  const existingRows = Math.max(sheet.getLastRow() - 1, 0);
  if (existingRows) sheet.getRange(2, 1, existingRows, 4).clearContent();

  const output = Object.keys(summary).sort().reverse().map(function(dateKey) {
    const item = summary[dateKey];
    return [dateKey, item.quantity, item.amount, item.lastSoldAt];
  });

  if (output.length) {
    sheet.getRange(2, 1, output.length, 4).setValues(output);
    sheet.getRange(2, 3, output.length, 1).setNumberFormat('¥#,##0');
    sheet.getRange(2, 4, output.length, 1).setNumberFormat('yyyy/MM/dd HH:mm');
  }
}

