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

  return template.evaluate()
    .setTitle('cocokaruco 商品登録')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function registerProduct(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
