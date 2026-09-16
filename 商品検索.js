function onEdit(e) {
  const sheet = e.range.getSheet();

  if (e.range.getRow() !== 1 || e.range.getColumn() !== 2) return;
  if (sheet.getRange('A2').getDisplayValue() !== '商品番号') return;

  const firstDataRow = 3;
  const maxRows = sheet.getMaxRows();

  sheet.showRows(firstDataRow, maxRows - firstDataRow + 1);

  const keyword = String(e.range.getDisplayValue())
    .trim()
    .toLowerCase();

  if (!keyword) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < firstDataRow) return;

  const data = sheet
    .getRange(firstDataRow, 1, lastRow - firstDataRow + 1, 15)
    .getDisplayValues();

  data.forEach(function (row, index) {
    const searchableText = [
      row[0],
      row[1],
      row[2],
      row[3],
      row[4]
    ].join(' ').toLowerCase();

    if (!searchableText.includes(keyword)) {
      sheet.hideRows(firstDataRow + index);
    }
  });
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

  return '登録しました！';
}

function checkBarcode(barcode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('商品在庫一覧');

  const lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    return { found: false };
  }

  // B列〜L列を取得
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

    barcode: row[0],      // B バーコード番号
    productName: row[1],  // C 商品名
    type: row[2],         // D 種類
    brand: row[3],        // E ブランド名
    color: row[4],        // F カラー
    size: row[5],         // G サイズ
    stock: row[10]        // L 在庫数
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
      const stockCell = sheet.getRange(targetRow, 12); // L列
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

