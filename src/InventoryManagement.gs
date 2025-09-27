/**
 * 在庫管理シートの作成と設定
 * 在庫管理ツール システム開発
 * 作成日: 2025-09-27
 */

/**
 * 在庫管理シートの初期化
 */
function initializeInventorySheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 在庫管理シートを作成または取得
  let inventorySheet = spreadsheet.getSheetByName('在庫管理');
  if (!inventorySheet) {
    inventorySheet = spreadsheet.insertSheet('在庫管理');
  }
  
  // 既存のデータをクリア
  inventorySheet.clear();
  
  // ヘッダー行を設定
  const headers = [
    '商品ID',
    '商品名', 
    'SKU',
    'ASIN',
    '仕入れ元',
    '仕入れ元URL',
    '仕入れ価格',
    '販売価格',
    '重量',
    '在庫数',
    '在庫ステータス',
    '利益',
    '最終更新日時'
  ];
  
  // ヘッダーを設定
  inventorySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーの書式設定
  const headerRange = inventorySheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅の自動調整
  inventorySheet.autoResizeColumns(1, headers.length);
  
  // サンプルデータを追加
  addSampleData(inventorySheet);
  
  // 条件付き書式を設定
  setupConditionalFormatting(inventorySheet);
  
  // 利益計算式を設定
  setupProfitCalculation(inventorySheet);
  
  console.log('在庫管理シートの初期化が完了しました');
}

/**
 * サンプルデータの追加
 */
function addSampleData(sheet) {
  const sampleData = [
    [1, 'iPhone 15 Pro 128GB', 'IPH15P-128', 'B0CHX1W1XY', 'Amazon', 'https://amazon.co.jp/dp/B0CHX1W1XY', 120000, 150000, 187, 5, '在庫あり', '', '2025-09-27 00:00:00'],
    [2, 'MacBook Air M2 13インチ', 'MBA-M2-13', 'B0B3C2Q5XK', '楽天', 'https://item.rakuten.co.jp/example/macbook-air-m2', 140000, 180000, 1240, 3, '在庫あり', '', '2025-09-27 00:00:00'],
    [3, 'AirPods Pro 第2世代', 'APP-2ND', 'B0BDJDRJ9T', 'Yahooショッピング', 'https://shopping.yahoo.co.jp/products/airpods-pro-2nd', 25000, 35000, 56, 0, '売り切れ', '', '2025-09-27 00:00:00'],
    [4, 'iPad Air 第5世代', 'IPAD-AIR-5', 'B09V4HCN9V', 'メルカリ', 'https://mercari.com/items/m123456789', 60000, 80000, 461, 2, '在庫あり', '', '2025-09-27 00:00:00'],
    [5, 'Apple Watch Series 9', 'AWS-9', 'B0CHX1W1XZ', 'ヤフオク', 'https://page.auctions.yahoo.co.jp/jp/auction/example', 45000, 60000, 39, 1, '在庫あり', '', '2025-09-27 00:00:00']
  ];
  
  const dataRange = sheet.getRange(2, 1, sampleData.length, sampleData[0].length);
  dataRange.setValues(sampleData);
  
  // 数値列の書式設定
  sheet.getRange(2, 1, sampleData.length, 1).setNumberFormat('0'); // 商品ID
  sheet.getRange(2, 7, sampleData.length, 1).setNumberFormat('#,##0'); // 仕入れ価格
  sheet.getRange(2, 8, sampleData.length, 1).setNumberFormat('#,##0'); // 販売価格
  sheet.getRange(2, 9, sampleData.length, 1).setNumberFormat('0'); // 重量
  sheet.getRange(2, 10, sampleData.length, 1).setNumberFormat('0'); // 在庫数
  sheet.getRange(2, 12, sampleData.length, 1).setNumberFormat('#,##0'); // 利益
}

/**
 * 条件付き書式の設定
 */
function setupConditionalFormatting(sheet) {
  // 在庫ステータスによる色分け
  const statusRange = sheet.getRange('K:K');
  
  // 在庫あり - 緑色
  const inStockRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('在庫あり')
    .setBackground('#d9ead3')
    .setRanges([statusRange])
    .build();
  
  // 売り切れ - 赤色
  const outOfStockRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('売り切れ')
    .setBackground('#f4cccc')
    .setRanges([statusRange])
    .build();
  
  // 在庫数による色分け
  const stockRange = sheet.getRange('J:J');
  
  // 在庫数1以下 - 黄色
  const lowStockRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(1)
    .setBackground('#fff2cc')
    .setRanges([stockRange])
    .build();
  
  // 在庫数0 - 赤色
  const noStockRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground('#f4cccc')
    .setRanges([stockRange])
    .build();
  
  // ルールを適用
  const rules = [inStockRule, outOfStockRule, lowStockRule, noStockRule];
  sheet.setConditionalFormatRules(rules);
}

/**
 * 利益計算式の設定
 */
function setupProfitCalculation(sheet) {
  const lastRow = sheet.getLastRow();
  
  // 利益計算式を設定（販売価格 - 仕入れ価格 - 手数料）
  // 手数料は仮に仕入れ価格の5%とする
  for (let row = 2; row <= lastRow; row++) {
    const formula = `=H${row}-G${row}-(G${row}*0.05)`;
    sheet.getRange(row, 12).setFormula(formula);
  }
}
