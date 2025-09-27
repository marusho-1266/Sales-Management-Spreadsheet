/**
 * 売上管理シートの作成と設定
 * 在庫管理ツール システム開発
 * 作成日: 2025-09-27
 */

/**
 * 売上管理シートの初期化
 */
function initializeSalesSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 売上管理シートを作成または取得
  let salesSheet = spreadsheet.getSheetByName('売上管理');
  if (!salesSheet) {
    salesSheet = spreadsheet.insertSheet('売上管理');
  }
  
  // 既存のデータをクリア
  salesSheet.clear();
  
  // ヘッダー行を設定
  const headers = [
    '注文ID',
    '注文日',
    '商品ID',
    '商品名',
    'SKU',
    '数量',
    '販売価格',
    '送料',
    '手数料',
    '純利益',
    '仕入れ価格',
    '仕入れ元',
    '備考'
  ];
  
  // ヘッダーを設定
  salesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーの書式設定
  const headerRange = salesSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅の自動調整
  salesSheet.autoResizeColumns(1, headers.length);
  
  // サンプルデータを追加
  addSampleSalesData(salesSheet);
  
  // 条件付き書式を設定
  setupSalesConditionalFormatting(salesSheet);
  
  // 利益計算式を設定
  setupSalesProfitCalculation(salesSheet);
  
  console.log('売上管理シートの初期化が完了しました');
}

/**
 * サンプル売上データの追加
 */
function addSampleSalesData(sheet) {
  const sampleData = [
    ['ORD001', '2025-09-25', 1, 'iPhone 15 Pro 128GB', 'IPH15P-128', 1, 150000, 0, 7500, '', 120000, 'Amazon', 'Joom注文'],
    ['ORD002', '2025-09-26', 2, 'MacBook Air M2 13インチ', 'MBA-M2-13', 1, 180000, 0, 9000, '', 140000, '楽天', 'Joom注文'],
    ['ORD003', '2025-09-26', 4, 'iPad Air 第5世代', 'IPAD-AIR-5', 2, 80000, 0, 4000, '', 60000, 'メルカリ', 'Joom注文'],
    ['ORD004', '2025-09-27', 5, 'Apple Watch Series 9', 'AWS-9', 1, 60000, 0, 3000, '', 45000, 'ヤフオク', 'Joom注文']
  ];
  
  const dataRange = sheet.getRange(2, 1, sampleData.length, sampleData[0].length);
  dataRange.setValues(sampleData);
  
  // 数値列の書式設定
  sheet.getRange(2, 6, sampleData.length, 1).setNumberFormat('0'); // 数量
  sheet.getRange(2, 7, sampleData.length, 1).setNumberFormat('#,##0'); // 販売価格
  sheet.getRange(2, 8, sampleData.length, 1).setNumberFormat('#,##0'); // 送料
  sheet.getRange(2, 9, sampleData.length, 1).setNumberFormat('#,##0'); // 手数料
  sheet.getRange(2, 10, sampleData.length, 1).setNumberFormat('#,##0'); // 純利益
  sheet.getRange(2, 11, sampleData.length, 1).setNumberFormat('#,##0'); // 仕入れ価格
}

/**
 * 売上管理の条件付き書式設定
 */
function setupSalesConditionalFormatting(sheet) {
  // 純利益による色分け
  const profitRange = sheet.getRange('J:J');
  
  // 利益が高い - 緑色
  const highProfitRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(50000)
    .setBackground('#d9ead3')
    .setRanges([profitRange])
    .build();
  
  // 利益が低い - 黄色
  const lowProfitRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(10000)
    .setBackground('#fff2cc')
    .setRanges([profitRange])
    .build();
  
  // 損失 - 赤色
  const lossRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#f4cccc')
    .setRanges([profitRange])
    .build();
  
  // ルールを適用
  const rules = [highProfitRule, lowProfitRule, lossRule];
  sheet.setConditionalFormatRules(rules);
}

/**
 * 売上管理の利益計算式設定
 */
function setupSalesProfitCalculation(sheet) {
  const lastRow = sheet.getLastRow();
  
  // 純利益計算式を設定（販売価格 - 仕入れ価格 - 手数料）
  for (let row = 2; row <= lastRow; row++) {
    const formula = `=G${row}-K${row}-I${row}`;
    sheet.getRange(row, 10).setFormula(formula);
  }
}
