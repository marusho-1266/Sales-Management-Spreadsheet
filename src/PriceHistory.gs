/**
 * 価格履歴シートの作成と設定
 * 在庫管理ツール システム開発
 * 作成日: 2025-09-27
 */

/**
 * 価格履歴シートの初期化
 */
function initializePriceHistorySheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 価格履歴シートを作成または取得
  let priceHistorySheet = spreadsheet.getSheetByName('価格履歴');
  if (!priceHistorySheet) {
    priceHistorySheet = spreadsheet.insertSheet('価格履歴');
  }
  
  // 既存のデータをクリア
  priceHistorySheet.clear();
  
  // ヘッダー行を設定
  const headers = [
    '商品ID',
    '商品名',
    'チェック日時',
    '仕入れ価格',
    '販売価格',
    '価格変動フラグ',
    '仕入れ価格変動',
    '販売価格変動',
    '変動率(%)',
    '備考'
  ];
  
  // ヘッダーを設定
  priceHistorySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーの書式設定
  const headerRange = priceHistorySheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅の自動調整
  priceHistorySheet.autoResizeColumns(1, headers.length);
  
  // サンプルデータを追加
  addSamplePriceHistoryData(priceHistorySheet);
  
  // 条件付き書式を設定
  setupPriceHistoryConditionalFormatting(priceHistorySheet);
  
  console.log('価格履歴シートの初期化が完了しました');
}

/**
 * サンプル価格履歴データの追加
 */
function addSamplePriceHistoryData(sheet) {
  const sampleData = [
    [1, 'iPhone 15 Pro 128GB', '2025-09-25 10:00:00', 120000, 150000, '変動なし', 0, 0, 0, '初回登録'],
    [1, 'iPhone 15 Pro 128GB', '2025-09-26 10:00:00', 118000, 150000, '仕入れ価格下落', -2000, 0, -1.67, 'セール価格'],
    [1, 'iPhone 15 Pro 128GB', '2025-09-27 10:00:00', 118000, 148000, '販売価格下落', 0, -2000, -1.33, '競合価格調整'],
    [2, 'MacBook Air M2 13インチ', '2025-09-25 10:00:00', 140000, 180000, '変動なし', 0, 0, 0, '初回登録'],
    [2, 'MacBook Air M2 13インチ', '2025-09-26 10:00:00', 140000, 180000, '変動なし', 0, 0, 0, '価格維持'],
    [2, 'MacBook Air M2 13インチ', '2025-09-27 10:00:00', 142000, 180000, '仕入れ価格上昇', 2000, 0, 1.43, '為替変動'],
    [3, 'AirPods Pro 第2世代', '2025-09-25 10:00:00', 25000, 35000, '変動なし', 0, 0, 0, '初回登録'],
    [3, 'AirPods Pro 第2世代', '2025-09-26 10:00:00', 25000, 35000, '変動なし', 0, 0, 0, '価格維持'],
    [3, 'AirPods Pro 第2世代', '2025-09-27 10:00:00', 25000, 35000, '変動なし', 0, 0, 0, '価格維持']
  ];
  
  const dataRange = sheet.getRange(2, 1, sampleData.length, sampleData[0].length);
  dataRange.setValues(sampleData);
  
  // 数値列の書式設定
  sheet.getRange(2, 1, sampleData.length, 1).setNumberFormat('0'); // 商品ID
  sheet.getRange(2, 4, sampleData.length, 1).setNumberFormat('#,##0'); // 仕入れ価格
  sheet.getRange(2, 5, sampleData.length, 1).setNumberFormat('#,##0'); // 販売価格
  sheet.getRange(2, 7, sampleData.length, 1).setNumberFormat('#,##0'); // 仕入れ価格変動
  sheet.getRange(2, 8, sampleData.length, 1).setNumberFormat('#,##0'); // 販売価格変動
  sheet.getRange(2, 9, sampleData.length, 1).setNumberFormat('0.00%'); // 変動率
}

/**
 * 価格履歴の条件付き書式設定
 */
function setupPriceHistoryConditionalFormatting(sheet) {
  // 価格変動フラグによる色分け
  const flagRange = sheet.getRange('F:F');
  
  // 変動なし - グレー
  const noChangeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('変動なし')
    .setBackground('#f5f5f5')
    .setRanges([flagRange])
    .build();
  
  // 価格上昇 - 赤色
  const priceUpRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('上昇')
    .setBackground('#f4cccc')
    .setRanges([flagRange])
    .build();
  
  // 価格下落 - 緑色
  const priceDownRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('下落')
    .setBackground('#d9ead3')
    .setRanges([flagRange])
    .build();
  
  // 変動率による色分け
  const changeRateRange = sheet.getRange('I:I');
  
  // 大幅上昇 - 濃い赤
  const highIncreaseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0.05)
    .setBackground('#ffcdd2')
    .setRanges([changeRateRange])
    .build();
  
  // 大幅下落 - 濃い緑
  const highDecreaseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(-0.05)
    .setBackground('#c8e6c9')
    .setRanges([changeRateRange])
    .build();
  
  // ルールを適用
  const rules = [noChangeRule, priceUpRule, priceDownRule, highIncreaseRule, highDecreaseRule];
  sheet.setConditionalFormatRules(rules);
}
