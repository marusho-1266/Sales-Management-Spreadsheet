/**
 * 仕入れ元マスターシートの作成と設定
 * 在庫管理ツール システム開発
 * 作成日: 2025-09-27
 */

/**
 * 仕入れ元マスターシートの初期化
 */
function initializeSupplierMasterSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 仕入れ元マスターシートを作成または取得
  let supplierSheet = spreadsheet.getSheetByName('仕入れ元マスター');
  if (!supplierSheet) {
    supplierSheet = spreadsheet.insertSheet('仕入れ元マスター');
  }
  
  // 既存のデータをクリア
  supplierSheet.clear();
  
  // ヘッダー行を設定
  const headers = [
    'サイト名',
    'サイトURL',
    '在庫あり判定文字列',
    '売り切れ判定文字列',
    '価格取得セレクタ',
    '手数料率(%)',
    'アクセス間隔(秒)',
    '有効フラグ',
    '備考'
  ];
  
  // ヘッダーを設定
  supplierSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーの書式設定
  const headerRange = supplierSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ea4335');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅の自動調整
  supplierSheet.autoResizeColumns(1, headers.length);
  
  // サンプルデータを追加
  addSampleSupplierData(supplierSheet);
  
  // 条件付き書式を設定
  setupSupplierConditionalFormatting(supplierSheet);
  
  console.log('仕入れ元マスターシートの初期化が完了しました');
}

/**
 * サンプル仕入れ元データの追加
 */
function addSampleSupplierData(sheet) {
  const sampleData = [
    ['Amazon', 'https://amazon.co.jp', 'カートに入れる', '在庫切れ', '.a-price-whole', 5.0, 2, '有効', 'ASINコード必須'],
    ['楽天市場', 'https://item.rakuten.co.jp', 'カートに入れる', '売り切れ', '.price', 3.0, 3, '有効', '楽天ポイント考慮'],
    ['Yahooショッピング', 'https://shopping.yahoo.co.jp', '購入手続きへ', 'SOLD OUT', '.price', 3.5, 3, '有効', 'Yahooポイント考慮'],
    ['メルカリ', 'https://mercari.com', '購入手続きへ', '売り切れました', '.item-price', 10.0, 5, '有効', '個人売買サイト'],
    ['ヤフオク', 'https://page.auctions.yahoo.co.jp', '入札する', '終了', '.price', 8.0, 5, '有効', 'オークションサイト'],
    ['ヤフーフリマ', 'https://fril.jp', '購入手続きへ', '売り切れ', '.price', 10.0, 5, '有効', '個人売買サイト'],
    ['個人ショップ', 'https://example.com', '購入手続きへ', '売り切れ', '.price', 0.0, 10, '無効', '手動設定用']
  ];
  
  const dataRange = sheet.getRange(2, 1, sampleData.length, sampleData[0].length);
  dataRange.setValues(sampleData);
  
  // 数値列の書式設定
  sheet.getRange(2, 6, sampleData.length, 1).setNumberFormat('0.0%'); // 手数料率
  sheet.getRange(2, 7, sampleData.length, 1).setNumberFormat('0'); // アクセス間隔
}

/**
 * 仕入れ元マスターの条件付き書式設定
 */
function setupSupplierConditionalFormatting(sheet) {
  // 有効フラグによる色分け
  const statusRange = sheet.getRange('H:H');
  
  // 有効 - 緑色
  const activeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('有効')
    .setBackground('#d9ead3')
    .setRanges([statusRange])
    .build();
  
  // 無効 - 赤色
  const inactiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('無効')
    .setBackground('#f4cccc')
    .setRanges([statusRange])
    .build();
  
  // 手数料率による色分け
  const feeRange = sheet.getRange('F:F');
  
  // 手数料率が高い - 黄色
  const highFeeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0.08)
    .setBackground('#fff2cc')
    .setRanges([feeRange])
    .build();
  
  // ルールを適用
  const rules = [activeRule, inactiveRule, highFeeRule];
  sheet.setConditionalFormatRules(rules);
}
