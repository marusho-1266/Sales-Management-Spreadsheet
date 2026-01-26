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
  
  // ヘッダー行を設定（拡張版）
  const headers = [
    'サイト名',
    'URLパターン（カンマ区切り）',
    '価格セレクタ（カンマ区切り）',
    '価格除外セレクタ（カンマ区切り）',
    '在庫セレクタ（カンマ区切り）',
    '在庫ありキーワード（カンマ区切り）',
    '売り切れキーワード（カンマ区切り）',
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
    ['Amazon', 'amazon.co.jp,amazon.com', '#priceblock_ourprice,#priceblock_dealprice,.a-price-whole', '', '#availability span,#availability', '在庫,stock,available', '売り切れ,out of stock,unavailable', 5.0, 2, '有効', 'ASINコード必須'],
    ['楽天市場', 'rakuten.co.jp', '#itemPrice,[id*="itemPrice"],[class*="item-price"],span.priceBox__price,.priceBox__price', '[class*="special"],[class*="offer"],[class*="card"]', '.stock,[class*="stock"],[class*="在庫"]', '在庫あり,在庫,購入可能,available', '売り切れ,在庫なし,完売,sold out,out of stock', 3.0, 3, '有効', '楽天ポイント考慮'],
    ['Yahoo!ショッピング', 'yahoo.co.jp,shopping.yahoo.co.jp', '.elPriceNumber,.Price__value,[data-testid="price"]', '', '.elStockStatus,[data-testid="stock-status"]', '在庫あり', '在庫なし,売り切れ', 3.5, 3, '有効', 'Yahooポイント考慮'],
    ['メルカリ', 'mercari.com,mercari.jp', '.item-price,[data-testid="price"],.merPrice', '', '.item-status,[data-testid="status"],.sold-out', '在庫あり', '売り切れ,取引中,sold', 10.0, 5, '有効', '個人売買サイト'],
    ['ヤフオク', 'page.auctions.yahoo.co.jp', '.price', '', '', '入札する', '終了', 8.0, 5, '有効', 'オークションサイト'],
    ['ヤフーフリマ', 'fril.jp', '.price', '', '', '購入手続きへ', '売り切れ', 10.0, 5, '有効', '個人売買サイト'],
    ['個人ショップ', 'example.com', '.price,#price', '', '.stock', '在庫あり', '売り切れ', 0.0, 10, '無効', '手動設定用']
  ];
  
  const dataRange = sheet.getRange(2, 1, sampleData.length, sampleData[0].length);
  dataRange.setValues(sampleData);
  
  // 数値列の書式設定
  sheet.getRange(2, 8, sampleData.length, 1).setNumberFormat('0.0'); // 手数料率（パーセント値として表示）
  sheet.getRange(2, 9, sampleData.length, 1).setNumberFormat('0'); // アクセス間隔
}

/**
 * 仕入れ元マスターの条件付き書式設定
 */
function setupSupplierConditionalFormatting(sheet) {
  // 有効フラグによる色分け（列J）
  const statusRange = sheet.getRange('J:J');
  
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
  
  // 手数料率による色分け（列H）
  const feeRange = sheet.getRange('H:H');
  
  // 手数料率が高い - 黄色（8%以上）
  const highFeeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(8.0)
    .setBackground('#fff2cc')
    .setRanges([feeRange])
    .build();
  
  // ルールを適用
  const rules = [activeRule, inactiveRule, highFeeRule];
  sheet.setConditionalFormatRules(rules);
}
