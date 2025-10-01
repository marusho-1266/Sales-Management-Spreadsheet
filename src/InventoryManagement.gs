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
    '在庫ステータス',
    '利益',
    '最終更新日時',
    '備考・メモ'
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
  
  // 価格変更時の自動履歴更新機能を設定
  setupPriceChangeTrigger(inventorySheet);
  
  console.log('在庫管理シートの初期化が完了しました');
}

/**
 * 既存の在庫管理シートに備考列を追加
 */
function addNotesColumnToExistingSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = spreadsheet.getSheetByName('在庫管理');
  
  if (!inventorySheet) {
    console.log('在庫管理シートが見つかりません');
    return;
  }
  
  try {
    // 現在のヘッダー行を取得
    const headerRange = inventorySheet.getRange(1, 1, 1, inventorySheet.getLastColumn());
    const headers = headerRange.getValues()[0];
    
    // 備考列が既に存在するかチェック
    if (headers.includes('備考・メモ')) {
      console.log('備考列は既に存在します');
      return;
    }
    
    // 備考列を追加（M列に挿入）
    inventorySheet.insertColumnAfter(12); // L列の後に挿入
    
    // 新しいヘッダーを設定
    inventorySheet.getRange(1, 13).setValue('備考・メモ');
    
    // ヘッダーの書式設定
    const notesHeaderRange = inventorySheet.getRange(1, 13);
    notesHeaderRange.setBackground('#4285f4');
    notesHeaderRange.setFontColor('#ffffff');
    notesHeaderRange.setFontWeight('bold');
    notesHeaderRange.setHorizontalAlignment('center');
    
    // 列幅を自動調整
    inventorySheet.autoResizeColumn(13);
    
    console.log('備考列が正常に追加されました');
    
  } catch (error) {
    console.error('備考列の追加中にエラーが発生しました:', error);
    throw error;
  }
}

/**
 * サンプルデータの追加
 */
function addSampleData(sheet) {
  const sampleData = [
    [1, 'iPhone 15 Pro 128GB', 'IPH15P-128', 'B0CHX1W1XY', 'Amazon', 'https://amazon.co.jp/dp/B0CHX1W1XY', 120000, 150000, 187, '在庫あり', '', '2025-09-27 00:00:00'],
    [2, 'MacBook Air M2 13インチ', 'MBA-M2-13', 'B0B3C2Q5XK', '楽天', 'https://item.rakuten.co.jp/example/macbook-air-m2', 140000, 180000, 1240, '在庫あり', '', '2025-09-27 00:00:00'],
    [3, 'AirPods Pro 第2世代', 'APP-2ND', 'B0BDJDRJ9T', 'Yahooショッピング', 'https://shopping.yahoo.co.jp/products/airpods-pro-2nd', 25000, 35000, 56, '売り切れ', '', '2025-09-27 00:00:00'],
    [4, 'iPad Air 第5世代', 'IPAD-AIR-5', 'B09V4HCN9V', 'メルカリ', 'https://mercari.com/items/m123456789', 60000, 80000, 461, '在庫あり', '', '2025-09-27 00:00:00'],
    [5, 'Apple Watch Series 9', 'AWS-9', 'B0CHX1W1XZ', 'ヤフオク', 'https://page.auctions.yahoo.co.jp/jp/auction/example', 45000, 60000, 39, '在庫あり', '', '2025-09-27 00:00:00']
  ];
  
  const dataRange = sheet.getRange(2, 1, sampleData.length, sampleData[0].length);
  dataRange.setValues(sampleData);
  
  // 数値列の書式設定
  sheet.getRange(2, 1, sampleData.length, 1).setNumberFormat('0'); // 商品ID
  sheet.getRange(2, 7, sampleData.length, 1).setNumberFormat('#,##0'); // 仕入れ価格
  sheet.getRange(2, 8, sampleData.length, 1).setNumberFormat('#,##0'); // 販売価格
  sheet.getRange(2, 9, sampleData.length, 1).setNumberFormat('0'); // 重量
  sheet.getRange(2, 11, sampleData.length, 1).setNumberFormat('#,##0'); // 利益
}

/**
 * 条件付き書式の設定
 */
function setupConditionalFormatting(sheet) {
  // 在庫ステータスによる色分け
  const statusRange = sheet.getRange('J:J');
  
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
  
  // ルールを適用
  const rules = [inStockRule, outOfStockRule];
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

/**
 * 価格変更時の自動履歴更新機能を設定
 * 注意: この機能は手動でトリガーを設定する必要があります
 */
function setupPriceChangeTrigger(sheet) {
  console.log('価格変更時の自動履歴更新機能の設定をスキップしました');
  console.log('手動でトリガーを設定する場合は、以下の手順に従ってください：');
  console.log('1. スクリプトエディタで「トリガー」を選択');
  console.log('2. 関数: onInventorySheetEdit');
  console.log('3. イベントソース: スプレッドシートから');
  console.log('4. イベントタイプ: 編集時');
  console.log('5. 保存');
}

/**
 * 在庫管理シートの編集時のイベントハンドラー
 * 価格変更を検出して価格履歴を自動更新
 */
function onInventorySheetEdit(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    
    // 在庫管理シートでない場合は処理しない
    if (sheet.getName() !== SHEET_NAMES.INVENTORY) {
      return;
    }
    
    // 仕入れ価格（G列）または販売価格（H列）が変更された場合のみ処理
    const column = range.getColumn();
    if (column !== COLUMN_INDEXES.INVENTORY.PURCHASE_PRICE && 
        column !== COLUMN_INDEXES.INVENTORY.SELLING_PRICE) {
      return;
    }
    
    // 変更された行を取得
    const row = range.getRow();
    if (row < 2) return; // ヘッダー行はスキップ
    
    // 商品IDを取得
    const productId = sheet.getRange(row, COLUMN_INDEXES.INVENTORY.PRODUCT_ID).getValue();
    if (!productId) return;
    
    // 現在の価格を取得
    const purchasePrice = sheet.getRange(row, COLUMN_INDEXES.INVENTORY.PURCHASE_PRICE).getValue();
    const sellingPrice = sheet.getRange(row, COLUMN_INDEXES.INVENTORY.SELLING_PRICE).getValue();
    
    if (purchasePrice && sellingPrice) {
      // 価格履歴を更新
      updatePriceHistory(productId, purchasePrice, sellingPrice, '在庫管理シートから自動更新');
    }
    
  } catch (error) {
    console.error('価格履歴の自動更新中にエラーが発生しました:', error);
  }
}

/**
 * 手動で価格履歴を更新する関数
 * @param {number} productId - 商品ID
 * @param {number} purchasePrice - 仕入れ価格
 * @param {number} sellingPrice - 販売価格
 * @param {string} notes - 備考
 */
function updateInventoryPriceHistory(productId, purchasePrice, sellingPrice, notes = '') {
  try {
    const result = updatePriceHistory(productId, purchasePrice, sellingPrice, notes);
    if (result) {
      console.log(`商品ID ${productId} の価格履歴を更新しました`);
    } else {
      console.error(`商品ID ${productId} の価格履歴更新に失敗しました`);
    }
    return result;
  } catch (error) {
    console.error('価格履歴の更新中にエラーが発生しました:', error);
    return false;
  }
}


