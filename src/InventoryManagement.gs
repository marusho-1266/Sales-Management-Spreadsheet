/**
 * 在庫管理シートの作成と設定
 * 在庫管理ツール システム開発
 * 作成日: 2025-09-27
 */

// 設定値キャッシュ（キー→値のマップ）
let settingsCache = null;

/**
 * 在庫管理シートの初期化
 */
function initializeInventorySheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 在庫管理シートを作成または取得
  let inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  if (!inventorySheet) {
    inventorySheet = spreadsheet.insertSheet(SHEET_NAMES.INVENTORY);
  }
  
  // ユーザー確認ダイアログを表示
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '在庫管理シートの初期化',
    '在庫管理シートを初期化すると、既存のデータと書式が削除されます。\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  // ユーザーが「はい」を選択した場合のみクリアを実行
  if (response === ui.Button.YES) {
    inventorySheet.clear();
  } else {
    console.log('在庫管理シートの初期化がキャンセルされました');
    return;
  }
  
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
    // Joom対応フィールド（重量と在庫ステータスの間に挿入）
    '商品説明',
    'メイン画像URL',
    '通貨',
    '配送価格',
    '在庫数量',
    // 既存フィールド
    '在庫ステータス',
    '利益',
    '最終更新日時',
    '備考・メモ',
    // Joom連携管理列
    'Joom連携ステータス',
    '最終出力日時'
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
  
  // Joom対応列のデフォルト値を設定
  setupJoomDefaultValues(inventorySheet);
  
  // 条件付き書式を設定
  setupConditionalFormatting(inventorySheet);
  
  // 利益計算式を設定
  setupProfitCalculation(inventorySheet);
  
  // 価格変更時の自動履歴更新機能を設定
  setupPriceChangeTrigger(inventorySheet);
  
  console.log('在庫管理シートの初期化が完了しました');
}


/**
 * 既存の在庫管理シートにJoom連携管理列を追加
 */
function addJoomStatusColumnsToExistingSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  
  if (!inventorySheet) {
    console.log('在庫管理シートが見つかりません');
    return;
  }
  
  try {
    // 現在のヘッダー行を取得
    const headerRange = inventorySheet.getRange(1, 1, 1, inventorySheet.getLastColumn());
    const headers = headerRange.getValues()[0];
    
    // Joom連携ステータス列が既に存在するかチェック
    if (headers.includes('Joom連携ステータス')) {
      console.log('Joom連携管理列は既に存在します');
      return;
    }
    
    // Joom連携管理列を追加（最後の列の後に2列挿入）
    const lastColumn = inventorySheet.getLastColumn();
    inventorySheet.insertColumnAfter(lastColumn); // Joom連携ステータス列
    inventorySheet.insertColumnAfter(lastColumn + 1); // 最終出力日時列
    
    // 新しいヘッダーを設定
    inventorySheet.getRange(1, lastColumn + 1).setValue('Joom連携ステータス');
    inventorySheet.getRange(1, lastColumn + 2).setValue('最終出力日時');
    
    // ヘッダーの書式設定
    const joomStatusHeaderRange = inventorySheet.getRange(1, lastColumn + 1, 1, 2);
    joomStatusHeaderRange.setBackground('#4285f4');
    joomStatusHeaderRange.setFontColor('#ffffff');
    joomStatusHeaderRange.setFontWeight('bold');
    joomStatusHeaderRange.setHorizontalAlignment('center');
    
    // 列幅を自動調整
    inventorySheet.autoResizeColumns(lastColumn + 1, 2);
    
    // 既存データのJoom連携ステータスを「未連携」に設定
    const lastRow = inventorySheet.getLastRow();
    if (lastRow > 1) {
      const unlinkedRange = inventorySheet.getRange(2, lastColumn + 1, lastRow - 1, 1);
      unlinkedRange.setValue(JOOM_STATUS.UNLINKED);
    }
    
    console.log('Joom連携管理列が正常に追加されました');
    
  } catch (error) {
    console.error('Joom連携管理列の追加中にエラーが発生しました:', error);
    throw error;
  }
}

/**
 * 既存の在庫管理シートに備考列を追加（後方互換性のため残す）
 */
function addNotesColumnToExistingSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  
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
    [1, 'iPhone 15 Pro 128GB', 'IPH15P-128', 'B0CHX1W1XY', 'Amazon', 'https://amazon.co.jp/dp/B0CHX1W1XY', 120000, 150000, 187, '最新のiPhone 15 Pro 128GBモデル。A17 Proチップ搭載で高性能。', 'https://via.placeholder.com/500x500.jpg', 'JPY', 0, 1, '在庫あり', '', '2025-09-27 00:00:00', ''],
    [2, 'MacBook Air M2 13インチ', 'MBA-M2-13', 'B0B3C2Q5XK', '楽天', 'https://item.rakuten.co.jp/example/macbook-air-m2', 140000, 180000, 1240, 'MacBook Air M2 13インチ。M2チップで高速処理。軽量設計。', 'https://example.com/images/macbook-air-m2.jpg', 'JPY', 0, 1, '在庫あり', '', '2025-09-27 00:00:00', ''],
    [3, 'AirPods Pro 第2世代', 'APP-2ND', 'B0BDJDRJ9T', 'Yahooショッピング', 'https://shopping.yahoo.co.jp/products/airpods-pro-2nd', 25000, 35000, 56, 'AirPods Pro 第2世代。ノイズキャンセリング機能搭載。', 'https://example.com/images/airpods-pro-2nd.jpg', 'JPY', 0, 0, '売り切れ', '', '2025-09-27 00:00:00', ''],
    [4, 'iPad Air 第5世代', 'IPAD-AIR-5', 'B09V4HCN9V', 'メルカリ', 'https://mercari.com/items/m123456789', 60000, 80000, 461, 'iPad Air 第5世代。M1チップ搭載で高性能タブレット。', 'https://example.com/images/ipad-air-5.jpg', 'JPY', 0, 1, '在庫あり', '', '2025-09-27 00:00:00', ''],
    [5, 'Apple Watch Series 9', 'AWS-9', 'B0CHX1W1XZ', 'ヤフオク', 'https://page.auctions.yahoo.co.jp/jp/auction/example', 45000, 60000, 39, 'Apple Watch Series 9。健康管理とスマートウォッチ機能。', 'https://example.com/images/apple-watch-s9.jpg', 'JPY', 0, 1, '在庫あり', '', '2025-09-27 00:00:00', '']
  ];
  
  const dataRange = sheet.getRange(2, 1, sampleData.length, sampleData[0].length);
  dataRange.setValues(sampleData);
  
  // 数値列の書式設定
  sheet.getRange(2, 1, sampleData.length, 1).setNumberFormat('0'); // 商品ID
  sheet.getRange(2, 7, sampleData.length, 1).setNumberFormat('#,##0'); // 仕入れ価格
  sheet.getRange(2, 8, sampleData.length, 1).setNumberFormat('#,##0'); // 販売価格
  sheet.getRange(2, 9, sampleData.length, 1).setNumberFormat('0'); // 重量
  sheet.getRange(2, 13, sampleData.length, 1).setNumberFormat('#,##0'); // 配送価格
  sheet.getRange(2, 14, sampleData.length, 1).setNumberFormat('0'); // 在庫数量
  sheet.getRange(2, 16, sampleData.length, 1).setNumberFormat('#,##0'); // 利益
}

/**
 * 条件付き書式の設定
 */
function setupConditionalFormatting(sheet) {
  // 在庫ステータスによる色分け
  const statusRange = sheet.getRange('O:O');
  
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
    sheet.getRange(row, 16).setFormula(formula);
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

/**
 * Joom対応列のデフォルト値を設定
 */
function setupJoomDefaultValues(sheet) {
  // COLUMN_INDEXES定数の存在チェック
  if (typeof COLUMN_INDEXES === 'undefined') {
    console.error('COLUMN_INDEXES定数が定義されていません');
    return;
  }
  
  if (typeof COLUMN_INDEXES.INVENTORY === 'undefined') {
    console.error('COLUMN_INDEXES.INVENTORY定数が定義されていません');
    return;
  }
  
  // 必要な列インデックスの存在チェック
  const requiredIndexes = ['CURRENCY', 'STOCK_STATUS', 'STOCK_QUANTITY', 'SHIPPING_PRICE'];
  for (const indexName of requiredIndexes) {
    if (typeof COLUMN_INDEXES.INVENTORY[indexName] === 'undefined') {
      console.error(`COLUMN_INDEXES.INVENTORY.${indexName}定数が定義されていません`);
      return;
    }
  }
  
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return; // データ行がない場合はスキップ
  
  // デフォルト値を設定
  const defaultCurrency = getSetting('デフォルト通貨') || 'JPY';
  const defaultShippingPrice = getSetting('デフォルト配送価格') || '0';
  
  for (let row = 2; row <= lastRow; row++) {
    // 通貨のデフォルト値
    sheet.getRange(row, COLUMN_INDEXES.INVENTORY.CURRENCY).setValue(defaultCurrency);
    
    // 在庫数量のデフォルト値（在庫ステータスから変換）
    const stockStatus = sheet.getRange(row, COLUMN_INDEXES.INVENTORY.STOCK_STATUS).getValue();
    const stockQuantity = stockStatus === '在庫あり' ? 1 : 0;
    sheet.getRange(row, COLUMN_INDEXES.INVENTORY.STOCK_QUANTITY).setValue(stockQuantity);
    
    // 配送価格のデフォルト値（基数10を明示的に指定）
    sheet.getRange(row, COLUMN_INDEXES.INVENTORY.SHIPPING_PRICE).setValue(parseInt(defaultShippingPrice, 10));
  }
  
  console.log('Joom対応列のデフォルト値を設定しました');
}

/**
 * 在庫ステータスから在庫数量を自動更新
 */
function updateStockQuantityFromStatus() {
  // COLUMN_INDEXES定数の存在チェック
  if (typeof COLUMN_INDEXES === 'undefined') {
    console.error('COLUMN_INDEXES定数が定義されていません');
    return;
  }
  
  if (typeof COLUMN_INDEXES.INVENTORY === 'undefined') {
    console.error('COLUMN_INDEXES.INVENTORY定数が定義されていません');
    return;
  }
  
  if (typeof COLUMN_INDEXES.INVENTORY.STOCK_STATUS === 'undefined') {
    console.error('COLUMN_INDEXES.INVENTORY.STOCK_STATUS定数が定義されていません');
    return;
  }
  
  if (typeof COLUMN_INDEXES.INVENTORY.STOCK_QUANTITY === 'undefined') {
    console.error('COLUMN_INDEXES.INVENTORY.STOCK_QUANTITY定数が定義されていません');
    return;
  }
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  
  if (!inventorySheet) {
    console.log('在庫管理シートが見つかりません');
    return;
  }
  
  const lastRow = inventorySheet.getLastRow();
  
  try {
    for (let row = 2; row <= lastRow; row++) {
      try {
        const stockStatus = inventorySheet.getRange(row, COLUMN_INDEXES.INVENTORY.STOCK_STATUS).getValue();
        const stockQuantity = stockStatus === '在庫あり' ? 1 : 0;
        inventorySheet.getRange(row, COLUMN_INDEXES.INVENTORY.STOCK_QUANTITY).setValue(stockQuantity);
      } catch (error) {
        console.error(`行${row}の在庫数量更新中にエラーが発生しました:`, error);
        // 個別行のエラーは続行
      }
    }
    
    console.log('在庫数量を在庫ステータスから更新しました');
  } catch (error) {
    console.error('在庫数量の一括更新中にエラーが発生しました:', error);
    throw error; // 重大なエラーの場合は再スロー
  }
}

/**
 * 設定シートの初期化
 */
function initializeSettingsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 設定シートを作成または取得
  let settingsSheet = spreadsheet.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = spreadsheet.insertSheet(SHEET_NAMES.SETTINGS);
  }
  
  // ユーザー確認ダイアログを表示
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '設定シートの初期化',
    '設定シートを初期化すると、既存のデータと書式が削除されます。\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  // ユーザーが「はい」を選択した場合のみクリアを実行
  if (response === ui.Button.YES) {
    settingsSheet.clear();
  } else {
    console.log('設定シートの初期化がキャンセルされました');
    return;
  }
  
  // ヘッダー行を設定
  const headers = [
    '設定項目',
    '設定値',
    '説明',
    '最終更新日時'
  ];
  
  // ヘッダーを設定
  settingsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーの書式設定
  const headerRange = settingsSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ff9800');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅の自動調整
  settingsSheet.autoResizeColumns(1, headers.length);
  
  // デフォルト設定値を追加
  addDefaultSettings(settingsSheet);
  
  // 条件付き書式を設定
  setupSettingsConditionalFormatting(settingsSheet);
  
  console.log('設定シートの初期化が完了しました');
}

/**
 * デフォルトの設定値を追加
 */
function addDefaultSettings(sheet) {
  const now = new Date();
  const currentTime = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd HH:mm:ss');
  
  const defaultSettings = [
    ['ストアID', 'STORE001', 'JoomのストアID（全商品共通）', currentTime],
    ['デフォルト通貨', 'JPY', 'デフォルトの通貨設定', currentTime],
    ['デフォルト配送価格', '0', 'デフォルトの配送価格（円）', currentTime],
    ['カテゴリID', '', 'JoomカテゴリID（必要に応じて設定）', currentTime],
    ['検索タグ', '', '商品検索用タグ（カンマ区切り）', currentTime],
    ['危険物種類', 'notdangerous', '配送時の危険物種類', currentTime]
  ];
  
  const dataRange = sheet.getRange(2, 1, defaultSettings.length, defaultSettings[0].length);
  dataRange.setValues(defaultSettings);
  
  // 数値列の書式設定
  sheet.getRange(2, 2, defaultSettings.length, 1).setNumberFormat('@'); // 設定値（テキスト）
  sheet.getRange(2, 4, defaultSettings.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); // 最終更新日時
}

/**
 * 設定シートの条件付き書式設定
 */
function setupSettingsConditionalFormatting(sheet) {
  // 設定値列の書式設定
  const valueRange = sheet.getRange('B:B');
  
  // 必須項目のハイライト（ストアID、デフォルト通貨）
  const requiredRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('STORE001')
    .setBackground('#e8f5e8')
    .setRanges([valueRange])
    .build();
  
  const rules = [requiredRule];
  sheet.setConditionalFormatRules(rules);
}

/**
 * 設定値を取得
 */
function getSetting(settingName) {
  // キャッシュが存在しない場合は初期化
  if (settingsCache === null) {
    loadSettingsCache();
  }
  
  // キャッシュから設定値を取得
  if (settingsCache && settingsCache.has(settingName)) {
    return settingsCache.get(settingName);
  }
  
  console.warn(`設定項目「${settingName}」が見つかりません`);
  return null;
}

/**
 * 設定キャッシュをロード
 */
function loadSettingsCache() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = spreadsheet.getSheetByName(SHEET_NAMES.SETTINGS);
  
  if (!settingsSheet) {
    console.error('設定シートが見つかりません');
    settingsCache = new Map();
    return;
  }
  
  try {
    const data = settingsSheet.getDataRange().getValues();
    settingsCache = new Map();
    
    // ヘッダー行をスキップして設定値をキャッシュに格納
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1] !== undefined) {
        settingsCache.set(data[i][0], data[i][1]);
      }
    }
    
    console.log(`設定キャッシュをロードしました（${settingsCache.size}項目）`);
  } catch (error) {
    console.error('設定キャッシュのロード中にエラーが発生しました:', error);
    settingsCache = new Map();
  }
}

/**
 * 設定値を更新
 */
function updateSetting(settingName, settingValue) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = spreadsheet.getSheetByName(SHEET_NAMES.SETTINGS);
  
  if (!settingsSheet) {
    console.error('設定シートが見つかりません');
    return false;
  }
  
  const data = settingsSheet.getDataRange().getValues();
  
  // ヘッダー行をスキップして設定値を検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === settingName) {
      // 設定値を更新
      settingsSheet.getRange(i + 1, 2).setValue(settingValue);
      
      // 最終更新日時を更新
      const now = new Date();
      const currentTime = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd HH:mm:ss');
      settingsSheet.getRange(i + 1, 4).setValue(currentTime);
      
      // キャッシュを無効化
      invalidateSettingsCache();
      
      console.log(`設定項目「${settingName}」を「${settingValue}」に更新しました`);
      return true;
    }
  }
  
  console.warn(`設定項目「${settingName}」が見つかりません`);
  return false;
}

/**
 * 設定キャッシュを無効化
 */
function invalidateSettingsCache() {
  settingsCache = null;
  console.log('設定キャッシュを無効化しました');
}

/**
 * 設定キャッシュを手動で再読み込み
 */
function reloadSettingsCache() {
  invalidateSettingsCache();
  loadSettingsCache();
  console.log('設定キャッシュを手動で再読み込みしました');
}

