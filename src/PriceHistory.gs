/**
 * 価格履歴シートの作成と設定（1商品1行形式）
 * 在庫管理ツール システム開発
 * 作成日: 2025-09-27
 * 更新日: 2025-10-01
 */

/**
 * 商品IDから商品名を取得するヘルパー関数
 * @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
 * @param {number} productId - 商品ID
 * @returns {string} 商品名（見つからない場合は空文字列）
 */
function getProductNameById(spreadsheet, productId) {
  try {
    const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
    if (!inventorySheet) {
      return '';
    }
    
    const inventoryData = inventorySheet.getDataRange().getValues();
    for (let i = 1; i < inventoryData.length; i++) {
      if (inventoryData[i][0] === productId) {
        return inventoryData[i][1] || '';
      }
    }
    
    return '';
  } catch (error) {
    console.error('商品名取得エラー:', error);
    return '';
  }
}

/**
 * 価格履歴シートの初期化（1商品1行形式）
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
  
  // ヘッダー行を設定（1商品1行形式）
  const headers = [
    '商品ID',
    '商品名',
    '現在仕入れ価格',
    '前回仕入れ価格',
    '仕入れ価格変動',
    '仕入れ価格変動率(%)',
    '現在販売価格',
    '前回販売価格',
    '販売価格変動',
    '販売価格変動率(%)',
    '最終更新日時',
    '仕入れ価格変動回数',
    '販売価格変動回数',
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
  
  console.log('価格履歴シートの初期化が完了しました（1商品1行形式）');
}

/**
 * サンプル価格履歴データの追加（1商品1行形式）
 */
function addSamplePriceHistoryData(sheet) {
  const sampleData = [
    [1, 'iPhone 15 Pro 128GB', 118000, 120000, -2000, -1.67, 148000, 150000, -2000, -1.33, '2025-10-01 22:46:03', 1, 1, '価格調整済み'],
    [2, 'MacBook Air M2 13インチ', 142000, 140000, 2000, 1.43, 180000, 180000, 0, 0, '2025-10-01 22:46:03', 1, 0, '為替変動対応'],
    [3, 'AirPods Pro 第2世代', 25000, 25000, 0, 0, 35000, 35000, 0, 0, '2025-10-01 22:46:03', 0, 0, '価格維持'],
    [4, 'iPad Air 第5世代', 65000, 65000, 0, 0, 85000, 85000, 0, 0, '2025-10-01 22:46:03', 0, 0, '新規登録'],
    [5, 'Apple Watch Series 9', 45000, 45000, 0, 0, 60000, 60000, 0, 0, '2025-10-01 22:46:03', 0, 0, '新規登録']
  ];
  
  const dataRange = sheet.getRange(2, 1, sampleData.length, sampleData[0].length);
  dataRange.setValues(sampleData);
  
  // 数値列の書式設定
  sheet.getRange(2, 1, sampleData.length, 1).setNumberFormat('0'); // 商品ID
  sheet.getRange(2, 3, sampleData.length, 1).setNumberFormat('#,##0'); // 現在仕入れ価格
  sheet.getRange(2, 4, sampleData.length, 1).setNumberFormat('#,##0'); // 前回仕入れ価格
  sheet.getRange(2, 5, sampleData.length, 1).setNumberFormat('#,##0'); // 仕入れ価格変動
  sheet.getRange(2, 6, sampleData.length, 1).setNumberFormat('0.00%'); // 仕入れ価格変動率
  sheet.getRange(2, 7, sampleData.length, 1).setNumberFormat('#,##0'); // 現在販売価格
  sheet.getRange(2, 8, sampleData.length, 1).setNumberFormat('#,##0'); // 前回販売価格
  sheet.getRange(2, 9, sampleData.length, 1).setNumberFormat('#,##0'); // 販売価格変動
  sheet.getRange(2, 10, sampleData.length, 1).setNumberFormat('0.00%'); // 販売価格変動率
  sheet.getRange(2, 12, sampleData.length, 1).setNumberFormat('0'); // 仕入れ価格変動回数
  sheet.getRange(2, 13, sampleData.length, 1).setNumberFormat('0'); // 販売価格変動回数
}

/**
 * 価格履歴の条件付き書式設定（1商品1行形式）
 */
function setupPriceHistoryConditionalFormatting(sheet) {
  // 仕入れ価格変動による色分け
  const purchasePriceChangeRange = sheet.getRange('E:E');
  
  // 価格上昇 - 赤色
  const purchasePriceUpRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#ffebee')
    .setRanges([purchasePriceChangeRange])
    .build();
  
  // 価格下落 - 緑色
  const purchasePriceDownRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#e8f5e8')
    .setRanges([purchasePriceChangeRange])
    .build();
  
  // 販売価格変動による色分け
  const sellingPriceChangeRange = sheet.getRange('I:I');
  
  // 価格上昇 - 赤色
  const sellingPriceUpRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#ffebee')
    .setRanges([sellingPriceChangeRange])
    .build();
  
  // 価格下落 - 緑色
  const sellingPriceDownRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#e8f5e8')
    .setRanges([sellingPriceChangeRange])
    .build();
  
  // 仕入れ価格変動率による色分け
  const purchaseChangeRateRange = sheet.getRange('F:F');
  
  // 大幅上昇 - 濃い赤
  const highPurchaseIncreaseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0.05)
    .setBackground('#ffcdd2')
    .setRanges([purchaseChangeRateRange])
    .build();
  
  // 大幅下落 - 濃い緑
  const highPurchaseDecreaseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(-0.05)
    .setBackground('#c8e6c9')
    .setRanges([purchaseChangeRateRange])
    .build();
  
  // 販売価格変動率による色分け
  const sellingChangeRateRange = sheet.getRange('J:J');
  
  // 大幅上昇 - 濃い赤
  const highSellingIncreaseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0.05)
    .setBackground('#ffcdd2')
    .setRanges([sellingChangeRateRange])
    .build();
  
  // 大幅下落 - 濃い緑
  const highSellingDecreaseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(-0.05)
    .setBackground('#c8e6c9')
    .setRanges([sellingChangeRateRange])
    .build();
  
  // ルールを適用
  const rules = [
    purchasePriceUpRule, purchasePriceDownRule,
    sellingPriceUpRule, sellingPriceDownRule,
    highPurchaseIncreaseRule, highPurchaseDecreaseRule,
    highSellingIncreaseRule, highSellingDecreaseRule
  ];
  sheet.setConditionalFormatRules(rules);
}

/**
 * 価格変動の検出と履歴更新
 * @param {number} productId - 商品ID
 * @param {number} newPurchasePrice - 新しい仕入れ価格
 * @param {number} newSellingPrice - 新しい販売価格
 * @param {string} notes - 備考
 */
function updatePriceHistory(productId, newPurchasePrice, newSellingPrice, notes = '') {
  // 入力パラメータの検証
  if (productId === null || productId === undefined) {
    console.error('商品IDがnullまたはundefinedです');
    return false;
  }
  
  // 商品IDが有効な数値または文字列であることを確認
  const validProductId = typeof productId === 'number' ? productId : parseInt(productId);
  if (!Number.isInteger(validProductId) || validProductId <= 0) {
    console.error('商品IDが無効です:', productId);
    return false;
  }
  
  // 仕入れ価格の検証
  if (!Number.isFinite(newPurchasePrice) || newPurchasePrice < 0) {
    console.error('仕入れ価格が無効です:', newPurchasePrice);
    return false;
  }
  
  // 販売価格の検証
  if (!Number.isFinite(newSellingPrice) || newSellingPrice < 0) {
    console.error('販売価格が無効です:', newSellingPrice);
    return false;
  }
  
  // 備考の検証とサニタイズ
  let sanitizedNotes = '';
  if (notes !== null && notes !== undefined && notes !== '') {
    sanitizedNotes = String(notes).trim();
    // 長さ制限（500文字）
    if (sanitizedNotes.length > 500) {
      sanitizedNotes = sanitizedNotes.substring(0, 500);
      console.warn('備考が長すぎるため、500文字に切り詰めました');
    }
  }
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const priceHistorySheet = spreadsheet.getSheetByName(SHEET_NAMES.PRICE_HISTORY);
  
  if (!priceHistorySheet) {
    console.error('価格履歴シートが見つかりません');
    return false;
  }
  
  // 商品IDで既存の行を検索（検証済みのvalidProductIdを使用）
  const dataRange = priceHistorySheet.getDataRange();
  const values = dataRange.getValues();
  
  let existingRowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === validProductId) {
      existingRowIndex = i + 1; // 1ベースの行番号
      break;
    }
  }
  
  const currentTime = new Date();
  const timeString = Utilities.formatDate(currentTime, 'JST', 'yyyy-MM-dd HH:mm:ss');
  
  if (existingRowIndex > 0) {
    // 既存商品の価格更新
    const currentRow = priceHistorySheet.getRange(existingRowIndex, 1, 1, 14).getValues()[0];
    
    const currentPurchasePrice = currentRow[2]; // 現在仕入れ価格
    const currentSellingPrice = currentRow[6]; // 現在販売価格
    const purchaseChangeCount = currentRow[11] || 0; // 仕入れ価格変動回数
    const sellingChangeCount = currentRow[12] || 0; // 販売価格変動回数
    
    // 価格変動の計算
    const purchasePriceChange = newPurchasePrice - currentPurchasePrice;
    const sellingPriceChange = newSellingPrice - currentSellingPrice;
    const purchaseChangeRate = currentPurchasePrice > 0 ? purchasePriceChange / currentPurchasePrice : 0;
    const sellingChangeRate = currentSellingPrice > 0 ? sellingPriceChange / currentSellingPrice : 0;
    
    // 価格変動がない場合は更新をスキップ
    if (purchasePriceChange === 0 && sellingPriceChange === 0) {
      console.log(`商品ID ${productId} の価格に変動がないため、履歴更新をスキップしました`);
      return true; // 成功として扱う（エラーではない）
    }
    
    // 変動回数の更新
    const newPurchaseChangeCount = purchasePriceChange !== 0 ? purchaseChangeCount + 1 : purchaseChangeCount;
    const newSellingChangeCount = sellingPriceChange !== 0 ? sellingChangeCount + 1 : sellingChangeCount;
    
    // 商品名を取得（在庫管理シートから）
    const productName = getProductNameById(spreadsheet, validProductId);
    
    // データを更新
    const updateData = [
      validProductId,
      productName,
      newPurchasePrice, // 現在仕入れ価格
      currentPurchasePrice, // 前回仕入れ価格
      purchasePriceChange, // 仕入れ価格変動
      purchaseChangeRate, // 仕入れ価格変動率
      newSellingPrice, // 現在販売価格
      currentSellingPrice, // 前回販売価格
      sellingPriceChange, // 販売価格変動
      sellingChangeRate, // 販売価格変動率
      timeString, // 最終更新日時
      newPurchaseChangeCount, // 仕入れ価格変動回数
      newSellingChangeCount, // 販売価格変動回数
      sanitizedNotes // サニタイズ済み備考
    ];
    
    priceHistorySheet.getRange(existingRowIndex, 1, 1, 14).setValues([updateData]);
    
    // 数値列の書式設定
    priceHistorySheet.getRange(existingRowIndex, 1, 1, 1).setNumberFormat('0'); // 商品ID
    priceHistorySheet.getRange(existingRowIndex, 3, 1, 1).setNumberFormat('#,##0'); // 現在仕入れ価格
    priceHistorySheet.getRange(existingRowIndex, 4, 1, 1).setNumberFormat('#,##0'); // 前回仕入れ価格
    priceHistorySheet.getRange(existingRowIndex, 5, 1, 1).setNumberFormat('#,##0'); // 仕入れ価格変動
    priceHistorySheet.getRange(existingRowIndex, 6, 1, 1).setNumberFormat('0.00%'); // 仕入れ価格変動率
    priceHistorySheet.getRange(existingRowIndex, 7, 1, 1).setNumberFormat('#,##0'); // 現在販売価格
    priceHistorySheet.getRange(existingRowIndex, 8, 1, 1).setNumberFormat('#,##0'); // 前回販売価格
    priceHistorySheet.getRange(existingRowIndex, 9, 1, 1).setNumberFormat('#,##0'); // 販売価格変動
    priceHistorySheet.getRange(existingRowIndex, 10, 1, 1).setNumberFormat('0.00%'); // 販売価格変動率
    priceHistorySheet.getRange(existingRowIndex, 12, 1, 1).setNumberFormat('0'); // 仕入れ価格変動回数
    priceHistorySheet.getRange(existingRowIndex, 13, 1, 1).setNumberFormat('0'); // 販売価格変動回数
    
    console.log(`商品ID ${validProductId} の価格履歴を更新しました`);
    return true;
    
  } else {
    // 新規商品の追加
    const productName = getProductNameById(spreadsheet, validProductId);
    
    const newData = [
      validProductId,
      productName,
      newPurchasePrice, // 現在仕入れ価格
      newPurchasePrice, // 前回仕入れ価格（初回は同じ）
      0, // 仕入れ価格変動（初回は0）
      0, // 仕入れ価格変動率（初回は0）
      newSellingPrice, // 現在販売価格
      newSellingPrice, // 前回販売価格（初回は同じ）
      0, // 販売価格変動（初回は0）
      0, // 販売価格変動率（初回は0）
      timeString, // 最終更新日時
      0, // 仕入れ価格変動回数（初回は0）
      0, // 販売価格変動回数（初回は0）
      sanitizedNotes || '新規登録' // サニタイズ済み備考
    ];
    
    priceHistorySheet.appendRow(newData);
    
    // 数値列の書式設定
    const newRowIndex = priceHistorySheet.getLastRow();
    priceHistorySheet.getRange(newRowIndex, 1, 1, 1).setNumberFormat('0'); // 商品ID
    priceHistorySheet.getRange(newRowIndex, 3, 1, 1).setNumberFormat('#,##0'); // 現在仕入れ価格
    priceHistorySheet.getRange(newRowIndex, 4, 1, 1).setNumberFormat('#,##0'); // 前回仕入れ価格
    priceHistorySheet.getRange(newRowIndex, 5, 1, 1).setNumberFormat('#,##0'); // 仕入れ価格変動
    priceHistorySheet.getRange(newRowIndex, 6, 1, 1).setNumberFormat('0.00%'); // 仕入れ価格変動率
    priceHistorySheet.getRange(newRowIndex, 7, 1, 1).setNumberFormat('#,##0'); // 現在販売価格
    priceHistorySheet.getRange(newRowIndex, 8, 1, 1).setNumberFormat('#,##0'); // 前回販売価格
    priceHistorySheet.getRange(newRowIndex, 9, 1, 1).setNumberFormat('#,##0'); // 販売価格変動
    priceHistorySheet.getRange(newRowIndex, 10, 1, 1).setNumberFormat('0.00%'); // 販売価格変動率
    priceHistorySheet.getRange(newRowIndex, 12, 1, 1).setNumberFormat('0'); // 仕入れ価格変動回数
    priceHistorySheet.getRange(newRowIndex, 13, 1, 1).setNumberFormat('0'); // 販売価格変動回数
    
    console.log(`商品ID ${validProductId} の新規価格履歴を追加しました`);
    return true;
  }
}

/**
 * 在庫管理シートから価格履歴を同期
 * 在庫管理シートの価格変更を検出して価格履歴を更新
 */
function syncPriceHistoryFromInventory() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  const priceHistorySheet = spreadsheet.getSheetByName(SHEET_NAMES.PRICE_HISTORY);
  
  if (!inventorySheet || !priceHistorySheet) {
    console.error('必要なシートが見つかりません');
    return false;
  }
  
  const inventoryData = inventorySheet.getDataRange().getValues();
  let updateCount = 0;
  let skipCount = 0;
  
  // 価格履歴データを一度だけ取得してルックアップテーブルを作成
  const priceHistoryData = priceHistorySheet.getDataRange().getValues();
  const priceHistoryLookup = new Map();
  
  // 価格履歴データのルックアップテーブルを作成（O(m)）
  for (let j = 1; j < priceHistoryData.length; j++) {
    const productId = priceHistoryData[j][0];
    if (productId) {
      priceHistoryLookup.set(productId, {
        currentPurchasePrice: priceHistoryData[j][2], // 現在仕入れ価格
        currentSellingPrice: priceHistoryData[j][6]   // 現在販売価格
      });
    }
  }
  
  // ヘッダー行をスキップして処理（O(n)）
  for (let i = 1; i < inventoryData.length; i++) {
    const productId = inventoryData[i][0];
    const productName = inventoryData[i][1];
    const purchasePrice = inventoryData[i][6]; // 仕入れ価格
    const sellingPrice = inventoryData[i][7]; // 販売価格
    
    if (productId && productName && purchasePrice && sellingPrice) {
      // ルックアップテーブルから現在の価格を取得（O(1)）
      const existingRecord = priceHistoryLookup.get(productId);
      
      // 既存レコードがあり、価格に変動がない場合はスキップ
      if (existingRecord && 
          purchasePrice === existingRecord.currentPurchasePrice && 
          sellingPrice === existingRecord.currentSellingPrice) {
        skipCount++;
        continue;
      }
      
      const success = updatePriceHistory(productId, purchasePrice, sellingPrice, '在庫管理シートから同期');
      if (success) {
        updateCount++;
      }
    }
  }
  
  console.log(`${updateCount}件の価格履歴を更新、${skipCount}件をスキップしました`);
  return true;
}
