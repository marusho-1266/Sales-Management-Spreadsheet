/**
 * 売上管理シートの作成と管理
 * 在庫管理ツール システム開発
 * 作成日: 2025-09-28
 */

/**
 * 売上管理シートの初期化（Joom注文連携対応版）
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
  
  // 基本情報フィールド（1-12列）
  const basicHeaders = [
    '注文ID',           // A列: Joom注文ID（8文字英数字）
    '注文日',           // B列: 注文受付日
    '商品ID',           // C列: Joom製品SKU（在庫管理の商品IDと同一）
    '商品名',           // D列: 商品名称（在庫管理から取得）
    'SKU',              // E列: 商品管理コード（在庫管理から取得）
    'ASIN',             // F列: Amazon商品コード（在庫管理から取得）
    '数量',             // G列: 注文数量
    '販売価格',         // H列: 実際の販売価格（円）
    '仕入れ価格',       // I列: 仕入れ価格（円、在庫管理から取得）
    '送料',             // J列: 配送料（円）
    '純利益',           // K列: 手数料差し引き後利益（自動計算）
    '登録日時'          // L列: データ登録日時
  ];
  
  // 注文ステータス管理フィールド（13-15列）
  const statusHeaders = [
    '注文ステータス',   // M列: Joom注文ステータス（approved/shipped等）
    'Joom連携ステータス', // N列: 連携状況（synced/error等）
    '最終同期日時'      // O列: 最後に同期した日時
  ];
  
  // 価格詳細管理フィールド（16-19列）
  const priceHeaders = [
    '手数料',           // P列: Joom手数料（円）
    '付加価値税',       // Q列: VAT金額（円）
    '返金額',           // R列: 返金された金額（円）
    '購入者GMV'         // S列: 購入者の総商品価値（円）
  ];
  
  // 顧客情報フィールド（20-24列）
  const customerHeaders = [
    '顧客名',           // T列: 注文者の氏名
    '顧客メール',       // U列: 注文者のメールアドレス
    '顧客電話番号',     // V列: 注文者の電話番号
    '顧客国コード',     // W列: 顧客の国コード
    '顧客都道府県'      // X列: 顧客の都道府県
  ];
  
  // 配送情報フィールド（25-30列）
  const shippingHeaders = [
    '配送国',           // Y列: 配送先の国コード
    '配送都道府県',     // Z列: 配送先の都道府県
    '配送市区町村',     // AA列: 配送先の市区町村
    '配送住所',         // AB列: 配送先の住所
    '配送郵便番号',     // AC列: 配送先の郵便番号
    '配送先住所（完全版）' // AD列: 配送先住所の完全版文字列
  ];
  
  // 出荷・配送管理フィールド（31-35列）
  const deliveryHeaders = [
    '追跡番号',         // AE列: 配送追跡番号
    '配送業者',         // AF列: 配送を担当する業者
    '出荷日時',         // AG列: 実際の出荷日時
    '履行日時',         // AH列: 注文履行完了日時
    '配送ステータス'    // AI列: 配送の現在状況
  ];
  
  // 連携管理フィールド（36-37列）
  const syncHeaders = [
    '同期エラーメッセージ', // AJ列: 同期時のエラーメッセージ
    'データソース'       // AK列: データの取得元（Joom/Manual）
  ];
  
  // 全ヘッダーを結合
  const allHeaders = [
    ...basicHeaders,
    ...statusHeaders,
    ...priceHeaders,
    ...customerHeaders,
    ...shippingHeaders,
    ...deliveryHeaders,
    ...syncHeaders
  ];
  
  // ヘッダーを設定
  salesSheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
  
  // ヘッダーの書式設定
  const headerRange = salesSheet.getRange(1, 1, 1, allHeaders.length);
  headerRange.setBackground('#1e7e34');  // より濃い緑色に変更
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // セクション別の色分け
  const sectionColors = [
    { start: 1, end: 12, color: '#c8e6c9' },    // 基本情報（より濃い緑系）
    { start: 13, end: 15, color: '#ffe0b2' },   // 注文ステータス（より濃いオレンジ系）
    { start: 16, end: 19, color: '#bbdefb' },   // 価格詳細（より濃い青系）
    { start: 20, end: 24, color: '#e1bee7' },   // 顧客情報（より濃い紫系）
    { start: 25, end: 30, color: '#b2dfdb' },   // 配送情報（より濃いティール系）
    { start: 31, end: 35, color: '#f8bbd9' },   // 出荷・配送（より濃いピンク系）
    { start: 36, end: 37, color: '#dcedc8' }    // 連携管理（より濃いライトグリーン系）
  ];
  
  sectionColors.forEach(section => {
    const range = salesSheet.getRange(1, section.start, 1, section.end - section.start + 1);
    range.setBackground(section.color);
  });
  
  // 列幅の設定
  const columnWidths = [
    120, 100, 120, 200, 100, 120, 60, 100, 100, 80, 100, 150,  // 基本情報（1-12列）
    120, 120, 150,  // 注文ステータス（13-15列）
    80, 80, 80, 100,  // 価格詳細（16-19列）
    120, 200, 120, 80, 100,  // 顧客情報（20-24列）
    80, 100, 120, 200, 100, 200,  // 配送情報（25-30列）
    120, 100, 120, 120, 120,  // 出荷・配送（31-35列）
    200, 100  // 連携管理（36-37列）
  ];
  
  columnWidths.forEach((width, index) => {
    salesSheet.setColumnWidth(index + 1, width);
  });
  
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
    [1, '2025-09-25', 1, 'iPhone 15 Pro 128GB', 'IPH15P-128', 'B0CHX1W1XY', 1, 150000, 120000, 0, 30000, '2025-09-25 10:30:00'],
    [2, '2025-09-26', 2, 'MacBook Air M2 13インチ', 'MBA-M2-13', 'B0B3C2Q5XK', 1, 180000, 140000, 0, 40000, '2025-09-26 14:15:00'],
    [3, '2025-09-27', 4, 'iPad Air 第5世代', 'IPAD-AIR-5', 'B09V4HCN9V', 2, 80000, 60000, 0, 20000, '2025-09-27 09:45:00']
  ];
  
  const dataRange = sheet.getRange(2, 1, sampleData.length, sampleData[0].length);
  dataRange.setValues(sampleData);
  
  // 数値列の書式設定
  sheet.getRange(2, 1, sampleData.length, 1).setNumberFormat('0'); // 注文ID
  sheet.getRange(2, 3, sampleData.length, 1).setNumberFormat('0'); // 商品ID
  sheet.getRange(2, 7, sampleData.length, 1).setNumberFormat('0'); // 数量
  sheet.getRange(2, 8, sampleData.length, 1).setNumberFormat('#,##0'); // 販売価格
  sheet.getRange(2, 9, sampleData.length, 1).setNumberFormat('#,##0'); // 仕入れ価格
  sheet.getRange(2, 10, sampleData.length, 1).setNumberFormat('#,##0'); // 送料
  sheet.getRange(2, 11, sampleData.length, 1).setNumberFormat('#,##0'); // 純利益
}

/**
 * 売上管理シートの条件付き書式設定
 */
function setupSalesConditionalFormatting(sheet) {
  // 純利益による色分け
  const profitRange = sheet.getRange('K:K');
  
  // 純利益が正 - 緑色
  const positiveProfitRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#d9ead3')
    .setRanges([profitRange])
    .build();
  
  // 純利益が負 - 赤色
  const negativeProfitRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#f4cccc')
    .setRanges([profitRange])
    .build();
  
  // ルールを適用
  const rules = [positiveProfitRule, negativeProfitRule];
  sheet.setConditionalFormatRules(rules);
}

/**
 * 売上管理シートの利益計算式設定
 */
function setupSalesProfitCalculation(sheet) {
  const lastRow = sheet.getLastRow();
  
  // 純利益計算式を設定（販売価格 - 仕入れ価格）
  for (let row = 2; row <= lastRow; row++) {
    const formula = `=H${row}-I${row}`;
    sheet.getRange(row, 11).setFormula(formula);
  }
}

/**
 * 注文データ入力フォームの表示
 */
function showSalesInputForm() {
  const htmlTemplate = HtmlService.createTemplate(getSalesInputFormHtml());
  htmlTemplate.nextOrderId = getNextOrderId();
  htmlTemplate.inventoryData = getInventoryDataForSelection();
  const htmlOutput = htmlTemplate.evaluate()
    .setWidth(800)
    .setHeight(600)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '注文データ入力');
}

/**
 * 次の注文IDを取得（自然数1から順に付番）
 */
function getNextOrderId() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = spreadsheet.getSheetByName('売上管理');
  
  if (!salesSheet) {
    return 1;
  }
  
  const data = salesSheet.getDataRange().getValues();
  let maxOrderId = 0;
  
  // ヘッダー行をスキップして注文IDをチェック
  for (let i = 1; i < data.length; i++) {
    const orderId = data[i][0];
    if (orderId && !isNaN(orderId) && orderId > 0) {
      if (orderId > maxOrderId) {
        maxOrderId = orderId;
      }
    }
  }
  
  return maxOrderId + 1;
}

/**
 * 商品選択用の在庫データを取得
 */
function getInventoryDataForSelection() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = spreadsheet.getSheetByName('在庫管理');
  
  if (!inventorySheet) {
    return [];
  }
  
  const data = inventorySheet.getDataRange().getValues();
  const products = [];
  
  // ヘッダー行をスキップして商品データを取得
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] && row[1] && row[COLUMN_INDEXES.INVENTORY.STOCK_STATUS - 1] === '在庫あり') { // 商品ID、商品名が存在し、在庫ステータスが「在庫あり」の場合
      products.push({
        id: row[0],
        name: row[1],
        sku: row[COLUMN_INDEXES.INVENTORY.SKU - 1],
        asin: row[COLUMN_INDEXES.INVENTORY.ASIN - 1],
        stockStatus: row[COLUMN_INDEXES.INVENTORY.STOCK_STATUS - 1],
        purchasePrice: row[COLUMN_INDEXES.INVENTORY.PURCHASE_PRICE - 1],
        sellingPrice: row[COLUMN_INDEXES.INVENTORY.SELLING_PRICE - 1]
      });
    }
  }
  
  return products;
}

/**
 * 売上データを追加（HTMLフォーム用）
 */
function saveSalesData(salesData) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = spreadsheet.getSheetByName('売上管理');
    
    if (!salesSheet) {
      return { success: false, message: '売上管理シートが見つかりません' };
    }
    
    // 現在の日時を取得
    const now = new Date();
    const orderDate = salesData.orderDate || Utilities.formatDate(now, 'JST', 'yyyy-MM-dd');
    const registrationTime = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd HH:mm:ss');
    
    // 売上データを追加（K列は後で計算式を設定するため0を仮置き）
    const newRow = [
      salesData.orderId,
      orderDate,
      salesData.productId,
      salesData.productName,
      salesData.sku,
      salesData.asin,
      salesData.quantity,
      salesData.sellingPrice,
      salesData.purchasePrice,
      salesData.shippingCost,
      0,
      registrationTime
    ];
    
    // データを追加
    const lastRow = salesSheet.getLastRow();
    const newRowIndex = lastRow + 1;
    salesSheet.getRange(newRowIndex, 1, 1, newRow.length).setValues([newRow]);
    
    // K列に純利益の計算式を設定: =H-(I+J+P+Q+R)
    const profitFormula = '=H' + newRowIndex + '-(I' + newRowIndex + '+J' + newRowIndex + '+P' + newRowIndex + '+Q' + newRowIndex + '+R' + newRowIndex + ')';
    salesSheet.getRange(newRowIndex, COLUMN_INDEXES.SALES.NET_PROFIT).setFormula(profitFormula);
    
    // P列(手数料)・Q列(VAT)・R列(返金額)を書き込み
    const commission = salesData.commission || 0;
    const vat = salesData.vat || 0;
    const refundAmount = salesData.refundAmount || 0;
    if (commission > 0) {
      salesSheet.getRange(newRowIndex, COLUMN_INDEXES.SALES.COMMISSION).setValue(commission);
    }
    if (vat > 0) {
      salesSheet.getRange(newRowIndex, COLUMN_INDEXES.SALES.VAT).setValue(vat);
    }
    if (refundAmount > 0) {
      salesSheet.getRange(newRowIndex, COLUMN_INDEXES.SALES.REFUND_AMOUNT).setValue(refundAmount);
    }
    
    // 在庫管理シートの在庫数量を減算
    const inventoryMessage = deductInventoryStock(spreadsheet, salesData.productId, salesData.quantity);
    
     console.log('売上データが正常に追加されました:', salesData.orderId);
     const resultMessage = '注文データが正常に保存されました' + (inventoryMessage ? '\n' + inventoryMessage : '');
     return { 
       success: true, 
       message: resultMessage,
       orderId: salesData.orderId,
       productName: salesData.productName
     };
    
  } catch (error) {
    console.error('売上データの追加中にエラーが発生しました:', error);
    return { success: false, message: 'データの保存中にエラーが発生しました: ' + error.message };
  }
}

/**
 * 在庫管理シートの在庫数量を減算する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {number|string} productId - 商品ID
 * @param {number} quantity - 減算する数量
 * @returns {string} 結果メッセージ（空文字は変更なし）
 */
function deductInventoryStock(spreadsheet, productId, quantity) {
  try {
    const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
    if (!inventorySheet) {
      console.log('在庫管理シートが見つからないため在庫減算をスキップしました');
      return '';
    }

    const data = inventorySheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(productId).trim()) {
        const row = i + 1;
        const currentStock = parseInt(data[i][COLUMN_INDEXES.INVENTORY.STOCK_QUANTITY - 1]) || 0;
        const newStock = Math.max(0, currentStock - quantity);
        inventorySheet.getRange(row, COLUMN_INDEXES.INVENTORY.STOCK_QUANTITY).setValue(newStock);

        if (newStock === 0) {
          inventorySheet.getRange(row, COLUMN_INDEXES.INVENTORY.STOCK_STATUS).setValue(STOCK_STATUS.OUT_OF_STOCK);
        }

        console.log('在庫減算: 商品ID=' + productId + ' ' + currentStock + ' → ' + newStock);
        return '在庫数量: ' + currentStock + ' → ' + newStock;
      }
    }
    console.log('在庫減算: 商品ID ' + productId + ' が在庫管理シートに見つかりません');
    return '';
  } catch (error) {
    console.error('在庫減算中にエラーが発生しました:', error);
    return '';
  }
}

/**
 * 売上データを追加（従来の関数）
 */
function addSalesData(salesData) {
  const result = saveSalesData(salesData);
  return result.success;
}



/**
 * 売上データの削除
 */
function deleteSalesData(orderId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = spreadsheet.getSheetByName('売上管理');
    
    if (!salesSheet) {
      console.error('売上管理シートが見つかりません');
      return false;
    }
    
    const data = salesSheet.getDataRange().getValues();
    
    // 注文IDで検索（数値として比較）
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === parseInt(orderId)) {
        // 行を削除
        salesSheet.deleteRow(i + 1);
        console.log(`注文ID ${orderId} のデータを削除しました`);
        return true;
      }
    }
    
    console.error(`注文ID ${orderId} が見つかりません`);
    return false;
    
  } catch (error) {
    console.error('売上データの削除中にエラーが発生しました:', error);
    return false;
  }
}


/**
 * 売上入力フォームのHTMLテンプレート
 */
function getSalesInputFormHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>注文データ入力</title>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 0;
      padding: 20px;
      background-color: #f8f9fa;
      color: #333;
    }
    
    .container {
      max-width: 100%;
      margin: 0 auto;
      background: white;
      border-radius: 8px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      overflow: hidden;
    }
    
    .header {
      background: linear-gradient(135deg, #34a853, #2e7d32);
      color: white;
      padding: 20px;
      text-align: center;
    }
    
    .header h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 600;
    }
    
    .form-container {
      padding: 30px;
    }
    
    .form-group {
      margin-bottom: 20px;
    }
    
    .form-group label {
      display: block;
      margin-bottom: 8px;
      font-weight: 600;
      color: #555;
    }
    
    .form-group input,
    .form-group select {
      width: 100%;
      padding: 12px;
      border: 2px solid #e1e5e9;
      border-radius: 6px;
      font-size: 14px;
      transition: border-color 0.3s ease;
      box-sizing: border-box;
    }
    
    .form-group input:focus,
    .form-group select:focus {
      outline: none;
      border-color: #34a853;
      box-shadow: 0 0 0 3px rgba(52, 168, 83, 0.1);
    }
    
    .form-group input[readonly] {
      background-color: #f8f9fa;
      color: #6c757d;
      cursor: not-allowed;
    }
    
    .form-row {
      display: flex;
      gap: 20px;
    }
    
    .form-row .form-group {
      flex: 1;
    }
    
    .product-info {
      background-color: #f8f9fa;
      padding: 15px;
      border-radius: 6px;
      margin-top: 10px;
      border-left: 4px solid #34a853;
    }
    
    .product-info h4 {
      margin: 0 0 10px 0;
      color: #34a853;
    }
    
    .product-info p {
      margin: 5px 0;
      color: #666;
    }
    
    .button-group {
      display: flex;
      gap: 15px;
      justify-content: center;
      margin-top: 30px;
    }
    
    .btn {
      padding: 12px 30px;
      border: none;
      border-radius: 6px;
      font-size: 16px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.3s ease;
      min-width: 120px;
    }
    
    .btn-primary {
      background: linear-gradient(135deg, #34a853, #2e7d32);
      color: white;
    }
    
    .btn-primary:hover {
      background: linear-gradient(135deg, #2e7d32, #1b5e20);
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(52, 168, 83, 0.3);
    }
    
    .btn-secondary {
      background: #6c757d;
      color: white;
    }
    
    .btn-secondary:hover {
      background: #5a6268;
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(108, 117, 125, 0.3);
    }
    
    .error-message {
      color: #dc3545;
      font-size: 14px;
      margin-top: 5px;
      display: none;
    }
    
    .success-message {
      color: #28a745;
      font-size: 14px;
      margin-top: 5px;
      display: none;
    }
    
     .loading-overlay {
       position: fixed;
       top: 0;
       left: 0;
       width: 100%;
       height: 100%;
       background: rgba(0, 0, 0, 0.8);
       display: none;
       justify-content: center;
       align-items: center;
       z-index: 1000;
       flex-direction: column;
     }
     
     .loading-spinner {
       width: 60px;
       height: 60px;
       border: 4px solid #f3f3f3;
       border-top: 4px solid #34a853;
       border-radius: 50%;
       animation: spin 1s linear infinite;
       margin-bottom: 20px;
     }
     
     .loading-message {
       color: white;
       font-size: 18px;
       font-weight: 600;
       margin-bottom: 10px;
       text-align: center;
     }
     
     .loading-details {
       color: #ccc;
       font-size: 14px;
       margin-bottom: 20px;
       text-align: center;
     }
     
     .progress-bar {
       width: 300px;
       height: 6px;
       background: rgba(255, 255, 255, 0.2);
       border-radius: 3px;
       overflow: hidden;
       margin-top: 15px;
     }
     
     .progress-fill {
       height: 100%;
       background: linear-gradient(90deg, #4285f4, #34a853);
       width: 0%;
       transition: width 0.3s ease;
     }
     
     .btn:disabled {
       opacity: 0.6;
       cursor: not-allowed;
       transform: none !important;
     }
     
     .price-preview {
       background: #e3f2fd;
       padding: 15px;
       border-radius: 6px;
       margin-top: 10px;
       border-left: 4px solid #4285f4;
     }
     
     .price-preview h4 {
       margin: 0 0 10px 0;
       color: #1565c0;
       font-size: 14px;
     }
     
     .price-item {
       display: flex;
       justify-content: space-between;
       margin-bottom: 5px;
     }
     
     .price-item.total {
       font-weight: 600;
       border-top: 1px solid #bbdefb;
       padding-top: 8px;
       margin-top: 8px;
     }
     
     @keyframes spin {
       0% { transform: rotate(0deg); }
       100% { transform: rotate(360deg); }
     }
     
     /* 商品選択モーダルスタイル */
     .product-selection-container {
       display: flex;
       gap: 10px;
       align-items: center;
     }
     
     .product-selection-container input[readonly] {
       flex: 1;
       background-color: #f8f9fa;
       color: #6c757d;
       cursor: pointer;
     }
     
     .btn-select-product {
       padding: 12px 20px;
       background: #34a853;
       color: white;
       border: none;
       border-radius: 6px;
       cursor: pointer;
       font-size: 14px;
       font-weight: 600;
       transition: background-color 0.3s ease;
     }
     
     .btn-select-product:hover {
       background: #2e7d32;
     }
     
     /* モーダルダイアログ */
     .modal {
       display: none;
       position: fixed;
       z-index: 2000;
       left: 0;
       top: 0;
       width: 100%;
       height: 100%;
       background-color: rgba(0, 0, 0, 0.5);
     }
     
     .modal-content {
       background-color: #fefefe;
       margin: 5% auto;
       padding: 0;
       border-radius: 8px;
       width: 90%;
       max-width: 1000px;
       max-height: 80vh;
       overflow: hidden;
       box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
     }
     
     .modal-header {
       background: #34a853;
       color: white;
       padding: 20px;
       display: flex;
       justify-content: space-between;
       align-items: center;
     }
     
     .modal-header h3 {
       margin: 0;
       font-size: 18px;
     }
     
     .close {
       color: white;
       font-size: 28px;
       font-weight: bold;
       cursor: pointer;
       line-height: 1;
     }
     
     .close:hover {
       opacity: 0.7;
     }
     
     .modal-body {
       padding: 20px;
       max-height: 60vh;
       overflow-y: auto;
     }
     
     .search-container {
       margin-bottom: 20px;
     }
     
     .search-container input {
       width: 100%;
       padding: 12px;
       border: 2px solid #e1e5e9;
       border-radius: 6px;
       font-size: 14px;
     }
     
     .search-container input:focus {
       outline: none;
       border-color: #34a853;
       box-shadow: 0 0 0 3px rgba(52, 168, 83, 0.1);
     }
     
     .product-list-container {
       max-height: 400px;
       overflow-y: auto;
       border: 1px solid #e1e5e9;
       border-radius: 6px;
     }
     
     .product-table {
       width: 100%;
       border-collapse: collapse;
       margin: 0;
     }
     
     .product-table th {
       background: #f8f9fa;
       padding: 12px 8px;
       text-align: left;
       font-weight: 600;
       border-bottom: 2px solid #e1e5e9;
       position: sticky;
       top: 0;
       z-index: 10;
     }
     
     .product-table td {
       padding: 12px 8px;
       border-bottom: 1px solid #e1e5e9;
       vertical-align: middle;
     }
     
     .product-table tr:hover {
       background-color: #f8f9fa;
     }
     
     .product-table tr.hidden {
       display: none;
     }
     
     .btn-select {
       padding: 6px 12px;
       background: #34a853;
       color: white;
       border: none;
       border-radius: 4px;
       cursor: pointer;
       font-size: 12px;
       font-weight: 600;
       transition: background-color 0.3s ease;
     }
     
     .btn-select:hover {
       background: #2e7d32;
     }
  </style>
</head>
<body>
  <!-- ローディングオーバーレイ -->
  <div class="loading-overlay" id="loadingOverlay">
    <div class="loading-spinner"></div>
    <div class="loading-message" id="loadingMessage">注文データを保存中...</div>
    <div class="loading-details" id="loadingDetails">しばらくお待ちください</div>
    <div class="progress-bar">
      <div class="progress-fill" id="progressFill"></div>
    </div>
  </div>

  <div class="container">
    <div class="header">
      <h1>📈 注文データ入力</h1>
    </div>
    
    <div class="form-container">
      <div class="success-message" id="successMessage">
        ✅ 注文データが正常に保存されました！
      </div>
      
      <form id="salesForm">
        <div class="form-group">
          <label for="orderId">注文ID *</label>
          <input type="number" id="orderId" name="orderId" value="<?= nextOrderId ?>" min="1" required readonly>
        </div>
        
        <div class="form-group">
          <label for="orderDate">注文日 *</label>
          <input type="date" id="orderDate" name="orderDate" required>
        </div>
        
        <div class="form-group">
          <label for="productId">商品選択 *</label>
          <div class="product-selection-container">
            <input type="hidden" id="productId" name="productId" required>
            <input type="text" id="productDisplay" placeholder="商品を選択してください" readonly onclick="openProductModal()">
            <button type="button" class="btn-select-product" onclick="openProductModal()">選択</button>
          </div>
        </div>
        
        <div class="form-row">
          <div class="form-group">
            <label for="quantity">数量 *</label>
            <input type="number" id="quantity" name="quantity" min="1" value="1" required onchange="validateQuantity()">
            <div id="quantityError" class="error-message"></div>
          </div>
          
          <div class="form-group">
            <label for="sellingPrice">販売価格 (円) *</label>
            <input type="number" id="sellingPrice" name="sellingPrice" min="1" required onchange="calculateProfit()">
          </div>
        </div>
        
         <div class="form-row">
           <div class="form-group">
             <label for="purchasePrice">仕入れ価格 (円) *</label>
             <input type="number" id="purchasePrice" name="purchasePrice" min="0" required onchange="calculateProfit()">
           </div>
           
           <div class="form-group">
             <label for="shippingCost">送料 (円)</label>
             <input type="number" id="shippingCost" name="shippingCost" min="0" value="0" onchange="calculateProfit()">
           </div>
         </div>
         
         <div class="form-row">
           <div class="form-group">
             <label for="commission">手数料 (円)</label>
             <input type="number" id="commission" name="commission" min="0" value="0" onchange="calculateProfit()">
           </div>
           
           <div class="form-group">
             <label for="vat">付加価値税 (円)</label>
             <input type="number" id="vat" name="vat" min="0" value="0" onchange="calculateProfit()">
           </div>
         </div>
         
         <div class="form-row">
           <div class="form-group">
             <label for="refundAmount">返金額 (円)</label>
             <input type="number" id="refundAmount" name="refundAmount" min="0" value="0" onchange="calculateProfit()">
           </div>
         </div>
         
         <!-- 利益計算プレビュー -->
         <div class="price-preview" id="pricePreview" style="display: none;">
           <h4>💰 利益計算プレビュー</h4>
           <div class="price-item">
             <span>販売価格:</span>
             <span id="previewSellingPrice">¥0</span>
           </div>
           <div class="price-item">
             <span>仕入れ価格:</span>
             <span id="previewPurchasePrice">¥0</span>
           </div>
           <div class="price-item">
             <span>送料:</span>
             <span id="previewShippingCost">¥0</span>
           </div>
           <div class="price-item">
             <span>手数料:</span>
             <span id="previewCommission">¥0</span>
           </div>
           <div class="price-item">
             <span>付加価値税:</span>
             <span id="previewVat">¥0</span>
           </div>
           <div class="price-item">
             <span>返金額:</span>
             <span id="previewRefundAmount">¥0</span>
           </div>
           <div class="price-item total">
             <span><strong>純利益:</strong></span>
             <span id="previewNetProfit"><strong>¥0</strong></span>
           </div>
         </div>
         
         <div class="form-group">
           <label for="netProfit">純利益 (円)</label>
           <input type="number" id="netProfit" name="netProfit" readonly>
         </div>
        
        <div id="errorMessage" class="error-message"></div>
        
        <div class="button-group">
          <button type="button" class="btn btn-secondary" onclick="closeDialog()">キャンセル</button>
          <button type="submit" class="btn btn-primary">保存</button>
        </div>
      </form>
    </div>
  </div>

  <!-- 商品選択モーダル -->
  <div id="productModal" class="modal">
    <div class="modal-content">
      <div class="modal-header">
        <h3>商品選択</h3>
        <span class="close" onclick="closeProductModal()">&times;</span>
      </div>
      <div class="modal-body">
        <div class="search-container">
          <input type="text" id="productSearch" placeholder="商品名、SKU、ASINで検索..." onkeyup="filterProducts()">
        </div>
        <div class="product-list-container">
          <table id="productTable" class="product-table">
            <thead>
              <tr>
                <th>商品名</th>
                <th>SKU</th>
                <th>ASIN</th>
                <th>在庫</th>
                <th>仕入れ価格</th>
                <th>販売価格</th>
                <th>選択</th>
              </tr>
            </thead>
            <tbody id="productTableBody">
              <? for (var i = 0; i < inventoryData.length; i++) { ?>
                <tr class="product-row" 
                    data-id="<?= inventoryData[i].id ?>"
                    data-name="<?= inventoryData[i].name ?>"
                    data-sku="<?= inventoryData[i].sku ?>"
                    data-asin="<?= inventoryData[i].asin ?>"
                    data-stock-status="<?= inventoryData[i].stockStatus ?>"
                    data-purchase-price="<?= inventoryData[i].purchasePrice ?>"
                    data-selling-price="<?= inventoryData[i].sellingPrice ?>">
                  <td><?= inventoryData[i].name ?></td>
                  <td><?= inventoryData[i].sku ?></td>
                  <td><?= inventoryData[i].asin ?></td>
                  <td><?= inventoryData[i].stockStatus ?></td>
                  <td>¥<?= parseInt(inventoryData[i].purchasePrice).toLocaleString() ?></td>
                  <td>¥<?= parseInt(inventoryData[i].sellingPrice).toLocaleString() ?></td>
                  <td><button type="button" class="btn-select" onclick="selectProduct(this)">選択</button></td>
                </tr>
              <? } ?>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

   <script>
     // グローバル変数
     var isProcessing = false;
     var currentStep = 0;
     var progressSteps = [
       '入力データを検証しています...',
       'スプレッドシートに保存中...',
       '書式設定中...',
       '保存完了'
     ];
     
     // 現在の日付を設定
     document.getElementById('orderDate').value = new Date().toISOString().split('T')[0];
     
     // ローディング表示の制御
     function showLoadingOverlay() {
       isProcessing = true;
       document.getElementById('loadingOverlay').style.display = 'flex';
       
       // ボタンを無効化
       var buttons = document.querySelectorAll('.btn');
       buttons.forEach(function(btn) {
         btn.disabled = true;
       });
     }
     
     function hideLoadingOverlay() {
       isProcessing = false;
       document.getElementById('loadingOverlay').style.display = 'none';
       
       // ボタンを有効化
       var buttons = document.querySelectorAll('.btn');
       buttons.forEach(function(btn) {
         btn.disabled = false;
       });
     }
     
     function updateProgress(step, details) {
       currentStep = step;
       var progress = ((step + 1) / progressSteps.length) * 100;
       
       document.getElementById('loadingMessage').textContent = progressSteps[step];
       document.getElementById('loadingDetails').textContent = details || '処理中...';
       document.getElementById('progressFill').style.width = progress + '%';
     }
     
    
    // 数量の検証（在庫数管理がないため、基本的な検証のみ）
    function validateQuantity() {
      const quantity = parseInt(document.getElementById('quantity').value);
      const errorDiv = document.getElementById('quantityError');
      
      if (quantity <= 0) {
        errorDiv.textContent = '数量は1以上で入力してください。';
        errorDiv.style.display = 'block';
        return false;
      } else {
        errorDiv.style.display = 'none';
        return true;
      }
    }
    
     // 利益計算: K列の計算式 =H-(I+J+P+Q+R) と一致させる
     function calculateProfit() {
       const sellingPrice = parseFloat(document.getElementById('sellingPrice').value) || 0;
       const purchasePrice = parseFloat(document.getElementById('purchasePrice').value) || 0;
       const shippingCost = parseFloat(document.getElementById('shippingCost').value) || 0;
       const commission = parseFloat(document.getElementById('commission').value) || 0;
       const vat = parseFloat(document.getElementById('vat').value) || 0;
       const refundAmount = parseFloat(document.getElementById('refundAmount').value) || 0;
       const netProfit = sellingPrice - (purchasePrice + shippingCost + commission + vat + refundAmount);
       
       document.getElementById('netProfit').value = Math.round(netProfit);
       
       updatePricePreview(sellingPrice, purchasePrice, shippingCost, commission, vat, refundAmount, netProfit);
     }
     
     function updatePricePreview(sellingPrice, purchasePrice, shippingCost, commission, vat, refundAmount, netProfit) {
       const pricePreview = document.getElementById('pricePreview');
       
       if (sellingPrice > 0) {
         document.getElementById('previewSellingPrice').textContent = '¥' + sellingPrice.toLocaleString();
         document.getElementById('previewPurchasePrice').textContent = '¥' + purchasePrice.toLocaleString();
         document.getElementById('previewShippingCost').textContent = '¥' + shippingCost.toLocaleString();
         document.getElementById('previewCommission').textContent = '¥' + commission.toLocaleString();
         document.getElementById('previewVat').textContent = '¥' + vat.toLocaleString();
         document.getElementById('previewRefundAmount').textContent = '¥' + refundAmount.toLocaleString();
         document.getElementById('previewNetProfit').textContent = '¥' + Math.round(netProfit).toLocaleString();
         
         // 利益の色分け
         const netProfitElement = document.getElementById('previewNetProfit');
         if (netProfit > 0) {
           netProfitElement.style.color = '#2e7d32';
         } else if (netProfit < 0) {
           netProfitElement.style.color = '#d32f2f';
         } else {
           netProfitElement.style.color = '#666';
         }
         
         pricePreview.style.display = 'block';
       } else {
         pricePreview.style.display = 'none';
       }
     }
    
     // フォーム送信処理
     document.getElementById('salesForm').addEventListener('submit', function(e) {
       e.preventDefault();
       
       // 重複送信防止
       if (isProcessing) {
         console.log('処理中です。しばらくお待ちください。');
         return;
       }
       
       // エラーメッセージをクリア
       document.getElementById('errorMessage').style.display = 'none';
       document.getElementById('successMessage').style.display = 'none';
       
       // 数量の検証
       if (!validateQuantity()) {
         return;
       }
       
       // ローディング表示開始
       showLoadingOverlay();
       
       // プログレス更新（データ検証）
       updateProgress(0, '入力データを検証しています...');
       
       // 商品選択の検証
       const productId = document.getElementById('productId').value;
       if (!productId || !window.selectedProductData) {
         hideLoadingOverlay();
         document.getElementById('errorMessage').textContent = '商品を選択してください。';
         document.getElementById('errorMessage').style.display = 'block';
         return;
       }
       
       // フォームデータを取得
       const formData = {
         orderId: parseInt(document.getElementById('orderId').value),
         orderDate: document.getElementById('orderDate').value,
         productId: parseInt(productId),
         sku: window.selectedProductData.sku,
         asin: window.selectedProductData.asin,
         productName: window.selectedProductData.name,
         quantity: parseInt(document.getElementById('quantity').value),
        sellingPrice: parseInt(document.getElementById('sellingPrice').value),
        purchasePrice: parseInt(document.getElementById('purchasePrice').value),
        shippingCost: parseInt(document.getElementById('shippingCost').value) || 0,
        commission: parseInt(document.getElementById('commission').value) || 0,
        vat: parseInt(document.getElementById('vat').value) || 0,
        refundAmount: parseInt(document.getElementById('refundAmount').value) || 0
      };
       
       // プログレス更新（保存開始）
       setTimeout(function() {
         updateProgress(1, 'スプレッドシートに保存中...');
       }, 500);
       
       // データを保存
       google.script.run
         .withSuccessHandler(onSaveSuccess)
         .withFailureHandler(onSaveError)
         .saveSalesData(formData);
     });
     
     // 保存成功時の処理
     function onSaveSuccess(result) {
       if (result.success) {
         // プログレス更新（書式設定）
         updateProgress(2, '書式設定中...');
         
         setTimeout(function() {
           // プログレス更新（完了）
           updateProgress(3, '保存完了');
           
           setTimeout(function() {
             // ローディング表示を非表示
             hideLoadingOverlay();
             
             // 成功メッセージを表示
             var successMessage = document.getElementById('successMessage');
             successMessage.innerHTML = '✅ ' + result.message + '<br>注文ID: ' + result.orderId + '<br>商品名: ' + result.productName;
             successMessage.style.display = 'block';
             successMessage.style.background = '#d9ead3';
             successMessage.style.color = '#2d5016';
             
             // フォームをクリア
             document.getElementById('salesForm').reset();
             document.getElementById('pricePreview').style.display = 'none';
             
             // 3秒後にフォームを閉じる
             setTimeout(function() {
               closeDialog();
             }, 3000);
           }, 800);
         }, 500);
       } else {
         // エラー処理
         hideLoadingOverlay();
         document.getElementById('errorMessage').textContent = result.message || 'データの保存に失敗しました。';
         document.getElementById('errorMessage').style.display = 'block';
       }
     }
     
     // 保存エラー時の処理
     function onSaveError(error) {
       hideLoadingOverlay();
       document.getElementById('errorMessage').textContent = 'エラーが発生しました: ' + error.message;
       document.getElementById('errorMessage').style.display = 'block';
     }
    
    // ダイアログを閉じる
    function closeDialog() {
      google.script.host.close();
    }
    
    // 商品選択モーダルを開く
    function openProductModal() {
      document.getElementById('productModal').style.display = 'block';
      document.getElementById('productSearch').value = '';
      filterProducts(); // 全商品を表示
    }
    
    // 商品選択モーダルを閉じる
    function closeProductModal() {
      document.getElementById('productModal').style.display = 'none';
    }
    
    // 商品検索・フィルタ機能
    function filterProducts() {
      const searchTerm = document.getElementById('productSearch').value.toLowerCase();
      const rows = document.querySelectorAll('.product-row');
      
      rows.forEach(function(row) {
        const productName = row.querySelector('td:nth-child(1)').textContent.toLowerCase();
        const sku = row.querySelector('td:nth-child(2)').textContent.toLowerCase();
        const asin = row.querySelector('td:nth-child(3)').textContent.toLowerCase();
        
        if (productName.includes(searchTerm) || sku.includes(searchTerm) || asin.includes(searchTerm)) {
          row.classList.remove('hidden');
        } else {
          row.classList.add('hidden');
        }
      });
    }
    
    // 商品を選択
    function selectProduct(button) {
      const row = button.closest('tr');
      const productId = row.dataset.id;
      const productName = row.dataset.name;
      const sku = row.dataset.sku;
      const asin = row.dataset.asin;
      const stockStatus = row.dataset.stockStatus;
      const purchasePrice = row.dataset.purchasePrice;
      const sellingPrice = row.dataset.sellingPrice;
      
      // フォームに値を設定
      document.getElementById('productId').value = productId;
      document.getElementById('productDisplay').value = productName + ' (SKU: ' + sku + ', 在庫: ' + stockStatus + ')';
      
      // 価格フィールドに値を設定
      document.getElementById('purchasePrice').value = purchasePrice;
      document.getElementById('sellingPrice').value = sellingPrice;
      
      // 商品データをグローバル変数に保存（フォーム送信用）
      window.selectedProductData = {
        id: productId,
        name: productName,
        sku: sku,
        asin: asin,
        stockStatus: stockStatus,
        purchasePrice: purchasePrice,
        sellingPrice: sellingPrice
      };
      
      // 利益計算を実行
      calculateProfit();
      
      // モーダルを閉じる
      closeProductModal();
    }
    
    // モーダル外をクリックした時に閉じる
    window.onclick = function(event) {
      const modal = document.getElementById('productModal');
      if (event.target == modal) {
        closeProductModal();
      }
    }
  </script>
</body>
</html>
  `;
}