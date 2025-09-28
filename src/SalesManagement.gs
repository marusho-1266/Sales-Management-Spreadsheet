/**
 * 売上管理シートの作成と管理
 * 在庫管理ツール システム開発
 * 作成日: 2025-09-28
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
    'ASIN',
    '数量',
    '販売価格',
    '仕入れ価格',
    '送料',
    '純利益',
    '登録日時'
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
    if (row[0] && row[1] && row[9] === '在庫あり') { // 商品ID、商品名が存在し、在庫ステータスが「在庫あり」の場合
      products.push({
        id: row[0],
        name: row[1],
        sku: row[2],
        asin: row[3],
        stockStatus: row[9],
        purchasePrice: row[6]
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
    
    // 純利益を計算（手数料は考慮しない）
    const netProfit = salesData.sellingPrice - salesData.purchasePrice;
    
    // 売上データを追加
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
      netProfit,
      registrationTime
    ];
    
    // データを追加
    const lastRow = salesSheet.getLastRow();
    salesSheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
    
     console.log('売上データが正常に追加されました:', salesData.orderId);
     return { 
       success: true, 
       message: '注文データが正常に保存されました',
       orderId: salesData.orderId,
       productName: salesData.productName
     };
    
  } catch (error) {
    console.error('売上データの追加中にエラーが発生しました:', error);
    return { success: false, message: 'データの保存中にエラーが発生しました: ' + error.message };
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
 * 売上データの一覧表示
 */
function showSalesList() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = spreadsheet.getSheetByName('売上管理');
  
  if (!salesSheet) {
    SpreadsheetApp.getUi().alert('エラー', '売上管理シートが見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // 売上管理シートをアクティブにする
  salesSheet.activate();
  
  console.log('売上管理シートを表示しました');
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
 * 売上統計の取得
 */
function getSalesStatistics() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = spreadsheet.getSheetByName('売上管理');
    
    if (!salesSheet) {
      console.error('売上管理シートが見つかりません');
      return null;
    }
    
    const data = salesSheet.getDataRange().getValues();
    let totalSales = 0;
    let totalProfit = 0;
    let orderCount = 0;
    
    // ヘッダー行をスキップして計算
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && !isNaN(data[i][0]) && data[i][0] > 0) { // 注文IDが存在し、数値で正の値の場合
        totalSales += data[i][7] || 0; // 販売価格（列番号8）
        totalProfit += data[i][10] || 0; // 純利益（列番号11）
        orderCount++;
      }
    }
    
    return {
      totalSales: totalSales,
      totalProfit: totalProfit,
      orderCount: orderCount,
      averageOrderValue: orderCount > 0 ? totalSales / orderCount : 0
    };
    
  } catch (error) {
    console.error('売上統計の取得中にエラーが発生しました:', error);
    return null;
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
          <input type="number" id="orderId" name="orderId" value="<?= nextOrderId ?>" min="1" required>
        </div>
        
        <div class="form-group">
          <label for="orderDate">注文日 *</label>
          <input type="date" id="orderDate" name="orderDate" required>
        </div>
        
        <div class="form-group">
          <label for="productId">商品選択 *</label>
          <select id="productId" name="productId" required onchange="updateProductInfo()">
            <option value="">商品を選択してください</option>
            <? for (var i = 0; i < inventoryData.length; i++) { ?>
              <option value="<?= inventoryData[i].id ?>" 
                      data-name="<?= inventoryData[i].name ?>"
                      data-sku="<?= inventoryData[i].sku ?>"
                      data-asin="<?= inventoryData[i].asin ?>"
                      data-stock-status="<?= inventoryData[i].stockStatus ?>"
                      data-purchase-price="<?= inventoryData[i].purchasePrice ?>">
                <?= inventoryData[i].name ?> (SKU: <?= inventoryData[i].sku ?>, 在庫: <?= inventoryData[i].stockStatus ?>)
              </option>
            <? } ?>
          </select>
        </div>
        
        <div id="productInfo" class="product-info" style="display: none;">
          <h4>選択された商品情報</h4>
          <p><strong>商品名:</strong> <span id="selectedProductName"></span></p>
          <p><strong>SKU:</strong> <span id="selectedSKU"></span></p>
          <p><strong>ASIN:</strong> <span id="selectedASIN"></span></p>
          <p><strong>在庫ステータス:</strong> <span id="selectedStockStatus"></span></p>
          <p><strong>仕入れ価格:</strong> ¥<span id="selectedPurchasePrice"></span></p>
        </div>
        
        <div class="form-row">
          <div class="form-group">
            <label for="quantity">数量 *</label>
            <input type="number" id="quantity" name="quantity" min="1" required onchange="validateQuantity()">
            <div id="quantityError" class="error-message"></div>
          </div>
          
          <div class="form-group">
            <label for="sellingPrice">販売価格 (円) *</label>
            <input type="number" id="sellingPrice" name="sellingPrice" min="1" required onchange="calculateProfit()">
          </div>
        </div>
        
         <div class="form-row">
           <div class="form-group">
             <label for="shippingCost">送料 (円)</label>
             <input type="number" id="shippingCost" name="shippingCost" min="0" value="0" onchange="calculateProfit()">
           </div>
           
           <div class="form-group">
             <!-- 空のスペース -->
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
     
     // 商品選択時の処理
     function updateProductInfo() {
      const select = document.getElementById('productId');
      const option = select.options[select.selectedIndex];
      const productInfo = document.getElementById('productInfo');
      
      if (option.value) {
        document.getElementById('selectedProductName').textContent = option.dataset.name;
        document.getElementById('selectedSKU').textContent = option.dataset.sku;
        document.getElementById('selectedASIN').textContent = option.dataset.asin;
        document.getElementById('selectedStockStatus').textContent = option.dataset.stockStatus;
        document.getElementById('selectedPurchasePrice').textContent = parseInt(option.dataset.purchasePrice).toLocaleString();
        
        productInfo.style.display = 'block';
        
        // 利益計算を実行
        calculateProfit();
      } else {
        productInfo.style.display = 'none';
      }
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
    
     // 利益計算
     function calculateProfit() {
       const sellingPrice = parseFloat(document.getElementById('sellingPrice').value) || 0;
       const purchasePrice = parseFloat(document.getElementById('productId').options[document.getElementById('productId').selectedIndex].dataset.purchasePrice) || 0;
       const netProfit = sellingPrice - purchasePrice;
       
       document.getElementById('netProfit').value = Math.round(netProfit);
       
       // プレビュー表示を更新
       updatePricePreview(sellingPrice, purchasePrice, netProfit);
     }
     
     // 価格プレビューの更新
     function updatePricePreview(sellingPrice, purchasePrice, netProfit) {
       const pricePreview = document.getElementById('pricePreview');
       
       if (sellingPrice > 0) {
         document.getElementById('previewSellingPrice').textContent = '¥' + sellingPrice.toLocaleString();
         document.getElementById('previewPurchasePrice').textContent = '¥' + purchasePrice.toLocaleString();
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
       
       // フォームデータを取得
       const formData = {
         orderId: parseInt(document.getElementById('orderId').value),
         orderDate: document.getElementById('orderDate').value,
         productId: parseInt(document.getElementById('productId').value),
         sku: document.getElementById('productId').options[document.getElementById('productId').selectedIndex].dataset.sku,
         asin: document.getElementById('productId').options[document.getElementById('productId').selectedIndex].dataset.asin,
         productName: document.getElementById('productId').options[document.getElementById('productId').selectedIndex].dataset.name,
         quantity: parseInt(document.getElementById('quantity').value),
         sellingPrice: parseInt(document.getElementById('sellingPrice').value),
         purchasePrice: parseInt(document.getElementById('productId').options[document.getElementById('productId').selectedIndex].dataset.purchasePrice),
         shippingCost: parseInt(document.getElementById('shippingCost').value) || 0
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
             document.getElementById('productInfo').style.display = 'none';
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
  </script>
</body>
</html>
  `;
}