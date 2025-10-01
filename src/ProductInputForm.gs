/**
 * 新商品追加の入力ウィンドウ
 * 在庫管理ツール システム開発
 * 作成日: 2025-09-27
 */

/**
 * 新商品追加の入力フォームを表示
 */
function showProductInputForm() {
  const htmlTemplate = HtmlService.createTemplate(getProductInputFormHtml());
  htmlTemplate.nextProductId = getNextProductId();
  htmlTemplate.supplierData = getSupplierData();
  const htmlOutput = htmlTemplate.evaluate()
    .setWidth(800)
    .setHeight(600)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '新商品追加');
}

/**
 * 新商品を在庫管理シートに保存
 */
function saveNewProduct(formData) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const inventorySheet = spreadsheet.getSheetByName('在庫管理');
    
    if (!inventorySheet) {
      throw new Error('在庫管理シートが見つかりません');
    }
    
    // 備考列が存在するかチェックし、存在しない場合は追加
    const headerRange = inventorySheet.getRange(1, 1, 1, inventorySheet.getLastColumn());
    const headers = headerRange.getValues()[0];
    if (!headers.includes('備考・メモ')) {
      console.log('備考列が存在しないため、追加します');
      addNotesColumnToExistingSheet();
    }
    
    // 現在の日時を取得
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    // 利益計算（販売価格 - 仕入れ価格）
    // 商品の仕入れには手数料が発生しないため、手数料は考慮しない
    const profit = formData.sellingPrice - formData.purchasePrice;
    
    // 新しい行のデータを準備
    const newRowData = [
      formData.productId,           // 商品ID
      formData.productName,         // 商品名
      formData.sku || '',           // SKU
      formData.asin || '',          // ASIN
      formData.supplier,            // 仕入れ元
      formData.supplierUrl,         // 仕入れ元URL
      formData.purchasePrice,       // 仕入れ価格
      formData.sellingPrice,        // 販売価格
      formData.weight,              // 重量
      formData.stockStatus,         // 在庫ステータス
      profit,                       // 利益（計算値）
      timestamp,                    // 最終更新日時
      formData.notes || ''          // 備考・メモ
    ];
    
    // 在庫管理シートに新しい行を追加
    inventorySheet.appendRow(newRowData);
    
    // 新しく追加された行の書式設定
    const lastRow = inventorySheet.getLastRow();
    
    // 数値列の書式設定
    inventorySheet.getRange(lastRow, 1, 1, 1).setNumberFormat('0');        // 商品ID
    inventorySheet.getRange(lastRow, 7, 1, 1).setNumberFormat('#,##0');    // 仕入れ価格
    inventorySheet.getRange(lastRow, 8, 1, 1).setNumberFormat('#,##0');    // 販売価格
    inventorySheet.getRange(lastRow, 9, 1, 1).setNumberFormat('0');        // 重量
    inventorySheet.getRange(lastRow, 11, 1, 1).setNumberFormat('#,##0');   // 利益
    
    // 価格履歴を自動で作成
    try {
      updatePriceHistory(formData.productId, formData.purchasePrice, formData.sellingPrice, '新商品登録');
      console.log('価格履歴が正常に作成されました:', formData.productName);
    } catch (priceHistoryError) {
      console.warn('価格履歴の作成中にエラーが発生しました:', priceHistoryError);
      // 価格履歴の作成に失敗しても商品追加は継続
    }
    
    console.log('新商品が正常に保存されました:', formData.productName);
    
    return {
      success: true,
      message: '商品が正常に保存されました',
      productId: formData.productId,
      productName: formData.productName
    };
    
  } catch (error) {
    console.error('商品保存エラー:', error);
    return {
      success: false,
      message: '商品の保存中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * 仕入れ元情報を取得（仕入れ元マスターシートから）
 */
function getSupplierData() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const supplierSheet = spreadsheet.getSheetByName('仕入れ元マスター');
    
    if (!supplierSheet) {
      console.warn('仕入れ元マスターシートが見つかりません');
      return [];
    }
    
    const lastRow = supplierSheet.getLastRow();
    if (lastRow <= 1) {
      console.warn('仕入れ元マスターシートにデータがありません');
      return [];
    }
    
    // 仕入れ元データを取得（サイト名、手数料率、アクセス間隔、有効フラグ）
    const data = supplierSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    const suppliers = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const siteName = row[0];
      const feeRate = row[5]; // 手数料率（%）
      const accessInterval = row[6]; // アクセス間隔（秒）
      const activeFlag = row[7]; // 有効フラグ
      
      // 有効な仕入れ元のみを追加
      if (siteName && activeFlag === '有効') {
        const supplier = {
          name: siteName,
          feeRate: feeRate / 100, // パーセントを小数に変換（5.0 → 0.05）
          feeRatePercent: feeRate + '%',
          accessInterval: accessInterval
        };
        suppliers.push(supplier);
      }
    }
    
    return suppliers;
    
  } catch (error) {
    console.error('仕入れ元データ取得エラー:', error);
    return [];
  }
}

/**
 * 次の商品IDを生成（在庫管理シートの最大ID + 1）
 */
function getNextProductId() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const inventorySheet = spreadsheet.getSheetByName('在庫管理');
    
    if (!inventorySheet) {
      // 在庫管理シートが存在しない場合は初期値から開始
      return 1;
    }
    
    const lastRow = inventorySheet.getLastRow();
    if (lastRow <= 1) {
      // ヘッダーのみの場合は初期値から開始
      return 1;
    }
    
    // 商品ID列（A列）からデータを取得
    const productIds = inventorySheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    let maxNumber = 0;
    for (let i = 0; i < productIds.length; i++) {
      const productId = productIds[i][0];
      if (productId && typeof productId === 'number' && productId > 0) {
        if (productId > maxNumber) {
          maxNumber = productId;
        }
      }
    }
    
    // 次の番号を生成（自然数）
    return maxNumber + 1;
    
  } catch (error) {
    console.error('商品ID生成エラー:', error);
    // エラー時は初期値1を返す
    return 1;
  }
}

/**
 * 商品追加フォームのHTMLテンプレート
 */
function getProductInputFormHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>新商品追加</title>
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
      background: linear-gradient(135deg, #4285f4, #34a853);
      color: white;
      padding: 20px;
      text-align: center;
    }
    
    .header h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 300;
    }
    
    .form-container {
      padding: 30px;
    }
    
    .form-row {
      display: flex;
      gap: 20px;
      margin-bottom: 20px;
    }
    
    .form-group {
      flex: 1;
    }
    
    .form-group.full-width {
      flex: 100%;
    }
    
    .form-group label {
      display: block;
      margin-bottom: 8px;
      font-weight: 600;
      color: #555;
      font-size: 14px;
    }
    
    .form-group input,
    .form-group select,
    .form-group textarea {
      width: 100%;
      padding: 12px;
      border: 2px solid #e1e5e9;
      border-radius: 6px;
      font-size: 14px;
      transition: border-color 0.3s ease;
      box-sizing: border-box;
    }
    
    .form-group input:focus,
    .form-group select:focus,
    .form-group textarea:focus {
      outline: none;
      border-color: #4285f4;
      box-shadow: 0 0 0 3px rgba(66, 133, 244, 0.1);
    }
    
    .form-group textarea {
      resize: vertical;
      min-height: 80px;
    }
    
    .required {
      color: #ea4335;
    }
    
    .section-title {
      font-size: 18px;
      font-weight: 600;
      color: #333;
      margin: 30px 0 20px 0;
      padding-bottom: 10px;
      border-bottom: 2px solid #e1e5e9;
    }
    
    .section-title:first-child {
      margin-top: 0;
    }
    
    .button-group {
      display: flex;
      gap: 15px;
      justify-content: flex-end;
      margin-top: 30px;
      padding-top: 20px;
      border-top: 1px solid #e1e5e9;
    }
    
    .btn {
      padding: 12px 24px;
      border: none;
      border-radius: 6px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.3s ease;
      min-width: 100px;
    }
    
    .btn-primary {
      background: #4285f4;
      color: white;
    }
    
    .btn-primary:hover {
      background: #3367d6;
      transform: translateY(-1px);
    }
    
    .btn-secondary {
      background: #f8f9fa;
      color: #666;
      border: 2px solid #e1e5e9;
    }
    
    .btn-secondary:hover {
      background: #e9ecef;
      border-color: #adb5bd;
    }
    
    .error-message {
      color: #ea4335;
      font-size: 12px;
      margin-top: 5px;
      display: none;
    }
    
    .success-message {
      background: #d9ead3;
      color: #2d5016;
      padding: 15px;
      border-radius: 6px;
      margin-bottom: 20px;
      display: none;
    }
    
    .loading {
      opacity: 0.6;
      pointer-events: none;
    }
    
    .loading-overlay {
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(255, 255, 255, 0.9);
      display: none;
      justify-content: center;
      align-items: center;
      z-index: 1000;
      flex-direction: column;
    }
    
    .loading-spinner {
      width: 50px;
      height: 50px;
      border: 4px solid #f3f3f3;
      border-top: 4px solid #4285f4;
      border-radius: 50%;
      animation: spin 1s linear infinite;
      margin-bottom: 20px;
    }
    
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
    
    .loading-message {
      font-size: 16px;
      color: #333;
      font-weight: 600;
      text-align: center;
      margin-bottom: 10px;
    }
    
    .loading-details {
      font-size: 14px;
      color: #666;
      text-align: center;
    }
    
    .progress-bar {
      width: 300px;
      height: 6px;
      background: #e1e5e9;
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
    
    .supplier-info {
      background: #f8f9fa;
      padding: 15px;
      border-radius: 6px;
      margin-top: 10px;
      font-size: 12px;
      color: #666;
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
  </style>
</head>
<body>
  <!-- ローディングオーバーレイ -->
  <div class="loading-overlay" id="loadingOverlay">
    <div class="loading-spinner"></div>
    <div class="loading-message" id="loadingMessage">商品を保存中...</div>
    <div class="loading-details" id="loadingDetails">しばらくお待ちください</div>
    <div class="progress-bar">
      <div class="progress-fill" id="progressFill"></div>
    </div>
  </div>

  <div class="container">
    <div class="header">
      <h1>📦 新商品追加</h1>
    </div>
    
    <div class="form-container">
      <div class="success-message" id="successMessage">
        ✅ 商品が正常に追加されました！
      </div>
      
      <form id="productForm">
        <!-- 基本情報セクション -->
        <div class="section-title">📋 基本情報</div>
        
        <div class="form-row">
          <div class="form-group">
            <label for="productId">商品ID <span class="required">*</span></label>
            <input type="text" id="productId" name="productId" value="<?= nextProductId ?>" readonly style="background-color: #f5f5f5; color: #666;">
            <div class="error-message" id="productIdError">商品IDは自動生成されます（自然数）</div>
          </div>
          
          <div class="form-group">
            <label for="productName">商品名 <span class="required">*</span></label>
            <input type="text" id="productName" name="productName" placeholder="例: iPhone 15 Pro 128GB" required>
            <div class="error-message" id="productNameError">商品名は必須です</div>
          </div>
        </div>
        
        <div class="form-row">
          <div class="form-group">
            <label for="sku">SKU</label>
            <input type="text" id="sku" name="sku" placeholder="例: IPH15P-128">
          </div>
          
          <div class="form-group">
            <label for="asin">ASIN (Amazon商品コード)</label>
            <input type="text" id="asin" name="asin" placeholder="例: B0CHX1W1XY">
          </div>
        </div>
        
        <!-- 仕入れ情報セクション -->
        <div class="section-title">🏪 仕入れ情報</div>
        
        <div class="form-row">
          <div class="form-group">
            <label for="supplier">仕入れ元 <span class="required">*</span></label>
            <select id="supplier" name="supplier" required>
              <option value="">選択してください</option>
              <? for (var i = 0; i < supplierData.length; i++) { ?>
                <option value="<?= supplierData[i].name ?>"><?= supplierData[i].name ?></option>
              <? } ?>
            </select>
            <div class="error-message" id="supplierError">仕入れ元は必須です</div>
            <div class="supplier-info" id="supplierInfo" style="display: none;">
              アクセス間隔: <span id="accessInterval"></span>秒
            </div>
          </div>
          
          <div class="form-group">
            <label for="supplierUrl">仕入れ元URL <span class="required" id="supplierUrlRequired">*</span></label>
            <input type="url" id="supplierUrl" name="supplierUrl" placeholder="https://..." required>
            <div class="error-message" id="supplierUrlError">有効なURLを入力してください</div>
          </div>
        </div>
        
        <div class="form-row">
          <div class="form-group">
            <label for="purchasePrice">仕入れ価格 <span class="required">*</span></label>
            <input type="number" id="purchasePrice" name="purchasePrice" placeholder="120000" min="0" required>
            <div class="error-message" id="purchasePriceError">仕入れ価格は必須です</div>
          </div>
          
          <div class="form-group">
            <label for="sellingPrice">販売価格 <span class="required">*</span></label>
            <input type="number" id="sellingPrice" name="sellingPrice" placeholder="150000" min="0" required>
            <div class="error-message" id="sellingPriceError">販売価格は必須です</div>
          </div>
        </div>
        
        <div class="form-row">
          <div class="form-group">
            <label for="weight">重量 <span class="required">*</span></label>
            <input type="number" id="weight" name="weight" placeholder="187" min="0" required>
            <div class="error-message" id="weightError">重量は必須です（単位：g）</div>
          </div>
          
          <div class="form-group">
            <!-- 空のスペース -->
          </div>
        </div>
        
        <!-- 在庫情報セクション -->
        <div class="section-title">📦 在庫情報</div>
        
        <div class="form-group">
          <label for="stockStatus">在庫ステータス <span class="required">*</span></label>
          <select id="stockStatus" name="stockStatus" required>
            <option value="">在庫ステータスを選択してください</option>
            <option value="在庫あり">在庫あり</option>
            <option value="売り切れ">売り切れ</option>
            <option value="予約受付中">予約受付中</option>
          </select>
          <div class="error-message" id="stockStatusError">在庫ステータスは必須です</div>
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
            <span>予想利益:</span>
            <span id="previewProfit">¥0</span>
          </div>
        </div>
        
        <!-- 備考セクション -->
        <div class="section-title">📝 備考</div>
        
        <div class="form-group full-width">
          <label for="notes">備考・メモ</label>
          <textarea id="notes" name="notes" placeholder="商品に関する追加情報や注意事項を入力してください..."></textarea>
        </div>
      </form>
      
      <div class="button-group">
        <button type="button" class="btn btn-secondary" onclick="closeForm()">キャンセル</button>
        <button type="button" class="btn btn-primary" onclick="saveProduct()">商品を追加</button>
      </div>
    </div>
  </div>

  <script>
    // サーバーサイドから取得した仕入れ元データ
    var supplierData = <?= JSON.stringify(supplierData) ?>;
    
    // データが文字列の場合は解析する
    if (typeof supplierData === 'string') {
      try {
        supplierData = JSON.parse(supplierData);
      } catch (e) {
        console.error('JSON解析エラー:', e);
        supplierData = [];
      }
    }
    
    // 処理状態管理
    var isProcessing = false;
    var progressSteps = [
      'データ検証中...',
      'スプレッドシートに保存中...',
      '書式設定中...',
      '完了！'
    ];
    var currentStep = 0;
    
    // 仕入れ元選択時の処理
    document.getElementById('supplier').addEventListener('change', function() {
      updateSupplierInfo();
      updateSupplierUrlRequirement(); // URL必須属性の制御
      updatePricePreview(); // 価格プレビューを更新
    });
    
    // 価格入力時の処理
    document.getElementById('purchasePrice').addEventListener('input', updatePricePreview);
    document.getElementById('sellingPrice').addEventListener('input', updatePricePreview);
    
    
    // 仕入れ元情報の更新
    function updateSupplierInfo() {
      var supplier = document.getElementById('supplier').value;
      var supplierInfo = document.getElementById('supplierInfo');
      var accessIntervalSpan = document.getElementById('accessInterval');
      
      // サーバーサイドから取得したデータから仕入れ元情報を検索
      var selectedSupplier = null;
      for (var i = 0; i < supplierData.length; i++) {
        if (supplierData[i].name === supplier) {
          selectedSupplier = supplierData[i];
          break;
        }
      }
      
      if (selectedSupplier) {
        accessIntervalSpan.textContent = selectedSupplier.accessInterval;
        supplierInfo.style.display = 'block';
      } else {
        supplierInfo.style.display = 'none';
      }
    }
    
    // 仕入れ元URL必須属性の制御
    function updateSupplierUrlRequirement() {
      var supplier = document.getElementById('supplier').value;
      var supplierUrlInput = document.getElementById('supplierUrl');
      var supplierUrlRequired = document.getElementById('supplierUrlRequired');
      var supplierUrlError = document.getElementById('supplierUrlError');
      
      // Amazonが選択された場合は必須を外す
      if (supplier === 'Amazon') {
        supplierUrlInput.removeAttribute('required');
        supplierUrlRequired.style.display = 'none';
        supplierUrlError.textContent = '有効なURLを入力してください（任意）';
      } else {
        supplierUrlInput.setAttribute('required', 'required');
        supplierUrlRequired.style.display = 'inline';
        supplierUrlError.textContent = '有効なURLを入力してください';
      }
    }
    
    // 価格プレビューの更新
    function updatePricePreview() {
      var purchasePrice = parseFloat(document.getElementById('purchasePrice').value) || 0;
      var sellingPrice = parseFloat(document.getElementById('sellingPrice').value) || 0;
      
      // 利益計算（手数料は考慮しない）
      var profit = sellingPrice - purchasePrice;
      
      // プレビューを表示
      document.getElementById('previewSellingPrice').textContent = '¥' + sellingPrice.toLocaleString();
      document.getElementById('previewPurchasePrice').textContent = '¥' + purchasePrice.toLocaleString();
      document.getElementById('previewProfit').textContent = '¥' + Math.round(profit).toLocaleString();
      
      // 利益の色分け
      var profitElement = document.getElementById('previewProfit');
      if (profit > 0) {
        profitElement.style.color = '#2d5016';
      } else {
        profitElement.style.color = '#ea4335';
      }
      
      // プレビューを表示
      if (sellingPrice > 0 || purchasePrice > 0) {
        document.getElementById('pricePreview').style.display = 'block';
      }
    }
    
    // ローディング表示の制御
    function showLoadingOverlay() {
      isProcessing = true;
      currentStep = 0;
      document.getElementById('loadingOverlay').style.display = 'flex';
      document.getElementById('loadingMessage').textContent = progressSteps[0];
      document.getElementById('loadingDetails').textContent = 'しばらくお待ちください';
      document.getElementById('progressFill').style.width = '0%';
      
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
    
    // フォームの保存
    function saveProduct() {
      // 重複送信防止
      if (isProcessing) {
        console.log('処理中です。しばらくお待ちください。');
        return;
      }
      
      if (validateForm()) {
        // ローディング表示開始
        showLoadingOverlay();
        
        // プログレス更新（データ検証）
        updateProgress(0, '入力データを検証しています...');
        
        // フォームデータを収集
        var formData = {
          productId: parseInt(document.getElementById('productId').value),
          productName: document.getElementById('productName').value,
          sku: document.getElementById('sku').value,
          asin: document.getElementById('asin').value,
          supplier: document.getElementById('supplier').value,
          supplierUrl: document.getElementById('supplierUrl').value,
          purchasePrice: parseFloat(document.getElementById('purchasePrice').value),
          sellingPrice: parseFloat(document.getElementById('sellingPrice').value),
          weight: parseInt(document.getElementById('weight').value),
          stockStatus: document.getElementById('stockStatus').value,
          notes: document.getElementById('notes').value
        };
        
        // プログレス更新（保存開始）
        setTimeout(function() {
          updateProgress(1, 'スプレッドシートに保存中...');
        }, 500);
        
        // サーバーサイドの保存関数を呼び出し
        google.script.run
          .withSuccessHandler(onSaveSuccess)
          .withFailureHandler(onSaveError)
          .saveNewProduct(formData);
      }
    }
    
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
            successMessage.innerHTML = '✅ ' + result.message + '<br>商品ID: ' + result.productId + '<br>商品名: ' + result.productName;
            successMessage.style.display = 'block';
            successMessage.style.background = '#d9ead3';
            successMessage.style.color = '#2d5016';
            
            // フォームをクリア（商品IDは再生成）
            document.getElementById('productForm').reset();
            // 商品IDを再生成（新しい商品追加のため）
            generateNewProductId();
            
            // 3秒後にフォームを閉じる
            setTimeout(function() {
              closeForm();
            }, 3000);
          }, 800);
        }, 500);
      } else {
        // エラーメッセージを表示
        hideLoadingOverlay();
        showError(result.message);
      }
    }
    
    // 保存失敗時の処理
    function onSaveError(error) {
      hideLoadingOverlay();
      console.error('保存エラー:', error);
      showError('サーバーエラーが発生しました: ' + error.message);
    }
    
    // エラーメッセージ表示
    function showError(message) {
      var successMessage = document.getElementById('successMessage');
      successMessage.innerHTML = '❌ ' + message;
      successMessage.style.display = 'block';
      successMessage.style.background = '#f4cccc';
      successMessage.style.color = '#ea4335';
      
      // エラーメッセージを自動で非表示にしない（ユーザーが確認できるように）
      // 5秒後にエラーメッセージを非表示
      setTimeout(function() {
        successMessage.style.display = 'none';
        successMessage.style.background = '#d9ead3';
        successMessage.style.color = '#2d5016';
      }, 8000);
    }
    
    // フォームバリデーション
    function validateForm() {
      var isValid = true;
      
      // エラーメッセージをクリア
      document.querySelectorAll('.error-message').forEach(function(error) {
        error.style.display = 'none';
      });
      
      // 必須項目のチェック（商品IDは自動生成のため除外）
      var requiredFields = ['productName', 'supplier', 'purchasePrice', 'sellingPrice', 'weight', 'stockStatus'];
      
      // 仕入れ元がAmazonでない場合のみURLを必須項目に追加
      var supplier = document.getElementById('supplier').value;
      if (supplier !== 'Amazon') {
        requiredFields.push('supplierUrl');
      }
      
      requiredFields.forEach(function(fieldName) {
        var field = document.getElementById(fieldName);
        var errorElement = document.getElementById(fieldName + 'Error');
        
        if (!field.value.trim()) {
          errorElement.style.display = 'block';
          isValid = false;
        }
      });
      
      // URLの形式チェック（Amazonでない場合のみ、または値が入力されている場合）
      var urlField = document.getElementById('supplierUrl');
      var urlError = document.getElementById('supplierUrlError');
      
      if (urlField.value && !isValidUrl(urlField.value)) {
        urlError.style.display = 'block';
        isValid = false;
      }
      
      // 価格の妥当性チェック
      var purchasePrice = parseFloat(document.getElementById('purchasePrice').value);
      var sellingPrice = parseFloat(document.getElementById('sellingPrice').value);
      
      if (sellingPrice <= purchasePrice) {
        document.getElementById('sellingPriceError').textContent = '販売価格は仕入れ価格より高く設定してください';
        document.getElementById('sellingPriceError').style.display = 'block';
        isValid = false;
      }
      
      return isValid;
    }
    
    // URL形式チェック
    function isValidUrl(string) {
      try {
        new URL(string);
        return true;
      } catch (_) {
        return false;
      }
    }
    
    // フォームを閉じる
    function closeForm() {
      google.script.host.close();
    }
    
    // 新しい商品IDを生成
    function generateNewProductId() {
      google.script.run
        .withSuccessHandler(function(newId) {
          document.getElementById('productId').value = newId;
        })
        .withFailureHandler(function(error) {
          console.error('商品ID生成エラー:', error);
        })
        .getNextProductId();
    }
    
    // ページ読み込み時の初期化
    document.addEventListener('DOMContentLoaded', function() {
      // 商品IDはサーバーサイドで自動生成済み
      console.log('商品ID:', document.getElementById('productId').value);
      
      // 初期状態でURL必須属性を制御
      updateSupplierUrlRequirement();
    });
  </script>
</body>
</html>
  `;
}

/**
 * 備考機能のテスト用関数
 */
function testNotesFunctionality() {
  console.log('備考機能のテストを開始します...');
  
  try {
    // テスト用の商品データを作成
    const testFormData = {
      productId: 999,
      productName: 'テスト商品（備考機能テスト）',
      sku: 'TEST-001',
      asin: 'B0TEST001',
      supplier: 'Amazon',
      supplierUrl: 'https://amazon.co.jp/test',
      purchasePrice: 1000,
      sellingPrice: 1500,
      weight: 100,
      stockStatus: '在庫あり',
      notes: 'これは備考機能のテスト用商品です。備考・メモが正常に保存されることを確認します。'
    };
    
    // 商品を保存
    const result = saveNewProduct(testFormData);
    
    if (result.success) {
      console.log('✅ 備考機能のテストが成功しました');
      console.log('保存された商品:', result.productName);
      
      // 保存されたデータを確認
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const inventorySheet = spreadsheet.getSheetByName('在庫管理');
      const lastRow = inventorySheet.getLastRow();
      const notesValue = inventorySheet.getRange(lastRow, 13).getValue(); // M列（備考列）
      
      console.log('保存された備考:', notesValue);
      
      if (notesValue === testFormData.notes) {
        console.log('✅ 備考データが正常に保存されました');
        return true;
      } else {
        console.log('❌ 備考データの保存に問題があります');
        return false;
      }
    } else {
      console.log('❌ 商品の保存に失敗しました:', result.message);
      return false;
    }
    
  } catch (error) {
    console.error('❌ 備考機能のテスト中にエラーが発生しました:', error);
    return false;
  }
}
