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
        suppliers.push({
          name: siteName,
          feeRate: feeRate / 100, // パーセントを小数に変換
          feeRatePercent: feeRate + '%',
          accessInterval: accessInterval
        });
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
              手数料率: <span id="feeRate"></span> | アクセス間隔: <span id="accessInterval"></span>秒
            </div>
          </div>
          
          <div class="form-group">
            <label for="supplierUrl">仕入れ元URL <span class="required">*</span></label>
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
        
        <div class="form-row">
          <div class="form-group">
            <label for="stockQuantity">在庫数 <span class="required">*</span></label>
            <input type="number" id="stockQuantity" name="stockQuantity" placeholder="5" min="0" value="1" required>
            <div class="error-message" id="stockQuantityError">在庫数は必須です</div>
          </div>
          
          <div class="form-group">
            <label for="stockStatus">在庫ステータス</label>
            <select id="stockStatus" name="stockStatus">
              <option value="在庫あり">在庫あり</option>
              <option value="売り切れ">売り切れ</option>
              <option value="予約受付中">予約受付中</option>
            </select>
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
            <span>手数料 (<span id="previewFeeRate">0%</span>):</span>
            <span id="previewFee">¥0</span>
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
    const supplierData = <?= JSON.stringify(supplierData) ?>;
    
    // 仕入れ元選択時の処理
    document.getElementById('supplier').addEventListener('change', function() {
      updateSupplierInfo();
      updatePricePreview();
    });
    
    // 価格入力時の処理
    document.getElementById('purchasePrice').addEventListener('input', updatePricePreview);
    document.getElementById('sellingPrice').addEventListener('input', updatePricePreview);
    
    // 在庫数入力時の処理
    document.getElementById('stockQuantity').addEventListener('input', function() {
      const quantity = parseInt(this.value);
      const statusSelect = document.getElementById('stockStatus');
      
      if (quantity === 0) {
        statusSelect.value = '売り切れ';
      } else if (quantity > 0) {
        statusSelect.value = '在庫あり';
      }
    });
    
    // 仕入れ元情報の更新
    function updateSupplierInfo() {
      const supplier = document.getElementById('supplier').value;
      const supplierInfo = document.getElementById('supplierInfo');
      const feeRateSpan = document.getElementById('feeRate');
      const accessIntervalSpan = document.getElementById('accessInterval');
      
      // サーバーサイドから取得したデータから仕入れ元情報を検索
      const selectedSupplier = supplierData.find(s => s.name === supplier);
      
      if (selectedSupplier) {
        feeRateSpan.textContent = selectedSupplier.feeRatePercent;
        accessIntervalSpan.textContent = selectedSupplier.accessInterval;
        supplierInfo.style.display = 'block';
      } else {
        supplierInfo.style.display = 'none';
      }
    }
    
    // 価格プレビューの更新
    function updatePricePreview() {
      const purchasePrice = parseFloat(document.getElementById('purchasePrice').value) || 0;
      const sellingPrice = parseFloat(document.getElementById('sellingPrice').value) || 0;
      const supplier = document.getElementById('supplier').value;
      
      // サーバーサイドから取得したデータから手数料率を取得
      const selectedSupplier = supplierData.find(s => s.name === supplier);
      const feeRate = selectedSupplier ? selectedSupplier.feeRate : 0;
      const fee = purchasePrice * feeRate;
      const profit = sellingPrice - purchasePrice - fee;
      
      // プレビューを表示
      document.getElementById('previewSellingPrice').textContent = '¥' + sellingPrice.toLocaleString();
      document.getElementById('previewPurchasePrice').textContent = '¥' + purchasePrice.toLocaleString();
      document.getElementById('previewFeeRate').textContent = (feeRate * 100).toFixed(1) + '%';
      document.getElementById('previewFee').textContent = '¥' + fee.toLocaleString();
      document.getElementById('previewProfit').textContent = '¥' + profit.toLocaleString();
      
      // 利益の色分け
      const profitElement = document.getElementById('previewProfit');
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
    
    // フォームの保存
    function saveProduct() {
      if (validateForm()) {
        // ローディング状態
        document.body.classList.add('loading');
        
        // フォームデータを収集
        const formData = {
          productId: document.getElementById('productId').value,
          productName: document.getElementById('productName').value,
          sku: document.getElementById('sku').value,
          asin: document.getElementById('asin').value,
          supplier: document.getElementById('supplier').value,
          supplierUrl: document.getElementById('supplierUrl').value,
          purchasePrice: parseFloat(document.getElementById('purchasePrice').value),
          sellingPrice: parseFloat(document.getElementById('sellingPrice').value),
          weight: parseInt(document.getElementById('weight').value),
          stockQuantity: parseInt(document.getElementById('stockQuantity').value),
          stockStatus: document.getElementById('stockStatus').value,
          notes: document.getElementById('notes').value
        };
        
        // 成功メッセージを表示（実際の保存処理は実装しない）
        setTimeout(() => {
          document.body.classList.remove('loading');
          document.getElementById('successMessage').style.display = 'block';
          
          // 3秒後にフォームを閉じる
          setTimeout(() => {
            closeForm();
          }, 3000);
        }, 1000);
      }
    }
    
    // フォームバリデーション
    function validateForm() {
      let isValid = true;
      
      // エラーメッセージをクリア
      document.querySelectorAll('.error-message').forEach(error => {
        error.style.display = 'none';
      });
      
      // 必須項目のチェック（商品IDは自動生成のため除外）
      const requiredFields = ['productName', 'supplier', 'supplierUrl', 'purchasePrice', 'sellingPrice', 'weight', 'stockQuantity'];
      
      requiredFields.forEach(fieldName => {
        const field = document.getElementById(fieldName);
        const errorElement = document.getElementById(fieldName + 'Error');
        
        if (!field.value.trim()) {
          errorElement.style.display = 'block';
          isValid = false;
        }
      });
      
      // URLの形式チェック
      const urlField = document.getElementById('supplierUrl');
      const urlError = document.getElementById('supplierUrlError');
      
      if (urlField.value && !isValidUrl(urlField.value)) {
        urlError.style.display = 'block';
        isValid = false;
      }
      
      // 価格の妥当性チェック
      const purchasePrice = parseFloat(document.getElementById('purchasePrice').value);
      const sellingPrice = parseFloat(document.getElementById('sellingPrice').value);
      
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
    
    // ページ読み込み時の初期化
    document.addEventListener('DOMContentLoaded', function() {
      // 商品IDはサーバーサイドで自動生成済み
      console.log('商品ID:', document.getElementById('productId').value);
    });
  </script>
</body>
</html>
  `;
}
