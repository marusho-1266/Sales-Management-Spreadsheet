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
  htmlTemplate.categoryData = getCategoryData();
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
    
    // 販売価格（USD）列が存在するかチェックし、存在しない場合は追加
    const headerRange = inventorySheet.getRange(1, 1, 1, inventorySheet.getLastColumn());
    const headers = headerRange.getValues()[0];
    if (!headers.includes('JoomID')) {
      console.log('JoomID列が存在しないため、追加します');
      addJoomIdColumnToExistingSheet();
    }
    if (!headers.includes('販売価格（USD）')) {
      console.log('販売価格（USD）列が存在しないため、追加します');
      addSellingPriceUsdColumnToExistingSheet();
    }
    
    // 備考列が存在するかチェックし、存在しない場合は追加
    if (!headers.includes('備考・メモ')) {
      console.log('備考列が存在しないため、追加します');
      addNotesColumnToExistingSheet();
    }
    
    // カテゴリー列が存在するかチェックし、存在しない場合は追加
    if (!headers.includes('商品カテゴリー')) {
      console.log('カテゴリー列が存在しないため、追加します');
      addCategoryColumnToExistingSheet();
    }
    
    // 利益計算関連列が存在するかチェックし、存在しない場合は追加
    if (!headers.includes('返金額(円)') || !headers.includes('最終為替レート')) {
      console.log('利益計算関連列が存在しないため、追加します');
      addProfitRelatedColumnsToExistingSheet();
    }
    
    // 現在の日時を取得
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    // 販売価格（USD）: パースして NaN の場合は空にする
    const sellingPriceUsdParsed = (formData.sellingPriceUsd !== '' && formData.sellingPriceUsd != null)
      ? parseFloat(formData.sellingPriceUsd)
      : NaN;
    const sellingPriceUsd = Number.isNaN(sellingPriceUsdParsed) ? '' : sellingPriceUsdParsed;
    
    // 新しい行のデータを準備（Joom対応フィールド含む）
    const newRowData = [
      formData.productId,           // 商品ID
      formData.productName,         // 商品名
      formData.joomId || '',        // JoomID
      formData.sku || '',           // SKU
      formData.asin || '',          // ASIN
      formData.supplier,            // 仕入れ元
      formData.supplierUrl,         // 仕入れ元URL
      formData.purchasePrice,       // 仕入れ価格
      formData.sellingPrice || 0,        // 販売価格
      sellingPriceUsd,              // 販売価格（USD）任意（不正時は空）
      formData.weight,              // 重量
      // 容積重量計算用寸法フィールド（重量の後ろに配置）
      formData.heightCm || 0,       // 高さ(cm)
      formData.lengthCm || 0,       // 長さ(cm)
      formData.widthCm || 0,        // 幅(cm)
      // 利益計算用カテゴリーフィールド
      formData.category || '',      // 商品カテゴリー
      // Joom対応フィールド
      formData.description || '',   // 商品説明
      formData.mainImageUrl || '',  // メイン画像URL
      formData.currency || 'JPY',   // 通貨
      formData.shippingPrice || 0,  // 配送価格
      formData.stockQuantity || (formData.stockStatus === '在庫あり' ? 1 : 0), // 在庫数量
      // 既存フィールド
      formData.stockStatus,         // 在庫ステータス
      // 利益計算関連フィールド（初期値0）
      0,                            // 返金額(円)
      0,                            // Joom手数料(円)
      0,                            // サーチャージ(円)
      0,                            // 繁忙期料金(円)
      '',                           // 利益（計算式で設定）
      0,                            // 最終為替レート
      timestamp,                    // 最終更新日時
      formData.notes || '',         // 備考・メモ
      // Joom連携管理列
      '未連携',                      // Joom連携ステータス
      ''                            // 最終出力日時
    ];
    
    // 在庫管理シートに新しい行を追加
    inventorySheet.appendRow(newRowData);
    
    // 新しく追加された行の書式設定
    const lastRow = inventorySheet.getLastRow();
    
    // 数値列の書式設定
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.PRODUCT_ID, 1, 1).setNumberFormat('0');        // 商品ID
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.PURCHASE_PRICE, 1, 1).setNumberFormat('#,##0');    // 仕入れ価格
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.SELLING_PRICE, 1, 1).setNumberFormat('#,##0');    // 販売価格
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.SELLING_PRICE_USD, 1, 1).setNumberFormat('#,##0.00');  // 販売価格（USD）
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.WEIGHT, 1, 1).setNumberFormat('0');        // 重量
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.SHIPPING_PRICE, 1, 1).setNumberFormat('#,##0');   // 配送価格
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.STOCK_QUANTITY, 1, 1).setNumberFormat('0');       // 在庫数量
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.REFUND_AMOUNT, 1, 1).setNumberFormat('#,##0');   // 返金額(円)
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.JOOM_FEE, 1, 1).setNumberFormat('#,##0');   // Joom手数料(円)
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.SURCHARGE, 1, 1).setNumberFormat('#,##0');   // サーチャージ(円)
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.PEAK_SEASON_FEE, 1, 1).setNumberFormat('#,##0');   // 繁忙期料金(円)
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.EXCHANGE_RATE, 1, 1).setNumberFormat('#,##0.00');   // 最終為替レート
    
    // 利益計算式: 販売価格-(仕入価格+配送価格+返金額+Joom手数料+サーチャージ+繁忙料金)
    // I列: 販売価格, H列: 仕入価格, S列: 配送価格, V列: 返金額, W列: Joom手数料, X列: サーチャージ, Y列: 繁忙料金
    const profitFormula = `=I${lastRow}-(H${lastRow}+S${lastRow}+V${lastRow}+W${lastRow}+X${lastRow}+Y${lastRow})`;
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.PROFIT, 1, 1).setFormula(profitFormula);
    inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.PROFIT, 1, 1).setNumberFormat('#,##0');   // 利益
    
    // 価格履歴を自動で作成
    try {
      updatePriceHistory(formData.productId, formData.purchasePrice, formData.sellingPrice, '新商品登録');
      console.log('価格履歴が正常に作成されました:', formData.productName);
    } catch (priceHistoryError) {
      console.warn('価格履歴の作成中にエラーが発生しました:', priceHistoryError);
      // 価格履歴の作成に失敗しても商品追加は継続
    }
    
    // 利益計算シートの商品IDドロップダウンを更新（新規追加をリストに反映）
    try {
      if (typeof refreshProfitSheetProductIdDropdown === 'function') {
        refreshProfitSheetProductIdDropdown();
      }
    } catch (dropdownError) {
      console.warn('利益計算ドロップダウン更新中にエラーが発生しました:', dropdownError);
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
 * カテゴリー一覧を取得（関税率マスタから）
 */
function getCategoryData() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const dutyMaster = spreadsheet.getSheetByName('関税率マスタ');
    
    if (!dutyMaster) {
      console.warn('関税率マスタシートが見つかりません');
      // デフォルトのカテゴリーリストを返す
      return ['一般', '家電', 'アパレル', '雑貨', 'ホビー'];
    }
    
    const lastRow = dutyMaster.getLastRow();
    if (lastRow <= 1) {
      console.warn('関税率マスタシートにデータがありません');
      // デフォルトのカテゴリーリストを返す
      return ['一般', '家電', 'アパレル', '雑貨', 'ホビー'];
    }
    
    // カテゴリー名を取得（B列、2行目以降）
    const categoryRange = dutyMaster.getRange(2, 2, lastRow - 1, 1);
    const categories = categoryRange.getValues();
    
    // 空の値を除外してカテゴリーリストを作成
    const categoryList = [];
    for (let i = 0; i < categories.length; i++) {
      const category = categories[i][0];
      if (category && category.toString().trim() !== '') {
        categoryList.push(category.toString().trim());
      }
    }
    
    // カテゴリーが存在しない場合はデフォルトリストを返す
    if (categoryList.length === 0) {
      return ['一般', '家電', 'アパレル', '雑貨', 'ホビー'];
    }
    
    return categoryList;
    
  } catch (error) {
    console.error('カテゴリーデータ取得エラー:', error);
    // エラー時はデフォルトのカテゴリーリストを返す
    return ['一般', '家電', 'アパレル', '雑貨', 'ホビー'];
  }
}

/**
 * 仕入れ元情報を取得（仕入れ元マスターシートから）
 */
function getSupplierData() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // Constants.gsの定数を使用してシート名を取得
    const supplierSheet = spreadsheet.getSheetByName(SHEET_NAMES.SUPPLIER_MASTER);
    
    if (!supplierSheet) {
      console.warn('仕入れ元マスターシートが見つかりません。シート名:', SHEET_NAMES.SUPPLIER_MASTER);
      return [];
    }
    
    const lastRow = supplierSheet.getLastRow();
    if (lastRow <= 1) {
      console.warn('仕入れ元マスターシートにデータがありません');
      return [];
    }
    
    // 仕入れ元データを取得（全11列を取得）
    // 列構成: 1=サイト名, 2=URLパターン, 3=価格セレクタ, 4=価格除外セレクタ, 
    //         5=在庫セレクタ, 6=在庫ありキーワード, 7=売り切れキーワード,
    //         8=手数料率(%), 9=アクセス間隔(秒), 10=有効フラグ, 11=備考
    const data = supplierSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    
    const suppliers = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const siteName = row[0]; // 列1: サイト名
      const feeRate = row[7]; // 列8: 手数料率（%）
      const accessInterval = row[8]; // 列9: アクセス間隔（秒）
      const activeFlag = row[9]; // 列10: 有効フラグ
      
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
    
    console.log(`仕入れ元データを取得しました: ${suppliers.length}件`);
    return suppliers;
    
  } catch (error) {
    console.error('仕入れ元データ取得エラー:', error);
    console.error('エラースタック:', error.stack);
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
    
    // 商品ID列からデータを取得
    const productIds = inventorySheet.getRange(2, COLUMN_INDEXES.INVENTORY.PRODUCT_ID, lastRow - 1, 1).getValues();
    
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
    
    .validation-error-banner {
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      padding: 12px 20px;
      background: #f4cccc;
      color: #ea4335;
      font-weight: 600;
      font-size: 14px;
      text-align: center;
      z-index: 999;
      box-shadow: 0 2px 8px rgba(0,0,0,0.15);
      display: none;
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
  <!-- バリデーションエラー表示（スクロール位置に関係なく常に表示） -->
  <div class="validation-error-banner" id="validationErrorBanner">
    ⚠️ 入力内容に誤りがあります。該当箇所へスクロールしました。
  </div>
  
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
          <div class="form-group full-width">
            <label for="joomId">JoomID</label>
            <input type="text" id="joomId" name="joomId" placeholder="Joom出品後の製品ID（任意）">
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
            <label for="sellingPrice">販売価格</label>
            <input type="number" id="sellingPrice" name="sellingPrice" placeholder="150000" min="0">
            <div class="error-message" id="sellingPriceError">販売価格は仕入れ価格より高く設定してください</div>
          </div>
        </div>
        
        <div class="form-row">
          <div class="form-group">
            <label for="weight">重量 <span class="required">*</span></label>
            <input type="number" id="weight" name="weight" placeholder="187" min="0" required>
            <div class="error-message" id="weightError">重量は必須です（単位：g）</div>
          </div>
          
          <div class="form-group">
            <label for="sellingPriceUsd">販売価格（USD）</label>
            <input type="number" id="sellingPriceUsd" name="sellingPriceUsd" placeholder="1000" min="0" step="0.01">
          </div>
        </div>
        
        <!-- 容積重量計算用寸法セクション -->
        <div class="section-title">📏 容積重量計算用寸法</div>
        <div class="form-row">
          <div class="form-group">
            <label for="heightCm">高さ (cm)</label>
            <input type="number" id="heightCm" name="heightCm" placeholder="10" min="0" step="0.1">
            <div class="error-message" id="heightCmError">高さを入力してください（cm）</div>
          </div>
          
          <div class="form-group">
            <label for="lengthCm">長さ (cm)</label>
            <input type="number" id="lengthCm" name="lengthCm" placeholder="20" min="0" step="0.1">
            <div class="error-message" id="lengthCmError">長さを入力してください（cm）</div>
          </div>
        </div>
        
        <div class="form-row">
          <div class="form-group">
            <label for="widthCm">幅 (cm)</label>
            <input type="number" id="widthCm" name="widthCm" placeholder="15" min="0" step="0.1">
            <div class="error-message" id="widthCmError">幅を入力してください（cm）</div>
          </div>
        </div>
        
        <!-- Joom対応フィールドセクション -->
        
        <div class="form-group full-width">
          <label for="description">商品説明</label>
          <textarea id="description" name="description" placeholder="商品の詳細説明を入力してください（Joom出品用）..."></textarea>
        </div>
        
        <div class="form-row">
          <div class="form-group">
            <label for="mainImageUrl">メイン画像URL</label>
            <input type="url" id="mainImageUrl" name="mainImageUrl" placeholder="https://example.com/images/product.jpg">
            <div class="error-message" id="mainImageUrlError">有効なURLを入力してください</div>
          </div>
          
          <div class="form-group">
            <label for="currency">通貨</label>
            <select id="currency" name="currency">
              <option value="JPY">JPY（日本円）</option>
              <option value="USD">USD（米ドル）</option>
              <option value="EUR">EUR（ユーロ）</option>
            </select>
          </div>
        </div>
        
        <div class="form-row">
          <div class="form-group">
            <label for="shippingPrice">配送価格</label>
            <input type="number" id="shippingPrice" name="shippingPrice" placeholder="0" min="0" value="0">
            <div class="error-message" id="shippingPriceError">配送価格を入力してください（円）</div>
          </div>
          
          <div class="form-group">
            <label for="stockQuantity">在庫数量</label>
            <input type="number" id="stockQuantity" name="stockQuantity" placeholder="1" min="0" value="1">
            <div class="error-message" id="stockQuantityError">在庫数量を入力してください</div>
          </div>
        </div>
        
        <!-- 在庫情報セクション -->
        <div class="section-title">📦 在庫情報</div>
        
        <div class="form-row">
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
          
          <div class="form-group">
            <label for="category">商品カテゴリー <span class="required">*</span></label>
            <select id="category" name="category" required>
              <option value="">カテゴリーを選択してください</option>
              <? for (var i = 0; i < categoryData.length; i++) { ?>
                <option value="<?= categoryData[i] ?>"><?= categoryData[i] ?></option>
              <? } ?>
            </select>
            <div class="error-message" id="categoryError">商品カテゴリーは必須です</div>
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
    
    // サーバーサイドから取得したカテゴリーデータ
    var categoryData = <?= JSON.stringify(categoryData) ?>;
    
    // データが文字列の場合は解析する
    if (typeof categoryData === 'string') {
      try {
        categoryData = JSON.parse(categoryData);
      } catch (e) {
        console.error('JSON解析エラー:', e);
        categoryData = ['一般', '家電', 'アパレル', '雑貨', 'ホビー'];
      }
    }
    
    // 処理状態管理
    var isProcessing = false;
    var validationBannerHideTimeout = null;
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
    
    // 在庫ステータス変更時の処理
    document.getElementById('stockStatus').addEventListener('change', updateStockQuantity);
    
    
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
      var sellingPrice = document.getElementById('sellingPrice').value ? parseFloat(document.getElementById('sellingPrice').value) : null;
      
      // 利益計算（手数料は考慮しない）
      var profit = (sellingPrice || 0) - purchasePrice;
      
      // プレビューを表示
      document.getElementById('previewSellingPrice').textContent = '¥' + (sellingPrice || 0).toLocaleString();
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
      if ((sellingPrice && sellingPrice > 0) || purchasePrice > 0) {
        document.getElementById('pricePreview').style.display = 'block';
      }
    }
    
    // 在庫ステータス変更時の在庫数量自動更新
    function updateStockQuantity() {
      var stockStatus = document.getElementById('stockStatus').value;
      var stockQuantityField = document.getElementById('stockQuantity');
      
      if (stockStatus === '在庫あり') {
        stockQuantityField.value = 1;
      } else if (stockStatus === '売り切れ') {
        stockQuantityField.value = 0;
      } else if (stockStatus === '予約受付中') {
        stockQuantityField.value = 0;
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
          joomId: document.getElementById('joomId').value || '',
          sku: document.getElementById('sku').value,
          asin: document.getElementById('asin').value,
          supplier: document.getElementById('supplier').value,
          supplierUrl: document.getElementById('supplierUrl').value,
          purchasePrice: parseFloat(document.getElementById('purchasePrice').value),
          sellingPrice: document.getElementById('sellingPrice').value ? parseFloat(document.getElementById('sellingPrice').value) : null,
          sellingPriceUsd: document.getElementById('sellingPriceUsd').value,
          weight: parseInt(document.getElementById('weight').value),
          // Joom対応フィールド
          description: document.getElementById('description').value,
          mainImageUrl: document.getElementById('mainImageUrl').value,
          currency: document.getElementById('currency').value,
          shippingPrice: parseFloat(document.getElementById('shippingPrice').value) || 0,
          stockQuantity: parseInt(document.getElementById('stockQuantity').value) || 1,
          // 容積重量計算用寸法フィールド
          heightCm: parseFloat(document.getElementById('heightCm').value) || 0,
          lengthCm: parseFloat(document.getElementById('lengthCm').value) || 0,
          widthCm: parseFloat(document.getElementById('widthCm').value) || 0,
          // 既存フィールド
          stockStatus: document.getElementById('stockStatus').value,
          category: document.getElementById('category').value,
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
            
            // 0.5秒後にフォームを閉じる（成功メッセージを短く表示）
            setTimeout(function() {
              closeForm();
            }, 500);
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
    
    // フォームバリデーション（失敗時は固定バナー表示＋最初のエラー箇所へスクロール）
    function validateForm() {
      var isValid = true;
      var firstInvalidField = null;
      
      // バリデーションエラーバナーを非表示（再検証時に前回の表示をクリア）
      if (validationBannerHideTimeout) {
        clearTimeout(validationBannerHideTimeout);
        validationBannerHideTimeout = null;
      }
      document.getElementById('validationErrorBanner').style.display = 'none';
      
      // エラーメッセージをクリア
      document.querySelectorAll('.error-message').forEach(function(error) {
        error.style.display = 'none';
      });
      
      // 必須項目のチェック（商品IDは自動生成のため除外）
      var requiredFields = ['productName', 'supplier', 'purchasePrice', 'weight', 'stockStatus', 'category'];
      
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
          if (!firstInvalidField) firstInvalidField = field;
          isValid = false;
        }
      });
      
      // URLの形式チェック（Amazonでない場合のみ、または値が入力されている場合）
      var urlField = document.getElementById('supplierUrl');
      var urlError = document.getElementById('supplierUrlError');
      
      if (urlField.value && !isValidUrl(urlField.value)) {
        urlError.style.display = 'block';
        if (!firstInvalidField) firstInvalidField = urlField;
        isValid = false;
      }
      
      // 価格の妥当性チェック
      var purchasePrice = parseFloat(document.getElementById('purchasePrice').value);
      var sellingPriceField = document.getElementById('sellingPrice');
      var sellingPrice = sellingPriceField.value ? parseFloat(sellingPriceField.value) : null;
      
      if (sellingPrice && sellingPrice <= purchasePrice) {
        document.getElementById('sellingPriceError').textContent = '販売価格は仕入れ価格より高く設定してください';
        document.getElementById('sellingPriceError').style.display = 'block';
        if (!firstInvalidField) firstInvalidField = sellingPriceField;
        isValid = false;
      }
      
      // バリデーション失敗時：画面上部に固定バナー表示＋最初のエラー箇所へスクロール＋5秒後に非表示
      if (!isValid && firstInvalidField) {
        document.getElementById('validationErrorBanner').style.display = 'block';
        firstInvalidField.scrollIntoView({ behavior: 'smooth', block: 'center' });
        firstInvalidField.focus();
        validationBannerHideTimeout = setTimeout(function() {
          document.getElementById('validationErrorBanner').style.display = 'none';
          validationBannerHideTimeout = null;
        }, 5000);
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
      joomId: '',
      sku: 'TEST-001',
      asin: 'B0TEST001',
      supplier: 'Amazon',
      supplierUrl: 'https://amazon.co.jp/test',
      purchasePrice: 1000,
      sellingPrice: 1500,
      weight: 100,
      heightCm: 7.6,
      lengthCm: 14.7,
      widthCm: 0.8,
      stockStatus: '在庫あり',
      notes: 'これは容積重量計算機能のテスト用商品です。寸法データが正常に保存され、利益計算シートで自動読み込みされることを確認します。'
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
      const notesValue = inventorySheet.getRange(lastRow, COLUMN_INDEXES.INVENTORY.NOTES).getValue(); // 備考列
      
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
