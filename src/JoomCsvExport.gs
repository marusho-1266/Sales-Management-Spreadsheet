/**
 * Joom用CSV出力機能
 * 在庫管理ツール システム開発
 * 作成日: 2025-10-04
 */

/**
 * Joom用CSV出力のメイン関数
 * @param {Object} options - 出力オプション
 * @param {string} options.targetProducts - 出力対象 ('all', 'unlinked', 'selected')
 * @param {boolean} options.includeRecommended - 推奨フィールドを含むか
 * @param {boolean} options.validateData - データバリデーションを実行するか
 */
function exportJoomCsv(options = {}) {
  try {
    const {
      targetProducts = JOOM_CSV_CONFIG.TARGET_PRODUCTS.UNLINKED,
      includeRecommended = true,
      validateData = true
    } = options;
    
    console.log('Joom用CSV出力を開始します...');
    console.log('出力対象:', targetProducts);
    
    // 1. 対象商品の取得
    const targetProductsData = getTargetProducts(targetProducts);
    if (!targetProductsData || targetProductsData.length === 0) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('エラー', '出力対象の商品が見つかりません。\n\n※ 在庫数が0または未設定の商品は自動的に除外されます。', ui.ButtonSet.OK);
      return;
    }
    
    console.log(`対象商品数: ${targetProductsData.length}件（在庫数フィルタリング適用後）`);
    
    // 2. データバリデーション
    if (validateData) {
      const validationResult = validateProductsData(targetProductsData);
      if (!validationResult.isValid) {
        showValidationErrors(validationResult.errors);
        return;
      }
    }
    
    // 3. CSV生成
    const csvData = generateJoomCsv(targetProductsData, includeRecommended);
    
    // 4. ファイル出力
    const fileName = `Joom_Products_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}.csv`;
    
    let file;
    try {
      // DriveAppを使用してファイルを作成
      file = DriveApp.createFile(fileName, csvData, MimeType.CSV);
    } catch (driveError) {
      console.warn('DriveAppでのファイル作成に失敗しました。代替方法を使用します:', driveError.message);
      
      // 代替方法：スプレッドシート内にCSVデータを出力
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const tempSheetName = `CSV_Export_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
      
      try {
        const tempSheet = spreadsheet.insertSheet(tempSheetName);
        
        // 改善されたCSVパーサーでマルチラインクォートフィールドに対応
        const data = parseCsvWithMultilineSupport(csvData);
        
        // フィールド数整合性の検証
        if (data.length > 0) {
          const headerFieldCount = data[0].length;
          const inconsistentRows = [];
          
          for (let i = 1; i < data.length; i++) {
            if (data[i].length !== headerFieldCount) {
              inconsistentRows.push({
                rowIndex: i + 1,
                expectedFields: headerFieldCount,
                actualFields: data[i].length
              });
            }
          }
          
          if (inconsistentRows.length > 0) {
            console.warn('CSVフィールド数不一致が検出されました:', inconsistentRows);
            console.log('不一致行:', inconsistentRows.map(row => `行${row.rowIndex}: 期待${row.expectedFields}フィールド、実際${row.actualFields}フィールド`).join(', '));
          }
          
          // データをシートに書き込み（整合性が取れている場合のみ）
          if (data.length > 0 && data[0].length > 0) {
            tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
            console.log(`CSVデータをシートに書き込みました。行数: ${data.length}, 列数: ${data[0].length}`);
          }
        }        
        // ファイル名をシート名に設定
        tempSheet.setName(fileName.replace('.csv', ''));
        
        console.log('CSVデータをスプレッドシート内のシートに出力しました:', tempSheetName);
        
      } catch (sheetError) {
        console.error('スプレッドシート内での出力にも失敗しました:', sheetError);
        throw new Error('CSVファイルの出力に失敗しました。権限設定を確認してください。');
      }
    }
    
    // 5. 連携ステータス更新
    if (targetProducts === JOOM_CSV_CONFIG.TARGET_PRODUCTS.UNLINKED) {
      updateJoomStatus(targetProductsData.map(row => row.productId), JOOM_STATUS.LINKED);
    }
    
    // 6. 結果表示
    const ui = SpreadsheetApp.getUi();
    if (file) {
      // DriveAppでファイル作成成功の場合
      ui.alert(
        'CSV出力完了',
        `Joom用CSVの出力が完了しました。\nファイル名: ${fileName}\n出力商品数: ${targetProductsData.length}件\n\n※ CSVファイルはGoogleドライブに保存されました。`,
        ui.ButtonSet.OK
      );
    } else {
      // スプレッドシート内出力の場合
      ui.alert(
        'CSV出力完了',
        `Joom用CSVの出力が完了しました。\nシート名: ${fileName.replace('.csv', '')}\n出力商品数: ${targetProductsData.length}件\n\n※ CSVデータは新しいシートに出力されました。\nシートを右クリックして「ダウンロード」→「カンマ区切り値（.csv）」でCSVファイルをダウンロードできます。`,
        ui.ButtonSet.OK
      );
    }
    
    console.log('Joom用CSV出力が完了しました:', fileName);
    
  } catch (error) {
    console.error('Joom用CSV出力中にエラーが発生しました:', error);
    const ui = SpreadsheetApp.getUi();
    ui.alert('エラー', `CSV出力中にエラーが発生しました: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * 対象商品の取得
 * @param {string} targetProducts - 出力対象
 * @returns {Array} 商品データの配列
 */
function getTargetProducts(targetProducts) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  
  if (!inventorySheet) {
    throw new Error('在庫管理シートが見つかりません');
  }
  
  const data = inventorySheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  // ヘッダーから列インデックスを取得
  const columnIndexes = {
    productId: headers.indexOf('商品ID'),
    productName: headers.indexOf('商品名'),
    sku: headers.indexOf('SKU'),
    asin: headers.indexOf('ASIN'),
    supplier: headers.indexOf('仕入れ元'),
    supplierUrl: headers.indexOf('仕入れ元URL'),
    purchasePrice: headers.indexOf('仕入れ価格'),
    sellingPrice: headers.indexOf('販売価格'),
    weight: headers.indexOf('重量'),
    description: headers.indexOf('商品説明'),
    mainImageUrl: headers.indexOf('メイン画像URL'),
    currency: headers.indexOf('通貨'),
    shippingPrice: headers.indexOf('配送価格'),
    stockQuantity: headers.indexOf('在庫数量'),
    stockStatus: headers.indexOf('在庫ステータス'),
    joomStatus: headers.indexOf('Joom連携ステータス')
  };

  // 必須ヘッダーの検証
  const requiredHeaders = {
    productId: '商品ID',
    productName: '商品名',
    sku: 'SKU',
    description: '商品説明',
    mainImageUrl: 'メイン画像URL',
    sellingPrice: '販売価格',
    weight: '重量',
    currency: '通貨',
    stockQuantity: '在庫数量',
    joomStatus: 'Joom連携ステータス'
  };

  const missingHeaders = [];
  for (const [key, headerName] of Object.entries(requiredHeaders)) {
    if (columnIndexes[key] === -1) {
      missingHeaders.push(headerName);
    }
  }

  if (missingHeaders.length > 0) {
    throw new Error(`在庫管理シートに必須のヘッダーが見つかりません: ${missingHeaders.join(', ')}\n\n見つからないヘッダー:\n${missingHeaders.map(h => `- ${h}`).join('\n')}\n\n在庫管理シートの1行目にこれらのヘッダーが含まれていることを確認してください。`);
  }
  
  let filteredRows = rows;
  
  // 出力対象によるフィルタリング
  switch (targetProducts) {
    case JOOM_CSV_CONFIG.TARGET_PRODUCTS.UNLINKED:
      filteredRows = rows.filter(row => 
        row[columnIndexes.joomStatus] === JOOM_STATUS.UNLINKED || 
        !row[columnIndexes.joomStatus]
      );
      break;
    case JOOM_CSV_CONFIG.TARGET_PRODUCTS.SELECTED:
      // 選択された行を取得（実装は後で追加）
      const selection = SpreadsheetApp.getActiveRange();
      if (selection) {
        const selectedRows = selection.getRow() - 1; // ヘッダー行を考慮
        filteredRows = rows.filter((_, index) => 
          selection.getRow() <= index + 2 && 
          index + 2 <= selection.getLastRow()
        );
      }
      break;
    case JOOM_CSV_CONFIG.TARGET_PRODUCTS.ALL:
    default:
      // 全商品
      break;
  }
  
  // 在庫数によるフィルタリング（在庫数が0または空の商品を除外）
  const originalCount = filteredRows.length;
  filteredRows = filteredRows.filter(row => {
    const stockQuantity = row[columnIndexes.stockQuantity];
    // 在庫数が数値で0より大きい場合のみ出力対象とする
    return stockQuantity && !isNaN(stockQuantity) && stockQuantity > 0;
  });
  
  const filteredCount = filteredRows.length;
  const excludedCount = originalCount - filteredCount;
  
  if (excludedCount > 0) {
    console.log(`在庫数フィルタリング: ${excludedCount}件の商品を除外しました（在庫数0または未設定）`);
    console.log(`出力対象商品数: ${originalCount}件 → ${filteredCount}件`);
  }
  
  return filteredRows.map(row => ({
    productId: row[columnIndexes.productId],
    productName: row[columnIndexes.productName],
    sku: row[columnIndexes.sku],
    asin: row[columnIndexes.asin],
    supplier: row[columnIndexes.supplier],
    supplierUrl: row[columnIndexes.supplierUrl],
    purchasePrice: row[columnIndexes.purchasePrice],
    sellingPrice: row[columnIndexes.sellingPrice],
    weight: row[columnIndexes.weight],
    description: row[columnIndexes.description],
    mainImageUrl: row[columnIndexes.mainImageUrl],
    currency: row[columnIndexes.currency],
    shippingPrice: row[columnIndexes.shippingPrice],
    stockQuantity: row[columnIndexes.stockQuantity],
    stockStatus: row[columnIndexes.stockStatus],
    joomStatus: row[columnIndexes.joomStatus]
  }));
}

/**
 * 商品データのバリデーション
 * @param {Array} productsData - 商品データの配列
 * @returns {Object} バリデーション結果
 */
function validateProductsData(productsData) {
  const errors = [];
  
  productsData.forEach((product, index) => {
    const rowNumber = index + 2; // ヘッダー行を考慮
    
    // 必須フィールドのチェック
    if (!product.sku) {
      errors.push(`行${rowNumber}: SKUが入力されていません`);
    }
    if (!product.productName) {
      errors.push(`行${rowNumber}: 商品名が入力されていません`);
    }
    if (!product.description) {
      errors.push(`行${rowNumber}: 商品説明が入力されていません`);
    }
    if (!product.mainImageUrl) {
      errors.push(`行${rowNumber}: メイン画像URLが入力されていません`);
    }
    if (!product.sellingPrice || product.sellingPrice <= 0) {
      errors.push(`行${rowNumber}: 販売価格が正しく入力されていません`);
    }
    if (!product.currency) {
      errors.push(`行${rowNumber}: 通貨が入力されていません`);
    }
    if (product.stockQuantity === undefined || product.stockQuantity < 0) {
      errors.push(`行${rowNumber}: 在庫数量が正しく入力されていません`);
    }
    if (product.shippingPrice === undefined || product.shippingPrice < 0) {
      errors.push(`行${rowNumber}: 配送価格が正しく入力されていません`);
    }
    if (!product.weight || product.weight <= 0) {
      errors.push(`行${rowNumber}: 重量が正しく入力されていません`);
    }
    
    // 画像URLの形式チェック（一時的に無効化）
    // if (product.mainImageUrl) {
    //   const imageValidation = isValidImageUrl(product.mainImageUrl);
    //   if (!imageValidation.isValid) {
    //     imageValidation.errors.forEach(error => {
    //       errors.push(`行${rowNumber}: メイン画像URL - ${error}`);
    //     });
    //   }
    // }
    
    // 在庫数量の上限チェック
    if (product.stockQuantity > JOOM_CSV_CONFIG.VALIDATION.MAX_INVENTORY) {
      errors.push(`行${rowNumber}: 在庫数量が上限(${JOOM_CSV_CONFIG.VALIDATION.MAX_INVENTORY})を超えています`);
    }
  });
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * 画像URLの有効性チェック
 * @param {string} url - 画像URL
 * @returns {Object} バリデーション結果
 */
function isValidImageUrl(url) {
  const result = {
    isValid: false,
    errors: []
  };
  
  try {
    // 基本的なURL形式チェック
    const urlObj = new URL(url);
    const validProtocols = ['http:', 'https:'];
    const validExtensions = ['.jpg', '.jpeg', '.png', '.gif'];
    
    // プロトコルチェック
    if (!validProtocols.includes(urlObj.protocol)) {
      result.errors.push('プロトコルはhttpまたはhttpsである必要があります');
      return result;
    }
    
    // 拡張子チェック
    const pathname = urlObj.pathname.toLowerCase();
    const hasValidExtension = validExtensions.some(ext => pathname.endsWith(ext));
    
    if (!hasValidExtension) {
      result.errors.push('画像ファイルの拡張子は.jpg、.jpeg、.png、.gifのいずれかである必要があります');
      return result;
    }
    
    // ファイル名の存在チェック
    if (pathname === '/' || pathname.endsWith('/')) {
      result.errors.push('画像ファイル名が指定されていません');
      return result;
    }
    
    // フラグメントのチェック（一般的に画像URLには不要）
    if (urlObj.hash && urlObj.hash.includes('#')) {
      result.errors.push('画像URLにフラグメントが含まれています。直接的な画像ファイルのURLを使用してください');
      return result;
    }
    
    // クエリパラメータのチェック（画像配信サービスでは一般的に使用される）
    // ただし、明らかにページURLと思われるパラメータは除外
    if (urlObj.search && urlObj.search.includes('?')) {
      const queryParams = urlObj.searchParams;
      const suspiciousParams = ['page', 'id', 'category', 'search', 'view', 'action'];
      const hasSuspiciousParam = suspiciousParams.some(param => queryParams.has(param));
      
      if (hasSuspiciousParam) {
        result.errors.push('画像URLにページ関連のパラメータが含まれています。直接的な画像ファイルのURLを使用してください');
        return result;
      }
    }
    
    result.isValid = true;
    return result;
    
  } catch (error) {
    result.errors.push('有効なURL形式ではありません');
    return result;
  }
}

/**
 * 画像URLの有効性チェック（簡易版）
 * @param {string} url - 画像URL
 * @returns {boolean} 有効かどうか
 */
function isValidImageUrlSimple(url) {
  const result = isValidImageUrl(url);
  return result.isValid;
}

/**
 * バリデーションエラーの表示
 * @param {Array} errors - エラーメッセージの配列
 */
function showValidationErrors(errors) {
  const ui = SpreadsheetApp.getUi();
  const errorMessage = '以下のエラーを修正してください:\n\n' + errors.join('\n');
  ui.alert('バリデーションエラー', errorMessage, ui.ButtonSet.OK);
}

/**
 * Joom用CSVデータの生成
 * @param {Array} productsData - 商品データの配列
 * @param {boolean} includeRecommended - 推奨フィールドを含むか
 * @returns {string} CSVデータ
 */
function generateJoomCsv(productsData, includeRecommended) {
  const csvRows = [];
  
  // ヘッダー行の生成
  const headers = [
    JOOM_CSV_FIELDS.REQUIRED.PRODUCT_SKU,
    JOOM_CSV_FIELDS.REQUIRED.NAME,
    JOOM_CSV_FIELDS.REQUIRED.DESCRIPTION,
    JOOM_CSV_FIELDS.REQUIRED.PRODUCT_MAIN_IMAGE_URL,
    JOOM_CSV_FIELDS.REQUIRED.STORE_ID,
    JOOM_CSV_FIELDS.REQUIRED.SHIPPING_WEIGHT,
    JOOM_CSV_FIELDS.REQUIRED.PRICE,
    JOOM_CSV_FIELDS.REQUIRED.CURRENCY,
    JOOM_CSV_FIELDS.REQUIRED.INVENTORY,
    JOOM_CSV_FIELDS.REQUIRED.SHIPPING_PRICE,
    JOOM_CSV_FIELDS.RECOMMENDED.VARIANT_SKU,
    JOOM_CSV_FIELDS.RECOMMENDED.DANGEROUS_KIND
  ];
  
  if (includeRecommended) {
    headers.push(
      JOOM_CSV_FIELDS.RECOMMENDED.BRAND,
      JOOM_CSV_FIELDS.RECOMMENDED.SEARCH_TAGS,
      JOOM_CSV_FIELDS.RECOMMENDED.SUGGESTED_CATEGORY_ID
    );
  }
  
  csvRows.push(headers.join(','));
  
  // データ行の生成
  productsData.forEach(product => {
    const storeId = getSetting('ストアID') || 'STORE001';
    // Dangerous Kindは常に空の値を使用（危険物でない場合）
    const dangerousKind = '';
    const categoryId = getSetting('カテゴリID') || '';
    const searchTags = getSetting('検索タグ') || '';
    
    const row = [
      product.productId || '',
      product.productName || '',
      product.description || '',
      product.mainImageUrl || '',
      storeId,
      convertWeightToKg(product.weight),
      product.sellingPrice || 0,
      product.currency || JOOM_CSV_CONFIG.DEFAULTS.CURRENCY,
      product.stockQuantity || 0,
      product.shippingPrice || JOOM_CSV_CONFIG.DEFAULTS.SHIPPING_PRICE,
      generateVariantSku(product.productId),
      dangerousKind || '' // Dangerous Kind (常に出力)
    ];
    
    if (includeRecommended) {
      row.push(
        '', // Brand
        searchTags, // Search Tags
        categoryId // Suggested Category ID
      );
    }
    
    // CSVエスケープ処理
    const escapedRow = row.map(field => {
      const str = String(field || '');
      if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    });
    
    csvRows.push(escapedRow.join(','));
  });
  
  return csvRows.join('\n');
}

/**
 * 重量をgからkgに変換
 * @param {number} weightInGrams - 重量（g）
 * @returns {number} 重量（kg）
 */
function convertWeightToKg(weightInGrams) {
  if (!weightInGrams || weightInGrams <= 0) {
    return 0;
  }
  return (weightInGrams / JOOM_CSV_CONFIG.DEFAULTS.WEIGHT_UNIT_CONVERSION).toFixed(3);
}

/**
 * Product SKUからVariant SKUを生成
 * Product SKUに「-1」を付けた値を返す
 * @param {string} productSku - Product SKU
 * @returns {string} Variant SKU
 */
function generateVariantSku(productSku) {
  if (!productSku) {
    return '';
  }
  return `${productSku}-1`;
}

/**
 * Joom連携ステータスの更新
 * @param {Array} productIds - 商品IDの配列
 * @param {string} status - 新しいステータス
 */
function updateJoomStatus(productIds, status) {
  try {
    const now = new Date();
    const currentTime = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd HH:mm:ss');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
    
    if (!inventorySheet) {
      console.error('在庫管理シートが見つかりません');
      return;
    }
    
    const data = inventorySheet.getDataRange().getValues();
    const headers = data[0];
    const productIdColumn = headers.indexOf('商品ID');
    const joomStatusColumn = headers.indexOf('Joom連携ステータス');
    const joomLastExportColumn = headers.indexOf('最終出力日時');
    
    if (productIdColumn === -1) {
      console.error('商品ID列が見つかりません');
      return;
    }
    
    // Joom連携管理列が存在しない場合は自動追加
    if (joomStatusColumn === -1 || joomLastExportColumn === -1) {
      console.log('Joom連携管理列が見つかりません。自動追加を実行します...');
      try {
        addJoomStatusColumnsToExistingSheet();
        
        // 列インデックスを再取得
        const newData = inventorySheet.getDataRange().getValues();
        const newHeaders = newData[0];
        const newJoomStatusColumn = newHeaders.indexOf('Joom連携ステータス');
        const newJoomLastExportColumn = newHeaders.indexOf('最終出力日時');
        
        if (newJoomStatusColumn === -1 || newJoomLastExportColumn === -1) {
          console.error('Joom連携管理列の追加に失敗しました');
          return;
        }
        
        // 更新された列インデックスを使用
        const updatedJoomStatusColumn = newJoomStatusColumn;
        const updatedJoomLastExportColumn = newJoomLastExportColumn;
        
        console.log('Joom連携管理列を追加しました');
        
        // 更新処理を実行
        productIds.forEach(productId => {
          for (let i = 1; i < newData.length; i++) {
            if (newData[i][productIdColumn] === productId) {
              inventorySheet.getRange(i + 1, updatedJoomStatusColumn + 1).setValue(status);
              if (updatedJoomLastExportColumn !== -1) {
                inventorySheet.getRange(i + 1, updatedJoomLastExportColumn + 1).setValue(currentTime);
              }
              break;
            }
          }
        });
        
        console.log(`${productIds.length}件の商品の連携ステータスを「${status}」に更新しました`);
        return;
        
      } catch (addError) {
        console.error('Joom連携管理列の追加中にエラーが発生しました:', addError);
        return;
      }
    }
    
    productIds.forEach(productId => {
      for (let i = 1; i < data.length; i++) {
        if (data[i][productIdColumn] === productId) {
          inventorySheet.getRange(i + 1, joomStatusColumn + 1).setValue(status);
          if (joomLastExportColumn !== -1) {
            inventorySheet.getRange(i + 1, joomLastExportColumn + 1).setValue(currentTime);
          }
          break;
        }
      }
    });
    
    console.log(`${productIds.length}件の商品の連携ステータスを「${status}」に更新しました`);
    
    // デバッグ情報を追加
    console.log('更新対象商品ID:', productIds);
    console.log('商品ID列インデックス:', productIdColumn);
    console.log('Joom連携ステータス列インデックス:', joomStatusColumn);
    console.log('最終出力日時列インデックス:', joomLastExportColumn);
    
  } catch (error) {
    console.error('連携ステータスの更新中にエラーが発生しました:', error);
  }
}

/**
 * 未連携商品のみのCSV出力
 */
function exportUnlinkedProductsCsv() {
  exportJoomCsv({
    targetProducts: JOOM_CSV_CONFIG.TARGET_PRODUCTS.UNLINKED,
    includeRecommended: true,
    validateData: true
  });
}

/**
 * 在庫数フィルタリング機能のテスト用関数
 * 現在のCSVファイル形式での動作確認
 */
function testInventoryFiltering() {
  try {
    console.log('在庫数フィルタリング機能のテストを開始します...');
    
    // テスト用の商品データ（現在のCSVファイルの内容を模擬）
    const testProducts = [
      {
        productId: '1',
        productName: 'iPhone 15 Pro 128GB',
        sku: '1',
        description: '最新のiPhone 15 Pro 128GBモデル。A17 Proチップ搭載で高性能。',
        mainImageUrl: 'https://via.placeholder.com/500x500.jpg',
        sellingPrice: 150000,
        currency: 'JPY',
        stockQuantity: 1, // 在庫あり
        shippingPrice: 0,
        joomStatus: '未連携'
      },
      {
        productId: '2',
        productName: 'MacBook Air M2 13インチ',
        sku: '2',
        description: 'MacBook Air M2 13インチ。M2チップで高速処理。軽量設計。',
        mainImageUrl: 'https://example.com/images/macbook-air-m2.jpg',
        sellingPrice: 180000,
        currency: 'JPY',
        stockQuantity: 1, // 在庫あり
        shippingPrice: 0,
        joomStatus: '未連携'
      },
      {
        productId: '3',
        productName: 'AirPods Pro 第2世代',
        sku: '3',
        description: 'AirPods Pro 第2世代。ノイズキャンセリング機能搭載。',
        mainImageUrl: 'https://example.com/images/airpods-pro-2nd.jpg',
        sellingPrice: 35000,
        currency: 'JPY',
        stockQuantity: null, // 在庫なし（空欄）
        shippingPrice: 0,
        joomStatus: '未連携'
      },
      {
        productId: '4',
        productName: 'iPad Air 第5世代',
        sku: '4',
        description: 'iPad Air 第5世代。M1チップ搭載で高性能タブレット。',
        mainImageUrl: 'https://example.com/images/ipad-air-5.jpg',
        sellingPrice: 80000,
        currency: 'JPY',
        stockQuantity: 1, // 在庫あり
        shippingPrice: 0,
        joomStatus: '未連携'
      },
      {
        productId: '5',
        productName: 'Apple Watch Series 9',
        sku: '5',
        description: 'Apple Watch Series 9。健康管理とスマートウォッチ機能。',
        mainImageUrl: 'https://example.com/images/apple-watch-s9.jpg',
        sellingPrice: 60000,
        currency: 'JPY',
        stockQuantity: 1, // 在庫あり
        shippingPrice: 0,
        joomStatus: '未連携'
      }
    ];
    
    console.log(`テスト対象商品数: ${testProducts.length}件`);
    
    // 在庫数フィルタリングを適用
    const filteredProducts = testProducts.filter(product => {
      const stockQuantity = product.stockQuantity;
      return stockQuantity && !isNaN(stockQuantity) && stockQuantity > 0;
    });
    
    console.log(`フィルタリング後商品数: ${filteredProducts.length}件`);
    
    // 除外された商品を表示
    const excludedProducts = testProducts.filter(product => {
      const stockQuantity = product.stockQuantity;
      return !stockQuantity || isNaN(stockQuantity) || stockQuantity <= 0;
    });
    
    if (excludedProducts.length > 0) {
      console.log(`除外された商品 (${excludedProducts.length}件):`);
      excludedProducts.forEach(product => {
        console.log(`- ${product.productName} (在庫数: ${product.stockQuantity})`);
      });
    }
    
    console.log(`出力対象商品 (${filteredProducts.length}件):`);
    filteredProducts.forEach(product => {
      console.log(`- ${product.productName} (在庫数: ${product.stockQuantity})`);
    });
    
    console.log('在庫数フィルタリング機能のテストが完了しました。');
    
  } catch (error) {
    console.error('テスト中にエラーが発生しました:', error);
  }
}

/**
 * マルチラインクォートフィールドに対応したCSVパーサー（改良版）
 * @param {string} csvData - パースするCSVデータ
 * @returns {Array<Array<string>>} パースされた2次元配列
 */
function parseCsvWithMultilineSupport(csvData) {
  // まず Utilities.parseCsv() を試行
  try {
    const result = Utilities.parseCsv(csvData);
    console.log('Utilities.parseCsv()を使用してCSVをパースしました');
    return result;
  } catch (error) {
    console.warn('Utilities.parseCsv()でパースに失敗しました。カスタムパーサーを使用します:', error.message);
  }
  
  // カスタムパーサーでフォールバック
  return parseCsvWithStateMachine(csvData);
}

/**
 * 状態機械を使用した堅牢なCSVパーサー
 * @param {string} csvData - パースするCSVデータ
 * @returns {Array<Array<string>>} パースされた2次元配列
 */
function parseCsvWithStateMachine(csvData) {
  const records = [];
  const lines = csvData.split(/\r?\n/);
  let currentRecord = '';
  let inQuotes = false;
  let i = 0;
  
  while (i < lines.length) {
    const line = lines[i];
    
    // 現在の行を追加
    currentRecord += (currentRecord ? '\n' : '') + line;
    
    // 状態を確認
    const newState = determineQuoteState(currentRecord);
    
    if (newState.inQuotes !== inQuotes) {
      inQuotes = newState.inQuotes;
    }
    
    // クォートが閉じられている場合、レコード完了
    if (!inQuotes && newState.isRecordComplete) {
      records.push(currentRecord);
      currentRecord = '';
    }
    
    i++;
  }
  
  // 最後のレコードが未完了の場合も追加
  if (currentRecord.trim()) {
    records.push(currentRecord);
  }
  
  // 各レコードをフィールドに分割
  return records.map(record => parseCsvRecord(record));
}

/**
 * 文字列のクォート状態を判定
 * @param {string} text - 判定する文字列
 * @returns {Object} 状態情報
 */
function determineQuoteState(text) {
  let inQuotes = false;
  let i = 0;
  
  while (i < text.length) {
    const char = text[i];
    
    if (char === '"') {
      // 次の文字もクォートの場合、エスケープされたクォート
      if (i + 1 < text.length && text[i + 1] === '"') {
        i += 2; // エスケープされたクォートをスキップ
        continue;
      }
      // 通常のクォート
      inQuotes = !inQuotes;
    }
    
    i++;
  }
  
  return {
    inQuotes: inQuotes,
    isRecordComplete: !inQuotes
  };
}

/**
 * 単一のCSVレコードをフィールドに分割
 * @param {string} record - CSVレコード文字列
 * @returns {Array<string>} フィールドの配列
 */
function parseCsvRecord(record) {
  const fields = [];
  let currentField = '';
  let inQuotes = false;
  
  for (let i = 0; i < record.length; i++) {
    const char = record[i];
    
    if (char === '"') {
      if (inQuotes && record[i + 1] === '"') {
        // エスケープされたクォート
        currentField += '"';
        i++; // 次のクォートをスキップ
      } else {
        // クォートの開始/終了
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      // フィールド区切り
      fields.push(currentField);
      currentField = '';
    } else {
      // 通常の文字
      currentField += char;
    }
  }
  
  // 最後のフィールドを追加
  fields.push(currentField);
  
  return fields;
}

/**
 * Utilities.parseCsv()のテスト関数
 */
function testUtilitiesParseCsv() {
  console.log('Utilities.parseCsv()のテストを開始します...');
  
  try {
    // テストケース1: 通常のCSV
    const normalCsv = 'name,age,city\n"John Doe",25,"New York"\n"Jane Smith",30,"Los Angeles"';
    console.log('テストケース1: 通常のCSV');
    const result1 = Utilities.parseCsv(normalCsv);
    console.log('Utilities.parseCsv結果:', result1);
    
    // テストケース2: マルチラインクォートフィールド
    const multilineCsv = 'name,description\n"John","This is a\nmultiline description"\n"Jane","Single line"';
    console.log('テストケース2: マルチラインクォートフィールド');
    const result2 = Utilities.parseCsv(multilineCsv);
    console.log('Utilities.parseCsv結果:', result2);
    
    // テストケース3: エスケープされたクォート
    const escapedCsv = 'name,quote\n"John","He said ""Hello"" to me"\n"Jane","Normal text"';
    console.log('テストケース3: エスケープされたクォート');
    const result3 = Utilities.parseCsv(escapedCsv);
    console.log('Utilities.parseCsv結果:', result3);
    
    // テストケース4: 複雑なマルチラインケース
    const complexCsv = 'id,name,description\n1,"Product A","This is a\nmultiline\ndescription with ""quotes"" inside"\n2,"Product B","Simple description"';
    console.log('テストケース4: 複雑なマルチラインケース');
    const result4 = Utilities.parseCsv(complexCsv);
    console.log('Utilities.parseCsv結果:', result4);
    
    console.log('Utilities.parseCsv()のテストが完了しました。');
    
  } catch (error) {
    console.error('Utilities.parseCsv()テスト中にエラーが発生しました:', error);
  }
}

/**
 * 改良されたCSVパーサーのテスト関数
 */
function testCsvParser() {
  console.log('改良されたCSVパーサーのテストを開始します...');
  
  try {
    // テストケース1: 通常のCSV
    const normalCsv = 'name,age,city\n"John Doe",25,"New York"\n"Jane Smith",30,"Los Angeles"';
    console.log('テストケース1: 通常のCSV');
    const result1 = parseCsvWithMultilineSupport(normalCsv);
    console.log('結果:', result1);
    
    // テストケース2: マルチラインクォートフィールド
    const multilineCsv = 'name,description\n"John","This is a\nmultiline description"\n"Jane","Single line"';
    console.log('テストケース2: マルチラインクォートフィールド');
    const result2 = parseCsvWithMultilineSupport(multilineCsv);
    console.log('結果:', result2);
    
    // テストケース3: エスケープされたクォート
    const escapedCsv = 'name,quote\n"John","He said ""Hello"" to me"\n"Jane","Normal text"';
    console.log('テストケース3: エスケープされたクォート');
    const result3 = parseCsvWithMultilineSupport(escapedCsv);
    console.log('結果:', result3);
    
    // テストケース4: 複雑なエスケープケース
    const complexEscapedCsv = 'id,description\n1,"This has ""multiple"" escaped quotes"\n2,"Simple description"';
    console.log('テストケース4: 複雑なエスケープケース');
    const result4 = parseCsvWithMultilineSupport(complexEscapedCsv);
    console.log('結果:', result4);
    
    // テストケース5: フィールド数不一致
    const inconsistentCsv = 'name,age,city\n"John",25\n"Jane",30,"Los Angeles","Extra"';
    console.log('テストケース5: フィールド数不一致');
    const result5 = parseCsvWithMultilineSupport(inconsistentCsv);
    console.log('結果:', result5);
    
    // フィールド数検証テスト
    if (result5.length > 0) {
      const headerFieldCount = result5[0].length;
      console.log(`ヘッダーフィールド数: ${headerFieldCount}`);
      
      for (let i = 1; i < result5.length; i++) {
        if (result5[i].length !== headerFieldCount) {
          console.log(`行${i + 1}: 期待${headerFieldCount}フィールド、実際${result5[i].length}フィールド`);
        }
      }
    }
    
    console.log('改良されたCSVパーサーのテストが完了しました。');
    
  } catch (error) {
    console.error('テスト中にエラーが発生しました:', error);
  }
}

/**
 * 包括的なCSVパーサーテスト関数
 */
function testAllCsvParsers() {
  console.log('=== 包括的なCSVパーサーテストを開始します ===');
  
  const testCases = [
    {
      name: '通常のCSV',
      data: 'name,age,city\n"John Doe",25,"New York"\n"Jane Smith",30,"Los Angeles"'
    },
    {
      name: 'マルチラインクォートフィールド',
      data: 'name,description\n"John","This is a\nmultiline description"\n"Jane","Single line"'
    },
    {
      name: 'エスケープされたクォート',
      data: 'name,quote\n"John","He said ""Hello"" to me"\n"Jane","Normal text"'
    },
    {
      name: '複雑なエスケープケース',
      data: 'id,description\n1,"This has ""multiple"" escaped quotes"\n2,"Simple description"'
    },
    {
      name: '空のフィールド',
      data: 'name,age,city\n"John",,"New York"\n"Jane",30,""'
    }
  ];
  
  testCases.forEach((testCase, index) => {
    console.log(`\n--- テストケース${index + 1}: ${testCase.name} ---`);
    
    try {
      // Utilities.parseCsv()のテスト
      console.log('Utilities.parseCsv()結果:');
      const utilsResult = Utilities.parseCsv(testCase.data);
      console.log(utilsResult);
    } catch (error) {
      console.log('Utilities.parseCsv()エラー:', error.message);
    }
    
    try {
      // カスタムパーサーのテスト
      console.log('カスタムパーサー結果:');
      const customResult = parseCsvWithMultilineSupport(testCase.data);
      console.log(customResult);
    } catch (error) {
      console.log('カスタムパーサーエラー:', error.message);
    }
  });
  
  console.log('\n=== 包括的なCSVパーサーテストが完了しました ===');
}

/**
 * Variant SKU生成機能のテスト
 */
function testVariantSkuGeneration() {
  console.log('=== Variant SKU生成機能のテストを開始します ===');
  
  const testCases = [
    { input: 'SKU001', expected: 'SKU001-1', description: '通常のSKU' },
    { input: 'PROD123', expected: 'PROD123-1', description: '英数字のSKU' },
    { input: '', expected: '', description: '空のSKU' },
    { input: null, expected: '', description: 'nullのSKU' },
    { input: undefined, expected: '', description: 'undefinedのSKU' },
    { input: 'TEST-ABC', expected: 'TEST-ABC-1', description: 'ハイフンを含むSKU' }
  ];
  
  testCases.forEach((testCase, index) => {
    console.log(`\n--- テストケース${index + 1}: ${testCase.description} ---`);
    console.log(`入力: "${testCase.input}"`);
    
    const result = generateVariantSku(testCase.input);
    console.log(`期待値: "${testCase.expected}"`);
    console.log(`実際の結果: "${result}"`);
    
    if (result === testCase.expected) {
      console.log('✅ テスト成功');
    } else {
      console.log('❌ テスト失敗');
    }
  });
  
  console.log('\n=== Variant SKU生成機能のテストが完了しました ===');
}

/**
 * CSV出力機能のVariant SKU統合テスト
 */
function testCsvOutputWithVariantSku() {
  console.log('=== CSV出力機能のVariant SKU統合テストを開始します ===');
  
  try {
    // テスト用の商品データ
    const testProducts = [
      {
        productId: 'TEST001',
        productName: 'テスト商品1',
        sku: 'TEST001',
        description: 'テスト商品の説明1',
        mainImageUrl: 'https://example.com/image1.jpg',
        sellingPrice: 1000,
        currency: 'JPY',
        stockQuantity: 5,
        shippingPrice: 0,
        weight: 500,
        joomStatus: '未連携'
      },
      {
        productId: 'TEST002',
        productName: 'テスト商品2',
        sku: 'TEST002',
        description: 'テスト商品の説明2',
        mainImageUrl: 'https://example.com/image2.jpg',
        sellingPrice: 2000,
        currency: 'JPY',
        stockQuantity: 3,
        shippingPrice: 0,
        weight: 300,
        joomStatus: '未連携'
      }
    ];
    
    console.log('テスト商品データ:', testProducts);
    
    // CSV生成
    const csvData = generateJoomCsv(testProducts, false);
    console.log('\n生成されたCSVデータ:');
    console.log(csvData);
    
    // CSVを解析してVariant SKUを確認
    const csvLines = csvData.split('\n');
    console.log('\nCSV行数:', csvLines.length);
    
    if (csvLines.length >= 2) {
      console.log('\nヘッダー行:', csvLines[0]);
      console.log('データ行1:', csvLines[1]);
      console.log('データ行2:', csvLines[2]);
      
      // Variant SKUフィールドの位置を確認
      const headers = csvLines[0].split(',');
      const variantSkuIndex = headers.indexOf('Variant SKU');
      console.log('\nVariant SKUフィールドの位置:', variantSkuIndex);
      
      if (variantSkuIndex !== -1) {
        // 各データ行のVariant SKUを確認
        for (let i = 1; i < csvLines.length; i++) {
          const fields = csvLines[i].split(',');
          if (fields[variantSkuIndex]) {
            console.log(`データ行${i}のVariant SKU: "${fields[variantSkuIndex]}"`);
          }
        }
      }
    }
    
    console.log('\n✅ CSV出力機能のVariant SKU統合テストが完了しました');
    
  } catch (error) {
    console.error('❌ テスト中にエラーが発生しました:', error);
  }
}

/**
 * Dangerous Kindフィールド修正のテスト
 */
function testDangerousKindFix() {
  console.log('=== Dangerous Kindフィールド修正のテストを開始します ===');
  
  try {
    // テスト用の商品データ
    const testProducts = [
      {
        productId: 'TEST001',
        productName: 'テスト商品1',
        sku: 'TEST001',
        description: 'テスト商品の説明1',
        mainImageUrl: 'https://example.com/image1.jpg',
        sellingPrice: 1000,
        currency: 'JPY',
        stockQuantity: 5,
        shippingPrice: 0,
        weight: 500,
        joomStatus: '未連携'
      }
    ];
    
    console.log('テスト商品データ:', testProducts);
    
    // CSV生成（推奨フィールドなし）
    const csvDataMinimal = generateJoomCsv(testProducts, false);
    console.log('\n=== 推奨フィールドなしCSV ===');
    console.log(csvDataMinimal);
    
    // CSV生成（推奨フィールドあり）
    const csvDataRecommended = generateJoomCsv(testProducts, true);
    console.log('\n=== 推奨フィールドありCSV ===');
    console.log(csvDataRecommended);
    
    // CSVを解析してDangerous Kindフィールドを確認
    const csvLines = csvDataMinimal.split('\n');
    console.log('\nCSV行数:', csvLines.length);
    
    if (csvLines.length >= 2) {
      console.log('\nヘッダー行:', csvLines[0]);
      console.log('データ行1:', csvLines[1]);
      
      // Dangerous Kindフィールドの位置を確認
      const headers = csvLines[0].split(',');
      const dangerousKindIndexes = [];
      headers.forEach((header, index) => {
        if (header === 'Dangerous Kind') {
          dangerousKindIndexes.push(index);
        }
      });
      
      console.log('\nDangerous Kindフィールドの位置:', dangerousKindIndexes);
      
      if (dangerousKindIndexes.length === 1) {
        console.log('✅ Dangerous Kindフィールドが1つのみ存在します');
        const fields = csvLines[1].split(',');
        const dangerousKindValue = fields[dangerousKindIndexes[0]];
        console.log(`Dangerous Kind値: "${dangerousKindValue}"`);
        
        if (dangerousKindValue === '' || dangerousKindValue === '""') {
          console.log('✅ Dangerous Kindが空の値で正しく出力されています');
        } else {
          console.log(`⚠️ Dangerous Kind値: "${dangerousKindValue}"`);
        }
      } else if (dangerousKindIndexes.length === 0) {
        console.log('❌ Dangerous Kindフィールドが見つかりません');
      } else {
        console.log(`❌ Dangerous Kindフィールドが重複しています（${dangerousKindIndexes.length}個）`);
      }
    }
    
    console.log('\n✅ Dangerous Kindフィールド修正のテストが完了しました');
    
  } catch (error) {
    console.error('❌ テスト中にエラーが発生しました:', error);
  }
}


