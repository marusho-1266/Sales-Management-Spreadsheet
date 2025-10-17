/**
 * ==========================================
 * Joom注文データ変換・統合機能
 * ==========================================
 * 
 * このファイルは、Joom APIから取得した注文データを
 * Google Sheetsの売上管理シート形式に変換し、
 * 在庫管理シートと連携する機能を提供します。
 */

/**
 * ==========================================
 * 1. メインデータ変換関数
 * ==========================================
 */

/**
 * Joom注文データを売上管理シート形式に変換
 * @param {Object} joomOrder - Joom APIから取得した注文データ
 * @returns {Array} 売上管理シートの1行分のデータ配列
 */
function transformJoomOrderToSalesRow(joomOrder) {
  try {
    console.log('注文データ変換開始:', joomOrder.id);
    
    // 商品情報を在庫管理シートから取得
    const productInfo = getProductInfoFromInventory(joomOrder.product.sku);
    
    // 日時の変換
    const orderDate = convertRfc3339ToJstDate(joomOrder.orderTimestamp);
    const registrationTime = new Date();
    
    // 価格情報の変換
    const priceInfo = transformPriceInfo(joomOrder.priceInfo, joomOrder.currency);
    
    // 配送情報の変換
    const shippingInfo = transformShippingAddress(joomOrder.shippingAddress);
    
    // 利益計算
    const netProfit = calculateNetProfit(
      priceInfo.sellingPrice,
      productInfo.purchasePrice,
      priceInfo.shippingCost,
      priceInfo.commission
    );
    
    // 売上管理シート形式の配列を作成（37列）
    const salesRow = [
      orderDate,                           // A列: 販売日
      joomOrder.id,                        // B列: 注文ID
      productInfo.productId,               // C列: 商品ID（在庫管理シートのSKU）
      productInfo.productName,             // D列: 商品名
      productInfo.sku,                     // E列: SKU
      productInfo.asin,                    // F列: ASIN
      joomOrder.quantity || 1,             // G列: 数量
      priceInfo.sellingPrice,              // H列: 販売価格
      productInfo.purchasePrice,           // I列: 仕入れ価格
      priceInfo.shippingCost,              // J列: 配送料
      netProfit,                           // K列: 純利益
      registrationTime,                    // L列: データ登録日時
      
      // 注文ステータス管理
      joomOrder.status || '',              // M列: 注文ステータス
      'synced',                            // N列: 連携状況
      registrationTime,                    // O列: 最終同期日時
      
      // 価格詳細
      priceInfo.commission,                // P列: 手数料
      priceInfo.vat,                       // Q列: VAT
      priceInfo.refundAmount,              // R列: 返金額
      priceInfo.customerGmv,               // S列: 購入者GMV
      
      // 顧客情報
      shippingInfo.name || '',             // T列: 顧客名
      shippingInfo.email || '',            // U列: メールアドレス
      shippingInfo.phone || '',            // V列: 電話番号
      shippingInfo.country || '',          // W列: 顧客の国
      shippingInfo.state || '',            // X列: 都道府県
      
      // 配送情報
      shippingInfo.country || '',          // Y列: 配送先国
      shippingInfo.state || '',            // Z列: 配送先都道府県
      shippingInfo.city || '',             // AA列: 配送先市区町村
      shippingInfo.address || '',          // AB列: 配送先住所
      shippingInfo.zipCode || '',          // AC列: 郵便番号
      shippingInfo.fullAddress || '',      // AD列: 完全住所
      
      // 出荷・配送管理
      joomOrder.shipment?.trackingNumber || '', // AE列: 追跡番号
      joomOrder.shipment?.provider || '',       // AF列: 配送業者
      convertRfc3339ToJstDateTime(joomOrder.shipment?.shippedTimestamp) || '', // AG列: 出荷日
      convertRfc3339ToJstDateTime(joomOrder.shipment?.fulfilledTimestamp) || '', // AH列: 履行日
      getDeliveryStatus(joomOrder.status),      // AI列: 配送状況
      
      // 連携管理
      '',                                  // AJ列: エラーメッセージ
      'Joom'                               // AK列: データソース
    ];
    
    console.log('注文データ変換完了:', joomOrder.id);
    return salesRow;
    
  } catch (error) {
    console.error('注文データ変換エラー:', error);
    throw new Error(`注文データ変換エラー (${joomOrder.id}): ${error.message}`);
  }
}

/**
 * ==========================================
 * 2. 商品情報取得機能
 * ==========================================
 */

/**
 * 在庫管理シートから商品情報を取得
 * @param {string} sku - 商品SKU（Joomの製品SKU）
 * @returns {Object} 商品情報オブジェクト
 */
function getProductInfoFromInventory(sku) {
  try {
    console.log('商品情報取得開始:', sku);
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
    
    if (!inventorySheet) {
      throw new Error('在庫管理シートが見つかりません');
    }
    
    const lastRow = inventorySheet.getLastRow();
    if (lastRow <= 1) {
      throw new Error('在庫管理シートにデータがありません');
    }
    
    // 商品IDでデータを検索（A列）
    const productIdRange = inventorySheet.getRange(2, COLUMN_INDEXES.INVENTORY.PRODUCT_ID, lastRow - 1, 1);
    const productIds = productIdRange.getValues();
    
    // SKUに一致する行を探す
    let targetRow = -1;
    for (let i = 0; i < productIds.length; i++) {
      if (productIds[i][0] === sku) {
        targetRow = i + 2; // ヘッダー行を考慮
        break;
      }
    }
    
    if (targetRow === -1) {
      console.warn(`商品情報が見つかりません: ${sku}。デフォルト値を使用します。`);
      return {
        productId: sku,
        productName: '商品名不明（在庫管理シートに未登録）',
        sku: sku,
        asin: '',
        purchasePrice: 0
      };
    }
    
    // 商品情報を取得
    const rowData = inventorySheet.getRange(targetRow, 1, 1, 20).getValues()[0];
    
    const productInfo = {
      productId: rowData[COLUMN_INDEXES.INVENTORY.PRODUCT_ID - 1],
      productName: rowData[COLUMN_INDEXES.INVENTORY.PRODUCT_NAME - 1],
      sku: rowData[COLUMN_INDEXES.INVENTORY.SKU - 1],
      asin: rowData[COLUMN_INDEXES.INVENTORY.ASIN - 1],
      purchasePrice: parseFloat(rowData[COLUMN_INDEXES.INVENTORY.PURCHASE_PRICE - 1]) || 0
    };
    
    console.log('商品情報取得完了:', productInfo);
    return productInfo;
    
  } catch (error) {
    console.error('商品情報取得エラー:', error);
    // エラー時はデフォルト値を返す
    return {
      productId: sku,
      productName: '商品名不明（取得エラー）',
      sku: sku,
      asin: '',
      purchasePrice: 0
    };
  }
}

/**
 * ==========================================
 * 3. 価格情報変換機能
 * ==========================================
 */

/**
 * 価格情報を変換
 * @param {Object} priceInfo - Joom APIの価格情報
 * @param {string} currency - 通貨コード
 * @returns {Object} 変換後の価格情報
 */
function transformPriceInfo(priceInfo, currency) {
  try {
    const exchangeRate = getExchangeRate(currency, 'JPY');
    
    return {
      sellingPrice: convertPriceToNumber(priceInfo.orderPrice, exchangeRate),
      shippingCost: convertPriceToNumber(priceInfo.shippingPrice || priceInfo.joomShippingPrice, exchangeRate),
      commission: convertPriceToNumber(priceInfo.commissionAmount, exchangeRate),
      vat: convertPriceToNumber(priceInfo.vat, exchangeRate),
      refundAmount: convertPriceToNumber(priceInfo.refundAmount, exchangeRate),
      customerGmv: convertPriceToNumber(priceInfo.buyerGmv || priceInfo.orderPrice, exchangeRate)
    };
    
  } catch (error) {
    console.error('価格情報変換エラー:', error);
    return {
      sellingPrice: 0,
      shippingCost: 0,
      commission: 0,
      vat: 0,
      refundAmount: 0,
      customerGmv: 0
    };
  }
}

/**
 * 価格文字列を数値に変換（通貨換算含む）
 * @param {string|number} priceString - 価格文字列または数値
 * @param {number} exchangeRate - 為替レート
 * @returns {number} 変換後の数値
 */
function convertPriceToNumber(priceString, exchangeRate = 1) {
  try {
    if (!priceString || priceString === '') {
      return 0;
    }
    
    // 既に数値の場合
    if (typeof priceString === 'number') {
      return Math.round(priceString * exchangeRate * 100) / 100;
    }
    
    // 文字列から数値に変換
    const numericValue = parseFloat(priceString);
    if (isNaN(numericValue)) {
      return 0;
    }
    
    // 為替換算して小数点2桁まで
    return Math.round(numericValue * exchangeRate * 100) / 100;
    
  } catch (error) {
    console.error('価格変換エラー:', priceString, error);
    return 0;
  }
}

/**
 * 為替レートを取得
 * @param {string} fromCurrency - 変換元通貨
 * @param {string} toCurrency - 変換先通貨
 * @returns {number} 為替レート
 */
function getExchangeRate(fromCurrency, toCurrency) {
  try {
    // 同じ通貨の場合
    if (fromCurrency === toCurrency) {
      return 1;
    }
    
    // JPYへの変換の場合、設定シートから取得
    if (toCurrency === 'JPY') {
      const settingKey = `為替レート ${fromCurrency}/JPY`;
      const rate = getSetting(settingKey);
      
      if (rate && !isNaN(parseFloat(rate))) {
        return parseFloat(rate);
      }
      
      // デフォルトレート（よく使われる通貨のみ）
      const defaultRates = {
        'USD': 150.0,
        'EUR': 160.0,
        'GBP': 180.0,
        'CNY': 20.0
      };
      
      return defaultRates[fromCurrency] || 1;
    }
    
    // その他の通貨変換は未対応
    console.warn(`未対応の通貨変換: ${fromCurrency} → ${toCurrency}`);
    return 1;
    
  } catch (error) {
    console.error('為替レート取得エラー:', error);
    return 1;
  }
}

/**
 * ==========================================
 * 4. 配送情報変換機能
 * ==========================================
 */

/**
 * 配送先住所情報を変換
 * @param {Object} shippingAddress - Joom APIの配送先住所
 * @returns {Object} 変換後の配送情報
 */
function transformShippingAddress(shippingAddress) {
  try {
    if (!shippingAddress) {
      return {
        name: '',
        email: '',
        phone: '',
        country: '',
        state: '',
        city: '',
        address: '',
        zipCode: '',
        fullAddress: ''
      };
    }
    
    // 住所文字列の構築
    const addressParts = [];
    if (shippingAddress.street || shippingAddress.streetAddress1) {
      addressParts.push(shippingAddress.street || shippingAddress.streetAddress1);
    }
    if (shippingAddress.streetAddress2) {
      addressParts.push(shippingAddress.streetAddress2);
    }
    if (shippingAddress.building) {
      addressParts.push(shippingAddress.building);
    }
    if (shippingAddress.flat) {
      addressParts.push(shippingAddress.flat);
    }
    
    const address = addressParts.join(' ');
    
    // 完全住所の構築
    const fullAddressParts = [
      shippingAddress.country,
      shippingAddress.state,
      shippingAddress.city,
      address,
      shippingAddress.zipCode
    ].filter(part => part && part !== '');
    
    const fullAddress = fullAddressParts.join(', ');
    
    return {
      name: shippingAddress.name || '',
      email: shippingAddress.email || '',
      phone: shippingAddress.phoneNumber || '',
      country: shippingAddress.country || '',
      state: shippingAddress.state || '',
      city: shippingAddress.city || '',
      address: address,
      zipCode: shippingAddress.zipCode || '',
      fullAddress: fullAddress
    };
    
  } catch (error) {
    console.error('配送情報変換エラー:', error);
    return {
      name: '',
      email: '',
      phone: '',
      country: '',
      state: '',
      city: '',
      address: '',
      zipCode: '',
      fullAddress: ''
    };
  }
}

/**
 * ==========================================
 * 5. 利益計算機能
 * ==========================================
 */

/**
 * 純利益を計算
 * @param {number} sellingPrice - 販売価格
 * @param {number} purchasePrice - 仕入れ価格
 * @param {number} shippingCost - 配送料
 * @param {number} commission - 手数料
 * @returns {number} 純利益
 */
function calculateNetProfit(sellingPrice, purchasePrice, shippingCost, commission) {
  try {
    const profit = sellingPrice - purchasePrice - shippingCost - commission;
    return Math.round(profit * 100) / 100;
  } catch (error) {
    console.error('利益計算エラー:', error);
    return 0;
  }
}

/**
 * ==========================================
 * 6. ユーティリティ関数
 * ==========================================
 */

/**
 * 配送状況を取得
 * @param {string} orderStatus - Joom注文ステータス
 * @returns {string} 配送状況
 */
function getDeliveryStatus(orderStatus) {
  const statusMap = {
    'approved': '承認済み',
    'fulfilledOnline': 'オンライン履行済み',
    'shipped': '出荷済み',
    'cancelled': 'キャンセル',
    'paidByJoomRefund': 'Joom返金済み',
    'refunded': '返金済み',
    'returnInitiated': '返品開始',
    'returnExpired': '返品期限切れ',
    'returnArrived': '返品到着',
    'returnCompleted': '返品完了',
    'returnDeclined': '返品拒否'
  };
  
  return statusMap[orderStatus] || orderStatus || '不明';
}

/**
 * ==========================================
 * 7. 売上管理シートへの挿入機能
 * ==========================================
 */

/**
 * 変換済み注文データを売上管理シートに挿入
 * @param {Array} salesRows - 売上管理シート形式のデータ配列（複数行）
 * @returns {number} 挿入された行数
 */
function insertOrdersToSalesSheet(salesRows) {
  try {
    console.log('売上管理シートへのデータ挿入開始:', salesRows.length, '件');
    
    if (!salesRows || salesRows.length === 0) {
      console.log('挿入するデータがありません');
      return 0;
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = spreadsheet.getSheetByName(SHEET_NAMES.SALES);
    
    if (!salesSheet) {
      throw new Error('売上管理シートが見つかりません');
    }
    
    // 最終行を取得
    const lastRow = salesSheet.getLastRow();
    
    // データを挿入
    const startRow = lastRow + 1;
    const numRows = salesRows.length;
    const numCols = salesRows[0].length;
    
    const range = salesSheet.getRange(startRow, 1, numRows, numCols);
    range.setValues(salesRows);
    
    // 日時列のフォーマット設定
    const dateColumns = [1, 12, 15, 33, 34]; // A, L, O, AG, AH列
    dateColumns.forEach(col => {
      const dateRange = salesSheet.getRange(startRow, col, numRows, 1);
      dateRange.setNumberFormat('yyyy-mm-dd HH:mm:ss');
    });
    
    // 数値列のフォーマット設定
    const numberColumns = [7, 8, 9, 10, 11, 16, 17, 18, 19]; // G, H, I, J, K, P, Q, R, S列
    numberColumns.forEach(col => {
      const numberRange = salesSheet.getRange(startRow, col, numRows, 1);
      numberRange.setNumberFormat('#,##0.00');
    });
    
    console.log('売上管理シートへのデータ挿入完了:', numRows, '件');
    return numRows;
    
  } catch (error) {
    console.error('売上管理シートへのデータ挿入エラー:', error);
    throw error;
  }
}

/**
 * ==========================================
 * 8. 在庫管理シート更新機能
 * ==========================================
 */

/**
 * 在庫管理シートを更新（在庫数量減算、最終注文日更新、総注文数更新）
 * @param {Array} orders - 注文データ配列
 * @returns {number} 更新された商品数
 */
function updateInventorySheet(orders) {
  try {
    console.log('在庫管理シート更新開始:', orders.length, '件の注文');
    
    if (!orders || orders.length === 0) {
      return 0;
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
    
    if (!inventorySheet) {
      throw new Error('在庫管理シートが見つかりません');
    }
    
    const lastRow = inventorySheet.getLastRow();
    if (lastRow <= 1) {
      console.warn('在庫管理シートにデータがありません');
      return 0;
    }
    
    // SKUごとの注文集計
    const skuOrderMap = new Map();
    orders.forEach(order => {
      const sku = order.product.sku;
      const quantity = order.quantity || 1;
      const orderDate = new Date(order.orderTimestamp);
      
      if (!skuOrderMap.has(sku)) {
        skuOrderMap.set(sku, {
          totalQuantity: 0,
          orderCount: 0,
          latestOrderDate: orderDate
        });
      }
      
      const orderInfo = skuOrderMap.get(sku);
      orderInfo.totalQuantity += quantity;
      orderInfo.orderCount += 1;
      
      if (orderDate > orderInfo.latestOrderDate) {
        orderInfo.latestOrderDate = orderDate;
      }
    });
    
    // 在庫管理シートのデータを一括取得
    const dataRange = inventorySheet.getRange(2, 1, lastRow - 1, 20);
    const data = dataRange.getValues();
    
    let updatedCount = 0;
    
    // 各商品について更新
    data.forEach((row, index) => {
      const sku = row[COLUMN_INDEXES.INVENTORY.PRODUCT_ID - 1];
      
      if (skuOrderMap.has(sku)) {
        const orderInfo = skuOrderMap.get(sku);
        const actualRow = index + 2;
        
        // 在庫数量を減算（N列）
        const currentStock = parseInt(row[COLUMN_INDEXES.INVENTORY.STOCK_QUANTITY - 1]) || 0;
        const newStock = Math.max(0, currentStock - orderInfo.totalQuantity);
        inventorySheet.getRange(actualRow, COLUMN_INDEXES.INVENTORY.STOCK_QUANTITY).setValue(newStock);
        
        // 最終注文日を更新（P列）
        inventorySheet.getRange(actualRow, COLUMN_INDEXES.INVENTORY.LAST_ORDER_DATE)
          .setValue(orderInfo.latestOrderDate)
          .setNumberFormat('yyyy-mm-dd');
        
        // 総注文数を更新（Q列）
        const currentOrderCount = parseInt(row[COLUMN_INDEXES.INVENTORY.TOTAL_ORDERS - 1]) || 0;
        const newOrderCount = currentOrderCount + orderInfo.orderCount;
        inventorySheet.getRange(actualRow, COLUMN_INDEXES.INVENTORY.TOTAL_ORDERS).setValue(newOrderCount);
        
        updatedCount++;
        console.log(`在庫更新: ${sku}, 在庫: ${currentStock} → ${newStock}, 注文数: ${currentOrderCount} → ${newOrderCount}`);
      }
    });
    
    console.log('在庫管理シート更新完了:', updatedCount, '商品');
    return updatedCount;
    
  } catch (error) {
    console.error('在庫管理シート更新エラー:', error);
    throw error;
  }
}

/**
 * ==========================================
 * 9. バッチ変換・挿入機能
 * ==========================================
 */

/**
 * 複数のJoom注文を一括変換して売上管理シートに挿入
 * @param {Array} joomOrders - Joom注文データ配列
 * @returns {Object} 処理結果 {inserted: number, failed: number, errors: Array}
 */
function batchTransformAndInsertOrders(joomOrders) {
  try {
    console.log('一括変換・挿入開始:', joomOrders.length, '件');
    
    if (!joomOrders || joomOrders.length === 0) {
      return {
        inserted: 0,
        failed: 0,
        errors: []
      };
    }
    
    const salesRows = [];
    const errors = [];
    
    // 各注文を変換
    joomOrders.forEach((order, index) => {
      try {
        const salesRow = transformJoomOrderToSalesRow(order);
        salesRows.push(salesRow);
      } catch (error) {
        console.error(`注文変換エラー [${index + 1}/${joomOrders.length}]:`, error);
        errors.push({
          orderId: order.id || 'Unknown',
          error: error.message
        });
      }
    });
    
    // 売上管理シートに挿入
    let insertedCount = 0;
    if (salesRows.length > 0) {
      insertedCount = insertOrdersToSalesSheet(salesRows);
    }
    
    // 在庫管理シートを更新
    if (insertedCount > 0) {
      updateInventorySheet(joomOrders);
    }
    
    const result = {
      inserted: insertedCount,
      failed: errors.length,
      errors: errors
    };
    
    console.log('一括変換・挿入完了:', result);
    return result;
    
  } catch (error) {
    console.error('一括変換・挿入エラー:', error);
    throw error;
  }
}

