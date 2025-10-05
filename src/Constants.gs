/**
 * 定数定義ファイル
 * 在庫管理ツール システム開発
 * 作成日: 2025-09-29
 */

/**
 * スプレッドシートの列インデックス定数
 */
const COLUMN_INDEXES = {
  // 在庫管理シート
  INVENTORY: {
    PRODUCT_ID: 1,           // A列: 商品ID
    PRODUCT_NAME: 2,         // B列: 商品名
    SKU: 3,                  // C列: SKU
    ASIN: 4,                 // D列: ASIN
    SUPPLIER: 5,             // E列: 仕入れ元
    SUPPLIER_URL: 6,         // F列: 仕入れ元URL
    PURCHASE_PRICE: 7,       // G列: 仕入れ価格
    SELLING_PRICE: 8,        // H列: 販売価格
    WEIGHT: 9,               // I列: 重量
    // Joom対応フィールド（I列とO列の間に挿入）
    DESCRIPTION: 10,         // J列: 商品説明
    MAIN_IMAGE_URL: 11,      // K列: メイン画像URL
    CURRENCY: 12,            // L列: 通貨
    SHIPPING_PRICE: 13,      // M列: 配送価格
    STOCK_QUANTITY: 14,      // N列: 在庫数量
    // 既存フィールド（列番号がシフト）
    STOCK_STATUS: 15,        // O列: 在庫ステータス
    PROFIT: 16,              // P列: 利益
    LAST_UPDATED: 17,        // Q列: 最終更新日時
    NOTES: 18,               // R列: 備考・メモ
    // Joom連携管理列
    JOOM_STATUS: 19,         // S列: Joom連携ステータス
    JOOM_LAST_EXPORT: 20     // T列: 最終出力日時
  },
  
  // 売上管理シート
  SALES: {
    ORDER_ID: 1,             // A列: 注文ID
    ORDER_DATE: 2,           // B列: 注文日
    PRODUCT_ID: 3,           // C列: 商品ID
    PRODUCT_NAME: 4,         // D列: 商品名
    SKU: 5,                  // E列: SKU
    ASIN: 6,                 // F列: ASIN
    QUANTITY: 7,             // G列: 数量
    SELLING_PRICE: 8,        // H列: 販売価格
    PURCHASE_PRICE: 9,       // I列: 仕入れ価格
    SHIPPING_COST: 10,       // J列: 送料
    NET_PROFIT: 11,          // K列: 純利益
    REGISTRATION_TIME: 12    // L列: 登録日時
  },
  
  // 価格履歴シート（1商品1行形式）
  PRICE_HISTORY: {
    PRODUCT_ID: 1,           // A列: 商品ID
    PRODUCT_NAME: 2,         // B列: 商品名
    CURRENT_PURCHASE_PRICE: 3, // C列: 現在仕入れ価格
    PREVIOUS_PURCHASE_PRICE: 4, // D列: 前回仕入れ価格
    PURCHASE_PRICE_CHANGE: 5, // E列: 仕入れ価格変動
    PURCHASE_PRICE_CHANGE_RATE: 6, // F列: 仕入れ価格変動率
    CURRENT_SELLING_PRICE: 7, // G列: 現在販売価格
    PREVIOUS_SELLING_PRICE: 8, // H列: 前回販売価格
    SELLING_PRICE_CHANGE: 9, // I列: 販売価格変動
    SELLING_PRICE_CHANGE_RATE: 10, // J列: 販売価格変動率
    LAST_UPDATED: 11,        // K列: 最終更新日時
    PURCHASE_CHANGE_COUNT: 12, // L列: 仕入れ価格変動回数
    SELLING_CHANGE_COUNT: 13, // M列: 販売価格変動回数
    NOTES: 14                // N列: 備考
  }
};

/**
 * シート名定数
 */
const SHEET_NAMES = {
  INVENTORY: '在庫管理',
  SALES: '売上管理',
  SUPPLIER_MASTER: '仕入れ元マスター',
  PRICE_HISTORY: '価格履歴',
  SETTINGS: '設定'
};

/**
 * 在庫ステータス定数
 */
const STOCK_STATUS = {
  IN_STOCK: '在庫あり',
  OUT_OF_STOCK: '売り切れ',
  RESERVED: '予約受付中'
};

/**
 * 仕入れ元サイト定数
 */
const SUPPLIER_SITES = {
  AMAZON: 'Amazon',
  RAKUTEN: '楽天市場',
  YAHOO_SHOPPING: 'Yahooショッピング',
  MERUKARI: 'メルカリ',
  YAHOO_AUCTION: 'ヤフオク',
  YAHOO_FLEA: 'ヤフーフリマ',
  PRIVATE_SHOP: '個人ショップ'
};

/**
 * フォーム関連定数
 */
const FORM_CONSTANTS = {
  MIN_QUANTITY: 1,
  MIN_PRICE: 0,
  DEFAULT_SHIPPING_COST: 0
};

/**
 * Joom連携ステータス定数
 */
const JOOM_STATUS = {
  UNLINKED: '未連携',
  LINKED: '連携済み'
};

/**
 * Joom CSVフィールドマッピング定数
 */
const JOOM_CSV_FIELDS = {
  // 必須フィールド
  REQUIRED: {
    PRODUCT_SKU: 'Product SKU',
    NAME: 'Name',
    DESCRIPTION: 'Description',
    PRODUCT_MAIN_IMAGE_URL: 'Product Main Image URL',
    STORE_ID: 'Store ID',
    SHIPPING_WEIGHT: 'Shipping Weight (Kg)',
    PRICE: 'Price (without VAT)',
    CURRENCY: 'Currency',
    INVENTORY: 'Inventory (default warehouse)',
    SHIPPING_PRICE: 'Shipping Price (without VAT) (default warehouse)'
  },
  // 推奨フィールド
  RECOMMENDED: {
    BRAND: 'Brand',
    SEARCH_TAGS: 'Search Tags',
    DANGEROUS_KIND: 'Dangerous Kind',
    SUGGESTED_CATEGORY_ID: 'Suggested Category ID',
    LANDING_PAGE_URL: 'Landing Page URL',
    EXTRA_IMAGE_URLS: 'Extra Image URLs (max 20)',
    MANUFACTURE_GTIN: 'Manufacture GTIN',
    COLOR: 'Color',
    SIZE: 'Size',
    MSRP: 'MSRP',
    VARIANT_SKU: 'Variant SKU',
    VARIANT_MAIN_IMAGE_URL: 'Variant Main Image URL',
    SHIPPING_LENGTH: 'Shipping Length (cm)',
    SHIPPING_WIDTH: 'Shipping Width (cm)',
    SHIPPING_HEIGHT: 'Shipping Height (cm)',
    HS_CODE: 'HS Code',
    DECLARED_VALUE: 'Declared Value'
  }
};

/**
 * Joom CSV出力設定定数
 */
const JOOM_CSV_CONFIG = {
  // 出力対象
  TARGET_PRODUCTS: {
    ALL: 'all',
    UNLINKED: 'unlinked',
    SELECTED: 'selected'
  },
  // デフォルト値
  DEFAULTS: {
    CURRENCY: 'JPY',
    DANGEROUS_KIND: 'notdangerous',
    SHIPPING_PRICE: 0,
    WEIGHT_UNIT_CONVERSION: 1000 // g to kg
  },
  // バリデーション
  VALIDATION: {
    MIN_IMAGE_SIZE: 500,
    MAX_IMAGES: 20,
    MAX_INVENTORY: 100000,
    REQUIRED_FIELDS: Object.values(JOOM_CSV_FIELDS.REQUIRED)
  }
};

/**
 * メッセージ定数
 */
const MESSAGES = {
  SUCCESS: {
    PRODUCT_SAVED: '商品が正常に保存されました',
    SALES_SAVED: '注文データが正常に保存されました',
    SHEET_INITIALIZED: 'シートの初期化が完了しました',
    CSV_EXPORTED: 'Joom用CSVの出力が完了しました',
    STATUS_UPDATED: '連携ステータスが更新されました'
  },
  ERROR: {
    SHEET_NOT_FOUND: 'シートが見つかりません',
    SAVE_FAILED: 'データの保存中にエラーが発生しました',
    VALIDATION_FAILED: '入力データの検証に失敗しました',
    CSV_EXPORT_FAILED: 'CSV出力中にエラーが発生しました',
    INVALID_DATA: '無効なデータが検出されました'
  }
};
