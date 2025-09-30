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
    STOCK_STATUS: 10,        // J列: 在庫ステータス
    PROFIT: 11,              // K列: 利益
    LAST_UPDATED: 12,        // L列: 最終更新日時
    NOTES: 13                // M列: 備考・メモ
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
  }
};

/**
 * シート名定数
 */
const SHEET_NAMES = {
  INVENTORY: '在庫管理',
  SALES: '売上管理',
  SUPPLIER_MASTER: '仕入れ元マスター',
  PRICE_HISTORY: '価格履歴'
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
 * メッセージ定数
 */
const MESSAGES = {
  SUCCESS: {
    PRODUCT_SAVED: '商品が正常に保存されました',
    SALES_SAVED: '注文データが正常に保存されました',
    SHEET_INITIALIZED: 'シートの初期化が完了しました'
  },
  ERROR: {
    SHEET_NOT_FOUND: 'シートが見つかりません',
    SAVE_FAILED: 'データの保存中にエラーが発生しました',
    VALIDATION_FAILED: '入力データの検証に失敗しました'
  }
};
