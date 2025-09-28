/**
 * 全シートの初期化を管理するメインスクリプト
 * 在庫管理ツール システム開発
 * 作成日: 2025-09-27
 */

/**
 * 全シートの初期化を実行
 */
function initializeAllSheets() {
  console.log('在庫管理ツールのシート初期化を開始します...');
  
  try {
    // 各シートを順次初期化
    initializeInventorySheet();
    initializeSalesSheet();
    initializeSupplierMasterSheet();
    initializePriceHistorySheet();
    
    // カスタムメニューを設定
    setupCustomMenu();
    
    console.log('全シートの初期化が完了しました');
    SpreadsheetApp.getUi().alert('シート初期化完了', '在庫管理ツールの全シートが正常に作成されました。', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('シート初期化中にエラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラー', 'シート初期化中にエラーが発生しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * カスタムメニューの設定
 */
function setupCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📊 EC管理システム')
    .addSubMenu(
      ui.createMenu('📦 商品管理')
        .addItem('新商品追加', 'showProductInputForm')
        .addItem('商品削除', 'showProductDeleteMenu')
    )
    .addSubMenu(
      ui.createMenu('📈 売上管理')
        .addItem('注文データ入力', 'showSalesInputFormMenu')
        .addItem('売上一覧表示', 'showSalesList')
        .addItem('売上統計表示', 'showSalesStatistics')
    )
    .addSubMenu(
      ui.createMenu('🔄 在庫チェック')
        .addItem('全商品在庫チェック', 'checkAllInventory')
        .addItem('選択商品チェック', 'checkSelectedInventory')
        .addItem('定期チェック設定', 'setupScheduledCheck')
    )
    .addSubMenu(
      ui.createMenu('📤 データ出力')
        .addItem('Joom用CSV出力', 'exportJoomCsv')
    )
    .addSubMenu(
      ui.createMenu('⚙️ システム設定')
        .addItem('通知設定', 'showNotificationSettings')
    )
    .addSeparator()
    .addItem('🔄 全シート初期化', 'initializeAllSheets')
    .addToUi();
}

/**
 * 商品削除メニューの表示
 */
function showProductDeleteMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('商品削除', '商品削除機能は今後実装予定です。', ui.ButtonSet.OK);
}

/**
 * 売上データ入力フォームの表示
 */
function showSalesInputFormMenu() {
  // SalesManagement.gsの関数を呼び出し
  try {
    // 直接関数を呼び出し（GASでは同じプロジェクト内の関数は直接呼び出し可能）
    showSalesInputForm();
  } catch (error) {
    console.error('売上データ入力フォームの表示中にエラーが発生しました:', error);
    const ui = SpreadsheetApp.getUi();
    ui.alert('エラー', '売上データ入力フォームの表示中にエラーが発生しました。', ui.ButtonSet.OK);
  }
}

/**
 * 全商品在庫チェックの実行
 */
function checkAllInventory() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('全商品在庫チェック', '全商品在庫チェック機能は今後実装予定です。', ui.ButtonSet.OK);
}

/**
 * 選択商品在庫チェックの実行
 */
function checkSelectedInventory() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('選択商品チェック', '選択商品チェック機能は今後実装予定です。', ui.ButtonSet.OK);
}

/**
 * 定期チェック設定の表示
 */
function setupScheduledCheck() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('定期チェック設定', '定期チェック設定機能は今後実装予定です。', ui.ButtonSet.OK);
}

/**
 * Joom用CSV出力の実行
 */
function exportJoomCsv() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Joom用CSV出力', 'Joom用CSV出力機能は今後実装予定です。', ui.ButtonSet.OK);
}

/**
 * 通知設定の表示
 */
function showNotificationSettings() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('通知設定', '通知設定機能は今後実装予定です。', ui.ButtonSet.OK);
}

/**
 * 売上統計表示
 */
function showSalesStatistics() {
  try {
    const stats = getSalesStatistics();
    if (stats) {
      const message = `売上統計:\n\n` +
        `総売上: ¥${stats.totalSales.toLocaleString()}\n` +
        `総利益: ¥${stats.totalProfit.toLocaleString()}\n` +
        `注文数: ${stats.orderCount}件\n` +
        `平均注文単価: ¥${Math.round(stats.averageOrderValue).toLocaleString()}`;
      
      const ui = SpreadsheetApp.getUi();
      ui.alert('売上統計', message, ui.ButtonSet.OK);
    } else {
      const ui = SpreadsheetApp.getUi();
      ui.alert('エラー', '売上統計の取得に失敗しました。', ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('売上統計表示中にエラーが発生しました:', error);
    const ui = SpreadsheetApp.getUi();
    ui.alert('エラー', '売上統計の表示中にエラーが発生しました。', ui.ButtonSet.OK);
  }
}

/**
 * スプレッドシートが開かれた時の初期化
 */
function onOpen() {
  setupCustomMenu();
}
