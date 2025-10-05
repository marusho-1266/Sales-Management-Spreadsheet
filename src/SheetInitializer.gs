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
    initializeSettingsSheet();
    
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
    )
    .addSubMenu(
      ui.createMenu('🔄 在庫チェック')
        .addItem('全商品在庫チェック', 'checkAllInventory')
        .addItem('選択商品チェック', 'checkSelectedInventory')
        .addItem('定期チェック設定', 'setupScheduledCheck')
    )
    .addSubMenu(
      ui.createMenu('📤 データ出力')
        .addItem('Joom用CSV出力', 'exportUnlinkedProductsCsv')
    )
    .addSubMenu(
      ui.createMenu('⚙️ 設定')
        .addItem('設定シート初期化', 'initializeSettingsSheet')
        .addItem('設定値一括更新', 'showSettingsUpdateForm')
    )
    .addSubMenu(
      ui.createMenu('💰 価格履歴')
        .addItem('価格履歴同期', 'syncPriceHistoryMenu')
        .addSeparator()
        .addItem('📧 メールアドレス設定', 'showEmailAddressSettings')
        .addItem('🔔 通知有効化設定', 'showNotificationEnableSettings')
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
 * 価格履歴同期メニュー
 */
function syncPriceHistoryMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const result = syncPriceHistoryFromInventory();
    if (result) {
      ui.alert('価格履歴同期完了', '在庫管理シートから価格履歴を正常に同期しました。', ui.ButtonSet.OK);
    } else {
      ui.alert('エラー', '価格履歴の同期中にエラーが発生しました。', ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('価格履歴同期中にエラーが発生しました:', error);
    ui.alert('エラー', '価格履歴の同期中にエラーが発生しました: ' + error.message, ui.ButtonSet.OK);
  }
}


/**
 * メールアドレス設定
 */
function showEmailAddressSettings() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const currentEmail = getSetting('価格変動通知メールアドレス') || '';
    
    const emailResponse = ui.prompt(
      '価格変動通知メールアドレス設定',
      `現在の設定: ${currentEmail}\n\n通知先のメールアドレスを入力してください:\n（空欄にするとメールアドレスをクリアします）`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (emailResponse.getSelectedButton() === ui.Button.OK) {
      const newEmail = emailResponse.getResponseText().trim();
      
      // メールアドレスの検証
      if (newEmail === '') {
        // 空の場合はクリア
        updateSetting('価格変動通知メールアドレス', '');
        ui.alert(
          '設定完了',
          'メールアドレスをクリアしました。',
          ui.ButtonSet.OK
        );
        console.log('価格変動通知メールアドレスをクリアしました');
        
      } else if (newEmail.includes('@') && newEmail.includes('.')) {
        // 有効なメールアドレスの場合
        updateSetting('価格変動通知メールアドレス', newEmail);
        ui.alert(
          '設定完了',
          `メールアドレスを設定しました: ${newEmail}`,
          ui.ButtonSet.OK
        );
        console.log(`価格変動通知メールアドレスを設定しました: ${newEmail}`);
        
      } else {
        ui.alert(
          'エラー',
          '有効なメールアドレスを入力してください。',
          ui.ButtonSet.OK
        );
      }
    }
    
  } catch (error) {
    console.error('メールアドレス設定中にエラーが発生しました:', error);
    ui.alert(
      'エラー',
      'メールアドレス設定中にエラーが発生しました: ' + error.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 通知有効化設定
 */
function showNotificationEnableSettings() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const currentEnabled = getSetting('価格変動通知有効化') || 'true';
    const currentEmail = getSetting('価格変動通知メールアドレス') || '';
    
    if (!currentEmail) {
      ui.alert(
        'エラー',
        'メールアドレスが設定されていません。\n先にメールアドレスを設定してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    const statusText = currentEnabled === 'true' ? '有効' : '無効';
    const enableResponse = ui.alert(
      '価格変動通知の有効化設定',
      `現在の状態: ${statusText}\nメールアドレス: ${currentEmail}\n\n価格変動通知を${currentEnabled === 'true' ? '無効' : '有効'}にしますか？`,
      ui.ButtonSet.YES_NO
    );
    
    if (enableResponse === ui.Button.YES) {
      const newStatus = currentEnabled === 'true' ? 'false' : 'true';
      updateSetting('価格変動通知有効化', newStatus);
      
      const newStatusText = newStatus === 'true' ? '有効' : '無効';
      ui.alert(
        '設定完了',
        `価格変動通知を${newStatusText}に設定しました。`,
        ui.ButtonSet.OK
      );
      console.log(`価格変動通知の有効化を${newStatusText}に設定しました`);
    }
    
  } catch (error) {
    console.error('通知有効化設定中にエラーが発生しました:', error);
    ui.alert(
      'エラー',
      '通知有効化設定中にエラーが発生しました: ' + error.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 通知設定の表示
 */
function showNotificationSettings() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('通知設定', '通知設定機能は今後実装予定です。', ui.ButtonSet.OK);
}

/**
 * 設定更新フォームの表示
 */
function showSettingsUpdateForm() {
  const ui = SpreadsheetApp.getUi();
  
  // 現在の設定値を取得
  const storeId = getSetting('ストアID') || 'STORE001';
  
  // プロンプトで設定値を入力
  const newStoreId = ui.prompt('ストアID設定', `現在のストアID: ${storeId}\n新しいストアIDを入力してください:`, ui.ButtonSet.OK_CANCEL);
  
  if (newStoreId.getSelectedButton() === ui.Button.OK) {
    const storeIdValue = newStoreId.getResponseText().trim();
    if (storeIdValue) {
      updateSetting('ストアID', storeIdValue);
      
      ui.alert('設定更新完了', `ストアIDを「${storeIdValue}」に更新しました。\nCSV出力時にこの設定値が使用されます。`, ui.ButtonSet.OK);
    } else {
      ui.alert('エラー', 'ストアIDが入力されていません。', ui.ButtonSet.OK);
    }
  }
}


/**
 * スプレッドシートが開かれた時の初期化
 */
function onOpen() {
  setupCustomMenu();
}
