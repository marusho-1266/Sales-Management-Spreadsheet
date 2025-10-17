/**
 * ==========================================
 * Joom注文 3層同期システム
 * ==========================================
 * 
 * このファイルは、Joom注文データの3層同期システムを提供します：
 * 1. 手動即時処理: ユーザーが手動で実行
 * 2. 定期処理: 1時間間隔で自動実行
 * 3. 日次総チェック: 毎日午前2時に前日分を包括チェック
 */

/**
 * ==========================================
 * 1. 手動即時処理
 * ==========================================
 */

/**
 * 手動即時処理: 前回連携時間以降の注文を即座に取得
 * メニューから実行される（既に実装済み）
 */
function manualImmediateSync() {
  console.log('=== 手動即時処理開始 ===');
  
  try {
    const startTime = new Date();
    
    // トークンの有効性確認
    if (!ensureValidToken()) {
      throw new Error('有効なアクセストークンがありません');
    }
    
    // 前回連携時間以降の注文を取得
    const orders = fetchJoomOrders({ mode: 'lastSync' });
    
    if (!orders || orders.length === 0) {
      console.log('新しい注文はありません');
      return {
        success: true,
        fetched: 0,
        inserted: 0,
        failed: 0,
        duration: (new Date().getTime() - startTime.getTime()) / 1000
      };
    }
    
    // データ変換・挿入
    const result = batchTransformAndInsertOrders(orders);
    
    // 前回連携時間を更新
    const endTime = new Date();
    updateSetting('Joom 前回連携時間', Utilities.formatDate(endTime, 'JST', 'yyyy-MM-dd HH:mm:ss'));
    
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    
    console.log('=== 手動即時処理完了 ===');
    console.log(`取得: ${orders.length}件, 登録: ${result.inserted}件, 失敗: ${result.failed}件, 処理時間: ${duration}秒`);
    
    return {
      success: true,
      fetched: orders.length,
      inserted: result.inserted,
      failed: result.failed,
      errors: result.errors,
      duration: duration
    };
    
  } catch (error) {
    console.error('手動即時処理エラー:', error);
    return {
      success: false,
      error: error.message,
      duration: (new Date().getTime() - startTime.getTime()) / 1000
    };
  }
}

/**
 * ==========================================
 * 2. 定期処理（1時間間隔）
 * ==========================================
 */

/**
 * 定期処理: 1時間間隔で前回連携時間以降の注文を取得
 * トリガーにより自動実行される
 */
function hourlyScheduledSync() {
  console.log('=== 定期処理（1時間間隔）開始 ===');
  
  try {
    const startTime = new Date();
    
    // トークンの有効性確認
    if (!ensureValidToken()) {
      throw new Error('有効なアクセストークンがありません');
    }
    
    // 前回連携時間以降の注文を取得
    const orders = fetchJoomOrders({ mode: 'lastSync' });
    
    if (!orders || orders.length === 0) {
      console.log('新しい注文はありません');
      
      // 同期履歴を記録
      recordSyncHistory({
        syncType: '定期処理',
        fetched: 0,
        inserted: 0,
        failed: 0,
        duration: (new Date().getTime() - startTime.getTime()) / 1000,
        status: 'success',
        message: '新しい注文なし'
      });
      
      return {
        success: true,
        fetched: 0,
        inserted: 0,
        failed: 0
      };
    }
    
    // データ変換・挿入
    const result = batchTransformAndInsertOrders(orders);
    
    // 前回連携時間を更新
    const endTime = new Date();
    updateSetting('Joom 前回連携時間', Utilities.formatDate(endTime, 'JST', 'yyyy-MM-dd HH:mm:ss'));
    
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    
    // 同期履歴を記録
    recordSyncHistory({
      syncType: '定期処理',
      fetched: orders.length,
      inserted: result.inserted,
      failed: result.failed,
      duration: duration,
      status: 'success',
      message: `${result.inserted}件登録成功`
    });
    
    // 通知送信（エラーがある場合のみ）
    if (result.failed > 0) {
      sendSyncNotification({
        syncType: '定期処理',
        fetched: orders.length,
        inserted: result.inserted,
        failed: result.failed,
        errors: result.errors
      });
    }
    
    console.log('=== 定期処理完了 ===');
    console.log(`取得: ${orders.length}件, 登録: ${result.inserted}件, 失敗: ${result.failed}件, 処理時間: ${duration}秒`);
    
    return {
      success: true,
      fetched: orders.length,
      inserted: result.inserted,
      failed: result.failed,
      errors: result.errors,
      duration: duration
    };
    
  } catch (error) {
    console.error('定期処理エラー:', error);
    
    // エラー履歴を記録
    recordSyncHistory({
      syncType: '定期処理',
      fetched: 0,
      inserted: 0,
      failed: 0,
      duration: (new Date().getTime() - startTime.getTime()) / 1000,
      status: 'error',
      message: error.message
    });
    
    // エラー通知送信
    sendErrorNotification({
      syncType: '定期処理',
      error: error.message
    });
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ==========================================
 * 3. 日次総チェック（毎日午前2時）
 * ==========================================
 */

/**
 * 日次総チェック: 前日1日分の注文を包括的にチェック
 * 6時間単位で分割して取得し、漏れを防ぐ
 * トリガーにより毎日午前2時に自動実行される
 */
function dailyComprehensiveCheck() {
  console.log('=== 日次総チェック開始 ===');
  
  try {
    const startTime = new Date();
    
    // トークンの有効性確認
    if (!ensureValidToken()) {
      throw new Error('有効なアクセストークンがありません');
    }
    
    // 前日の0時から今日の0時までの期間を設定
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    yesterday.setHours(0, 0, 0, 0);
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    console.log(`日次総チェック期間: ${yesterday.toISOString()} - ${today.toISOString()}`);
    
    // 6時間単位で分割して取得（API負荷分散）
    const timeSlots = [
      { start: new Date(yesterday.getTime()), end: new Date(yesterday.getTime() + 6 * 60 * 60 * 1000) },
      { start: new Date(yesterday.getTime() + 6 * 60 * 60 * 1000), end: new Date(yesterday.getTime() + 12 * 60 * 60 * 1000) },
      { start: new Date(yesterday.getTime() + 12 * 60 * 60 * 1000), end: new Date(yesterday.getTime() + 18 * 60 * 60 * 1000) },
      { start: new Date(yesterday.getTime() + 18 * 60 * 60 * 1000), end: today }
    ];
    
    let totalFetched = 0;
    let totalInserted = 0;
    let totalFailed = 0;
    const allErrors = [];
    
    // 各時間帯の注文を取得
    for (let i = 0; i < timeSlots.length; i++) {
      const slot = timeSlots[i];
      console.log(`時間帯 ${i + 1}/4: ${slot.start.toISOString()} - ${slot.end.toISOString()}`);
      
      try {
        const orders = fetchJoomOrders({
          mode: 'dateRange',
          startDate: slot.start,
          endDate: slot.end
        });
        
        if (orders && orders.length > 0) {
          const result = batchTransformAndInsertOrders(orders);
          
          totalFetched += orders.length;
          totalInserted += result.inserted;
          totalFailed += result.failed;
          
          if (result.errors && result.errors.length > 0) {
            allErrors.push(...result.errors);
          }
          
          console.log(`時間帯 ${i + 1} 完了: 取得 ${orders.length}件, 登録 ${result.inserted}件`);
        } else {
          console.log(`時間帯 ${i + 1}: 注文なし`);
        }
        
        // レート制限対応（次の時間帯まで少し待機）
        if (i < timeSlots.length - 1) {
          Utilities.sleep(2000);
        }
        
      } catch (error) {
        console.error(`時間帯 ${i + 1} エラー:`, error);
        allErrors.push({
          timeSlot: `${i + 1}/4`,
          error: error.message
        });
      }
    }
    
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    
    // 同期履歴を記録
    recordSyncHistory({
      syncType: '日次総チェック',
      fetched: totalFetched,
      inserted: totalInserted,
      failed: totalFailed,
      duration: duration,
      status: totalFailed > 0 ? 'warning' : 'success',
      message: `前日分チェック完了: ${totalInserted}件登録`
    });
    
    // 日次レポート送信
    sendDailyReport({
      date: yesterday,
      fetched: totalFetched,
      inserted: totalInserted,
      failed: totalFailed,
      errors: allErrors,
      duration: duration
    });
    
    console.log('=== 日次総チェック完了 ===');
    console.log(`取得: ${totalFetched}件, 登録: ${totalInserted}件, 失敗: ${totalFailed}件, 処理時間: ${duration}秒`);
    
    return {
      success: true,
      fetched: totalFetched,
      inserted: totalInserted,
      failed: totalFailed,
      errors: allErrors,
      duration: duration
    };
    
  } catch (error) {
    console.error('日次総チェックエラー:', error);
    
    // エラー履歴を記録
    recordSyncHistory({
      syncType: '日次総チェック',
      fetched: 0,
      inserted: 0,
      failed: 0,
      duration: (new Date().getTime() - startTime.getTime()) / 1000,
      status: 'error',
      message: error.message
    });
    
    // エラー通知送信
    sendErrorNotification({
      syncType: '日次総チェック',
      error: error.message
    });
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ==========================================
 * 4. トリガー設定機能
 * ==========================================
 */

/**
 * 定期処理トリガーを設定（1時間間隔）
 */
function setupHourlyTrigger() {
  try {
    // 既存のトリガーを削除
    deleteHourlyTrigger();
    
    // 新しいトリガーを作成
    ScriptApp.newTrigger('hourlyScheduledSync')
      .timeBased()
      .everyHours(1)
      .create();
    
    console.log('定期処理トリガー（1時間間隔）を設定しました');
    return true;
    
  } catch (error) {
    console.error('定期処理トリガー設定エラー:', error);
    throw error;
  }
}

/**
 * 日次総チェックトリガーを設定（毎日午前2時）
 */
function setupDailyTrigger() {
  try {
    // 既存のトリガーを削除
    deleteDailyTrigger();
    
    // 新しいトリガーを作成
    ScriptApp.newTrigger('dailyComprehensiveCheck')
      .timeBased()
      .atHour(2)
      .everyDays(1)
      .create();
    
    console.log('日次総チェックトリガー（毎日午前2時）を設定しました');
    return true;
    
  } catch (error) {
    console.error('日次総チェックトリガー設定エラー:', error);
    throw error;
  }
}

/**
 * すべてのトリガーを一括設定
 */
function setupAllTriggers() {
  try {
    setupHourlyTrigger();
    setupDailyTrigger();
    
    console.log('すべてのトリガーを設定しました');
    return {
      success: true,
      message: '定期処理（1時間間隔）と日次総チェック（毎日午前2時）のトリガーを設定しました'
    };
    
  } catch (error) {
    console.error('トリガー一括設定エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 定期処理トリガーを削除
 */
function deleteHourlyTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'hourlyScheduledSync') {
        ScriptApp.deleteTrigger(trigger);
        console.log('定期処理トリガーを削除しました');
      }
    });
    
    return true;
    
  } catch (error) {
    console.error('定期処理トリガー削除エラー:', error);
    throw error;
  }
}

/**
 * 日次総チェックトリガーを削除
 */
function deleteDailyTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'dailyComprehensiveCheck') {
        ScriptApp.deleteTrigger(trigger);
        console.log('日次総チェックトリガーを削除しました');
      }
    });
    
    return true;
    
  } catch (error) {
    console.error('日次総チェックトリガー削除エラー:', error);
    throw error;
  }
}

/**
 * すべてのJoom同期トリガーを削除
 */
function deleteAllSyncTriggers() {
  try {
    deleteHourlyTrigger();
    deleteDailyTrigger();
    
    console.log('すべての同期トリガーを削除しました');
    return {
      success: true,
      message: 'すべての同期トリガーを削除しました'
    };
    
  } catch (error) {
    console.error('トリガー一括削除エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 現在のトリガー状態を確認
 */
function checkTriggerStatus() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    let hourlyTrigger = null;
    let dailyTrigger = null;
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'hourlyScheduledSync') {
        hourlyTrigger = trigger;
      } else if (trigger.getHandlerFunction() === 'dailyComprehensiveCheck') {
        dailyTrigger = trigger;
      }
    });
    
    return {
      hourlyTrigger: {
        enabled: hourlyTrigger !== null,
        info: hourlyTrigger ? `1時間間隔で実行` : '未設定'
      },
      dailyTrigger: {
        enabled: dailyTrigger !== null,
        info: dailyTrigger ? `毎日午前2時に実行` : '未設定'
      }
    };
    
  } catch (error) {
    console.error('トリガー状態確認エラー:', error);
    // エラー時もデフォルト構造を返す
    return {
      hourlyTrigger: {
        enabled: false,
        info: 'エラー: ' + error.message
      },
      dailyTrigger: {
        enabled: false,
        info: 'エラー: ' + error.message
      },
      error: error.message
    };
  }
}

/**
 * ==========================================
 * 5. 同期履歴記録機能
 * ==========================================
 */

/**
 * 同期履歴を設定シートに記録
 * @param {Object} syncInfo - 同期情報
 */
function recordSyncHistory(syncInfo) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = spreadsheet.getSheetByName(SHEET_NAMES.SETTINGS);
    
    if (!settingsSheet) {
      console.warn('設定シートが見つかりません');
      return;
    }
    
    // 履歴エリアを探す（設定シートの下部に追加）
    const lastRow = settingsSheet.getLastRow();
    
    // 履歴ヘッダーが存在しない場合は作成
    const headerRow = lastRow + 2;
    const headers = ['同期日時', '同期タイプ', '取得件数', '登録件数', '失敗件数', '処理時間(秒)', 'ステータス', 'メッセージ'];
    
    // ヘッダーが存在するか確認
    const existingHeaders = settingsSheet.getRange(headerRow, 1, 1, headers.length).getValues()[0];
    if (existingHeaders[0] !== '同期日時') {
      // ヘッダーを作成
      settingsSheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
      settingsSheet.getRange(headerRow, 1, 1, headers.length).setFontWeight('bold');
    }
    
    // 履歴データを追加
    const historyRow = [
      new Date(),
      syncInfo.syncType || '',
      syncInfo.fetched || 0,
      syncInfo.inserted || 0,
      syncInfo.failed || 0,
      syncInfo.duration || 0,
      syncInfo.status || '',
      syncInfo.message || ''
    ];
    
    const dataRow = settingsSheet.getLastRow() + 1;
    settingsSheet.getRange(dataRow, 1, 1, historyRow.length).setValues([historyRow]);
    
    // 日時列のフォーマット
    settingsSheet.getRange(dataRow, 1).setNumberFormat('yyyy-mm-dd HH:mm:ss');
    
    console.log('同期履歴を記録しました:', syncInfo);
    
  } catch (error) {
    console.error('同期履歴記録エラー:', error);
  }
}

/**
 * ==========================================
 * 6. 通知機能
 * ==========================================
 */

/**
 * 同期完了通知を送信
 * @param {Object} syncResult - 同期結果
 */
function sendSyncNotification(syncResult) {
  try {
    const notificationEnabled = getSetting('Joom 通知有効化');
    if (notificationEnabled !== 'true' && notificationEnabled !== true) {
      console.log('通知が無効化されています');
      return;
    }
    
    const emailAddress = getSetting('通知先メールアドレス');
    if (!emailAddress) {
      console.log('通知先メールアドレスが設定されていません');
      return;
    }
    
    const subject = `[Joom注文同期] ${syncResult.syncType} - エラーあり`;
    
    let body = `Joom注文同期が完了しましたが、一部エラーがありました。\n\n`;
    body += `【同期タイプ】${syncResult.syncType}\n`;
    body += `【同期日時】${Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd HH:mm:ss')}\n\n`;
    body += `【結果】\n`;
    body += `取得件数: ${syncResult.fetched}件\n`;
    body += `登録成功: ${syncResult.inserted}件\n`;
    body += `登録失敗: ${syncResult.failed}件\n\n`;
    
    if (syncResult.errors && syncResult.errors.length > 0) {
      body += `【エラー詳細】\n`;
      syncResult.errors.forEach((error, index) => {
        body += `${index + 1}. 注文ID: ${error.orderId}, エラー: ${error.error}\n`;
      });
    }
    
    body += `\n売上管理シートを確認してください。`;
    
    MailApp.sendEmail(emailAddress, subject, body);
    console.log('同期完了通知を送信しました:', emailAddress);
    
  } catch (error) {
    console.error('同期完了通知送信エラー:', error);
  }
}

/**
 * エラー通知を送信
 * @param {Object} errorInfo - エラー情報
 */
function sendErrorNotification(errorInfo) {
  try {
    const notificationEnabled = getSetting('Joom 通知有効化');
    if (notificationEnabled !== 'true' && notificationEnabled !== true) {
      console.log('通知が無効化されています');
      return;
    }
    
    const emailAddress = getSetting('通知先メールアドレス');
    if (!emailAddress) {
      console.log('通知先メールアドレスが設定されていません');
      return;
    }
    
    const subject = `[Joom注文同期] ${errorInfo.syncType} - エラー発生`;
    
    let body = `Joom注文同期中にエラーが発生しました。\n\n`;
    body += `【同期タイプ】${errorInfo.syncType}\n`;
    body += `【同期日時】${Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd HH:mm:ss')}\n\n`;
    body += `【エラー内容】\n${errorInfo.error}\n\n`;
    body += `設定シートとログを確認してください。`;
    
    MailApp.sendEmail(emailAddress, subject, body);
    console.log('エラー通知を送信しました:', emailAddress);
    
  } catch (error) {
    console.error('エラー通知送信エラー:', error);
  }
}

/**
 * 日次レポートを送信
 * @param {Object} reportData - レポートデータ
 */
function sendDailyReport(reportData) {
  try {
    const notificationEnabled = getSetting('Joom 通知有効化');
    if (notificationEnabled !== 'true' && notificationEnabled !== true) {
      console.log('通知が無効化されています');
      return;
    }
    
    const emailAddress = getSetting('通知先メールアドレス');
    if (!emailAddress) {
      console.log('通知先メールアドレスが設定されていません');
      return;
    }
    
    const dateStr = Utilities.formatDate(reportData.date, 'JST', 'yyyy年MM月dd日');
    const subject = `[Joom注文同期] 日次レポート - ${dateStr}`;
    
    let body = `Joom注文の日次総チェックが完了しました。\n\n`;
    body += `【対象日】${dateStr}\n`;
    body += `【チェック日時】${Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd HH:mm:ss')}\n\n`;
    body += `【結果】\n`;
    body += `取得件数: ${reportData.fetched}件\n`;
    body += `登録成功: ${reportData.inserted}件\n`;
    body += `登録失敗: ${reportData.failed}件\n`;
    body += `処理時間: ${reportData.duration.toFixed(1)}秒\n\n`;
    
    if (reportData.errors && reportData.errors.length > 0) {
      body += `【エラー詳細】\n`;
      reportData.errors.forEach((error, index) => {
        body += `${index + 1}. ${JSON.stringify(error)}\n`;
      });
      body += `\n`;
    }
    
    body += `売上管理シートと設定シートの同期履歴を確認してください。`;
    
    MailApp.sendEmail(emailAddress, subject, body);
    console.log('日次レポートを送信しました:', emailAddress);
    
  } catch (error) {
    console.error('日次レポート送信エラー:', error);
  }
}

/**
 * ==========================================
 * 7. メニュー機能
 * ==========================================
 */

/**
 * 自動同期設定メニューを表示
 */
function showSyncTriggerSettings() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const triggerStatus = checkTriggerStatus();
    
    let message = '【現在の自動同期設定】\n\n';
    
    // エラーチェック
    if (triggerStatus.error) {
      message += '⚠️ トリガー状態の取得中にエラーが発生しました\n';
      message += `エラー: ${triggerStatus.error}\n\n`;
    }
    
    // hourlyTriggerの安全なチェック
    if (triggerStatus.hourlyTrigger) {
      message += `定期処理（1時間間隔）: ${triggerStatus.hourlyTrigger.enabled ? '✅ 有効' : '❌ 無効'}\n`;
      message += `  ${triggerStatus.hourlyTrigger.info}\n\n`;
    } else {
      message += `定期処理（1時間間隔）: ❌ 状態不明\n\n`;
    }
    
    // dailyTriggerの安全なチェック
    if (triggerStatus.dailyTrigger) {
      message += `日次総チェック（毎日午前2時）: ${triggerStatus.dailyTrigger.enabled ? '✅ 有効' : '❌ 無効'}\n`;
      message += `  ${triggerStatus.dailyTrigger.info}\n\n`;
    } else {
      message += `日次総チェック（毎日午前2時）: ❌ 状態不明\n\n`;
    }
    
    message += '自動同期を設定しますか？\n\n';
    message += '「はい」: すべての自動同期を有効化\n';
    message += '「いいえ」: すべての自動同期を無効化\n';
    message += '「キャンセル」: 設定を変更しない';
    
    const response = ui.alert('自動同期設定', message, ui.ButtonSet.YES_NO_CANCEL);
    
    if (response === ui.Button.YES) {
      // 自動同期を有効化
      const result = setupAllTriggers();
      
      if (result.success) {
        ui.alert(
          '設定完了',
          '✅ 自動同期を有効化しました！\n\n' +
          '• 定期処理: 1時間ごとに自動実行\n' +
          '• 日次総チェック: 毎日午前2時に自動実行\n\n' +
          '注文データが自動的に同期されます。',
          ui.ButtonSet.OK
        );
      } else {
        ui.alert('エラー', '自動同期の設定に失敗しました: ' + result.error, ui.ButtonSet.OK);
      }
      
    } else if (response === ui.Button.NO) {
      // 自動同期を無効化
      const result = deleteAllSyncTriggers();
      
      if (result.success) {
        ui.alert(
          '設定完了',
          '✅ 自動同期を無効化しました。\n\n' +
          '手動での注文取得は引き続き利用できます。',
          ui.ButtonSet.OK
        );
      } else {
        ui.alert('エラー', '自動同期の無効化に失敗しました: ' + result.error, ui.ButtonSet.OK);
      }
    }
    
  } catch (error) {
    console.error('自動同期設定メニューエラー:', error);
    ui.alert('エラー', '自動同期設定中にエラーが発生しました: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * 同期状況確認メニューを表示
 */
function showSyncStatus() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const triggerStatus = checkTriggerStatus();
    const lastSyncTime = getSetting('Joom 前回連携時間');
    
    let message = '【Joom注文同期状況】\n\n';
    
    // エラーチェック
    if (triggerStatus.error) {
      message += '⚠️ トリガー状態の取得中にエラーが発生しました\n';
      message += `エラー: ${triggerStatus.error}\n\n`;
    }
    
    message += '【自動同期設定】\n';
    
    // hourlyTriggerの安全なチェック
    if (triggerStatus.hourlyTrigger) {
      message += `定期処理（1時間間隔）: ${triggerStatus.hourlyTrigger.enabled ? '✅ 有効' : '❌ 無効'}\n`;
      message += `  ${triggerStatus.hourlyTrigger.info}\n`;
    } else {
      message += `定期処理（1時間間隔）: ❌ 状態不明\n`;
    }
    
    // dailyTriggerの安全なチェック
    if (triggerStatus.dailyTrigger) {
      message += `日次総チェック（毎日午前2時）: ${triggerStatus.dailyTrigger.enabled ? '✅ 有効' : '❌ 無効'}\n`;
      message += `  ${triggerStatus.dailyTrigger.info}\n`;
    } else {
      message += `日次総チェック（毎日午前2時）: ❌ 状態不明\n`;
    }
    
    message += '\n【前回連携時間】\n';
    message += `最終同期: ${lastSyncTime || '未設定'}\n\n`;
    
    message += '【利用可能な機能】\n';
    message += '• 手動即時処理: メニューから「最新注文を取得」を実行\n';
    message += '• 定期処理: 1時間ごとに自動実行（有効時）\n';
    message += '• 日次総チェック: 毎日午前2時に自動実行（有効時）\n\n';
    
    // 推奨アクションの安全なチェック
    const hourlyEnabled = triggerStatus.hourlyTrigger && triggerStatus.hourlyTrigger.enabled;
    const dailyEnabled = triggerStatus.dailyTrigger && triggerStatus.dailyTrigger.enabled;
    
    if (!hourlyEnabled && !dailyEnabled) {
      message += '【推奨アクション】\n';
      message += '• 「自動同期設定」から自動同期を有効化してください\n';
    }
    
    message += '\n設定シートの同期履歴で詳細を確認できます。';
    
    ui.alert('同期状況', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('同期状況確認エラー:', error);
    ui.alert('エラー', '同期状況確認中にエラーが発生しました: ' + error.message, ui.ButtonSet.OK);
  }
}

