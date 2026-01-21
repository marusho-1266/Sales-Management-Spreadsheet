/**
 * ウェブスクレイピング結果直接更新ハンドラー
 * URLパラメータ経由でCSVデータを受け取り、スプレッドシートを直接更新する
 * Pythonから直接呼び出せるようにする
 * 
 * 使用方法:
 * 1. この関数をGoogle Apps Scriptエディタで実行可能にする
 * 2. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」として公開
 * 3. Python側からURLにアクセスして関数を実行
 * 4. CSVデータをURLパラメータとして送信
 */

/**
 * Webアプリのエントリーポイント（POSTリクエスト）
 * Python側からPOSTリクエストでCSVデータを受信する
 * 
 * @param {GoogleAppsScript.Events.DoPost} e - POSTリクエストイベント
 * @returns {TextOutput} JSONレスポンス
 */
function doPost(e) {
  try {
    console.log('=== doPost関数が呼び出されました ===');
    
    // POSTボディからCSVデータを取得
    let csvContent = null;
    
    // JSON形式のPOSTボディの場合
    if (e.postData && e.postData.type === 'application/json') {
      try {
        const jsonData = JSON.parse(e.postData.contents);
        csvContent = jsonData.csvData;
        console.log('JSON形式のPOSTボディからCSVデータを取得しました（長さ:', csvContent ? csvContent.length : 0, '文字）');
      } catch (parseError) {
        console.error('JSON解析エラー:', parseError);
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'JSON解析に失敗しました: ' + parseError.message
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    // プレーンテキストの場合（フォールバック）
    else if (e.postData && e.postData.contents) {
      csvContent = e.postData.contents;
      console.log('プレーンテキストのPOSTボディからCSVデータを取得しました（長さ:', csvContent.length, '文字）');
    }
    
    if (!csvContent || csvContent.trim() === '') {
      const errorMsg = 'POSTボディにCSVデータが含まれていません';
      console.error(errorMsg);
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: errorMsg
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // CSVをパース
    let csvRows;
    try {
      csvRows = parseCsvWithMultilineSupport(csvContent);
      console.log('CSVデータをパースしました:', csvRows.length, '行');
    } catch (parseError) {
      console.error('CSVパースエラー:', parseError);
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'CSVパースに失敗しました: ' + parseError.message
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (!csvRows || csvRows.length === 0) {
      const errorMsg = 'CSVデータが空です';
      console.error(errorMsg);
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: errorMsg
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // スプレッドシートを更新
    const result = updateInventorySheet(csvRows);
    
    // 結果をJSON形式で返す
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('doPost関数でエラーが発生しました:', error, error.stack);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'サーバーエラーが発生しました: ' + error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Webアプリのエントリーポイント（GETリクエスト）
 * Google Apps ScriptのWebアプリとして公開する際に必要
 * 既存の互換性のために維持
 * 
 * @param {GoogleAppsScript.Events.DoGet} e - GETリクエストイベント
 * @returns {HtmlOutput|TextOutput} レスポンス
 */
function doGet(e) {
  try {
    console.log('=== doGet関数が呼び出されました ===');
    console.log('パラメータ:', JSON.stringify(e.parameter));
    
    // 方法1: ファイルIDからCSVを読み込む（推奨）
    const fileId = e.parameter.fileId;
    if (fileId) {
      console.log('ファイルIDからCSVを読み込みます:', fileId);
      try {
        const csvFile = DriveApp.getFileById(fileId);
        const csvContent = csvFile.getBlob().getDataAsString('UTF-8');
        console.log('CSVファイルを読み込みました（長さ:', csvContent.length, '文字）');
        
        const csvRows = parseCsvWithMultilineSupport(csvContent);
        console.log('CSVデータをパースしました:', csvRows.length, '行');
        
        const result = updateInventorySheet(csvRows);
        
        // ファイルを削除（成功時のみ）
        if (result && result.success) {
          try {
            csvFile.setTrashed(true);
            console.log('CSVファイルをゴミ箱へ移動しました');
          } catch (trashError) {
            console.warn('ファイル削除エラー（無視）:', trashError);
          }
        } else {
          console.log('更新が失敗したため、CSVファイルを保持します（リトライのため）');
        }
        
        return ContentService.createTextOutput(JSON.stringify(result))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (fileError) {
        console.error('ファイル読み込みエラー:', fileError);
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'ファイル読み込みに失敗しました: ' + fileError.message
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // 方法2: URLパラメータから直接CSVデータを取得（小規模データ用）
    const csvData = e.parameter.csvData;
    if (csvData) {
      console.log('URLパラメータからCSVデータを取得します（長さ:', csvData.length, '文字）');
      
      // URLデコード
      let decodedCsv;
      try {
        decodedCsv = decodeURIComponent(csvData);
        console.log('CSVデータをデコードしました');
      } catch (decodeError) {
        console.error('URLデコードエラー:', decodeError);
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'URLデコードに失敗しました: ' + decodeError.message
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // CSVをパース
      let csvRows;
      try {
        csvRows = parseCsvWithMultilineSupport(decodedCsv);
        console.log('CSVデータをパースしました:', csvRows.length, '行');
      } catch (parseError) {
        console.error('CSVパースエラー:', parseError);
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'CSVパースに失敗しました: ' + parseError.message
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      if (!csvRows || csvRows.length === 0) {
        const errorMsg = 'CSVデータが空です';
        console.error(errorMsg);
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: errorMsg
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // スプレッドシートを更新
      const result = updateInventorySheet(csvRows);
      
      // 結果をJSON形式で返す
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // パラメータがない場合
    const errorMsg = 'fileIdまたはcsvDataパラメータが必要です';
    console.error(errorMsg);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: errorMsg
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('doGet関数でエラーが発生しました:', error, error.stack);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'サーバーエラーが発生しました'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * URLパラメータ経由でスプレッドシートを更新する（内部関数）
 * doGet関数から呼び出される
 * 
 * @param {string} csvData - CSVデータ（URLエンコード済み）
 * @returns {Object} 更新結果
 */
function updateSpreadsheetFromUrl(csvData) {
  try {
    // URLデコード
    const decodedCsv = decodeURIComponent(csvData);
    
    // CSVをパース
    const csvRows = parseCsvWithMultilineSupport(decodedCsv);
    
    if (!csvRows || csvRows.length === 0) {
      console.error('CSVデータのパースに失敗しました');
      return { success: false, error: 'CSVパース失敗' };
    }
    
    // スプレッドシートを更新
    return updateInventorySheet(csvRows);
    
  } catch (error) {
    console.error('スプレッドシート更新エラー:', error);
    return { success: false, error: error.message };
  }
}

/**
 * CSVデータをパースしてスプレッドシートを更新する
 * 
 * @param {Array<Array<string>>} csvRows - パース済みCSVデータ
 * @returns {Object} 更新結果
 */
function updateInventorySheet(csvRows) {
  try {
    // ヘッダー行を取得
    const headers = csvRows[0];
    
    // 列名からインデックスを取得
    const columnIndexes = {
      supplierUrl: headers.indexOf('仕入れ元URL'),
      purchasePrice: headers.indexOf('仕入れ価格'),
      stockStatus: headers.indexOf('在庫ステータス'),
      lastUpdated: headers.indexOf('最終更新日時')
    };
    
    // 必須列の確認
    if (columnIndexes.supplierUrl === -1) {
      console.error('CSVに「仕入れ元URL」列が見つかりません');
      return { success: false, error: '仕入れ元URL列が見つかりません' };
    }
    
    // スプレッドシートを取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
    
    if (!inventorySheet) {
      console.error('在庫管理シートが見つかりません');
      return { success: false, error: '在庫管理シートが見つかりません' };
    }
    
    // 在庫管理シートのデータを取得
    const sheetData = inventorySheet.getDataRange().getValues();
    const sheetHeaders = sheetData[0];
    
    // 仕入れ元URL列のインデックスを取得
    const sheetSupplierUrlIndex = sheetHeaders.indexOf('仕入れ元URL');
    
    if (sheetSupplierUrlIndex === -1) {
      console.error('在庫管理シートに「仕入れ元URL」列が見つかりません');
      return { success: false, error: '在庫管理シートに仕入れ元URL列が見つかりません' };
    }
    
    // 更新対象列のインデックスを取得
    const purchasePriceCol = COLUMN_INDEXES.INVENTORY.PURCHASE_PRICE; // G列（7列目）
    const stockStatusCol = COLUMN_INDEXES.INVENTORY.STOCK_STATUS; // S列（19列目）
    const lastUpdatedCol = COLUMN_INDEXES.INVENTORY.LAST_UPDATED; // Z列（26列目）
    
    // URL正規化関数
    function normalizeUrl(url) {
      if (!url || typeof url !== 'string') {
        return '';
      }
      let normalized = url.trim();
      // 末尾のスラッシュを削除
      if (normalized.endsWith('/')) {
        normalized = normalized.slice(0, -1);
      }
      // URLデコードを試行（エンコードされた文字をデコード）
      try {
        normalized = decodeURIComponent(normalized);
      } catch (e) {
        // デコードに失敗した場合は元のURLを使用
        console.log('URLデコードに失敗しました（無視して続行）:', e.message);
      }
      // 再度エンコードして統一（クエリパラメータの順序を正規化するため）
      try {
        const urlObj = new URL(normalized);
        // クエリパラメータをソートして正規化
        const sortedParams = new URLSearchParams(urlObj.search);
        sortedParams.sort();
        urlObj.search = sortedParams.toString();
        normalized = urlObj.toString();
      } catch (e) {
        // URLオブジェクトの作成に失敗した場合は元のURLを使用
        console.log('URLオブジェクトの作成に失敗しました（無視して続行）:', e.message);
      }
      return normalized;
    }
    
    // URL類似度を計算する関数（デバッグ用）
    function calculateUrlSimilarity(url1, url2) {
      if (!url1 || !url2) return 0;
      
      // ベースURL（クエリパラメータを除く）を比較
      try {
        const base1 = url1.split('?')[0];
        const base2 = url2.split('?')[0];
        if (base1 === base2) return 0.8;  // ベースURLが同じなら類似度0.8
      } catch (e) {
        // 無視
      }
      
      // 文字列の類似度を計算（簡易版）
      const longer = url1.length > url2.length ? url1 : url2;
      const editDistance = levenshteinDistance(url1, url2);
      return 1 - (editDistance / longer.length);
    }
    
    // レーベンシュタイン距離を計算する関数
    function levenshteinDistance(str1, str2) {
      const len1 = str1.length;
      const len2 = str2.length;
      const matrix = [];
      
      for (let i = 0; i <= len1; i++) {
        matrix[i] = [i];
      }
      
      for (let j = 0; j <= len2; j++) {
        matrix[0][j] = j;
      }
      
      for (let i = 1; i <= len1; i++) {
        for (let j = 1; j <= len2; j++) {
          if (str1[i - 1] === str2[j - 1]) {
            matrix[i][j] = matrix[i - 1][j - 1];
          } else {
            matrix[i][j] = Math.min(
              matrix[i - 1][j] + 1,
              matrix[i][j - 1] + 1,
              matrix[i - 1][j - 1] + 1
            );
          }
        }
      }
      
      return matrix[len1][len2];
    }
    
    // CSVデータをURLをキーとしたMapに変換
    const csvMap = new Map();
    for (let i = 1; i < csvRows.length; i++) {
      const row = csvRows[i];
      const supplierUrl = row[columnIndexes.supplierUrl];
      
      if (supplierUrl && supplierUrl.trim() !== '') {
        const normalizedUrl = normalizeUrl(supplierUrl);
        const purchasePriceRaw = columnIndexes.purchasePrice !== -1 ? row[columnIndexes.purchasePrice] : '';
        
        // 仕入れ価格を数値に変換（-1の場合は無効値として扱う）
        let purchasePrice = '';
        if (purchasePriceRaw !== '' && purchasePriceRaw !== undefined && purchasePriceRaw !== null) {
          const priceNum = parseFloat(purchasePriceRaw);
          if (!isNaN(priceNum) && priceNum >= 0) {
            purchasePrice = priceNum;
          }
        }
        
        csvMap.set(normalizedUrl, {
          purchasePrice: purchasePrice,
          stockStatus: columnIndexes.stockStatus !== -1 ? row[columnIndexes.stockStatus] : '',
          lastUpdated: columnIndexes.lastUpdated !== -1 ? row[columnIndexes.lastUpdated] : '',
          originalUrl: supplierUrl  // デバッグ用に元のURLも保存
        });
        
        console.log(`CSV行[${i}]: 元のURL=${supplierUrl.substring(0, 80)}...`);
        console.log(`CSV行[${i}]: 正規化URL=${normalizedUrl.substring(0, 80)}...`);
      }
    }
    
    console.log(`CSVデータをMapに変換しました: ${csvMap.size}件`);
    
    // 在庫管理シートを更新（バッチ書き込み用の更新データを蓄積）
    let updateCount = 0;
    let priceUpdateCount = 0;
    let statusUpdateCount = 0;
    let dateUpdateCount = 0;
    let notFoundCount = 0;
    
    // 更新データを蓄積するMap（行番号 -> 値）
    const purchasePriceUpdates = new Map();
    const stockStatusUpdates = new Map();
    const lastUpdatedUpdates = new Map();
    
    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      const sheetSupplierUrl = row[sheetSupplierUrlIndex];
      
      if (!sheetSupplierUrl || sheetSupplierUrl.trim() === '') {
        continue;
      }
      
      const normalizedSheetUrl = normalizeUrl(sheetSupplierUrl);
      const csvRow = csvMap.get(normalizedSheetUrl);
      
      // デバッグ用: マッチしない場合のログ
      if (!csvRow) {
        console.log(`行[${i + 1}]: マッチしませんでした`);
        console.log(`  スプレッドシートURL: ${sheetSupplierUrl.substring(0, 80)}...`);
        console.log(`  正規化URL: ${normalizedSheetUrl.substring(0, 80)}...`);
        // CSV内のすべてのURLと比較して、類似度を確認
        csvMap.forEach((value, key) => {
          const similarity = calculateUrlSimilarity(normalizedSheetUrl, key);
          if (similarity > 0.5) {
            console.log(`  類似URL発見 (類似度: ${similarity}): ${key.substring(0, 80)}...`);
          }
        });
      }
      
      if (csvRow) {
        const rowNumber = i + 1;
        let rowUpdated = false;
        
        // 仕入れ価格を更新（配列に蓄積）
        if (csvRow.purchasePrice !== undefined && csvRow.purchasePrice !== '') {
          purchasePriceUpdates.set(rowNumber, csvRow.purchasePrice);
          priceUpdateCount++;
          rowUpdated = true;
        }
        
        // 在庫ステータスを更新（配列に蓄積）
        if (csvRow.stockStatus !== undefined && csvRow.stockStatus !== '') {
          stockStatusUpdates.set(rowNumber, csvRow.stockStatus);
          statusUpdateCount++;
          rowUpdated = true;
        }
        
        // 最終更新日時を更新（配列に蓄積）
        if (csvRow.lastUpdated !== undefined && csvRow.lastUpdated !== '') {
          lastUpdatedUpdates.set(rowNumber, csvRow.lastUpdated);
          dateUpdateCount++;
          rowUpdated = true;
        }
        
        if (rowUpdated) {
          updateCount++;
        }
        
        csvMap.delete(normalizedSheetUrl);
      }
    }
    
    // バッチ書き込み: 仕入れ価格
    if (purchasePriceUpdates.size > 0) {
      const sortedRows = Array.from(purchasePriceUpdates.keys()).sort((a, b) => a - b);
      const startRow = sortedRows[0];
      const endRow = sortedRows[sortedRows.length - 1];
      const numRows = endRow - startRow + 1;
      
      // 現在の値を取得して、更新が必要な行のみ新しい値で置き換え
      const currentValues = inventorySheet.getRange(startRow, purchasePriceCol, numRows, 1).getValues();
      const purchasePriceValues = currentValues.map((value, index) => {
        const rowNum = startRow + index;
        return purchasePriceUpdates.has(rowNum) ? [purchasePriceUpdates.get(rowNum)] : value;
      });
      
      inventorySheet.getRange(startRow, purchasePriceCol, numRows, 1).setValues(purchasePriceValues);
    }
    
    // バッチ書き込み: 在庫ステータス
    if (stockStatusUpdates.size > 0) {
      const sortedRows = Array.from(stockStatusUpdates.keys()).sort((a, b) => a - b);
      const startRow = sortedRows[0];
      const endRow = sortedRows[sortedRows.length - 1];
      const numRows = endRow - startRow + 1;
      
      // 現在の値を取得して、更新が必要な行のみ新しい値で置き換え
      const currentValues = inventorySheet.getRange(startRow, stockStatusCol, numRows, 1).getValues();
      const stockStatusValues = currentValues.map((value, index) => {
        const rowNum = startRow + index;
        return stockStatusUpdates.has(rowNum) ? [stockStatusUpdates.get(rowNum)] : value;
      });
      
      inventorySheet.getRange(startRow, stockStatusCol, numRows, 1).setValues(stockStatusValues);
    }
    
    // バッチ書き込み: 最終更新日時
    if (lastUpdatedUpdates.size > 0) {
      const sortedRows = Array.from(lastUpdatedUpdates.keys()).sort((a, b) => a - b);
      const startRow = sortedRows[0];
      const endRow = sortedRows[sortedRows.length - 1];
      const numRows = endRow - startRow + 1;
      
      // 現在の値を取得して、更新が必要な行のみ新しい値で置き換え
      const currentValues = inventorySheet.getRange(startRow, lastUpdatedCol, numRows, 1).getValues();
      const lastUpdatedValues = currentValues.map((value, index) => {
        const rowNum = startRow + index;
        return lastUpdatedUpdates.has(rowNum) ? [lastUpdatedUpdates.get(rowNum)] : value;
      });
      
      inventorySheet.getRange(startRow, lastUpdatedCol, numRows, 1).setValues(lastUpdatedValues);
    }
    
    // 見つからなかったURLをカウント
    notFoundCount = csvMap.size;
    
    // 見つからなかったURLのリストを作成（デバッグ用）
    const notFoundUrls = [];
    csvMap.forEach((value, key) => {
      notFoundUrls.push({
        normalizedUrl: key.substring(0, 100),
        originalUrl: value.originalUrl ? value.originalUrl.substring(0, 100) : key.substring(0, 100)
      });
    });
    
    const result = {
      success: true,
      updateCount: updateCount,
      priceUpdateCount: priceUpdateCount,
      statusUpdateCount: statusUpdateCount,
      dateUpdateCount: dateUpdateCount,
      notFoundCount: notFoundCount,
      notFoundUrls: notFoundUrls  // デバッグ用
    };
    
    console.log('更新処理が完了しました:', result);
    if (notFoundCount > 0) {
      console.warn(`マッチしなかったURL: ${notFoundCount}件`);
      notFoundUrls.forEach((item, index) => {
        console.warn(`  [${index + 1}] ${item.originalUrl}...`);
      });
    }
    return result;
    
  } catch (error) {
    console.error('スプレッドシート更新エラー:', error);
    return { success: false, error: error.message };
  }
}

/**
 * テスト用関数: CSVデータを直接渡して更新をテストする
 * 
 * @param {string} csvContent - CSVデータ（文字列）
 */
function testUpdateFromCsv(csvContent) {
  const csvRows = parseCsvWithMultilineSupport(csvContent);
  return updateInventorySheet(csvRows);
}

/**
 * デバッグ用: Webアプリが正しく動作しているか確認する
 * 
 * @returns {TextOutput} テスト結果
 */
function testWebApp() {
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: 'Webアプリは正常に動作しています',
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * デバッグ用: 簡単なテスト用doGet関数
 * この関数をdoGetとして使用して、Webアプリが正しく動作するか確認
 */
function doGetTest(e) {
  console.log('=== doGetTest関数が呼び出されました ===');
  console.log('パラメータ:', JSON.stringify(e.parameter));
  
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: 'doGet関数は正常に動作しています',
    parameters: e.parameter,
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

