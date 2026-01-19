/**
 * ウェブスクレイピング結果フォームハンドラー
 * Googleフォーム経由でアップロードされたCSVを処理し、在庫管理シートを更新する
 * 在庫管理ツール システム開発
 * 作成日: 2025-12-13
 */

/**
 * フォーム送信時トリガー
 * GoogleフォームにCSVファイルがアップロードされたときに自動実行される
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e - フォーム送信イベント
 */
function onFormSubmit(e) {
  try {
    console.log('フォーム送信イベントを受信しました');
    
    // フォーム添付ファイル（CSV）のIDを取得
    const formResponse = e.response;
    const itemResponses = formResponse.getItemResponses();
    
    let csvFileId = null;
    
    // ファイルアップロード項目を探す
    for (let i = 0; i < itemResponses.length; i++) {
      const itemResponse = itemResponses[i];
      const item = itemResponse.getItem();
      
      if (item.getType() === FormApp.ItemType.FILE_UPLOAD) {
        const fileIds = itemResponse.getResponse();
        if (fileIds && fileIds.length > 0) {
          csvFileId = fileIds[0]; // 最初のファイルIDを取得
          break;
        }
      }
    }
    
    if (!csvFileId) {
      console.error('CSVファイルが見つかりません');
      return;
    }
    
    console.log('CSVファイルIDを取得しました:', csvFileId);
    
    // CSVファイルを取得してパース
    const csvFile = DriveApp.getFileById(csvFileId);
    const csvContent = csvFile.getBlob().getDataAsString('UTF-8');
    
    console.log('CSVファイルを読み込みました');
    
    // 既存のparseCsvWithMultilineSupport関数を使用してCSVをパース
    const csvData = parseCsvWithMultilineSupport(csvContent);
    
    if (!csvData || csvData.length === 0) {
      console.error('CSVデータのパースに失敗しました');
      return;
    }
    
    console.log(`CSVデータをパースしました: ${csvData.length}行`);
    
    // ヘッダー行を取得
    const headers = csvData[0];
    
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
      return;
    }
    
    console.log('列インデックス:', columnIndexes);
    
    // スプレッドシートを取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
    
    if (!inventorySheet) {
      console.error('在庫管理シートが見つかりません');
      return;
    }
    
    // 在庫管理シートのデータを取得
    const sheetData = inventorySheet.getDataRange().getValues();
    const sheetHeaders = sheetData[0];
    
    // 仕入れ元URL列のインデックスを取得（ヘッダーから動的に取得）
    const sheetSupplierUrlIndex = sheetHeaders.indexOf('仕入れ元URL');
    
    if (sheetSupplierUrlIndex === -1) {
      console.error('在庫管理シートに「仕入れ元URL」列が見つかりません');
      return;
    }
    
    // 更新対象列のインデックスを取得（Constants.gsの定数を使用）
    const purchasePriceCol = COLUMN_INDEXES.INVENTORY.PURCHASE_PRICE; // G列（7列目）
    const stockStatusCol = COLUMN_INDEXES.INVENTORY.STOCK_STATUS; // S列（19列目）
    const lastUpdatedCol = COLUMN_INDEXES.INVENTORY.LAST_UPDATED; // Z列（26列目）
    
    console.log('更新対象列:', {
      purchasePrice: purchasePriceCol,
      stockStatus: stockStatusCol,
      lastUpdated: lastUpdatedCol
    });
    
    /**
     * URLを正規化する関数
     * 末尾のスラッシュを削除し、URLを統一形式に変換
     */
    function normalizeUrl(url) {
      if (!url || typeof url !== 'string') {
        return '';
      }
      let normalized = url.trim();
      // 末尾のスラッシュを削除
      if (normalized.endsWith('/')) {
        normalized = normalized.slice(0, -1);
      }
      return normalized;
    }
    
    // CSVデータをURLをキーとしたMapに変換
    const csvMap = new Map();
    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      const supplierUrl = row[columnIndexes.supplierUrl];
      
      if (supplierUrl && supplierUrl.trim() !== '') {
        const normalizedUrl = normalizeUrl(supplierUrl);
        const purchasePriceRaw = columnIndexes.purchasePrice !== -1 ? row[columnIndexes.purchasePrice] : '';
        
        // 仕入れ価格を数値に変換（-1の場合は無効値として扱う）
        let purchasePrice = '';
        if (purchasePriceRaw !== '' && purchasePriceRaw !== undefined && purchasePriceRaw !== null) {
          const priceNum = parseFloat(purchasePriceRaw);
          // -1は「価格が取得できなかった」ことを示すので、更新しない
          if (!isNaN(priceNum) && priceNum >= 0) {
            purchasePrice = priceNum;
          }
        }
        
        csvMap.set(normalizedUrl, {
          purchasePrice: purchasePrice,
          stockStatus: columnIndexes.stockStatus !== -1 ? row[columnIndexes.stockStatus] : '',
          lastUpdated: columnIndexes.lastUpdated !== -1 ? row[columnIndexes.lastUpdated] : ''
        });
        
        console.log(`CSV行[${i}]: URL=${normalizedUrl}, 価格=${purchasePrice !== '' ? purchasePrice : '(更新なし)'}, ステータス=${columnIndexes.stockStatus !== -1 ? row[columnIndexes.stockStatus] : '(なし)'}`);
      }
    }
    
    console.log(`CSVデータをMapに変換しました: ${csvMap.size}件`);
    
    // 在庫管理シートを更新
    let updateCount = 0;
    let notFoundCount = 0;
    let priceUpdateCount = 0;
    let statusUpdateCount = 0;
    let dateUpdateCount = 0;
    
    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      const sheetSupplierUrl = row[sheetSupplierUrlIndex];
      
      if (!sheetSupplierUrl || sheetSupplierUrl.trim() === '') {
        continue;
      }
      
      const normalizedSheetUrl = normalizeUrl(sheetSupplierUrl);
      const csvRow = csvMap.get(normalizedSheetUrl);
      
      if (csvRow) {
        // 該当行を更新
        const rowNumber = i + 1; // シートの行番号（1ベース）
        let rowUpdated = false;
        
        // 仕入れ価格を更新（-1の場合は更新しない）
        if (csvRow.purchasePrice !== undefined && csvRow.purchasePrice !== '') {
          const currentPrice = inventorySheet.getRange(rowNumber, purchasePriceCol).getValue();
          inventorySheet.getRange(rowNumber, purchasePriceCol).setValue(csvRow.purchasePrice);
          priceUpdateCount++;
          rowUpdated = true;
          console.log(`行[${rowNumber}]: 仕入れ価格を更新 ${currentPrice} → ${csvRow.purchasePrice} (URL: ${normalizedSheetUrl.substring(0, 50)}...)`);
        }
        
        // 在庫ステータスを更新
        if (csvRow.stockStatus !== undefined && csvRow.stockStatus !== '') {
          const currentStatus = inventorySheet.getRange(rowNumber, stockStatusCol).getValue();
          inventorySheet.getRange(rowNumber, stockStatusCol).setValue(csvRow.stockStatus);
          statusUpdateCount++;
          rowUpdated = true;
          console.log(`行[${rowNumber}]: 在庫ステータスを更新 ${currentStatus} → ${csvRow.stockStatus}`);
        }
        
        // 最終更新日時を更新
        if (csvRow.lastUpdated !== undefined && csvRow.lastUpdated !== '') {
          inventorySheet.getRange(rowNumber, lastUpdatedCol).setValue(csvRow.lastUpdated);
          dateUpdateCount++;
          rowUpdated = true;
        }
        
        if (rowUpdated) {
          updateCount++;
        }
        
        // Mapから削除（処理済み）
        csvMap.delete(normalizedSheetUrl);
      } else {
        // マッチしなかったURLをログに記録（デバッグ用）
        console.log(`行[${i + 1}]: URLがマッチしませんでした: ${normalizedSheetUrl.substring(0, 50)}...`);
      }
    }
    
    // 見つからなかったURLをログに記録
    if (csvMap.size > 0) {
      notFoundCount = csvMap.size;
      console.warn(`在庫管理シートに見つからなかったURL: ${csvMap.size}件`);
      csvMap.forEach((value, key) => {
        console.warn(`  - ${key}`);
      });
    }
    
    console.log(`更新処理が完了しました:`);
    console.log(`  - 更新行数: ${updateCount}行`);
    console.log(`  - 仕入れ価格更新: ${priceUpdateCount}件`);
    console.log(`  - 在庫ステータス更新: ${statusUpdateCount}件`);
    console.log(`  - 最終更新日時更新: ${dateUpdateCount}件`);
    console.log(`  - マッチしなかったURL: ${notFoundCount}件`);
    
    // CSVファイルをゴミ箱へ移動
    try {
      csvFile.setTrashed(true);
      console.log('CSVファイルをゴミ箱へ移動しました');
    } catch (trashError) {
      console.warn('CSVファイルのゴミ箱移動に失敗しました:', trashError.message);
    }
    
    console.log('フォーム送信処理が正常に完了しました');
    
  } catch (error) {
    console.error('フォーム送信処理中にエラーが発生しました:', error);
    console.error('エラースタック:', error.stack);
    
    // エラーが発生した場合でも、CSVファイルはゴミ箱へ移動しない（デバッグのため）
  }
}

/**
 * トリガーを設定する関数
 * 初回セットアップ時に実行する
 * @param {string} formId - フォームID（必須）
 */
function setupWebScrapingFormTrigger(formId) {
  try {
    // 既存のトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onFormSubmit') {
        ScriptApp.deleteTrigger(trigger);
        console.log('既存のトリガーを削除しました');
      }
    });
    
    // フォームIDが指定されていない場合はエラー
    if (!formId || formId === 'YOUR_FORM_ID_HERE') {
      throw new Error('フォームIDが指定されていません。createWebScrapingForm()でフォームを作成するか、フォームIDを指定してください。');
    }
    
    try {
      const form = FormApp.openById(formId);
      
      // フォーム送信時トリガーを設定
      ScriptApp.newTrigger('onFormSubmit')
        .onFormSubmit(form)
        .create();
      
      console.log('フォーム送信時トリガーを設定しました');
      return true;
    } catch (formError) {
      console.error('フォームの取得に失敗しました。フォームIDを確認してください:', formError);
      console.log('手動でトリガーを設定する場合は、以下の手順を実行してください:');
      console.log('1. 拡張機能 > Apps Script エディタを開く');
      console.log('2. 左側の時計アイコン（トリガー）をクリック');
      console.log('3. 「トリガーを追加」をクリック');
      console.log('4. 関数: onFormSubmit、イベント: フォームから、送信時 を選択');
      console.log('5. 「保存」をクリック');
      throw formError;
    }
    
  } catch (error) {
    console.error('トリガー設定中にエラーが発生しました:', error);
    throw error;
  }
}

/**
 * Googleフォームを作成する関数
 * ウェブスクレイピング結果アップロード用のフォームを自動的に作成する
 * 
 * 注意: Google Apps ScriptのFormApp APIでは、ファイルアップロード項目を
 * プログラムで追加することができません。フォームの基本構造のみを作成し、
 * ファイルアップロード項目は手動で追加する必要があります。
 * 
 * @returns {Object} フォーム情報（id, url, editUrl）
 */
function createWebScrapingForm() {
  try {
    console.log('Googleフォームの作成を開始します...');
    
    // フォームを作成
    const form = FormApp.create('在庫管理スクレイピング結果アップロード');
    console.log('フォームを作成しました:', form.getId());
    
    // フォームの説明を設定
    form.setDescription('スクレイピング結果のCSVファイルをアップロードしてください。\n\n【重要】このフォームにはファイルアップロード項目がまだ追加されていません。\nフォーム編集画面で手動でファイルアップロード項目を追加してください。');
    
    // 注意: FormApp APIではファイルアップロード項目をプログラムで追加できません
    // ユーザーが手動で追加する必要があります
    // 参考: https://developers.google.com/apps-script/reference/forms/form
    
    // フォームの設定を変更
    // 「回答を1回に制限する」を外す（スクレイピング結果を複数回送信するため）
    form.setLimitOneResponsePerUser(false);
    
    // 送信完了時の確認メッセージを設定
    form.setConfirmationMessage('送信が完了しました。ありがとうございました。');
    
    // 「回答を別のシートに収集する」を有効化（回答をスプレッドシートに保存）
    form.setCollectEmail(false); // メールアドレスは収集しない
    
    console.log('フォーム設定を完了しました');
    console.log('注意: ファイルアップロード項目は手動で追加する必要があります');
    
    // フォームIDとURLを取得
    const formId = form.getId();
    const formUrl = form.getPublishedUrl();
    const formEditUrl = form.getEditUrl();
    
    console.log('フォーム作成が完了しました');
    console.log('フォームID:', formId);
    console.log('フォームURL:', formUrl);
    console.log('フォーム編集URL:', formEditUrl);
    
    return {
      id: formId,
      url: formUrl,
      editUrl: formEditUrl
    };
    
  } catch (error) {
    console.error('フォーム作成中にエラーが発生しました:', error);
    console.error('エラースタック:', error.stack);
    throw error;
  }
}

/**
 * Googleフォーム作成メニューを表示する関数
 * カスタムメニューから呼び出される
 */
function showWebScrapingFormCreationMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 確認ダイアログを表示
    const response = ui.alert(
      'Googleフォーム作成',
      'ウェブスクレイピング結果アップロード用のGoogleフォームを作成しますか？\n\n' +
      '作成されるフォームの設定:\n' +
      '・タイトル: 在庫管理スクレイピング結果アップロード\n' +
      '・ファイルアップロード項目（CSV形式、最大1ファイル）\n' +
      '・回答制限なし（複数回送信可能）\n' +
      '・確認メール送信なし',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // フォームを作成
    const formInfo = createWebScrapingForm();
    
    // トリガーを自動設定
    let triggerSetupSuccess = false;
    try {
      setupWebScrapingFormTrigger(formInfo.id);
      triggerSetupSuccess = true;
    } catch (triggerError) {
      console.warn('トリガーの自動設定に失敗しました:', triggerError);
    }
    
    // 結果を表示
    const message = 
      'Googleフォームの作成が完了しました！\n\n' +
      '【フォーム情報】\n' +
      `フォームID: ${formInfo.id}\n` +
      `フォームURL: ${formInfo.url}\n` +
      `フォーム編集URL: ${formInfo.editUrl}\n\n` +
      '【重要】ファイルアップロード項目の追加が必要です\n' +
      'Google Apps ScriptのAPI制限により、ファイルアップロード項目は\n' +
      'プログラムで自動追加できません。以下の手順で手動で追加してください:\n\n' +
      '1. フォーム編集URLを開く\n' +
      '2. 「+」ボタンをクリック\n' +
      '3. 「ファイルアップロード」を選択\n' +
      '4. 質問タイトルを「CSVファイルをアップロード」に設定\n' +
      '5. ファイルの種類を「CSV」に設定（オプション）\n' +
      '6. 最大ファイル数を「1」に設定\n' +
      '7. 必須項目にチェックを入れる\n\n' +
      (triggerSetupSuccess 
        ? '✅ フォーム送信時トリガーも自動設定されました。\n\n'
        : '⚠️ フォーム送信時トリガーの自動設定に失敗しました。\n手動で設定してください。\n\n') +
      '【次のステップ】\n' +
      '1. 上記の手順でファイルアップロード項目を追加\n' +
      '2. フォームURLをコピーして、.envファイルのGOOGLE_FORM_URLに設定\n' +
      '3. フォーム編集URLからフォームの詳細設定を確認・変更\n' +
      (triggerSetupSuccess 
        ? '4. トリガーは既に設定済みです\n'
        : '4. setupWebScrapingFormTrigger()関数を実行してトリガーを設定\n');
    
    ui.alert('フォーム作成完了', message, ui.ButtonSet.OK);
    
    // フォームURLをクリップボードにコピーするためのHTMLダイアログを表示
    const htmlTemplate = HtmlService.createTemplate(`
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            .url-box { 
              background-color: #f5f5f5; 
              padding: 10px; 
              border-radius: 4px; 
              margin: 10px 0;
              word-break: break-all;
            }
            button { 
              background-color: #4285f4; 
              color: white; 
              border: none; 
              padding: 10px 20px; 
              border-radius: 4px; 
              cursor: pointer;
              margin: 5px;
            }
            button:hover { background-color: #357ae8; }
          </style>
        </head>
        <body>
          <h3>フォーム情報</h3>
          <p><strong>フォームID:</strong></p>
          <div class="url-box">${formInfo.id}</div>
          <button onclick="copyToClipboard('${formInfo.id}')">フォームIDをコピー</button>
          
          <p><strong>フォームURL:</strong></p>
          <div class="url-box">${formInfo.url}</div>
          <button onclick="copyToClipboard('${formInfo.url}')">フォームURLをコピー</button>
          
          <p><strong>フォーム編集URL:</strong></p>
          <div class="url-box">${formInfo.editUrl}</div>
          <button onclick="copyToClipboard('${formInfo.editUrl}')">編集URLをコピー</button>
          
          <hr>
          <div style="background-color: #fff3cd; padding: 15px; border-radius: 4px; margin: 15px 0; border-left: 4px solid #ffc107;">
            <h4 style="margin-top: 0; color: #856404;">⚠️ 重要: ファイルアップロード項目の追加が必要です</h4>
            <p style="margin-bottom: 5px;"><strong>Google Apps ScriptのAPI制限により、ファイルアップロード項目はプログラムで自動追加できません。</strong></p>
            <p style="margin-bottom: 10px;">以下の手順で手動で追加してください:</p>
            <ol style="margin-top: 5px; padding-left: 20px;">
              <li>フォーム編集URLを開く（上記の「編集URLをコピー」ボタンでコピーできます）</li>
              <li>「+」ボタンをクリック</li>
              <li>「ファイルアップロード」を選択</li>
              <li>質問タイトルを「CSVファイルをアップロード」に設定</li>
              <li>ファイルの種類を「CSV」に設定（オプション）</li>
              <li>最大ファイル数を「1」に設定</li>
              <li>必須項目にチェックを入れる</li>
            </ol>
          </div>
          <p><small>※ ファイルアップロード項目を追加した後、フォームURLを.envファイルのGOOGLE_FORM_URLに設定してください</small></p>
          
          <script>
            function copyToClipboard(text) {
              const textarea = document.createElement('textarea');
              textarea.value = text;
              document.body.appendChild(textarea);
              textarea.select();
              document.execCommand('copy');
              document.body.removeChild(textarea);
              alert('コピーしました: ' + text);
            }
          </script>
        </body>
      </html>
    `);
    
    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(600)
      .setHeight(500);
    
    ui.showModalDialog(htmlOutput, 'フォーム作成完了 - URLをコピー');
    
  } catch (error) {
    console.error('フォーム作成メニューでエラーが発生しました:', error);
    ui.alert(
      'エラー',
      'フォーム作成中にエラーが発生しました:\n' + error.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * トリガー設定用のラッパー関数
 * Apps Scriptエディタから実行する場合に使用
 * フォームIDを指定して実行してください
 * 
 * 使用方法:
 * 1. この関数内の`YOUR_FORM_ID_HERE`を実際のフォームIDに置き換える
 * 2. 関数を選択して「実行」ボタンをクリック
 */
function setupTrigger() {
  const formId = 'YOUR_FORM_ID_HERE'; // 実際のフォームIDに置き換える
  setupWebScrapingFormTrigger(formId);
}

/**
 * フォームURLからフォームIDを抽出する関数
 * @param {string} formUrl - フォームURL（/viewformまたは/editで終わる）
 * @returns {string} フォームID
 */
function extractFormIdFromUrl(formUrl) {
  if (!formUrl || typeof formUrl !== 'string') {
    throw new Error('フォームURLが指定されていません');
  }
  
  // URLからフォームIDを抽出
  // 形式: https://docs.google.com/forms/d/e/FORM_ID/viewform
  // または: https://docs.google.com/forms/d/e/FORM_ID/edit
  const match = formUrl.match(/\/d\/e\/([^\/]+)/);
  if (!match || !match[1]) {
    throw new Error('フォームURLからフォームIDを抽出できませんでした。URLの形式を確認してください。');
  }
  
  return match[1];
}

/**
 * 既存のフォームの確認メッセージを修正する関数
 * 「false」が表示される問題を修正するために使用
 * 
 * 使用方法（方法1: フォームIDを直接指定）:
 * 1. この関数内の`YOUR_FORM_ID_HERE`を実際のフォームIDに置き換える
 * 2. 関数を選択して「実行」ボタンをクリック
 * 
 * 使用方法（方法2: フォームURLから自動抽出）:
 * 1. この関数内の`YOUR_FORM_URL_HERE`を実際のフォームURLに置き換える
 * 2. `useUrl`を`true`に設定
 * 3. 関数を選択して「実行」ボタンをクリック
 */
function fixFormConfirmationMessage() {
  const useUrl = false; // trueに設定すると、フォームURLから自動抽出
  const formId = 'YOUR_FORM_ID_HERE'; // 実際のフォームIDに置き換える（useUrl=falseの場合）
  const formUrl = 'YOUR_FORM_URL_HERE'; // 実際のフォームURLに置き換える（useUrl=trueの場合）
  // 例: 'https://docs.google.com/forms/d/e/1FAIpQLSdpa1BhsNC3w7QpkoJe4EVDKktMxfa_eVPSKOfXNw3L8aim9g/viewform'
  
  try {
    let targetFormId = formId;
    
    if (useUrl && formUrl !== 'YOUR_FORM_URL_HERE') {
      targetFormId = extractFormIdFromUrl(formUrl);
      console.log(`フォームURLからフォームIDを抽出しました: ${targetFormId}`);
    }
    
    if (targetFormId === 'YOUR_FORM_ID_HERE') {
      throw new Error('フォームIDまたはフォームURLを指定してください。');
    }
    
    const form = FormApp.openById(targetFormId);
    form.setConfirmationMessage('送信が完了しました。ありがとうございました。');
    
    console.log('✅ フォームの確認メッセージを更新しました');
    console.log('フォームID:', targetFormId);
    console.log('フォームURL:', form.getPublishedUrl());
    console.log('フォーム編集URL:', form.getEditUrl());
    
    return {
      success: true,
      formId: targetFormId,
      formUrl: form.getPublishedUrl(),
      editUrl: form.getEditUrl()
    };
  } catch (error) {
    console.error('❌ フォームの確認メッセージ更新中にエラーが発生しました:', error);
    throw error;
  }
}
