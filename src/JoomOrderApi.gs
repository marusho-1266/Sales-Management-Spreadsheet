/**
 * Joom注文連携API機能
 * 在庫管理ツール システム開発
 * 作成日: 2025-10-07
 */

/**
 * ==========================================
 * 1. OAuth 2.0 認証システム
 * ==========================================
 */

/**
 * Joom API認証フローの開始
 */
function authenticateWithJoom() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Client IDの確認
    const clientId = getSetting('Joom Client ID');
    if (!clientId) {
      ui.alert(
        '認証エラー',
        'Joom Client IDが設定されていません。\n設定シートでClient IDを設定してください。',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // 認証URLの生成
    const authUrl = generateJoomAuthUrl(clientId);
    
    // 認証手順の説明ダイアログ
    const instructionResponse = ui.alert(
      'Joom API 認証',
      '以下の手順でJoom APIの認証を行います：\n\n' +
      '1. 次のダイアログに表示されるURLにアクセス\n' +
      '2. Joomアカウントでログイン\n' +
      '3. アクセス許可をクリック\n' +
      '4. 表示された認証コードをコピー\n' +
      '5. 次のダイアログに認証コードを入力\n\n' +
      '準備ができたら「OK」を押してください。',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (instructionResponse !== ui.Button.OK) {
      console.log('認証がキャンセルされました');
      return false;
    }
    
    // 認証URLを表示
    const urlResponse = ui.alert(
      '認証URL',
      '以下のURLにアクセスして認証コードを取得してください：\n\n' +
      authUrl + '\n\n' +
      'URLをコピーしてブラウザで開いてください。',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (urlResponse !== ui.Button.OK) {
      console.log('認証がキャンセルされました');
      return false;
    }
    
    // 認証コードの入力
    const codeResponse = ui.prompt(
      '認証コード入力',
      'Joomから取得した認証コードを入力してください：',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (codeResponse.getSelectedButton() !== ui.Button.OK) {
      console.log('認証がキャンセルされました');
      return false;
    }
    
    const authCode = codeResponse.getResponseText().trim();
    if (!authCode) {
      ui.alert('エラー', '認証コードが入力されていません。', ui.ButtonSet.OK);
      return false;
    }
    
    // アクセストークンの取得
    const success = exchangeCodeForToken(authCode);
    
    if (success) {
      ui.alert(
        '認証成功',
        'Joom APIの認証が完了しました。\n注文データの同期が可能になりました。',
        ui.ButtonSet.OK
      );
      return true;
    } else {
      ui.alert(
        '認証失敗',
        '認証コードの交換に失敗しました。\n認証コードを確認して再試行してください。',
        ui.ButtonSet.OK
      );
      return false;
    }
    
  } catch (error) {
    console.error('認証エラー:', error);
    ui.alert(
      '認証エラー',
      '認証中にエラーが発生しました: ' + error.message,
      ui.ButtonSet.OK
    );
    return false;
  }
}

/**
 * Joom認証URLの生成
 * @param {string} clientId - Joom Client ID
 * @returns {string} 認証URL
 */
function generateJoomAuthUrl(clientId) {
  const baseUrl = 'https://api-merchant.joom.com/api/v2/oauth/authorize';
  const redirectUri = getSetting('Joom Redirect URI') || 'urn:ietf:wg:oauth:2.0:oob';
  
  // GAS用のURLパラメータ構築
  const params = [];
  params.push(`client_id=${encodeURIComponent(clientId)}`);
  params.push(`response_type=code`);
  params.push(`redirect_uri=${encodeURIComponent(redirectUri)}`);
  
  const authUrl = `${baseUrl}?${params.join('&')}`;
  
  // デバッグログ
  console.log('認証URL生成:');
  console.log('- Base URL:', baseUrl);
  console.log('- Client ID:', clientId);
  console.log('- Redirect URI:', redirectUri);
  console.log('- 生成されたURL:', authUrl);
  
  return authUrl;
}

/**
 * 認証コードをアクセストークンに交換
 * @param {string} authCode - 認証コード
 * @returns {boolean} 成功フラグ
 */
function exchangeCodeForToken(authCode) {
  try {
    const clientId = getSetting('Joom Client ID');
    const clientSecret = getSetting('Joom Client Secret');
    
    if (!clientId || !clientSecret) {
      throw new Error('Client IDまたはClient Secretが設定されていません');
    }
    
    const tokenUrl = 'https://api-merchant.joom.com/api/v2/oauth/access_token';
    
    // x-www-form-urlencoded形式でペイロードを作成
    const redirectUri = getSetting('Joom Redirect URI') || 'urn:ietf:wg:oauth:2.0:oob';
    const payload = `client_id=${encodeURIComponent(clientId)}&client_secret=${encodeURIComponent(clientSecret)}&grant_type=authorization_code&code=${encodeURIComponent(authCode)}&redirect_uri=${encodeURIComponent(redirectUri)}`;
    
    const options = {
      method: 'post',
      payload: payload,
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(tokenUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('トークン取得レスポンスコード:', responseCode);
    
    if (responseCode === 200) {
      const tokenData = JSON.parse(responseText);
      
      // レスポンスの検証
      if (tokenData.code === 0 && tokenData.data) {
        const data = tokenData.data;
        
        // トークンを設定シートに保存
        updateSetting('Joom Access Token', data.access_token);
        updateSetting('Joom Refresh Token', data.refresh_token);
        
        // 有効期限を計算して保存
        const expiryTime = new Date(data.expiry_time * 1000);
        updateSetting('Joom Token Expiry', Utilities.formatDate(expiryTime, 'JST', 'yyyy-MM-dd HH:mm:ss'));
        
        console.log('アクセストークン取得成功');
        console.log('有効期限:', expiryTime.toISOString());
        
        return true;
      } else {
        throw new Error(`トークン取得失敗: ${tokenData.message || '不明なエラー'}`);
      }
    } else {
      throw new Error(`HTTPエラー: ${responseCode} - ${responseText}`);
    }
    
  } catch (error) {
    console.error('トークン交換エラー:', error);
    return false;
  }
}

/**
 * ==========================================
 * 2. トークン管理機能
 * ==========================================
 */

/**
 * トークンの有効性チェック
 * @returns {boolean} 有効フラグ
 */
function isJoomTokenValid() {
  try {
    const accessToken = getSetting('Joom Access Token');
    if (!accessToken) {
      console.log('アクセストークンが設定されていません');
      return false;
    }
    
    const expiryStr = getSetting('Joom Token Expiry');
    if (!expiryStr) {
      console.log('トークン有効期限が設定されていません');
      return false;
    }
    
    const expiryDate = new Date(expiryStr);
    const now = new Date();
    
    // 5分前からリフレッシュが必要と判定
    const remainingTime = expiryDate.getTime() - now.getTime();
    const fiveMinutes = 5 * 60 * 1000;
    
    if (remainingTime > fiveMinutes) {
      console.log('トークンは有効です');
      return true;
    } else {
      console.log('トークンの有効期限が近づいています');
      return false;
    }
    
  } catch (error) {
    console.error('トークン有効性チェックエラー:', error);
    return false;
  }
}

/**
 * リフレッシュトークンを使用してアクセストークンを更新
 * @returns {boolean} 成功フラグ
 */
function refreshJoomToken() {
  try {
    console.log('アクセストークンのリフレッシュを開始します...');
    
    const clientId = getSetting('Joom Client ID');
    const clientSecret = getSetting('Joom Client Secret');
    const refreshToken = getSetting('Joom Refresh Token');
    
    if (!clientId || !clientSecret || !refreshToken) {
      throw new Error('認証情報が不足しています');
    }
    
    const tokenUrl = 'https://api-merchant.joom.com/api/v2/oauth/refresh_token';
    
    // x-www-form-urlencoded形式でペイロードを作成
    const redirectUri = getSetting('Joom Redirect URI') || 'urn:ietf:wg:oauth:2.0:oob';
    const payload = `client_id=${encodeURIComponent(clientId)}&client_secret=${encodeURIComponent(clientSecret)}&grant_type=refresh_token&refresh_token=${encodeURIComponent(refreshToken)}&redirect_uri=${encodeURIComponent(redirectUri)}`;
    
    const options = {
      method: 'post',
      payload: payload,
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(tokenUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('トークンリフレッシュレスポンスコード:', responseCode);
    
    if (responseCode === 200) {
      const tokenData = JSON.parse(responseText);
      
      if (tokenData.code === 0 && tokenData.data) {
        const data = tokenData.data;
        
        // 新しいトークンを保存
        updateSetting('Joom Access Token', data.access_token);
        updateSetting('Joom Refresh Token', data.refresh_token);
        
        // 有効期限を計算して保存
        const expiryTime = new Date(data.expiry_time * 1000);
        updateSetting('Joom Token Expiry', Utilities.formatDate(expiryTime, 'JST', 'yyyy-MM-dd HH:mm:ss'));
        
        console.log('トークンリフレッシュ成功');
        console.log('新しい有効期限:', expiryTime.toISOString());
        
        return true;
      } else {
        throw new Error(`トークンリフレッシュ失敗: ${tokenData.message || '不明なエラー'}`);
      }
    } else {
      throw new Error(`HTTPエラー: ${responseCode} - ${responseText}`);
    }
    
  } catch (error) {
    console.error('トークンリフレッシュエラー:', error);
    return false;
  }
}

/**
 * トークンの自動リフレッシュチェック
 * リクエスト前に呼び出して、必要に応じてトークンをリフレッシュ
 * @returns {boolean} トークンが有効かどうか
 */
function ensureValidToken() {
  try {
    if (!isJoomTokenValid()) {
      console.log('トークンが無効または期限切れ間近です。リフレッシュを実行します...');
      
      if (refreshJoomToken()) {
        console.log('トークンリフレッシュ成功');
        return true;
      } else {
        console.error('トークンリフレッシュ失敗');
        return false;
      }
    }
    
    return true;
  } catch (error) {
    console.error('トークン検証エラー:', error);
    return false;
  }
}

/**
 * アクセストークンのテスト
 * @returns {boolean} トークンが有効かどうか
 */
function testJoomToken() {
  try {
    const accessToken = getSetting('Joom Access Token');
    if (!accessToken) {
      console.log('アクセストークンが設定されていません');
      return false;
    }
    
    const testUrl = 'https://api-merchant.joom.com/api/v2/auth_test';
    
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(testUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      console.log('トークンは有効です');
      return true;
    } else {
      console.log('トークンが無効です:', responseCode);
      return false;
    }
    
  } catch (error) {
    console.error('トークンテストエラー:', error);
    return false;
  }
}

/**
 * 認証設定の詳細テスト
 */
function testJoomAuthConfiguration() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const clientId = getSetting('Joom Client ID');
    const clientSecret = getSetting('Joom Client Secret');
    const redirectUri = getSetting('Joom Redirect URI');
    const webAppUrl = getWebAppUrl();
    
    const message = `
【Joom認証設定テスト】

Client ID: ${clientId ? '✅ 設定済み' : '❌ 未設定'}
Client Secret: ${clientSecret ? '✅ 設定済み' : '❌ 未設定'}
Redirect URI: ${redirectUri || '❌ 未設定'}
Web App URL: ${webAppUrl}

【設定確認】
1. Joom設定のRedirect URL: ${redirectUri}
2. GAS Web App URL: ${webAppUrl}
3. URL一致: ${redirectUri === webAppUrl ? '✅ 一致' : '❌ 不一致'}

【次のステップ】
${!clientId || !clientSecret ? '1. 設定シートでClient IDとSecretを設定\n' : ''}
${redirectUri !== webAppUrl ? '2. Joom設定でRedirect URLを更新\n' : ''}
3. GAS Webアプリをデプロイ
4. 認証を実行
    `;
    
    ui.alert('認証設定テスト', message, ui.ButtonSet.OK);
    
    // 認証URLの生成テスト
    if (clientId) {
      const authUrl = generateJoomAuthUrl(clientId);
      console.log('生成された認証URL:', authUrl);
    }
    
  } catch (error) {
    console.error('設定テストエラー:', error);
    ui.alert(
      'エラー',
      '設定テスト中にエラーが発生しました: ' + error.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * ==========================================
 * 3. 基本APIリクエスト機能
 * ==========================================
 */

/**
 * Joom APIへのリクエスト送信
 * @param {string} endpoint - APIエンドポイント（例: '/orders/multi'）
 * @param {string} method - HTTPメソッド（GET, POST等）
 * @param {Object} payload - リクエストボディ（POST時）
 * @returns {Object} APIレスポンス
 */
function makeJoomApiRequest(endpoint, method = 'GET', payload = null) {
  try {
    // トークンの有効性を確認（必要に応じて自動リフレッシュ）
    if (!ensureValidToken()) {
      throw new Error('有効なアクセストークンがありません。認証を実行してください。');
    }
    
    const baseUrl = getSetting('Joom API Base URL') || 'https://api-merchant.joom.com/api/v3';
    const accessToken = getSetting('Joom Access Token');
    
    const url = `${baseUrl}${endpoint}`;
    
    const options = {
      method: method.toLowerCase(),
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    if (payload && (method === 'POST' || method === 'PUT')) {
      options.payload = JSON.stringify(payload);
    }
    
    console.log(`Joom API リクエスト: ${method} ${url}`);
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    // レート制限ヘッダーの取得
    const rateLimitRemaining = response.getHeaders()['X-Rate-Limit-Remaining'];
    const rateLimitReset = response.getHeaders()['X-Rate-Limit-Reset'];
    
    if (rateLimitRemaining !== undefined) {
      console.log(`レート制限残り: ${rateLimitRemaining} リクエスト`);
    }
    
    if (responseCode >= 200 && responseCode < 300) {
      return JSON.parse(responseText);
    } else if (responseCode === 401) {
      throw new Error('認証エラー: アクセストークンが無効です');
    } else if (responseCode === 429) {
      throw new Error(`レート制限超過: ${rateLimitReset}秒後にリトライしてください`);
    } else {
      throw new Error(`API Error: ${responseCode} - ${responseText}`);
    }
    
  } catch (error) {
    console.error('Joom API リクエストエラー:', error);
    throw error;
  }
}

/**
 * ==========================================
 * 4. トークン取得機能
 * ==========================================
 */

/**
 * トークン取得のメイン機能
 */
function acquireJoomToken() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 設定の確認
    const clientId = getSetting('Joom Client ID');
    const clientSecret = getSetting('Joom Client Secret');
    
    if (!clientId || !clientSecret) {
      ui.alert(
        '設定エラー',
        'Joom Client IDまたはClient Secretが設定されていません。\n' +
        '設定シートで認証情報を設定してください。',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // 認証方法の選択
    const methodResponse = ui.alert(
      'トークン取得方法の選択',
      'トークン取得方法を選択してください：\n\n' +
      '1. Webアプリ方式（推奨）\n' +
      '   - 自動的に認証が完了\n' +
      '   - より安全で使いやすい\n\n' +
      '2. 手動入力方式\n' +
      '   - 認証コードを手動で入力\n' +
      '   - 従来の方法',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (methodResponse === ui.Button.YES) {
      // Webアプリ方式
      return acquireTokenWebApp();
    } else if (methodResponse === ui.Button.NO) {
      // 手動入力方式
      return acquireTokenManual();
    } else {
      // キャンセル
      console.log('トークン取得がキャンセルされました');
      return false;
    }
    
  } catch (error) {
    console.error('トークン取得エラー:', error);
    ui.alert(
      'エラー',
      'トークン取得中にエラーが発生しました: ' + error.message,
      ui.ButtonSet.OK
    );
    return false;
  }
}

/**
 * Webアプリ方式でのトークン取得
 */
function acquireTokenWebApp() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const clientId = getSetting('Joom Client ID');
    const webAppUrl = getWebAppUrl();
    
    // Redirect URIを設定
    updateSetting('Joom Redirect URI', webAppUrl);
    
    // 認証URLの生成
    const authUrl = generateJoomAuthUrl(clientId);
    
    // 手順の説明
    const instructionResponse = ui.alert(
      'Webアプリ方式でのトークン取得',
      '以下の手順でトークンを取得します：\n\n' +
      '【事前準備】\n' +
      '1. Joom設定でRedirect URLを設定\n' +
      `   URL: ${webAppUrl}\n\n` +
      '【認証手順】\n' +
      '1. 次のダイアログのURLにアクセス\n' +
      '2. Joomアカウントでログイン\n' +
      '3. アクセス許可をクリック\n' +
      '4. 自動的にトークンが取得されます\n\n' +
      '準備ができたら「OK」を押してください。',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (instructionResponse !== ui.Button.OK) {
      return false;
    }
    
    // 認証URLの表示
    const urlResponse = ui.alert(
      '認証URL',
      '以下のURLにアクセスして認証を完了してください：\n\n' +
      authUrl + '\n\n' +
      '【注意事項】\n' +
      '• URLをコピーしてブラウザで開いてください\n' +
      '• 認証完了後、自動的にスプレッドシートに戻ります\n' +
      '• エラーが出る場合は、Joom設定のRedirect URLを確認してください',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (urlResponse !== ui.Button.OK) {
      return false;
    }
    
    return true;
    
  } catch (error) {
    console.error('Webアプリ方式トークン取得エラー:', error);
    ui.alert(
      'エラー',
      'Webアプリ方式でのトークン取得中にエラーが発生しました: ' + error.message,
      ui.ButtonSet.OK
    );
    return false;
  }
}

/**
 * 手動入力方式でのトークン取得
 */
function acquireTokenManual() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const clientId = getSetting('Joom Client ID');
    
    // ローカルリダイレクトURLを使用
    const localRedirectUri = 'http://localhost:8080/callback';
    updateSetting('Joom Redirect URI', localRedirectUri);
    
    // 認証URLの生成
    const authUrl = generateJoomAuthUrl(clientId);
    
    // 手順の説明
    const instructionResponse = ui.alert(
      '手動入力方式でのトークン取得',
      '以下の手順でトークンを取得します：\n\n' +
      '【重要】事前にJoom設定でRedirect URLを設定してください\n' +
      `Redirect URL: ${localRedirectUri}\n\n` +
      '1. 次のダイアログのURLにアクセス\n' +
      '2. Joomアカウントでログイン\n' +
      '3. アクセス許可をクリック\n' +
      '4. ローカルサーバーにリダイレクトされます（エラーは無視）\n' +
      '5. ブラウザのアドレスバーから認証コードをコピー\n' +
      '6. 次のダイアログに認証コードを入力\n\n' +
      '準備ができたら「OK」を押してください。',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (instructionResponse !== ui.Button.OK) {
      return false;
    }
    
    // 認証URLの表示
    const urlResponse = ui.alert(
      '認証URL',
      '以下のURLにアクセスして認証コードを取得してください：\n\n' +
      authUrl + '\n\n' +
      '【注意事項】\n' +
      '• URLをコピーしてブラウザで開いてください\n' +
      '• 認証後、localhost:8080にリダイレクトされます\n' +
      '• ブラウザのアドレスバーに認証コードが表示されます\n' +
      '• 認証コードは5分で期限切れになります\n' +
      '• 認証コードをコピーしたら速やかに入力してください',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (urlResponse !== ui.Button.OK) {
      return false;
    }
    
    // 認証コードの入力
    const codeResponse = ui.prompt(
      '認証コード入力',
      'Joomから取得した認証コードを入力してください：\n\n' +
      '【入力例】\n' +
      '4/0AX4XfWh...（長い文字列）\n\n' +
      '認証コード：',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (codeResponse.getSelectedButton() !== ui.Button.OK) {
      return false;
    }
    
    const authCode = codeResponse.getResponseText().trim();
    if (!authCode) {
      ui.alert('エラー', '認証コードが入力されていません。', ui.ButtonSet.OK);
      return false;
    }
    
    // トークンの取得
    console.log('認証コード受信:', authCode.substring(0, 20) + '...');
    const success = exchangeCodeForToken(authCode);
    
    if (success) {
      ui.alert(
        'トークン取得成功',
        '✅ Joom APIのトークンが正常に取得されました！\n\n' +
        'アクセストークンとリフレッシュトークンが設定シートに保存されました。\n' +
        'これで注文データの同期が可能になりました。\n\n' +
        '「トークン取得状況」で確認してください。',
        ui.ButtonSet.OK
      );
      return true;
    } else {
      ui.alert(
        'トークン取得失敗',
        '❌ トークンの取得に失敗しました。\n\n' +
        '【可能な原因】\n' +
        '• 認証コードが間違っている\n' +
        '• 認証コードの有効期限が切れている（5分）\n' +
        '• Client IDまたはClient Secretが間違っている\n' +
        '• Joom APIサーバーの一時的な問題\n\n' +
        '【対処法】\n' +
        '1. 認証コードを再取得\n' +
        '2. 設定シートでClient ID/Secretを確認\n' +
        '3. 数分後に再試行',
        ui.ButtonSet.OK
      );
      return false;
    }
    
  } catch (error) {
    console.error('手動入力方式トークン取得エラー:', error);
    ui.alert(
      'エラー',
      '手動入力方式でのトークン取得中にエラーが発生しました: ' + error.message + '\n\n' +
      '【デバッグ情報】\n' +
      'エラーの詳細はログを確認してください。\n' +
      'GASエディタの「実行」→「ログを表示」で確認できます。',
      ui.ButtonSet.OK
    );
    return false;
  }
}

/**
 * トークン取得状況の確認
 */
function checkTokenAcquisitionStatus() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const accessToken = getSetting('Joom Access Token');
    const refreshToken = getSetting('Joom Refresh Token');
    const tokenExpiry = getSetting('Joom Token Expiry');
    const clientId = getSetting('Joom Client ID');
    const clientSecret = getSetting('Joom Client Secret');
    
    let statusMessage = '【Joom APIトークン取得状況】\n\n';
    
    // 基本設定の確認
    statusMessage += '【基本設定】\n';
    statusMessage += `Client ID: ${clientId ? '✅ 設定済み' : '❌ 未設定'}\n`;
    statusMessage += `Client Secret: ${clientSecret ? '✅ 設定済み' : '❌ 未設定'}\n\n`;
    
    // トークンの確認
    statusMessage += '【トークン状況】\n';
    statusMessage += `Access Token: ${accessToken ? '✅ 取得済み' : '❌ 未取得'}\n`;
    statusMessage += `Refresh Token: ${refreshToken ? '✅ 取得済み' : '❌ 未取得'}\n`;
    statusMessage += `有効期限: ${tokenExpiry || '❌ 未設定'}\n\n`;
    
    // 有効性の確認
    if (accessToken) {
      const isValid = isJoomTokenValid();
      statusMessage += `【有効性】\n`;
      statusMessage += `トークン状態: ${isValid ? '✅ 有効' : '⚠️ 期限切れまたは無効'}\n\n`;
      
      if (!isValid) {
        statusMessage += '【推奨アクション】\n';
        statusMessage += '• トークンリフレッシュを実行\n';
        statusMessage += '• または再認証を実行\n';
      }
    } else {
      statusMessage += '【推奨アクション】\n';
      statusMessage += '• トークン取得を実行\n';
    }
    
    ui.alert('トークン取得状況', statusMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('トークン取得状況確認エラー:', error);
    ui.alert(
      'エラー',
      'トークン取得状況確認中にエラーが発生しました: ' + error.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * トークンの完全リセット
 */
function resetJoomTokens() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'トークンリセット',
    'Joom APIのトークンを完全にリセットしますか？\n\n' +
    '【削除される情報】\n' +
    '• Access Token\n' +
    '• Refresh Token\n' +
    '• Token Expiry\n' +
    '• Auth Code（一時保存分）\n\n' +
    '【注意】\n' +
    'この操作は取り消せません。\n' +
    '必要に応じて再認証を実行してください。',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      // トークン関連の設定をクリア
      updateSetting('Joom Access Token', '');
      updateSetting('Joom Refresh Token', '');
      updateSetting('Joom Token Expiry', '');
      updateSetting('Joom Auth Code', '');
      
      ui.alert(
        'リセット完了',
        '✅ Joom APIのトークンがリセットされました。\n\n' +
        '必要に応じて「トークン取得」を実行してください。',
        ui.ButtonSet.OK
      );
      
      console.log('Joom APIトークンがリセットされました');
      
    } catch (error) {
      console.error('トークンリセットエラー:', error);
      ui.alert(
        'エラー',
        'トークンリセット中にエラーが発生しました: ' + error.message,
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * デバッグ用：トークン取得の詳細テスト
 */
function debugTokenAcquisition() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const clientId = getSetting('Joom Client ID');
    const clientSecret = getSetting('Joom Client Secret');
    
    console.log('=== デバッグ：トークン取得テスト ===');
    console.log('Client ID:', clientId ? '設定済み' : '未設定');
    console.log('Client Secret:', clientSecret ? '設定済み' : '未設定');
    
    if (!clientId || !clientSecret) {
      ui.alert(
        'デバッグ結果',
        '❌ 基本設定が不完全です\n\n' +
        'Client ID: ' + (clientId ? '✅' : '❌') + '\n' +
        'Client Secret: ' + (clientSecret ? '✅' : '❌') + '\n\n' +
        '設定シートで認証情報を確認してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 認証URLの生成テスト
    const authUrl = generateJoomAuthUrl(clientId);
    console.log('生成された認証URL:', authUrl);
    
    ui.alert(
      'デバッグ結果',
      '✅ 基本設定は正常です\n\n' +
      'Client ID: ✅ 設定済み\n' +
      'Client Secret: ✅ 設定済み\n\n' +
      '【次のステップ】\n' +
      '1. Joom設定でRedirect URLを設定\n' +
      '2. 手動入力方式でトークン取得を試行\n' +
      '3. 認証コードを正確にコピー&ペースト\n' +
      '4. エラーが発生した場合はログを確認\n\n' +
      '認証URL:\n' + authUrl.substring(0, 100) + '...',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('デバッグテストエラー:', error);
    ui.alert(
      'デバッグエラー',
      'デバッグテスト中にエラーが発生しました: ' + error.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 代替のリダイレクトURL設定
 */
function setAlternativeRedirectUri() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'リダイレクトURL設定',
    'Joom OAuth認証用のリダイレクトURLを設定します。\n\n' +
    '以下のオプションから選択してください：\n\n' +
    '1. localhost:8080（推奨）\n' +
    '2. カスタムURL\n' +
    '3. 現在の設定を確認',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (response === ui.Button.YES) {
    // localhost:8080を設定
    updateSetting('Joom Redirect URI', 'http://localhost:8080/callback');
    ui.alert(
      '設定完了',
      '✅ リダイレクトURLを設定しました\n\n' +
      '設定値: http://localhost:8080/callback\n\n' +
      'このURLをJoom設定のRedirect URLに設定してください。',
      ui.ButtonSet.OK
    );
  } else if (response === ui.Button.NO) {
    // カスタムURLの入力
    const customResponse = ui.prompt(
      'カスタムリダイレクトURL',
      '使用するリダイレクトURLを入力してください：\n\n' +
      '例: http://localhost:8080/callback\n' +
      '例: https://example.com/callback\n\n' +
      'URL:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (customResponse.getSelectedButton() === ui.Button.OK) {
      const customUrl = customResponse.getResponseText().trim();
      if (customUrl) {
        updateSetting('Joom Redirect URI', customUrl);
        ui.alert(
          '設定完了',
          '✅ カスタムリダイレクトURLを設定しました\n\n' +
          `設定値: ${customUrl}\n\n` +
          'このURLをJoom設定のRedirect URLに設定してください。',
          ui.ButtonSet.OK
        );
      }
    }
  } else {
    // 現在の設定を確認
    const currentUri = getSetting('Joom Redirect URI');
    ui.alert(
      '現在の設定',
      '現在のリダイレクトURL設定:\n\n' +
      `設定値: ${currentUri || '未設定'}\n\n` +
      'Joom設定で同じURLを設定してください。',
      ui.ButtonSet.OK
    );
  }
}

/**
 * ==========================================
 * 注文取得メニュー機能
 * ==========================================
 */

/**
 * 最新注文の取得メニュー
 */
function fetchLatestJoomOrders() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // トークンの有効性確認
    if (!ensureValidToken()) {
      ui.alert(
        '認証エラー',
        '有効なアクセストークンがありません。\n「JoomAPI設定」→「トークン取得」を実行してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    const response = ui.alert(
      '最新注文の取得',
      '前回連携時間以降の最新注文を取得し、売上管理シートに追加します。\n\n' +
      'この処理には時間がかかる場合があります。\n' +
      '続行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // 注文取得の実行
    const startTime = new Date();
    console.log('最新注文取得開始:', startTime.toISOString());
    
    const orders = fetchJoomOrders({ mode: 'lastSync' });
    
    if (orders && orders.length > 0) {
      // データ変換・挿入
      const result = batchTransformAndInsertOrders(orders);
      
      const endTime = new Date();
      const duration = (endTime.getTime() - startTime.getTime()) / 1000;
      
      let message = `✅ 最新注文の取得・登録が完了しました！\n\n`;
      message += `取得件数: ${orders.length}件\n`;
      message += `登録成功: ${result.inserted}件\n`;
      
      if (result.failed > 0) {
        message += `登録失敗: ${result.failed}件\n`;
      }
      
      message += `処理時間: ${duration.toFixed(1)}秒\n\n`;
      message += '売上管理シートと在庫管理シートを確認してください。';
      
      ui.alert('注文取得完了', message, ui.ButtonSet.OK);
      
      // 前回連携時間を更新
      updateSetting('Joom 前回連携時間', Utilities.formatDate(endTime, 'JST', 'yyyy-MM-dd HH:mm:ss'));
      console.log('前回連携時間を更新:', endTime.toISOString());
      
    } else {
      const endTime = new Date();
      const duration = (endTime.getTime() - startTime.getTime()) / 1000;
      
      ui.alert(
        '注文取得完了',
        `✅ 注文取得が完了しました。\n\n` +
        `取得件数: 0件（新しい注文はありません）\n` +
        `処理時間: ${duration.toFixed(1)}秒\n\n` +
        '前回連携時間以降に新しい注文はありませんでした。',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('最新注文取得エラー:', error);
    ui.alert(
      'エラー',
      '最新注文の取得中にエラーが発生しました: ' + error.message + '\n\n' +
      '詳細はログを確認してください。',
      ui.ButtonSet.OK
    );
  }
}

/**
 * 日時範囲での注文取得メニュー
 */
function fetchJoomOrdersByDateMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // トークンの有効性確認
    if (!ensureValidToken()) {
      ui.alert(
        '認証エラー',
        '有効なアクセストークンがありません。\n「JoomAPI設定」→「トークン取得」を実行してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 開始日時の入力
    const startDateResponse = ui.prompt(
      '開始日時の設定',
      '注文取得の開始日時を入力してください：\n\n' +
      '例: 2025-10-01 00:00:00\n' +
      '例: 2025-10-01（時刻は00:00:00になります）\n\n' +
      '開始日時:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (startDateResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const startDateText = startDateResponse.getResponseText().trim();
    if (!startDateText) {
      ui.alert('エラー', '開始日時が入力されていません。', ui.ButtonSet.OK);
      return;
    }
    
    // 終了日時の入力
    const endDateResponse = ui.prompt(
      '終了日時の設定',
      '注文取得の終了日時を入力してください：\n\n' +
      '例: 2025-10-07 23:59:59\n' +
      '例: 2025-10-07（時刻は23:59:59になります）\n\n' +
      '終了日時:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (endDateResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const endDateText = endDateResponse.getResponseText().trim();
    if (!endDateText) {
      ui.alert('エラー', '終了日時が入力されていません。', ui.ButtonSet.OK);
      return;
    }
    
    // 日時の解析（終了日は日付のみの場合 23:59:59 を補完）
    let startDate, endDate;
    try {
      startDate = parseDateTime(startDateText, false);
      endDate = parseDateTime(endDateText, true);
    } catch (error) {
      ui.alert('エラー', '日時の形式が正しくありません。\n例: 2025-10-01 00:00:00', ui.ButtonSet.OK);
      return;
    }
    
    // 確認
    const confirmResponse = ui.alert(
      '日時範囲の確認',
      `以下の日時範囲で注文を取得します：\n\n` +
      `開始: ${startDate.toLocaleString('ja-JP')}\n` +
      `終了: ${endDate.toLocaleString('ja-JP')}\n\n` +
      'この処理には時間がかかる場合があります。\n' +
      '続行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    // 注文取得の実行
    const startTime = new Date();
    console.log('日時範囲注文取得開始:', startTime.toISOString());
    
    const orders = fetchJoomOrders({
      mode: 'dateRange',
      startDate: startDate,
      endDate: endDate
    });
    
    if (orders && orders.length > 0) {
      // データ変換・挿入
      const result = batchTransformAndInsertOrders(orders);
      
      const endTime = new Date();
      const duration = (endTime.getTime() - startTime.getTime()) / 1000;
      
      let message = `✅ 日時範囲での注文取得・登録が完了しました！\n\n`;
      message += `取得件数: ${orders.length}件\n`;
      message += `登録成功: ${result.inserted}件\n`;
      
      if (result.failed > 0) {
        message += `登録失敗: ${result.failed}件\n`;
      }
      
      message += `処理時間: ${duration.toFixed(1)}秒\n\n`;
      message += '売上管理シートと在庫管理シートを確認してください。';
      
      ui.alert('注文取得完了', message, ui.ButtonSet.OK);
      
    } else {
      const endTime = new Date();
      const duration = (endTime.getTime() - startTime.getTime()) / 1000;
      
      ui.alert(
        '注文取得完了',
        `✅ 注文取得が完了しました。\n\n` +
        `取得件数: 0件\n` +
        `処理時間: ${duration.toFixed(1)}秒\n\n` +
        '指定した日時範囲に注文はありませんでした。',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('日時範囲注文取得エラー:', error);
    ui.alert(
      'エラー',
      '日時範囲での注文取得中にエラーが発生しました: ' + error.message + '\n\n' +
      '詳細はログを確認してください。',
      ui.ButtonSet.OK
    );
  }
}

/**
 * 特定注文の取得メニュー
 */
function fetchSpecificJoomOrder() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // トークンの有効性確認
    if (!ensureValidToken()) {
      ui.alert(
        '認証エラー',
        '有効なアクセストークンがありません。\n「JoomAPI設定」→「トークン取得」を実行してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    const orderIdResponse = ui.prompt(
      '注文IDの入力',
      '取得する注文のIDを入力してください：\n\n' +
      '例: 12345678\n' +
      '例: ORD-2025-001\n\n' +
      '注文ID:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (orderIdResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const orderId = orderIdResponse.getResponseText().trim();
    if (!orderId) {
      ui.alert('エラー', '注文IDが入力されていません。', ui.ButtonSet.OK);
      return;
    }
    
    // 注文取得の実行
    const startTime = new Date();
    console.log('特定注文取得開始:', orderId);
    
    const order = fetchJoomOrders({
      mode: 'single',
      orderId: orderId
    });
    
    if (order) {
      // データ変換・挿入
      const result = batchTransformAndInsertOrders([order]);
      
      const endTime = new Date();
      const duration = (endTime.getTime() - startTime.getTime()) / 1000;
      
      let message = `✅ 特定注文の取得・登録が完了しました！\n\n`;
      message += `注文ID: ${orderId}\n`;
      message += `登録成功: ${result.inserted}件\n`;
      
      if (result.failed > 0) {
        message += `登録失敗: ${result.failed}件\n`;
      }
      
      message += `処理時間: ${duration.toFixed(1)}秒\n\n`;
      message += '売上管理シートと在庫管理シートを確認してください。';
      
      ui.alert('注文取得完了', message, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('特定注文取得エラー:', error);
    ui.alert(
      'エラー',
      '特定注文の取得中にエラーが発生しました: ' + error.message + '\n\n' +
      '注文IDが正しいか確認してください。',
      ui.ButtonSet.OK
    );
  }
}

/**
 * 注文取得状況の確認
 */
function checkOrderFetchStatus() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const lastSyncTime = getSetting('Joom 前回連携時間');
    const existingOrderIds = getExistingOrderIds();
    
    let statusMessage = '【Joom注文取得状況】\n\n';
    
    statusMessage += '【前回連携時間】\n';
    statusMessage += `最終同期: ${lastSyncTime || '未設定'}\n\n`;
    
    statusMessage += '【売上管理シート】\n';
    statusMessage += `登録済み注文数: ${existingOrderIds.size}件\n\n`;
    
    statusMessage += '【利用可能な機能】\n';
    statusMessage += '• 最新注文の取得（前回連携時間以降）\n';
    statusMessage += '• 日時範囲での注文取得\n';
    statusMessage += '• 特定注文の取得\n\n';
    
    if (!lastSyncTime) {
      statusMessage += '【推奨アクション】\n';
      statusMessage += '• 「最新注文を取得」を実行して初回同期を行ってください\n';
    }
    
    ui.alert('注文取得状況', statusMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('注文取得状況確認エラー:', error);
    ui.alert(
      'エラー',
      '注文取得状況確認中にエラーが発生しました: ' + error.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * ==========================================
 * デバッグ機能
 * ==========================================
 */

/**
 * デバッグ: 最新注文の生データを表示
 */
function debugShowLatestOrdersRawData() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // トークンの有効性確認
    if (!ensureValidToken()) {
      ui.alert(
        '認証エラー',
        '有効なアクセストークンがありません。\n「JoomAPI設定」→「トークン取得」を実行してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    const response = ui.alert(
      'デバッグ: 最新注文データ表示',
      '前回連携時間以降の注文データ（生データ）を取得して表示します。\n\n' +
      '最大3件まで表示します。\n' +
      '続行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // 注文取得の実行
    console.log('デバッグ: 最新注文生データ取得開始');
    const orders = fetchJoomOrders({ mode: 'lastSync' });
    
    if (!orders || orders.length === 0) {
      ui.alert(
        'デバッグ結果',
        '取得された注文データはありません。\n\n' +
        '前回連携時間以降に新しい注文がない可能性があります。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 最大3件まで表示
    const displayCount = Math.min(orders.length, 3);
    let debugMessage = `【デバッグ: 注文データ（生データ）】\n\n`;
    debugMessage += `取得件数: ${orders.length}件\n`;
    debugMessage += `表示件数: ${displayCount}件\n\n`;
    debugMessage += `${'='.repeat(40)}\n\n`;
    
    for (let i = 0; i < displayCount; i++) {
      const order = orders[i];
      debugMessage += `【注文 ${i + 1}】\n`;
      debugMessage += formatOrderDataForDisplay(order);
      debugMessage += `\n${'='.repeat(40)}\n\n`;
    }
    
    if (orders.length > displayCount) {
      debugMessage += `※ ${orders.length - displayCount}件の注文データは省略されています\n`;
      debugMessage += `詳細はログを確認してください（実行 → ログを表示）\n`;
    }
    
    // ログに全データを出力
    console.log('取得した全注文データ:', JSON.stringify(orders, null, 2));
    
    ui.alert('デバッグ結果', debugMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    ui.alert(
      'エラー',
      'デバッグ実行中にエラーが発生しました: ' + error.message + '\n\n' +
      '詳細はログを確認してください。',
      ui.ButtonSet.OK
    );
  }
}

/**
 * デバッグ: 特定注文の生データを表示
 */
function debugShowSpecificOrderRawData() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // トークンの有効性確認
    if (!ensureValidToken()) {
      ui.alert(
        '認証エラー',
        '有効なアクセストークンがありません。\n「JoomAPI設定」→「トークン取得」を実行してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    const orderIdResponse = ui.prompt(
      'デバッグ: 注文IDの入力',
      '生データを表示する注文のIDを入力してください：\n\n' +
      '例: 12345678\n\n' +
      '注文ID:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (orderIdResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const orderId = orderIdResponse.getResponseText().trim();
    if (!orderId) {
      ui.alert('エラー', '注文IDが入力されていません。', ui.ButtonSet.OK);
      return;
    }
    
    // 注文取得の実行
    console.log('デバッグ: 特定注文生データ取得開始:', orderId);
    const order = fetchJoomOrders({
      mode: 'single',
      orderId: orderId
    });
    
    let debugMessage = `【デバッグ: 注文データ（生データ）】\n\n`;
    debugMessage += `注文ID: ${orderId}\n\n`;
    debugMessage += `${'='.repeat(40)}\n\n`;
    debugMessage += formatOrderDataForDisplay(order);
    debugMessage += `\n${'='.repeat(40)}\n\n`;
    debugMessage += `※ 完全なJSONデータはログで確認できます\n`;
    debugMessage += `（実行 → ログを表示）`;
    
    // ログに全データを出力
    console.log('取得した注文データ（完全版）:', JSON.stringify(order, null, 2));
    
    ui.alert('デバッグ結果', debugMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    ui.alert(
      'エラー',
      'デバッグ実行中にエラーが発生しました: ' + error.message + '\n\n' +
      '注文IDが正しいか確認してください。',
      ui.ButtonSet.OK
    );
  }
}

/**
 * デバッグ: 日時範囲の注文生データを表示
 */
function debugShowDateRangeOrdersRawData() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // トークンの有効性確認
    if (!ensureValidToken()) {
      ui.alert(
        '認証エラー',
        '有効なアクセストークンがありません。\n「JoomAPI設定」→「トークン取得」を実行してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 開始日時の入力
    const startDateResponse = ui.prompt(
      'デバッグ: 開始日時の設定',
      '注文取得の開始日時を入力してください：\n\n' +
      '例: 2025-10-13\n' +
      '例: 2025-10-13 00:00:00\n\n' +
      '開始日時:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (startDateResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const startDateText = startDateResponse.getResponseText().trim();
    if (!startDateText) {
      ui.alert('エラー', '開始日時が入力されていません。', ui.ButtonSet.OK);
      return;
    }
    
    // 終了日時の入力
    const endDateResponse = ui.prompt(
      'デバッグ: 終了日時の設定',
      '注文取得の終了日時を入力してください：\n\n' +
      '例: 2025-10-13\n' +
      '例: 2025-10-13 23:59:59\n\n' +
      '終了日時:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (endDateResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const endDateText = endDateResponse.getResponseText().trim();
    if (!endDateText) {
      ui.alert('エラー', '終了日時が入力されていません。', ui.ButtonSet.OK);
      return;
    }
    
    // 日時の解析（終了日は日付のみの場合 23:59:59 を補完）
    let startDate, endDate;
    try {
      startDate = parseDateTime(startDateText, false);
      endDate = parseDateTime(endDateText, true);
    } catch (error) {
      ui.alert('エラー', '日時の形式が正しくありません。\n例: 2025-10-13 00:00:00', ui.ButtonSet.OK);
      return;
    }
    
    // 注文取得の実行
    console.log('デバッグ: 日時範囲注文生データ取得開始');
    const orders = fetchJoomOrders({
      mode: 'dateRange',
      startDate: startDate,
      endDate: endDate
    });
    
    if (!orders || orders.length === 0) {
      ui.alert(
        'デバッグ結果',
        '取得された注文データはありません。\n\n' +
        '指定した日時範囲に注文がない可能性があります。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 最大3件まで表示
    const displayCount = Math.min(orders.length, 3);
    let debugMessage = `【デバッグ: 注文データ（生データ）】\n\n`;
    debugMessage += `期間: ${startDate.toLocaleString('ja-JP')} - ${endDate.toLocaleString('ja-JP')}\n`;
    debugMessage += `取得件数: ${orders.length}件\n`;
    debugMessage += `表示件数: ${displayCount}件\n\n`;
    debugMessage += `${'='.repeat(40)}\n\n`;
    
    for (let i = 0; i < displayCount; i++) {
      const order = orders[i];
      debugMessage += `【注文 ${i + 1}】\n`;
      debugMessage += formatOrderDataForDisplay(order);
      debugMessage += `\n${'='.repeat(40)}\n\n`;
    }
    
    if (orders.length > displayCount) {
      debugMessage += `※ ${orders.length - displayCount}件の注文データは省略されています\n`;
      debugMessage += `詳細はログを確認してください（実行 → ログを表示）\n`;
    }
    
    // ログに全データを出力
    console.log('取得した全注文データ:', JSON.stringify(orders, null, 2));
    
    ui.alert('デバッグ結果', debugMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    ui.alert(
      'エラー',
      'デバッグ実行中にエラーが発生しました: ' + error.message + '\n\n' +
      '詳細はログを確認してください。',
      ui.ButtonSet.OK
    );
  }
}

/**
 * 注文データを表示用にフォーマット
 * @param {Object} order - 注文データ
 * @returns {string} フォーマットされた文字列
 */
function formatOrderDataForDisplay(order) {
  try {
    let formatted = '';
    
    // 基本情報
    formatted += `【基本情報】\n`;
    formatted += `注文ID: ${order.order_id || order.id || 'N/A'}\n`;
    formatted += `注文日時: ${order.created || order.order_date || 'N/A'}\n`;
    formatted += `更新日時: ${order.updated || order.updated_at || 'N/A'}\n`;
    formatted += `ステータス: ${order.status || 'N/A'}\n`;
    formatted += `Store ID: ${order.store_id || 'N/A'}\n\n`;
    
    // 価格情報
    formatted += `【価格情報】\n`;
    formatted += `商品小計: ${order.items_price || 'N/A'}\n`;
    formatted += `配送料: ${order.shipping_price || 'N/A'}\n`;
    formatted += `合計金額: ${order.total_price || 'N/A'}\n`;
    formatted += `手数料: ${order.fee_price || order.commission || 'N/A'}\n`;
    formatted += `VAT: ${order.vat_price || 'N/A'}\n\n`;
    
    // 商品情報
    if (order.items && Array.isArray(order.items) && order.items.length > 0) {
      formatted += `【商品情報】(${order.items.length}件)\n`;
      order.items.slice(0, 2).forEach((item, index) => {
        formatted += `[商品${index + 1}]\n`;
        formatted += `  SKU: ${item.sku || item.product_sku || 'N/A'}\n`;
        formatted += `  商品名: ${item.name || item.product_name || 'N/A'}\n`;
        formatted += `  数量: ${item.quantity || 'N/A'}\n`;
        formatted += `  価格: ${item.price || 'N/A'}\n`;
      });
      if (order.items.length > 2) {
        formatted += `  ... 他${order.items.length - 2}件\n`;
      }
      formatted += `\n`;
    }
    
    // 配送情報
    if (order.shipping || order.shipping_address) {
      const shipping = order.shipping || order.shipping_address;
      formatted += `【配送情報】\n`;
      formatted += `配送先名: ${shipping.name || shipping.recipient_name || 'N/A'}\n`;
      formatted += `国: ${shipping.country || shipping.country_code || 'N/A'}\n`;
      formatted += `都市: ${shipping.city || 'N/A'}\n`;
      formatted += `郵便番号: ${shipping.zip || shipping.postal_code || 'N/A'}\n\n`;
    }
    
    // 追跡情報
    if (order.tracking_code || order.tracking_number) {
      formatted += `【追跡情報】\n`;
      formatted += `追跡番号: ${order.tracking_code || order.tracking_number}\n\n`;
    }
    
    // その他の重要フィールド
    formatted += `【その他】\n`;
    formatted += `支払方法: ${order.payment_method || 'N/A'}\n`;
    formatted += `配送方法: ${order.shipping_method || 'N/A'}\n`;
    formatted += `通貨: ${order.currency || 'N/A'}\n`;
    
    return formatted;
    
  } catch (error) {
    console.error('データフォーマットエラー:', error);
    return `データフォーマット中にエラーが発生しました。\n生データをログで確認してください。`;
  }
}

/**
 * 日時文字列の解析
 * @param {string} dateTimeString - 日時文字列
 * @param {boolean} isEndOfDay - 終了日用のとき true の場合、日付のみ入力で 23:59:59 を補完
 * @returns {Date} 解析された日時
 */
function parseDateTime(dateTimeString, isEndOfDay) {
  try {
    // 日付のみの場合（時刻を補完）
    if (!dateTimeString.includes(' ')) {
      if (dateTimeString.includes(':')) {
        // 時刻のみの場合は今日の日付を補完
        const today = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
        dateTimeString = `${today} ${dateTimeString}`;
      } else {
        // 日付のみ: 終了日なら23:59:59、開始日なら00:00:00
        dateTimeString = isEndOfDay
          ? `${dateTimeString} 23:59:59`
          : `${dateTimeString} 00:00:00`;
      }
    }
    
    // JST形式で解析
    const date = new Date(dateTimeString);
    if (isNaN(date.getTime())) {
      throw new Error('Invalid date');
    }
    
    return date;
    
  } catch (error) {
    throw new Error('日時の解析に失敗しました: ' + dateTimeString);
  }
}

/**
 * ==========================================
 * 5. カスタムメニュー拡張
 * ==========================================
 */

/**
 * Joom API認証メニューの表示
 */
function showJoomAuthMenu() {
  const ui = SpreadsheetApp.getUi();
  
  // WebアプリのURLを取得
  const webAppUrl = getWebAppUrl();
  
  // 設定シートのRedirect URIを更新
  updateSetting('Joom Redirect URI', webAppUrl);
  
  const response = ui.alert(
    'Joom API 認証',
    'Joom APIの認証を開始しますか？\n\n' +
    '※ 事前に設定シートでClient IDとClient Secretを設定してください。\n' +
    '※ この認証ではWebアプリを使用します。',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    authenticateWithJoomWebApp();
  }
}

/**
 * Webアプリ方式でのJoom API認証
 */
function authenticateWithJoomWebApp() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Client IDの確認
    const clientId = getSetting('Joom Client ID');
    if (!clientId) {
      ui.alert(
        '認証エラー',
        'Joom Client IDが設定されていません。\n設定シートでClient IDを設定してください。',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // WebアプリのURLを取得
    const webAppUrl = getWebAppUrl();
    updateSetting('Joom Redirect URI', webAppUrl);
    
    // 認証手順の説明ダイアログ
    const instructionResponse = ui.alert(
      'Joom API 認証（Webアプリ版）',
      '以下の手順でJoom APIの認証を行います：\n\n' +
      '【重要】事前にJoom設定でRedirect URLを設定してください\n' +
      `Redirect URL: ${webAppUrl}\n\n` +
      '1. 次のダイアログに表示されるURLにアクセス\n' +
      '2. Joomアカウントでログイン\n' +
      '3. アクセス許可をクリック\n' +
      '4. 自動的に認証が完了します\n\n' +
      '準備ができたら「OK」を押してください。',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (instructionResponse !== ui.Button.OK) {
      console.log('認証がキャンセルされました');
      return false;
    }
    
    // 認証URLの生成
    const authUrl = generateJoomAuthUrl(clientId);
    
    // 認証URLを表示
    const urlResponse = ui.alert(
      '認証URL',
      '以下のURLにアクセスして認証を完了してください：\n\n' +
      authUrl + '\n\n' +
      'URLをコピーしてブラウザで開いてください。\n' +
      '認証完了後、自動的にスプレッドシートに戻ります。\n\n' +
      '【エラーが出る場合】\n' +
      'Joom設定のRedirect URLが以下と一致しているか確認してください：\n' +
      webAppUrl,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (urlResponse !== ui.Button.OK) {
      console.log('認証がキャンセルされました');
      return false;
    }
    
    return true;
    
  } catch (error) {
    console.error('認証エラー:', error);
    ui.alert(
      '認証エラー',
      '認証中にエラーが発生しました: ' + error.message,
      ui.ButtonSet.OK
    );
    return false;
  }
}

/**
 * トークンステータス確認メニュー
 */
function showJoomTokenStatus() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const accessToken = getSetting('Joom Access Token');
    const expiryStr = getSetting('Joom Token Expiry');
    
    if (!accessToken) {
      ui.alert(
        'トークンステータス',
        'アクセストークンが設定されていません。\n認証を実行してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    const isValid = isJoomTokenValid();
    const validityText = isValid ? '有効' : '期限切れまたは無効';
    
    const message = `
【Joom APIトークンステータス】

アクセストークン: ${accessToken.substring(0, 20)}...
有効期限: ${expiryStr || '不明'}
ステータス: ${validityText}

${!isValid ? '\n※ トークンをリフレッシュする必要があります。' : ''}
    `;
    
    ui.alert('トークンステータス', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('トークンステータス確認エラー:', error);
    ui.alert(
      'エラー',
      'トークンステータス確認中にエラーが発生しました: ' + error.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * トークン手動リフレッシュメニュー
 */
function showJoomTokenRefreshMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'トークンリフレッシュ',
    'アクセストークンをリフレッシュしますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    if (refreshJoomToken()) {
      ui.alert(
        'リフレッシュ成功',
        'アクセストークンが更新されました。',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'リフレッシュ失敗',
        'アクセストークンの更新に失敗しました。\n再認証が必要な場合があります。',
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * ==========================================
 * 5. OAuth認証コールバック処理
 * ==========================================
 */

/**
 * OAuth認証コールバック処理（GAS Webアプリ用）
 * @param {Object} e - リクエストパラメータ
 * @returns {HtmlOutput} 認証結果ページ
 */
function doGet(e) {
  try {
    const authCode = e.parameter.code;
    const error = e.parameter.error;
    
    if (error) {
      return HtmlService.createHtmlOutput(`
        <html>
          <head>
            <title>Joom API認証エラー</title>
            <style>
              body { font-family: Arial, sans-serif; padding: 20px; }
              .error { color: red; }
              .success { color: green; }
            </style>
          </head>
          <body>
            <h1>Joom API認証エラー</h1>
            <p class="error">エラー: ${error}</p>
            <p>ブラウザを閉じて、スプレッドシートに戻ってください。</p>
          </body>
        </html>
      `);
    }
    
    if (authCode) {
      // 認証コードを設定シートに一時保存
      updateSetting('Joom Auth Code', authCode);
      
      // 自動的にトークン取得を実行
      const success = exchangeCodeForToken(authCode);
      
      if (success) {
        return HtmlService.createHtmlOutput(`
          <html>
            <head>
              <title>Joom API認証完了</title>
              <style>
                body { font-family: Arial, sans-serif; padding: 20px; }
                .error { color: red; }
                .success { color: green; }
              </style>
            </head>
            <body>
              <h1>Joom API認証完了</h1>
              <p class="success">✅ 認証が正常に完了しました！</p>
              <p>アクセストークンが取得され、設定シートに保存されました。</p>
              <p>ブラウザを閉じて、スプレッドシートに戻ってください。</p>
            </body>
          </html>
        `);
      } else {
        return HtmlService.createHtmlOutput(`
          <html>
            <head>
              <title>Joom API認証失敗</title>
              <style>
                body { font-family: Arial, sans-serif; padding: 20px; }
                .error { color: red; }
                .success { color: green; }
              </style>
            </head>
            <body>
              <h1>Joom API認証失敗</h1>
              <p class="error">❌ トークンの取得に失敗しました</p>
              <p>ブラウザを閉じて、スプレッドシートに戻って再試行してください。</p>
            </body>
          </html>
        `);
      }
    }
    
    // 認証コードが取得できない場合
    return HtmlService.createHtmlOutput(`
      <html>
        <head>
          <title>Joom API認証</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            .error { color: red; }
            .success { color: green; }
          </style>
        </head>
        <body>
          <h1>Joom API認証</h1>
          <p>認証コードが取得できませんでした。</p>
          <p>ブラウザを閉じて、スプレッドシートに戻ってください。</p>
        </body>
      </html>
    `);
    
  } catch (error) {
    console.error('OAuth認証コールバックエラー:', error);
    return HtmlService.createHtmlOutput(`
      <html>
        <head>
          <title>Joom API認証エラー</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            .error { color: red; }
          </style>
        </head>
        <body>
          <h1>Joom API認証エラー</h1>
          <p class="error">エラーが発生しました: ${error.message}</p>
          <p>ブラウザを閉じて、スプレッドシートに戻ってください。</p>
        </body>
      </html>
    `);
  }
}

/**
 * GAS WebアプリのURL取得
 * @returns {string} WebアプリのURL
 */
function getWebAppUrl() {
  const scriptId = ScriptApp.getScriptId();
  return `https://script.google.com/macros/s/${scriptId}/usercallback`;
}

/**
 * 現在のWebアプリURLを表示（デバッグ用）
 */
function showCurrentWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  const webAppUrl = getWebAppUrl();
  
  ui.alert(
    '現在のWebアプリURL',
    `現在のGAS WebアプリURL:\n\n${webAppUrl}\n\n` +
    'このURLをJoom設定のRedirect URLに設定してください。',
    ui.ButtonSet.OK
  );
  
  console.log('WebアプリURL:', webAppUrl);
}

/**
 * GAS Webアプリの設定確認
 */
function checkWebAppConfiguration() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // スクリプトIDの確認
    const scriptId = ScriptApp.getScriptId();
    const webAppUrl = getWebAppUrl();
    
    // doGet関数の存在確認
    const hasDoGet = typeof doGet === 'function';
    
    const message = `
【GAS Webアプリ設定確認】

スクリプトID: ${scriptId}
WebアプリURL: ${webAppUrl}
doGet関数: ${hasDoGet ? '✅ 存在' : '❌ 未定義'}

【確認事項】
1. GASエディタで「デプロイ」→「新しいデプロイ」を実行
2. 種類: ウェブアプリ
3. 実行者: 自分
4. アクセス: 自分
5. デプロイ後にURLを確認

【Joom設定】
Redirect URL: ${webAppUrl}
    `;
    
    ui.alert('Webアプリ設定確認', message, ui.ButtonSet.OK);
    
    if (!hasDoGet) {
      ui.alert(
        '警告',
        'doGet関数が定義されていません。\n' +
        'OAuth認証コールバックが処理できません。',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('設定確認エラー:', error);
    ui.alert(
      'エラー',
      '設定確認中にエラーが発生しました: ' + error.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * ==========================================
 * 6. 注文データ取得機能
 * ==========================================
 */

/**
 * 単一注文の取得
 * @param {string} orderId - 注文ID
 * @returns {Object} 注文データ
 */
function fetchSingleJoomOrder(orderId) {
  try {
    console.log(`単一注文取得開始: ${orderId}`);
    // Joom API v3: 単一注文は GET /orders?id=注文ID 形式（パスではなくクエリパラメータ）
    // 参照: doc/Joom_API_v3_Get_Order_詳細資料.md
    const endpoint = `/orders?id=${encodeURIComponent(orderId)}`;
    const response = makeJoomApiRequest(endpoint, 'GET');
    
    if (response && response.data) {
      console.log(`単一注文取得成功: ${orderId}`);
      return response.data;
    } else {
      throw new Error(`注文取得失敗: ${response?.message || response?.errors?.[0]?.message || '不明なエラー'}`);
    }
    
  } catch (error) {
    console.error(`単一注文取得エラー (${orderId}):`, error);
    throw error;
  }
}

/**
 * 複数注文の取得（ページネーション対応）
 * @param {Object} options - 取得オプション
 * @param {number} options.limit - 取得件数（デフォルト: 100）
 * @param {string} options.after - 次ページのカーソル
 * @param {string} options.updatedFrom - 更新日時の開始
 * @param {string} options.updatedTo - 更新日時の終了
 * @param {string} options.status - 注文ステータス
 * @param {string} options.storeId - ストアID
 * @returns {Object} 注文データとページネーション情報
 */
function fetchMultipleJoomOrders(options = {}) {
  try {
    console.log('複数注文取得開始:', options);
    
    // デフォルト値の設定
    const params = {
      limit: options.limit || getSetting('Joom 最大取得件数') || 100,
      ...options
    };
    
    // パラメータの構築
    const queryParams = [];
    Object.keys(params).forEach(key => {
      if (params[key] !== undefined && params[key] !== null && params[key] !== '') {
        queryParams.push(`${key}=${encodeURIComponent(params[key])}`);
      }
    });
    
    const endpoint = `/orders/multi${queryParams.length > 0 ? '?' + queryParams.join('&') : ''}`;
    const response = makeJoomApiRequest(endpoint, 'GET');
    
    // Joom API v3: 実レスポンスは { code, data: { items: [...] } } の形式。資料の data 配列や data.orders にも対応。
    let orders = [];
    if (Array.isArray(response.data)) {
      orders = response.data;
    } else if (response.data && Array.isArray(response.data.items)) {
      orders = response.data.items;
    } else if (response.data && Array.isArray(response.data.orders)) {
      orders = response.data.orders;
    } else if (Array.isArray(response.orders)) {
      orders = response.orders;
    } else if (response.data && typeof response.data === 'object' && response.data !== null && Array.isArray(response.data.data)) {
      orders = response.data.data;
    }
    
    // paging: ルートの paging または data.paging を参照
    const paging = response.paging || (response.data && response.data.paging) || {};
    const nextUrl = paging.next || null;
    let after = null;
    if (nextUrl) {
      try {
        const u = new URL(nextUrl);
        after = u.searchParams.get('after');
      } catch (e) { /* ignore */ }
    }
    
    console.log(`複数注文取得成功: ${orders.length}件`);
    return {
      orders: orders,
      pagination: {
        has_more: !!nextUrl,
        after: after
      },
      totalCount: (response.data && response.data.total_count !== undefined) ? response.data.total_count : orders.length
    };
    
  } catch (error) {
    console.error('複数注文取得エラー:', error);
    throw error;
  }
}

/**
 * 日時範囲指定での注文取得
 * @param {Date} startDate - 開始日時
 * @param {Date} endDate - 終了日時
 * @param {Object} additionalOptions - 追加オプション
 * @returns {Array} 注文データの配列
 */
function fetchJoomOrdersByDateRange(startDate, endDate, additionalOptions = {}) {
  try {
    console.log(`日時範囲指定注文取得開始: ${startDate.toISOString()} - ${endDate.toISOString()}`);
    
    // 日時をRFC3339形式に変換
    const updatedFrom = formatDateForApi(startDate);
    const updatedTo = formatDateForApi(endDate);
    
    const options = {
      updatedFrom,
      updatedTo,
      ...additionalOptions
    };
    
    let allOrders = [];
    let hasMore = true;
    let cursor = null;
    
    while (hasMore) {
      const requestOptions = {
        ...options,
        limit: 100 // レート制限対応のため小さいバッチで取得
      };
      
      if (cursor) {
        requestOptions.after = cursor;
      }
      
      const result = fetchMultipleJoomOrders(requestOptions);
      allOrders.push(...result.orders);
      
      // ページネーションの確認
      hasMore = result.pagination.has_more || false;
      cursor = result.pagination.after || null;
      
      // レート制限対応（50 rpm）
      if (hasMore) {
        Utilities.sleep(1200); // 1.2秒待機（50 rpm対応）
      }
      
      console.log(`取得済み: ${allOrders.length}件, 次ページ: ${hasMore}`);
    }
    
    // 0件の場合: APIのupdatedFrom/updatedToは「最終更新日(updateTimestamp)」で絞る。
    // Approval Date（承認日）は orderTimestamp に近いため、更新日で0件なら
    // updated を付けずに取得し、orderTimestamp で絞り込むフォールバックを試行。
    if (allOrders.length === 0) {
      console.log('更新日時(updatedFrom/updatedTo)で0件のため、注文日(orderTimestamp)でフォールバック取得を試行します');
      allOrders = [];
      let fbHasMore = true;
      let fbCursor = null;
      let fbPageCount = 0;
      const maxFallbackPages = 5; // 最大500件を注文日で確認
      while (fbHasMore && fbPageCount < maxFallbackPages) {
        const fbOpts = { limit: 100 };
        if (fbCursor) fbOpts.after = fbCursor;
        const fbResult = fetchMultipleJoomOrders(fbOpts);
        for (const order of fbResult.orders) {
          const ts = order.orderTimestamp || order.updateTimestamp;
          if (ts) {
            const d = new Date(ts);
            if (!isNaN(d.getTime()) && d >= startDate && d <= endDate) {
              allOrders.push(order);
            }
          }
        }
        fbHasMore = fbResult.pagination.has_more || false;
        fbCursor = fbResult.pagination.after || null;
        fbPageCount++;
        if (fbHasMore) Utilities.sleep(1200);
      }
      console.log(`フォールバック取得完了: 注文日(orderTimestamp)で絞り込み後 ${allOrders.length}件`);
    }
    
    console.log(`日時範囲指定注文取得完了: ${allOrders.length}件`);
    return allOrders;
    
  } catch (error) {
    console.error('日時範囲指定注文取得エラー:', error);
    throw error;
  }
}

/**
 * 前回連携時間以降の注文取得
 * @param {Object} additionalOptions - 追加オプション
 * @returns {Array} 注文データの配列
 */
function fetchJoomOrdersSinceLastSync(additionalOptions = {}) {
  try {
    const lastSyncTime = getSetting('Joom 前回連携時間');
    
    if (!lastSyncTime) {
      console.log('前回連携時間が設定されていません。過去24時間の注文を取得します。');
      const yesterday = new Date();
      yesterday.setHours(yesterday.getHours() - 24);
      return fetchJoomOrdersByDateRange(yesterday, new Date(), additionalOptions);
    }
    
    const lastSyncDate = new Date(lastSyncTime);
    console.log(`前回連携時間以降の注文取得: ${lastSyncDate.toISOString()}`);
    
    return fetchJoomOrdersByDateRange(lastSyncDate, new Date(), additionalOptions);
    
  } catch (error) {
    console.error('前回連携時間以降の注文取得エラー:', error);
    throw error;
  }
}

/**
 * 重複チェック機能
 * @param {Array} newOrders - 新しく取得した注文データ
 * @returns {Array} 重複を除いた注文データ
 */
function removeDuplicateOrders(newOrders) {
  try {
    console.log(`重複チェック開始: ${newOrders.length}件`);
    
    // 売上管理シートから既存の注文IDを取得
    const existingOrderIds = getExistingOrderIds();
    console.log(`既存注文ID数: ${existingOrderIds.size}`);
    
    // 重複を除いた注文をフィルタリング
    const uniqueOrders = newOrders.filter(order => {
      const orderId = order.order_id || order.id;
      const isDuplicate = existingOrderIds.has(orderId);
      
      if (isDuplicate) {
        console.log(`重複注文をスキップ: ${orderId}`);
      }
      
      return !isDuplicate;
    });
    
    console.log(`重複チェック完了: ${uniqueOrders.length}件（重複除外: ${newOrders.length - uniqueOrders.length}件）`);
    return uniqueOrders;
    
  } catch (error) {
    console.error('重複チェックエラー:', error);
    throw error;
  }
}

/**
 * 売上管理シートから既存の注文IDを取得
 * @returns {Set} 既存の注文IDのセット
 */
function getExistingOrderIds() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = spreadsheet.getSheetByName('売上管理');
    
    if (!salesSheet) {
      console.log('売上管理シートが存在しません');
      return new Set();
    }
    
    const lastRow = salesSheet.getLastRow();
    if (lastRow <= 1) {
      console.log('売上管理シートにデータがありません');
      return new Set();
    }
    
    // 注文ID列（B列）のデータを取得
    const orderIdRange = salesSheet.getRange(2, 2, lastRow - 1, 1);
    const orderIdValues = orderIdRange.getValues();
    
    const existingIds = new Set();
    orderIdValues.forEach(row => {
      const orderId = row[0];
      if (orderId && orderId.toString().trim() !== '') {
        existingIds.add(orderId.toString().trim());
      }
    });
    
    return existingIds;
    
  } catch (error) {
    console.error('既存注文ID取得エラー:', error);
    return new Set();
  }
}

/**
 * 注文データの取得（統合関数）
 * @param {Object} options - 取得オプション
 * @param {string} options.mode - 取得モード ('lastSync', 'dateRange', 'all', 'single')
 * @param {Date} options.startDate - 開始日時（dateRangeモード時）
 * @param {Date} options.endDate - 終了日時（dateRangeモード時）
 * @param {string} options.orderId - 注文ID（singleモード時）
 * @param {Object} options.additionalOptions - 追加オプション
 * @returns {Array|Object} 注文データ
 */
function fetchJoomOrders(options = {}) {
  try {
    const mode = options.mode || 'lastSync';
    console.log(`注文データ取得開始 (モード: ${mode})`);
    
    switch (mode) {
      case 'single':
        if (!options.orderId) {
          throw new Error('単一注文取得モードでは注文IDが必要です');
        }
        return fetchSingleJoomOrder(options.orderId);
        
      case 'dateRange':
        if (!options.startDate || !options.endDate) {
          throw new Error('日時範囲取得モードでは開始日時と終了日時が必要です');
        }
        const orders = fetchJoomOrdersByDateRange(options.startDate, options.endDate, options.additionalOptions);
        return removeDuplicateOrders(orders);
        
      case 'all':
        // 過去30日間の全注文を取得
        const thirtyDaysAgo = new Date();
        thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
        const allOrders = fetchJoomOrdersByDateRange(thirtyDaysAgo, new Date(), options.additionalOptions);
        return removeDuplicateOrders(allOrders);
        
      case 'lastSync':
      default:
        const lastSyncOrders = fetchJoomOrdersSinceLastSync(options.additionalOptions);
        return removeDuplicateOrders(lastSyncOrders);
    }
    
  } catch (error) {
    console.error('注文データ取得エラー:', error);
    throw error;
  }
}

/**
 * ==========================================
 * 7. ユーティリティ関数
 * ==========================================
 */

/**
 * 日時をRFC3339形式に変換
 * @param {Date} date - 変換する日時
 * @returns {string} RFC3339形式の文字列
 */
function formatDateForApi(date) {
  return Utilities.formatDate(date, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

/**
 * RFC3339形式の日時をJST日付に変換
 * @param {string} rfc3339String - RFC3339形式の文字列
 * @returns {string} JST形式の日付文字列（YYYY-MM-DD）
 */
function convertRfc3339ToJstDate(rfc3339String) {
  try {
    const utcDate = new Date(rfc3339String);
    return Utilities.formatDate(utcDate, 'JST', 'yyyy-MM-dd');
  } catch (error) {
    console.error('日時変換エラー:', error);
    return '';
  }
}

/**
 * RFC3339形式の日時をJST日時に変換
 * @param {string} rfc3339String - RFC3339形式の文字列
 * @returns {string} JST形式の日時文字列（YYYY-MM-DD HH:mm:ss）
 */
function convertRfc3339ToJstDateTime(rfc3339String) {
  try {
    const utcDate = new Date(rfc3339String);
    return Utilities.formatDate(utcDate, 'JST', 'yyyy-MM-dd HH:mm:ss');
  } catch (error) {
    console.error('日時変換エラー:', error);
    return '';
  }
}

