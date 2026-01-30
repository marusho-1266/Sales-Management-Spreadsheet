/**
 * 利益計算（事前利益計算特化）イベントコーディネータ
 * 仕様参照: doc/Joom利益計算機能仕様書.md
 */

// 変更ソースの識別子
const CHANGE_SOURCES = {
  USER_INPUT: 'user_input',
  PROGRAMMATIC: 'programmatic',
  FORMULA: 'formula'
};

// 計算モジュール識別子
const CALCULATION_MODULES = {
  SHIPPING: 'shipping',
  COST: 'cost',
  PROFIT: 'profit'
};

// デバウンス設定
const DEBOUNCE_CONFIG = {
  WINDOW_MS: 300,
  MAX_BATCH_SIZE: 50,
  LOCK_TIMEOUT_MS: 30000
};

// ループ防止用フラグ（ScriptProperties）
const PROGRAMMATIC_FLAG_KEY = 'PC_IGNORE_ONEDIT';

// 内部バッファ（簡易実装）
let __editEventQueue = [];
let __debounceTimer = null;

/**
 * onEdit: 編集イベントのエントリーポイント
 * - イベント収集・バッチ化
 * - 影響モジュールの解析
 * - 実行順序に従って計算をスケジュール
 */
function onEdit(e) {
  try {
    if (isProgrammaticUpdateActive()) {
      // プログラム更新中はonEditを無視
      return;
    }
    if (!e) return;
    const range = e.range;
    if (!range) return;

    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    if (sheetName !== SHEET_NAMES.PROFIT) {
      // 利益計算シート以外の編集は対象外
      return;
    }

    // B2セル（商品ID）の編集を検出して自動データ読み込み
    if (range.getA1Notation() === 'B2') {
      const productId = range.getValue();
      if (productId && productId !== '') {
        // 商品データの自動読み込み（onEdit時点でセル値は確定しているため即時実行）
        loadProductDataFromInventory(productId);
      }
    }

    // イベントをキューに追加
    __editEventQueue.push({
      a1: range.getA1Notation(),
      row: range.getRow(),
      column: range.getColumn(),
      value: range.getValue(),
      timestamp: Date.now()
    });

    // デバウンススケジュール
    scheduleExecution();
  } catch (error) {
    console.error('onEditエラー:', error);
  }
}

/**
 * デバウンスによる実行スケジュール
 */
function scheduleExecution() {
  if (__debounceTimer) {
    clearTimeout(__debounceTimer);
  }
  __debounceTimer = Utilities.sleep ? (function() {
    // GASにはsetTimeoutがないため、疑似的に待機は行わず即時実行
    executeQueuedEdits();
  })() : null;
}

/**
 * キューされた編集イベントを処理し、影響モジュールを決定して順次実行
 */
function executeQueuedEdits() {
  // スナップショットを取得してキューをリセット
  const events = __editEventQueue.splice(0, DEBOUNCE_CONFIG.MAX_BATCH_SIZE);
  if (events.length === 0) return;

  const modules = analyzeAffectedModules(events);
  const ordered = determineExecutionOrder(modules);

  const lock = LockService.getScriptLock();
  try {
    if (lock.tryLock(DEBOUNCE_CONFIG.LOCK_TIMEOUT_MS)) {
      executeModulesInOrder(ordered);
      // 配送方法選択ロジックも併走（対象セル変更時）
      tryRunShippingSelection(events);
    } else {
      console.log('計算実行中のためスキップ');
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * 配送方法選択ロジックの実行（簡易）
 * - 推奨配送方法の提示
 * - DDP対応チェック（EU圏はDDP必須）
 * - 利用可能配送方法の表示（簡易）
 */
function tryRunShippingSelection(events) {
  var need = false;
  for (var i = 0; i < events.length; i++) {
    var r = events[i];
    if (r.row === 37 || r.row === 16 || (r.row === 15 && r.column === 5)) {
      need = true; break;
    }
  }
  if (!need) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.PROFIT);
  if (!sheet) return;

  var targetRegion = (sheet.getRange(PROFIT_CELLS.TARGET_REGION).getDisplayValue() || '').toUpperCase();
  var shippingMethod = (sheet.getRange(PROFIT_CELLS.SHIPPING_METHOD).getDisplayValue() || '').toUpperCase();

  // 簡易ルール: EU/ロシア/東欧→Joom Logistics、その他→eLogi DDP
  var recommended = 'eLogi DDP';
  if (targetRegion.indexOf('EU') >= 0 || targetRegion.indexOf('ロシア') >= 0 || targetRegion.indexOf('東欧') >= 0) {
    recommended = 'Joom Logistics';
  }
  updateCellSilently(sheet.getRange(PROFIT_CELLS.RECOMMENDED_METHOD), recommended);

  // DDP警告: EU圏でDDP非対応の方法を選んだ場合、または日本郵便が選択された場合
  var warning = '';
  // 日本郵便（EMS/ePacket）はDDP非対応・非推奨
  var isJapanPost = shippingMethod.indexOf('EMS') >= 0 || shippingMethod.indexOf('Eパケット') >= 0 || 
                    shippingMethod.indexOf('Eパケ') >= 0 || shippingMethod.indexOf('郵便局') >= 0 ||
                    shippingMethod.indexOf('郵便') >= 0;
  if (isJapanPost) {
    warning = '⚠️ 日本郵便（EMS/ePacket）はDDP非対応・非推奨です。Joom LogisticsまたはeLogi DDPをご利用ください';
  } else if (targetRegion.indexOf('EU') >= 0) {
    var isDdp = shippingMethod.indexOf('DDP') >= 0 || shippingMethod.indexOf('LOGISTICS') >= 0; // 簡易判定
    if (!isDdp && shippingMethod) {
      warning = '⚠️ EU圏への配送にはDDP対応配送方法が必要です';
    }
  }
  updateCellSilently(sheet.getRange(PROFIT_CELLS.DDP_WARNING), warning);

  // 利用可能配送方法の簡易比較を作成
  var comparison = buildShippingComparison(sheet, targetRegion, shippingMethod, recommended);
  updateCellSilently(sheet.getRange(PROFIT_CELLS.AVAILABLE_METHODS), comparison);
}

/**
 * 配送方法の簡易比較（送料・納期・追跡）をテキストで返す
 * 仕様書の推奨ルールをベースに、簡易ヒューリスティクスで生成
 */
function buildShippingComparison(sheet, targetRegion, selectedMethod, recommended) {
  var weight = Number(sheet.getRange(PROFIT_CELLS.WEIGHT_GRAM).getValue()) || 0;
  var entries = [];

  // 候補リスト（簡易）
  var candidates = [recommended, 'DHL', 'FedEx'];
  for (var i = 0; i < candidates.length; i++) {
    var name = candidates[i];
    var eta = estimateLeadTimeDays(name, targetRegion);
    var track = isTrackable(name) ? '追跡:可' : '追跡:不可';
    var fee = estimateShippingCostPlaceholder(name, targetRegion, weight);
    var tag = (selectedMethod && name && selectedMethod.toUpperCase().indexOf(name.toUpperCase()) >= 0) ? '〈選択中〉' : '';
    entries.push(name + ' | 送料:~' + fee + '円 | 納期:~' + eta + '日 | ' + track + (tag ? (' ' + tag) : ''));
  }
  return entries.join(' / ');
}

function estimateLeadTimeDays(methodName, regionUpper) {
  if (!regionUpper) return 7;
  var base = 7;
  if (methodName.indexOf('JOOM') >= 0) base = 8;
  if (methodName.indexOf('ELOGI') >= 0 || methodName.indexOf('DDP') >= 0) base = 10;
  if (methodName.indexOf('DHL') >= 0) base = 5;
  if (methodName.indexOf('FEDEX') >= 0) base = 6;
  if (regionUpper.indexOf('EU') >= 0) base += 1;
  if (regionUpper.indexOf('ロシア') >= 0) base += 3;
  return base;
}

function isTrackable(methodName) {
  methodName = (methodName || '').toUpperCase();
  return (methodName.indexOf('JOOM') >= 0 || methodName.indexOf('DHL') >= 0 || methodName.indexOf('FEDEX') >= 0 || methodName.indexOf('DDP') >= 0);
}

function estimateShippingCostPlaceholder(methodName, regionUpper, weightGram) {
  // 簡易コスト見積り（ダミー式）: 重量係数 + 地域係数 + キャリア係数
  var kg = Math.max(1, Math.ceil((weightGram || 0) / 1000));
  var regionK = 1.0;
  if (regionUpper.indexOf('EU') >= 0) regionK = 1.4;
  else if (regionUpper.indexOf('北米') >= 0) regionK = 1.5;
  else if (regionUpper.indexOf('アジア') >= 0) regionK = 0.9;

  var carrierK = 1.0;
  var up = (methodName || '').toUpperCase();
  if (up.indexOf('JOOM') >= 0) carrierK = 1.1;
  else if (up.indexOf('ELOGI') >= 0 || up.indexOf('DDP') >= 0) carrierK = 1.2;
  else if (up.indexOf('DHL') >= 0) carrierK = 1.5;
  else if (up.indexOf('FEDEX') >= 0) carrierK = 1.4;

  var base = 900; // ベース送料（円）
  var estimate = Math.round(base * kg * regionK * carrierK);
  return estimate;
}

/**
 * 設定値の数値取得（未設定時はデフォルト）
 */
function getNumberSetting(key, defaultValue) {
  try {
    if (typeof getSetting === 'function') {
      var v = getSetting(key);
      if (v === null || v === undefined || v === '') return defaultValue;
      var num = Number(v);
      return isNaN(num) ? defaultValue : num;
    }
  } catch (e) {
    // ignore and use default
  }
  return defaultValue;
}

/**
 * システム設定の読み込みと反映
 * 仕様参照: doc/Joom利益計算機能仕様書.md 3.6
 */
function loadSystemSettings() {
  const settings = {
    // 同時ユーザー数制限
    maxConcurrentUsers: getNumberSetting('利益計算 最大同時ユーザー数', 10),
    
    // デバウンス設定
    debounceTimeMs: getNumberSetting('利益計算 デバウンス時間（ミリ秒）', 300),
    
    // ロックタイムアウト
    lockTimeoutMs: getNumberSetting('利益計算 ロックタイムアウト（ミリ秒）', 30000),
    
    // 諸費用設定
    miscFeesRate: getNumberSetting('利益計算 諸費用（%）', 0),
    miscFeesFixed: getNumberSetting('利益計算 諸費用固定額（円）', 0),
    
    // 各種手数料率
    fedexSurchargeRate: getNumberSetting('利益計算 FedEx割増料金率（%）', 0),
    dhlSurchargeRate: getNumberSetting('利益計算 DHL割増料金率（%）', 0),
    overseasSalesFeeRate: getNumberSetting('利益計算 海外販売手数料率（%）', 0),
    payoneerFeeRate: getNumberSetting('利益計算 Payoneer手数料率（%）', 0),
    
    // ピークシーズン設定
    peakSeasonEnabled: getSetting('利益計算 ピークシーズン有効') === 'true',
    peakSeasonAmount: getNumberSetting('利益計算 ピークシーズン料金（円）', 0),
    
    // 為替レート設定
    exchangeRateManual: getSetting('利益計算 為替レート手動設定') === 'true',
    exchangeRateManualValue: getNumberSetting('利益計算 出品管理為替レート USD/JPY', 150.0)
  };
  
  console.log('システム設定を読み込みました:', settings);
  return settings;
}

/**
 * USD/JPY為替レート取得（手動設定があれば優先）
 * 設定キー:
 *  - 利益計算 為替レート手動設定 (true/false)
 *  - 利益計算 出品管理為替レート USD/JPY (数値)
 *  - 最新為替レート (USD/JPY) (数値)
 */
function getUsdJpyExchangeRate() {
  var manualOn = false;
  try {
    if (typeof getSetting === 'function') {
      var flag = getSetting('利益計算 為替レート手動設定');
      manualOn = String(flag).toLowerCase() === 'true';
    }
  } catch (e) {}

  if (manualOn) {
    var manualRate = getNumberSetting('利益計算 出品管理為替レート USD/JPY', 150.0);
    return manualRate || 150.0;
  }

  // 自動参照（設定シートが更新済み前提）。未設定時はデフォルト150
  var latest = getNumberSetting('最新為替レート (USD/JPY)', 150.0);
  return latest || 150.0;
}

/**
 * 変更セルから影響を受けるモジュールを推定（簡易版）
 */
function analyzeAffectedModules(events) {
  const set = {};
  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    // 仕様の入力セルに基づく簡易判定
    // 販売価格(B12), 仕入価格(B13), 重量(B15), 発送方法(E15), 配送地帯(B16), 寸法(B20,D20,B21), カテゴリー(B19)
    if (ev.row === 15 || ev.row === 20 || ev.row === 21 || ev.column === 5 && ev.row === 15 || ev.row === 16) {
      set[CALCULATION_MODULES.SHIPPING] = true;
    }
    if (ev.row === 13 || (ev.row === 14) || (ev.row === 13 && ev.column === 5)) {
      set[CALCULATION_MODULES.COST] = true;
    }
    if (ev.row === 12 || ev.row === 13 || ev.row === 16 || ev.row === 19) {
      set[CALCULATION_MODULES.PROFIT] = true;
    }
  }
  return Object.keys(set);
}

/**
 * 決定論的実行順序（送料→原価→利益）
 */
function determineExecutionOrder(modules) {
  const order = [
    CALCULATION_MODULES.SHIPPING,
    CALCULATION_MODULES.COST,
    CALCULATION_MODULES.PROFIT
  ];
  return order.filter(function(m){ return modules.indexOf(m) >= 0; });
}

/**
 * モジュールを順序通り実行
 */
function executeModulesInOrder(modules) {
  for (var i = 0; i < modules.length; i++) {
    var m = modules[i];
    if (m === CALCULATION_MODULES.SHIPPING) calculateShippingModule();
    if (m === CALCULATION_MODULES.COST) calculateCostModule();
    if (m === CALCULATION_MODULES.PROFIT) calculateProfitModule();
  }
}

/**
 * 送料計算モジュール（スタブ）
 */
function calculateShippingModule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.PROFIT);
  if (!sheet) return;

  var weight = Number(sheet.getRange(PROFIT_CELLS.WEIGHT_GRAM).getValue()) || 0;
  var h = Number(sheet.getRange(PROFIT_CELLS.HEIGHT_CM).getValue()) || 0;
  var l = Number(sheet.getRange(PROFIT_CELLS.LENGTH_CM).getValue()) || 0;
  var w = Number(sheet.getRange(PROFIT_CELLS.WIDTH_CM).getValue()) || 0;
  var method = sheet.getRange(PROFIT_CELLS.SHIPPING_METHOD).getDisplayValue();
  var region = sheet.getRange(PROFIT_CELLS.REGION).getDisplayValue();

  // 容積重量係数を送料マスタから取得（発送方法に応じて）
  var volumetricFactor = 6000; // デフォルト値
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const methodRange = ss.getRangeByName(NAMED_RANGES.SHIPPING_METHODS);
    const factorRange = ss.getRangeByName(NAMED_RANGES.SHIPPING_VOLUME_FACTOR);
    const methods = methodRange ? methodRange.getValues().flat() : [];
    const factors = factorRange ? factorRange.getValues().flat() : [];
    const normalizedMethod = String(method || '').trim().toLowerCase();

    for (var i = 0; i < methods.length; i++) {
      var candidate = String(methods[i] || '').trim().toLowerCase();
      if (candidate === normalizedMethod) {
        var factor = Number(factors[i]);
        if (!isNaN(factor) && factor > 0) {
          volumetricFactor = factor;
        }
        break;
      }
    }
  } catch (error) {
    console.log('容積重量係数取得エラー、デフォルト値6000を使用:', error);
  }
  
  var volumetricWeight = 0;
  if (h > 0 && l > 0 && w > 0) {
    volumetricWeight = Math.round((h * l * w) / volumetricFactor * 1000); // cm^3/係数 → kg相当をgへ
  }
  var finalWeight = Math.max(weight, volumetricWeight);

  // 地域別送料はシート関数参照が基本。ここではIFERRORを前提に、空なら0。
  // 注意: 送料(E16)と容積重量(D22)は参照式（数式）で自動計算されるため、
  // トリガーからの更新は行わない（数式による自動計算を優先）
  // var shippingCost = Number(sheet.getRange(PROFIT_CELLS.SHIPPING_COST).getValue()) || 0;
  // updateCellSilently(sheet.getRange(PROFIT_CELLS.VOLUMETRIC_WEIGHT), volumetricWeight);
  // updateCellSilently(sheet.getRange(PROFIT_CELLS.SHIPPING_COST), shippingCost);
}

/**
 * 重量と地域を考慮した送料計算（カスタム関数）
 * シート関数として使用可能: =getWeightShippingCost(配送方法, 最終重量, 地域)
 * 
 * @param {string} shippingMethod - 配送方法名（例: "DHL"）
 * @param {number} finalWeight - 最終重量(g)
 * @param {string} region - 地域名（例: "アジア"）
 * @return {number} 送料(¥)
 */
function getWeightShippingCost(shippingMethod, finalWeight, region) {
  try {
    if (!shippingMethod || !finalWeight || !region) {
      return 0;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const weightMaster = ss.getSheetByName('重量別送料マスタ');
    
    if (!weightMaster) {
      console.log('重量別送料マスタが見つかりません');
      return 0;
    }
    
    // データ範囲を取得
    const dataRange = weightMaster.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length < 2) {
      return 0;
    }
    
    // ヘッダー行から地域列のインデックスを取得
    const headers = values[0];
    // 地域名のマッチングを改善（(¥)やスペースを除去して比較）
    const normalizedRegion = String(region).trim();
    const regionColumnIndex = headers.findIndex(h => {
      if (!h) return false;
      const normalizedHeader = String(h).replace(/\s*\(¥\)\s*/g, '').trim();
      return normalizedHeader === normalizedRegion;
    });
    
    if (regionColumnIndex < 2) { // C列(2)以降である必要がある
      // デバッグ用: 利用可能な地域ヘッダーをログ出力
      const availableRegions = headers.slice(2).map(h => String(h || '').replace(/\s*\(¥\)\s*/g, '').trim()).filter(h => h);
      console.log('地域が見つかりません:', normalizedRegion, '利用可能:', availableRegions);
      return 0;
    }
    
    // 配送方法が一致し、最終重量が重量上限以下の行を検索
    // 重量上限は「その値までの重量」を表す（例：500g = 1g〜500gまで）
    let minWeightLimit = Infinity;  // 最小の重量上限を持つ行を選択（最も近い上限値）
    let selectedRow = -1;
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const method = String(row[0] || '').trim();
      const weightLimit = Number(row[1] || 0);
      
      // 最終重量が重量上限以下（finalWeight <= weightLimit）で、かつ最も小さい重量上限を選択
      if (method === String(shippingMethod).trim() && finalWeight <= weightLimit) {
        if (weightLimit < minWeightLimit) {
          minWeightLimit = weightLimit;
          selectedRow = i;
        }
      }
    }
    
    if (selectedRow >= 0 && values[selectedRow][regionColumnIndex] !== null && values[selectedRow][regionColumnIndex] !== undefined && values[selectedRow][regionColumnIndex] !== '') {
      const shippingCost = Number(values[selectedRow][regionColumnIndex]);
      if (isNaN(shippingCost)) {
        console.log('送料値が数値ではありません:', values[selectedRow][regionColumnIndex]);
        return 0;
      }
      return shippingCost;
    }
    
    // デバッグ用: 検索結果をログ出力
    if (selectedRow < 0) {
      console.log('該当する重量段階が見つかりません:', {
        method: shippingMethod,
        finalWeight: finalWeight,
        region: region,
        note: '最終重量が重量上限を超えています。より大きい重量段階のデータが必要です。'
      });
    } else {
      console.log('送料を取得しました:', {
        method: shippingMethod,
        finalWeight: finalWeight,
        region: region,
        weightLimit: minWeightLimit,
        shippingCost: values[selectedRow][regionColumnIndex]
      });
    }
    
    return 0;
  } catch (error) {
    console.log('送料計算エラー:', error);
    return 0;
  }
}

/**
 * 配送方法×地域の繁忙期料金を取得
 * シート関数としても使用可能
 * @param {string} shippingMethod
 * @param {string} region
 * @return {number}
 */
function getPeakSeasonFee(shippingMethod, region) {
  try {
    if (!shippingMethod || !region) {
      return 0;
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const methodsRange = ss.getRangeByName(NAMED_RANGES.PEAK_SEASON_METHODS);
    const regionsRange = ss.getRangeByName(NAMED_RANGES.PEAK_SEASON_REGIONS);
    const amountsRange = ss.getRangeByName(NAMED_RANGES.PEAK_SEASON_AMOUNT);
    if (!methodsRange || !regionsRange || !amountsRange) {
      return 0;
    }

    const methods = methodsRange.getValues().flat();
    const regions = regionsRange.getValues().flat();
    const amounts = amountsRange.getValues().flat();
    const normalizedMethod = String(shippingMethod).trim().toLowerCase();
    const normalizedRegion = String(region).trim().toLowerCase();

    for (let i = 0; i < methods.length; i++) {
      const methodCandidate = String(methods[i] || '').trim().toLowerCase();
      const regionCandidate = String(regions[i] || '').trim().toLowerCase();
      if (methodCandidate === normalizedMethod && regionCandidate === normalizedRegion) {
        const amount = Number(amounts[i]);
        return isNaN(amount) ? 0 : amount;
      }
    }
  } catch (error) {
    console.log('繁忙期料金取得エラー:', error);
  }
  return 0;
}

/**
 * 仕入原価計算モジュール（スタブ）
 */
function calculateCostModule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.PROFIT);
  if (!sheet) return;

  var purchase = Number(sheet.getRange(PROFIT_CELLS.PURCHASE_PRICE).getValue()) || 0;
  var discountRate = Number(sheet.getRange(PROFIT_CELLS.DISCOUNT_RATE).getValue()) || 0;
  var discountPoints = Number(sheet.getRange(PROFIT_CELLS.DISCOUNT_POINTS).getValue()) || 0;

  var cost = Math.max(0, Math.round(purchase * (1 - discountRate / 100) - discountPoints));
  // 注意: 仕入原価(E14)は参照式（数式）で自動計算されるため、
  // トリガーからの更新は行わない（数式による自動計算を優先）
  // updateCellSilently(sheet.getRange(PROFIT_CELLS.COST_RESULT), cost);
}

/**
 * 基本利益計算モジュール（スタブ）
 */
function calculateProfitModule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.PROFIT);
  if (!sheet) return;

  var selling = Number(sheet.getRange(PROFIT_CELLS.SELLING_PRICE).getValue()) || 0;
  var cost = Number(sheet.getRange(PROFIT_CELLS.COST_RESULT).getValue()) || 0;
  var shipping = Number(sheet.getRange(PROFIT_CELLS.SHIPPING_COST).getValue()) || 0;
  var joomFeeRate = getNumberSetting('利益計算 Joom手数料率（%）', Number(sheet.getRange(PROFIT_CELLS.JOOM_FEE_RATE).getValue()) || 12.7);
  var miscPercent = getNumberSetting('利益計算 諸費用（%）', 0);
  var miscFixed = getNumberSetting('利益計算 諸費用固定額（円）', 0);
  var joomFee = Math.round(selling * joomFeeRate / 100);
  var shippingSurcharge = Number(sheet.getRange(PROFIT_CELLS.SHIPPING_SURCHARGE).getValue()) || 0;
  var peakSeasonFee = Number(sheet.getRange(PROFIT_CELLS.PEAK_SEASON_FEE).getValue()) || 0;
  var refund = Number(sheet.getRange(PROFIT_CELLS.REFUND_AMOUNT).getValue()) || 0;

  var miscCost = Math.round(selling * (miscPercent / 100)) + Math.round(miscFixed);
  var profit = Math.round(selling - cost - shipping - shippingSurcharge - peakSeasonFee - joomFee - miscCost + refund);
  var profitRate = selling > 0 ? Math.round((profit / selling) * 10000) / 100 : 0;
  var profitWithRefund = Math.round(profit + refund);
  var profitRateWithRefund = selling > 0 ? Math.round((profitWithRefund / selling) * 10000) / 100 : 0;

  // 注意: 以下のセルは参照式（数式）で自動計算されるため、
  // トリガーからの更新は行わない（数式による自動計算を優先）
  // - E17 (Joom手数料): =ROUND(B12*IFERROR(B17,0.127),0)
  // - B8 (利益額): =B12-E14-E16-E17+B18
  // - B9 (利益率): =IF(B12>0,B8/B12,0)
  // - B10 (還付込利益額): =B8+B18
  // - B11 (還付込利益率): =IF(B12>0,B10/B12,0)
  // updateCellSilently(sheet.getRange(PROFIT_CELLS.JOOM_FEE_YEN), joomFee);
  // updateCellSilently(sheet.getRange(PROFIT_CELLS.PROFIT_YEN), profit);
  // updateCellSilently(sheet.getRange(PROFIT_CELLS.PROFIT_RATE), profitRate);
  // updateCellSilently(sheet.getRange(PROFIT_CELLS.PROFIT_YEN_REFUND), profitWithRefund);
  // updateCellSilently(sheet.getRange(PROFIT_CELLS.PROFIT_RATE_REFUND), profitRateWithRefund);
}

/**
 * ループ防止: プログラム更新フラグの操作
 */
function isProgrammaticUpdateActive() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty(PROGRAMMATIC_FLAG_KEY) === '1';
}

function beginProgrammaticUpdate() {
  PropertiesService.getScriptProperties().setProperty(PROGRAMMATIC_FLAG_KEY, '1');
}

function endProgrammaticUpdate() {
  PropertiesService.getScriptProperties().deleteProperty(PROGRAMMATIC_FLAG_KEY);
}

/**
 * onEditを発火させないようにセル更新を行う
 */
function updateCellSilently(range, value) {
  beginProgrammaticUpdate();
  try {
    range.setValue(value);
  } finally {
    endProgrammaticUpdate();
  }
}


