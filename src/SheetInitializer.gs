/**
 * 全シートの初期化を管理するメインスクリプト
 * 在庫管理ツール システム開発
 * 作成日: 2025-09-27
 */

/**
 * 全シートの初期化を実行（Joom注文連携対応版）
 */
function initializeAllSheets() {
  console.log('在庫管理ツールのシート初期化を開始します（Joom注文連携対応版）...');
  
  const initializationResults = [];
  
  try {
    // 各シートを順次初期化
    console.log('1. 在庫管理シートの初期化...');
    initializeInventorySheet();
    initializationResults.push('✓ 在庫管理シート');
    
    console.log('2. 売上管理シートの初期化...');
    initializeSalesSheet();
    initializationResults.push('✓ 売上管理シート');
    
    console.log('3. 仕入れ元マスターシートの初期化...');
    initializeSupplierMasterSheet();
    initializationResults.push('✓ 仕入れ元マスターシート');
    
    console.log('4. 価格履歴シートの初期化...');
    initializePriceHistorySheet();
    initializationResults.push('✓ 価格履歴シート');
    
    console.log('5. 設定シートの初期化...');
    initializeSettingsSheet();
    initializationResults.push('✓ 設定シート');
    
    console.log('6. 利益計算シートの初期化...');
    initializeProfitSheet();
    initializationResults.push('✓ 利益計算シート（事前利益計算特化）');
    
    // 利益計算シートの検証を実行
    console.log('7. 利益計算シートの検証...');
    const profitSheetValid = verifyProfitSheetLayout();
    if (profitSheetValid) {
      initializationResults.push('✓ 利益計算シート検証');
    } else {
      initializationResults.push('⚠ 利益計算シート検証（一部警告あり）');
    }
    
    // カスタムメニューを設定
    console.log('8. カスタムメニューの設定...');
    setupCustomMenu();
    initializationResults.push('✓ カスタムメニュー');
    
    console.log('全シートの初期化が完了しました（Joom注文連携対応版）');
    
    // 詳細な結果を表示
    const resultMessage = '在庫管理ツールの全シートが正常に作成されました（Joom注文連携対応版）。\n\n' +
                         '初期化結果:\n' + initializationResults.join('\n') + '\n\n' +
                         '利益計算機能が利用可能になりました。';
    
    SpreadsheetApp.getUi().alert('シート初期化完了', resultMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('シート初期化中にエラーが発生しました:', error);
    
    // エラー発生時の結果表示
    const errorMessage = 'シート初期化中にエラーが発生しました。\n\n' +
                        '完了した項目:\n' + initializationResults.join('\n') + '\n\n' +
                        'エラー詳細: ' + error.message + '\n\n' +
                        'ログを確認して詳細を確認してください。';
    
    SpreadsheetApp.getUi().alert('エラー', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 利益計算シート初期化（名前付き範囲・データ検証）
 * 仕様参照: doc/Joom利益計算機能仕様書.md 3.1 / 3.1.0.2 / 3.1.0.3
 */
function initializeProfitSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.PROFIT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.PROFIT);
  }

  // ヘッダーと項目名の設定
  setupProfitSheetHeaders(sheet);
  
  // セルの書式設定
  setupProfitSheetFormatting(sheet);
  
  // レイアウト調整
  setupProfitSheetLayout(sheet);
  
  // データ検証の強化
  setupProfitSheetDataValidation(sheet);
  
  // セル保護設定
  setupProfitSheetProtection(sheet);
  
  // 条件付き書式設定
  setupProfitSheetConditionalFormatting(sheet);

  // 名前付き範囲（簡易）
  try {
    // 送料・関税・為替（実データは各マスタにて管理。ここではセル/列のプレースホルダ設定）
    // 為替レート: 設定シートのA列想定（見出しを除く）→ 最新為替レートセルの仮置き
    const settings = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    if (settings) {
      const latestFxCell = settings.getRange('A2');
      ss.setNamedRange('為替レート_USDJPY', latestFxCell);
    }
  } catch (e) {
    console.log('名前付き範囲の設定中に警告:', e);
  }

  // マスタ参照の名前付き範囲（存在時のみ設定）
  try {
    const shippingMaster = ss.getSheetByName('Joom送料マスタ');
    if (shippingMaster) {
      // 例: 配送方法名B列、容積重量係数N列
      // 注意: G-M列（地域別送料）は削除済み。送料計算には重量別送料マスタを使用
      ss.setNamedRange(NAMED_RANGES.SHIPPING_METHODS, shippingMaster.getRange('B2:B'));
      ss.setNamedRange(NAMED_RANGES.SHIPPING_SURCHARGE, shippingMaster.getRange('N2:N'));
      ss.setNamedRange(NAMED_RANGES.SHIPPING_VOLUME_FACTOR, shippingMaster.getRange('O2:O'));
    }
    const dutyMaster = ss.getSheetByName('関税率マスタ');
    if (dutyMaster) {
      ss.setNamedRange('関税率マスタ_カテゴリー', dutyMaster.getRange('B2:B'));
      ss.setNamedRange('関税率マスタ_税率', dutyMaster.getRange('D2:J'));
    }
    const joomFeeMaster = ss.getSheetByName('Joom手数料マスタ');
    if (joomFeeMaster) {
      ss.setNamedRange('Joom手数料マスタ_カテゴリー', joomFeeMaster.getRange('B2:B'));  // B列: カテゴリー名
      ss.setNamedRange('Joom手数料マスタ_手数料率', joomFeeMaster.getRange('D2:J'));     // D-J列: 各地域の手数料率
    }
    const peakSeasonMaster = ss.getSheetByName(SHEET_NAMES.PEAK_SEASON_MASTER);
    if (peakSeasonMaster) {
      ss.setNamedRange(NAMED_RANGES.PEAK_SEASON_METHODS, peakSeasonMaster.getRange('A2:A'));
      ss.setNamedRange(NAMED_RANGES.PEAK_SEASON_REGIONS, peakSeasonMaster.getRange('B2:B'));
      ss.setNamedRange(NAMED_RANGES.PEAK_SEASON_AMOUNT, peakSeasonMaster.getRange('C2:C'));
    }
  } catch (e) {
    console.log('マスタ名前付き範囲の設定中に警告:', e);
  }

  // データ検証の設定
  setupProfitSheetDataValidation(sheet);

  // 参照式の設定
  setupProfitSheetFormulas(sheet);
  
    // マスタシートの存在確認と簡易マスタの作成
    createMasterSheetsIfNeeded(ss);
    
    // 既存の送料マスタを更新（U列追加対応）
    updateExistingShippingMaster(ss);
    
    console.log('利益計算シートの初期化が完了しました');
}

/**
 * マスタシートの存在確認と簡易マスタの作成
 */
function createMasterSheetsIfNeeded(ss) {
  try {
    // 送料マスタの確認・作成
    let shippingMaster = ss.getSheetByName('Joom送料マスタ');
    if (!shippingMaster) {
      console.log('送料マスタが存在しません。簡易マスタを作成します。');
      shippingMaster = ss.insertSheet('Joom送料マスタ');
      setupShippingMaster(shippingMaster);
    }
    
    // 関税率マスタの確認・作成
    let dutyMaster = ss.getSheetByName('関税率マスタ');
    if (!dutyMaster) {
      console.log('関税率マスタが存在しません。簡易マスタを作成します。');
      dutyMaster = ss.insertSheet('関税率マスタ');
      setupDutyMaster(dutyMaster);
    }
    
    // Joom手数料マスタの確認・作成
    let joomFeeMaster = ss.getSheetByName('Joom手数料マスタ');
    if (!joomFeeMaster) {
      console.log('Joom手数料マスタが存在しません。簡易マスタを作成します。');
      joomFeeMaster = ss.insertSheet('Joom手数料マスタ');
      setupJoomFeeMaster(joomFeeMaster);
    }
    
    // 重量別送料マスタの確認・作成
    let weightShippingMaster = ss.getSheetByName('重量別送料マスタ');
    if (!weightShippingMaster) {
      console.log('重量別送料マスタが存在しません。簡易マスタを作成します。');
      weightShippingMaster = ss.insertSheet('重量別送料マスタ');
      setupWeightShippingMaster(weightShippingMaster);
    }

    // 繁忙期料金マスタの確認・作成
    let peakSeasonMaster = ss.getSheetByName(SHEET_NAMES.PEAK_SEASON_MASTER);
    if (!peakSeasonMaster) {
      console.log('繁忙期料金マスタが存在しません。簡易マスタを作成します。');
      peakSeasonMaster = ss.insertSheet(SHEET_NAMES.PEAK_SEASON_MASTER);
      setupPeakSeasonMaster(peakSeasonMaster);
    }
    
  } catch (error) {
    console.log('マスタシート作成中にエラーが発生しました:', error);
  }
}

/**
 * 既存の送料マスタを更新（U列追加対応）
 */
function updateExistingShippingMaster(ss) {
  try {
    const shippingMaster = ss.getSheetByName('Joom送料マスタ');
    if (shippingMaster) {
      console.log('既存の送料マスタを更新します（U列追加対応）...');
      setupShippingMaster(shippingMaster);
    }
  } catch (error) {
    console.log('送料マスタ更新中にエラー:', error);
  }
}

/**
 * Joom送料マスタの簡易セットアップ
 * 既存データがある場合はヘッダーの更新のみ、なければ新規作成
 * 
 * 役割:
 * - 配送方法の基本情報管理（容積重量係数、Joom手数料率、DDP対応など）
 * - G-M列（地域別送料）は削除済み。送料計算には重量別送料マスタを使用してください
 * - このマスタは主に「配送方法マスタ」として機能
 */
function setupShippingMaster(sheet) {
  const SURCHARGE_HEADER = 'サーチャージ(¥)';
  const VOLUME_HEADER = '容積重量係数';

  // ヘッダー行（仕様書に基づく全列構成）
  // G-M列（地域別送料）は削除済み。送料計算には重量別送料マスタを使用してください。
  const headers = [
    ['配送方法ID', '配送方法名', '配送タイプ', 'DDP対応', '対応地域', '重量上限(g)',
     'Joom手数料率(%)', '代行手数料(¥)', '関税込み', '追跡可能', '納期(日)', '有効フラグ', '備考',
     SURCHARGE_HEADER, VOLUME_HEADER]
  ];
  
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const headerRowValues = sheet.getRange(1, 1, 1, Math.max(lastColumn, headers[0].length)).getValues()[0];
  const hasData = lastRow > 1;
  const isHeaderEmpty = headerRowValues.every(value => value === '');

  if (isHeaderEmpty) {
    try {
      sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    } catch (e) {
      console.log('ヘッダー初期設定でエラー:', e);
    }
  } else {
    let surchargeCol = headerRowValues.indexOf(SURCHARGE_HEADER) + 1;
    let volumeCol = headerRowValues.indexOf(VOLUME_HEADER) + 1;

    if (surchargeCol === 0) {
      if (volumeCol > 0) {
        sheet.insertColumnBefore(volumeCol);
        surchargeCol = volumeCol;
        volumeCol += 1;
      } else {
        sheet.insertColumnAfter(sheet.getLastColumn());
        surchargeCol = sheet.getLastColumn();
      }
      sheet.getRange(1, surchargeCol).setValue(SURCHARGE_HEADER);
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, surchargeCol, sheet.getLastRow() - 1, 1).setValue(0);
      }
    } else {
      sheet.getRange(1, surchargeCol).setValue(SURCHARGE_HEADER);
      if (sheet.getLastRow() > 1) {
        const surchargeRange = sheet.getRange(2, surchargeCol, sheet.getLastRow() - 1, 1);
        const surchargeValues = surchargeRange.getValues();
        let surchargeNeedsUpdate = false;
        const normalized = surchargeValues.map(row => {
          const num = Number(row[0]);
          if (row[0] === '' || isNaN(num)) {
            surchargeNeedsUpdate = true;
            return [0];
          }
          return [num];
        });
        if (surchargeNeedsUpdate) {
          surchargeRange.setValues(normalized);
        }
      }
    }

    if (volumeCol === 0) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      volumeCol = sheet.getLastColumn();
      sheet.getRange(1, volumeCol).setValue(VOLUME_HEADER);
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, volumeCol, sheet.getLastRow() - 1, 1).setValue(6000);
      }
    } else {
      sheet.getRange(1, volumeCol).setValue(VOLUME_HEADER);
      if (sheet.getLastRow() > 1) {
        const volumeRange = sheet.getRange(2, volumeCol, sheet.getLastRow() - 1, 1);
        const volumeValues = volumeRange.getValues();
        let volumeNeedsUpdate = false;
        const normalizedVolume = volumeValues.map(row => {
          const num = Number(row[0]);
          if (!num) {
            volumeNeedsUpdate = true;
            return [6000];
          }
          return [num];
        });
        if (volumeNeedsUpdate) {
          volumeRange.setValues(normalizedVolume);
        }
      }
    }

    try {
      sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    } catch (e) {
      console.log('ヘッダー更新でエラー:', e);
    }
  }
  
  if (!hasData) {
    console.log('送料マスタのサンプルデータを設定します...');
    const sampleData = [
      ['LOG001', 'Joom Logistics', '香港倉庫', true, 'EU圏,ロシア,東欧', 0, 12.7, 0, true, true, 7, true, 'Joom推奨配送（送料は重量別送料マスタを参照）', 0, 6000],
      ['ELO001', 'eLogi DDP', 'DDP直送', true, 'EU圏,ロシア,東欧', 0, 12.7, 0, true, true, 6, true, 'eLogi DDP配送（送料は重量別送料マスタを参照）', 0, 6000],
      ['DHL001', 'DHL', 'DDP直送', true, '北米,アジア,オセアニア,南米', 0, 12.7, 0, true, true, 5, true, 'DHL国際配送（送料は重量別送料マスタを参照）', 0, 5000],
      ['FED001', 'FedEx', 'DDP直送', true, '北米,アジア,オセアニア,南米', 0, 12.7, 0, true, true, 4, true, 'FedEx国際配送（送料は重量別送料マスタを参照）', 0, 3571]
    ];

    try {
      for (let i = 0; i < sampleData.length; i++) {
        sheet.getRange(i + 2, 1, 1, sampleData[i].length).setValues([sampleData[i]]);
      }
      console.log('送料マスタのサンプルデータを設定しました');
    } catch (e) {
      console.log('サンプルデータ設定でエラー:', e);
    }
  } else {
    console.log('送料マスタは既にデータが存在します（サーチャージ列を確認済み）');
  }
  
  console.log('送料マスタのセットアップが完了しました（サーチャージ列・容積重量係数列を整備）');
}

/**
 * 関税率マスタの簡易セットアップ
 */
function setupDutyMaster(sheet) {
  // ヘッダー行
  sheet.getRange('A1').setValue('カテゴリー');
  sheet.getRange('B1').setValue('カテゴリー名');
  sheet.getRange('C1').setValue('EU圏');
  sheet.getRange('D1').setValue('ロシア');
  sheet.getRange('E1').setValue('東欧');
  sheet.getRange('F1').setValue('北米');
  sheet.getRange('G1').setValue('アジア');
  sheet.getRange('H1').setValue('オセアニア');
  sheet.getRange('I1').setValue('南米');
  
  // データ行
  const data = [
    ['一般', '一般', 0, 0, 0, 0, 0, 0, 0],
    ['家電', '家電', 5, 5, 5, 0, 0, 0, 0],
    ['アパレル', 'アパレル', 12, 12, 12, 0, 0, 0, 0],
    ['雑貨', '雑貨', 3, 3, 3, 0, 0, 0, 0],
    ['ホビー', 'ホビー', 8, 8, 8, 0, 0, 0, 0]
  ];
  
  for (let i = 0; i < data.length; i++) {
    sheet.getRange(i + 2, 1, 1, data[i].length).setValues([data[i]]);
  }
  
  console.log('関税率マスタの簡易セットアップが完了しました');
}

/**
 * Joom手数料マスタの簡易セットアップ
 */
function setupJoomFeeMaster(sheet) {
  // ヘッダー行
  sheet.getRange('A1').setValue('カテゴリーID');
  sheet.getRange('B1').setValue('カテゴリー名');
  sheet.getRange('C1').setValue('商品説明');
  sheet.getRange('D1').setValue('EU圏');
  sheet.getRange('E1').setValue('ロシア');
  sheet.getRange('F1').setValue('東欧');
  sheet.getRange('G1').setValue('北米');
  sheet.getRange('H1').setValue('アジア');
  sheet.getRange('I1').setValue('オセアニア');
  sheet.getRange('J1').setValue('南米');
  sheet.getRange('K1').setValue('更新日時');
  sheet.getRange('L1').setValue('有効フラグ');
  sheet.getRange('M1').setValue('備考');
  
  // ヘッダーの書式設定
  const headerRange = sheet.getRange('A1:M1');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // データ行（関税率マスタと同じカテゴリーで、デフォルト手数料率12.7%を設定）
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const data = [
    ['CAT001', '一般', '一般商品', 12.7, 12.7, 12.7, 12.7, 12.7, 12.7, 12.7, now, true, ''],
    ['CAT002', '家電', '家電製品', 12.7, 12.7, 12.7, 12.7, 12.7, 12.7, 12.7, now, true, ''],
    ['CAT003', 'アパレル', '衣料品・ファッション', 12.7, 12.7, 12.7, 12.7, 12.7, 12.7, 12.7, now, true, ''],
    ['CAT004', '雑貨', '日用品・雑貨', 12.7, 12.7, 12.7, 12.7, 12.7, 12.7, 12.7, now, true, ''],
    ['CAT005', 'ホビー', '趣味・ホビー用品', 12.7, 12.7, 12.7, 12.7, 12.7, 12.7, 12.7, now, true, '']
  ];
  
  for (let i = 0; i < data.length; i++) {
    sheet.getRange(i + 2, 1, 1, data[i].length).setValues([data[i]]);
  }
  
  // 手数料率列の書式設定（パーセント表示）
  const feeRateRange = sheet.getRange('D2:J' + (data.length + 1));
  feeRateRange.setNumberFormat('0.0');
  
  // 列幅の自動調整
  sheet.autoResizeColumns(1, 13);
  
  console.log('Joom手数料マスタの簡易セットアップが完了しました');
}


/**
 * 重量別送料マスタのセットアップ
 * eShip方式：重量段階別×地域別の送料テーブル
 * 
 * 役割:
 * - 送料計算の主データソース
 * - 配送方法×重量段階×地域の組み合わせで送料を管理
 * - Joom送料マスタの送料列（G-M列）は非推奨のため、こちらを使用
 */
function setupWeightShippingMaster(sheet) {
  // ヘッダー行（仕様書とeShipの例を参考）
  const headers = [
    ['配送方法名', '重量上限(g)', 'EU圏(¥)', 'ロシア(¥)', '東欧(¥)', '北米(¥)', 'アジア(¥)', 'オセアニア(¥)', '南米(¥)']
  ];
  
  // 既存データの確認
  const existingData = sheet.getDataRange().getValues();
  const hasData = existingData.length > 1; // ヘッダー行以外にデータがあるか
  
  // ヘッダーを設定
  if (!hasData || sheet.getRange('A1').getValue() !== '配送方法名') {
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    console.log('重量別送料マスタのヘッダーを設定しました');
  }
  
  // 既存データがない場合のみサンプルデータを設定
  if (!hasData) {
    console.log('重量別送料マスタのサンプルデータを設定します...');
  } else {
    // 既存データがある場合でも、軽量データ（100g、200g）が不足していれば追加
    console.log('重量別送料マスタの既存データを確認中...');
    checkAndAddLightWeightData(sheet);
  }
  
  if (!hasData) {
    
    // DHLの例（eShipのDHL送料設定を参考）
    // 軽量商品（100g以下、200g以下）の送料を追加
    const dhlData = [
      ['DHL', 100, 1500, 0, 0, 1800, 1300, 0, 3000],  // 100g以下
      ['DHL', 200, 1650, 0, 0, 2000, 1450, 0, 3250],  // 200g以下
      ['DHL', 500, 1822, 0, 0, 2191, 1619, 0, 3527],
      ['DHL', 1000, 1822, 0, 0, 2191, 1619, 0, 3527],
      ['DHL', 1500, 2462, 0, 0, 3000, 2150, 0, 4500],
      ['DHL', 2000, 3102, 0, 0, 3809, 2681, 0, 5473],
      ['DHL', 2500, 3742, 0, 0, 4618, 3212, 0, 6446],
      ['DHL', 3000, 4382, 0, 0, 5427, 3743, 0, 7419],
      ['DHL', 3500, 5022, 0, 0, 6236, 4274, 0, 8392],
      ['DHL', 4000, 5662, 0, 0, 7045, 4805, 0, 9365],
      ['DHL', 4500, 6302, 0, 0, 7854, 5336, 0, 10338],
      ['DHL', 5000, 6942, 0, 0, 8663, 5867, 0, 11311],
      ['DHL', 5500, 7582, 0, 0, 8774, 6398, 0, 12284],
      ['DHL', 6000, 8222, 0, 0, 8885, 6929, 0, 13257],
      ['DHL', 6500, 8862, 0, 0, 8996, 7460, 0, 14230],
      ['DHL', 7000, 9502, 0, 0, 9107, 7991, 0, 15203],
      ['DHL', 7500, 10142, 0, 0, 9218, 8522, 0, 16176],
      ['DHL', 8000, 10782, 0, 0, 9329, 9053, 0, 17149],
      ['DHL', 8500, 11422, 0, 0, 9240, 9584, 0, 18122],
      ['DHL', 9000, 12062, 0, 0, 9251, 10115, 0, 19095],
      ['DHL', 9500, 12702, 0, 0, 9262, 10646, 0, 20068],
      ['DHL', 10000, 13342, 0, 0, 9273, 11177, 0, 21041],
      ['DHL', 10500, 13982, 0, 0, 9284, 11708, 0, 21861]
    ];
    
    // その他の配送方法のサンプルデータ（簡易版）
    const otherData = [
      ['Joom Logistics', 100, 1000, 1300, 1100, 0, 700, 0, 0],  // 軽量対応
      ['Joom Logistics', 200, 1100, 1400, 1200, 0, 750, 0, 0],  // 軽量対応
      ['Joom Logistics', 500, 1200, 1500, 1300, 0, 800, 0, 0],
      ['Joom Logistics', 1000, 1300, 1600, 1400, 0, 900, 0, 0],
      ['Joom Logistics', 2000, 1500, 1800, 1600, 0, 1100, 0, 0],
      ['Joom Logistics', 5000, 2000, 2300, 2100, 0, 1500, 0, 0],
      ['eLogi DDP', 100, 1300, 1600, 1400, 0, 0, 0, 0],  // 軽量対応
      ['eLogi DDP', 200, 1400, 1700, 1500, 0, 0, 0, 0],  // 軽量対応
      ['eLogi DDP', 500, 1500, 1800, 1600, 0, 0, 0, 0],
      ['eLogi DDP', 1000, 1600, 1900, 1700, 0, 0, 0, 0],
      ['eLogi DDP', 2000, 1800, 2100, 1900, 0, 0, 0, 0],
      ['eLogi DDP', 5000, 2300, 2600, 2400, 0, 0, 0, 0]
    ];
    
    const allData = [...dhlData, ...otherData];
    
    try {
      for (let i = 0; i < allData.length; i++) {
        sheet.getRange(i + 2, 1, 1, allData[i].length).setValues([allData[i]]);
      }
      console.log('重量別送料マスタのサンプルデータを設定しました（' + allData.length + '行）');
    } catch (e) {
      console.log('サンプルデータ設定でエラー:', e);
    }
  } else {
    console.log('重量別送料マスタは既にデータが存在します');
  }
  
  console.log('重量別送料マスタのセットアップが完了しました');
}

/**
 * 繁忙期料金マスタのセットアップ
 * 配送方法×地域別の繁忙期追加料金を管理
 */
function setupPeakSeasonMaster(sheet) {
  const headers = [
    ['配送方法名', '地域', '繁忙期料金(¥)', '開始日', '終了日', '備考']
  ];

  const existingData = sheet.getDataRange().getValues();
  const hasData = existingData.length > 1;

  const headerRow = sheet.getRange(1, 1, 1, headers[0].length);
  headerRow.setValues(headers);
  headerRow.setFontWeight('bold');
  headerRow.setBackground('#F3F2FF');

  if (!hasData) {
    const sampleData = [
      ['Joom Logistics', 'EU圏', 200, '', '', 'EUハイシーズン追加料金'],
      ['Joom Logistics', 'ロシア', 220, '', '', 'ロシア冬季物流加算'],
      ['Joom Logistics', '東欧', 210, '', '', '東欧繁忙期間料金'],
      ['eLogi DDP', '北米', 150, '', '', '北米向けピークシーズン'],
      ['DHL', 'アジア', 100, '', '', '旧正月期間対応'],
      ['DHL', 'オセアニア', 130, '', '', '夏季観光シーズン'],
      ['FedEx', '南米', 180, '', '', '年末セール対応']
    ];
    for (let i = 0; i < sampleData.length; i++) {
      sheet.getRange(i + 2, 1, 1, sampleData[i].length).setValues([sampleData[i]]);
    }
  }

  // 書式設定
  sheet.getRange('C2:C').setNumberFormat('¥#,##0');
  sheet.autoResizeColumns(1, headers[0].length);
}

/**
 * 既存の重量別送料マスタに軽量データ（100g、200g）が不足している場合に追加
 */
function checkAndAddLightWeightData(sheet) {
  try {
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length < 2) return;
    
    // 既存の配送方法と重量上限の組み合わせを確認
    const existingCombinations = new Set();
    for (let i = 1; i < values.length; i++) {
      const method = String(values[i][0] || '').trim();
      const weightLimit = Number(values[i][1] || 0);
      if (method && weightLimit > 0) {
        existingCombinations.add(`${method}_${weightLimit}`);
      }
    }
    
    // 追加すべき軽量データ
    const lightWeightData = [
      ['DHL', 100, 1500, 0, 0, 1800, 1300, 0, 3000],
      ['DHL', 200, 1650, 0, 0, 2000, 1450, 0, 3250],
      ['Joom Logistics', 100, 1000, 1300, 1100, 0, 700, 0, 0],
      ['Joom Logistics', 200, 1100, 1400, 1200, 0, 750, 0, 0],
      ['eLogi DDP', 100, 1300, 1600, 1400, 0, 0, 0, 0],
      ['eLogi DDP', 200, 1400, 1700, 1500, 0, 0, 0, 0]
    ];
    
    const rowsToAdd = [];
    for (const rowData of lightWeightData) {
      const method = rowData[0];
      const weightLimit = rowData[1];
      const key = `${method}_${weightLimit}`;
      
      if (!existingCombinations.has(key)) {
        rowsToAdd.push(rowData);
      }
    }
    
    if (rowsToAdd.length > 0) {
      console.log(`軽量データ（100g、200g）が不足しています。${rowsToAdd.length}件を追加します...`);
      const lastRow = sheet.getLastRow();
      
      for (let i = 0; i < rowsToAdd.length; i++) {
        const rowData = rowsToAdd[i];
        // 最後に追加
        sheet.getRange(lastRow + i + 1, 1, 1, rowData.length).setValues([rowData]);
      }
      
      console.log(`${rowsToAdd.length}件の軽量データを追加しました`);
    } else {
      console.log('軽量データは既に存在しています');
    }
  } catch (error) {
    console.log('軽量データ追加中にエラー:', error);
  }
}

/**
 * 利益計算シートのヘッダーと項目名を設定
 */
function setupProfitSheetHeaders(sheet) {
  try {
    // 1. メインタイトル
    sheet.getRange('A1').setValue('Joom利益計算シート');
    sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
    
    // 2. 商品ID入力セル
    sheet.getRange('A2').setValue('商品ID:');
    sheet.getRange('A2').setFontWeight('bold');
    
    // 2. 出品情報セクション
    sheet.getRange('A4').setValue('【出品情報】');
    sheet.getRange('A4').setFontWeight('bold').setBackground('#E8F4FD');
    sheet.getRange('A5').setValue('Joom商品SKU:');
    sheet.getRange('D5').setValue('商品名:');
    sheet.getRange('A6').setValue('仕入URL:');
    sheet.getRange('D6').setValue('仕入元:');
    
    // 3. 利益計算セクション
    sheet.getRange('A7').setValue('【利益計算】');
    sheet.getRange('A7').setFontWeight('bold').setBackground('#E8F4FD');
    sheet.getRange('A8').setValue('利益額(¥):');
    sheet.getRange('A9').setValue('利益率(%):');
    sheet.getRange('A10').setValue('還付込利益額(¥):');
    sheet.getRange('A11').setValue('還付込利益率(%):');
    sheet.getRange('A12').setValue('販売価格(¥):');
    sheet.getRange('D12').setValue('出品数量:');
    sheet.getRange('A13').setValue('仕入価格(¥):');
    sheet.getRange('D13').setValue('割引率(%):');
    sheet.getRange('A14').setValue('割引ポイント:');
    sheet.getRange('D14').setValue('仕入原価(¥):');
    sheet.getRange('A15').setValue('重量(g):');
    sheet.getRange('D15').setValue('発送方法:');
    sheet.getRange('A16').setValue('配送地帯:');
    sheet.getRange('D16').setValue('送料(¥):');
    sheet.getRange('A17').setValue('Joom手数料率(%):');
    sheet.getRange('D17').setValue('Joom手数料(¥):');
    sheet.getRange('A18').setValue('返金額(¥):');
  sheet.getRange('D18').setValue('サーチャージ(¥):');
  sheet.getRange('D19').setValue('繁忙期料金(¥):');
    sheet.getRange('A19').setValue('商品カテゴリー:');
    
    // 4. 容積重量チェックセクション
    sheet.getRange('A20').setValue('【容積重量チェック】');
    sheet.getRange('A20').setFontWeight('bold').setBackground('#E8F4FD');
    sheet.getRange('A21').setValue('高さ(cm):');
    sheet.getRange('C21').setValue('長さ(cm):');
    sheet.getRange('A22').setValue('幅(cm):');
    sheet.getRange('C22').setValue('容積重量(g):');
    sheet.getRange('A23').setValue('最終重量(g):');
    sheet.getRange('C23').setValue('重量差(g):');
    
    // 5. DDP対応・地域別設定セクション
    sheet.getRange('A26').setValue('【DDP対応・地域別設定】');
    sheet.getRange('A26').setFontWeight('bold').setBackground('#E8F4FD');
    sheet.getRange('A27').setValue('販売ターゲット地域:');
    sheet.getRange('A28').setValue('推奨配送方法:');
    sheet.getRange('A29').setValue('利用可能配送方法:');
    sheet.getRange('A30').setValue('DDP対応警告:');
    
    // 6. 為替レート管理セクション
    sheet.getRange('A31').setValue('【為替レート管理】');
    sheet.getRange('A31').setFontWeight('bold').setBackground('#E8F4FD');
    sheet.getRange('A32').setValue('手動設定:');
    sheet.getRange('A33').setValue('手動為替レート:');
    sheet.getRange('A34').setValue('自動為替レート:');
    sheet.getRange('A35').setValue('最終為替レート:');
  sheet.getRange('A36').setValue('USD換算販売額:');
    
    console.log('利益計算シートのヘッダー設定が完了しました');
    
  } catch (error) {
    console.error('ヘッダー設定中にエラーが発生しました:', error);
  }
}

/**
 * 利益計算シートのセル書式設定
 */
function setupProfitSheetFormatting(sheet) {
  try {
    // 1. 通貨フォーマット（円）
    const jpyCurrencyCells = ['B8', 'B10', 'B12', 'B13', 'B14', 'E14', 'E16', 'E17', 'E18', 'E19', 'B18', 'B33', 'B34', 'B35'];
    jpyCurrencyCells.forEach(cell => {
      sheet.getRange(cell).setNumberFormat('¥#,##0');
    });
    sheet.getRange(PROFIT_CELLS.USD_CONVERTED_PRICE).setNumberFormat('$#,##0.00');
    
    // 2. パーセントフォーマット
    const percentCells = ['B9', 'B11', 'B17', 'D13'];
    percentCells.forEach(cell => {
      sheet.getRange(cell).setNumberFormat('0.0%');
    });
    
    // 3. 数値フォーマット（整数）
    const integerCells = ['D12', 'B15'];
    integerCells.forEach(cell => {
      sheet.getRange(cell).setNumberFormat('0');
    });
    
    // 4. 数値フォーマット（小数点表示）
    const decimalCells = ['B21', 'D21', 'B22', 'D22', 'B23', 'D23'];
    decimalCells.forEach(cell => {
      sheet.getRange(cell).setNumberFormat('0.0');
    });
    
    // 5. 計算結果セルの黄色ハイライト
    const resultCells = ['B8', 'B9', 'B10', 'B11', 'E14', 'E16', 'E17', 'E18', 'E19', 'D22', 'B23', 'D23'];
    resultCells.forEach(cell => {
      sheet.getRange(cell).setBackground('#FFFF99'); // 薄い黄色
    });
    
    // 5. 入力セルの薄い青色
    const inputCells = ['B2', 'B5', 'E5', 'B6', 'E6', 'B12', 'D12', 'B13', 'D13', 'B14', 'B15', 'E15', 'B16', 'B18', 'B19', 'B21', 'D21', 'B22', 'B32', 'B33'];
    inputCells.forEach(cell => {
      sheet.getRange(cell).setBackground('#E6F3FF'); // 薄い青色
    });
    
    // 6. 設定値セルの薄い緑色
    const settingCells = ['B17', 'B27', 'B34', 'B35', 'B36']; // B27はB16から自動取得
    settingCells.forEach(cell => {
      sheet.getRange(cell).setBackground('#E6FFE6'); // 薄い緑色
    });
    
    // 7. 警告・注意セルの薄いオレンジ色
    const warningCells = ['B28', 'B29', 'B30'];
    warningCells.forEach(cell => {
      sheet.getRange(cell).setBackground('#FFF2E6'); // 薄いオレンジ色
    });
    
    console.log('利益計算シートの書式設定が完了しました');
    
  } catch (error) {
    console.error('書式設定中にエラーが発生しました:', error);
  }
}

/**
 * 利益計算シートのレイアウト調整
 */
function setupProfitSheetLayout(sheet) {
  try {
    // 1. 列幅の設定
    sheet.setColumnWidth(1, 15);   // A列: 15px
    sheet.setColumnWidth(2, 120);  // B列: 120px
    sheet.setColumnWidth(3, 80);   // C列: 80px
    sheet.setColumnWidth(4, 100);  // D列: 100px
    sheet.setColumnWidth(5, 120);  // E列: 120px
    sheet.setColumnWidth(6, 80);   // F列: 80px
    
    // G-Z列の列幅を60pxに設定
    for (let col = 7; col <= 26; col++) {
      sheet.setColumnWidth(col, 60);
    }
    
    // 2. 行高の設定
    sheet.setRowHeight(1, 30);   // ヘッダー行: 30px
    sheet.setRowHeight(2, 25);   // データ行: 25px
    sheet.setRowHeight(3, 25);   // データ行: 25px
    
    // セクション見出し行の行高を30pxに設定
    const sectionRows = [4, 7, 20, 26, 31];
    sectionRows.forEach(row => {
      sheet.setRowHeight(row, 30);
    });
    
    // 計算結果行の行高を30pxに設定
    const resultRows = [8, 9, 10, 11, 14, 16, 17, 18, 21, 22, 32, 33, 34, 35, 36, 37];
    resultRows.forEach(row => {
      sheet.setRowHeight(row, 30);
    });
    
    // 3. セクション間の区切り線
    const sectionBorders = [
      { range: 'A3:Z3', style: 'SOLID_MEDIUM' },  // ヘッダー下
      { range: 'A6:Z6', style: 'SOLID_THIN' },    // 出品情報下
      { range: 'A19:Z19', style: 'SOLID_THIN' },  // 利益計算下
      { range: 'A24:Z24', style: 'SOLID_THIN' },  // 容積重量チェック下（行23の下、行24）
      { range: 'A30:Z30', style: 'SOLID_THIN' },  // DDP対応下（行30の下、行31の上）
      { range: 'A37:Z37', style: 'SOLID_THIN' }   // 為替レート・諸費用設定下
    ];
    
    sectionBorders.forEach(border => {
      sheet.getRange(border.range).setBorder(
        true, true, true, true, true, true, border.style, SpreadsheetApp.BorderStyle.SOLID
      );
    });
    
    // 4. セルの結合（タイトル行）
    sheet.getRange('A1:F1').merge();
    sheet.getRange('A1').setHorizontalAlignment('center');
    
    // 5. セクション見出しの結合
    const sectionHeaders = [
      { range: 'A4:F4', text: '【出品情報】' },
      { range: 'A7:F7', text: '【利益計算】' },
      { range: 'A20:F20', text: '【容積重量チェック】' },
      { range: 'A26:F26', text: '【DDP対応・地域別設定】' },
      { range: 'A31:F31', text: '【為替レート管理】' },
    ];
    
    sectionHeaders.forEach(header => {
      sheet.getRange(header.range).merge();
      sheet.getRange(header.range).setHorizontalAlignment('center');
    });
    
    console.log('利益計算シートのレイアウト調整が完了しました');
    
  } catch (error) {
    console.error('レイアウト調整中にエラーが発生しました:', error);
  }
}

/**
 * 利益計算シートのデータ検証設定
 */
function setupProfitSheetDataValidation(sheet) {
  try {
    // 発送方法候補（簡易リスト）
    const shippingList = ['Joom Logistics', 'eLogi DDP', 'DHL', 'FedEx', 'UPS', '西濃運輸', '日本通運'];
    const shippingValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(shippingList, true)
      .setAllowInvalid(false)
      .setHelpText('配送業者を選択してください')
      .build();
    sheet.getRange('E15').setDataValidation(shippingValidation);

    // 配送地帯候補（簡易リスト）
    const regionList = ['EU圏', 'ロシア', '東欧', '北米', 'アジア', 'オセアニア', '南米'];
    const regionValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(regionList, true)
      .setAllowInvalid(false)
      .setHelpText('配送先地域を選択してください')
      .build();
    sheet.getRange('B16').setDataValidation(regionValidation);
    // B27はB16から自動取得するため、データ検証は不要
    
    // カテゴリー候補（簡易）
    const categoryList = ['一般', '家電', 'アパレル', '雑貨', 'ホビー'];
    const categoryValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(categoryList, true)
      .setAllowInvalid(false)
      .setHelpText('商品カテゴリーを選択してください')
      .build();
    sheet.getRange('B19').setDataValidation(categoryValidation);
    
    // 数値入力の制限
    // 販売価格（0以上）
    const priceValidation = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0)
      .setAllowInvalid(false)
      .setHelpText('0以上の数値を入力してください')
      .build();
    sheet.getRange('B12').setDataValidation(priceValidation);
    sheet.getRange('B13').setDataValidation(priceValidation);
    
    // 出品数量（1以上）
    const quantityValidation = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(1)
      .setAllowInvalid(false)
      .setHelpText('1以上の整数を入力してください')
      .build();
    sheet.getRange('D12').setDataValidation(quantityValidation);
    
    // 割引率（0-100%）
    const discountValidation = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 100)
      .setAllowInvalid(false)
      .setHelpText('0-100の数値を入力してください')
      .build();
    sheet.getRange('D13').setDataValidation(discountValidation);
    
    // 重量（0以上）
    const weightValidation = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0)
      .setAllowInvalid(false)
      .setHelpText('0以上の数値を入力してください')
      .build();
    sheet.getRange('B15').setDataValidation(weightValidation);
    sheet.getRange('B21').setDataValidation(weightValidation);
    sheet.getRange('D21').setDataValidation(weightValidation);
    sheet.getRange('B22').setDataValidation(weightValidation);
    
    // 返金額（0以上）
    sheet.getRange('B18').setDataValidation(priceValidation);
    
    // 為替レート（0以上）
    const exchangeValidation = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0)
      .setAllowInvalid(false)
      .setHelpText('0より大きい数値を入力してください')
      .build();
    sheet.getRange('B33').setDataValidation(exchangeValidation);
    
    // 商品ID入力のデータ検証（在庫管理シートのA列を参照）
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const inventorySheet = ss.getSheetByName('在庫管理');
      if (inventorySheet) {
        const productIdValidation = SpreadsheetApp.newDataValidation()
          .requireValueInRange(inventorySheet.getRange('A:A'))
          .setAllowInvalid(false)
          .setHelpText('在庫管理シートの商品IDを選択してください')
          .build();
        sheet.getRange('B2').setDataValidation(productIdValidation);
      }
    } catch (e) {
      console.log('商品IDデータ検証の設定中に警告:', e);
    }
    
    console.log('利益計算シートのデータ検証設定が完了しました');
    
  } catch (error) {
    console.error('データ検証設定中にエラーが発生しました:', error);
  }
}

/**
 * 利益計算シートのセル保護設定
 */
function setupProfitSheetProtection(sheet) {
  try {
    // 1. 計算結果セルを保護（読み取り専用）
    const protectedCells = [
      'B8', 'B9', 'B10', 'B11',  // 利益計算結果
      'E14', 'E16', 'E17', 'E18', 'E19',       // 原価・送料・手数料・サーチャージ・繁忙期料金
      'D22', 'B23', 'D23',       // 容積重量・最終重量・重量差
      'B17', 'B27', 'B34', 'B35', 'B36' // 設定値・自動取得値（B27はB16から自動取得）
    ];
    
    protectedCells.forEach(cell => {
      sheet.getRange(cell).protect();
    });
    
    // 2. 入力可能セルを編集可能に設定
    const editableCells = [
      'B2',                       // 商品ID
      'B5', 'E5', 'B6', 'E6',     // 商品情報
      'B12', 'D12', 'B13', 'D13', 'B14', 'B15', 'E15', 'B16', 'B18', // 価格・数量・重量・配送
      'B21', 'D21', 'B22',        // 寸法
      'B19',                      // カテゴリー
      'B32', 'B33'                // 為替レート設定
    ];
    
    editableCells.forEach(cell => {
      sheet.getRange(cell).setNote('編集可能セル');
    });
    
    // 3. セクション見出しとヘッダーを保護
    const headerCells = [
      'A1', 'A4', 'A7', 'A20', 'A26', 'A31' // セクション見出し
    ];
    
    headerCells.forEach(cell => {
      sheet.getRange(cell).protect();
    });
    
    console.log('利益計算シートのセル保護設定が完了しました');
    
  } catch (error) {
    console.error('セル保護設定中にエラーが発生しました:', error);
  }
}

/**
 * 利益計算シートの条件付き書式設定
 */
function setupProfitSheetConditionalFormatting(sheet) {
  try {
    // 1. 利益率による色分け（B9, B11）
    const profitRateRange = sheet.getRange('B9:B11');
    const profitRateRules = [
      // 利益率 20%以上: 緑色
      SpreadsheetApp.newConditionalFormatRule()
        .setRanges([profitRateRange])
        .whenNumberGreaterThanOrEqualTo(0.2)
        .setBackground('#90EE90') // 薄い緑色
        .build(),
      // 利益率 10-20%: 黄色
      SpreadsheetApp.newConditionalFormatRule()
        .setRanges([profitRateRange])
        .whenNumberBetween(0.1, 0.2)
        .setBackground('#FFFF99') // 黄色
        .build(),
      // 利益率 0-10%: オレンジ色
      SpreadsheetApp.newConditionalFormatRule()
        .setRanges([profitRateRange])
        .whenNumberBetween(0, 0.1)
        .setBackground('#FFB366') // オレンジ色
        .build(),
      // 利益率 0%未満: 赤色
      SpreadsheetApp.newConditionalFormatRule()
        .setRanges([profitRateRange])
        .whenNumberLessThan(0)
        .setBackground('#FF6B6B') // 赤色
        .build()
    ];
    
    // 2. 利益額による色分け（B8, B10）
    const profitAmountRange = sheet.getRange('B8:B10');
    const profitAmountRules = [
      // 利益額 1000円以上: 緑色
      SpreadsheetApp.newConditionalFormatRule()
        .setRanges([profitAmountRange])
        .whenNumberGreaterThanOrEqualTo(1000)
        .setBackground('#90EE90') // 薄い緑色
        .build(),
      // 利益額 0-1000円: 黄色
      SpreadsheetApp.newConditionalFormatRule()
        .setRanges([profitAmountRange])
        .whenNumberBetween(0, 1000)
        .setBackground('#FFFF99') // 黄色
        .build(),
      // 利益額 0円未満: 赤色
      SpreadsheetApp.newConditionalFormatRule()
        .setRanges([profitAmountRange])
        .whenNumberLessThan(0)
        .setBackground('#FF6B6B') // 赤色
        .build()
    ];
    
    // 3. エラー値の検出（#ERROR!, #N/A等）
    const errorRange = sheet.getRange('B8:B40');
    const errorRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([errorRange])
      .whenFormulaSatisfies('=ISERROR(B8)')
      .setBackground('#FF0000') // 赤色
      .setFontColor('#FFFFFF') // 白色文字
      .build();
    
    // 4. 条件付き書式ルールを適用
    const allRules = [...profitRateRules, ...profitAmountRules, errorRule];
    sheet.setConditionalFormatRules(allRules);
    
    console.log('利益計算シートの条件付き書式設定が完了しました');
    
  } catch (error) {
    console.error('条件付き書式設定中にエラーが発生しました:', error);
  }
}

/**
 * 商品IDに基づいて在庫管理シートからデータを自動設定
 */
function loadProductDataFromInventory(productId, showFeedback = false) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inventorySheet = ss.getSheetByName('在庫管理');
    const profitSheet = ss.getSheetByName(SHEET_NAMES.PROFIT);
    
    if (!inventorySheet || !profitSheet) {
      const message = '在庫管理シートまたは利益計算シートが見つかりません';
      console.log(message);
      if (showFeedback) {
        SpreadsheetApp.getUi().alert('エラー', message, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return false;
    }
    
    if (!productId || productId === '') {
      const message = '商品IDが指定されていません';
      console.log(message);
      if (showFeedback) {
        SpreadsheetApp.getUi().alert('エラー', message, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return false;
    }
    
    // 在庫管理シートから商品データを検索
    const dataRange = inventorySheet.getDataRange();
    const values = dataRange.getValues();
    
    let productRow = -1;
    let productData = null;
    
    for (let i = 1; i < values.length; i++) { // ヘッダー行をスキップ
      if (values[i][0] === productId) { // A列（商品ID）
        productRow = i + 1; // 1ベースの行番号
        productData = values[i]; // 0ベースのインデックス
        break;
      }
    }
    
    if (productRow === -1) {
      const message = `商品ID "${productId}" が見つかりません。在庫管理シートに登録されている商品IDを入力してください。`;
      console.log(message);
      if (showFeedback) {
        SpreadsheetApp.getUi().alert('商品が見つかりません', message, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return false;
    }
    
    // データ検証
    const validationResult = validateProductData(productData);
    if (!validationResult.isValid) {
      const message = `商品データに問題があります: ${validationResult.errors.join(', ')}`;
      console.log(message);
      if (showFeedback) {
        SpreadsheetApp.getUi().alert('データ検証エラー', message, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return false;
    }
    
    // データを利益計算シートに設定（プログラム更新フラグを設定）
    if (typeof beginProgrammaticUpdate === 'function') {
      beginProgrammaticUpdate();
    }
    
    try {
      // 在庫管理シートの列構造に基づいてデータを設定
      // A: 商品ID, B: 商品名, C: SKU, D: ASIN, E: 仕入れ元, F: 仕入れ元URL, G: 仕入れ価格, I: 重量
      profitSheet.getRange('B5').setValue(productData[2] || '');  // SKU
      profitSheet.getRange('E5').setValue(productData[1] || '');  // 商品名
      profitSheet.getRange('B6').setValue(productData[5] || '');  // 仕入れ元URL
      profitSheet.getRange('E6').setValue(productData[4] || '');  // 仕入れ元
      profitSheet.getRange('B13').setValue(productData[6] || 0);  // 仕入れ価格
      profitSheet.getRange('B15').setValue(productData[8] || 0);  // 重量
      
      // 容積重量計算用寸法データの設定
      // U: 高さ(cm), V: 長さ(cm), W: 幅(cm), X: 容積重量係数
      profitSheet.getRange('B21').setValue(productData[20] || 0);  // 高さ(cm)
      profitSheet.getRange('D21').setValue(productData[21] || 0);  // 長さ(cm)
      profitSheet.getRange('B22').setValue(productData[22] || 0);  // 幅(cm)
      
      // 商品カテゴリーの設定
      // Y: 商品カテゴリー（25列目、インデックス24）
      profitSheet.getRange(PROFIT_CELLS.CATEGORY).setValue(productData[24] || '');  // 商品カテゴリー
      
      const message = `商品ID "${productId}" のデータを正常に読み込みました`;
      console.log(message);
      if (showFeedback) {
        SpreadsheetApp.getUi().alert('データ読み込み完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return true;
      
    } finally {
      if (typeof endProgrammaticUpdate === 'function') {
        endProgrammaticUpdate();
      }
    }
    
  } catch (error) {
    const message = `商品データの読み込み中にエラーが発生しました: ${error.message}`;
    console.error(message, error);
    if (showFeedback) {
      SpreadsheetApp.getUi().alert('エラー', message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * 商品データの検証
 */
function validateProductData(productData) {
  const errors = [];
  
  try {
    // 必須フィールドの検証
    if (!productData[1] || productData[1] === '') {
      errors.push('商品名が設定されていません');
    }
    
    if (!productData[6] || productData[6] <= 0) {
      errors.push('仕入価格が正しく設定されていません');
    }
    
    if (!productData[8] || productData[8] <= 0) {
      errors.push('重量が正しく設定されていません');
    }
    
    // 寸法データの検証（オプション）
    if (productData[20] && productData[20] < 0) {
      errors.push('高さが負の値です');
    }
    
    if (productData[21] && productData[21] < 0) {
      errors.push('長さが負の値です');
    }
    
    if (productData[22] && productData[22] < 0) {
      errors.push('幅が負の値です');
    }
    
    // 容積重量係数の検証
    if (productData[23] && (productData[23] < 1000 || productData[23] > 10000)) {
      errors.push('容積重量係数が範囲外です（1000-10000）');
    }
    
    return {
      isValid: errors.length === 0,
      errors: errors
    };
    
  } catch (error) {
    return {
      isValid: false,
      errors: ['データ検証中にエラーが発生しました: ' + error.message]
    };
  }
}

/**
 * 利益計算シートにテストデータを設定
 */
function setupProfitSheetTestData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const profitSheet = ss.getSheetByName(SHEET_NAMES.PROFIT);
    
    if (!profitSheet) {
      SpreadsheetApp.getUi().alert('利益計算シートが見つかりません');
      return;
    }
    
    // テストデータを設定
    profitSheet.getRange('B2').setValue('1'); // 商品ID（iPhone 15 Pro）
    profitSheet.getRange('B5').setValue('IPH15P-128'); // SKU
    profitSheet.getRange('E5').setValue('iPhone 15 Pro 128GB'); // 商品名
    profitSheet.getRange('B6').setValue('https://amazon.co.jp/dp/B0CHX1W1XY'); // 仕入URL
    profitSheet.getRange('E6').setValue('Amazon'); // 仕入元
    profitSheet.getRange('B12').setValue(150000); // 販売価格
    profitSheet.getRange('D12').setValue(1); // 出品数量
    profitSheet.getRange('B13').setValue(120000); // 仕入価格
    profitSheet.getRange('D13').setValue(0); // 割引率
    profitSheet.getRange('B14').setValue(0); // 割引ポイント
    profitSheet.getRange('B15').setValue(187); // 重量
    profitSheet.getRange('E15').setValue('Joom Logistics'); // 発送方法
    profitSheet.getRange('B16').setValue('EU圏'); // 配送地帯（B27も自動で同期）
    profitSheet.getRange('B18').setValue(0); // 返金額
    profitSheet.getRange('B21').setValue(7.6); // 高さ(cm)
    profitSheet.getRange('D21').setValue(14.7); // 長さ(cm)
    profitSheet.getRange('B22').setValue(0.8); // 幅(cm)
    profitSheet.getRange(PROFIT_CELLS.CATEGORY).setValue('家電'); // 商品カテゴリー
    profitSheet.getRange(PROFIT_CELLS.EXCHANGE_RATE_MANUAL_VALUE).setValue(150); // 手動為替レート
    
    SpreadsheetApp.getUi().alert('利益計算シートにテストデータを設定しました');
    
  } catch (error) {
    console.error('テストデータ設定中にエラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
  }
}

/**
 * 複数商品の一括データ読み込み
 */
function loadMultipleProductsData(productIds) {
  try {
    const results = [];
    let successCount = 0;
    let errorCount = 0;
    
    for (const productId of productIds) {
      const result = loadProductDataFromInventory(productId, false);
      if (result) {
        successCount++;
        results.push({ productId, status: 'success' });
      } else {
        errorCount++;
        results.push({ productId, status: 'error' });
      }
    }
    
    const message = `一括読み込み完了: 成功 ${successCount}件, エラー ${errorCount}件`;
    console.log(message);
    SpreadsheetApp.getUi().alert('一括読み込み結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
    return results;
    
  } catch (error) {
    const message = `一括読み込み中にエラーが発生しました: ${error.message}`;
    console.error(message, error);
    SpreadsheetApp.getUi().alert('エラー', message, SpreadsheetApp.getUi().ButtonSet.OK);
    return [];
  }
}

/**
 * 複数商品一括読み込みメニューの表示
 */
function showBulkLoadMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      '複数商品一括読み込み',
      '読み込む商品IDをカンマ区切りで入力してください（例: 1,2,3,4,5）:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const input = response.getResponseText().trim();
      if (input === '') {
        ui.alert('エラー', '商品IDを入力してください', ui.ButtonSet.OK);
        return;
      }
      
      // カンマ区切りで分割して、空白を除去
      const productIds = input.split(',').map(id => id.trim()).filter(id => id !== '');
      
      if (productIds.length === 0) {
        ui.alert('エラー', '有効な商品IDが入力されていません', ui.ButtonSet.OK);
        return;
      }
      
      // 一括読み込み実行
      loadMultipleProductsData(productIds);
    }
    
  } catch (error) {
    console.error('一括読み込みメニューでエラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラー', '一括読み込みメニューでエラーが発生しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 商品ID入力時の自動データ読み込み（メニュー用）
 */
function loadProductDataMenu() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const profitSheet = ss.getSheetByName(SHEET_NAMES.PROFIT);
    
    if (!profitSheet) {
      SpreadsheetApp.getUi().alert('利益計算シートが見つかりません');
      return;
    }
    
    const productId = profitSheet.getRange('B2').getValue();
    
    if (!productId || productId === '') {
      SpreadsheetApp.getUi().alert('B2セルに商品IDを入力してください');
      return;
    }
    
    const success = loadProductDataFromInventory(productId);
    
    if (success) {
      SpreadsheetApp.getUi().alert(`商品ID "${productId}" のデータを読み込みました`);
    } else {
      SpreadsheetApp.getUi().alert(`商品ID "${productId}" のデータが見つかりませんでした`);
    }
    
  } catch (error) {
    console.error('商品データ読み込みメニューでエラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
  }
}

/**
 * 利益計算シートの参照式を更新（マスタシート作成後）
 */
function updateProfitSheetFormulas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.PROFIT);
    if (!sheet) {
      console.log('利益計算シートが見つかりません');
      return false;
    }
    
    // マスタシートの存在確認と作成（必要に応じて）
    createMasterSheetsIfNeeded(ss);
    
    // 名前付き範囲の設定（参照式設定の前に必要）
    // マスタ参照の名前付き範囲（存在時のみ設定）
    try {
      const shippingMaster = ss.getSheetByName('Joom送料マスタ');
      if (shippingMaster) {
        // 例: 配送方法名B列、容積重量係数N列
        // 注意: G-M列（地域別送料）は削除済み。送料計算には重量別送料マスタを使用
        ss.setNamedRange(NAMED_RANGES.SHIPPING_METHODS, shippingMaster.getRange('B2:B'));
        ss.setNamedRange(NAMED_RANGES.SHIPPING_SURCHARGE, shippingMaster.getRange('N2:N'));
        ss.setNamedRange(NAMED_RANGES.SHIPPING_VOLUME_FACTOR, shippingMaster.getRange('O2:O'));
      }
      const dutyMaster = ss.getSheetByName('関税率マスタ');
      if (dutyMaster) {
        ss.setNamedRange('関税率マスタ_カテゴリー', dutyMaster.getRange('B2:B'));
        ss.setNamedRange('関税率マスタ_税率', dutyMaster.getRange('D2:J'));
      }
      const peakSeasonMaster = ss.getSheetByName(SHEET_NAMES.PEAK_SEASON_MASTER);
      if (peakSeasonMaster) {
        ss.setNamedRange(NAMED_RANGES.PEAK_SEASON_METHODS, peakSeasonMaster.getRange('A2:A'));
        ss.setNamedRange(NAMED_RANGES.PEAK_SEASON_REGIONS, peakSeasonMaster.getRange('B2:B'));
        ss.setNamedRange(NAMED_RANGES.PEAK_SEASON_AMOUNT, peakSeasonMaster.getRange('C2:C'));
      }
      console.log('名前付き範囲を設定しました');
    } catch (e) {
      console.log('名前付き範囲の設定中に警告:', e);
    }
    
    // 参照式を再設定
    setupProfitSheetFormulas(sheet);
    
    console.log('利益計算シートの参照式を更新しました');
    return true;
    
  } catch (error) {
    console.error('参照式更新中にエラーが発生しました:', error);
    return false;
  }
}

/**
 * 利益計算シートのみの初期化（メニュー用）
 */
function initializeProfitSheetOnly() {
  console.log('利益計算シートの初期化を開始します...');
  
  try {
    // 利益計算シートの初期化
    initializeProfitSheet();
    
    // 検証を実行
    const isValid = verifyProfitSheetLayout();
    
    // 設定値の更新
    if (typeof updateProfitCalculationSettings === 'function') {
      updateProfitCalculationSettings();
    }
    
    // 為替レートの検証
    if (typeof validateAndUpdateExchangeRate === 'function') {
      validateAndUpdateExchangeRate();
    }
    
    const resultMessage = '利益計算シートの初期化が完了しました。\n\n' +
                         '初期化内容:\n' +
                         '✓ シート作成・設定\n' +
                         '✓ 名前付き範囲設定\n' +
                         '✓ データ検証（ドロップダウン）設定\n' +
                         '✓ 参照式設定\n' +
                         '✓ 設定値反映\n' +
                         '✓ 為替レート検証\n\n' +
                         '検証結果: ' + (isValid ? '正常' : '一部警告あり') + '\n\n' +
                         '利益計算機能が利用可能になりました。';
    
    SpreadsheetApp.getUi().alert('利益計算シート初期化完了', resultMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('利益計算シート初期化中にエラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラー', '利益計算シート初期化中にエラーが発生しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 利益計算シートの参照式設定
 * 仕様参照: doc/Joom利益計算機能仕様書.md 3.1.0.3
 */
function setupProfitSheetFormulas(sheet) {
  try {
    // B17を一旦保護解除して参照式を設定
    const b17Range = sheet.getRange('B17');
    let wasProtected = false;
    try {
      if (b17Range.getProtections().length > 0) {
        wasProtected = true;
        b17Range.getProtections()[0].remove();
      }
    } catch (e) {
      // 保護が解除できない場合はスキップ
    }
    
    // 為替レート参照（手動/自動切替対応）
    // B32: 為替レート手動設定フラグ
    // B33: 手動為替レート値
    // B34: 自動為替レート値（GoogleFinance or 設定シート）
    // B35: 最終為替レート（IF文で手動優先）
    
    // 手動設定フラグのデフォルト値
    sheet.getRange(PROFIT_CELLS.EXCHANGE_RATE_MANUAL_TOGGLE).setValue(false);
    const manualToggleValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['TRUE', 'FALSE'], true)
      .setAllowInvalid(false)
      .setHelpText('TRUE または FALSE を選択してください')
      .build();
    sheet.getRange(PROFIT_CELLS.EXCHANGE_RATE_MANUAL_TOGGLE).setDataValidation(manualToggleValidation);
    
    // 手動為替レートのデフォルト値
    sheet.getRange(PROFIT_CELLS.EXCHANGE_RATE_MANUAL_VALUE).setValue(150);
    
    // 自動為替レート（GoogleFinance参照）
    sheet.getRange(PROFIT_CELLS.EXCHANGE_RATE_AUTO_VALUE).setFormula('=GOOGLEFINANCE("CURRENCY:USDJPY")');
    
    // 最終為替レート（手動優先）
    sheet.getRange(PROFIT_CELLS.EXCHANGE_RATE_FINAL_VALUE).setFormula(
      '=IF(' + PROFIT_CELLS.EXCHANGE_RATE_MANUAL_TOGGLE + '=TRUE,' +
      PROFIT_CELLS.EXCHANGE_RATE_MANUAL_VALUE + ',' +
      PROFIT_CELLS.EXCHANGE_RATE_AUTO_VALUE + ')'
    );
    sheet.getRange(PROFIT_CELLS.USD_CONVERTED_PRICE).setFormula(
      '=IF(AND(' + PROFIT_CELLS.EXCHANGE_RATE_FINAL_VALUE + '>0,B12<>""),B12/' + PROFIT_CELLS.EXCHANGE_RATE_FINAL_VALUE + ',0)'
    );
    
    // 送料計算の参照式は、B23（最終重量）の設定後に配置する（依存関係のため）
    
    // B27（販売ターゲット地域）をB16（配送地帯）から自動取得（統一）
    sheet.getRange('B27').setFormula('=B16');
    sheet.getRange('B27').setBackground('#E6FFE6'); // 設定値セルの薄い緑色（自動取得のため）
    
    // 容積重量計算式（送料マスタから容積重量係数を取得）
    // D22: 容積重量（cm^3 / 容積係数）
    sheet.getRange(PROFIT_CELLS.VOLUMETRIC_WEIGHT).setFormula('=IF(AND(B21>0,D21>0,B22>0),ROUND((B21*D21*B22)/IFERROR(INDEX(送料マスタ_容積重量係数,MATCH(E15,送料マスタ_配送方法,0)),6000)*1000,0),0)');
    
    // Joom手数料率の参照式（Joom手数料マスタから取得）
    // B17: Joom手数料率（カテゴリーB19と配送地帯B16からJoom手数料マスタを参照）
    // Joom手数料マスタの構造:
    //   - B列=カテゴリー名
    //   - D列=EU圏、E列=ロシア、F列=東欧、G列=北米、H列=アジア、I列=オセアニア、J列=南米
    // MATCHでカテゴリー行を特定し、配送地帯に応じた列をINDEXで取得
    // パーセント表示に合わせて小数(例: 12.7% -> 0.127)で保持
    // フォールバック: デフォルト値12.7%（0.127）
    sheet.getRange('B17').setFormula('=IFERROR(INDEX(Joom手数料マスタ!D:J,MATCH(B19,Joom手数料マスタ!B:B,0),MATCH(B16,Joom手数料マスタ!D1:J1,0))/100,0.127)');
    
    // 最終重量（実重量と容積重量の大きい方）
    // B23: 最終重量
    sheet.getRange('B23').setFormula('=MAX(B15,D22)');
    
    // 重量差（実重量と容積重量の差）
    // D23: 重量差
    sheet.getRange('D23').setFormula('=ABS(B15-D22)');
    
    // 送料計算の参照式（重量別送料マスタ参照、フォールバック付き）
    // E16: 送料（配送方法・最終重量・地域を考慮して計算）
    // 主に重量別送料マスタを使用し、フォールバックとしてデフォルト値を参照
    // 注意: B23（最終重量）の設定後に配置（依存関係のため）
    // B23がエラーの場合はB15（実重量）をフォールバックとして使用
    const weightShippingFormula = '=IFERROR(getWeightShippingCost(E15,IFERROR(B23,B15),B16),' +
      'IF(AND(E15="Joom Logistics",OR(B16="EU圏",B16="ロシア",B16="東欧")),1200,' +
      'IF(AND(E15="eLogi DDP",OR(B16="EU圏",B16="ロシア",B16="東欧")),1500,' +
      'IF(AND(E15="DHL",OR(B16="EU圏",B16="ロシア",B16="東欧")),1800,' +  // DHL + EU圏対応
      'IF(AND(E15="DHL",B16="北米"),2000,' +
      'IF(AND(E15="DHL",B16="アジア"),1600,' +  // DHL + アジア対応
      'IF(AND(E15="FedEx",B16="北米"),1800,' +
      'IF(AND(E15="FedEx",OR(B16="EU圏",B16="ロシア",B16="東欧")),1700,1000)))))))';  // FedEx + EU圏対応
    sheet.getRange('E16').setFormula(weightShippingFormula);
    sheet.getRange(PROFIT_CELLS.SHIPPING_SURCHARGE).setFormula(
      '=IF(E15="","",IFERROR(INDEX(' + NAMED_RANGES.SHIPPING_SURCHARGE + ',MATCH(E15,' + NAMED_RANGES.SHIPPING_METHODS + ',0)),0))'
    );
    sheet.getRange(PROFIT_CELLS.PEAK_SEASON_FEE).setFormula(
      '=IF(OR(E15="",' + PROFIT_CELLS.REGION + '=""),0,IFERROR(getPeakSeasonFee(E15,' + PROFIT_CELLS.REGION + '),0))'
    );
    
    // 利益計算の基本式
    // B8: 利益額 = 販売価格 - 仕入原価 - 送料 - Joom手数料 + 返金額
    sheet.getRange('B8').setFormula('=B12-E14-E16-E17-E18-E19+B18');
    
    // B9: 利益率 = 利益額 / 販売価格
    sheet.getRange('B9').setFormula('=IF(B12>0,B8/B12,0)');
    
    // B10: 還付込利益額 = 利益額 + 返金額
    sheet.getRange('B10').setFormula('=B8+B18');
    
    // B11: 還付込利益率 = 還付込利益額 / 販売価格
    sheet.getRange('B11').setFormula('=IF(B12>0,B10/B12,0)');
    
    // E14: 仕入原価 = 仕入価格 - 割引ポイント
    sheet.getRange('E14').setFormula('=B13-B14');
    
    // E17: Joom手数料 = 販売価格 * Joom手数料率
    // B17は既に/100で小数（0.127）になっているので、そのまま使用
    // B17がエラーや未定義の場合は0として扱う
    sheet.getRange('E17').setFormula('=ROUND(B12*IFERROR(B17,0.127),0)');
    
    console.log('利益計算シートの参照式設定が完了しました');
    
    // B17を再保護（元々保護されていた場合）
    if (wasProtected) {
      try {
        b17Range.protect();
      } catch (e) {
        console.log('B17の再保護でエラー:', e);
      }
    }
    
  } catch (error) {
    console.error('参照式設定中にエラーが発生しました:', error);
  }
}

/**
 * 利益計算シートの主要セル配置の簡易検証（ログ出力）
 */
function verifyProfitSheetLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PROFIT);
  if (!sheet) {
    console.log('利益計算シートが存在しません');
    return false;
  }
  
  // 主要セルの存在確認と値の表示
  const keyCells = [
    { cell: 'B8', name: '利益（円）' },
    { cell: 'B9', name: '利益率（%）' },
    { cell: 'B10', name: '利益（返金込み）' },
    { cell: 'B11', name: '利益率（返金込み）' },
    { cell: 'B12', name: '販売価格' },
    { cell: 'B13', name: '仕入価格' },
    { cell: 'B15', name: '重量（g）' },
    { cell: 'E15', name: '発送方法' },
    { cell: 'B16', name: '配送地帯' },
    { cell: 'B27', name: '販売ターゲット地域（B16から自動取得）' },
    { cell: 'B19', name: '商品カテゴリー' }
  ];
  
  let ok = true;
  console.log('=== 利益計算シート セル配置検証 ===');
  
  for (let i = 0; i < keyCells.length; i++) {
    const { cell, name } = keyCells[i];
    try {
      const range = sheet.getRange(cell);
      const value = range.getValue();
      console.log(`${cell} (${name}): ${value || '(空)'}`);
    } catch (e) {
      ok = false;
      console.log(`セル参照エラー: ${cell} (${name}) - ${e.message}`);
    }
  }
  
  // 名前付き範囲の確認
  console.log('=== 名前付き範囲確認 ===');
  const namedRanges = [
    '為替レート_USDJPY',
    '送料マスタ_配送方法',
    '送料マスタ_容積重量係数',
    '関税率マスタ_カテゴリー',
    '関税率マスタ_税率',
  ];
  
  for (let i = 0; i < namedRanges.length; i++) {
    const rangeName = namedRanges[i];
    try {
      const range = ss.getRangeByName(rangeName);
      if (range) {
        console.log(`${rangeName}: ${range.getA1Notation()} (設定済み)`);
      } else {
        console.log(`${rangeName}: 未設定`);
      }
    } catch (e) {
      console.log(`${rangeName}: エラー - ${e.message}`);
    }
  }
  
  console.log('利益計算シート配置検証結果:', ok ? 'OK' : '一部NG');
  return ok;
}

/**
 * カスタムメニューの設定（Joom注文連携対応版）
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
        .addItem('Joom注文同期', 'showJoomOrderSyncMenu')
        .addItem('注文ステータス管理', 'showOrderStatusManagement')
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
      ui.createMenu('🛠️ システム管理')
        .addItem('全シート初期化', 'initializeAllSheets')
        .addItem('利益計算シート初期化', 'initializeProfitSheetOnly')
        .addItem('データバックアップ', 'showDataBackupMenu')
    )
    .addSubMenu(
      ui.createMenu('⚙️ 設定')
        .addItem('設定シート初期化', 'initializeSettingsSheet')
        .addItem('設定値一括更新', 'showSettingsUpdateForm')
    )
    .addSubMenu(
      ui.createMenu('💰 利益計算')
        .addItem('利益計算シート初期化', 'initializeProfitSheetOnly')
        .addItem('利益計算シート検証', 'verifyProfitSheetLayout')
        .addItem('参照式更新', 'updateProfitSheetFormulas')
        .addSeparator()
        .addItem('商品データ読み込み', 'loadProductDataMenu')
        .addItem('複数商品一括読み込み', 'showBulkLoadMenu')
        .addItem('テストデータ設定', 'setupProfitSheetTestData')
        .addSeparator()
        .addItem('為替レート検証', 'validateAndUpdateExchangeRate')
        .addItem('設定値更新', 'updateProfitCalculationSettings')
        .addItem('為替レート切替', 'toggleExchangeRateMode')
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
    .addSubMenu(
      ui.createMenu('🔗 JoomAPI設定')
        .addItem('🎫 トークン取得', 'acquireJoomToken')
        .addItem('📊 トークン取得状況', 'checkTokenAcquisitionStatus')
        .addItem('🔑 トークンステータス確認', 'showJoomTokenStatus')
        .addItem('🔄 トークンリフレッシュ', 'showJoomTokenRefreshMenu')
        .addItem('🗑️ トークンリセット', 'resetJoomTokens')
        .addSeparator()
        .addItem('⏰ 自動同期設定', 'showSyncTriggerSettings')
        .addItem('📊 同期状況確認', 'showSyncStatus')
    )
    .addSubMenu(
      ui.createMenu('📋 Joom注文取得')
        .addItem('🔄 最新注文を取得', 'fetchLatestJoomOrders')
        .addItem('📅 日時範囲で取得', 'fetchJoomOrdersByDateMenu')
        .addItem('🔍 特定注文を取得', 'fetchSpecificJoomOrder')
        .addItem('📊 取得状況確認', 'checkOrderFetchStatus')
        .addSeparator()
        .addItem('🐛 デバッグ: 最新注文データ表示', 'debugShowLatestOrdersRawData')
        .addItem('🐛 デバッグ: 特定注文データ表示', 'debugShowSpecificOrderRawData')
        .addItem('🐛 デバッグ: 日時範囲データ表示', 'debugShowDateRangeOrdersRawData')
    )
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
 * 新規メニュー関数の追加
 */

/**
 * Joom注文同期メニューの表示
 */
function showJoomOrderSyncMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Joom注文同期', 'Joom注文同期機能は今後実装予定です。', ui.ButtonSet.OK);
}

/**
 * 注文ステータス管理メニューの表示
 */
function showOrderStatusManagement() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('注文ステータス管理', '注文ステータス管理機能は今後実装予定です。', ui.ButtonSet.OK);
}

/**
 * データバックアップメニューの表示
 */
function showDataBackupMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('データバックアップ', 'データバックアップ機能は今後実装予定です。', ui.ButtonSet.OK);
}


/**
 * スプレッドシートが開かれた時の初期化
 */
function onOpen() {
  setupCustomMenu();
}
