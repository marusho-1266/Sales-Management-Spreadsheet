/**
 * 商品削除機能
 * 在庫管理ツール システム開発
 * 作成日: 2026-03-06
 *
 * 商品IDを指定して在庫管理シートから物理削除する。
 * 影響範囲: 利益計算B2のクリア・ドロップダウン更新を実施。
 */

/**
 * 商品IDを指定して商品を削除する（エントリーポイント）
 * メニュー「商品管理」→「商品削除」から呼ばれる。
 * 商品ID入力ダイアログで削除対象を指定する。
 */
function deleteSelectedProduct() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 商品IDの入力を求める
  const promptResult = ui.prompt(
    '商品削除',
    MESSAGES.PROMPT.PRODUCT_DELETE_ENTER_ID,
    ui.ButtonSet.OK_CANCEL
  );

  if (promptResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const productIdStr = String(promptResult.getResponseText() || '').trim();
  if (!productIdStr) {
    ui.alert('商品削除', MESSAGES.ERROR.PRODUCT_DELETE_ENTER_ID_EMPTY, ui.ButtonSet.OK);
    return;
  }

  // 在庫管理シートを取得
  const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  if (!inventorySheet) {
    ui.alert('商品削除', MESSAGES.ERROR.SHEET_NOT_FOUND, ui.ButtonSet.OK);
    return;
  }
  const dataRange = inventorySheet.getDataRange();
  const values = dataRange.getValues();
  let foundRow = -1;
  let productName = '';
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === productIdStr) {
      foundRow = i + 1;
      productName = values[i][COLUMN_INDEXES.INVENTORY.PRODUCT_NAME - 1] || '';
      break;
    }
  }
  if (foundRow === -1) {
    ui.alert('商品削除', MESSAGES.ERROR.PRODUCT_DELETE_NOT_FOUND.replace('{id}', productIdStr), ui.ButtonSet.OK);
    return;
  }

  // 売上履歴の有無をチェック
  const salesCount = getSalesCountByProductId(spreadsheet, productIdStr);
  let confirmMessage = '';
  if (salesCount > 0) {
    confirmMessage = `この商品には売上履歴が${salesCount}件あります。\n在庫管理からは削除されますが、売上履歴は残ります。\n\n削除しますか？`;
  } else {
    confirmMessage = `商品ID ${productIdStr} (${productName || '(商品名なし)'}) を削除しますか？\n在庫管理から削除されますが、売上履歴は残ります。`;
  }

  const confirmResult = ui.alert('商品削除の確認', confirmMessage, ui.ButtonSet.OK_CANCEL);
  if (confirmResult !== ui.Button.OK) {
    return;
  }

  // 利益計算B2が削除対象の商品IDと一致する場合はB2をクリア
  const profitSheet = spreadsheet.getSheetByName(SHEET_NAMES.PROFIT);
  if (profitSheet) {
    const b2Value = profitSheet.getRange('B2').getValue();
    if (b2Value !== null && b2Value !== '' && String(b2Value).trim() === productIdStr) {
      profitSheet.getRange('B2').clearContent();
    }
  }

  // 行を削除
  inventorySheet.deleteRow(foundRow);

  // 利益計算シートの商品IDドロップダウンを更新
  refreshProfitSheetProductIdDropdown();

  const successMsg = MESSAGES.SUCCESS.PRODUCT_DELETED
    .replace('{id}', productIdStr)
    .replace('{name}', productName || '(商品名なし)');
  ui.alert('商品削除', successMsg, ui.ButtonSet.OK);
}

/**
 * 売上管理シートで指定した商品IDの売上件数をカウントする
 * @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
 * @param {string} productId - 商品ID（比較時は文字列として扱う）
 * @returns {number} 売上件数
 */
function getSalesCountByProductId(spreadsheet, productId) {
  try {
    const salesSheet = spreadsheet.getSheetByName(SHEET_NAMES.SALES);
    if (!salesSheet) {
      return 0;
    }
    const lastRow = salesSheet.getLastRow();
    if (lastRow < 2) {
      return 0;
    }
    const productIdCol = COLUMN_INDEXES.SALES.PRODUCT_ID;
    const dataRange = salesSheet.getRange(2, productIdCol, lastRow, productIdCol);
    const values = dataRange.getValues();
    const productIdStr = String(productId || '').trim();
    let count = 0;
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim() === productIdStr) {
        count++;
      }
    }
    return count;
  } catch (e) {
    console.warn('売上件数取得エラー:', e);
    return 0;
  }
}
