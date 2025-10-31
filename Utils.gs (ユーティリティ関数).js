/**
 * ユーティリティ関数
 */

/**
 * シートを取得または作成
 * @param {string} sheetName - シート名
 * @param {number} index - シートの位置（オプション）
 * @return {Sheet} シートオブジェクト
 */
function getOrCreateSheet(sheetName, index = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    if (index !== null) {
      sheet = ss.insertSheet(sheetName, index);
    } else {
      sheet = ss.insertSheet(sheetName);
    }
    Logger.log(`✅ シート「${sheetName}」を作成しました`);
  } else {
    Logger.log(`ℹ️ シート「${sheetName}」は既に存在します`);
  }
  
  return sheet;
}

/**
 * ヘッダー行を設定
 * @param {Sheet} sheet - シートオブジェクト
 * @param {Array} headers - ヘッダーの配列
 */
function setHeaders(sheet, headers) {
  // ヘッダー行を設定
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行の書式設定
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // ヘッダー行を固定
  sheet.setFrozenRows(1);
}

/**
 * 列幅を設定
 * @param {Sheet} sheet - シートオブジェクト
 * @param {Array} widths - 列幅の配列（ピクセル）
 */
function setColumnWidths(sheet, widths) {
  widths.forEach((width, index) => {
    if (width === 'auto') {
      sheet.autoResizeColumn(index + 1);
    } else {
      sheet.setColumnWidth(index + 1, width);
    }
  });
}

/**
 * データ検証（ドロップダウン）を設定
 * @param {Sheet} sheet - シートオブジェクト
 * @param {string} range - 範囲（例: 'C2:C1000'）
 * @param {Array} options - 選択肢の配列
 * @param {string} helpText - ヘルプテキスト（オプション）
 */
function setDropdownValidation(sheet, range, options, helpText = '') {
  const targetRange = sheet.getRange(range);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .setHelpText(helpText)
    .build();
  targetRange.setDataValidation(rule);
}

/**
 * 条件付き書式を設定（3段階：高・中・低）
 * @param {Sheet} sheet - シートオブジェクト
 * @param {string} range - 範囲（例: 'H2:H1000'）
 * @param {number} highThreshold - 高スコアの閾値
 * @param {number} lowThreshold - 低スコアの閾値
 */
function setThreeColorConditionalFormat(sheet, range, highThreshold, lowThreshold) {
  const targetRange = sheet.getRange(range);
  
  // 高スコア（緑）
  const rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(highThreshold)
    .setBackground(CONFIG.COLORS.HIGH_SCORE)
    .setRanges([targetRange])
    .build();
  
  // 中スコア（黄）
  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(lowThreshold, highThreshold - 0.01)
    .setBackground(CONFIG.COLORS.MEDIUM_SCORE)
    .setRanges([targetRange])
    .build();
  
  // 低スコア（赤）
  const rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(lowThreshold)
    .setBackground(CONFIG.COLORS.LOW_SCORE)
    .setRanges([targetRange])
    .build();
  
  // ルールを適用
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule1, rule2, rule3);
  sheet.setConditionalFormatRules(rules);
}

/**
 * 日付を YYYY-MM-DD HH:mm:ss 形式でフォーマット
 * @param {Date} date - 日付オブジェクト
 * @return {string} フォーマットされた日付文字列
 */
function formatDateTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/**
 * エラーログを記録
 * @param {string} functionName - 関数名
 * @param {Error} error - エラーオブジェクト
 */
function logError(functionName, error) {
  Logger.log(`❌ [${functionName}] エラー: ${error.message}`);
  Logger.log(error.stack);
}