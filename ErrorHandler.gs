/**
 * ========================================
 * Phase 3.5: エラーハンドリング
 * ========================================
 */

/**
 * エラーログの記録
 *
 * @param {string} functionName - 関数名
 * @param {Error} error - エラーオブジェクト
 */
function logError(functionName, error) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let errorLog = ss.getSheetByName('Error_Log');

    // Error_Logシートが存在しない場合は作成
    if (!errorLog) {
      errorLog = ss.insertSheet('Error_Log');

      // ヘッダー行を追加
      errorLog.getRange(1, 1, 1, 5).setValues([[
        'Timestamp',
        'Function Name',
        'Error Message',
        'Error Stack',
        'Status'
      ]]);

      // ヘッダー行の書式設定
      const headerRange = errorLog.getRange(1, 1, 1, 5);
      headerRange.setBackground(CONFIG.COLORS.HEADER_BG);
      headerRange.setFontColor(CONFIG.COLORS.HEADER_TEXT);
      headerRange.setFontWeight('bold');
    }

    // エラー情報を追加
    errorLog.appendRow([
      new Date(),
      functionName,
      error.message,
      error.stack || 'N/A',
      'Unresolved'
    ]);

    Logger.log(`✅ エラーログ記録: ${functionName} - ${error.message}`);

  } catch (logError) {
    Logger.log(`❌ エラーログ記録失敗: ${logError}`);
  }
}

/**
 * エラー通知の送信
 *
 * @param {string} message - エラーメッセージ
 */
function sendErrorNotification(message) {
  try {
    // メール通知（オプション）
    // const email = 'your-email@example.com';
    // MailApp.sendEmail(email, 'エラー通知', message);

    Logger.log(`⚠️ エラー通知: ${message}`);

    // Slack通知などを追加する場合はここに実装

  } catch (error) {
    Logger.log(`❌ エラー通知送信失敗: ${error}`);
  }
}

/**
 * リトライ処理
 *
 * @param {Function} func - 実行する関数
 * @param {number} maxRetries - 最大リトライ回数
 * @param {number} delay - リトライ間隔（ミリ秒）
 * @return {*} 関数の戻り値
 */
function retryOnError(func, maxRetries = 3, delay = 1000) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      return func();
    } catch (error) {
      Logger.log(`⚠️ リトライ ${i + 1}/${maxRetries}: ${error.message}`);

      if (i === maxRetries - 1) {
        throw error; // 最後のリトライで失敗したらエラーをスロー
      }

      Utilities.sleep(delay);
    }
  }
}

/**
 * データ検証
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {boolean} 検証結果
 */
function validateData(candidateId, phase) {
  try {
    // 候補者IDの検証
    if (!candidateId || candidateId === '') {
      Logger.log('❌ 候補者IDが空です');
      return false;
    }

    // フェーズの検証
    const validPhases = CONFIG.VALIDATION_OPTIONS.SURVEY_PHASE;
    if (!validPhases.includes(phase)) {
      Logger.log(`❌ 不正なフェーズ: ${phase}`);
      return false;
    }

    // 候補者の存在確認
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);

    if (!master) {
      Logger.log('❌ Candidates_Masterシートが見つかりません');
      return false;
    }

    const data = master.getDataRange().getValues();
    let found = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === candidateId) {
        found = true;
        break;
      }
    }

    if (!found) {
      Logger.log(`❌ 候補者が見つかりません: ${candidateId}`);
      return false;
    }

    Logger.log(`✅ データ検証成功: ${candidateId}, ${phase}`);
    return true;

  } catch (error) {
    Logger.log(`❌ データ検証エラー: ${error}`);
    return false;
  }
}
