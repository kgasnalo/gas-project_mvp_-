/**
 * TestDataGenerator.gs
 * アンケート回答速度計算機能のテストデータ生成スクリプト
 *
 * 【主要機能】
 * - Survey_Send_Logへのテストデータ投入
 * - Survey_Responseへのテストデータ投入
 * - 様々な回答速度パターンのテストデータ生成
 * - テストデータのクリア
 * - データバリデーション（列数チェック）
 *
 * @version 1.1
 * @date 2025-11-13
 */

/**
 * テストデータを一括生成
 * Survey_Send_LogとSurvey_Responseに様々なパターンのテストデータを投入
 */
function generateAllTestData() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'テストデータ生成',
      'Survey_Send_LogとSurvey_Responseにテストデータを生成しますか？\n\n' +
      '※既存のテストデータは上書きされます',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      Logger.log('テストデータ生成がキャンセルされました');
      return;
    }

    Logger.log('📊 テストデータ生成を開始します...');

    // 1. テストデータをクリア
    clearTestData();

    // 2. Survey_Send_Logにテストデータを投入
    const sendLogCount = generateSurveySendLogTestData();

    // 3. Survey_Responseにテストデータを投入
    const responseCount = generateSurveyResponseTestData();

    Logger.log(`✅ テストデータ生成完了: 送信ログ ${sendLogCount}件、回答 ${responseCount}件`);

    ui.alert(
      'テストデータ生成完了',
      `以下のテストデータを生成しました:\n\n` +
      `📤 Survey_Send_Log: ${sendLogCount}件\n` +
      `📥 Survey_Response: ${responseCount}件\n\n` +
      '次に「📈 回答速度を一括計算」を実行してください。',
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log(`❌ generateAllTestDataエラー: ${error.message}`);
    Logger.log(error.stack);

    SpreadsheetApp.getUi().alert(
      'エラー',
      `テストデータ生成中にエラーが発生しました:\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Survey_Send_Logにテストデータを投入
 *
 * @return {number} 投入したデータ件数
 */
function generateSurveySendLogTestData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sendLogSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);
    const masterSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);

    if (!sendLogSheet || !masterSheet) {
      throw new Error('必要なシートが見つかりません');
    }

    // Candidates_Masterから候補者情報を取得（最初の5名）
    const masterData = masterSheet.getDataRange().getValues();
    const candidates = [];

    for (let i = 1; i < Math.min(6, masterData.length); i++) {
      const candidateId = masterData[i][CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID];
      const name = masterData[i][CONFIG.COLUMNS.CANDIDATES_MASTER.NAME];
      const email = masterData[i][CONFIG.COLUMNS.CANDIDATES_MASTER.EMAIL];

      if (candidateId && name) {
        candidates.push({ candidateId, name, email: email || `${candidateId}@example.com` });
      }
    }

    if (candidates.length === 0) {
      throw new Error('Candidates_Masterに候補者データがありません');
    }

    Logger.log(`📋 ${candidates.length}名の候補者にテストデータを生成します`);

    // アンケート種別
    const phases = ['初回面談', '社員面談', '2次面接', '内定後'];

    let count = 0;
    const now = new Date();

    // 各候補者に対して、各フェーズのアンケート送信記録を作成
    candidates.forEach((candidate, idx) => {
      phases.forEach((phase, phaseIdx) => {
        // 送信日時: 現在から1-10日前のランダムな時刻
        const daysAgo = 1 + idx + phaseIdx;
        const sendTime = new Date(now.getTime() - daysAgo * 24 * 60 * 60 * 1000);

        // log_id生成
        const logId = `LOG-TEST-${candidate.candidateId}-${phase.replace(/\s/g, '')}-${Date.now()}-${count}`;

        // Survey_Send_Logに追加
        sendLogSheet.appendRow([
          logId,                    // A: SEND_ID
          candidate.candidateId,    // B: CANDIDATE_ID
          candidate.name,           // C: NAME
          candidate.email,          // D: EMAIL
          phase,                    // E: PHASE
          sendTime,                 // F: SEND_TIME
          '成功',                   // G: STATUS
          ''                        // H: ERROR_MSG
        ]);

        count++;
        Logger.log(`✅ 送信ログ追加: ${candidate.name} (${phase}) - ${Utilities.formatDate(sendTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`);
      });
    });

    Logger.log(`✅ Survey_Send_Logに${count}件のテストデータを投入しました`);
    return count;

  } catch (error) {
    Logger.log(`❌ generateSurveySendLogTestDataエラー: ${error.message}`);
    throw error;
  }
}

/**
 * Survey_Responseにテストデータを投入
 * 送信ログに対応する回答データを、様々な回答速度パターンで生成
 *
 * @return {number} 投入したデータ件数
 */
function generateSurveyResponseTestData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sendLogSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);
    const responseSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_RESPONSE);
    const masterSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);

    if (!sendLogSheet || !responseSheet || !masterSheet) {
      throw new Error('必要なシートが見つかりません');
    }

    // Survey_Send_Logからテストデータを取得
    const sendLogData = sendLogSheet.getDataRange().getValues();

    // 回答速度パターン（時間単位）
    const responsePatterns = [
      1,      // 0-2時間: 100点
      3,      // 2-6時間: 100-80点
      4,      // 2-6時間: 100-80点
      12,     // 6-24時間: 80-50点
      18,     // 6-24時間: 80-50点
      30,     // 24-48時間: 50-20点
      36,     // 24-48時間: 50-20点
      60,     // 48時間以上: 20点以下
      72,     // 48時間以上: 20点以下
      96      // 48時間以上: 20点以下
    ];

    let count = 0;
    let patternIndex = 0;

    // ヘッダー行をスキップ
    for (let i = 1; i < sendLogData.length; i++) {
      const candidateId = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.CANDIDATE_ID];
      const name = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.NAME];
      const phase = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.PHASE];
      const sendTime = new Date(sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.SEND_TIME]);

      // 回答速度パターンをローテーション
      const hoursDelay = responsePatterns[patternIndex % responsePatterns.length];
      patternIndex++;

      // 回答日時 = 送信日時 + 回答速度
      const responseDate = new Date(sendTime.getTime() + hoursDelay * 60 * 60 * 1000);

      // response_id生成
      const responseId = `RESP-TEST-${candidateId}-${phase.replace(/\s/g, '')}-${Date.now()}-${count}`;

      // 志望度（ランダム: 6-10）
      const aspiration = Math.floor(Math.random() * 5) + 6;

      // Survey_Responseに追加
      responseSheet.appendRow([
        responseId,               // A: RESPONSE_ID
        candidateId,              // B: CANDIDATE_ID
        name,                     // C: NAME
        responseDate,             // D: RESPONSE_DATE
        aspiration,               // E: ASPIRATION (志望度)
        '',                       // F: CONCERNS (懸念事項)
        '',                       // G: OTHER_COMPANIES (他社選考状況)
        '',                       // H: COMMENTS (その他コメント)
        phase                     // I: PHASE (アンケート種別)
      ]);

      count++;
      Logger.log(`✅ 回答データ追加: ${name} (${phase}) - ${hoursDelay}時間後に回答`);
    }

    Logger.log(`✅ Survey_Responseに${count}件のテストデータを投入しました`);
    return count;

  } catch (error) {
    Logger.log(`❌ generateSurveyResponseTestDataエラー: ${error.message}`);
    throw error;
  }
}

/**
 * テストデータをクリア
 * Survey_Send_Log、Survey_Response、Survey_Analysisのテストデータを削除
 */
function clearTestData() {
  try {
    Logger.log('🗑️ テストデータをクリアします...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Survey_Send_Logのテストデータをクリア
    const sendLogSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);
    if (sendLogSheet) {
      const sendLogData = sendLogSheet.getDataRange().getValues();
      const rowsToDelete = [];

      for (let i = sendLogData.length - 1; i >= 1; i--) {
        const logId = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.LOG_ID];
        if (logId && logId.toString().startsWith('LOG-TEST-')) {
          rowsToDelete.push(i + 1); // 行番号は1始まり
        }
      }

      rowsToDelete.forEach(row => {
        sendLogSheet.deleteRow(row);
      });

      Logger.log(`✅ Survey_Send_Logから${rowsToDelete.length}件のテストデータを削除しました`);
    }

    // Survey_Responseのテストデータをクリア
    const responseSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_RESPONSE);
    if (responseSheet) {
      const responseData = responseSheet.getDataRange().getValues();
      const rowsToDelete = [];

      for (let i = responseData.length - 1; i >= 1; i--) {
        const responseId = responseData[i][CONFIG.COLUMNS.SURVEY_RESPONSE.RESPONSE_ID];
        if (responseId && responseId.toString().startsWith('RESP-TEST-')) {
          rowsToDelete.push(i + 1);
        }
      }

      rowsToDelete.forEach(row => {
        responseSheet.deleteRow(row);
      });

      Logger.log(`✅ Survey_Responseから${rowsToDelete.length}件のテストデータを削除しました`);
    }

    // Survey_Analysisのテストデータをクリア
    const analysisSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_ANALYSIS);
    if (analysisSheet) {
      const analysisData = analysisSheet.getDataRange().getValues();
      const rowsToDelete = [];

      for (let i = analysisData.length - 1; i >= 1; i--) {
        const analysisId = analysisData[i][CONFIG.COLUMNS.SURVEY_ANALYSIS.ANALYSIS_ID];
        if (analysisId && analysisId.toString().includes('TEST')) {
          rowsToDelete.push(i + 1);
        }
      }

      rowsToDelete.forEach(row => {
        analysisSheet.deleteRow(row);
      });

      Logger.log(`✅ Survey_Analysisから${rowsToDelete.length}件のテストデータを削除しました`);
    }

    Logger.log('✅ テストデータのクリアが完了しました');

  } catch (error) {
    Logger.log(`❌ clearTestDataエラー: ${error.message}`);
    throw error;
  }
}

/**
 * テストデータをクリア（メニュー用）
 */
function clearAllTestData() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'テストデータクリア',
      '以下のテストデータを削除しますか？\n\n' +
      '・Survey_Send_Log (LOG-TEST-で始まるデータ)\n' +
      '・Survey_Response (RESP-TEST-で始まるデータ)\n' +
      '・Survey_Analysis (テストデータ)\n\n' +
      '※この操作は取り消せません',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      Logger.log('テストデータクリアがキャンセルされました');
      return;
    }

    clearTestData();

    ui.alert(
      'テストデータクリア完了',
      'テストデータを削除しました。',
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log(`❌ clearAllTestDataエラー: ${error.message}`);

    SpreadsheetApp.getUi().alert(
      'エラー',
      `テストデータクリア中にエラーが発生しました:\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * テストデータの状況を確認
 */
function checkTestDataStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let message = '【📊 テストデータ状況】\n\n';

    // Survey_Send_Logのテストデータ数
    const sendLogSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);
    if (sendLogSheet) {
      const sendLogData = sendLogSheet.getDataRange().getValues();
      let testCount = 0;
      let successCount = 0;

      for (let i = 1; i < sendLogData.length; i++) {
        const logId = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.LOG_ID];
        const status = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.STATUS];

        if (logId && logId.toString().startsWith('LOG-TEST-')) {
          testCount++;
          if (status === '成功') successCount++;
        }
      }

      message += `📤 Survey_Send_Log\n`;
      message += `   テストデータ: ${testCount}件\n`;
      message += `   送信成功: ${successCount}件\n\n`;
    }

    // Survey_Responseのテストデータ数
    const responseSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_RESPONSE);
    if (responseSheet) {
      const responseData = responseSheet.getDataRange().getValues();
      let testCount = 0;

      for (let i = 1; i < responseData.length; i++) {
        const responseId = responseData[i][CONFIG.COLUMNS.SURVEY_RESPONSE.RESPONSE_ID];
        if (responseId && responseId.toString().startsWith('RESP-TEST-')) {
          testCount++;
        }
      }

      message += `📥 Survey_Response\n`;
      message += `   テストデータ: ${testCount}件\n\n`;
    }

    // Survey_Analysisのデータ数
    const analysisSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_ANALYSIS);
    if (analysisSheet) {
      const analysisData = analysisSheet.getDataRange().getValues();
      const dataCount = analysisData.length - 1; // ヘッダーを除く

      message += `📊 Survey_Analysis\n`;
      message += `   分析データ: ${dataCount}件\n\n`;
    }

    message += '━━━━━━━━━━━━━━━━\n';
    message += '💡 次のステップ\n';
    message += '━━━━━━━━━━━━━━━━\n';
    message += '1. テストデータがない場合:\n';
    message += '   →「テストデータを生成」を実行\n\n';
    message += '2. テストデータがある場合:\n';
    message += '   →「📈 回答速度を一括計算」を実行';

    SpreadsheetApp.getUi().alert(
      'テストデータ状況',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log(`❌ checkTestDataStatusエラー: ${error.message}`);
  }
}

/**
 * データバリデーション：ヘッダー数とデータ列数の一致をチェック
 *
 * @param {string} sheetName - シート名
 * @return {Object} { valid: boolean, errors: Array }
 */
function validateSheetColumnCount(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return {
        valid: false,
        errors: [`シート「${sheetName}」が見つかりません`]
      };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return {
        valid: true,
        errors: []
      };
    }

    const headerCount = data[0].filter(cell => cell !== '').length;
    const errors = [];

    // 各データ行の列数をチェック
    for (let i = 1; i < data.length; i++) {
      const rowData = data[i];
      const nonEmptyCount = rowData.filter(cell => cell !== '').length;

      if (nonEmptyCount !== headerCount) {
        errors.push(
          `行${i + 1}: ヘッダー${headerCount}列に対してデータ${nonEmptyCount}列（不一致）`
        );
      }
    }

    return {
      valid: errors.length === 0,
      headerCount: headerCount,
      dataRowCount: data.length - 1,
      errors: errors
    };

  } catch (error) {
    Logger.log(`❌ validateSheetColumnCountエラー: ${error.message}`);
    return {
      valid: false,
      errors: [`エラー: ${error.message}`]
    };
  }
}

/**
 * 全テストデータのバリデーションを実行
 */
function validateAllTestData() {
  try {
    Logger.log('🔍 テストデータのバリデーションを開始します...');

    const results = [];

    // Survey_Send_Logのバリデーション
    const sendLogResult = validateSheetColumnCount(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);
    results.push({
      sheet: 'Survey_Send_Log',
      ...sendLogResult
    });

    // Survey_Responseのバリデーション
    const responseResult = validateSheetColumnCount(CONFIG.SHEET_NAMES.SURVEY_RESPONSE);
    results.push({
      sheet: 'Survey_Response',
      ...responseResult
    });

    // Survey_Analysisのバリデーション
    const analysisResult = validateSheetColumnCount(CONFIG.SHEET_NAMES.SURVEY_ANALYSIS);
    results.push({
      sheet: 'Survey_Analysis',
      ...analysisResult
    });

    // 結果をログ出力
    let allValid = true;
    let message = '【🔍 データバリデーション結果】\n\n';

    results.forEach(result => {
      if (result.valid) {
        Logger.log(`✅ ${result.sheet}: 問題なし（ヘッダー${result.headerCount}列、データ${result.dataRowCount}行）`);
        message += `✅ ${result.sheet}\n`;
        message += `   ヘッダー: ${result.headerCount}列\n`;
        message += `   データ: ${result.dataRowCount}行\n`;
        message += `   検証: 問題なし\n\n`;
      } else {
        allValid = false;
        Logger.log(`❌ ${result.sheet}: エラーあり`);
        result.errors.forEach(error => {
          Logger.log(`   - ${error}`);
        });
        message += `❌ ${result.sheet}\n`;
        message += `   エラー:\n`;
        result.errors.forEach(error => {
          message += `   - ${error}\n`;
        });
        message += '\n';
      }
    });

    if (allValid) {
      message += '━━━━━━━━━━━━━━━━\n';
      message += '✅ 全シートのデータ構造が正しいです';
    } else {
      message += '━━━━━━━━━━━━━━━━\n';
      message += '⚠️ データ構造にエラーがあります\n';
      message += '上記のエラーを修正してください';
    }

    SpreadsheetApp.getUi().alert(
      'データバリデーション結果',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return allValid;

  } catch (error) {
    Logger.log(`❌ validateAllTestDataエラー: ${error.message}`);
    Logger.log(error.stack);

    SpreadsheetApp.getUi().alert(
      'エラー',
      `バリデーション中にエラーが発生しました:\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return false;
  }
}
