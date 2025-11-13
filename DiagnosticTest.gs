/**
 * DiagnosticTest.gs
 * テストデータの診断用スクリプト
 */

/**
 * テストデータの診断を実行
 * Survey_Send_LogとSurvey_Responseのデータをチェック
 */
function diagnoseTestData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sendLogSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);
    const responseSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_RESPONSE);

    Logger.log('=== テストデータ診断開始 ===');

    // Survey_Send_Logの診断
    Logger.log('\n【Survey_Send_Log】');
    const sendLogData = sendLogSheet.getDataRange().getValues();
    Logger.log(`総行数: ${sendLogData.length - 1}件（ヘッダー除く）`);
    Logger.log(`データ範囲: ${sendLogSheet.getDataRange().getA1Notation()}`);

    let sendLogTestCount = 0;
    let sendLogAllCount = 0;

    for (let i = 1; i < sendLogData.length; i++) {
      const sendId = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.SEND_ID];

      if (sendId && sendId !== '') {
        sendLogAllCount++;

        if (sendId.toString().startsWith('LOG-TEST-')) {
          sendLogTestCount++;
          const candidateId = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.CANDIDATE_ID];
          const phase = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.PHASE];
          const email = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.EMAIL];
          const status = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.STATUS];

          if (sendLogTestCount <= 5) {  // 最初の5件のみ詳細表示
            Logger.log(`  行${i+1}: ${candidateId} | phase="${phase}" | email="${email}" | status="${status}"`);
          }

          // phaseに@が含まれているかチェック
          if (phase && phase.toString().includes('@')) {
            Logger.log(`    ❌ エラー: phaseにメールアドレスが入っています！`);
          }
        }
      }
    }
    Logger.log(`全データ件数: ${sendLogAllCount}件`);
    Logger.log(`テストデータ件数: ${sendLogTestCount}件`);
    if (sendLogTestCount === 0) {
      Logger.log(`⚠️ 警告: テストデータが見つかりません！`);
      Logger.log(`⚠️ 「➕ テストデータを生成」を実行してください`);
    }

    // Survey_Responseの診断
    Logger.log('\n【Survey_Response】');
    const responseData = responseSheet.getDataRange().getValues();
    Logger.log(`総行数: ${responseData.length - 1}件（ヘッダー除く）`);
    Logger.log(`データ範囲: ${responseSheet.getDataRange().getA1Notation()}`);

    let responseTestCount = 0;
    let responseAllCount = 0;

    for (let i = 1; i < responseData.length; i++) {
      const responseId = responseData[i][CONFIG.COLUMNS.SURVEY_RESPONSE.RESPONSE_ID];

      if (responseId && responseId !== '') {
        responseAllCount++;

        if (responseId.toString().startsWith('RESP-TEST-')) {
          responseTestCount++;
          const candidateId = responseData[i][CONFIG.COLUMNS.SURVEY_RESPONSE.CANDIDATE_ID];
          const phase = responseData[i][CONFIG.COLUMNS.SURVEY_RESPONSE.PHASE];
          const responseDate = responseData[i][CONFIG.COLUMNS.SURVEY_RESPONSE.RESPONSE_DATE];

          if (responseTestCount <= 5) {  // 最初の5件のみ詳細表示
            Logger.log(`  行${i+1}: ${candidateId} | phase="${phase}" | responseDate="${responseDate}"`);
          }

          // phaseに@が含まれているかチェック
          if (phase && phase.toString().includes('@')) {
            Logger.log(`    ❌ エラー: phaseにメールアドレスが入っています！`);
          }

          // phaseがundefinedまたは空かチェック
          if (!phase || phase === '') {
            Logger.log(`    ❌ エラー: phaseが空です！`);
          }
        }
      }
    }
    Logger.log(`全データ件数: ${responseAllCount}件`);
    Logger.log(`テストデータ件数: ${responseTestCount}件`);
    if (responseTestCount === 0) {
      Logger.log(`⚠️ 警告: テストデータが見つかりません！`);
      Logger.log(`⚠️ 「➕ テストデータを生成」を実行してください`);
    }

    // データマッチングテスト
    Logger.log('\n【データマッチングテスト】');
    Logger.log('C001 + 初回面談 でデータを取得してみます...');

    const testCandidateId = 'C001';
    const testPhase = '初回面談';

    // getSurveySendLogでデータ取得
    const sendLogs = getSurveySendLog(testCandidateId, testPhase);
    Logger.log(`  getSurveySendLog("${testCandidateId}", "${testPhase}"): ${sendLogs.length}件`);
    if (sendLogs.length > 0) {
      Logger.log(`    送信日時: ${sendLogs[0].send_time}`);
      Logger.log(`    ステータス: ${sendLogs[0].send_status}`);
    }

    // getSurveyResponseでデータ取得
    const responses = getSurveyResponse(testCandidateId, testPhase);
    Logger.log(`  getSurveyResponse("${testCandidateId}", "${testPhase}"): ${responses.length}件`);
    if (responses.length > 0) {
      Logger.log(`    回答日時: ${responses[0].response_date}`);
      Logger.log(`    phase: "${responses[0].phase}"`);
    }

    // getResponseSpeedDataでデータ取得
    const speedData = getResponseSpeedData(testCandidateId, testPhase);
    if (speedData) {
      Logger.log(`  getResponseSpeedData: ✅ 成功`);
      Logger.log(`    回答速度: ${speedData.hours}時間`);
      Logger.log(`    スコア: ${speedData.score}点`);
    } else {
      Logger.log(`  getResponseSpeedData: ❌ 失敗（nullが返された）`);
    }

    Logger.log('\n=== 診断完了 ===');

    SpreadsheetApp.getUi().alert(
      '診断完了',
      '診断が完了しました。\n\n実行ログ（表示 → ログ）で詳細を確認してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log(`❌ 診断エラー: ${error.message}`);
    Logger.log(error.stack);

    SpreadsheetApp.getUi().alert(
      'エラー',
      `診断中にエラーが発生しました:\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
