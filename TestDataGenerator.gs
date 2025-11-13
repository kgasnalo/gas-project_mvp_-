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
 * Survey_Send_Logにテストデータを投入（明示的な配列インデックス版）
 *
 * 【重要】配列のインデックスを明示して、列のズレを完全に防止
 * Survey_Send_Log構造（8列）:
 *   [0] A: send_id
 *   [1] B: candidate_id
 *   [2] C: candidate_name
 *   [3] D: email
 *   [4] E: phase
 *   [5] F: send_time
 *   [6] G: send_status
 *   [7] H: error_message
 *
 * @return {number} 投入したデータ件数
 */
function generateSurveySendLogTestData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sendLogSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);

    if (!sendLogSheet) {
      throw new Error('Survey_Send_Logシートが見つかりません');
    }

    Logger.log('📋 Survey_Send_Logにテストデータを生成します（20件）');

    const now = new Date();

    // テストデータ：20件の送信ログ（5名 × 4フェーズ）
    // 配列のインデックスを明示的に記載
    const testData = [
      // C001 - 田中太郎
      [
        'LOG-TEST-001',           // [0] send_id
        'C001',                   // [1] candidate_id
        '田中太郎',               // [2] candidate_name
        'tanaka@example.com',     // [3] email
        '初回面談',               // [4] phase ← 重要！
        new Date(now.getTime() - 1 * 24 * 60 * 60 * 1000),  // [5] send_time (1日前)
        '成功',                   // [6] send_status
        ''                        // [7] error_message
      ],
      [
        'LOG-TEST-002',
        'C001',
        '田中太郎',
        'tanaka@example.com',
        '社員面談',               // [4] phase
        new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000),  // 2日前
        '成功',
        ''
      ],
      [
        'LOG-TEST-003',
        'C001',
        '田中太郎',
        'tanaka@example.com',
        '2次面接',                // [4] phase
        new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000),  // 3日前
        '成功',
        ''
      ],
      [
        'LOG-TEST-004',
        'C001',
        '田中太郎',
        'tanaka@example.com',
        '内定後',                 // [4] phase
        new Date(now.getTime() - 4 * 24 * 60 * 60 * 1000),  // 4日前
        '成功',
        ''
      ],

      // C002 - 佐藤花子
      [
        'LOG-TEST-005',
        'C002',
        '佐藤花子',
        'sato@example.com',
        '初回面談',
        new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],
      [
        'LOG-TEST-006',
        'C002',
        '佐藤花子',
        'sato@example.com',
        '社員面談',
        new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],
      [
        'LOG-TEST-007',
        'C002',
        '佐藤花子',
        'sato@example.com',
        '2次面接',
        new Date(now.getTime() - 4 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],
      [
        'LOG-TEST-008',
        'C002',
        '佐藤花子',
        'sato@example.com',
        '内定後',
        new Date(now.getTime() - 5 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],

      // C003 - 鈴木一郎
      [
        'LOG-TEST-009',
        'C003',
        '鈴木一郎',
        'suzuki@example.com',
        '初回面談',
        new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],
      [
        'LOG-TEST-010',
        'C003',
        '鈴木一郎',
        'suzuki@example.com',
        '社員面談',
        new Date(now.getTime() - 4 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],
      [
        'LOG-TEST-011',
        'C003',
        '鈴木一郎',
        'suzuki@example.com',
        '2次面接',
        new Date(now.getTime() - 5 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],
      [
        'LOG-TEST-012',
        'C003',
        '鈴木一郎',
        'suzuki@example.com',
        '内定後',
        new Date(now.getTime() - 6 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],

      // C004 - 高橋美咲
      [
        'LOG-TEST-013',
        'C004',
        '高橋美咲',
        'takahashi@example.com',
        '初回面談',
        new Date(now.getTime() - 4 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],
      [
        'LOG-TEST-014',
        'C004',
        '高橋美咲',
        'takahashi@example.com',
        '社員面談',
        new Date(now.getTime() - 5 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],
      [
        'LOG-TEST-015',
        'C004',
        '高橋美咲',
        'takahashi@example.com',
        '2次面接',
        new Date(now.getTime() - 6 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],
      [
        'LOG-TEST-016',
        'C004',
        '高橋美咲',
        'takahashi@example.com',
        '内定後',
        new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],

      // C005 - 渡辺健太
      [
        'LOG-TEST-017',
        'C005',
        '渡辺健太',
        'watanabe@example.com',
        '初回面談',
        new Date(now.getTime() - 5 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],
      [
        'LOG-TEST-018',
        'C005',
        '渡辺健太',
        'watanabe@example.com',
        '社員面談',
        new Date(now.getTime() - 6 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],
      [
        'LOG-TEST-019',
        'C005',
        '渡辺健太',
        'watanabe@example.com',
        '2次面接',
        new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ],
      [
        'LOG-TEST-020',
        'C005',
        '渡辺健太',
        'watanabe@example.com',
        '内定後',
        new Date(now.getTime() - 8 * 24 * 60 * 60 * 1000),
        '成功',
        ''
      ]
    ];

    // データを一括書き込み
    if (testData.length > 0) {
      const startRow = sendLogSheet.getLastRow() + 1;
      sendLogSheet.getRange(startRow, 1, testData.length, 8).setValues(testData);
    }

    Logger.log(`✅ Survey_Send_Logに${testData.length}件のテストデータを投入しました`);
    return testData.length;

  } catch (error) {
    Logger.log(`❌ generateSurveySendLogTestDataエラー: ${error.message}`);
    throw error;
  }
}

/**
 * Survey_Responseにテストデータを投入（明示的な配列インデックス版）
 *
 * 【重要】配列のインデックスを明示して、列のズレを完全に防止
 * Survey_Response構造（9列）:
 *   [0] A: response_id
 *   [1] B: candidate_id
 *   [2] C: 氏名
 *   [3] D: 回答日時
 *   [4] E: 志望度
 *   [5] F: 懸念事項
 *   [6] G: 他社選考状況
 *   [7] H: その他コメント
 *   [8] I: アンケート種別 ← 重要！
 *
 * @return {number} 投入したデータ件数
 */
function generateSurveyResponseTestData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const responseSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_RESPONSE);

    if (!responseSheet) {
      throw new Error('Survey_Responseシートが見つかりません');
    }

    Logger.log('📋 Survey_Responseにテストデータを生成します（20件）');

    const now = new Date();

    // 回答速度パターン（時間単位）: 1h, 3h, 4h, 12h, 18h, 30h, 36h, 60h, 72h, 96h
    const responseDelays = [1, 3, 4, 12, 18, 30, 36, 60, 72, 96, 1, 3, 4, 12, 18, 30, 36, 60, 72, 96];

    // テストデータ：20件の回答（5名 × 4フェーズ）
    // 配列のインデックスを明示的に記載
    const testData = [
      // C001 - 田中太郎（回答速度: 1h, 3h, 4h, 12h）
      [
        'RESP-TEST-001',                                              // [0] response_id
        'C001',                                                       // [1] candidate_id
        '田中太郎',                                                   // [2] 氏名
        new Date(now.getTime() - 1 * 24 * 60 * 60 * 1000 + 1 * 60 * 60 * 1000),   // [3] 回答日時 (送信1日前 + 1時間後)
        8,                                                            // [4] 志望度
        '',                                                           // [5] 懸念事項
        '',                                                           // [6] 他社選考状況
        '',                                                           // [7] その他コメント
        '初回面談'                                                    // [8] アンケート種別 ← 重要！
      ],
      [
        'RESP-TEST-002',
        'C001',
        '田中太郎',
        new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000 + 3 * 60 * 60 * 1000),   // 送信2日前 + 3時間後
        9,
        '',
        '',
        '',
        '社員面談'                                                    // [8] phase
      ],
      [
        'RESP-TEST-003',
        'C001',
        '田中太郎',
        new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000 + 4 * 60 * 60 * 1000),   // 送信3日前 + 4時間後
        7,
        '',
        '',
        '',
        '2次面接'                                                     // [8] phase
      ],
      [
        'RESP-TEST-004',
        'C001',
        '田中太郎',
        new Date(now.getTime() - 4 * 24 * 60 * 60 * 1000 + 12 * 60 * 60 * 1000),  // 送信4日前 + 12時間後
        10,
        '',
        '',
        '',
        '内定後'                                                      // [8] phase
      ],

      // C002 - 佐藤花子（回答速度: 18h, 30h, 36h, 60h）
      [
        'RESP-TEST-005',
        'C002',
        '佐藤花子',
        new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000 + 18 * 60 * 60 * 1000),  // 18時間後
        6,
        '',
        '',
        '',
        '初回面談'
      ],
      [
        'RESP-TEST-006',
        'C002',
        '佐藤花子',
        new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000 + 30 * 60 * 60 * 1000),  // 30時間後
        7,
        '',
        '',
        '',
        '社員面談'
      ],
      [
        'RESP-TEST-007',
        'C002',
        '佐藤花子',
        new Date(now.getTime() - 4 * 24 * 60 * 60 * 1000 + 36 * 60 * 60 * 1000),  // 36時間後
        8,
        '',
        '',
        '',
        '2次面接'
      ],
      [
        'RESP-TEST-008',
        'C002',
        '佐藤花子',
        new Date(now.getTime() - 5 * 24 * 60 * 60 * 1000 + 60 * 60 * 60 * 1000),  // 60時間後
        9,
        '',
        '',
        '',
        '内定後'
      ],

      // C003 - 鈴木一郎（回答速度: 72h, 96h, 1h, 3h）
      [
        'RESP-TEST-009',
        'C003',
        '鈴木一郎',
        new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000 + 72 * 60 * 60 * 1000),  // 72時間後
        7,
        '',
        '',
        '',
        '初回面談'
      ],
      [
        'RESP-TEST-010',
        'C003',
        '鈴木一郎',
        new Date(now.getTime() - 4 * 24 * 60 * 60 * 1000 + 96 * 60 * 60 * 1000),  // 96時間後
        6,
        '',
        '',
        '',
        '社員面談'
      ],
      [
        'RESP-TEST-011',
        'C003',
        '鈴木一郎',
        new Date(now.getTime() - 5 * 24 * 60 * 60 * 1000 + 1 * 60 * 60 * 1000),   // 1時間後
        10,
        '',
        '',
        '',
        '2次面接'
      ],
      [
        'RESP-TEST-012',
        'C003',
        '鈴木一郎',
        new Date(now.getTime() - 6 * 24 * 60 * 60 * 1000 + 3 * 60 * 60 * 1000),   // 3時間後
        9,
        '',
        '',
        '',
        '内定後'
      ],

      // C004 - 高橋美咲（回答速度: 4h, 12h, 18h, 30h）
      [
        'RESP-TEST-013',
        'C004',
        '高橋美咲',
        new Date(now.getTime() - 4 * 24 * 60 * 60 * 1000 + 4 * 60 * 60 * 1000),   // 4時間後
        8,
        '',
        '',
        '',
        '初回面談'
      ],
      [
        'RESP-TEST-014',
        'C004',
        '高橋美咲',
        new Date(now.getTime() - 5 * 24 * 60 * 60 * 1000 + 12 * 60 * 60 * 1000),  // 12時間後
        7,
        '',
        '',
        '',
        '社員面談'
      ],
      [
        'RESP-TEST-015',
        'C004',
        '高橋美咲',
        new Date(now.getTime() - 6 * 24 * 60 * 60 * 1000 + 18 * 60 * 60 * 1000),  // 18時間後
        9,
        '',
        '',
        '',
        '2次面接'
      ],
      [
        'RESP-TEST-016',
        'C004',
        '高橋美咲',
        new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000 + 30 * 60 * 60 * 1000),  // 30時間後
        6,
        '',
        '',
        '',
        '内定後'
      ],

      // C005 - 渡辺健太（回答速度: 36h, 60h, 72h, 96h）
      [
        'RESP-TEST-017',
        'C005',
        '渡辺健太',
        new Date(now.getTime() - 5 * 24 * 60 * 60 * 1000 + 36 * 60 * 60 * 1000),  // 36時間後
        10,
        '',
        '',
        '',
        '初回面談'
      ],
      [
        'RESP-TEST-018',
        'C005',
        '渡辺健太',
        new Date(now.getTime() - 6 * 24 * 60 * 60 * 1000 + 60 * 60 * 60 * 1000),  // 60時間後
        8,
        '',
        '',
        '',
        '社員面談'
      ],
      [
        'RESP-TEST-019',
        'C005',
        '渡辺健太',
        new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000 + 72 * 60 * 60 * 1000),  // 72時間後
        7,
        '',
        '',
        '',
        '2次面接'
      ],
      [
        'RESP-TEST-020',
        'C005',
        '渡辺健太',
        new Date(now.getTime() - 8 * 24 * 60 * 60 * 1000 + 96 * 60 * 60 * 1000),  // 96時間後
        9,
        '',
        '',
        '',
        '内定後'
      ]
    ];

    // データを一括書き込み
    if (testData.length > 0) {
      const startRow = responseSheet.getLastRow() + 1;
      responseSheet.getRange(startRow, 1, testData.length, 9).setValues(testData);
    }

    Logger.log(`✅ Survey_Responseに${testData.length}件のテストデータを投入しました`);
    return testData.length;

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

      // データを走査してテストデータの行番号を収集
      for (let i = 1; i < sendLogData.length; i++) {
        const sendId = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.SEND_ID]; // ✅ LOG_ID → SEND_ID に修正
        if (sendId && sendId.toString().startsWith('LOG-TEST-')) {
          rowsToDelete.push(i + 1); // 行番号は1始まり
        }
      }

      // 後ろの行から削除（行番号のズレを防ぐ）
      rowsToDelete.sort((a, b) => b - a); // 降順ソート
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

      // データを走査してテストデータの行番号を収集
      for (let i = 1; i < responseData.length; i++) {
        const responseId = responseData[i][CONFIG.COLUMNS.SURVEY_RESPONSE.RESPONSE_ID];
        if (responseId && responseId.toString().startsWith('RESP-TEST-')) {
          rowsToDelete.push(i + 1);
        }
      }

      // 後ろの行から削除（行番号のズレを防ぐ）
      rowsToDelete.sort((a, b) => b - a); // 降順ソート
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

      // データを走査してテストデータの行番号を収集
      for (let i = 1; i < analysisData.length; i++) {
        const analysisId = analysisData[i][CONFIG.COLUMNS.SURVEY_ANALYSIS.ANALYSIS_ID];
        if (analysisId && analysisId.toString().includes('TEST')) {
          rowsToDelete.push(i + 1);
        }
      }

      // 後ろの行から削除（行番号のズレを防ぐ）
      rowsToDelete.sort((a, b) => b - a); // 降順ソート
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
        const sendId = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.SEND_ID]; // ✅ LOG_ID → SEND_ID に修正
        const status = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.STATUS];

        if (sendId && sendId.toString().startsWith('LOG-TEST-')) {
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
 * テストデータ構造を詳細に検証（配列インデックスレベルでチェック）
 *
 * 【検証内容】
 * - Survey_Send_Log: E列（phase）にメールアドレスが入っていないかチェック
 * - Survey_Response: I列（アンケート種別）にメールアドレスが入っていないかチェック
 * - 各列の期待されるデータ型と実際のデータが一致するかチェック
 */
function validateTestDataStructure() {
  try {
    Logger.log('🔍 テストデータ構造の詳細検証を開始します...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const errors = [];
    let successCount = 0;

    // ========== Survey_Send_Log の検証 ==========
    const sendLogSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);
    if (sendLogSheet) {
      const sendLogData = sendLogSheet.getDataRange().getValues();
      let sendLogTestCount = 0;

      for (let i = 1; i < sendLogData.length; i++) {
        const row = sendLogData[i];
        const sendId = row[0];  // A: send_id

        // テストデータのみチェック
        if (sendId && sendId.toString().startsWith('LOG-TEST-')) {
          sendLogTestCount++;

          const email = row[3];   // D: email [3]
          const phase = row[4];   // E: phase [4]

          // E列（phase）にメールアドレスが入っていないかチェック
          if (phase && phase.toString().includes('@')) {
            errors.push(
              `❌ Survey_Send_Log 行${i + 1}: E列（phase）にメールアドレスが入っています: "${phase}"`
            );
          }

          // E列（phase）が正しいフェーズ名かチェック
          const validPhases = ['初回面談', '社員面談', '2次面接', '内定後'];
          if (phase && !validPhases.includes(phase)) {
            errors.push(
              `❌ Survey_Send_Log 行${i + 1}: E列（phase）の値が不正です: "${phase}"`
            );
          }

          // D列（email）にメールアドレスが入っているかチェック
          if (email && !email.toString().includes('@')) {
            errors.push(
              `❌ Survey_Send_Log 行${i + 1}: D列（email）にメールアドレスが入っていません: "${email}"`
            );
          }

          if (errors.length === 0) {
            successCount++;
          }
        }
      }

      Logger.log(`Survey_Send_Log: ${sendLogTestCount}件のテストデータをチェックしました`);
    }

    // ========== Survey_Response の検証 ==========
    const responseSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_RESPONSE);
    if (responseSheet) {
      const responseData = responseSheet.getDataRange().getValues();
      let responseTestCount = 0;

      for (let i = 1; i < responseData.length; i++) {
        const row = responseData[i];
        const responseId = row[0];  // A: response_id

        // テストデータのみチェック
        if (responseId && responseId.toString().startsWith('RESP-TEST-')) {
          responseTestCount++;

          const phase = row[8];  // I: アンケート種別 [8]

          // I列（アンケート種別）にメールアドレスが入っていないかチェック
          if (phase && phase.toString().includes('@')) {
            errors.push(
              `❌ Survey_Response 行${i + 1}: I列（アンケート種別）にメールアドレスが入っています: "${phase}"`
            );
          }

          // I列（アンケート種別）が正しいフェーズ名かチェック
          const validPhases = ['初回面談', '社員面談', '2次面接', '内定後'];
          if (phase && !validPhases.includes(phase)) {
            errors.push(
              `❌ Survey_Response 行${i + 1}: I列（アンケート種別）の値が不正です: "${phase}"`
            );
          }

          // 列数チェック（9列であること）
          const nonEmptyCount = row.filter(cell => cell !== '').length;
          if (nonEmptyCount !== 9) {
            errors.push(
              `❌ Survey_Response 行${i + 1}: 列数が不正です（期待: 9列, 実際: ${nonEmptyCount}列）`
            );
          }

          if (errors.length === 0) {
            successCount++;
          }
        }
      }

      Logger.log(`Survey_Response: ${responseTestCount}件のテストデータをチェックしました`);
    }

    // ========== 結果をダイアログ表示 ==========
    let message = '【🔍 テストデータ構造検証結果】\n\n';

    if (errors.length === 0) {
      message += '✅ 全てのテストデータが正しい構造です\n\n';
      message += `検証件数: ${successCount}件\n\n`;
      message += '━━━━━━━━━━━━━━━━\n';
      message += '【検証内容】\n';
      message += '✓ Survey_Send_Log E列にphaseが正しく入っている\n';
      message += '✓ Survey_Response I列にアンケート種別が正しく入っている\n';
      message += '✓ メールアドレスが誤った列に入っていない\n';
      message += '✓ 各列のデータ型が正しい\n\n';
      message += '次のステップ:\n';
      message += '→「📈 回答速度を一括計算」を実行してください';

      Logger.log('✅ テストデータ構造検証: 問題なし');
    } else {
      message += `❌ ${errors.length}件のエラーが見つかりました\n\n`;
      message += '━━━━━━━━━━━━━━━━\n';
      message += 'エラー詳細:\n';
      message += '━━━━━━━━━━━━━━━━\n';

      errors.slice(0, 10).forEach(error => {
        message += `${error}\n`;
        Logger.log(error);
      });

      if (errors.length > 10) {
        message += `\n... 他 ${errors.length - 10}件のエラー\n`;
      }

      message += '\n━━━━━━━━━━━━━━━━\n';
      message += '対処方法:\n';
      message += '1. テストデータをクリア\n';
      message += '2. TestDataGenerator.gsを確認\n';
      message += '3. 再度テストデータを生成';
    }

    SpreadsheetApp.getUi().alert(
      'テストデータ構造検証',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return errors.length === 0;

  } catch (error) {
    Logger.log(`❌ validateTestDataStructureエラー: ${error.message}`);
    Logger.log(error.stack);

    SpreadsheetApp.getUi().alert(
      'エラー',
      `検証中にエラーが発生しました:\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return false;
  }
}

/**
 * 全テストデータのバリデーションを実行（列数チェック）
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
