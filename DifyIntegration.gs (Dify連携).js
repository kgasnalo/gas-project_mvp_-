/**
 * DifyIntegration.gs
 * Difyとの連携機能を管理
 */

/**
 * Webhook URLを取得（デプロイ後に設定）
 *
 * 【デプロイ手順】
 * 1. GASエディタで「デプロイ」→「新しいデプロイ」
 * 2. 種類を「ウェブアプリ」に設定
 * 3. 「アクセスできるユーザー」を「全員」に設定
 * 4. デプロイ後、URLをDifyのWebhook設定に貼り付け
 */
function getWebhookUrl() {
  const deploymentUrl = ScriptApp.getService().getUrl();
  Logger.log('Webhook URL: ' + deploymentUrl);
  return deploymentUrl;
}


/**
 * 評価データを処理
 */
function handleEvaluationData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Evaluation_Log');

  // TODO: Dify環境構築後、実際のデータ形式に合わせて実装
  Logger.log('Processing evaluation data for: ' + data.candidate_id);

  // データを追記（サンプル）
  // sheet.appendRow([...]);
}

/**
 * 承諾可能性データを処理
 */
function handleEngagementData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Engagement_Log');

  // TODO: Dify環境構築後、実際のデータ形式に合わせて実装
  Logger.log('Processing engagement data for: ' + data.candidate_id);

  // データを追記（サンプル）
  // sheet.appendRow([...]);
}

/**
 * エビデンスデータを処理
 */
function handleEvidenceData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Evidence');

  // TODO: 実装
  Logger.log('Processing evidence data for: ' + data.candidate_id);
}

/**
 * リスクデータを処理
 */
function handleRiskData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Risk');

  // TODO: 実装
  Logger.log('Processing risk data for: ' + data.candidate_id);
}

/**
 * 次回質問データを処理
 */
function handleNextQData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('NextQ');

  // TODO: 実装
  Logger.log('Processing next_q data for: ' + data.candidate_id);
}

/**
 * 承諾ストーリーデータを処理
 */
function handleAcceptanceStoryData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Acceptance_Story');

  // TODO: 実装
  Logger.log('Processing acceptance_story data for: ' + data.candidate_id);
}

/**
 * 競合比較データを処理
 */
function handleCompetitorComparisonData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Competitor_Comparison');

  // TODO: 実装
  Logger.log('Processing competitor_comparison data for: ' + data.candidate_id);
}

/**
 * Dify APIを呼び出す（承諾可能性予測）
 *
 * @param {string} candidateId - 候補者ID
 * @return {Object} APIレスポンス
 */
function callDifyAPI(candidateId) {
  try {
    // API設定を取得
    const apiUrl = PropertiesService.getScriptProperties()
      .getProperty('DIFY_API_URL');
    const apiKey = PropertiesService.getScriptProperties()
      .getProperty('DIFY_API_KEY');

    if (!apiUrl || !apiKey) {
      throw new Error('Dify API settings not configured. Please run setupDifyApiSettings() first.');
    }

    // 候補者データを取得
    const candidateData = getCandidateData(candidateId);

    Logger.log('Calling Dify API for candidate: ' + candidateId);

    // APIリクエストを送信
    const options = {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        inputs: candidateData,
        response_mode: 'blocking',
        user: 'gas-system'
      }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    Logger.log('Dify API response code: ' + responseCode);
    Logger.log('Dify API response body: ' + responseBody);

    if (responseCode !== 200) {
      throw new Error('Dify API returned error: ' + responseCode + ' - ' + responseBody);
    }

    // レスポンスをパース
    const result = JSON.parse(responseBody);

    // TODO: Dify環境構築後、実際のレスポンス形式に合わせて処理

    return result;

  } catch (error) {
    Logger.log('Error in callDifyAPI: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}

/**
 * 候補者データを取得（Dify APIに送信するデータを構築）
 *
 * @param {string} candidateId - 候補者ID
 * @return {Object} 候補者データ
 */
function getCandidateData(candidateId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('Candidates_Master');

  // 候補者のデータを検索
  const data = masterSheet.getDataRange().getValues();
  const candidateRow = data.find(row => row[0] === candidateId);

  if (!candidateRow) {
    throw new Error('Candidate not found: ' + candidateId);
  }

  // TODO: Dify環境構築後、必要なデータ形式に合わせて構築
  const candidateData = {
    candidate_id: candidateRow[0],
    name: candidateRow[1],
    status: candidateRow[2],
    // ... 必要なデータを追加
  };

  Logger.log('Candidate data: ' + JSON.stringify(candidateData));

  return candidateData;
}

/**
 * Dify API設定をセットアップ
 *
 * 【使用方法】
 * 1. GASエディタでこの関数を実行
 * 2. API URLとAPI Keyを入力
 * 3. スクリプトプロパティに保存される
 */
function setupDifyApiSettings() {
  const ui = SpreadsheetApp.getUi();

  // API URLを入力
  const apiUrlResponse = ui.prompt(
    'Dify API URL',
    'Dify APIのエンドポイントURLを入力してください\n例: https://api.dify.ai/v1/workflows/run',
    ui.ButtonSet.OK_CANCEL
  );

  if (apiUrlResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('キャンセルされました');
    return;
  }

  const apiUrl = apiUrlResponse.getResponseText();

  // API Keyを入力
  const apiKeyResponse = ui.prompt(
    'Dify API Key',
    'Dify APIのAPI Keyを入力してください',
    ui.ButtonSet.OK_CANCEL
  );

  if (apiKeyResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('キャンセルされました');
    return;
  }

  const apiKey = apiKeyResponse.getResponseText();

  // スクリプトプロパティに保存
  const props = PropertiesService.getScriptProperties();
  props.setProperty('DIFY_API_URL', apiUrl);
  props.setProperty('DIFY_API_KEY', apiKey);

  ui.alert(
    '✅ Dify API設定が完了しました\n\n' +
    'API URL: ' + apiUrl + '\n' +
    'API Key: ' + apiKey.substring(0, 10) + '...'
  );

  Logger.log('Dify API settings saved successfully');
}

/**
 * Webhook URLを表示
 */
function showWebhookUrl() {
  const url = getWebhookUrl();
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    '📡 Webhook URL\n\n' +
    url + '\n\n' +
    'このURLをDifyのWebhook設定に貼り付けてください'
  );
}

/**
 * Webhook受信テスト
 */
function testWebhook() {
  // テストデータ
  const testData = {
    type: 'evaluation',
    candidate_id: 'C001',
    data: {
      overall_score: 0.85,
      philosophy_score: 0.90
    }
  };

  // doPost()をシミュレート
  const e = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };

  const response = doPost(e);
  Logger.log('Test webhook response: ' + response.getContent());

  SpreadsheetApp.getUi().alert(
    '✅ Webhookテストが完了しました\n\n' +
    '詳細はログを確認してください（表示 → ログ）'
  );
}

/**
 * Evaluation_Masterシートにデータを書き込む
 */
function writeToEvaluationMaster(data) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Evaluation_Master');

  if (!sheet) {
    throw new Error('Evaluation_Master sheet not found');
  }

  // 評価IDの生成
  const evaluationId = generateEvaluationId();

  // 行データの組み立て
  const row = [
    evaluationId,                           // A: 評価ID
    data.interview_datetime || '',          // B: 面接日時
    data.candidate_id || '',                // C: 候補者ID
    data.candidate_name || '',              // D: 候補者氏名
    data.recruit_type || '',                // E: 採用区分
    data.selection_phase || '',             // F: 選考フェーズ
    data.dify_report_url || '',             // G: ドキュメントURL

    // Philosophy
    data.philosophy_rank || '',             // H
    data.philosophy_score || 0,             // I
    data.philosophy_reason || '',           // J

    // Strategy
    data.strategy_rank || '',               // K
    data.strategy_score || 0,               // L
    data.strategy_reason || '',             // M

    // Motivation
    data.motivation_rank || '',             // N
    data.motivation_score || 0,             // O
    data.motivation_reason || '',           // P

    // Execution
    data.execution_rank || '',              // Q
    data.execution_score || 0,              // R
    data.execution_reason || '',            // S

    // 総合
    data.total_score || 0,                  // T
    data.total_rank || '',                  // U
    data.summary || '',                     // V

    // その他
    data.transcript || '',                  // W
    data.interview_memo || '',              // X
    data.concerns || '',                    // Y
    data.next_check_points || '',           // Z
    new Date(),                             // AA
    data.workflow_id || '',                 // AB

    // Phase A追加: 次回質問（AC-AG列）
    data.next_question_1 || '',             // AC
    data.next_question_2 || '',             // AD
    data.next_question_3 || '',             // AE
    data.next_question_4 || '',             // AF
    data.next_question_5 || '',             // AG
    data.competitor_analysis || '',         // AH
    data.evaluation_report_url || '',       // AI
    data.strategy_report_url || ''          // AJ
  ];

  // データを追加
  sheet.appendRow(row);

  Logger.log(`Evaluation data written: ${evaluationId}`);
  return evaluationId;
}

/**
 * 評価ID生成関数
 * 形式: EVAL_YYYYMMDD_NNN
 */
function generateEvaluationId() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Evaluation_Master');

  // 今日の評価数をカウント
  const data = sheet.getDataRange().getValues();
  const todayPrefix = `EVAL_${dateStr}_`;
  const todayCount = data.filter(row =>
    String(row[0]).startsWith(todayPrefix)
  ).length;

  const sequence = String(todayCount + 1).padStart(3, '0');
  return `${todayPrefix}${sequence}`;
}

/**
 * Processing_Log確認関数
 * Phase 1-1のテスト実行結果を確認
 */
function checkProcessingLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Processing_Log');

  if (!sheet) {
    Logger.log('❌ Processing_Logシートが見つかりません');
    SpreadsheetApp.getUi().alert(
      '❌ エラー\n\n' +
      'Processing_Logシートが見つかりません。\n' +
      'シート名を確認してください。'
    );
    return;
  }

  // 最新行（最後のデータ行）を取得
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    Logger.log('❌ データが記録されていません（ヘッダー行のみ）');
    SpreadsheetApp.getUi().alert(
      '❌ データなし\n\n' +
      'Processing_Logにデータが記録されていません。\n' +
      'Difyからテスト実行を行ってください。'
    );
    return;
  }

  // 最新行のデータを取得（8列分）
  const data = sheet.getRange(lastRow, 1, 1, 8).getValues()[0];

  Logger.log('=== Processing_Log 最新行 ===');
  Logger.log('行番号: ' + lastRow);
  Logger.log('A列 (timestamp): ' + data[0]);
  Logger.log('B列 (phase): ' + data[1]);
  Logger.log('C列 (candidate): ' + data[2]);
  Logger.log('D列 (event): ' + data[3]);
  Logger.log('E列 (status): ' + data[4]);
  Logger.log('F列 (input_data): ' + data[5]);
  Logger.log('G列 (output_data): ' + data[6]);
  Logger.log('H列 (notes): ' + data[7]);

  // 期待値との比較
  const expectedPhase = 'Phase1-1_Test';
  const expectedCandidate = 'テスト太郎';
  const expectedStatus = 'SUCCESS';

  let isValid = true;
  let errorMessages = [];

  if (data[1] !== expectedPhase) {
    Logger.log('⚠️ Phase が期待値と異なります: ' + data[1]);
    errorMessages.push('Phase: ' + data[1] + ' (期待値: ' + expectedPhase + ')');
    isValid = false;
  }

  if (data[2] !== expectedCandidate) {
    Logger.log('⚠️ Candidate が期待値と異なります: ' + data[2]);
    errorMessages.push('Candidate: ' + data[2] + ' (期待値: ' + expectedCandidate + ')');
    isValid = false;
  }

  if (data[4] !== expectedStatus) {
    Logger.log('⚠️ Status が期待値と異なります: ' + data[4]);
    errorMessages.push('Status: ' + data[4] + ' (期待値: ' + expectedStatus + ')');
    isValid = false;
  }

  // 結果をUIに表示（UIコンテキストがある場合のみ）
  if (isValid) {
    Logger.log('✅ Phase 1-1 完全成功！');
    Logger.log('✅ Processing_Logに正しく記録されています');

    try {
      SpreadsheetApp.getUi().alert(
        '🎉 Phase 1-1 完全成功！\n\n' +
        '【記録内容】\n' +
        '行番号: ' + lastRow + '\n' +
        'タイムスタンプ: ' + data[0] + '\n' +
        'Phase: ' + data[1] + '\n' +
        '候補者: ' + data[2] + '\n' +
        'イベント: ' + data[3] + '\n' +
        'ステータス: ' + data[4] + '\n' +
        '実行時間: ' + data[7] + '\n\n' +
        '✅ Processing_Logに正しく記録されています\n' +
        '✅ Dify → GAS データフロー確立完了'
      );
    } catch (e) {
      // UIコンテキストがない場合はログのみ
      Logger.log('ℹ️ UIアラートをスキップ（UIコンテキストなし）');
    }
  } else {
    Logger.log('⚠️ 一部データが期待値と異なります');

    try {
      SpreadsheetApp.getUi().alert(
        '⚠️ データ不一致\n\n' +
        '一部データが期待値と異なります:\n\n' +
        errorMessages.join('\n') + '\n\n' +
        '詳細はログを確認してください。\n' +
        '（表示 → ログ）'
      );
    } catch (e) {
      // UIコンテキストがない場合はログのみ
      Logger.log('ℹ️ UIアラートをスキップ（UIコンテキストなし）');
    }
  }

  return isValid;
}

// ========================================
// Phase 1-2: シート構造確認
// ========================================

/**
 * シート構造確認用関数
 * 各シートのヘッダー行を確認し、データマッピングを確定
 */
function checkSheetStructures() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = [
    'Dify_Workflow_Log',
    'Engagement_Log',
    'Candidates_Master',
    'Candidate_Scores',
    'Candidate_Insights'
  ];

  Logger.log('========================================');
  Logger.log('シート構造確認開始');
  Logger.log('========================================\n');

  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`❌ ${sheetName} が見つかりません\n`);
      return;
    }

    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
      Logger.log(`⚠️ ${sheetName} にデータがありません\n`);
      return;
    }

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    Logger.log(`=== ${sheetName} ===`);
    Logger.log(`列数: ${headers.length}`);
    headers.forEach((header, index) => {
      Logger.log(`  ${String.fromCharCode(65 + index)}列: ${header}`);
    });
    Logger.log('');
  });

  Logger.log('========================================');
  Logger.log('シート構造確認完了');
  Logger.log('========================================');
}

// ========================================
// Phase 1-2: コア関数（5つ）
// ========================================

/**
 * Dify_Workflow_Logシートにワークフロー実行ログを追記
 *
 * @param {Object} data - ワークフローログデータ
 * @return {string} - ログID
 */
function appendToDifyWorkflowLog(data) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Dify_Workflow_Log');

  if (!sheet) {
    throw new Error('Dify_Workflow_Logシートが見つかりません');
  }

  // ログIDの生成
  const logId = data.workflow_id || `WF_${new Date().getTime()}`;

  // 行データの組み立て
  const row = [
    logId,                                    // A: workflow_log_id
    data.workflow_name || 'Phase1_Workflow',  // B: workflow_name
    data.candidate_id || '',                  // C: candidate_id
    data.execution_date || new Date(),        // D: execution_date
    data.status || 'SUCCESS',                 // E: status
    data.duration_seconds || data.execution_time_seconds || 0, // F: duration_seconds
    data.input_summary || JSON.stringify(data.input || {}).substring(0, 200), // G: input_summary
    data.output_summary || JSON.stringify(data.output || {}).substring(0, 200), // H: output_summary
    data.error_message || ''                  // I: error_message
  ];

  // データ追加
  sheet.appendRow(row);

  Logger.log(`✅ Dify_Workflow_Log追記完了: ${logId}`);
  return logId;
}

/**
 * Engagement_Logシートにエンゲージメントログを追記
 *
 * @param {Object} data - エンゲージメントデータ
 * @return {string} - ログID
 */
function appendToEngagementLog(data) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Engagement_Log');

  if (!sheet) {
    throw new Error('Engagement_Logシートが見つかりません');
  }

  // ログIDの生成
  const logId = data.log_id || `LOG_${new Date().getTime()}`;

  // 行データの組み立て（シート構造に合わせて調整が必要な場合があります）
  const row = [
    logId,                                    // A: log_id
    data.candidate_id || '',                  // B: candidate_id
    data.candidate_name || data['氏名'] || '', // C: 氏名
    data.timestamp || new Date(),             // D: timestamp
    data.contact_type || '',                  // E: contact_type
    data.acceptance_rate_ai || 0,             // F: acceptance_rate_ai
    data.confidence_level || '',              // G: confidence_level
    data.motivation_score || 0                // H: motivation_score
  ];

  // 追加のフィールド（シート構造により異なる）
  if (sheet.getLastColumn() > 8) {
    row.push(
      data.core_motivation || '',             // I: core_motivation
      data.top_concern || '',                 // J: top_concern
      data.competitors || '',                 // K: competitors
      data.next_action || '',                 // L: next_action
      data.action_priority || '',             // M: action_priority
      data.action_deadline || '',             // N: action_deadline
      data.action_status || '未実施',         // O: action_status
      data.action_result || '',               // P: action_result
      data.notes || ''                        // Q: notes
    );
  }

  // データ追加
  sheet.appendRow(row);

  Logger.log(`✅ Engagement_Log追記完了: ${logId}`);
  return logId;
}

/**
 * Candidates_Masterシートのデータを更新または追加
 * candidate_idで検索し、存在すれば更新、なければ追加
 *
 * @param {Object} data - 候補者マスターデータ
 * @return {string} - 'UPDATED' or 'INSERTED'
 */
function updateOrInsertCandidatesMaster(data) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Candidates_Master');

  if (!sheet) {
    throw new Error('Candidates_Masterシートが見つかりません');
  }

  const candidateId = data.candidate_id;
  if (!candidateId) {
    throw new Error('candidate_idが必須です');
  }

  // 既存データを検索
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];

  // candidate_id列のインデックスを取得
  const candidateIdColIndex = headers.indexOf('candidate_id');
  if (candidateIdColIndex === -1) {
    throw new Error('candidate_id列が見つかりません');
  }

  // 既存行を検索
  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][candidateIdColIndex] === candidateId) {
      targetRow = i + 1; // 1-indexed
      break;
    }
  }

  // ヘッダーに基づいて行データを構築
  const row = [];
  headers.forEach(header => {
    if (header === 'candidate_id') {
      row.push(candidateId);
    } else if (header === '最終更新日時') {
      row.push(new Date());
    } else {
      // データオブジェクトから対応する値を取得
      row.push(data[header] || '');
    }
  });

  if (targetRow > 0) {
    // 更新
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    Logger.log(`✅ Candidates_Master更新完了: ${candidateId} (行${targetRow})`);
    return 'UPDATED';
  } else {
    // 追加
    sheet.appendRow(row);
    Logger.log(`✅ Candidates_Master追加完了: ${candidateId}`);
    return 'INSERTED';
  }
}

/**
 * Candidate_Scoresシートにスコアデータを追記
 *
 * @param {Object} data - スコアデータ
 * @return {number} - 追加された行番号
 */
function appendToCandidateScores(data) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Candidate_Scores');

  if (!sheet) {
    throw new Error('Candidate_Scoresシートが見つかりません');
  }

  // ヘッダー取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // ヘッダーに基づいて行データを構築
  const row = [];
  headers.forEach(header => {
    if (header === 'candidate_id') {
      row.push(data.candidate_id || '');
    } else if (header === '氏名') {
      row.push(data['氏名'] || data.candidate_name || '');
    } else if (header === '最終更新日時') {
      row.push(new Date());
    } else {
      row.push(data[header] || '');
    }
  });

  // データ追加
  sheet.appendRow(row);

  const lastRow = sheet.getLastRow();
  Logger.log(`✅ Candidate_Scores追記完了: ${data.candidate_id} (行${lastRow})`);
  return lastRow;
}

/**
 * Candidate_Insightsシートのデータを更新または追加
 * candidate_idで検索し、存在すれば更新、なければ追加
 *
 * @param {Object} data - インサイトデータ
 * @return {string} - 'UPDATED' or 'INSERTED'
 */
function updateOrInsertCandidateInsights(data) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Candidate_Insights');

  if (!sheet) {
    throw new Error('Candidate_Insightsシートが見つかりません');
  }

  const candidateId = data.candidate_id;
  if (!candidateId) {
    throw new Error('candidate_idが必須です');
  }

  // 既存データを検索
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];

  const candidateIdColIndex = headers.indexOf('candidate_id');
  if (candidateIdColIndex === -1) {
    throw new Error('candidate_id列が見つかりません');
  }

  // 既存行を検索
  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][candidateIdColIndex] === candidateId) {
      targetRow = i + 1;
      break;
    }
  }

  // ヘッダーに基づいて行データを構築
  const row = [];
  headers.forEach(header => {
    if (header === 'candidate_id') {
      row.push(candidateId);
    } else if (header === '最終更新日時') {
      row.push(new Date());
    } else {
      row.push(data[header] || '');
    }
  });

  if (targetRow > 0) {
    // 更新
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    Logger.log(`✅ Candidate_Insights更新完了: ${candidateId} (行${targetRow})`);
    return 'UPDATED';
  } else {
    // 追加
    sheet.appendRow(row);
    Logger.log(`✅ Candidate_Insights追加完了: ${candidateId}`);
    return 'INSERTED';
  }
}

// ========================================
// Phase 1 本番版: doPost関数
// ========================================

/**
 * Phase 1 本番版: Dify Webhookエンドポイント
 * テストモードと本番モードの両対応
 */
function doPost(e) {
  const startTime = new Date();

  try {
    const requestBody = e.postData ? e.postData.contents : null;

    if (!requestBody) {
      throw new Error('リクエストボディが空です');
    }

    const data = JSON.parse(requestBody);

    // ===== 行動データ取得エンドポイント =====
    if (data.action === 'get_behavior_data') {
      return handleGetBehaviorData(data);
    }

    Logger.log('=== Dify Webhook受信 ===');
    Logger.log('Mode: ' + (data.test_mode ? 'TEST' : 'PRODUCTION'));
    Logger.log('Candidate: ' + (data.validated_input ? data.validated_input.candidate_name : 'Unknown'));

    // テストモードの場合
    if (data.test_mode === true) {
      return handleTestMode(data, startTime);
    }

    // ===== 本番モード: データ保存処理 =====

    const results = {
      candidates_master: null,
      candidate_scores: null,
      candidate_insights: null,
      engagement_log: null,
      evaluation_master: null,
      workflow_log: null
    };

    // 1. Candidates_Master更新
    if (data.candidates_master) {
      const masterData = typeof data.candidates_master === 'string'
        ? JSON.parse(data.candidates_master)
        : data.candidates_master;
      results.candidates_master = updateOrInsertCandidatesMaster(masterData);
      Logger.log('✅ Candidates_Master: ' + results.candidates_master);
    }

    // 2. Candidate_Scores追加
    if (data.candidate_scores) {
      const scoresData = typeof data.candidate_scores === 'string'
        ? JSON.parse(data.candidate_scores)
        : data.candidate_scores;
      results.candidate_scores = appendToCandidateScores(scoresData);
      Logger.log('✅ Candidate_Scores: 行' + results.candidate_scores);
    }

    // 3. Candidate_Insights更新
    if (data.candidate_insights) {
      const insightsData = typeof data.candidate_insights === 'string'
        ? JSON.parse(data.candidate_insights)
        : data.candidate_insights;
      results.candidate_insights = updateOrInsertCandidateInsights(insightsData);
      Logger.log('✅ Candidate_Insights: ' + results.candidate_insights);
    }

    // 4. Engagement_Log追加
    if (data.engagement_log) {
      const engagementData = typeof data.engagement_log === 'string'
        ? JSON.parse(data.engagement_log)
        : data.engagement_log;
      results.engagement_log = appendToEngagementLog(engagementData);
      Logger.log('✅ Engagement_Log: ' + results.engagement_log);
    }

    // 5. Evaluation_Master追加
    let evaluationMasterData = null;
    if (data.evaluation_master) {
      evaluationMasterData = typeof data.evaluation_master === 'string'
        ? JSON.parse(data.evaluation_master)
        : data.evaluation_master;
      results.evaluation_master = writeToEvaluationMaster(evaluationMasterData);
      Logger.log('✅ Evaluation_Master: ' + results.evaluation_master);
    }

    // === レポート生成処理（V2版） ===
    Logger.log('=== レポート生成V2開始 ===');

    try {
      // 企業名取得（スクリプトプロパティまたはデフォルト）
      const companyName = PropertiesService.getScriptProperties().getProperty('COMPANY_NAME') || 'アマネク';

      // 候補者マスターデータの取得
      const candidateMasterData = typeof data.candidates_master === 'string'
        ? JSON.parse(data.candidates_master)
        : data.candidates_master;

      // 承諾可能性データの取得
      const acceptanceData = typeof data.engagement_log === 'string'
        ? JSON.parse(data.engagement_log)
        : data.engagement_log;

      // 評価マスターデータの取得（既にパース済み）
      const evalData = evaluationMasterData || {};

      // 採用区分と選考フェーズの取得
      const recruitType = candidateMasterData['採用区分'] || '新卒';
      const selectionPhase = candidateMasterData['現在ステータス'] || (data.validated_input ? data.validated_input.selection_phase : '1次面接');

      // スプレッドシートURL取得
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const spreadsheetUrl = ss.getUrl();

      // 1. 評価レポート生成（V2）
      if (candidateMasterData && evalData && data.validated_input) {
        Logger.log('--- 評価レポートV2 生成 ---');

        const evalReportDataV2 = {
          // 基本情報
          candidate_id: data.validated_input.candidate_id,
          candidate_name: candidateMasterData['氏名'],
          selection_phase: selectionPhase,
          interview_date: evalData.interview_datetime || Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd'),
          interviewer: data.validated_input.interviewer,

          // 総合評価
          total_rank: evalData.total_rank,
          recommendation: getRecommendation(evalData.total_rank),
          summary_reasons: [
            evalData.philosophy_reason ? '理念への共感が見られる' : null,
            evalData.strategy_reason ? '戦略的思考力がある' : null,
            evalData.motivation_reason ? '高い志望度が見られる' : null
          ].filter(r => r),

          // 4軸評価
          philosophy_rank: evalData.philosophy_rank,
          philosophy_score: evalData.philosophy_score,
          philosophy_summary: getSummary(evalData.philosophy_rank, 'Philosophy'),
          philosophy_reason: evalData.philosophy_reason,
          philosophy_evidence: evalData.philosophy_evidence || '',

          strategy_rank: evalData.strategy_rank,
          strategy_score: evalData.strategy_score,
          strategy_summary: getSummary(evalData.strategy_rank, 'Strategy'),
          strategy_reason: evalData.strategy_reason,
          strategy_evidence: evalData.strategy_evidence || '',

          motivation_rank: evalData.motivation_rank,
          motivation_score: evalData.motivation_score,
          motivation_summary: getSummary(evalData.motivation_rank, 'Motivation'),
          motivation_reason: evalData.motivation_reason,
          motivation_evidence: evalData.motivation_evidence || '',

          execution_rank: evalData.execution_rank,
          execution_score: evalData.execution_score,
          execution_summary: getSummary(evalData.execution_rank, 'Execution'),
          execution_reason: evalData.execution_reason,
          execution_evidence: evalData.execution_evidence || '',

          // 懸念事項
          critical_concerns: [],

          // 次回質問
          next_questions: [
            evalData.next_question_1,
            evalData.next_question_2,
            evalData.next_question_3,
            evalData.next_question_4,
            evalData.next_question_5
          ].filter(q => q),

          // 面接官コメント
          interviewer_comment: evalData.summary || '',

          // 議事録
          transcript: evalData.transcript || '（文字起こしデータなし）',

          // スプレッドシートリンク
          spreadsheet_url: spreadsheetUrl
        };

        const evalResultV2 = generateEvaluationReportV2(evalReportDataV2, recruitType, selectionPhase, companyName);

        if (evalResultV2.success) {
          Logger.log('✅ 評価レポートV2生成成功: ' + evalResultV2.url);
          results.evaluation_report_url = evalResultV2.url;

          // Evaluation_MasterにURL記録
          const evalSheet = ss.getSheetByName('Evaluation_Master');
          if (evalSheet) {
            const evalLastRow = evalSheet.getLastRow();
            evalSheet.getRange(evalLastRow, 32).setValue(evalResultV2.url);  // AF列 = 32
            Logger.log('✅ 評価レポートURL記録: AF列（行' + evalLastRow + '）');
          }
        }
      }

      // 2. 戦略レポート生成（V2）
      if (candidateMasterData && acceptanceData && data.validated_input) {
        Logger.log('--- 戦略レポートV2 生成 ---');

        const strategyReportDataV2 = {
          // 基本情報
          candidate_id: data.validated_input.candidate_id,
          candidate_name: candidateMasterData['氏名'],
          current_phase: selectionPhase,
          interviewer: data.validated_input.interviewer,

          // 承諾可能性
          acceptance_probability: acceptanceData.acceptance_rate_ai || 0,
          confidence_level: acceptanceData.confidence_level || 'MEDIUM',

          // 競合状況
          competitor_probabilities: [],

          // 24時間アクション
          immediate_action_24h: acceptanceData.next_action || '（アクションなし）',
          action_reason: '（理由不明）',
          expected_effect: '（効果不明）',

          // リスク要因
          risk_factors: (acceptanceData.key_risk_factors || []).map(factor => ({
            factor: factor,
            countermeasure: '（対策要検討）'
          })),

          // 自社の強み
          our_strengths: acceptanceData.key_positive_factors || [],

          // 承諾ストーリー
          acceptance_story: acceptanceData.engagement_strategy ? [
            `Step 1: ${acceptanceData.engagement_strategy.immediate_action_24h || '即時アクション'}`,
            `Step 2: ${acceptanceData.engagement_strategy.followup_action_48h || 'フォローアップ'}`,
            `Step 3: ${acceptanceData.engagement_strategy.longterm_action_72h || '長期施策'}`
          ].filter(s => s && !s.includes('undefined')) : [],

          // ポジティブ要因（詳細）
          positive_factors: (acceptanceData.key_positive_factors || []).map(factor => ({
            factor: factor,
            evidence: ''
          })),

          // リスク要因（詳細）
          risk_factors_detailed: (acceptanceData.key_risk_factors || []).map(factor => ({
            factor: factor,
            severity: 'MEDIUM',
            detailed_countermeasure: '（詳細対策要検討）'
          })),

          // 競合分析
          competitor_analysis: [],

          // 推奨施策
          engagement_recommendations: acceptanceData.engagement_strategy ? [
            `24時間以内: ${acceptanceData.engagement_strategy.immediate_action_24h || '要確認'}`,
            `48時間以内: ${acceptanceData.engagement_strategy.followup_action_48h || '要確認'}`,
            `72時間以内: ${acceptanceData.engagement_strategy.longterm_action_72h || '要確認'}`
          ].filter(s => s && !s.includes('undefined')) : [],

          // スプレッドシートリンク
          spreadsheet_url: spreadsheetUrl
        };

        const strategyResultV2 = generateStrategyReportV2(strategyReportDataV2, recruitType, selectionPhase, companyName);

        if (strategyResultV2.success) {
          Logger.log('✅ 戦略レポートV2生成成功: ' + strategyResultV2.url);
          results.strategy_report_url = strategyResultV2.url;

          // Evaluation_MasterにURL記録
          const evalSheet = ss.getSheetByName('Evaluation_Master');
          if (evalSheet) {
            const evalLastRow = evalSheet.getLastRow();
            evalSheet.getRange(evalLastRow, 33).setValue(strategyResultV2.url);  // AG列 = 33
            Logger.log('✅ 戦略レポートURL記録: AG列（行' + evalLastRow + '）');
          }
        }
      }

      // 3. 評価B以上の場合、フォルダコピー
      if (evalData && evalData.total_rank && ['A', 'B'].includes(evalData.total_rank)) {
        Logger.log('--- 評価B以上: フォルダコピー実行 ---');
        const copyResult = copyFolderToGradeB(
          data.validated_input.candidate_id,
          candidateMasterData['氏名'],
          recruitType,
          companyName
        );

        if (copyResult.success) {
          Logger.log('✅ 評価B以上フォルダコピー成功');
        } else {
          Logger.log('⚠️ 評価B以上フォルダコピー失敗: ' + (copyResult.error || 'Unknown error'));
        }
      }

      // 4. 内定・承諾の場合、フォルダ移動
      if (candidateMasterData && candidateMasterData['現在ステータス'] &&
          ['内定', '承諾'].includes(candidateMasterData['現在ステータス'])) {
        Logger.log('--- 内定・承諾: フォルダ移動実行 ---');
        const moveResult = moveFolderToAccepted(
          data.validated_input.candidate_id,
          candidateMasterData['氏名'],
          recruitType,
          selectionPhase,
          companyName
        );

        if (moveResult.success) {
          Logger.log('✅ 内定・承諾フォルダ移動成功');
        } else {
          Logger.log('⚠️ 内定・承諾フォルダ移動失敗: ' + (moveResult.error || 'Unknown error'));
        }
      }

      Logger.log('✅ レポート生成V2完了');

    } catch (reportError) {
      Logger.log('❌ レポート生成V2エラー: ' + reportError.message);
      Logger.log('スタックトレース: ' + reportError.stack);
      results.report_generation_error = reportError.message;
      // レポート生成エラーでも処理は継続
    }

    Logger.log('=== レポート生成V2完了 ===');

    // 6. Dify_Workflow_Log追加
    if (data.workflow_log) {
      const workflowData = typeof data.workflow_log === 'string'
        ? JSON.parse(data.workflow_log)
        : data.workflow_log;

      // 実行時間を追加
      workflowData.execution_time_seconds = ((new Date() - startTime) / 1000).toFixed(2);

      results.workflow_log = appendToDifyWorkflowLog(workflowData);
      Logger.log('✅ Dify_Workflow_Log: ' + results.workflow_log);
    }

    // 7. Processing_Log記録
    logProcessing({
      candidate_id: data.validated_input ? data.validated_input.candidate_id : 'unknown',
      candidate_name: data.validated_input ? data.validated_input.candidate_name : 'unknown',
      status: 'SUCCESS',
      phase: 'Phase1_Production',
      timestamp: new Date().toISOString(),
      execution_time: ((new Date() - startTime) / 1000).toFixed(2)
    });

    // 成功レスポンス
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        mode: 'PRODUCTION',
        message: 'Phase 1: データ保存完了',
        results: results,
        timestamp: new Date().toISOString(),
        execution_time_seconds: ((new Date() - startTime) / 1000).toFixed(2)
      }, null, 2))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('❌ エラー発生: ' + error.message);
    Logger.log('スタック: ' + error.stack);

    // エラーをProcessing_Logに記録
    logProcessing({
      candidate_id: 'error',
      candidate_name: 'error',
      status: 'FAILED',
      phase: 'Phase1_Production',
      timestamp: new Date().toISOString(),
      error: error.message
    });

    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        mode: 'PRODUCTION',
        error: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString()
      }, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ランクから推奨を取得（ヘルパー関数）
 */
function getRecommendation(rank) {
  const recommendations = {
    'A': '積極採用推奨',
    'B': '採用推奨',
    'C': '条件付き採用検討',
    'D': '慎重検討',
    'E': '見送り推奨'
  };
  return recommendations[rank] || '要検討';
}

/**
 * ランクからサマリーを取得（ヘルパー関数）
 */
function getSummary(rank, axis) {
  const summaries = {
    'Philosophy': {
      'A': '理念への深い共感が見られる',
      'B': '理念への一定の共感が見られる',
      'C': '理念理解は標準的',
      'D': '理念への共感がやや弱い',
      'E': '理念とのミスマッチが懸念される'
    },
    'Strategy': {
      'A': '優れた戦略的思考力',
      'B': '戦略理解は十分、実践経験で向上可',
      'C': '戦略理解は標準的',
      'D': '戦略的思考力の強化が必要',
      'E': '戦略的思考力に大きな課題'
    },
    'Motivation': {
      'A': '非常に高い志望度、成長意欲強',
      'B': '高い志望度が見られる',
      'C': '志望度は標準的',
      'D': '志望度がやや低い',
      'E': '志望度が低い、動機に懸念'
    },
    'Execution': {
      'A': '優れた実行力、実績あり',
      'B': '実行力は十分、実績あり',
      'C': '実行力は標準的',
      'D': '実行力の強化が必要',
      'E': '実行力に大きな課題'
    }
  };
  return summaries[axis]?.[rank] || '評価保留';
}

/**
 * テストモード処理（既存のdoPost関数内で使用）
 */
function handleTestMode(data, startTime) {
  Logger.log('テストモード実行');

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Processing_Log');

  if (sheet) {
    const logRow = [
      new Date(),
      'Phase1-1_Test',
      data.validated_input ? data.validated_input.candidate_name : 'Test',
      'webhook_test',
      'SUCCESS',
      JSON.stringify(data.validated_input || {}).substring(0, 500),
      'transcript: ' + (data.transcript ? data.transcript.length + '文字' : 'なし'),
      '実行時間: ' + ((new Date() - startTime) / 1000).toFixed(2) + '秒'
    ];
    sheet.appendRow(logRow);
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      mode: 'TEST_MODE',
      message: 'Phase 1-1: データ受信成功（テストモード）',
      received: {
        candidate_id: data.validated_input ? data.validated_input.candidate_id : null,
        candidate_name: data.validated_input ? data.validated_input.candidate_name : null,
        has_transcript: !!data.transcript,
        transcript_length: data.transcript ? data.transcript.length : 0
      },
      timestamp: new Date().toISOString(),
      execution_time_seconds: ((new Date() - startTime) / 1000).toFixed(2)
    }, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Processing_Log記録
 */
function logProcessing(logData) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Processing_Log');

  if (sheet) {
    sheet.appendRow([
      new Date(),
      logData.phase || 'Phase1',
      logData.candidate_name || logData.candidate_id || 'unknown',
      'workflow_execution',
      logData.status || 'SUCCESS',
      logData.message || '',
      logData.error || '',
      logData.execution_time ? '実行時間: ' + logData.execution_time + '秒' : logData.timestamp
    ]);
  }
}

// ========================================
// Phase 1-4: テスト・検証関数
// ========================================

/**
 * ワークフローデータの検証
 *
 * @param {Object} data - 検証するデータ
 * @return {Object} - {valid: boolean, errors: array}
 */
function validateWorkflowData(data) {
  const errors = [];

  // 必須フィールドのチェック
  if (!data.validated_input) {
    errors.push('validated_inputが存在しません');
  } else {
    if (!data.validated_input.candidate_id) {
      errors.push('candidate_idが存在しません');
    }
    if (!data.validated_input.candidate_name) {
      errors.push('candidate_nameが存在しません');
    }
  }

  // スコアデータのチェック
  if (data.candidate_scores) {
    const scores = typeof data.candidate_scores === 'string'
      ? JSON.parse(data.candidate_scores)
      : data.candidate_scores;

    const totalScore = (scores['最新_Philosophy'] || 0) +
                      (scores['最新_Strategy'] || 0) +
                      (scores['最新_Motivation'] || 0) +
                      (scores['最新_Execution'] || 0);

    if (totalScore !== scores['最新_合計スコア']) {
      errors.push(`スコア合計が不一致: 計算値${totalScore} vs 実値${scores['最新_合計スコア']}`);
    }
  }

  const isValid = errors.length === 0;
  Logger.log(isValid ? '✅ データ検証成功' : '⚠️ データ検証失敗: ' + errors.join(', '));

  return {
    valid: isValid,
    errors: errors
  };
}

/**
 * 最近のワークフローログを取得
 *
 * @param {number} n - 取得する件数（デフォルト: 10）
 * @return {Array} - ログデータの配列
 */
function getRecentWorkflowLogs(n = 10) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Dify_Workflow_Log');

  if (!sheet) {
    Logger.log('❌ Dify_Workflow_Logシートが見つかりません');
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('ℹ️ Dify_Workflow_Logにデータがありません');
    return [];
  }

  const startRow = Math.max(2, lastRow - n + 1);
  const numRows = lastRow - startRow + 1;

  const data = sheet.getRange(startRow, 1, numRows, 9).getValues();

  const logs = data.map(row => ({
    workflow_id: row[0],
    workflow_name: row[1],
    candidate_id: row[2],
    execution_date: row[3],
    status: row[4],
    duration_seconds: row[5],
    input_summary: row[6],
    output_summary: row[7],
    error_message: row[8]
  }));

  Logger.log(`✅ 最近のログ${logs.length}件を取得しました`);
  return logs;
}

/**
 * 統合テスト用関数
 * Phase 1の全機能を一度にテストします
 */
function testFullWorkflow() {
  Logger.log('========================================');
  Logger.log('Phase 1 統合テスト開始');
  Logger.log('========================================\n');

  // テストデータの作成
  const testCandidateId = 'TEST_' + new Date().getTime();
  const testData = {
    test_mode: false, // 本番モードでテスト
    validated_input: {
      candidate_id: testCandidateId,
      candidate_name: 'テスト統合_太郎',
      recruit_type: '新卒',
      selection_phase: '1次面接',
      interviewer: '統合テスト',
      timestamp: new Date().toISOString()
    },
    candidates_master: {
      candidate_id: testCandidateId,
      '氏名': 'テスト統合_太郎',
      '採用区分': '新卒',
      '現在ステータス': '1次面接',
      'メールアドレス': 'test@example.com'
    },
    candidate_scores: {
      candidate_id: testCandidateId,
      '氏名': 'テスト統合_太郎',
      '最新_Philosophy': 25,
      '最新_Strategy': 27,
      '最新_Motivation': 18,
      '最新_Execution': 16,
      '最新_合計スコア': 86,
      '最新_承諾可能性（AI予測）': 65,
      '予測の信頼度': 'HIGH'
    },
    candidate_insights: {
      candidate_id: testCandidateId,
      '氏名': 'テスト統合_太郎',
      'コアモチベーション': '社会貢献への強い意欲',
      '主要懸念事項': '待遇面での不安',
      '競合企業1': 'リブコンサルティング',
      '次推奨アクション': '待遇詳細の説明'
    },
    engagement_log: {
      log_id: 'TESTLOG_' + new Date().getTime(),
      candidate_id: testCandidateId,
      candidate_name: 'テスト統合_太郎',
      contact_type: '1次面接',
      acceptance_rate_ai: 65,
      confidence_level: 'HIGH',
      motivation_score: 18,
      core_motivation: '社会貢献への強い意欲'
    },
    evaluation_master: {
      candidate_id: testCandidateId,
      candidate_name: 'テスト統合_太郎',
      recruit_type: '新卒',
      selection_phase: '1次面接',
      philosophy_score: 25,
      strategy_score: 27,
      motivation_score: 18,
      execution_score: 16,
      total_score: 86,
      total_rank: 'A'
    },
    workflow_log: {
      workflow_id: 'TESTWF_' + new Date().getTime(),
      workflow_name: 'Phase1_IntegrationTest',
      candidate_id: testCandidateId,
      execution_date: new Date(),
      status: 'SUCCESS',
      phase: 'Phase1_IntegrationTest',
      duration_seconds: 3.5
    }
  };

  try {
    // 1. データ検証
    Logger.log('Step 1: データ検証');
    const validation = validateWorkflowData(testData);
    if (!validation.valid) {
      throw new Error('データ検証失敗: ' + validation.errors.join(', '));
    }

    // 2. 各関数を個別にテスト
    Logger.log('\nStep 2: 各関数の個別テスト');

    // 2-1. Dify_Workflow_Log
    Logger.log('  2-1. appendToDifyWorkflowLog()');
    const workflowLogId = appendToDifyWorkflowLog(testData.workflow_log);

    // 2-2. Engagement_Log
    Logger.log('  2-2. appendToEngagementLog()');
    const engagementLogId = appendToEngagementLog(testData.engagement_log);

    // 2-3. Candidates_Master
    Logger.log('  2-3. updateOrInsertCandidatesMaster()');
    const masterResult = updateOrInsertCandidatesMaster(testData.candidates_master);

    // 2-4. Candidate_Scores
    Logger.log('  2-4. appendToCandidateScores()');
    const scoresRow = appendToCandidateScores(testData.candidate_scores);

    // 2-5. Candidate_Insights
    Logger.log('  2-5. updateOrInsertCandidateInsights()');
    const insightsResult = updateOrInsertCandidateInsights(testData.candidate_insights);

    // 2-6. Evaluation_Master
    Logger.log('  2-6. writeToEvaluationMaster()');
    const evalId = writeToEvaluationMaster(testData.evaluation_master);

    // 3. 結果サマリー
    Logger.log('\n========================================');
    Logger.log('✅ Phase 1 統合テスト成功！');
    Logger.log('========================================');
    Logger.log('\n結果サマリー:');
    Logger.log(`  Workflow Log ID: ${workflowLogId}`);
    Logger.log(`  Engagement Log ID: ${engagementLogId}`);
    Logger.log(`  Candidates_Master: ${masterResult}`);
    Logger.log(`  Candidate_Scores: 行${scoresRow}`);
    Logger.log(`  Candidate_Insights: ${insightsResult}`);
    Logger.log(`  Evaluation_Master: ${evalId}`);
    Logger.log(`  Test Candidate ID: ${testCandidateId}`);

    return {
      success: true,
      test_candidate_id: testCandidateId,
      results: {
        workflow_log: workflowLogId,
        engagement_log: engagementLogId,
        candidates_master: masterResult,
        candidate_scores: scoresRow,
        candidate_insights: insightsResult,
        evaluation_master: evalId
      }
    };

  } catch (error) {
    Logger.log('\n========================================');
    Logger.log('❌ Phase 1 統合テスト失敗');
    Logger.log('========================================');
    Logger.log('エラー: ' + error.message);
    Logger.log('スタック: ' + error.stack);

    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * 個別関数のテスト
 * 各関数を1つずつテストします
 */
function testIndividualFunctions() {
  Logger.log('========================================');
  Logger.log('個別関数テスト開始');
  Logger.log('========================================\n');

  const testCandidateId = 'TEST_IND_' + new Date().getTime();
  let passCount = 0;
  let failCount = 0;

  // Test 1: appendToDifyWorkflowLog
  try {
    Logger.log('Test 1: appendToDifyWorkflowLog()');
    const logId = appendToDifyWorkflowLog({
      workflow_id: 'TEST_WF_' + new Date().getTime(),
      candidate_id: testCandidateId,
      status: 'SUCCESS',
      phase: 'IndividualTest',
      duration_seconds: 2.5
    });
    Logger.log(`  ✅ SUCCESS - ログID: ${logId}\n`);
    passCount++;
  } catch (e) {
    Logger.log(`  ❌ FAILED: ${e.message}\n`);
    failCount++;
  }

  // Test 2: appendToEngagementLog
  try {
    Logger.log('Test 2: appendToEngagementLog()');
    const engageId = appendToEngagementLog({
      candidate_id: testCandidateId,
      candidate_name: 'テスト個別_太郎',
      contact_type: 'テスト面接',
      acceptance_rate_ai: 75,
      confidence_level: 'HIGH'
    });
    Logger.log(`  ✅ SUCCESS - ログID: ${engageId}\n`);
    passCount++;
  } catch (e) {
    Logger.log(`  ❌ FAILED: ${e.message}\n`);
    failCount++;
  }

  // Test 3: updateOrInsertCandidatesMaster
  try {
    Logger.log('Test 3: updateOrInsertCandidatesMaster()');
    const masterResult = updateOrInsertCandidatesMaster({
      candidate_id: testCandidateId,
      '氏名': 'テスト個別_太郎',
      '採用区分': '新卒',
      '現在ステータス': 'テスト中'
    });
    Logger.log(`  ✅ SUCCESS - 結果: ${masterResult}\n`);
    passCount++;
  } catch (e) {
    Logger.log(`  ❌ FAILED: ${e.message}\n`);
    failCount++;
  }

  // Test 4: appendToCandidateScores
  try {
    Logger.log('Test 4: appendToCandidateScores()');
    const scoresRow = appendToCandidateScores({
      candidate_id: testCandidateId,
      '氏名': 'テスト個別_太郎',
      '最新_Philosophy': 25,
      '最新_Strategy': 28,
      '最新_Motivation': 18,
      '最新_Execution': 17,
      '最新_合計スコア': 88
    });
    Logger.log(`  ✅ SUCCESS - 行番号: ${scoresRow}\n`);
    passCount++;
  } catch (e) {
    Logger.log(`  ❌ FAILED: ${e.message}\n`);
    failCount++;
  }

  // Test 5: updateOrInsertCandidateInsights
  try {
    Logger.log('Test 5: updateOrInsertCandidateInsights()');
    const insightsResult = updateOrInsertCandidateInsights({
      candidate_id: testCandidateId,
      '氏名': 'テスト個別_太郎',
      'コアモチベーション': 'テスト動機',
      '主要懸念事項': 'テスト懸念'
    });
    Logger.log(`  ✅ SUCCESS - 結果: ${insightsResult}\n`);
    passCount++;
  } catch (e) {
    Logger.log(`  ❌ FAILED: ${e.message}\n`);
    failCount++;
  }

  // 結果サマリー
  Logger.log('========================================');
  Logger.log('個別関数テスト完了');
  Logger.log('========================================');
  Logger.log(`成功: ${passCount}/5`);
  Logger.log(`失敗: ${failCount}/5`);
  Logger.log(`テスト候補者ID: ${testCandidateId}`);

  return {
    passed: passCount,
    failed: failCount,
    total: 5,
    test_candidate_id: testCandidateId
  };
}

// ========================================
// Phase 2: 行動データ取得関数
// ========================================

/**
 * Survey_Send_Logから候補者の行動データを取得・分析
 *
 * @param {string} candidateId - 候補者ID
 * @return {Object} 行動データ（JSON）
 */
function getBehaviorData(candidateId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Survey_Send_Log');

  if (!sheet) {
    Logger.log('Survey_Send_Logシートが見つかりません');
    return null;
  }

  // 全データ取得
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];

  // candidate_id列のインデックスを取得
  const candidateIdColIndex = headers.indexOf('candidate_id');
  if (candidateIdColIndex === -1) {
    Logger.log('candidate_id列が見つかりません');
    return null;
  }

  // 候補者のデータを抽出
  const candidateRows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][candidateIdColIndex] === candidateId) {
      candidateRows.push(values[i]);
    }
  }

  if (candidateRows.length === 0) {
    Logger.log('候補者のデータが見つかりません: ' + candidateId);
    return {
      candidate_id: candidateId,
      survey_count: 0,
      response_count: 0,
      response_rate: 0,
      avg_response_time_hours: 0,
      engagement_score: 0,
      last_survey_date: null,
      has_data: false
    };
  }

  // 列インデックス取得
  const surveyDateCol = headers.indexOf('survey_date') || headers.indexOf('送信日時');
  const responseStatusCol = headers.indexOf('response_status') || headers.indexOf('回答状況');
  const responseTimeCol = headers.indexOf('response_time_hours') || headers.indexOf('回答所要時間（時間）');
  const surveyTypeCol = headers.indexOf('survey_type') || headers.indexOf('サーベイ種類');

  // 統計計算
  let surveyCount = candidateRows.length;
  let responseCount = 0;
  let totalResponseTime = 0;
  let responseTimeCount = 0;
  let lastSurveyDate = null;

  candidateRows.forEach(row => {
    // 回答数カウント
    if (responseStatusCol !== -1) {
      const status = row[responseStatusCol];
      if (status === '回答済み' || status === 'responded' || status === '完了') {
        responseCount++;
      }
    }

    // 回答時間集計
    if (responseTimeCol !== -1 && row[responseTimeCol]) {
      const time = parseFloat(row[responseTimeCol]);
      if (!isNaN(time)) {
        totalResponseTime += time;
        responseTimeCount++;
      }
    }

    // 最終サーベイ日時
    if (surveyDateCol !== -1 && row[surveyDateCol]) {
      const date = new Date(row[surveyDateCol]);
      if (!lastSurveyDate || date > lastSurveyDate) {
        lastSurveyDate = date;
      }
    }
  });

  // 回答率
  const responseRate = surveyCount > 0 ? (responseCount / surveyCount) * 100 : 0;

  // 平均回答時間（時間単位）
  const avgResponseTimeHours = responseTimeCount > 0 ? totalResponseTime / responseTimeCount : 0;

  // エンゲージメントスコア算出（0-100点）
  let engagementScore = 0;

  // 回答率の貢献（0-50点）
  engagementScore += responseRate * 0.5;

  // 回答速度の貢献（0-30点）
  // 24時間以内: 30点、48時間以内: 20点、72時間以内: 10点、それ以降: 0点
  if (avgResponseTimeHours > 0) {
    if (avgResponseTimeHours <= 24) {
      engagementScore += 30;
    } else if (avgResponseTimeHours <= 48) {
      engagementScore += 20;
    } else if (avgResponseTimeHours <= 72) {
      engagementScore += 10;
    }
  }

  // アクティビティの貢献（0-20点）
  // 5回以上: 20点、3-4回: 15点、1-2回: 10点
  if (surveyCount >= 5) {
    engagementScore += 20;
  } else if (surveyCount >= 3) {
    engagementScore += 15;
  } else if (surveyCount >= 1) {
    engagementScore += 10;
  }

  // 結果オブジェクト
  return {
    candidate_id: candidateId,
    survey_count: surveyCount,
    response_count: responseCount,
    response_rate: Math.round(responseRate * 10) / 10,  // 小数第1位まで
    avg_response_time_hours: Math.round(avgResponseTimeHours * 10) / 10,
    engagement_score: Math.round(engagementScore),
    last_survey_date: lastSurveyDate ? lastSurveyDate.toISOString() : null,
    has_data: true,
    behavior_summary: generateBehaviorSummary(responseRate, avgResponseTimeHours, surveyCount)
  };
}

/**
 * 行動データのサマリーテキスト生成
 */
function generateBehaviorSummary(responseRate, avgResponseTime, surveyCount) {
  let summary = [];

  // 回答率の評価
  if (responseRate >= 80) {
    summary.push('非常に高い回答率');
  } else if (responseRate >= 50) {
    summary.push('良好な回答率');
  } else if (responseRate >= 30) {
    summary.push('やや低い回答率');
  } else {
    summary.push('低い回答率');
  }

  // 回答速度の評価
  if (avgResponseTime > 0) {
    if (avgResponseTime <= 24) {
      summary.push('迅速な回答');
    } else if (avgResponseTime <= 48) {
      summary.push('適度な回答速度');
    } else {
      summary.push('回答に時間がかかる傾向');
    }
  }

  // アクティビティの評価
  if (surveyCount >= 5) {
    summary.push('高いエンゲージメント');
  } else if (surveyCount >= 3) {
    summary.push('適度なエンゲージメント');
  } else {
    summary.push('限定的なエンゲージメント');
  }

  return summary.join('、');
}

/**
 * テスト関数
 */
function testGetBehaviorData() {
  // テスト用候補者ID（実際のIDに置き換えてください）
  const testCandidateId = 'CAND_20251217034006';

  Logger.log('=== 行動データ取得テスト ===');
  const result = getBehaviorData(testCandidateId);

  if (result) {
    Logger.log('✅ テスト成功');
    Logger.log(JSON.stringify(result, null, 2));
  } else {
    Logger.log('❌ テスト失敗');
  }
}

// ========================================
// Phase 2: 行動データ取得エンドポイント
// ========================================

/**
 * 行動データ取得エンドポイント
 * Difyから呼び出される専用エンドポイント
 *
 * @param {Object} data - リクエストデータ
 * @return {ContentService.TextOutput} JSON レスポンス
 */
function handleGetBehaviorData(data) {
  const startTime = new Date();

  try {
    // candidate_idの取得
    const candidateId = data.candidate_id;

    if (!candidateId) {
      throw new Error('candidate_idが必須です');
    }

    Logger.log('=== 行動データ取得リクエスト ===');
    Logger.log('candidate_id: ' + candidateId);

    // getBehaviorData()を呼び出し
    const behaviorData = getBehaviorData(candidateId);

    if (!behaviorData) {
      // データが見つからない場合もエラーにせず、空データを返す
      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          data: {
            candidate_id: candidateId,
            survey_count: 0,
            response_count: 0,
            response_rate: 0,
            avg_response_time_hours: 0,
            engagement_score: 0,
            last_survey_date: null,
            has_data: false,
            behavior_summary: "データなし"
          },
          message: 'データが見つかりませんでした',
          execution_time_seconds: ((new Date() - startTime) / 1000).toFixed(2)
        }, null, 2))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 成功レスポンス
    Logger.log('✅ 行動データ取得成功');
    Logger.log('エンゲージメントスコア: ' + behaviorData.engagement_score);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        data: behaviorData,
        message: '行動データ取得成功',
        execution_time_seconds: ((new Date() - startTime) / 1000).toFixed(2)
      }, null, 2))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('❌ 行動データ取得エラー: ' + error.message);
    Logger.log('スタック: ' + error.stack);

    // エラーレスポンス
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.message,
        stack: error.stack
      }, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * handleGetBehaviorData のテスト
 */
function testHandleGetBehaviorData() {
  Logger.log('=== handleGetBehaviorData テスト ===');

  // モックリクエスト作成
  const mockData = {
    action: 'get_behavior_data',
    candidate_id: 'CAND_20251217034006'  // テスト用ID
  };

  // handleGetBehaviorData を直接呼び出し
  const response = handleGetBehaviorData(mockData);
  const result = JSON.parse(response.getContent());

  Logger.log('レスポンス:');
  Logger.log(JSON.stringify(result, null, 2));

  // 検証
  if (result.success) {
    Logger.log('✅ テスト成功');
    Logger.log('has_data: ' + result.data.has_data);
    Logger.log('engagement_score: ' + result.data.engagement_score);
  } else {
    Logger.log('❌ テスト失敗: ' + result.error);
  }
}
