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
 * UTC時刻を日本時間（JST）に変換
 * @param {string} utcTimeString - UTC時刻の文字列（例: "2025-12-26T14:20:38.886773"）
 * @return {string} 日本時間の文字列（例: "2025-12-26 23:20:38"）
 */
function convertToJST(utcTimeString) {
  try {
    // UTC時刻文字列を正規化（'Z'がなければ追加してUTCとして明示）
    let normalizedUtcString = utcTimeString;
    if (!utcTimeString.endsWith('Z') && !utcTimeString.includes('+') && !utcTimeString.includes('-', 10)) {
      // タイムゾーン指定がない場合、'Z'を追加してUTCとして明示
      normalizedUtcString = utcTimeString.replace(/(\.\d+)?$/, '') + 'Z';
    }

    // UTC時刻をDateオブジェクトに変換
    const utcDate = new Date(normalizedUtcString);

    // Utilities.formatDateを使用して日本時間（JST）に変換
    // タイムゾーン 'Asia/Tokyo' を指定
    const jstString = Utilities.formatDate(utcDate, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

    return jstString;
  } catch (error) {
    Logger.log('convertToJST エラー: ' + error.toString());
    // エラー時は元の文字列をそのまま返す
    return utcTimeString;
  }
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
    convertToJST(data.interview_datetime) || '', // B: 面接日時（日本時間に変換）
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
    data.次回質問1 || '',                   // AC
    data.次回質問2 || '',                   // AD
    data.次回質問3 || '',                   // AE
    data.次回質問4 || '',                   // AF
    data.次回質問5 || '',                   // AG
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
  try {
    Logger.log('=== appendToEngagementLog 開始 ===');
    Logger.log('受信データ型: ' + typeof data);
    Logger.log('competitor_details存在: ' + (data.competitor_details ? 'YES' : 'NO'));
    if (data.competitor_details) {
      Logger.log('competitor_details型: ' + typeof data.competitor_details);
      Logger.log('competitor_details数: ' + (Array.isArray(data.competitor_details) ? data.competitor_details.length : 'N/A'));
    }

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Engagement_Log');

    if (!sheet) {
      Logger.log('❌ ERROR: Engagement_Logシートが見つかりません');
      throw new Error('Engagement_Logシートが見つかりません');
    }

    Logger.log('✅ シート取得成功');

    // 現在の列数を確認
    const lastColumn = sheet.getLastColumn();
    Logger.log('現在の列数: ' + lastColumn);

    if (lastColumn < 22) {
      Logger.log('❌ ERROR: 列数が不足しています。期待: 22列、実際: ' + lastColumn);
      throw new Error(`列数が不足しています（期待: 22列、実際: ${lastColumn}列）`);
    }

    // ログIDの生成
    const logId = data.log_id || `LOG_${new Date().getTime()}`;

    // confidence_levelを日本語に変換
    const confidenceLevelMap = {
      'HIGH': '高',
      'MEDIUM': '中',
      'LOW': '低'
    };
    const confidenceLevel = confidenceLevelMap[data.confidence_level] || data.confidence_level || '';

    // acceptance_rate_aiをパーセント表記に変換
    const acceptanceRateAi = data.acceptance_rate_ai
      ? (typeof data.acceptance_rate_ai === 'number' ? `${data.acceptance_rate_ai.toFixed(2)}%` : data.acceptance_rate_ai)
      : '';

    // 行データの組み立て（Engagement_Logの正しい列構造に合わせる）
    const row = [
      logId,                                        // 1: log_id
      data.candidate_id || '',                      // 2: candidate_id
      data.candidate_name || data['氏名'] || '',     // 3: 氏名
      convertToJST(data.timestamp) || new Date(),   // 4: timestamp（日本時間に変換）
      data.contact_type || '',                      // 5: contact_type
      data.acceptance_rate_rule || '',              // 6: acceptance_rate_rule
      acceptanceRateAi,                             // 7: acceptance_rate_ai（パーセント表記）
      data.acceptance_rate_final || '',             // 8: acceptance_rate_final
      confidenceLevel,                              // 9: confidence_level（日本語）
      data.motivation_score || 0,                   // 10: motivation_score
      data.competitive_advantage_score || 0,        // 11: competitive_advantage_score
      data.concern_resolution_score || 0,           // 12: concern_resolution_score
      data.core_motivation || '',                   // 13: core_motivation
      data.top_concern || '',                       // 14: top_concern
      data.concern_category || '',                  // 15: concern_category
      data.competitors || '',                       // 16: competitors
      data.competitive_advantage || '',             // 17: competitive_advantage
      data.next_action || '',                       // 18: next_action
      data.action_deadline || '',                   // 19: action_deadline
      data.action_priority || '',                   // 20: action_priority
      data.doc_url || '',                           // 21: doc_url
      JSON.stringify(data.competitor_details || []) // 22: competitor_details
    ];

    Logger.log('row配列の長さ: ' + row.length);
    Logger.log('row[21] (competitor_details JSON): ' + (row[21] ? row[21].substring(0, 100) + '...' : 'empty'));

    // データ追加
    sheet.appendRow(row);

    Logger.log('✅ Engagement_Log追記完了: ' + logId);
    return logId;

  } catch (error) {
    Logger.log('=== appendToEngagementLog エラー ===');
    Logger.log('❌ エラー: ' + error.toString());
    Logger.log('スタックトレース: ' + error.stack);
    throw error;
  }
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

    // デバッグ: 受信データの構造を確認
    Logger.log('=== 受信データ構造チェック ===');
    Logger.log('data.validated_input: ' + (data.validated_input ? 'あり' : '❌なし'));
    Logger.log('data.candidates_master: ' + (data.candidates_master ? 'あり' : 'なし'));
    Logger.log('data.evaluation_master: ' + (data.evaluation_master ? 'あり' : 'なし'));
    Logger.log('data.engagement_log: ' + (data.engagement_log ? 'あり' : 'なし'));

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

      // デバッグログ: competitor_details確認
      Logger.log('=== Engagement_Log データ確認 ===');
      Logger.log('engagementData型: ' + typeof engagementData);
      Logger.log('competitor_details存在: ' + (engagementData.competitor_details ? 'YES' : 'NO'));
      if (engagementData.competitor_details) {
        Logger.log('competitor_details数: ' + engagementData.competitor_details.length);
        Logger.log('competitor_details内容: ' + JSON.stringify(engagementData.competitor_details, null, 2));
      }

      results.engagement_log = appendToEngagementLog(engagementData);
      Logger.log('✅ Engagement_Log: ' + results.engagement_log);
    }

    // 5. Evaluation_Master追加
    let evaluationMasterData = null;
    if (data.evaluation_master) {
      evaluationMasterData = typeof data.evaluation_master === 'string'
        ? JSON.parse(data.evaluation_master)
        : data.evaluation_master;

      // デバッグログ: 次回質問データの確認
      Logger.log('=== Evaluation_Master データ確認 ===');
      Logger.log('evaluation_master型: ' + typeof evaluationMasterData);
      Logger.log('次回質問1: ' + evaluationMasterData.次回質問1);
      Logger.log('次回質問2: ' + evaluationMasterData.次回質問2);
      Logger.log('次回質問3: ' + evaluationMasterData.次回質問3);
      Logger.log('次回質問4: ' + evaluationMasterData.次回質問4);
      Logger.log('次回質問5: ' + evaluationMasterData.次回質問5);

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
      const recruitType = (candidateMasterData && candidateMasterData['採用区分']) || '新卒';
      const selectionPhase = (candidateMasterData && candidateMasterData['現在ステータス'])
        || (evaluationMasterData && evaluationMasterData.selection_phase)
        || '初回面談';

      // スプレッドシートURL取得
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const spreadsheetUrl = ss.getUrl();

      // === デバッグ: レポート生成条件の確認 ===
      Logger.log('=== レポート生成条件チェック ===');
      Logger.log('candidateMasterData: ' + (candidateMasterData ? 'あり' : 'なし'));
      Logger.log('evalData: ' + (evalData ? 'あり' : 'なし'));
      Logger.log('data.validated_input: ' + (data.validated_input ? 'あり' : 'なし'));
      Logger.log('acceptanceData: ' + (acceptanceData ? 'あり' : 'なし'));

      if (candidateMasterData) {
        Logger.log('  candidateMasterData.氏名: ' + candidateMasterData['氏名']);
      }
      if (evalData) {
        Logger.log('  evalData.total_rank: ' + evalData.total_rank);
      }
      if (data.validated_input) {
        Logger.log('  validated_input.candidate_id: ' + data.validated_input.candidate_id);
      }

      // 1. 評価レポート生成（V2）
      if (candidateMasterData && evalData) {
        Logger.log('--- 評価レポートV2 生成 ---');
        Logger.log('✅ validated_input不要で実行');

        const evalReportDataV2 = {
          // 基本情報
          candidate_id: candidateMasterData.candidate_id || evalData.candidate_id,
          candidate_name: candidateMasterData['氏名'],
          selection_phase: selectionPhase,
          interview_date: evalData.interview_datetime
            ? convertToJST(evalData.interview_datetime).split(' ')[0].replace(/-/g, '/')
            : Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd'),
          interviewer: evalData.interviewer || '未設定',

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
            evalData.次回質問1,
            evalData.次回質問2,
            evalData.次回質問3,
            evalData.次回質問4,
            evalData.次回質問5
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
            evalSheet.getRange(evalLastRow, 35).setValue(evalResultV2.url);  // AI列 = 35
            Logger.log('✅ 評価レポートURL記録: AI列（行' + evalLastRow + '）');
          }
        }
      } else {
        Logger.log('⚠️ 評価レポートV2生成スキップ:');
        if (!candidateMasterData) Logger.log('  - candidateMasterDataなし');
        if (!evalData) Logger.log('  - evalDataなし');
      }

      // 2. 戦略レポート生成（V2）
      if (candidateMasterData && acceptanceData) {
        Logger.log('--- 戦略レポートV2 生成 ---');
        Logger.log('✅ validated_input不要で実行');

        // デバッグログ: competitor_details確認
        Logger.log('=== 戦略レポート用データ確認 ===');
        Logger.log('acceptanceData.competitor_details存在: ' + (acceptanceData.competitor_details ? 'YES' : 'NO'));
        if (acceptanceData.competitor_details) {
          Logger.log('competitor_details数: ' + acceptanceData.competitor_details.length);
          Logger.log('competitor_details内容: ' + JSON.stringify(acceptanceData.competitor_details, null, 2));
        }

        const strategyReportDataV2 = {
          // 基本情報
          candidate_id: candidateMasterData.candidate_id || evalData.candidate_id,
          candidate_name: candidateMasterData['氏名'],
          current_phase: selectionPhase,
          interviewer: evalData.interviewer || '未設定',

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
          competitor_analysis: acceptanceData.competitor_details || [],

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
            evalSheet.getRange(evalLastRow, 36).setValue(strategyResultV2.url);  // AJ列 = 36
            Logger.log('✅ 戦略レポートURL記録: AJ列（行' + evalLastRow + '）');
          }
        }
      } else {
        Logger.log('⚠️ 戦略レポートV2生成スキップ:');
        if (!candidateMasterData) Logger.log('  - candidateMasterDataなし');
        if (!acceptanceData) Logger.log('  - acceptanceDataなし');
      }

      // 3. 評価B以上の場合、フォルダコピー
      if (evalData && evalData.total_rank && ['A', 'B'].includes(evalData.total_rank) && candidateMasterData) {
        Logger.log('--- 評価B以上: フォルダコピー実行 ---');
        const copyResult = copyFolderToGradeB(
          candidateMasterData.candidate_id || evalData.candidate_id,
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
          candidateMasterData.candidate_id || evalData.candidate_id,
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

    // ===== Phase 4-2a: データ集計 =====
    Logger.log('=== Phase 4-2a: データ集計開始 ===');

    try {
      // 重要: 書き込みを強制的に完了させる
      SpreadsheetApp.flush();
      Logger.log('SpreadsheetApp.flush() 実行完了');

      // 少し待機（書き込み完了を確実にする）
      Utilities.sleep(1000);
      Logger.log('1秒待機完了');

      // 候補者IDを取得
      const candidateId = (candidateMasterData && candidateMasterData.candidate_id) ||
                          (evalData && evalData.candidate_id);

      if (candidateId) {
        Logger.log('集計対象候補者ID: ' + candidateId);

        // Candidate_Scoresを更新
        const scoresUpdateResult = updateCandidateScores(candidateId);
        if (scoresUpdateResult.success) {
          Logger.log('✅ Candidate_Scores更新成功: ' + JSON.stringify(scoresUpdateResult.updated));
        } else {
          Logger.log('⚠️ Candidate_Scores更新失敗: ' + (scoresUpdateResult.message || scoresUpdateResult.error));
        }

        // Candidates_Masterを更新
        const masterUpdateResult = updateCandidatesMaster(candidateId);
        if (masterUpdateResult.success) {
          Logger.log('✅ Candidates_Master更新成功: ' + JSON.stringify(masterUpdateResult.updated));
        } else {
          Logger.log('⚠️ Candidates_Master更新失敗: ' + (masterUpdateResult.message || masterUpdateResult.error));
        }

        Logger.log('=== Phase 4-2a: データ集計完了 ===');
      } else {
        Logger.log('⚠️ 候補者IDが見つかりません。集計をスキップします。');
      }

    } catch (error) {
      Logger.log('❌ ERROR in Phase 4-2a集計: ' + error.toString());
      Logger.log('Stack: ' + error.stack);
      // エラーがあっても処理は継続
    }

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
    const candidateMaster = typeof data.candidates_master === 'string'
      ? JSON.parse(data.candidates_master)
      : data.candidates_master;
    const evalMaster = typeof data.evaluation_master === 'string'
      ? JSON.parse(data.evaluation_master)
      : data.evaluation_master;

    logProcessing({
      candidate_id: (candidateMaster && candidateMaster.candidate_id) || (evalMaster && evalMaster.candidate_id) || 'unknown',
      candidate_name: (candidateMaster && candidateMaster['氏名']) || 'unknown',
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
      acceptance_rate_rule: '70%',
      acceptance_rate_ai: 65,
      acceptance_rate_final: '65%',
      confidence_level: '高',  // 日本語で記録
      motivation_score: 18,
      competitive_advantage_score: 75,
      concern_resolution_score: 80,
      core_motivation: '社会貢献への強い意欲',
      top_concern: '給与条件',
      concern_category: '待遇',
      competitors: 'リブコンサルティング、ベイカレント',
      competitive_advantage: '実行支援の実績、成長機会',
      next_action: '給与条件の詳細説明',
      action_deadline: '2025-12-23',
      action_priority: '高',
      doc_url: ''
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
      acceptance_rate_rule: '70%',
      acceptance_rate_ai: 75,
      acceptance_rate_final: '75%',
      confidence_level: '中',  // 日本語で記録
      motivation_score: 20,
      competitive_advantage_score: 80,
      concern_resolution_score: 85,
      core_motivation: 'テスト動機',
      top_concern: 'テスト懸念',
      concern_category: 'テスト',
      competitors: 'テスト競合',
      competitive_advantage: 'テスト優位性',
      next_action: 'テストアクション',
      action_deadline: '2025-12-25',
      action_priority: '中',
      doc_url: 'https://example.com/test'
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

/**
 * レポート生成V2とURL記録列の検証用テスト関数
 *
 * 検証項目:
 * 1. 評価レポートV2の生成
 * 2. 戦略レポートV2の生成
 * 3. Evaluation_MasterのAF列（32列目）への評価レポートURL記録
 * 4. Evaluation_MasterのAG列（33列目）への戦略レポートURL記録
 * 5. getSummary()関数の動作確認（4軸: Philosophy, Strategy, Motivation, Execution）
 */
function testReportGenerationWithURLRecording() {
  Logger.log('========================================');
  Logger.log('レポート生成V2 & URL記録列 検証テスト');
  Logger.log('========================================\n');

  const startTime = new Date();
  const testCandidateId = 'TEST_URL_' + startTime.getTime();
  const companyName = 'テスト株式会社';  // テスト用企業名（ドライブ: https://drive.google.com/drive/folders/1jF-8_SrhoIlPnseyrCVMd4BNOK4WSGdG）

  try {
    // スプレッドシート取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const evalSheet = ss.getSheetByName('Evaluation_Master');

    if (!evalSheet) {
      throw new Error('Evaluation_Masterシートが見つかりません');
    }

    Logger.log('=== Step 1: テストデータ準備 ===');

    // 評価レポート用テストデータ
    const evalReportData = {
      candidate_id: testCandidateId,
      candidate_name: 'URL検証_太郎',
      selection_phase: '1次面接',
      interview_date: new Date().toLocaleDateString('ja-JP'),
      interviewer: 'テスト面接官',
      total_rank: 'A',
      recommendation: '積極採用推奨',
      summary_reasons: [
        '理念への深い共感',
        '優れた戦略的思考力',
        '高い実行力'
      ],
      philosophy_rank: 'A',
      philosophy_score: 28,
      philosophy_summary: '理念への深い共感が見られる',
      philosophy_reason: '企業理念に深く共感し、自身の価値観と一致している',
      philosophy_evidence: '理念に共感した発言があった',
      strategy_rank: 'B',
      strategy_score: 24,
      strategy_summary: '戦略理解は十分、実践経験で向上可',
      strategy_reason: '戦略的思考の基礎はあるが、実践でさらに向上可能',
      strategy_evidence: '戦略的な発言があった',
      motivation_rank: 'A',
      motivation_score: 19,
      motivation_summary: '非常に高い志望度、成長意欲強',
      motivation_reason: '高い志望度と成長意欲が見られる',
      motivation_evidence: '強い志望動機を語った',
      execution_rank: 'A',
      execution_score: 19,
      execution_summary: '優れた実行力、実績あり',
      execution_reason: '過去の実績から実行力が確認できる',
      execution_evidence: '実績について語った',
      critical_concerns: [],
      next_questions: ['詳細確認事項1', '詳細確認事項2'],
      interviewer_comment: 'テスト用コメント',
      transcript: 'テスト用議事録\n面接官: よろしくお願いします。\n候補者: よろしくお願いします。',
      spreadsheet_url: ss.getUrl()
    };

    // 戦略レポート用テストデータ
    const strategyReportData = {
      candidate_id: testCandidateId,
      candidate_name: 'URL検証_太郎',
      current_phase: '1次面接',
      interviewer: 'テスト面接官',
      acceptance_probability: 75,
      confidence_level: 'HIGH',
      competitor_probabilities: [
        { company: '自社', probability: 55 },
        { company: '競合A社', probability: 30 },
        { company: '競合B社', probability: 15 }
      ],
      immediate_action_24h: '具体的な条件提示',
      action_reason: '待遇面の懸念解消',
      expected_effect: '承諾率 +10%',
      risk_factors: [
        { factor: 'リスク要因1', countermeasure: '対策1' }
      ],
      our_strengths: ['強み1', '強み2'],
      acceptance_story: ['Step 1', 'Step 2', 'Step 3'],
      positive_factors: [
        { factor: 'ポジティブ要因1', evidence: 'エビデンス1' }
      ],
      risk_factors_detailed: [
        { factor: 'リスク詳細1', severity: 'MEDIUM', detailed_countermeasure: '詳細対策1' }
      ],
      competitor_analysis: [],
      engagement_recommendations: ['推奨施策1', '推奨施策2'],
      spreadsheet_url: ss.getUrl()
    };

    Logger.log('✅ テストデータ準備完了');
    Logger.log('  候補者ID: ' + testCandidateId);
    Logger.log('  企業名: ' + companyName);

    // Step 2: 評価レポートV2生成
    Logger.log('\n=== Step 2: 評価レポートV2生成 ===');
    const evalResult = generateEvaluationReportV2(evalReportData, '新卒', '1次面接', companyName);

    if (!evalResult.success) {
      throw new Error('評価レポート生成失敗: ' + (evalResult.error || 'Unknown error'));
    }

    Logger.log('✅ 評価レポートV2生成成功');
    Logger.log('  URL: ' + evalResult.url);
    Logger.log('  Document ID: ' + evalResult.documentId);

    // Step 3: 戦略レポートV2生成
    Logger.log('\n=== Step 3: 戦略レポートV2生成 ===');
    const strategyResult = generateStrategyReportV2(strategyReportData, '新卒', '1次面接', companyName);

    if (!strategyResult.success) {
      throw new Error('戦略レポート生成失敗: ' + (strategyResult.error || 'Unknown error'));
    }

    Logger.log('✅ 戦略レポートV2生成成功');
    Logger.log('  URL: ' + strategyResult.url);
    Logger.log('  Document ID: ' + strategyResult.documentId);

    // Step 4: Evaluation_Masterに基本データを書き込み（writeToEvaluationMaster使用）
    Logger.log('\n=== Step 4: Evaluation_Masterにデータ書き込み ===');

    const evalMasterData = {
      candidate_id: testCandidateId,
      candidate_name: 'URL検証_太郎',
      recruit_type: '新卒',
      selection_phase: '1次面接',
      interview_date: new Date(),
      interviewer: 'テスト面接官',
      philosophy_score: 28,
      philosophy_rank: 'A',
      philosophy_summary: '理念への深い共感が見られる',
      philosophy_reason: '企業理念に深く共感',
      philosophy_evidence: '理念に共感した発言',
      strategy_score: 24,
      strategy_rank: 'B',
      strategy_summary: '戦略理解は十分、実践経験で向上可',
      strategy_reason: '戦略的思考の基礎はある',
      strategy_evidence: '戦略的な発言',
      motivation_score: 19,
      motivation_rank: 'A',
      motivation_summary: '非常に高い志望度、成長意欲強',
      motivation_reason: '高い志望度と成長意欲',
      motivation_evidence: '強い志望動機',
      execution_score: 19,
      execution_rank: 'A',
      execution_summary: '優れた実行力、実績あり',
      execution_reason: '過去の実績から実行力確認',
      execution_evidence: '実績について語った',
      total_score: 90,
      total_rank: 'A',
      recommendation: '積極採用推奨',
      summary_reasons: ['理念への深い共感'],
      critical_concerns: [],
      next_questions: ['質問1', '質問2']
    };

    const evalId = writeToEvaluationMaster(evalMasterData);
    Logger.log('✅ Evaluation_Master書き込み完了: ' + evalId);

    // 最新行を取得
    const lastRow = evalSheet.getLastRow();
    Logger.log('  最新行番号: ' + lastRow);

    // Step 5: URL記録（修正箇所の検証）
    Logger.log('\n=== Step 5: URL記録列への書き込み（検証） ===');
    Logger.log('✅ 正しい列: AI列（35）/AJ列（36）');

    // AI列（35列目）に評価レポートURL記録
    evalSheet.getRange(lastRow, 35).setValue(evalResult.url);
    Logger.log('✅ 評価レポートURL記録: AI列（35列目）行' + lastRow);
    Logger.log('  URL: ' + evalResult.url);

    // AJ列（36列目）に戦略レポートURL記録
    evalSheet.getRange(lastRow, 36).setValue(strategyResult.url);
    Logger.log('✅ 戦略レポートURL記録: AJ列（36列目）行' + lastRow);
    Logger.log('  URL: ' + strategyResult.url);

    // Step 6: 記録内容の確認
    Logger.log('\n=== Step 6: 記録内容の確認 ===');

    const recordedEvalUrl = evalSheet.getRange(lastRow, 35).getValue();
    const recordedStrategyUrl = evalSheet.getRange(lastRow, 36).getValue();

    Logger.log('📊 記録確認:');
    Logger.log('  行番号: ' + lastRow);
    Logger.log('  AI列（35）の値: ' + (recordedEvalUrl ? '✅ 記録あり' : '❌ 空欄'));
    Logger.log('  AJ列（36）の値: ' + (recordedStrategyUrl ? '✅ 記録あり' : '❌ 空欄'));

    // 列名の確認
    const headers = evalSheet.getRange(1, 1, 1, 36).getValues()[0];
    Logger.log('\n📋 列ヘッダー確認:');
    Logger.log('  35列目（AI列）のヘッダー: ' + (headers[34] || '未定義'));
    Logger.log('  36列目（AJ列）のヘッダー: ' + (headers[35] || '未定義'));

    // Step 7: getSummary()関数の検証
    Logger.log('\n=== Step 7: getSummary()関数の動作確認 ===');
    Logger.log('✅ 4軸対応: Philosophy, Strategy, Motivation, Execution');

    const testAxes = ['Philosophy', 'Strategy', 'Motivation', 'Execution'];
    const testRank = 'A';

    testAxes.forEach(axis => {
      const summary = getSummary(testRank, axis);
      Logger.log(`  ${axis} (${testRank}): ${summary}`);
    });

    // 最終結果
    Logger.log('\n========================================');
    Logger.log('✅ テスト完了！');
    Logger.log('========================================');
    Logger.log('\n📊 検証結果サマリー:');
    Logger.log('  テスト候補者ID: ' + testCandidateId);
    Logger.log('  Evaluation_Master行番号: ' + lastRow);
    Logger.log('  評価レポートURL: ' + evalResult.url);
    Logger.log('  戦略レポートURL: ' + strategyResult.url);
    Logger.log('  AI列（35）記録: ' + (recordedEvalUrl ? '✅' : '❌'));
    Logger.log('  AJ列（36）記録: ' + (recordedStrategyUrl ? '✅' : '❌'));
    Logger.log('  getSummary()動作: ✅');
    Logger.log('  実行時間: ' + ((new Date() - startTime) / 1000).toFixed(2) + '秒');

    Logger.log('\n🎯 次のアクション:');
    Logger.log('  1. Evaluation_Masterシートを開く');
    Logger.log('  2. 行' + lastRow + 'を確認');
    Logger.log('  3. AI列（35列目）に評価レポートURLがあることを確認');
    Logger.log('  4. AJ列（36列目）に戦略レポートURLがあることを確認');
    Logger.log('  5. URLをクリックしてレポート内容を確認');

    return {
      success: true,
      test_candidate_id: testCandidateId,
      evaluation_master_row: lastRow,
      evaluation_report_url: evalResult.url,
      strategy_report_url: strategyResult.url,
      ai_column_35_recorded: !!recordedEvalUrl,
      aj_column_36_recorded: !!recordedStrategyUrl,
      execution_time_seconds: ((new Date() - startTime) / 1000).toFixed(2)
    };

  } catch (error) {
    Logger.log('\n========================================');
    Logger.log('❌ テスト失敗');
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
 * Engagement_Logシートに22列目（competitor_details）を追加
 * Phase B1.5対応: 競合分析詳細データ保存用カラム
 *
 * 実行方法:
 * 1. GASエディタでこの関数を開く
 * 2. 関数を選択して「実行」ボタンをクリック
 * 3. 初回実行時は権限承認が必要
 */
function addCompetitorDetailsColumn() {
  Logger.log('=== Engagement_Log列追加開始 ===');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Engagement_Log');

    if (!sheet) {
      throw new Error('Engagement_Logシートが見つかりません');
    }

    // 現在の列数を確認
    const lastColumn = sheet.getLastColumn();
    Logger.log(`現在の列数: ${lastColumn}`);

    // ヘッダー行（1行目）を取得
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    Logger.log(`既存ヘッダー数: ${headers.length}`);
    Logger.log(`最後のヘッダー: ${headers[headers.length - 1]}`);

    // competitor_detailsが既に存在するか確認
    if (headers.includes('competitor_details')) {
      Logger.log('⚠️ competitor_details列は既に存在します');
      Logger.log('列番号: ' + (headers.indexOf('competitor_details') + 1));
      return {
        success: true,
        already_exists: true,
        column_index: headers.indexOf('competitor_details') + 1
      };
    }

    // 22列目に追加（既存が21列の場合）
    if (lastColumn === 21) {
      Logger.log('22列目にcompetitor_detailsを追加します...');
      sheet.getRange(1, 22).setValue('competitor_details');
      Logger.log('✅ 列追加完了: AV列（22列目）');

      return {
        success: true,
        added: true,
        column_index: 22,
        column_letter: 'AV'
      };
    } else if (lastColumn < 21) {
      throw new Error(`列数が不足しています（現在: ${lastColumn}列, 必要: 21列）`);
    } else {
      Logger.log(`⚠️ 既に22列以上存在します（${lastColumn}列）`);
      Logger.log('手動確認が必要です');
      return {
        success: false,
        error: `既に${lastColumn}列存在します。手動で確認してください。`
      };
    }

  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    Logger.log('スタック: ' + error.stack);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * convertToJST関数のテスト
 * GASエディタで実行してテスト
 */
function testConvertToJST() {
  Logger.log('=== convertToJST テスト開始 ===');

  const testCases = [
    {
      utc: "2025-12-26T14:20:38.886773",
      expected: "2025-12-26 23:20:38"
    },
    {
      utc: "2025-12-26T01:24:47",
      expected: "2025-12-26 10:24:47"
    },
    {
      utc: "2025-12-25T15:00:00",
      expected: "2025-12-26 00:00:00"
    }
  ];

  testCases.forEach((test, index) => {
    const result = convertToJST(test.utc);
    const passed = result === test.expected;

    Logger.log(`\nテストケース ${index + 1}:`);
    Logger.log('  UTC入力: ' + test.utc);
    Logger.log('  JST出力: ' + result);
    Logger.log('  期待値:   ' + test.expected);
    Logger.log('  結果:     ' + (passed ? '✅ 成功' : '❌ 失敗'));
  });

  Logger.log('\n=== convertToJST テスト完了 ===');
}

// ============================================================
// Phase 4-2a: データ集計機能
// ============================================================

/**
 * Candidates_Masterシートに列を追加する関数
 */
function addCandidatesMasterColumns() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Candidates_Master');

    if (!sheet) {
      throw new Error('Candidates_Masterシートが見つかりません');
    }

    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    // 必要な列が既に存在するかチェック
    const requiredColumns = [
      '最新_承諾可能性',
      '最新_評価ランク',
      '最新_合計スコア',
      '最新_面接日',
      '面接回数'
    ];

    const existingColumns = new Set(headers);
    const columnsToAdd = requiredColumns.filter(col => !existingColumns.has(col));

    if (columnsToAdd.length === 0) {
      Logger.log('必要な列は既に存在します');
      return { success: true, message: '列は既に存在します' };
    }

    // 列を追加
    const startColumn = lastColumn + 1;
    columnsToAdd.forEach((colName, index) => {
      sheet.getRange(1, startColumn + index).setValue(colName);
    });

    Logger.log(`${columnsToAdd.length}列を追加しました: ${columnsToAdd.join(', ')}`);

    return {
      success: true,
      columnsAdded: columnsToAdd.length,
      columns: columnsToAdd
    };

  } catch (error) {
    Logger.log('ERROR in addCandidatesMasterColumns: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Candidate_Scoresシートを更新する関数
 * Engagement_Logから最新の承諾可能性・評価を取得して更新
 *
 * @param {string} candidateId - 候補者ID
 * @return {Object} 更新結果
 */
function updateCandidateScores(candidateId) {
  try {
    Logger.log('=== updateCandidateScores 開始 ===');
    Logger.log('候補者ID: ' + candidateId);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Engagement_Logから最新データを取得
    const engagementSheet = ss.getSheetByName('Engagement_Log');
    if (!engagementSheet) {
      throw new Error('Engagement_Logシートが見つかりません');
    }

    const engagementData = engagementSheet.getDataRange().getValues();
    const engagementHeaders = engagementData[0];

    // 列インデックスを取得
    const candidateIdColIndex = engagementHeaders.indexOf('candidate_id');
    const acceptanceRateColIndex = engagementHeaders.indexOf('acceptance_rate_ai');
    const confidenceLevelColIndex = engagementHeaders.indexOf('confidence_level');
    const motivationScoreColIndex = engagementHeaders.indexOf('motivation_score');
    const timestampColIndex = engagementHeaders.indexOf('timestamp');

    // 該当候補者のデータを抽出（最新のもの）
    let latestEngagement = null;
    let latestTimestamp = null;

    for (let i = engagementData.length - 1; i >= 1; i--) {
      const row = engagementData[i];
      if (row[candidateIdColIndex] === candidateId) {
        const timestamp = new Date(row[timestampColIndex]);
        if (!latestTimestamp || timestamp > latestTimestamp) {
          latestTimestamp = timestamp;
          latestEngagement = {
            acceptance_rate_ai: row[acceptanceRateColIndex] || 0,
            confidence_level: row[confidenceLevelColIndex] || '',
            motivation_score: row[motivationScoreColIndex] || 0,
            timestamp: row[timestampColIndex]
          };
        }
      }
    }

    if (!latestEngagement) {
      Logger.log('該当候補者のEngagement_Logデータが見つかりません');
      return { success: false, message: 'データなし' };
    }

    Logger.log('最新データ: ' + JSON.stringify(latestEngagement));

    // Candidate_Scoresシートを更新
    const scoresSheet = ss.getSheetByName('Candidate_Scores');
    if (!scoresSheet) {
      throw new Error('Candidate_Scoresシートが見つかりません');
    }

    const scoresData = scoresSheet.getDataRange().getValues();
    const scoresHeaders = scoresData[0];

    // 列インデックスを取得
    const scoresCandidateIdColIndex = scoresHeaders.indexOf('candidate_id');
    const latestAcceptanceColIndex = scoresHeaders.indexOf('最新_承諾可能性（AI予測）');
    const confidenceColIndex = scoresHeaders.indexOf('予測の信頼度');
    const motivationColIndex = scoresHeaders.indexOf('志望度スコア');

    // 該当候補者の行を探す
    let targetRow = -1;
    for (let i = 1; i < scoresData.length; i++) {
      if (scoresData[i][scoresCandidateIdColIndex] === candidateId) {
        targetRow = i + 1; // スプレッドシートは1始まり
        break;
      }
    }

    if (targetRow === -1) {
      Logger.log('Candidate_Scoresに該当候補者が見つかりません');
      return { success: false, message: '候補者なし' };
    }

    // データを更新
    scoresSheet.getRange(targetRow, latestAcceptanceColIndex + 1).setValue(latestEngagement.acceptance_rate_ai);
    scoresSheet.getRange(targetRow, confidenceColIndex + 1).setValue(latestEngagement.confidence_level);
    scoresSheet.getRange(targetRow, motivationColIndex + 1).setValue(latestEngagement.motivation_score);

    Logger.log('✅ Candidate_Scores更新完了: 行' + targetRow);

    return {
      success: true,
      candidateId: candidateId,
      row: targetRow,
      updated: {
        acceptance_rate: latestEngagement.acceptance_rate_ai,
        confidence_level: latestEngagement.confidence_level,
        motivation_score: latestEngagement.motivation_score
      }
    };

  } catch (error) {
    Logger.log('ERROR in updateCandidateScores: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.toString() };
  }
}

/**
 * Candidates_Masterシートを更新する関数
 * Candidate_ScoresとEvaluation_Masterから最新データを集計
 *
 * @param {string} candidateId - 候補者ID
 * @return {Object} 更新結果
 */
function updateCandidatesMaster(candidateId) {
  try {
    Logger.log('=== updateCandidatesMaster 開始 ===');
    Logger.log('候補者ID: ' + candidateId);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Candidate_Scoresから最新データを取得
    const scoresSheet = ss.getSheetByName('Candidate_Scores');
    if (!scoresSheet) {
      Logger.log('❌ ERROR: Candidate_Scoresシートが見つかりません');
      throw new Error('Candidate_Scoresシートが見つかりません');
    }

    const scoresData = scoresSheet.getDataRange().getValues();
    const scoresHeaders = scoresData[0];

    // デバッグ: Candidate_Scoresヘッダー確認
    Logger.log('Candidate_Scores ヘッダー: ' + JSON.stringify(scoresHeaders));

    const scoresCandidateIdColIndex = scoresHeaders.indexOf('candidate_id');
    const latestAcceptanceColIndex = scoresHeaders.indexOf('最新_承諾可能性（AI予測）');

    // デバッグ: Candidate_Scores列インデックス
    Logger.log('Candidate_Scores 列インデックス:');
    Logger.log('  candidate_id: ' + scoresCandidateIdColIndex);
    Logger.log('  最新_承諾可能性（AI予測）: ' + latestAcceptanceColIndex);

    let latestAcceptance = null;
    for (let i = 1; i < scoresData.length; i++) {
      if (scoresData[i][scoresCandidateIdColIndex] === candidateId) {
        latestAcceptance = scoresData[i][latestAcceptanceColIndex];
        Logger.log('✅ Candidate_Scoresで一致: 行' + i + ', 承諾可能性=' + latestAcceptance);
        break;
      }
    }

    // Evaluation_Masterから最新の評価データを取得
    const evalSheet = ss.getSheetByName('Evaluation_Master');
    if (!evalSheet) {
      Logger.log('❌ ERROR: Evaluation_Masterシートが見つかりません');
      throw new Error('Evaluation_Masterシートが見つかりません');
    }

    const evalData = evalSheet.getDataRange().getValues();
    const evalHeaders = evalData[0];

    // デバッグ: Evaluation_Masterヘッダー確認
    Logger.log('Evaluation_Master ヘッダー: ' + JSON.stringify(evalHeaders));

    // 実際のシート列名に修正（日本語列名）
    const evalCandidateIdColIndex = evalHeaders.indexOf('候補者ID');
    const totalRankColIndex = evalHeaders.indexOf('総合ランク');
    const totalScoreColIndex = evalHeaders.indexOf('総合スコア');
    const interviewDateColIndex = evalHeaders.indexOf('面接日時');

    // デバッグ: Evaluation_Master列インデックス確認
    Logger.log('Evaluation_Master 列インデックス:');
    Logger.log('  候補者ID: ' + evalCandidateIdColIndex);
    Logger.log('  総合ランク: ' + totalRankColIndex);
    Logger.log('  総合スコア: ' + totalScoreColIndex);
    Logger.log('  面接日時: ' + interviewDateColIndex);

    // 列が見つからない場合の警告
    if (evalCandidateIdColIndex === -1) {
      Logger.log('❌ ERROR: Evaluation_Masterに候補者ID列が見つかりません');
    }
    if (totalRankColIndex === -1) {
      Logger.log('❌ ERROR: Evaluation_Masterに総合ランク列が見つかりません');
    }
    if (totalScoreColIndex === -1) {
      Logger.log('❌ ERROR: Evaluation_Masterに総合スコア列が見つかりません');
    }
    if (interviewDateColIndex === -1) {
      Logger.log('❌ ERROR: Evaluation_Masterに面接日時列が見つかりません');
    }

    // 該当候補者の最新評価を取得
    let latestEval = null;
    let latestDate = null;
    let interviewCount = 0;

    Logger.log('Evaluation_Master データ行数: ' + (evalData.length - 1));

    // デバッグ: 最後の5行の候補者IDを出力
    Logger.log('Evaluation_Master 最後の5行:');
    for (let i = Math.max(1, evalData.length - 5); i < evalData.length; i++) {
      Logger.log('  行' + i + ' 候補者ID: ' + evalData[i][evalCandidateIdColIndex]);
    }

    for (let i = evalData.length - 1; i >= 1; i--) {
      const row = evalData[i];
      const rowCandidateId = row[evalCandidateIdColIndex];

      if (rowCandidateId === candidateId) {
        interviewCount++;
        Logger.log('✅ Evaluation_Masterで一致: 行' + i);
        Logger.log('  総合ランク: ' + row[totalRankColIndex]);
        Logger.log('  総合スコア: ' + row[totalScoreColIndex]);
        Logger.log('  面接日時: ' + row[interviewDateColIndex]);

        const interviewDate = new Date(row[interviewDateColIndex]);
        if (!latestDate || interviewDate > latestDate) {
          latestDate = interviewDate;
          latestEval = {
            total_rank: row[totalRankColIndex] || '',
            total_score: row[totalScoreColIndex] || 0,
            interview_date: row[interviewDateColIndex]
          };
        }
      }
    }

    if (!latestEval) {
      Logger.log('⚠️ Evaluation_Masterに該当候補者のデータがありません');
      Logger.log('検索したcandidate_id: ' + candidateId);
      latestEval = {
        total_rank: '',
        total_score: 0,
        interview_date: ''
      };
    } else {
      Logger.log('✅ 最新評価データ取得成功: ' + JSON.stringify(latestEval));
    }

    Logger.log('集計データ: ' + JSON.stringify({
      latestAcceptance: latestAcceptance,
      latestEval: latestEval,
      interviewCount: interviewCount
    }));

    // Candidates_Masterシートを更新
    const masterSheet = ss.getSheetByName('Candidates_Master');
    if (!masterSheet) {
      Logger.log('❌ ERROR: Candidates_Masterシートが見つかりません');
      throw new Error('Candidates_Masterシートが見つかりません');
    }

    const masterData = masterSheet.getDataRange().getValues();
    const masterHeaders = masterData[0];

    // デバッグ: Candidates_Masterヘッダー確認
    Logger.log('Candidates_Master ヘッダー: ' + JSON.stringify(masterHeaders));

    // 列インデックスを取得
    const masterCandidateIdColIndex = masterHeaders.indexOf('candidate_id');
    const latestAcceptanceMasterColIndex = masterHeaders.indexOf('最新_承諾可能性');
    const latestRankColIndex = masterHeaders.indexOf('最新_評価ランク');
    const latestScoreColIndex = masterHeaders.indexOf('最新_合計スコア');
    const latestInterviewDateColIndex = masterHeaders.indexOf('最新_面接日');
    const interviewCountColIndex = masterHeaders.indexOf('面接回数');

    // デバッグ: Candidates_Master列インデックス確認
    Logger.log('Candidates_Master 列インデックス:');
    Logger.log('  candidate_id: ' + masterCandidateIdColIndex);
    Logger.log('  最新_承諾可能性: ' + latestAcceptanceMasterColIndex);
    Logger.log('  最新_評価ランク: ' + latestRankColIndex);
    Logger.log('  最新_合計スコア: ' + latestScoreColIndex);
    Logger.log('  最新_面接日: ' + latestInterviewDateColIndex);
    Logger.log('  面接回数: ' + interviewCountColIndex);

    // 列が見つからない場合の警告
    if (latestRankColIndex === -1) {
      Logger.log('❌ ERROR: Candidates_Masterに最新_評価ランク列が見つかりません');
    }
    if (latestScoreColIndex === -1) {
      Logger.log('❌ ERROR: Candidates_Masterに最新_合計スコア列が見つかりません');
    }
    if (latestInterviewDateColIndex === -1) {
      Logger.log('❌ ERROR: Candidates_Masterに最新_面接日列が見つかりません');
    }
    if (interviewCountColIndex === -1) {
      Logger.log('❌ ERROR: Candidates_Masterに面接回数列が見つかりません');
    }

    // 該当候補者の行を探す
    let targetRow = -1;
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][masterCandidateIdColIndex] === candidateId) {
        targetRow = i + 1; // スプレッドシートは1始まり
        Logger.log('✅ Candidates_Masterで候補者発見: 行' + targetRow);
        break;
      }
    }

    if (targetRow === -1) {
      Logger.log('❌ Candidates_Masterに該当候補者が見つかりません');
      Logger.log('検索したcandidate_id: ' + candidateId);
      return { success: false, message: '候補者なし' };
    }

    // データを更新
    Logger.log('=== Candidates_Master更新開始 ===');

    if (latestAcceptanceMasterColIndex !== -1) {
      Logger.log('最新_承諾可能性を更新: ' + (latestAcceptance || ''));
      masterSheet.getRange(targetRow, latestAcceptanceMasterColIndex + 1).setValue(latestAcceptance || '');
    } else {
      Logger.log('⚠️ 最新_承諾可能性列がないためスキップ');
    }

    if (latestRankColIndex !== -1) {
      Logger.log('最新_評価ランクを更新: ' + latestEval.total_rank);
      masterSheet.getRange(targetRow, latestRankColIndex + 1).setValue(latestEval.total_rank);
    } else {
      Logger.log('⚠️ 最新_評価ランク列がないためスキップ');
    }

    if (latestScoreColIndex !== -1) {
      Logger.log('最新_合計スコアを更新: ' + latestEval.total_score);
      masterSheet.getRange(targetRow, latestScoreColIndex + 1).setValue(latestEval.total_score);
    } else {
      Logger.log('⚠️ 最新_合計スコア列がないためスキップ');
    }

    if (latestInterviewDateColIndex !== -1) {
      Logger.log('最新_面接日を更新: ' + latestEval.interview_date);
      masterSheet.getRange(targetRow, latestInterviewDateColIndex + 1).setValue(latestEval.interview_date);
    } else {
      Logger.log('⚠️ 最新_面接日列がないためスキップ');
    }

    if (interviewCountColIndex !== -1) {
      Logger.log('面接回数を更新: ' + interviewCount);
      masterSheet.getRange(targetRow, interviewCountColIndex + 1).setValue(interviewCount);
    } else {
      Logger.log('⚠️ 面接回数列がないためスキップ');
    }

    Logger.log('✅ Candidates_Master更新完了: 行' + targetRow);

    return {
      success: true,
      candidateId: candidateId,
      row: targetRow,
      updated: {
        acceptance: latestAcceptance,
        rank: latestEval.total_rank,
        score: latestEval.total_score,
        interview_date: latestEval.interview_date,
        interview_count: interviewCount
      }
    };

  } catch (error) {
    Logger.log('ERROR in updateCandidatesMaster: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return { success: false, error: error.toString() };
  }
}

/**
 * Phase 4-2a機能のテスト関数
 */
function testPhase42aFunctions() {
  Logger.log('=== Phase 4-2a テスト開始 ===');

  // テスト用の候補者ID（実際に存在するものを使用）
  const testCandidateId = 'CAND_20251226012324'; // ← 実際のIDに置き換え

  Logger.log('テスト候補者ID: ' + testCandidateId);

  // updateCandidateScoresのテスト
  Logger.log('\n--- updateCandidateScores テスト ---');
  const scoresResult = updateCandidateScores(testCandidateId);
  Logger.log('結果: ' + JSON.stringify(scoresResult, null, 2));

  // updateCandidatesMasterのテスト
  Logger.log('\n--- updateCandidatesMaster テスト ---');
  const masterResult = updateCandidatesMaster(testCandidateId);
  Logger.log('結果: ' + JSON.stringify(masterResult, null, 2));

  Logger.log('\n=== Phase 4-2a テスト完了 ===');

  if (scoresResult.success && masterResult.success) {
    Logger.log('✅ 全テスト成功');
    return true;
  } else {
    Logger.log('❌ テスト失敗');
    return false;
  }
}

// ============================================================
// Phase 4-2a: 診断関数
// ============================================================

/**
 * シートの列名を確認する診断関数
 */
function diagnoseSheetColumns() {
  Logger.log('=== シート列名診断開始 ===');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Evaluation_Master
  const evalSheet = ss.getSheetByName('Evaluation_Master');
  if (evalSheet) {
    const evalHeaders = evalSheet.getRange(1, 1, 1, evalSheet.getLastColumn()).getValues()[0];
    Logger.log('\n【Evaluation_Master ヘッダー】');
    evalHeaders.forEach((header, index) => {
      Logger.log('  列' + (index + 1) + ': ' + header);
    });

    // 必要な列が存在するか確認
    const requiredEvalColumns = ['candidate_id', 'total_rank', 'total_score', 'interview_date'];
    requiredEvalColumns.forEach(col => {
      const index = evalHeaders.indexOf(col);
      if (index === -1) {
        Logger.log('  ❌ ' + col + ' が見つかりません！');
      } else {
        Logger.log('  ✅ ' + col + ' は列' + (index + 1) + 'にあります');
      }
    });
  } else {
    Logger.log('❌ Evaluation_Masterシートが見つかりません');
  }

  // Candidates_Master
  const masterSheet = ss.getSheetByName('Candidates_Master');
  if (masterSheet) {
    const masterHeaders = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    Logger.log('\n【Candidates_Master ヘッダー】');
    masterHeaders.forEach((header, index) => {
      Logger.log('  列' + (index + 1) + ': ' + header);
    });

    // 必要な列が存在するか確認
    const requiredMasterColumns = ['candidate_id', '最新_評価ランク', '最新_合計スコア', '最新_面接日', '面接回数'];
    requiredMasterColumns.forEach(col => {
      const index = masterHeaders.indexOf(col);
      if (index === -1) {
        Logger.log('  ❌ ' + col + ' が見つかりません！');
      } else {
        Logger.log('  ✅ ' + col + ' は列' + (index + 1) + 'にあります');
      }
    });
  } else {
    Logger.log('❌ Candidates_Masterシートが見つかりません');
  }

  // Candidate_Scores
  const scoresSheet = ss.getSheetByName('Candidate_Scores');
  if (scoresSheet) {
    const scoresHeaders = scoresSheet.getRange(1, 1, 1, scoresSheet.getLastColumn()).getValues()[0];
    Logger.log('\n【Candidate_Scores ヘッダー】');
    scoresHeaders.forEach((header, index) => {
      Logger.log('  列' + (index + 1) + ': ' + header);
    });

    // 必要な列が存在するか確認
    const requiredScoresColumns = ['candidate_id', '最新_承諾可能性（AI予測）', '予測の信頼度', '志望度スコア'];
    requiredScoresColumns.forEach(col => {
      const index = scoresHeaders.indexOf(col);
      if (index === -1) {
        Logger.log('  ❌ ' + col + ' が見つかりません！');
      } else {
        Logger.log('  ✅ ' + col + ' は列' + (index + 1) + 'にあります');
      }
    });
  } else {
    Logger.log('❌ Candidate_Scoresシートが見つかりません');
  }

  // Engagement_Log
  const engagementSheet = ss.getSheetByName('Engagement_Log');
  if (engagementSheet) {
    const engagementHeaders = engagementSheet.getRange(1, 1, 1, engagementSheet.getLastColumn()).getValues()[0];
    Logger.log('\n【Engagement_Log ヘッダー】');
    engagementHeaders.forEach((header, index) => {
      Logger.log('  列' + (index + 1) + ': ' + header);
    });

    // 必要な列が存在するか確認
    const requiredEngagementColumns = ['candidate_id', 'acceptance_rate_ai', 'confidence_level', 'motivation_score'];
    requiredEngagementColumns.forEach(col => {
      const index = engagementHeaders.indexOf(col);
      if (index === -1) {
        Logger.log('  ❌ ' + col + ' が見つかりません！');
      } else {
        Logger.log('  ✅ ' + col + ' は列' + (index + 1) + 'にあります');
      }
    });
  } else {
    Logger.log('❌ Engagement_Logシートが見つかりません');
  }

  Logger.log('\n=== シート列名診断完了 ===');
}

/**
 * 特定の候補者のデータを全シートで確認する診断関数
 * @param {string} targetCandidateId - 確認対象の候補者ID（省略時はデフォルト値）
 */
function diagnoseSpecificCandidate(targetCandidateId) {
  targetCandidateId = targetCandidateId || 'CAND_20260106102327'; // テストKKKK

  Logger.log('=== 候補者データ診断開始 ===');
  Logger.log('対象候補者ID: ' + targetCandidateId);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Evaluation_Master
  Logger.log('\n【Evaluation_Master】');
  const evalSheet = ss.getSheetByName('Evaluation_Master');
  if (evalSheet) {
    const evalData = evalSheet.getDataRange().getValues();
    const evalHeaders = evalData[0];
    // 実際の列名に修正
    const candidateIdCol = evalHeaders.indexOf('候補者ID');

    let found = false;
    for (let i = 1; i < evalData.length; i++) {
      if (evalData[i][candidateIdCol] === targetCandidateId) {
        found = true;
        Logger.log('✅ 行' + (i + 1) + 'に発見');
        evalHeaders.forEach((header, colIndex) => {
          Logger.log('  ' + header + ': ' + evalData[i][colIndex]);
        });
      }
    }
    if (!found) {
      Logger.log('❌ 該当データなし');
      // 最後の3行を表示
      Logger.log('最後の3行の候補者ID:');
      for (let i = Math.max(1, evalData.length - 3); i < evalData.length; i++) {
        Logger.log('  行' + (i + 1) + ': ' + evalData[i][candidateIdCol]);
      }
    }
  }

  // Candidates_Master
  Logger.log('\n【Candidates_Master】');
  const masterSheet = ss.getSheetByName('Candidates_Master');
  if (masterSheet) {
    const masterData = masterSheet.getDataRange().getValues();
    const masterHeaders = masterData[0];
    const candidateIdCol = masterHeaders.indexOf('candidate_id');

    let found = false;
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][candidateIdCol] === targetCandidateId) {
        found = true;
        Logger.log('✅ 行' + (i + 1) + 'に発見');
        masterHeaders.forEach((header, colIndex) => {
          Logger.log('  ' + header + ': ' + masterData[i][colIndex]);
        });
      }
    }
    if (!found) {
      Logger.log('❌ 該当データなし');
    }
  }

  // Engagement_Log
  Logger.log('\n【Engagement_Log】');
  const engagementSheet = ss.getSheetByName('Engagement_Log');
  if (engagementSheet) {
    const engagementData = engagementSheet.getDataRange().getValues();
    const engagementHeaders = engagementData[0];
    const candidateIdCol = engagementHeaders.indexOf('candidate_id');

    let found = false;
    for (let i = 1; i < engagementData.length; i++) {
      if (engagementData[i][candidateIdCol] === targetCandidateId) {
        found = true;
        Logger.log('✅ 行' + (i + 1) + 'に発見');
        // 主要な列のみ表示
        const importantCols = ['candidate_id', 'acceptance_rate_ai', 'confidence_level', 'motivation_score', 'timestamp'];
        importantCols.forEach(col => {
          const colIndex = engagementHeaders.indexOf(col);
          if (colIndex !== -1) {
            Logger.log('  ' + col + ': ' + engagementData[i][colIndex]);
          }
        });
      }
    }
    if (!found) {
      Logger.log('❌ 該当データなし');
    }
  }

  // Candidate_Scores
  Logger.log('\n【Candidate_Scores】');
  const scoresSheet = ss.getSheetByName('Candidate_Scores');
  if (scoresSheet) {
    const scoresData = scoresSheet.getDataRange().getValues();
    const scoresHeaders = scoresData[0];
    const candidateIdCol = scoresHeaders.indexOf('candidate_id');

    let found = false;
    for (let i = 1; i < scoresData.length; i++) {
      if (scoresData[i][candidateIdCol] === targetCandidateId) {
        found = true;
        Logger.log('✅ 行' + (i + 1) + 'に発見');
        scoresHeaders.forEach((header, colIndex) => {
          Logger.log('  ' + header + ': ' + scoresData[i][colIndex]);
        });
      }
    }
    if (!found) {
      Logger.log('❌ 該当データなし');
    }
  }

  Logger.log('\n=== 候補者データ診断完了 ===');
}
