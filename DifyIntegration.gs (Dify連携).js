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
 * Phase 1-1: Dify Webhookエンドポイント（テストモード）
 * 目的: データ受信確認とProcessing_Logへの記録
 */
function doPost(e) {
  const startTime = new Date();

  try {
    // リクエストボディの取得
    const requestBody = e.postData ? e.postData.contents : null;

    if (!requestBody) {
      throw new Error('リクエストボディが空です');
    }

    // JSONパース
    const data = JSON.parse(requestBody);

    // ログ出力
    Logger.log('=== Phase 1-1 テストモード ===');
    Logger.log('受信時刻: ' + new Date().toISOString());
    Logger.log('データサイズ: ' + requestBody.length + ' bytes');
    Logger.log('candidate_name: ' + (data.validated_input ? data.validated_input.candidate_name : 'なし'));
    Logger.log('transcript有無: ' + (data.transcript ? 'あり(' + data.transcript.length + '文字)' : 'なし'));

    // Processing_Logに記録
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Processing_Log');

    if (sheet) {
      const logRow = [
        new Date(),                                    // A: timestamp
        'Phase1-1_Test',                              // B: phase
        data.validated_input ? data.validated_input.candidate_name : 'Unknown',  // C: candidate
        'webhook_test',                                // D: event
        'SUCCESS',                                     // E: status
        JSON.stringify(data.validated_input || {}).substring(0, 500),  // F: input_data
        'transcript: ' + (data.transcript ? data.transcript.length + '文字' : 'なし'),  // G: output_data
        '実行時間: ' + ((new Date() - startTime) / 1000).toFixed(2) + '秒'  // H: notes
      ];

      sheet.appendRow(logRow);
      Logger.log('✅ Processing_Logに記録完了');
    } else {
      Logger.log('⚠️ Processing_Logシートが見つかりません');
    }

    // 成功レスポンス
    const response = {
      success: true,
      mode: 'TEST_MODE',
      message: 'Phase 1-1: データ受信成功（テストモード）',
      received: {
        candidate_id: data.validated_input ? data.validated_input.candidate_id : null,
        candidate_name: data.validated_input ? data.validated_input.candidate_name : null,
        recruit_type: data.validated_input ? data.validated_input.recruit_type : null,
        selection_phase: data.validated_input ? data.validated_input.selection_phase : null,
        has_transcript: !!data.transcript,
        transcript_length: data.transcript ? data.transcript.length : 0
      },
      timestamp: new Date().toISOString(),
      execution_time_seconds: ((new Date() - startTime) / 1000).toFixed(2)
    };

    return ContentService
      .createTextOutput(JSON.stringify(response, null, 2))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('❌ エラー発生: ' + error.message);
    Logger.log('スタック: ' + error.stack);

    // エラーレスポンス
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        mode: 'TEST_MODE',
        error: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString()
      }, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
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
    data.workflow_id || ''                  // AB
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
