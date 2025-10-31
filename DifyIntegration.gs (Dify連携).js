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
 * DifyからのWebhookを受信
 * POST: /exec
 * Body: JSON形式のデータ
 *
 * 【Difyからのデータ形式（想定）】
 * {
 *   "type": "evaluation" | "engagement",
 *   "candidate_id": "C001",
 *   "data": { ... }
 * }
 */
function doPost(e) {
  try {
    // リクエストボディを取得
    const requestBody = e.postData.contents;
    Logger.log('Received webhook: ' + requestBody);

    // JSONをパース
    const data = JSON.parse(requestBody);

    // データタイプに応じて処理を分岐
    switch (data.type) {
      case 'evaluation':
        handleEvaluationData(data);
        break;
      case 'engagement':
        handleEngagementData(data);
        break;
      case 'evidence':
        handleEvidenceData(data);
        break;
      case 'risk':
        handleRiskData(data);
        break;
      case 'next_q':
        handleNextQData(data);
        break;
      case 'acceptance_story':
        handleAcceptanceStoryData(data);
        break;
      case 'competitor_comparison':
        handleCompetitorComparisonData(data);
        break;
      default:
        throw new Error('Unknown data type: ' + data.type);
    }

    // 成功レスポンスを返す
    return ContentService.createTextOutput(
      JSON.stringify({
        status: 'success',
        message: 'Data processed successfully',
        timestamp: new Date().toISOString()
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('Error in doPost: ' + error.message);
    Logger.log(error.stack);

    // エラーレスポンスを返す
    return ContentService.createTextOutput(
      JSON.stringify({
        status: 'error',
        message: error.message,
        timestamp: new Date().toISOString()
      })
    ).setMimeType(ContentService.MimeType.JSON);
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
