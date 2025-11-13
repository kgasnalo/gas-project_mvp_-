/**
 * メイン実行関数
 */
function setupAllSheets() {
  const startTime = new Date();
  Logger.log('=== 採用参謀AI：スプレッドシート自動作成開始 ===');
  
  try {
    const steps = [
      { name: 'シート作成', func: createAllSheets },
      { name: 'Candidates_Master設定', func: setupCandidatesMaster },
      { name: 'Evaluation_Log設定', func: setupEvaluationLog },
      { name: 'Engagement_Log設定', func: setupEngagementLog },
      { name: 'Company_Assets設定', func: setupCompanyAssets },
      { name: 'Competitor_DB設定', func: setupCompetitorDB },
      { name: 'その他シート設定', func: setupOtherSheets },
      { name: 'Dashboard設定（完全版）', func: setupDashboardComplete }, // 追加
      { name: 'データ検証設定', func: setupAllDataValidation },
      { name: '条件付き書式設定', func: setupAllConditionalFormatting },
      { name: '初期データ投入', func: insertAllInitialData }
    ];
    
    steps.forEach((step, index) => {
      Logger.log(`[${index + 1}/${steps.length}] ${step.name}を開始...`);
      const stepStartTime = new Date();
      
      step.func();
      
      const stepEndTime = new Date();
      const stepDuration = (stepEndTime - stepStartTime) / 1000;
      Logger.log(`[${index + 1}/${steps.length}] ${step.name}が完了しました（${stepDuration}秒）`);
      
      const elapsedTime = (new Date() - startTime) / 1000;
      if (elapsedTime > 300) {
        Logger.log(`⚠️ 実行時間が5分を超えたため、一旦停止します。残りのステップは手動で実行してください。`);
        return;
      }
    });
    
    const endTime = new Date();
    const totalDuration = (endTime - startTime) / 1000;
    Logger.log(`=== 採用参謀AI：スプレッドシート自動作成完了（${totalDuration}秒） ===`);
    
    SpreadsheetApp.getUi().alert(
      '✅ スプレッドシートの自動作成が完了しました！\n\n' +
      '実行時間: ' + totalDuration + '秒\n' +
      '作成されたシート: 15シート（Dashboard完全版を含む）\n\n' +
      '次のステップ:\n' +
      '1. Dashboardで候補者一覧を確認してください\n' +
      '2. サンプルデータが正しく表示されているか確認してください\n' +
      '3. Difyとの連携設定を行ってください'
    );
    
  } catch (error) {
    Logger.log(`❌ エラーが発生しました: ${error.message}`);
    Logger.log(error.stack);
    
    SpreadsheetApp.getUi().alert(
      '❌ エラーが発生しました\n\n' +
      'エラー内容: ' + error.message + '\n\n' +
      '詳細はログを確認してください（表示 → ログ）'
    );
    
    throw error;
  }
}

/**
 * スプレッドシート起動時に実行（カスタムメニュー作成）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // ========== メニュー1: 📧 アンケート送信 ==========
  ui.createMenu('📧 アンケート送信')
    .addItem('初回面談アンケートを送信', 'showSendFirstInterviewSurvey')
    .addItem('社員面談アンケートを送信', 'showSendEmployeeInterviewSurvey')
    .addItem('2次面接アンケートを送信', 'showSendSecondInterviewSurvey')
    .addItem('内定後アンケートを送信', 'showSendFinalInterviewSurvey')
    .addSeparator()
    .addItem('📊 送信履歴を確認', 'showSendHistory')
    .addItem('📈 今日の送信状況', 'showTodaySendCount') // 【新規追加】
    .addToUi();

  // ========== メニュー2: 🤖 採用参謀AI ==========
  ui.createMenu('🤖 採用参謀AI')
    .addItem('📋 全シートを自動作成', 'setupAllSheets')
    .addItem('📊 Dashboardを再生成', 'setupDashboardComplete')
    .addSubMenu(ui.createMenu('🔗 Dify連携')
      .addItem('⚙️ API設定', 'setupDifyApiSettings')
      .addItem('📡 Webhook URL確認', 'showWebhookUrl')
      .addItem('🧪 Webhookテスト', 'testWebhook'))
    .addSeparator()
    .addItem('🔄 データ検証を再設定', 'setupAllDataValidation')
    .addItem('🎨 条件付き書式を再設定', 'setupAllConditionalFormatting')
    .addItem('📊 初期データを再投入', 'insertAllInitialData')
    .addItem('🔍 初期データを検証', 'validateTestData')
    .addSeparator()
    .addItem('📈 回答速度を一括計算', 'calculateAllResponseSpeeds') // 【Phase 2 Step 3追加】
    .addSubMenu(ui.createMenu('🧪 テストデータ')
      .addItem('📊 テストデータ状況を確認', 'checkTestDataStatus')
      .addItem('➕ テストデータを生成', 'generateAllTestData')
      .addItem('🗑️ テストデータをクリア', 'clearAllTestData')
      .addSeparator()
      .addItem('🔍 データ構造を検証（列数）', 'validateAllTestData')
      .addItem('🔬 データ構造を詳細検証', 'validateTestDataStructure'))
    .addSeparator()
    .addItem('🔍 システム状態を確認', 'showSystemStatus') // 【新規追加】
    .addItem('ℹ️ バージョン情報', 'showVersionInfo')
    .addToUi();
}

/**
 * バージョン情報を表示
 */
function showVersionInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // 候補者数を取得
  const masterSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);
  const candidateCount = masterSheet ? masterSheet.getLastRow() - 1 : 0;

  // 評価ログ数を取得
  const evalSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EVALUATION_LOG);
  const evalLogCount = evalSheet ? evalSheet.getLastRow() - 1 : 0;

  // 送信履歴数を取得
  const sendLogSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);
  const sendLogCount = sendLogSheet ? sendLogSheet.getLastRow() - 1 : 0;

  const message =
    '【採用参謀AI - システム情報】\n\n' +
    '📌 バージョン: v1.2.0\n' +
    '📅 最終更新日: 2025-11-12\n\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '📊 データ統計\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    `・シート数: ${sheets.length}\n` +
    `・候補者数: ${candidateCount}名\n` +
    `・評価ログ数: ${evalLogCount}件\n` +
    `・アンケート送信数: ${sendLogCount}件\n\n` +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '🔧 主要機能\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '✅ アンケート自動送信\n' +
    '✅ アンケート種別管理（Phase 2実装）\n' +
    '✅ 合格可能性AI評価\n' +
    '✅ 承諾可能性AI予測\n' +
    '✅ リスク検知\n' +
    '✅ ネクストアクション提案\n' +
    '✅ Dashboard自動更新\n\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '📚 ドキュメント\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    'GitHub: kgasnalo/gas-project_mvp_-\n' +
    'Branch: claude/add-column-constants-config';

  SpreadsheetApp.getUi().alert(
    'バージョン情報',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * システム状態を確認
 */
function showSystemStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 必須シートの存在チェック
  const requiredSheets = [
    'Dashboard',
    'Candidates_Master',
    'Evaluation_Log',
    'Engagement_Log',
    'Survey_Response',
    'Survey_Send_Log',
    'Evidence',
    'Risk',
    'NextQ',
    'Contact_History'
  ];

  let allSheetsExist = true;
  let missingSheets = [];

  requiredSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      allSheetsExist = false;
      missingSheets.push(sheetName);
    }
  });

  // メッセージ作成
  let statusIcon = allSheetsExist ? '✅' : '❌';
  let statusMessage = allSheetsExist ?
    'すべての必須シートが正常です' :
    `以下のシートが見つかりません:\n${missingSheets.join(', ')}`;

  const sheets = ss.getSheets();

  const message =
    '【🔍 システム状態確認】\n\n' +
    `${statusIcon} ${statusMessage}\n\n` +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '📊 シート構成\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    `・総シート数: ${sheets.length}\n` +
    `・必須シート: ${requiredSheets.length - missingSheets.length}/${requiredSheets.length}\n\n` +
    (missingSheets.length > 0 ?
      '⚠️ 不足しているシート:\n' + missingSheets.map(s => `  - ${s}`).join('\n') + '\n\n' :
      '') +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '🔧 推奨アクション\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    (missingSheets.length > 0 ?
      '「📋 全シートを自動作成」を実行して\n不足シートを作成してください' :
      'システムは正常に動作しています');

  SpreadsheetApp.getUi().alert(
    'システム状態',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}