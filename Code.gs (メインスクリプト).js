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
 * カスタムメニューの追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 採用参謀AI')
    .addItem('📋 全シートを自動作成', 'setupAllSheets')
    .addSeparator()
    .addItem('📊 Dashboardを再生成', 'setupDashboardComplete') // 追加
    .addSeparator()
    .addItem('🔄 データ検証を再設定', 'setupAllDataValidation')
    .addItem('🎨 条件付き書式を再設定', 'setupAllConditionalFormatting')
    .addItem('📊 初期データを再投入', 'insertAllInitialData')
    .addSeparator()
    .addItem('ℹ️ バージョン情報', 'showVersionInfo')
    .addToUi();
}