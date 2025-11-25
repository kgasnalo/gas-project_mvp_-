/**
 * ダッシュボードのチャート作成（Phase 4-1）
 *
 * Dashboardシートに以下のチャートを作成します:
 * 1. 候補者別承諾可能性（棒グラフ）
 * 2. フェーズ別スコア推移（折れ線グラフ）
 * 3. AI予測 vs 人間の直感（散布図）
 * 4. 候補者ステータス分布（円グラフ）
 *
 * @version 1.0
 * @date 2025-11-25
 */

/**
 * ダッシュボードの全チャートを作成（メイン関数）
 */
function setupDashboardCharts() {
  Logger.log('📊 チャート作成開始...');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  if (!sheet) {
    Logger.log('❌ Dashboardシートが見つかりません');
    return;
  }

  // 既存のチャートをクリア
  const charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));

  // 1. 候補者別承諾可能性（棒グラフ）
  createCandidateBarChart(sheet);

  // 2. フェーズ別スコア推移（折れ線グラフ）
  createPhaseLineChart(sheet);

  // 3. AI予測 vs 人間の直感（散布図）
  createAIvsHumanScatterChart(sheet);

  // 4. 候補者ステータス分布（円グラフ）
  createStatusPieChart(sheet);

  Logger.log('✅ チャート作成完了');
}

/**
 * 1. 候補者別承諾可能性（棒グラフ）
 *
 * 配置: J11:N25
 * 氏名と承諾可能性を表示
 */
function createCandidateBarChart(sheet) {
  Logger.log('  📊 候補者別承諾可能性（棒グラフ）を作成中...');

  // データ範囲: C13:D27（氏名と承諾可能性）
  const dataRange = sheet.getRange('C13:D27');

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(dataRange)
    .setPosition(11, 10, 0, 0) // J11セル
    .setOption('title', '候補者別承諾可能性')
    .setOption('width', 500)
    .setOption('height', 400)
    .setOption('hAxis', {
      title: '承諾可能性（%）',
      minValue: 0,
      maxValue: 100,
      format: '0"%"'
    })
    .setOption('vAxis', {
      title: '候補者名'
    })
    .setOption('colors', ['#4285f4'])
    .setOption('legend', { position: 'none' })
    .setOption('chartArea', { width: '65%', height: '85%' })
    .build();

  sheet.insertChart(chart);

  Logger.log('  ✅ 棒グラフ作成完了');
}

/**
 * 2. ステータス別人数分布（棒グラフ）
 *
 * 配置: J29:N43
 * ステータスごとの候補者数を表示
 */
function createPhaseLineChart(sheet) {
  Logger.log('  📊 ステータス別人数分布（棒グラフ）を作成中...');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const candidatesMaster = ss.getSheetByName('Candidates_Master');
  if (!candidatesMaster) {
    Logger.log('  ⚠️ Candidates_Masterシートが見つかりません');
    return;
  }

  // Candidates_MasterのステータスをQUERYで集計
  // 一時的なデータ範囲をDashboardシートに作成
  const tempDataRange = sheet.getRange('P29:Q40');
  tempDataRange.clearContent();

  // ヘッダー
  sheet.getRange('P29').setValue('ステータス');
  sheet.getRange('Q29').setValue('人数');

  // QUERYで集計
  sheet.getRange('P30').setFormula(
    '=QUERY(Candidates_Master!C:C, ' +
    '"SELECT C, COUNT(C) ' +
    'WHERE C IS NOT NULL ' +
    'GROUP BY C ' +
    'ORDER BY COUNT(C) DESC", 1)'
  );

  // データ範囲を取得
  SpreadsheetApp.flush(); // 数式を先に実行
  const dataRange = sheet.getRange('P29:Q40');

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    .setPosition(29, 10, 0, 0) // J29セル
    .setOption('title', 'ステータス別候補者数')
    .setOption('width', 500)
    .setOption('height', 350)
    .setOption('hAxis', {
      title: 'ステータス',
      slantedText: true,
      slantedTextAngle: 45
    })
    .setOption('vAxis', {
      title: '候補者数（人）',
      minValue: 0,
      format: '0'
    })
    .setOption('colors', ['#34a853'])
    .setOption('legend', { position: 'none' })
    .setOption('chartArea', { width: '75%', height: '65%' })
    .setOption('bar', { groupWidth: '70%' })
    .build();

  sheet.insertChart(chart);

  Logger.log('  ✅ 棒グラフ作成完了');
}

/**
 * 3. AI予測 vs 人間の直感（散布図）
 *
 * 配置: J45:N60
 * Candidates_Masterから直接取得
 */
function createAIvsHumanScatterChart(sheet) {
  Logger.log('  📊 AI予測 vs 人間の直感（散布図）を作成中...');

  // Dashboardシートに一時データを作成（A59:C73のデータを使用）
  // setupDashboardAIComparison()で既に作成されているデータを利用
  const dataRange = sheet.getRange('B59:C73');

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.SCATTER)
    .addRange(dataRange)
    .setPosition(45, 10, 0, 0) // J45セル
    .setOption('title', 'AI予測 vs 人間の直感')
    .setOption('width', 500)
    .setOption('height', 350)
    .setOption('hAxis', {
      title: 'AI予測（%）',
      minValue: 0,
      maxValue: 100,
      format: '0"%"'
    })
    .setOption('vAxis', {
      title: '人間の直感（%）',
      minValue: 0,
      maxValue: 100,
      format: '0"%"'
    })
    .setOption('legend', { position: 'none' })
    .setOption('chartArea', { width: '70%', height: '70%' })
    .setOption('pointSize', 5)
    .setOption('colors', ['#4285f4'])
    .setOption('trendlines', {
      0: {
        type: 'linear',
        color: '#ea4335',
        lineWidth: 2,
        opacity: 0.5,
        showR2: true,
        visibleInLegend: true
      }
    })
    .build();

  sheet.insertChart(chart);

  Logger.log('  ✅ 散布図作成完了');
}

/**
 * 4. 承諾可能性分布（円グラフ）
 *
 * 配置: J62:N75
 * Candidates_Masterから直接集計
 */
function createStatusPieChart(sheet) {
  Logger.log('  📊 承諾可能性分布（円グラフ）を作成中...');

  // 一時データ範囲をDashboardシートに作成
  const tempDataRange = sheet.getRange('P62:Q66');
  tempDataRange.clearContent();

  // ヘッダー
  sheet.getRange('P62').setValue('ステータス');
  sheet.getRange('Q62').setValue('人数');

  // データ行（承諾可能性の分布）
  const data = [
    ['高確率（80点以上）', '=COUNTIF(Candidates_Master!R:R,">=80")'],
    ['やや高（70-79点）', '=COUNTIFS(Candidates_Master!R:R,">=70",Candidates_Master!R:R,"<80")'],
    ['標準（60-69点）', '=COUNTIFS(Candidates_Master!R:R,">=60",Candidates_Master!R:R,"<70")'],
    ['要注意（60点未満）', '=COUNTIF(Candidates_Master!R:R,"<60")']
  ];

  sheet.getRange('P63:Q66').setValues(data);

  // データ範囲を取得
  SpreadsheetApp.flush(); // 数式を先に実行
  const dataRange = sheet.getRange('P62:Q66');

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataRange)
    .setPosition(62, 10, 0, 0) // J62セル
    .setOption('title', '承諾可能性分布')
    .setOption('width', 500)
    .setOption('height', 300)
    .setOption('is3D', false)
    .setOption('pieHole', 0.4) // ドーナツグラフ
    .setOption('legend', { position: 'right' })
    .setOption('chartArea', { width: '90%', height: '80%' })
    .setOption('slices', {
      0: { color: '#34a853' },  // 高確率: 緑
      1: { color: '#93c47d' },  // やや高: 薄い緑
      2: { color: '#fbbc04' },  // 標準: 黄
      3: { color: '#ea4335' }   // 要注意: 赤
    })
    .setOption('pieSliceText', 'percentage')
    .setOption('pieSliceTextStyle', {
      fontSize: 12,
      bold: true
    })
    .build();

  sheet.insertChart(chart);

  Logger.log('  ✅ 円グラフ作成完了');
}

/**
 * 特定のチャートのみを再作成（デバッグ用）
 */
function recreateSpecificChart(chartType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  if (!sheet) return;

  switch (chartType) {
    case 'bar':
      createCandidateBarChart(sheet);
      break;
    case 'line':
      createPhaseLineChart(sheet);
      break;
    case 'scatter':
      createAIvsHumanScatterChart(sheet);
      break;
    case 'pie':
      createStatusPieChart(sheet);
      break;
    default:
      Logger.log('❌ 無効なチャートタイプ: ' + chartType);
  }
}
