/**
 * スプレッドシート内ダッシュボード実装（Phase 4-1）
 *
 * 指示書に基づいて、Dashboard_DataとDashboardシートを作成し、
 * QUERY関数による自動集計、条件付き書式、チャートを設定します。
 *
 * 実装内容:
 * - Dashboard_Dataシート: 中間データ用（QUERY関数の結果を格納）
 * - Dashboardシート: メインダッシュボード（可視化とKPI）
 *
 * @version 1.0
 * @date 2025-11-25
 */

/**
 * ダッシュボード全体をセットアップ（メイン関数）
 *
 * この関数を実行することで、Dashboard_DataとDashboardシートが
 * 自動的に作成・設定されます。
 */
function setupDashboardPhase4() {
  try {
    Logger.log('====================================');
    Logger.log('📊 ダッシュボードセットアップ開始');
    Logger.log('====================================');

    // 1. Dashboard_Dataシートを作成
    setupDashboardDataSheet();

    // 2. Dashboardシートを作成
    setupDashboardSheet();

    // 3. 条件付き書式を設定
    setupDashboardConditionalFormats();

    // 4. チャートを作成
    setupDashboardCharts();

    Logger.log('====================================');
    Logger.log('✅ ダッシュボードセットアップ完了');
    Logger.log('====================================');

    // 完了メッセージ
    SpreadsheetApp.getUi().alert(
      '✅ ダッシュボードセットアップ完了\n\n' +
      '以下のシートが作成されました:\n' +
      '- Dashboard_Data (中間データ)\n' +
      '- Dashboard (メインダッシュボード)\n\n' +
      'Dashboardシートを開いてください。'
    );

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    throw error;
  }
}

/**
 * Dashboard_Dataシートをセットアップ
 *
 * 中間データを格納するシートを作成し、QUERY関数で
 * Engagement_Logから最新データを集計します。
 */
function setupDashboardDataSheet() {
  Logger.log('📝 Dashboard_Dataシートを作成中...');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存のシートを削除（再作成のため）
  let sheet = ss.getSheetByName('Dashboard_Data');
  if (sheet) {
    ss.deleteSheet(sheet);
  }

  // 新規作成
  sheet = ss.insertSheet('Dashboard_Data');

  // シートをEngagement_Logの後ろに移動
  const engagementSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ENGAGEMENT_LOG);
  if (engagementSheet) {
    sheet.moveAfter(engagementSheet);
  }

  // === セクション1: 最新承諾可能性（A1:E100） ===
  setupLatestAcceptanceData(sheet);

  // === セクション2: フェーズ別スコア推移（G1:K100） ===
  setupPhaseScoreData(sheet);

  // === セクション3: AI予測 vs 人間の直感（M1:P100） ===
  setupAIvsHumanData(sheet);

  // === セクション4: 懸念事項集計（R1:S20） ===
  setupConcernData(sheet);

  // === セクション5: ステータス分布（U1:V5） ===
  setupStatusDistributionData(sheet);

  Logger.log('✅ Dashboard_Dataシート作成完了');
}

/**
 * セクション1: 最新承諾可能性データ（A1:E100）
 */
function setupLatestAcceptanceData(sheet) {
  // ヘッダー行
  const headers = ['候補者ID', '最新承諾可能性', '最新フェーズ', '最終更新日', 'コアモチベーション'];
  sheet.getRange('A1:E1').setValues([headers]);
  sheet.getRange('A1:E1')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT);

  // QUERY関数（候補者別最新承諾可能性）
  // Engagement_Logから最新データを取得し、承諾可能性順にソート
  const query = `=QUERY(Engagement_Log!A:U,
    "SELECT B, MAX(H), E, MAX(D), M
     WHERE B IS NOT NULL
     GROUP BY B, E, M
     ORDER BY MAX(H) DESC
     LABEL B '候補者ID', MAX(H) '最新承諾可能性', E '最新フェーズ', MAX(D) '最終更新日', M 'コアモチベーション'",
    1)`;

  sheet.getRange('A2').setFormula(query);

  // 列幅設定
  sheet.setColumnWidth(1, 120); // A: 候補者ID
  sheet.setColumnWidth(2, 150); // B: 最新承諾可能性
  sheet.setColumnWidth(3, 120); // C: 最新フェーズ
  sheet.setColumnWidth(4, 140); // D: 最終更新日
  sheet.setColumnWidth(5, 200); // E: コアモチベーション

  // フォーマット設定
  sheet.getRange('B2:B100').setNumberFormat('0.00"%"');
  sheet.getRange('D2:D100').setNumberFormat('yyyy-mm-dd');
}

/**
 * セクション2: フェーズ別スコア推移データ（G1:K100）
 *
 * 各候補者のフェーズごとの承諾可能性を集計します。
 */
function setupPhaseScoreData(sheet) {
  // ヘッダー行
  const headers = ['候補者ID', '初回面談', '社員面談', '2次面接', '内定後'];
  sheet.getRange('G1:K1').setValues([headers]);
  sheet.getRange('G1:K1')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT);

  // 候補者IDリストを取得（UNIQUE関数）
  sheet.getRange('G2').setFormula('=UNIQUE(Engagement_Log!B2:B)');

  // 各フェーズの最大承諾可能性を取得
  // H列: 初回面談
  sheet.getRange('H2').setFormula(
    '=IFERROR(QUERY(Engagement_Log!B:H, ' +
    '"SELECT MAX(H) WHERE B=\'" & G2 & "\' AND E=\'初回面談\' GROUP BY B LABEL MAX(H) \'\'", 0), "")'
  );

  // I列: 社員面談
  sheet.getRange('I2').setFormula(
    '=IFERROR(QUERY(Engagement_Log!B:H, ' +
    '"SELECT MAX(H) WHERE B=\'" & G2 & "\' AND E=\'社員面談\' GROUP BY B LABEL MAX(H) \'\'", 0), "")'
  );

  // J列: 2次面接
  sheet.getRange('J2').setFormula(
    '=IFERROR(QUERY(Engagement_Log!B:H, ' +
    '"SELECT MAX(H) WHERE B=\'" & G2 & "\' AND E=\'2次面接\' GROUP BY B LABEL MAX(H) \'\'", 0), "")'
  );

  // K列: 内定後
  sheet.getRange('K2').setFormula(
    '=IFERROR(QUERY(Engagement_Log!B:H, ' +
    '"SELECT MAX(H) WHERE B=\'" & G2 & "\' AND E=\'内定後\' GROUP BY B LABEL MAX(H) \'\'", 0), "")'
  );

  // 列幅設定
  sheet.setColumnWidth(7, 120);  // G: 候補者ID
  sheet.setColumnWidth(8, 100);  // H: 初回面談
  sheet.setColumnWidth(9, 100);  // I: 社員面談
  sheet.setColumnWidth(10, 100); // J: 2次面接
  sheet.setColumnWidth(11, 100); // K: 内定後

  // フォーマット設定
  sheet.getRange('H2:K100').setNumberFormat('0.00"%"');
}

/**
 * セクション3: AI予測 vs 人間の直感（M1:P100）
 */
function setupAIvsHumanData(sheet) {
  // ヘッダー行
  const headers = ['候補者ID', 'AI予測', '人間の直感', '乖離'];
  sheet.getRange('M1:P1').setValues([headers]);
  sheet.getRange('M1:P1')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT);

  // QUERY関数（最新のAI予測と人間の直感を取得）
  const query = `=QUERY(Engagement_Log!A:U,
    "SELECT B, G, F
     WHERE B IS NOT NULL AND G IS NOT NULL
     ORDER BY B
     LABEL B '候補者ID', G 'AI予測', F '人間の直感'",
    1)`;

  sheet.getRange('M2').setFormula(query);

  // P列: 乖離計算
  sheet.getRange('P2').setFormula('=IF(AND(N2<>"", O2<>""), ABS(N2-O2), "")');

  // 列幅設定
  sheet.setColumnWidth(13, 120); // M: 候補者ID
  sheet.setColumnWidth(14, 100); // N: AI予測
  sheet.setColumnWidth(15, 120); // O: 人間の直感
  sheet.setColumnWidth(16, 80);  // P: 乖離

  // フォーマット設定
  sheet.getRange('N2:P100').setNumberFormat('0.00"%"');
}

/**
 * セクション4: 懸念事項集計（R1:S20）
 */
function setupConcernData(sheet) {
  // ヘッダー行
  const headers = ['懸念事項', '出現回数'];
  sheet.getRange('R1:S1').setValues([headers]);
  sheet.getRange('R1:S1')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT);

  // QUERY関数（懸念事項を集計）
  const query = `=QUERY(Engagement_Log!N:N,
    "SELECT N, COUNT(N)
     WHERE N IS NOT NULL AND N<>'なし' AND N<>''
     GROUP BY N
     ORDER BY COUNT(N) DESC
     LABEL N '懸念事項', COUNT(N) '回数'",
    1)`;

  sheet.getRange('R2').setFormula(query);

  // 列幅設定
  sheet.setColumnWidth(18, 250); // R: 懸念事項
  sheet.setColumnWidth(19, 100); // S: 出現回数

  // フォーマット設定
  sheet.getRange('S2:S20').setNumberFormat('0');
}

/**
 * セクション5: ステータス分布データ（U1:V5）
 */
function setupStatusDistributionData(sheet) {
  // ヘッダー行
  const headers = ['ステータス', '人数'];
  sheet.getRange('U1:V1').setValues([headers]);
  sheet.getRange('U1:V1')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT);

  // データ行（承諾可能性の分布）
  const data = [
    ['高確率（80点以上）', '=COUNTIF(Dashboard_Data!B:B,">=80")'],
    ['やや高（70-79点）', '=COUNTIFS(Dashboard_Data!B:B,">=70",Dashboard_Data!B:B,"<80")'],
    ['標準（60-69点）', '=COUNTIFS(Dashboard_Data!B:B,">=60",Dashboard_Data!B:B,"<70")'],
    ['要注意（60点未満）', '=COUNTIF(Dashboard_Data!B:B,"<60")']
  ];

  sheet.getRange('U2:V5').setValues(data);

  // 列幅設定
  sheet.setColumnWidth(21, 180); // U: ステータス
  sheet.setColumnWidth(22, 80);  // V: 人数

  // フォーマット設定
  sheet.getRange('V2:V5').setNumberFormat('0');
}

/**
 * Dashboardシートをセットアップ
 *
 * メインダッシュボードを作成し、KPIサマリー、
 * 候補者ランキング、AI予測比較などを表示します。
 */
function setupDashboardSheet() {
  Logger.log('📝 Dashboardシートを作成中...');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存のシートを削除（再作成のため）
  let sheet = ss.getSheetByName('Dashboard');
  if (sheet) {
    ss.deleteSheet(sheet);
  }

  // 新規作成
  sheet = ss.insertSheet('Dashboard');

  // シートを一番左に移動
  sheet.activate();
  ss.moveActiveSheet(1);

  // === ヘッダーセクション（A1:F2） ===
  setupDashboardHeader(sheet);

  // === KPIサマリーセクション（A4:F9） ===
  setupDashboardKPIs(sheet);

  // === 候補者ランキングセクション（A11:F30） ===
  setupDashboardRanking(sheet);

  // === AI予測 vs 人間の直感セクション（A32:E50） ===
  setupDashboardAIComparison(sheet);

  // === 候補者ステータス分布（A52:F65） ===
  // ※チャート作成は別関数で実装

  Logger.log('✅ Dashboardシート作成完了');
}

/**
 * ヘッダーセクション（A1:F2）
 */
function setupDashboardHeader(sheet) {
  // タイトル
  sheet.getRange('A1').setValue('📊 採用ダッシュボード');
  sheet.getRange('A1:F1').merge();
  sheet.getRange('A1')
    .setFontSize(24)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 50);

  // 最終更新日時
  sheet.getRange('A2').setFormula('="最終更新: " & TEXT(NOW(),"yyyy/mm/dd hh:mm")');
  sheet.getRange('A2:F2').merge();
  sheet.getRange('A2')
    .setFontSize(10)
    .setFontColor('#666666')
    .setHorizontalAlignment('center');
}

/**
 * KPIサマリーセクション（A4:F9）
 */
function setupDashboardKPIs(sheet) {
  // セクションヘッダー
  sheet.getRange('A4').setValue('【KPIサマリー】');
  sheet.getRange('A4:F4').merge();
  sheet.getRange('A4')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('left');

  // KPI項目
  const kpiData = [
    ['総候補者数', '=COUNTA(UNIQUE(Engagement_Log!B:B))-1'],
    ['平均承諾可能性', '=ROUND(AVERAGE(Dashboard_Data!B:B),1) & "点"'],
    ['高確率候補者数（80点以上）', '=COUNTIF(Dashboard_Data!B:B,">=80") & "名"'],
    ['要注意候補者数（60点未満）', '=COUNTIF(Dashboard_Data!B:B,"<60") & "名"'],
    ['本日の新規記録', '=COUNTIF(Engagement_Log!D:D,TODAY()) & "件"']
  ];

  sheet.getRange('A5:B9').setValues(kpiData);

  // 列幅設定
  sheet.setColumnWidth(1, 200); // A: ラベル
  sheet.setColumnWidth(2, 150); // B: 値
  sheet.setColumnWidth(3, 150); // C
  sheet.setColumnWidth(4, 150); // D
  sheet.setColumnWidth(5, 150); // E
  sheet.setColumnWidth(6, 150); // F

  // 書式設定
  sheet.getRange('A5:A9').setFontWeight('bold');
  sheet.getRange('B5:B9')
    .setFontSize(16)
    .setHorizontalAlignment('right');

  // 交互の背景色
  for (let i = 5; i <= 9; i++) {
    if (i % 2 === 0) {
      sheet.getRange(`A${i}:B${i}`).setBackground('#f9f9f9');
    }
  }
}

/**
 * 候補者ランキングセクション（A11:F30）
 */
function setupDashboardRanking(sheet) {
  // セクションヘッダー
  sheet.getRange('A11').setValue('【候補者別承諾可能性ランキング】');
  sheet.getRange('A11:F11').merge();
  sheet.getRange('A11')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('left');

  // ヘッダー行
  const headers = ['順位', '候補者ID', '承諾可能性', 'フェーズ', '更新日', 'モチベーション'];
  sheet.getRange('A12:F12').setValues([headers]);
  sheet.getRange('A12:F12')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  // 順位列（1-15）
  for (let i = 1; i <= 15; i++) {
    sheet.getRange(`A${12 + i}`).setValue(i);
  }

  // QUERY関数でデータを抽出
  const query = `=QUERY(Dashboard_Data!A:E,
    "SELECT A, B, C, D, E
     WHERE A IS NOT NULL
     ORDER BY B DESC
     LIMIT 15",
    0)`;

  sheet.getRange('B13').setFormula(query);

  // 書式設定
  sheet.getRange('A13:F27').setBorder(
    true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('C13:C27').setNumberFormat('0.00"%"');
  sheet.getRange('E13:E27').setNumberFormat('yyyy-mm-dd');
}

/**
 * AI予測 vs 人間の直感セクション（A32:E50）
 */
function setupDashboardAIComparison(sheet) {
  // セクションヘッダー
  sheet.getRange('A32').setValue('【AI予測 vs 人間の直感】');
  sheet.getRange('A32:E32').merge();
  sheet.getRange('A32')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('left');

  // ヘッダー行
  const headers = ['候補者ID', 'AI予測', '人間の直感', '乖離', '状態'];
  sheet.getRange('A33:E33').setValues([headers]);
  sheet.getRange('A33:E33')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  // QUERY関数でデータを抽出
  const query = `=QUERY(Dashboard_Data!M:P,
    "SELECT M, N, O, P
     WHERE M IS NOT NULL AND O IS NOT NULL
     ORDER BY P DESC
     LIMIT 15",
    0)`;

  sheet.getRange('A34').setFormula(query);

  // E列: 状態判定
  sheet.getRange('E34').setFormula(
    '=IF(D34="", "", ' +
    'IF(D34<=10, "✅ 一致", ' +
    'IF(D34<=20, "⚠️ やや乖離", "❌ 大きく乖離")))'
  );

  // 書式設定
  sheet.getRange('A34:E48').setBorder(
    true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('B34:D48').setNumberFormat('0.00"%"');
}

/**
 * ダッシュボードの条件付き書式を設定
 */
function setupDashboardConditionalFormats() {
  Logger.log('📝 条件付き書式を設定中...');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  if (!sheet) return;

  const rules = [];

  // === 候補者ランキング: 承諾可能性のヒートマップ（C13:C27） ===

  // 高確率（80点以上）: 緑
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(80)
      .setBackground('#d9ead3')
      .setRanges([sheet.getRange('C13:C27')])
      .build()
  );

  // 標準（60-79点）: 黄
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(60, 79)
      .setBackground('#fff2cc')
      .setRanges([sheet.getRange('C13:C27')])
      .build()
  );

  // 要注意（60点未満）: 赤
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(60)
      .setBackground('#f4cccc')
      .setFontColor('#cc0000')
      .setBold(true)
      .setRanges([sheet.getRange('C13:C27')])
      .build()
  );

  // === AI予測 vs 人間の直感: 乖離の強調（D34:D48） ===

  // 大きく乖離（20点以上）: オレンジ
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(20)
      .setBackground('#fce5cd')
      .setFontColor('#cc0000')
      .setBold(true)
      .setRanges([sheet.getRange('D34:D48')])
      .build()
  );

  sheet.setConditionalFormatRules(rules);

  Logger.log('✅ 条件付き書式設定完了');
}
