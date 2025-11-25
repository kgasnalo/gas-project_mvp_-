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

    // 完了メッセージ（UIコンテキストがある場合のみ表示）
    try {
      SpreadsheetApp.getUi().alert(
        '✅ ダッシュボードセットアップ完了\n\n' +
        '以下のシートが作成されました:\n' +
        '- Dashboard_Data (中間データ)\n' +
        '- Dashboard (メインダッシュボード)\n\n' +
        'Dashboardシートを開いてください。'
      );
    } catch (uiError) {
      // スクリプトエディタから実行した場合はUIが利用できないため、ログのみ
      Logger.log('💡 Dashboardシートを開いて確認してください。');
    }

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
    const engagementIndex = engagementSheet.getIndex();
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(engagementIndex + 1);
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
 * セクション1: 最新承諾可能性データ（A1:F100）
 * 氏名列を追加
 */
function setupLatestAcceptanceData(sheet) {
  // ヘッダー行（氏名を追加）
  const headers = ['候補者ID', '氏名', '最新承諾可能性', '最新フェーズ', '最終更新日', 'コアモチベーション'];
  sheet.getRange('A1:F1').setValues([headers]);
  sheet.getRange('A1:F1')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT);

  // QUERY関数（候補者別最新承諾可能性）
  // Engagement_Logから最新データを取得し、承諾可能性順にソート
  // 氏名を追加（C列）
  const query = `=QUERY(Engagement_Log!A:U,
    "SELECT B, C, MAX(H), E, MAX(D), M
     WHERE B IS NOT NULL
     GROUP BY B, C, E, M
     ORDER BY MAX(H) DESC
     LABEL B '候補者ID', C '氏名', MAX(H) '最新承諾可能性', E '最新フェーズ', MAX(D) '最終更新日', M 'コアモチベーション'",
    1)`;

  sheet.getRange('A2').setFormula(query);

  // 列幅設定
  sheet.setColumnWidth(1, 120); // A: 候補者ID
  sheet.setColumnWidth(2, 120); // B: 氏名
  sheet.setColumnWidth(3, 150); // C: 最新承諾可能性
  sheet.setColumnWidth(4, 120); // D: 最新フェーズ
  sheet.setColumnWidth(5, 140); // E: 最終更新日
  sheet.setColumnWidth(6, 200); // F: コアモチベーション

  // フォーマット設定
  sheet.getRange('C2:C100').setNumberFormat('0.00"%"');
  sheet.getRange('E2:E100').setNumberFormat('yyyy-mm-dd');
}

/**
 * セクション2: フェーズ別人数分布データ（G1:H10）
 *
 * 各フェーズの候補者数を集計します。
 */
function setupPhaseScoreData(sheet) {
  // ヘッダー行
  const headers = ['フェーズ', '人数'];
  sheet.getRange('G1:H1').setValues([headers]);
  sheet.getRange('G1:H1')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT);

  // QUERY関数でフェーズごとの人数を集計
  const query = `=QUERY(Engagement_Log!E:E,
    "SELECT E, COUNT(E)
     WHERE E IS NOT NULL
     GROUP BY E
     ORDER BY COUNT(E) DESC
     LABEL E 'フェーズ', COUNT(E) '人数'",
    1)`;

  sheet.getRange('G2').setFormula(query);

  // 列幅設定
  sheet.setColumnWidth(7, 120);  // G: フェーズ
  sheet.setColumnWidth(8, 80);   // H: 人数

  // フォーマット設定
  sheet.getRange('H2:H10').setNumberFormat('0');
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

  // === ヘッダーセクション（A1:H2） ===
  setupDashboardHeader(sheet);

  // === KPIサマリーセクション（A4:H9） ===
  setupDashboardKPIs(sheet);

  // === 候補者ランキングセクション（A11:H28） ===
  setupDashboardRanking(sheet);

  // === リスク候補者アラート（A30:H40） ===
  setupRiskAlert(sheet);

  // === 推奨アクション（A42:H55） ===
  setupRecommendedActions(sheet);

  // === AI予測 vs 人間の直感セクション（A57:E75） ===
  setupDashboardAIComparison(sheet);

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
 * 候補者ランキングセクション（A11:H30）
 * 氏名と承諾ストーリーを追加
 */
function setupDashboardRanking(sheet) {
  // セクションヘッダー
  sheet.getRange('A11').setValue('【候補者別承諾可能性ランキング】');
  sheet.getRange('A11:H11').merge();
  sheet.getRange('A11')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('left');

  // ヘッダー行（氏名と承諾ストーリーを追加）
  const headers = ['順位', '候補者ID', '氏名', '承諾可能性', 'フェーズ', '更新日', 'モチベーション', '承諾ストーリー'];
  sheet.getRange('A12:H12').setValues([headers]);
  sheet.getRange('A12:H12')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  // 順位列（1-15）
  for (let i = 1; i <= 15; i++) {
    sheet.getRange(`A${12 + i}`).setValue(i);
  }

  // QUERY関数でデータを抽出（Dashboard_Dataから）
  // 列順: A:候補者ID, B:氏名, C:承諾可能性, D:フェーズ, E:更新日, F:モチベーション
  const query = `=QUERY(Dashboard_Data!A:F,
    "SELECT A, B, C, D, E, F
     WHERE A IS NOT NULL
     ORDER BY C DESC
     LIMIT 15",
    0)`;

  sheet.getRange('B13').setFormula(query);

  // H列: 承諾ストーリー（Acceptance_Storyシートから取得）
  sheet.getRange('H13').setFormula(
    '=IFERROR(VLOOKUP(B13, Acceptance_Story!A:D, 4, FALSE), "未作成")'
  );

  // 列幅設定
  sheet.setColumnWidth(1, 50);   // A: 順位
  sheet.setColumnWidth(2, 100);  // B: 候補者ID
  sheet.setColumnWidth(3, 120);  // C: 氏名
  sheet.setColumnWidth(4, 120);  // D: 承諾可能性
  sheet.setColumnWidth(5, 100);  // E: フェーズ
  sheet.setColumnWidth(6, 110);  // F: 更新日
  sheet.setColumnWidth(7, 150);  // G: モチベーション
  sheet.setColumnWidth(8, 300);  // H: 承諾ストーリー

  // 書式設定
  sheet.getRange('A13:H27').setBorder(
    true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('D13:D27').setNumberFormat('0.00"%"');
  sheet.getRange('F13:F27').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('H13:H27').setWrap(true);
}

/**
 * リスク候補者アラートセクション（A30:H40）
 */
function setupRiskAlert(sheet) {
  // セクションヘッダー
  sheet.getRange('A30').setValue('【⚠️ リスク候補者アラート】');
  sheet.getRange('A30:H30').merge();
  sheet.getRange('A30')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#fce5cd')
    .setHorizontalAlignment('left');

  // ヘッダー行
  const headers = ['候補者ID', '氏名', '承諾可能性', 'フェーズ', '更新日', 'リスク内容', '推奨アクション'];
  sheet.getRange('A31:G31').setValues([headers]);
  sheet.getRange('A31:G31')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.CRITICAL)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  // QUERY関数でリスク候補者（承諾可能性60点未満）を抽出
  const query = `=QUERY(Dashboard_Data!A:F,
    "SELECT A, B, C, D, E
     WHERE A IS NOT NULL AND C < 60
     ORDER BY C ASC
     LIMIT 8",
    0)`;

  sheet.getRange('A32').setFormula(query);

  // G列: リスク内容（リスクシートから取得）
  sheet.getRange('F32').setFormula(
    '=IFERROR(QUERY(Risk!B:F, "SELECT F WHERE B=\'"&A32&"\' ORDER BY H DESC LIMIT 1", 0), "データなし")'
  );

  // H列: 推奨アクション（Engagement_Logから取得）
  sheet.getRange('G32').setFormula(
    '=IFERROR(QUERY(Engagement_Log!B:R, "SELECT R WHERE B=\'"&A32&"\' ORDER BY D DESC LIMIT 1", 0), "要検討")'
  );

  // 列幅設定
  sheet.setColumnWidth(1, 100);  // A: 候補者ID
  sheet.setColumnWidth(2, 120);  // B: 氏名
  sheet.setColumnWidth(3, 100);  // C: 承諾可能性
  sheet.setColumnWidth(4, 100);  // D: フェーズ
  sheet.setColumnWidth(5, 110);  // E: 更新日
  sheet.setColumnWidth(6, 200);  // F: リスク内容
  sheet.setColumnWidth(7, 250);  // G: 推奨アクション

  // 書式設定
  sheet.getRange('A32:G39').setBorder(
    true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('C32:C39').setNumberFormat('0.00"%"');
  sheet.getRange('E32:E39').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('F32:G39').setWrap(true);
  sheet.getRange('A32:G39').setBackground('#fff3cd');
}

/**
 * 推奨アクションセクション（A42:H55）
 */
function setupRecommendedActions(sheet) {
  // セクションヘッダー
  sheet.getRange('A42').setValue('【💡 今週の推奨アクション】');
  sheet.getRange('A42:H42').merge();
  sheet.getRange('A42')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setHorizontalAlignment('left');

  // ヘッダー行
  const headers = ['候補者ID', '氏名', '承諾可能性', 'フェーズ', '推奨アクション', '期限', '優先度', '実行状況'];
  sheet.getRange('A43:H43').setValues([headers]);
  sheet.getRange('A43:H43')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  // QUERY関数でアクションが必要な候補者を抽出
  const query = `=QUERY(Dashboard_Data!A:F,
    "SELECT A, B, C, D
     WHERE A IS NOT NULL AND C >= 60 AND C < 80
     ORDER BY C DESC
     LIMIT 10",
    0)`;

  sheet.getRange('A44').setFormula(query);

  // E列: 推奨アクション
  sheet.getRange('E44').setFormula(
    '=IF(D44="初回面談", "社員面談の設定", IF(D44="社員面談", "2次面接への推薦", IF(D44="2次面接", "最終面接への推薦", IF(D44="内定後", "承諾促進アクション", "フォローアップ"))))'
  );

  // F列: 期限
  sheet.getRange('F44').setFormula('=TODAY()+7');

  // G列: 優先度
  sheet.getRange('G44').setFormula(
    '=IF(C44>=70, "中", IF(C44>=60, "高", "CRITICAL"))'
  );

  // H列: 実行状況
  sheet.getRange('H44').setValue('未実行');

  // 列幅設定
  sheet.setColumnWidth(1, 100);  // A: 候補者ID
  sheet.setColumnWidth(2, 120);  // B: 氏名
  sheet.setColumnWidth(3, 100);  // C: 承諾可能性
  sheet.setColumnWidth(4, 100);  // D: フェーズ
  sheet.setColumnWidth(5, 250);  // E: 推奨アクション
  sheet.setColumnWidth(6, 110);  // F: 期限
  sheet.setColumnWidth(7, 80);   // G: 優先度
  sheet.setColumnWidth(8, 100);  // H: 実行状況

  // 書式設定
  sheet.getRange('A44:H53').setBorder(
    true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('C44:C53').setNumberFormat('0.00"%"');
  sheet.getRange('F44:F53').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('E44:E53').setWrap(true);
}

/**
 * AI予測 vs 人間の直感セクション（A57:E75）
 */
function setupDashboardAIComparison(sheet) {
  // セクションヘッダー
  sheet.getRange('A57').setValue('【AI予測 vs 人間の直感】');
  sheet.getRange('A57:E57').merge();
  sheet.getRange('A57')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('left');

  // ヘッダー行
  const headers = ['候補者ID', 'AI予測', '人間の直感', '乖離', '状態'];
  sheet.getRange('A58:E58').setValues([headers]);
  sheet.getRange('A58:E58')
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

  sheet.getRange('A59').setFormula(query);

  // E列: 状態判定（閾値を5点と15点に変更）
  sheet.getRange('E59').setFormula(
    '=IF(D59="", "", ' +
    'IF(D59<=5, "✅ 一致", ' +
    'IF(D59<=15, "⚠️ やや乖離", "❌ 大きく乖離")))'
  );

  // 書式設定
  sheet.getRange('A59:E73').setBorder(
    true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('B59:D73').setNumberFormat('0.00"%"');
}

/**
 * ダッシュボードの条件付き書式を設定
 */
function setupDashboardConditionalFormats() {
  Logger.log('📝 条件付き書式を設定中...');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  if (!sheet) return;

  const rules = [];

  // === 候補者ランキング: 承諾可能性のヒートマップ（D13:D27） ===
  // 氏名追加により列がC→Dに変更

  // 高確率（80点以上）: 緑
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(80)
      .setBackground('#d9ead3')
      .setRanges([sheet.getRange('D13:D27')])
      .build()
  );

  // 標準（60-79点）: 黄
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(60, 79)
      .setBackground('#fff2cc')
      .setRanges([sheet.getRange('D13:D27')])
      .build()
  );

  // 要注意（60点未満）: 赤
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(60)
      .setBackground('#f4cccc')
      .setFontColor('#cc0000')
      .setBold(true)
      .setRanges([sheet.getRange('D13:D27')])
      .build()
  );

  // === AI予測 vs 人間の直感: 乖離の強調（D59:D73） ===

  // 大きく乖離（15点以上）: オレンジ
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(15)
      .setBackground('#fce5cd')
      .setFontColor('#cc0000')
      .setBold(true)
      .setRanges([sheet.getRange('D59:D73')])
      .build()
  );

  sheet.setConditionalFormatRules(rules);

  Logger.log('✅ 条件付き書式設定完了');
}
