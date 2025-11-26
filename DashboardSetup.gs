/**
 * スプレッドシート内ダッシュボード実装（Phase 4-1）
 *
 * Candidates_Masterから直接データを取得し、Dashboardシートを作成します。
 * QUERY関数による自動集計、条件付き書式、チャートを設定します。
 *
 * 実装内容:
 * - Dashboardシート: メインダッシュボード（可視化とKPI）
 * - Dashboard_Dataシート: 削除（不要）
 *
 * @version 2.0（製品版対応）
 * @date 2025-11-25
 */

/**
 * ダッシュボード全体をセットアップ（メイン関数）
 *
 * この関数を実行することで、Dashboardシートが
 * 自動的に作成・設定されます。
 */
function setupDashboardPhase4() {
  try {
    Logger.log('====================================');
    Logger.log('📊 ダッシュボードセットアップ開始');
    Logger.log('====================================');

    // 1. Dashboard_Dataシートを削除（不要）
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
        '- Dashboard (メインダッシュボード)\n\n' +
        'すべてのデータはCandidates_Masterから自動取得されます。\n' +
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
 * ⚠️ このシートは削除されました（Phase 4-1改善）
 * すべてのデータはCandidates_Masterから直接取得します
 *
 * 理由: MVP版から製品版への移行時、スプレッドシートをコピーして
 * 販売する際の手直し工数を削減するため、中間シートを削除
 */
function setupDashboardDataSheet() {
  Logger.log('📝 Dashboard_Dataシートはスキップ（Candidates_Masterから直接取得）');

  // 既存のDashboard_Dataシートがあれば削除
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Dashboard_Data');
  if (sheet) {
    Logger.log('  🗑️ 既存のDashboard_Dataシートを削除中...');
    ss.deleteSheet(sheet);
  }

  Logger.log('✅ Dashboard_Data削除完了（不要）');
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

  // === PQ列の古いデータをクリア（QUERY関数の残骸を削除） ===
  Logger.log('📝 PQ列の古いデータをクリア中...');
  sheet.getRange('P29:Q40').clearContent();
  sheet.getRange('P62:Q66').clearContent();

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

  // KPI項目（Candidates_MasterのR列を使用）
  const kpiData = [
    ['総候補者数', '=COUNTA(Candidates_Master!A:A)-1 & "名"'],
    ['平均承諾可能性', '=ROUND(AVERAGE(Candidates_Master!R:R),1) & "点"'],
    ['高確率候補者数（80点以上）', '=COUNTIF(Candidates_Master!R:R,">=80") & "名"'],
    ['要注意候補者数（60点未満）', '=COUNTIF(Candidates_Master!R:R,"<60") & "名"'],
    ['本日の新規記録', '=COUNTIF(Engagement_Log!D:D,TODAY()) & "件"'],
    ['人間の直感入力率', '=TEXT(COUNTIF(Candidates_Master!Q:Q,">0")/(COUNTA(Candidates_Master!A:A)-1),"0%")']
  ];

  sheet.getRange('A5:B10').setValues(kpiData);

  // 列幅設定
  sheet.setColumnWidth(1, 200); // A: ラベル
  sheet.setColumnWidth(2, 150); // B: 値
  sheet.setColumnWidth(3, 150); // C
  sheet.setColumnWidth(4, 150); // D
  sheet.setColumnWidth(5, 150); // E
  sheet.setColumnWidth(6, 150); // F

  // 書式設定
  sheet.getRange('A5:A10').setFontWeight('bold');
  sheet.getRange('B5:B10')
    .setFontSize(16)
    .setHorizontalAlignment('right');

  // 交互の背景色
  for (let i = 5; i <= 10; i++) {
    if (i % 2 === 0) {
      sheet.getRange(`A${i}:B${i}`).setBackground('#f9f9f9');
    }
  }
}

/**
 * 候補者ランキングセクション（A11:H30）
 * Candidates_Masterから直接取得
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

  // ヘッダー行（承諾可能性（点）に変更）
  const headers = ['順位', '候補者ID', '氏名', '承諾可能性（点）', 'ステータス', '更新日', 'モチベーション', '承諾ストーリー'];
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

  // QUERY関数でCandidates_Masterから直接取得
  // A:候補者ID, B:氏名, R:承諾可能性（統合）, C:ステータス, D:最終更新日, Y:コアモチベーション
  const query = `=QUERY(Candidates_Master!A:Y,
    "SELECT A, B, R, C, D, Y
     WHERE A IS NOT NULL AND R IS NOT NULL
       AND C<>'辞退' AND C<>'見送り' AND C<>'承諾'
     ORDER BY R DESC
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
  sheet.setColumnWidth(5, 100);  // E: ステータス
  sheet.setColumnWidth(6, 110);  // F: 更新日
  sheet.setColumnWidth(7, 150);  // G: モチベーション
  sheet.setColumnWidth(8, 300);  // H: 承諾ストーリー

  // 書式設定（点数表記: 小数点1桁）
  sheet.getRange('A13:H27').setBorder(
    true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('D13:D27').setNumberFormat('0.0');
  sheet.getRange('F13:F27').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('H13:H27').setWrap(true);
}

/**
 * リスク候補者アラートセクション（A30:H40）
 * Candidates_Masterから直接取得
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

  // ヘッダー行（承諾可能性（点）に変更）
  const headers = ['候補者ID', '氏名', '承諾可能性（点）', 'ステータス', '更新日', '主要懸念事項', '推奨アクション'];
  sheet.getRange('A31:G31').setValues([headers]);
  sheet.getRange('A31:G31')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.CRITICAL)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  // QUERY関数でリスク候補者（承諾可能性60点未満）を抽出
  // A:候補者ID, B:氏名, R:承諾可能性, C:ステータス, D:更新日, Z:主要懸念事項
  const query = `=QUERY(Candidates_Master!A:Z,
    "SELECT A, B, R, C, D, Z
     WHERE A IS NOT NULL AND R IS NOT NULL AND R < 60
       AND C<>'辞退' AND C<>'見送り' AND C<>'承諾'
     ORDER BY R ASC
     LIMIT 8",
    0)`;

  sheet.getRange('A32').setFormula(query);

  // G列: 推奨アクション（承諾可能性に基づいて自動判定）
  sheet.getRange('G32').setFormula(
    '=IF(C32="", "", ' +
    'IF(C32<40, "🚨 即時対応: 採用マネージャーとの面談設定", ' +
    'IF(C32<50, "⚠️ 緊急フォローアップ（承諾可能性低下）", ' +
    '"📞 電話フォロー: 懸念事項の確認")))'
  );

  // 列幅設定
  sheet.setColumnWidth(1, 100);  // A: 候補者ID
  sheet.setColumnWidth(2, 120);  // B: 氏名
  sheet.setColumnWidth(3, 100);  // C: 承諾可能性
  sheet.setColumnWidth(4, 100);  // D: ステータス
  sheet.setColumnWidth(5, 110);  // E: 更新日
  sheet.setColumnWidth(6, 200);  // F: 主要懸念事項
  sheet.setColumnWidth(7, 250);  // G: 推奨アクション

  // 書式設定（点数表記: 小数点1桁）
  sheet.getRange('A32:G39').setBorder(
    true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('C32:C39').setNumberFormat('0.0');
  sheet.getRange('E32:E39').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('F32:G39').setWrap(true);
  sheet.getRange('A32:G39').setBackground('#fff3cd');
}

/**
 * 推奨アクションセクション（A42:H55）
 * Candidates_Masterから直接取得
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

  // ヘッダー行（承諾可能性（点）に変更）
  const headers = ['候補者ID', '氏名', '承諾可能性（点）', 'ステータス', '推奨アクション', '期限', '優先度', '実行状況'];
  sheet.getRange('A43:H43').setValues([headers]);
  sheet.getRange('A43:H43')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  // QUERY関数でアクションが必要な候補者を抽出（4列のみ）
  // A:候補者ID, B:氏名, R:承諾可能性, C:ステータス
  const query = `=QUERY(Candidates_Master!A:Y,
    "SELECT A, B, R, C
     WHERE A IS NOT NULL AND R IS NOT NULL
       AND R >= 60 AND R < 80
       AND C<>'辞退' AND C<>'見送り' AND C<>'承諾'
     ORDER BY R DESC
     LIMIT 10",
    0)`;

  sheet.getRange('A44').setFormula(query);

  // E列: 推奨アクション（ステータスと承諾可能性に基づいて自動判定）
  sheet.getRange('E44').setFormula(
    '=IF(A44="", "", ' +
    'IF(D44="初回面談", "社員面談の設定", ' +
    'IF(D44="1次面接", "2次面接への推薦", ' +
    'IF(D44="社員面談", "2次面接への推薦", ' +
    'IF(D44="2次面接", "最終面接への推薦", ' +
    'IF(D44="最終面接", "内定手続きの開始", ' +
    'IF(D44="内定通知済", "承諾促進アクション", "次ステップへの推薦")))))))'
  );

  // F列: 期限（今日から7日後）
  sheet.getRange('F44').setFormula('=IF(A44="", "", TODAY()+7)');

  // G列: 優先度（承諾可能性に基づいて自動判定）
  sheet.getRange('G44').setFormula(
    '=IF(C44="", "", IF(C44>=75, "低", IF(C44>=70, "中", "高")))'
  );

  // H列: 実行状況
  sheet.getRange('H44:H53').setValue('未実行');

  // 列幅設定
  sheet.setColumnWidth(1, 100);  // A: 候補者ID
  sheet.setColumnWidth(2, 120);  // B: 氏名
  sheet.setColumnWidth(3, 100);  // C: 承諾可能性
  sheet.setColumnWidth(4, 100);  // D: ステータス
  sheet.setColumnWidth(5, 250);  // E: 推奨アクション
  sheet.setColumnWidth(6, 110);  // F: 期限
  sheet.setColumnWidth(7, 80);   // G: 優先度
  sheet.setColumnWidth(8, 100);  // H: 実行状況

  // 書式設定（点数表記: 小数点1桁）
  sheet.getRange('A44:H53').setBorder(
    true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('C44:C53').setNumberFormat('0.0');
  sheet.getRange('F44:F53').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('E44:E53').setWrap(true);
}

/**
 * AI予測 vs 人間の直感セクション（A57:E75）
 * Candidates_Masterから直接取得
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

  // ヘッダー行（点数表記に変更）
  const headers = ['候補者ID', 'AI予測（点）', '人間の直感（点）', '乖離（点）', '状態'];
  sheet.getRange('A58:E58').setValues([headers]);
  sheet.getRange('A58:E58')
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center');

  // QUERY関数でCandidates_Masterから取得
  // A:候補者ID, P:AI予測, Q:人間の直感
  const query = `=QUERY(Candidates_Master!A:Q,
    "SELECT A, P, Q
     WHERE A IS NOT NULL AND P IS NOT NULL AND Q IS NOT NULL
       AND P > 0 AND Q > 0
     ORDER BY A
     LIMIT 15",
    0)`;

  sheet.getRange('A59').setFormula(query);

  // D列: 乖離計算
  sheet.getRange('D59').setFormula(
    '=IF(AND(B59<>"", C59<>""), ABS(B59-C59), "")'
  );

  // E列: 状態判定（閾値を5点と15点に変更）
  sheet.getRange('E59').setFormula(
    '=IF(D59="", "", ' +
    'IF(D59<=5, "✅ 一致", ' +
    'IF(D59<=15, "⚠️ やや乖離", "❌ 大きく乖離")))'
  );

  // 書式設定（点数表記: 小数点1桁）
  sheet.getRange('A59:E73').setBorder(
    true, true, true, true, true, true,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('B59:D73').setNumberFormat('0.0');
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

  // Step 1: 色スケール（ヒートマップ）を追加
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#34a853', SpreadsheetApp.InterpolationType.NUMBER, '100')
      .setGradientMidpointWithValue('#fbbc04', SpreadsheetApp.InterpolationType.NUMBER, '70')
      .setGradientMinpointWithValue('#ea4335', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setRanges([sheet.getRange('D13:D27')])
      .build()
  );

  // Step 2: 高確率（80点以上）: 緑背景 + 太字
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(80)
      .setBackground('#d9ead3')
      .setFontColor('#38761d')
      .setBold(true)
      .setRanges([sheet.getRange('D13:D27')])
      .build()
  );

  // Step 3: 要注意（60点未満）: 赤背景 + 太字
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(60)
      .setBackground('#f4cccc')
      .setFontColor('#cc0000')
      .setBold(true)
      .setRanges([sheet.getRange('D13:D27')])
      .build()
  );

  // === リスク候補者アラート: 承諾可能性の強調（C32:C39） ===

  // 極めて低い（40点未満）: 濃い赤
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(40)
      .setBackground('#cc0000')
      .setFontColor('#ffffff')
      .setBold(true)
      .setRanges([sheet.getRange('C32:C39')])
      .build()
  );

  // 低い（40-59点）: 薄い赤
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(40, 59)
      .setBackground('#f4cccc')
      .setFontColor('#cc0000')
      .setBold(true)
      .setRanges([sheet.getRange('C32:C39')])
      .build()
  );

  // === 推奨アクション: 優先度の色分け（G44:G53） ===

  // 優先度「高」: オレンジ
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('高')
      .setBackground('#fce5cd')
      .setFontColor('#cc0000')
      .setBold(true)
      .setRanges([sheet.getRange('G44:G53')])
      .build()
  );

  // 優先度「中」: 黄色
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('中')
      .setBackground('#fff2cc')
      .setRanges([sheet.getRange('G44:G53')])
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

  // やや乖離（5-15点）: 黄色
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(5, 15)
      .setBackground('#fff2cc')
      .setRanges([sheet.getRange('D59:D73')])
      .build()
  );

  sheet.setConditionalFormatRules(rules);

  Logger.log('✅ 条件付き書式設定完了');
}

/**
 * 列番号をアルファベットに変換（例: 1→A, 27→AA）
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * 現在ステータス列のアルファベットを取得
 */
function findStatusColumnLetter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('Candidates_Master');

  if (!master) {
    Logger.log('❌ Candidates_Masterシートが見つかりません');
    return null;
  }

  const headers = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i].toString().trim();
    if (header === '現在ステータス' ||
        header === 'ステータス' ||
        header === '選考ステータス' ||
        header === '現在のステータス') {
      return columnToLetter(i + 1);
    }
  }

  Logger.log('❌ 現在ステータス列が見つかりません');
  return null;
}

/**
 * 現在ステータス列を特定してログに出力（デバッグ用）
 */
function findStatusColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('Candidates_Master');

  if (!master) {
    Logger.log('❌ Candidates_Masterシートが見つかりません');
    return;
  }

  const headers = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i].toString().trim();
    if (header === '現在ステータス' ||
        header === 'ステータス' ||
        header === '選考ステータス' ||
        header === '現在のステータス') {
      const columnLetter = columnToLetter(i + 1);
      Logger.log(`✅ 現在ステータス列: ${columnLetter}列（${i + 1}列目）`);
      Logger.log(`📝 COUNTIF関数: =COUNTIF(Candidates_Master!${columnLetter}:${columnLetter},"初回面談")`);
      return columnLetter;
    }
  }

  Logger.log('❌ 現在ステータス列が見つかりません');
  Logger.log('📋 全ヘッダー一覧:');
  headers.forEach((h, i) => Logger.log(`  ${columnToLetter(i+1)}列: ${h}`));
}

/**
 * ステータス別候補者数のデータを確実に作成
 * Dashboard_DataシートのU1:V6にステータス別人数を作成
 */
function fixStatusChart() {
  Logger.log('====================================');
  Logger.log('🔧 ステータス別グラフデータ修正開始');
  Logger.log('====================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Dashboard_Dataシートを取得（なければ作成）
  let dataSheet = ss.getSheetByName('Dashboard_Data');
  if (!dataSheet) {
    Logger.log('📝 Dashboard_Dataシートを新規作成中...');
    dataSheet = ss.insertSheet('Dashboard_Data');

    // Engagement_Logの後ろに配置
    const engagementSheet = ss.getSheetByName('Engagement_Log');
    if (engagementSheet) {
      const engagementIndex = engagementSheet.getIndex();
      ss.setActiveSheet(dataSheet);
      ss.moveActiveSheet(engagementIndex + 1);
    }
  }

  // 現在ステータス列を特定
  const statusColumn = findStatusColumnLetter();

  if (!statusColumn) {
    throw new Error('❌ 現在ステータス列が見つかりません。Candidates_Masterシートのヘッダーを確認してください。');
  }

  Logger.log(`✅ 現在ステータス列を特定: ${statusColumn}列`);

  // ステータスの固定順序
  const statusOrder = ['初回面談', '1次面接', '社員面談', '2次面接', '最終面接'];

  // U1:V6の範囲をクリア
  dataSheet.getRange('U1:V6').clearContent();

  // ヘッダー行
  dataSheet.getRange('U1').setValue('ステータス');
  dataSheet.getRange('V1').setValue('人数');
  dataSheet.getRange('U1:V1')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');

  Logger.log('📝 ステータス別データを作成中...');

  // データ行
  statusOrder.forEach((status, index) => {
    const row = index + 2;
    dataSheet.getRange(`U${row}`).setValue(status);
    dataSheet.getRange(`V${row}`).setFormula(
      `=COUNTIF(Candidates_Master!${statusColumn}:${statusColumn},"${status}")`
    );
    Logger.log(`  ✓ ${status}: =COUNTIF(Candidates_Master!${statusColumn}:${statusColumn},"${status}")`);
  });

  // 列幅設定
  dataSheet.setColumnWidth(21, 120); // U列: ステータス
  dataSheet.setColumnWidth(22, 80);  // V列: 人数

  Logger.log('====================================');
  Logger.log('✅ ステータス別データを U1:V6 に作成完了');
  Logger.log('====================================');
  Logger.log('');
  Logger.log('📋 次の手順（手動作業）:');
  Logger.log('1. Dashboardシートを開く');
  Logger.log('2. ステータス別候補者数のグラフをクリック');
  Logger.log('3. 右上の「︙」→「グラフを編集」');
  Logger.log('4. データ範囲を「Dashboard_Data!U1:V6」に変更');
  Logger.log('5. 「更新」をクリック');
  Logger.log('');
}

/**
 * ステータス別グラフ用の固定データテーブルを作成
 * Dashboardシート内のR30:S35に配置
 */
function createStatusTableOnDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('Dashboard');

  if (!dashboard) {
    throw new Error('Dashboardシートが見つかりません');
  }

  // 現在ステータス列を特定
  const statusColumn = findStatusColumnLetter();

  if (!statusColumn) {
    throw new Error('現在ステータス列が見つかりません');
  }

  Logger.log('====================================');
  Logger.log('📊 ステータス別グラフ用データ作成開始');
  Logger.log('====================================');
  Logger.log(`✅ 現在ステータス列: ${statusColumn}列`);

  // ヘッダー
  dashboard.getRange('R30').setValue('ステータス');
  dashboard.getRange('S30').setValue('人数');

  // ヘッダーの書式設定
  const headerRange = dashboard.getRange('R30:S30');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // データ行（固定順序）
  const statusOrder = [
    '初回面談',
    '1次面接',
    '社員面談',
    '2次面接',
    '最終面接'
  ];

  statusOrder.forEach((status, index) => {
    const row = 31 + index;
    dashboard.getRange(`R${row}`).setValue(status);
    dashboard.getRange(`S${row}`).setFormula(
      `=COUNTIF(Candidates_Master!${statusColumn}:${statusColumn},"${status}")`
    );
    Logger.log(`  ✓ R${row}: ${status}`);
  });

  // データ範囲の書式設定
  const dataRange = dashboard.getRange('R31:S35');
  dataRange.setBorder(true, true, true, true, true, true);

  Logger.log('====================================');
  Logger.log('✅ ステータス別データ作成完了');
  Logger.log('====================================');
  Logger.log('');
  Logger.log('📋 次の手順（手動作業が必要）:');
  Logger.log('1. Dashboardシートを開く');
  Logger.log('2. ステータス別候補者数のグラフをクリック');
  Logger.log('3. 右上の「︙」→「グラフを編集」');
  Logger.log('4. データ範囲を「Dashboard!R30:S35」に変更');
  Logger.log('5. 「更新」をクリック');
  Logger.log('');
  Logger.log('⚠️ 重要: グラフのデータ範囲変更は手動で行う必要があります');
}
