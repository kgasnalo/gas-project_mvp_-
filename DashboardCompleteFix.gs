/**
 * Dashboard完全修正スクリプト
 *
 * 問題:
 * - DashboardのKPI指標が異常（#DIV/0!、間違った値）
 * - 候補者リストが壊れている（QUERY関数の列選択ミス）
 *
 * 原因:
 * - DashboardSetup.gsがR列を参照しているが、実際はH列が「最新_承諾可能性」
 * - QUERY関数の列選択が間違っている
 *
 * 修正内容:
 * 1. KPI指標: R列 → H列に変更
 * 2. 候補者ランキング: QUERY関数の列選択を修正
 * 3. リスク候補者アラート: QUERY関数の列選択を修正
 * 4. 推奨アクション: QUERY関数の列選択を修正
 *
 * 実行方法:
 * 1. Apps Scriptエディタで新規スクリプト「DashboardCompleteFix」を作成
 * 2. このコードをコピー&ペースト
 * 3. 関数 fixDashboardComplete を実行
 * 4. ログで結果を確認
 * 5. Dashboardシートで結果を確認
 *
 * 期待される結果:
 * - B6: 74.3点（平均承諾可能性）
 * - B7: 3名（高確率候補者数）
 * - B8: 3名（要注意候補者数）
 * - 候補者リスト: 正しいスコアとステータスが表示される
 */

/**
 * Dashboard完全修正のメイン関数
 */
function fixDashboardComplete() {
  // 既存のSpreadsheetCompleteFix.gsのSPREADSHEET_IDを使用
  // const SPREADSHEET_ID = '1CDsorhyXBj8aHcBoYFAT9S4FfpNQOvayeOAsOtkqsuM';
  Logger.log('========================================');
  Logger.log('📊 Dashboard完全修正を開始');
  Logger.log('========================================');
  Logger.log('');

  try {
    // Phase 1: Candidates_Masterの列構成を確認
    Logger.log('=== Phase 1: Candidates_Masterの列構成確認 ===');
    verifyCandidatesMasterColumns();
    Logger.log('');

    // Phase 2: DashboardのKPI指標を修正
    Logger.log('=== Phase 2: DashboardのKPI指標修正 ===');
    fixDashboardKPIs();
    Logger.log('✅ Phase 2 完了');
    Logger.log('');

    // Phase 3: 候補者ランキングのQUERY関数を修正
    Logger.log('=== Phase 3: 候補者ランキング修正 ===');
    fixDashboardRanking();
    Logger.log('✅ Phase 3 完了');
    Logger.log('');

    // Phase 4: リスク候補者アラートのQUERY関数を修正
    Logger.log('=== Phase 4: リスク候補者アラート修正 ===');
    fixDashboardRiskAlert();
    Logger.log('✅ Phase 4 完了');
    Logger.log('');

    // Phase 5: 推奨アクションのQUERY関数を修正
    Logger.log('=== Phase 5: 推奨アクション修正 ===');
    fixDashboardRecommendedActions();
    Logger.log('✅ Phase 5 完了');
    Logger.log('');

    // Phase 6: 最終確認
    Logger.log('=== Phase 6: 最終確認 ===');
    verifyDashboardFixes();
    Logger.log('');

    Logger.log('========================================');
    Logger.log('✅ Dashboard完全修正完了');
    Logger.log('========================================');
    Logger.log('');
    Logger.log('📋 確認事項:');
    Logger.log('1. Dashboardシートを開く');
    Logger.log('2. B6: 平均承諾可能性が正常な値（74.3点など）');
    Logger.log('3. B7: 高確率候補者数が正常な値（3名など）');
    Logger.log('4. B8: 要注意候補者数が正常な値（3名など）');
    Logger.log('5. 候補者リスト: 正しいスコアとステータスが表示');
    Logger.log('');

  } catch (error) {
    Logger.log('');
    Logger.log('========================================');
    Logger.log('❌ エラーが発生しました');
    Logger.log('========================================');
    Logger.log('エラー内容: ' + error.message);
    Logger.log('スタックトレース: ' + error.stack);
    throw error;
  }
}

/**
 * Candidates_Masterの列構成を確認
 */
function verifyCandidatesMasterColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('Candidates_Master');

  if (!masterSheet) {
    throw new Error('Candidates_Masterシートが見つかりません');
  }

  Logger.log('Candidates_Masterの列構成:');

  const headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];

  // 重要な列を確認
  const importantColumns = {
    'A': 'candidate_id',
    'B': '氏名',
    'C': 'メールアドレス',
    'D': '現在ステータス',
    'G': '最新_合格可能性',
    'H': '最新_承諾可能性',
    'I': '総合ランク'
  };

  for (const [col, expectedName] of Object.entries(importantColumns)) {
    const colIndex = col.charCodeAt(0) - 'A'.charCodeAt(0);
    const actualName = headers[colIndex];
    Logger.log(`  ${col}列: ${actualName} ${actualName === expectedName ? '✅' : '⚠️'}`);
  }

  // サンプルデータを確認
  Logger.log('');
  Logger.log('サンプルデータ（2行目）:');
  Logger.log(`  A2（candidate_id）: ${masterSheet.getRange('A2').getValue()}`);
  Logger.log(`  B2（氏名）: ${masterSheet.getRange('B2').getValue()}`);
  Logger.log(`  G2（合格可能性）: ${masterSheet.getRange('G2').getValue()}`);
  Logger.log(`  H2（承諾可能性）: ${masterSheet.getRange('H2').getValue()}`);
  Logger.log(`  I2（ランク）: ${masterSheet.getRange('I2').getValue()}`);
  Logger.log('');
}

/**
 * DashboardのKPI指標を修正
 */
function fixDashboardKPIs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');

  if (!dashboardSheet) {
    throw new Error('Dashboardシートが見つかりません');
  }

  Logger.log('KPI指標の修正:');

  // B5: 総候補者数（変更なし）
  const b5Formula = '=COUNTA(Candidates_Master!A:A)-1 & "名"';
  dashboardSheet.getRange('B5').setFormula(b5Formula);
  Logger.log('  B5（総候補者数）: ' + b5Formula);

  // B6: 平均承諾可能性（R列 → H列に修正、空セル除外）
  const b6Formula = '=IF(COUNTIF(Candidates_Master!H:H,">0")>0,ROUND(AVERAGEIF(Candidates_Master!H:H,">0"),1),"N/A") & "点"';
  dashboardSheet.getRange('B6').setFormula(b6Formula);
  Logger.log('  B6（平均承諾可能性）: ' + b6Formula);

  // B7: 高確率候補者数（R列 → H列に修正）
  const b7Formula = '=COUNTIF(Candidates_Master!H:H,">=80") & "名"';
  dashboardSheet.getRange('B7').setFormula(b7Formula);
  Logger.log('  B7（高確率候補者数）: ' + b7Formula);

  // B8: 要注意候補者数（R列 → H列に修正）
  const b8Formula = '=COUNTIF(Candidates_Master!H:H,"<60") & "名"';
  dashboardSheet.getRange('B8').setFormula(b8Formula);
  Logger.log('  B8（要注意候補者数）: ' + b8Formula);

  // B9: 本日の新規記録（変更なし）
  const b9Formula = '=COUNTIF(Engagement_Log!D:D,TODAY()) & "件"';
  dashboardSheet.getRange('B9').setFormula(b9Formula);
  Logger.log('  B9（本日の新規記録）: ' + b9Formula);

  // B10: 人間の直感入力率（変更なし）
  const b10Formula = '=TEXT(COUNTIF(Candidates_Master!Q:Q,">0")/(COUNTA(Candidates_Master!A:A)-1),"0%")';
  dashboardSheet.getRange('B10').setFormula(b10Formula);
  Logger.log('  B10（人間の直感入力率）: ' + b10Formula);

  Logger.log('');

  // 結果確認
  Utilities.sleep(1000);
  Logger.log('修正後の値:');
  Logger.log(`  B5: ${dashboardSheet.getRange('B5').getValue()}`);
  Logger.log(`  B6: ${dashboardSheet.getRange('B6').getValue()}`);
  Logger.log(`  B7: ${dashboardSheet.getRange('B7').getValue()}`);
  Logger.log(`  B8: ${dashboardSheet.getRange('B8').getValue()}`);
  Logger.log(`  B9: ${dashboardSheet.getRange('B9').getValue()}`);
  Logger.log(`  B10: ${dashboardSheet.getRange('B10').getValue()}`);
}

/**
 * 候補者ランキングのQUERY関数を修正
 */
function fixDashboardRanking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');

  if (!dashboardSheet) {
    throw new Error('Dashboardシートが見つかりません');
  }

  Logger.log('候補者ランキングのQUERY関数修正:');

  // 正しいQUERY関数
  // A: candidate_id
  // B: 氏名
  // H: 最新_承諾可能性
  // G: 最新_合格可能性
  // D: 現在ステータス
  // F: 最終更新日時
  const query = `=QUERY(Candidates_Master!A:Z,
    "SELECT A, B, H, G, D, F
     WHERE A IS NOT NULL AND H IS NOT NULL
       AND D<>'辞退' AND D<>'見送り' AND D<>'承諾'
     ORDER BY H DESC
     LIMIT 15",
    0)`;

  dashboardSheet.getRange('B13').setFormula(query);
  Logger.log('  B13にQUERY関数を設定しました');
  Logger.log('  列順序: candidate_id, 氏名, 承諾可能性, 合格可能性, ステータス, 更新日');

  // H列: 承諾ストーリー（変更なし）
  const h13Formula = '=IFERROR(VLOOKUP(B13, Acceptance_Story!A:D, 4, FALSE), "未作成")';
  dashboardSheet.getRange('H13').setFormula(h13Formula);
  Logger.log('  H13: 承諾ストーリー設定済み');

  Logger.log('');

  // 結果確認
  Utilities.sleep(1000);
  Logger.log('修正後のデータ（13行目）:');
  Logger.log(`  B13（ID）: ${dashboardSheet.getRange('B13').getValue()}`);
  Logger.log(`  C13（氏名）: ${dashboardSheet.getRange('C13').getValue()}`);
  Logger.log(`  D13（承諾可能性）: ${dashboardSheet.getRange('D13').getValue()}`);
  Logger.log(`  E13（合格可能性）: ${dashboardSheet.getRange('E13').getValue()}`);
  Logger.log(`  F13（ステータス）: ${dashboardSheet.getRange('F13').getValue()}`);
  Logger.log(`  G13（更新日）: ${dashboardSheet.getRange('G13').getValue()}`);
}

/**
 * リスク候補者アラートのQUERY関数を修正
 */
function fixDashboardRiskAlert() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');

  if (!dashboardSheet) {
    throw new Error('Dashboardシートが見つかりません');
  }

  Logger.log('リスク候補者アラートのQUERY関数修正:');

  // 正しいQUERY関数
  // A: candidate_id
  // B: 氏名
  // H: 最新_承諾可能性
  // D: 現在ステータス
  // F: 最終更新日時
  const query = `=QUERY(Candidates_Master!A:Z,
    "SELECT A, B, H, D, F
     WHERE A IS NOT NULL AND H IS NOT NULL AND H < 60
       AND D<>'辞退' AND D<>'見送り' AND D<>'承諾'
     ORDER BY H ASC
     LIMIT 8",
    0)`;

  dashboardSheet.getRange('A32').setFormula(query);
  Logger.log('  A32にQUERY関数を設定しました');
  Logger.log('  条件: 承諾可能性 < 60点');

  // G列: 推奨アクション（承諾可能性に基づいて自動判定）
  const g32Formula = `=IF(C32="", "",
    IF(C32<40, "🚨 即時対応: 採用マネージャーとの面談設定",
    IF(C32<50, "⚠️ 緊急フォローアップ（承諾可能性低下）",
    "📞 電話フォロー: 懸念事項の確認")))`;

  dashboardSheet.getRange('G32').setFormula(g32Formula);
  Logger.log('  G32: 推奨アクション設定済み');

  Logger.log('');
}

/**
 * 推奨アクションのQUERY関数を修正
 */
function fixDashboardRecommendedActions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');

  if (!dashboardSheet) {
    throw new Error('Dashboardシートが見つかりません');
  }

  Logger.log('推奨アクションのQUERY関数修正:');

  // 正しいQUERY関数
  // A: candidate_id
  // B: 氏名
  // H: 最新_承諾可能性
  // D: 現在ステータス
  const query = `=QUERY(Candidates_Master!A:Z,
    "SELECT A, B, H, D
     WHERE A IS NOT NULL AND H IS NOT NULL
       AND H >= 60 AND H < 80
       AND D<>'辞退' AND D<>'見送り' AND D<>'承諾'
     ORDER BY H DESC
     LIMIT 10",
    0)`;

  dashboardSheet.getRange('A44').setFormula(query);
  Logger.log('  A44にQUERY関数を設定しました');
  Logger.log('  条件: 60点 ≦ 承諾可能性 < 80点');

  // E列: 推奨アクション（ステータスと承諾可能性に基づいて自動判定）
  const e44Formula = `=IF(A44="", "",
    IF(D44="初回面談", "社員面談の設定",
    IF(D44="1次面接", "2次面接への推薦",
    IF(D44="社員面談", "2次面接への推薦",
    IF(D44="2次面接", "最終面接への推薦",
    IF(D44="最終面接", "内定手続きの開始",
    IF(D44="内定通知済", "承諾促進アクション", "次ステップへの推薦")))))))`;

  dashboardSheet.getRange('E44').setFormula(e44Formula);
  Logger.log('  E44: 推奨アクション設定済み');

  // F列: 期限（今日から7日後）
  const f44Formula = '=IF(A44="", "", TODAY()+7)';
  dashboardSheet.getRange('F44').setFormula(f44Formula);
  Logger.log('  F44: 期限設定済み');

  // G列: 優先度（承諾可能性に基づいて自動判定）
  const g44Formula = '=IF(C44="", "", IF(C44>=75, "低", IF(C44>=70, "中", "高")))';
  dashboardSheet.getRange('G44').setFormula(g44Formula);
  Logger.log('  G44: 優先度設定済み');

  // H列: 実行状況
  dashboardSheet.getRange('H44:H53').setValue('未実行');
  Logger.log('  H44:H53: 実行状況設定済み');

  Logger.log('');
}

/**
 * Dashboard修正の最終確認
 */
function verifyDashboardFixes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');

  if (!dashboardSheet) {
    throw new Error('Dashboardシートが見つかりません');
  }

  Logger.log('Dashboard修正の最終確認:');
  Logger.log('');

  // KPI指標の確認
  Logger.log('=== KPI指標 ===');
  const b5 = dashboardSheet.getRange('B5').getValue();
  const b6 = dashboardSheet.getRange('B6').getValue();
  const b7 = dashboardSheet.getRange('B7').getValue();
  const b8 = dashboardSheet.getRange('B8').getValue();
  const b9 = dashboardSheet.getRange('B9').getValue();
  const b10 = dashboardSheet.getRange('B10').getValue();

  Logger.log(`  B5（総候補者数）: ${b5}`);
  Logger.log(`  B6（平均承諾可能性）: ${b6} ${b6.toString().includes('点') ? '✅' : '❌'}`);
  Logger.log(`  B7（高確率候補者数）: ${b7} ${b7.toString().includes('名') ? '✅' : '❌'}`);
  Logger.log(`  B8（要注意候補者数）: ${b8} ${b8.toString().includes('名') ? '✅' : '❌'}`);
  Logger.log(`  B9（本日の新規記録）: ${b9}`);
  Logger.log(`  B10（人間の直感入力率）: ${b10}`);
  Logger.log('');

  // 候補者ランキングの確認
  Logger.log('=== 候補者ランキング（13行目） ===');
  const b13 = dashboardSheet.getRange('B13').getValue();
  const c13 = dashboardSheet.getRange('C13').getValue();
  const d13 = dashboardSheet.getRange('D13').getValue();
  const e13 = dashboardSheet.getRange('E13').getValue();
  const f13 = dashboardSheet.getRange('F13').getValue();
  const g13 = dashboardSheet.getRange('G13').getValue();

  Logger.log(`  B13（ID）: ${b13}`);
  Logger.log(`  C13（氏名）: ${c13}`);
  Logger.log(`  D13（承諾可能性）: ${d13} ${typeof d13 === 'number' && d13 >= 0 && d13 <= 100 ? '✅' : '❌'}`);
  Logger.log(`  E13（合格可能性）: ${e13}`);
  Logger.log(`  F13（ステータス）: ${f13}`);
  Logger.log(`  G13（更新日）: ${g13}`);
  Logger.log('');

  // 総合判定
  let allGood = true;
  if (!b6.toString().includes('点') || b6.toString().includes('#')) {
    Logger.log('❌ B6（平均承諾可能性）にエラーがあります');
    allGood = false;
  }
  if (!b7.toString().includes('名')) {
    Logger.log('❌ B7（高確率候補者数）にエラーがあります');
    allGood = false;
  }
  if (!b8.toString().includes('名')) {
    Logger.log('❌ B8（要注意候補者数）にエラーがあります');
    allGood = false;
  }
  if (typeof d13 !== 'number' || d13 < 0 || d13 > 100) {
    Logger.log('❌ D13（承諾可能性）が正常な値ではありません');
    allGood = false;
  }

  if (allGood) {
    Logger.log('✅ すべての修正が正常に完了しました！');
  } else {
    Logger.log('⚠️ 一部のデータに問題があります。手動で確認してください。');
  }
}

/**
 * デバッグ用: Dashboardの全データを表示
 */
function debugShowDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');

  if (!dashboardSheet) {
    throw new Error('Dashboardシートが見つかりません');
  }

  Logger.log('=== Dashboard 全データ ===');
  Logger.log('');

  // KPI指標
  Logger.log('【KPI指標】');
  for (let i = 5; i <= 10; i++) {
    const label = dashboardSheet.getRange(`A${i}`).getValue();
    const value = dashboardSheet.getRange(`B${i}`).getValue();
    const formula = dashboardSheet.getRange(`B${i}`).getFormula();
    Logger.log(`  ${label}: ${value}`);
    Logger.log(`    数式: ${formula}`);
  }
  Logger.log('');

  // 候補者ランキング
  Logger.log('【候補者ランキング】');
  for (let i = 13; i <= 17; i++) {
    const id = dashboardSheet.getRange(`B${i}`).getValue();
    const name = dashboardSheet.getRange(`C${i}`).getValue();
    const score = dashboardSheet.getRange(`D${i}`).getValue();
    const status = dashboardSheet.getRange(`F${i}`).getValue();
    Logger.log(`  ${i-12}位: ${id} | ${name} | ${score}点 | ${status}`);
  }
}
