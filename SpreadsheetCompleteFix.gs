/**
 * スプレッドシート数式完全修正スクリプト
 *
 * 目的:
 * - Candidates_Master G列・H列のVLOOKUP数式を修正
 * - Candidates_Master I列のランク計算数式を修正
 * - Dashboard B8・B9の集計数式を修正
 *
 * 実行方法:
 * 1. Google Apps Scriptエディタで開く
 * 2. 関数 executeAllFixes を実行
 * 3. ログで結果を確認
 *
 * 個別実行も可能:
 * - fixCandidatesMasterGColumn() : G列のみ修正
 * - fixCandidatesMasterHColumn() : H列のみ修正
 * - fixCandidatesMasterIColumn() : I列のみ修正
 * - fixDashboardFormulas() : Dashboard修正
 */

const SPREADSHEET_ID = '1CDsorhyXBj8aHcBoYFAT9S4FfpNQOvayeOAsOtkqsuM';

/**
 * 全修正を順番に実行するメイン関数
 */
function executeAllFixes() {
  Logger.log('========================================');
  Logger.log('スプレッドシート数式修正を開始します');
  Logger.log('========================================');
  Logger.log('');

  try {
    // Phase 1: Candidates_Master G列の修正
    Logger.log('=== Phase 1: Candidates_Master G列の修正 ===');
    fixCandidatesMasterGColumn();
    Logger.log('✅ Phase 1 完了');
    Logger.log('');
    Utilities.sleep(1000); // 1秒待機

    // Phase 2: Candidates_Master H列の修正
    Logger.log('=== Phase 2: Candidates_Master H列の修正 ===');
    fixCandidatesMasterHColumn();
    Logger.log('✅ Phase 2 完了');
    Logger.log('');
    Utilities.sleep(1000); // 1秒待機

    // Phase 3: Candidates_Master I列の修正
    Logger.log('=== Phase 3: Candidates_Master I列の修正 ===');
    fixCandidatesMasterIColumn();
    Logger.log('✅ Phase 3 完了');
    Logger.log('');
    Utilities.sleep(1000); // 1秒待機

    // Phase 4: Dashboard の修正
    Logger.log('=== Phase 4: Dashboard の修正 ===');
    fixDashboardFormulas();
    Logger.log('✅ Phase 4 完了');
    Logger.log('');

    // 最終確認
    Logger.log('=== 最終確認 ===');
    verifyAllFixes();
    Logger.log('');

    Logger.log('========================================');
    Logger.log('✅ 全修正タスク完了');
    Logger.log('========================================');

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
 * タスク1: Candidates_Master G列（最新_合格可能性）の数式修正
 *
 * 問題: =IFERROR(VLOOKUP(A2,Candidate_Scores!A:C,3,FALSE),"")
 *       → 3列目は「最終更新日時」なので日付が返る
 *
 * 修正: =IFERROR(VLOOKUP(A2,Candidate_Scores!A:D,4,FALSE),"")
 *       → 4列目は「最新_合格可能性」なので正しいスコアが返る
 */
function fixCandidatesMasterGColumn() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName('Candidates_Master');

  if (!masterSheet) {
    throw new Error('Candidates_Masterシートが見つかりません');
  }

  const lastRow = masterSheet.getLastRow();
  Logger.log(`Candidates_Master: ${lastRow}行のデータを処理します`);

  // 現在の数式を確認
  const currentFormula = masterSheet.getRange('G2').getFormula();
  Logger.log(`現在のG2数式: ${currentFormula}`);

  // 正しい数式
  const correctFormula = '=IFERROR(VLOOKUP(A2,Candidate_Scores!A:D,4,FALSE),"")';
  Logger.log(`修正後の数式: ${correctFormula}`);

  // G2に正しい数式を設定
  masterSheet.getRange('G2').setFormula(correctFormula);
  Logger.log('G2の数式を修正しました');

  // 数式を下方向にコピー（G2からG列の最終行まで）
  if (lastRow >= 2) {
    const sourceRange = masterSheet.getRange('G2');
    const targetRange = masterSheet.getRange(`G2:G${lastRow}`);
    sourceRange.copyTo(targetRange);
    Logger.log(`G2の数式をG${lastRow}までコピーしました`);
  }

  // 結果確認
  Utilities.sleep(500);
  const resultValue = masterSheet.getRange('G2').getValue();
  Logger.log(`G2の結果値: ${resultValue} (型: ${typeof resultValue})`);

  if (typeof resultValue === 'number' && resultValue > 40000) {
    Logger.log('⚠️ 警告: まだ日付のシリアル値が表示されています');
  } else if (typeof resultValue === 'number' && resultValue >= 0 && resultValue <= 100) {
    Logger.log('✅ 正常: 正しいスコアが表示されています');
  }
}

/**
 * タスク2: Candidates_Master H列（最新_承諾可能性）の数式修正
 *
 * 問題: =IFERROR(VLOOKUP(A2,Candidate_Scores!A:M,13,FALSE),"")
 *       → 13列目は間違った列
 *
 * 修正: =IFERROR(VLOOKUP(A2,Candidate_Scores!A:N,14,FALSE),"")
 *       → 14列目は「最新_承諾可能性」なので正しいスコアが返る
 */
function fixCandidatesMasterHColumn() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName('Candidates_Master');

  if (!masterSheet) {
    throw new Error('Candidates_Masterシートが見つかりません');
  }

  const lastRow = masterSheet.getLastRow();
  Logger.log(`Candidates_Master: ${lastRow}行のデータを処理します`);

  // 現在の数式を確認
  const currentFormula = masterSheet.getRange('H2').getFormula();
  Logger.log(`現在のH2数式: ${currentFormula}`);

  // 正しい数式
  const correctFormula = '=IFERROR(VLOOKUP(A2,Candidate_Scores!A:N,14,FALSE),"")';
  Logger.log(`修正後の数式: ${correctFormula}`);

  // H2に正しい数式を設定
  masterSheet.getRange('H2').setFormula(correctFormula);
  Logger.log('H2の数式を修正しました');

  // 数式を下方向にコピー（H2からH列の最終行まで）
  if (lastRow >= 2) {
    const sourceRange = masterSheet.getRange('H2');
    const targetRange = masterSheet.getRange(`H2:H${lastRow}`);
    sourceRange.copyTo(targetRange);
    Logger.log(`H2の数式をH${lastRow}までコピーしました`);
  }

  // 結果確認
  Utilities.sleep(500);
  const resultValue = masterSheet.getRange('H2').getValue();
  Logger.log(`H2の結果値: ${resultValue} (型: ${typeof resultValue})`);

  if (typeof resultValue === 'number' && resultValue >= 0 && resultValue <= 100) {
    Logger.log('✅ 正常: 正しいスコアが表示されています');
  }
}

/**
 * タスク3: Candidates_Master I列（ランク）の数式修正
 *
 * ランクはG列とH列の平均値に基づいて決定されます。
 * G列とH列の修正後、ランク計算も正しく動作するはずです。
 *
 * 期待される数式の例:
 * =IF(OR(G2="",H2=""),"",
 *    IF(AVERAGE(G2,H2)>=80,"A",
 *    IF(AVERAGE(G2,H2)>=60,"B",
 *    IF(AVERAGE(G2,H2)>=40,"C","D"))))
 */
function fixCandidatesMasterIColumn() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName('Candidates_Master');

  if (!masterSheet) {
    throw new Error('Candidates_Masterシートが見つかりません');
  }

  const lastRow = masterSheet.getLastRow();
  Logger.log(`Candidates_Master: ${lastRow}行のデータを処理します`);

  // 現在の数式を確認
  const currentFormula = masterSheet.getRange('I2').getFormula();
  Logger.log(`現在のI2数式: ${currentFormula}`);

  // 正しい数式（G列とH列の平均値でランク決定）
  const correctFormula = '=IF(OR(G2="",H2=""),"",IF(AVERAGE(G2,H2)>=80,"A",IF(AVERAGE(G2,H2)>=60,"B",IF(AVERAGE(G2,H2)>=40,"C","D"))))';
  Logger.log(`推奨される数式: ${correctFormula}`);

  // 現在の数式が空または明らかに間違っている場合のみ修正
  if (!currentFormula || currentFormula === '') {
    Logger.log('数式が空なので、推奨数式を設定します');
    masterSheet.getRange('I2').setFormula(correctFormula);

    // 数式を下方向にコピー
    if (lastRow >= 2) {
      const sourceRange = masterSheet.getRange('I2');
      const targetRange = masterSheet.getRange(`I2:I${lastRow}`);
      sourceRange.copyTo(targetRange);
      Logger.log(`I2の数式をI${lastRow}までコピーしました`);
    }
  } else {
    Logger.log('既存の数式が存在します。確認してください。');
    Logger.log('必要に応じて手動で修正してください。');
  }

  // 結果確認
  Utilities.sleep(500);
  const resultValue = masterSheet.getRange('I2').getValue();
  Logger.log(`I2の結果値: ${resultValue}`);

  if (resultValue === 'A' || resultValue === 'B' || resultValue === 'C' || resultValue === 'D') {
    Logger.log('✅ 正常: 正しいランクが表示されています');
  } else if (resultValue === '') {
    Logger.log('⚠️ I2が空です（G2またはH2が空の可能性）');
  } else {
    Logger.log('⚠️ 予期しない値: ' + resultValue);
  }
}

/**
 * タスク4: Dashboard B8・B9の数式修正
 *
 * B8: 平均合格可能性（Candidates_Master G列の平均）
 * B9: 平均承諾可能性（Candidates_Master H列の平均）
 *
 * 期待される数式:
 * B8: =IFERROR(AVERAGE(Candidates_Master!G:G)*100&"%","データなし")
 * B9: =IFERROR(AVERAGE(Candidates_Master!H:H)*100&"%","データなし")
 *
 * または:
 * B8: =IFERROR(AVERAGE(Candidates_Master!G2:G11),"")
 * B9: =IFERROR(AVERAGE(Candidates_Master!H2:H11),"")
 */
function fixDashboardFormulas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dashboardSheet = ss.getSheetByName('Dashboard');

  if (!dashboardSheet) {
    Logger.log('⚠️ Dashboardシートが見つかりません。スキップします。');
    return;
  }

  // B8の修正
  Logger.log('Dashboard B8（平均合格可能性）の修正');
  const b8CurrentFormula = dashboardSheet.getRange('B8').getFormula();
  Logger.log(`現在のB8数式: ${b8CurrentFormula}`);

  // シンプルな平均値の数式
  const b8CorrectFormula = '=IFERROR(AVERAGE(Candidates_Master!G2:G100),"")';
  Logger.log(`修正後のB8数式: ${b8CorrectFormula}`);

  dashboardSheet.getRange('B8').setFormula(b8CorrectFormula);
  Logger.log('B8の数式を修正しました');

  // B9の修正
  Logger.log('Dashboard B9（平均承諾可能性）の修正');
  const b9CurrentFormula = dashboardSheet.getRange('B9').getFormula();
  Logger.log(`現在のB9数式: ${b9CurrentFormula}`);

  const b9CorrectFormula = '=IFERROR(AVERAGE(Candidates_Master!H2:H100),"")';
  Logger.log(`修正後のB9数式: ${b9CorrectFormula}`);

  dashboardSheet.getRange('B9').setFormula(b9CorrectFormula);
  Logger.log('B9の数式を修正しました');

  // 結果確認
  Utilities.sleep(500);
  const b8Value = dashboardSheet.getRange('B8').getValue();
  const b9Value = dashboardSheet.getRange('B9').getValue();
  Logger.log(`B8の結果値: ${b8Value}`);
  Logger.log(`B9の結果値: ${b9Value}`);

  if (typeof b8Value === 'number' && b8Value >= 0 && b8Value <= 100) {
    Logger.log('✅ B8: 正常な平均値が表示されています');
  }
  if (typeof b9Value === 'number' && b9Value >= 0 && b9Value <= 100) {
    Logger.log('✅ B9: 正常な平均値が表示されています');
  }
}

/**
 * 全修正の検証
 */
function verifyAllFixes() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName('Candidates_Master');
  const dashboardSheet = ss.getSheetByName('Dashboard');

  Logger.log('修正結果の最終確認:');
  Logger.log('');

  // Candidates_Master の確認
  if (masterSheet) {
    Logger.log('=== Candidates_Master ===');
    const g2 = masterSheet.getRange('G2').getValue();
    const h2 = masterSheet.getRange('H2').getValue();
    const i2 = masterSheet.getRange('I2').getValue();

    Logger.log(`G2（合格可能性）: ${g2}`);
    Logger.log(`H2（承諾可能性）: ${h2}`);
    Logger.log(`I2（ランク）: ${i2}`);

    let allGood = true;

    if (typeof g2 === 'number' && g2 > 40000) {
      Logger.log('❌ G2: まだ日付のシリアル値です');
      allGood = false;
    } else if (typeof g2 === 'number' && g2 >= 0 && g2 <= 100) {
      Logger.log('✅ G2: 正常');
    }

    if (typeof h2 === 'number' && h2 >= 0 && h2 <= 100) {
      Logger.log('✅ H2: 正常');
    } else if (h2 === '' || h2 === null) {
      Logger.log('⚠️ H2: 空です');
    }

    if (i2 === 'A' || i2 === 'B' || i2 === 'C' || i2 === 'D') {
      Logger.log('✅ I2: 正常');
    } else if (i2 === '') {
      Logger.log('⚠️ I2: 空です');
    }

    Logger.log('');
  }

  // Dashboard の確認
  if (dashboardSheet) {
    Logger.log('=== Dashboard ===');
    const b8 = dashboardSheet.getRange('B8').getValue();
    const b9 = dashboardSheet.getRange('B9').getValue();

    Logger.log(`B8（平均合格可能性）: ${b8}`);
    Logger.log(`B9（平均承諾可能性）: ${b9}`);

    if (typeof b8 === 'number') {
      Logger.log('✅ B8: 数値が表示されています');
    }
    if (typeof b9 === 'number') {
      Logger.log('✅ B9: 数値が表示されています');
    }

    Logger.log('');
  }
}

/**
 * デバッグ用: すべてのデータを詳細表示
 */
function debugShowAllData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName('Candidates_Master');
  const scoresSheet = ss.getSheetByName('Candidate_Scores');

  Logger.log('=== Candidates_Master 全データ ===');
  if (masterSheet) {
    const lastRow = masterSheet.getLastRow();
    for (let i = 1; i <= lastRow; i++) {
      const id = masterSheet.getRange(i, 1).getValue();
      const name = masterSheet.getRange(i, 2).getValue();
      const g = masterSheet.getRange(i, 7).getValue();
      const h = masterSheet.getRange(i, 8).getValue();
      const rank = masterSheet.getRange(i, 9).getValue();
      Logger.log(`${i}行: ${id} | ${name} | G=${g} | H=${h} | Rank=${rank}`);
    }
  }
  Logger.log('');

  Logger.log('=== Candidate_Scores 全データ ===');
  if (scoresSheet) {
    const lastRow = scoresSheet.getLastRow();
    for (let i = 1; i <= lastRow; i++) {
      const id = scoresSheet.getRange(i, 1).getValue();
      const name = scoresSheet.getRange(i, 2).getValue();
      const col4 = scoresSheet.getRange(i, 4).getValue();
      const col14 = scoresSheet.getRange(i, 14).getValue();
      Logger.log(`${i}行: ${id} | ${name} | 4列目=${col4} | 14列目=${col14}`);
    }
  }
}
