/**
 * スプレッドシートの現状確認用スクリプト
 *
 * 目的:
 * - Candidates_MasterのG列・H列の数式が正しいか確認
 * - Candidate_Scoresのデータ構造を確認
 * - 修正が必要かどうか判定
 *
 * 実行方法:
 * 1. Google Apps Scriptエディタで開く
 * 2. 関数 checkSpreadsheetStatus を実行
 * 3. ログ（表示 → ログ）を確認
 */

function checkSpreadsheetStatus() {
  // 正しいスプレッドシートを開く
  const ss = SpreadsheetApp.openById('1CDsorhyXBj8aHcBoYFAT9S4FfpNQOvayeOAsOtkqsuM');

  Logger.log('========================================');
  Logger.log('スプレッドシート名: ' + ss.getName());
  Logger.log('スプレッドシートID: ' + ss.getId());
  Logger.log('========================================');
  Logger.log('');

  // 全シート一覧
  Logger.log('=== 全シート一覧 ===');
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const isHidden = sheet.isSheetHidden();
    Logger.log(`${sheet.getName()}: ${lastRow}行 x ${lastCol}列 ${isHidden ? '(非表示)' : ''}`);
  });
  Logger.log('');

  // Candidates_Masterの詳細確認
  Logger.log('=== Candidates_Master の詳細 ===');
  const masterSheet = ss.getSheetByName('Candidates_Master');
  if (masterSheet) {
    const lastRow = masterSheet.getLastRow();
    const lastCol = masterSheet.getLastColumn();
    Logger.log(`総行数: ${lastRow} (データ: ${lastRow - 1}件)`);
    Logger.log(`総列数: ${lastCol}`);
    Logger.log('');

    // ヘッダー行を確認
    Logger.log('ヘッダー行 (1-15列目):');
    for (let col = 1; col <= Math.min(lastCol, 15); col++) {
      const header = masterSheet.getRange(1, col).getValue();
      Logger.log(`  ${col}列目: ${header}`);
    }
    Logger.log('');

    // G列とH列の数式を確認（最重要）
    Logger.log('【重要】G列とH列の数式確認:');
    const g2Formula = masterSheet.getRange('G2').getFormula();
    const h2Formula = masterSheet.getRange('H2').getFormula();
    Logger.log(`G2の数式: ${g2Formula}`);
    Logger.log(`H2の数式: ${h2Formula}`);
    Logger.log('');

    // 期待される数式
    Logger.log('【期待される数式】');
    Logger.log('G2: =IFERROR(VLOOKUP(A2,Candidate_Scores!A:D,4,FALSE),"")');
    Logger.log('H2: =IFERROR(VLOOKUP(A2,Candidate_Scores!A:N,14,FALSE),"")');
    Logger.log('');

    // 数式が正しいかチェック
    const g2Expected = '=IFERROR(VLOOKUP(A2,Candidate_Scores!A:D,4,FALSE),"")';
    const h2Expected = '=IFERROR(VLOOKUP(A2,Candidate_Scores!A:N,14,FALSE),"")';

    if (g2Formula === g2Expected) {
      Logger.log('✅ G2の数式は正しい');
    } else {
      Logger.log('❌ G2の数式が間違っている');
      Logger.log('   現在: ' + g2Formula);
      Logger.log('   正解: ' + g2Expected);
    }

    if (h2Formula === h2Expected) {
      Logger.log('✅ H2の数式は正しい');
    } else {
      Logger.log('❌ H2の数式が間違っている');
      Logger.log('   現在: ' + h2Formula);
      Logger.log('   正解: ' + h2Expected);
    }
    Logger.log('');

    // 先頭3件のデータ確認
    Logger.log('先頭3件のデータ:');
    Logger.log('ID | 氏名 | G列(値) | H列(値) | I列(ランク)');
    Logger.log('-'.repeat(70));
    for (let i = 2; i <= Math.min(4, lastRow); i++) {
      const id = masterSheet.getRange(i, 1).getValue();
      const name = masterSheet.getRange(i, 2).getValue();
      const g = masterSheet.getRange(i, 7).getValue();
      const h = masterSheet.getRange(i, 8).getValue();
      const rank = masterSheet.getRange(i, 9).getValue();
      Logger.log(`${id} | ${name} | ${g} | ${h} | ${rank}`);

      // G列の値の型をチェック
      if (typeof g === 'number' && g > 40000) {
        Logger.log(`  ⚠️ G列が日付のシリアル値になっています（${g}）`);
      }
    }
  } else {
    Logger.log('⚠️ Candidates_Masterが見つかりません');
  }
  Logger.log('');

  // Candidate_Scoresの確認
  Logger.log('=== Candidate_Scores の詳細 ===');
  const scoresSheet = ss.getSheetByName('Candidate_Scores');
  if (scoresSheet) {
    const lastRow = scoresSheet.getLastRow();
    const lastCol = scoresSheet.getLastColumn();
    Logger.log(`総行数: ${lastRow} (データ: ${lastRow - 1}件)`);
    Logger.log(`総列数: ${lastCol}`);
    Logger.log('');

    // ヘッダー確認（1-14列目）
    Logger.log('ヘッダー行 (1-14列目):');
    for (let col = 1; col <= Math.min(14, lastCol); col++) {
      const header = scoresSheet.getRange(1, col).getValue();
      Logger.log(`  ${col}列目: ${header}`);
    }
    Logger.log('');

    // データ確認
    if (lastRow >= 2) {
      Logger.log('先頭3件のデータ (重要な列のみ):');
      Logger.log('ID | 氏名 | 4列目 | 14列目');
      Logger.log('-'.repeat(60));
      for (let i = 2; i <= Math.min(4, lastRow); i++) {
        const id = scoresSheet.getRange(i, 1).getValue();
        const name = scoresSheet.getRange(i, 2).getValue();
        const col4 = scoresSheet.getRange(i, 4).getValue();
        const col14 = scoresSheet.getRange(i, 14).getValue();
        Logger.log(`${i}行目: ${id} | ${name} | ${col4} | ${col14}`);
      }
    }
  } else {
    Logger.log('⚠️ Candidate_Scoresが見つかりません');
  }
  Logger.log('');

  // Dashboard の確認
  Logger.log('=== Dashboard の確認 ===');
  const dashboardSheet = ss.getSheetByName('Dashboard');
  if (dashboardSheet) {
    Logger.log('Dashboard B8 (平均合格可能性):');
    const b8Value = dashboardSheet.getRange('B8').getValue();
    const b8Formula = dashboardSheet.getRange('B8').getFormula();
    Logger.log(`  値: ${b8Value}`);
    Logger.log(`  数式: ${b8Formula}`);
    Logger.log('');

    Logger.log('Dashboard B9 (平均承諾可能性):');
    const b9Value = dashboardSheet.getRange('B9').getValue();
    const b9Formula = dashboardSheet.getRange('B9').getFormula();
    Logger.log(`  値: ${b9Value}`);
    Logger.log(`  数式: ${b9Formula}`);
  } else {
    Logger.log('⚠️ Dashboardが見つかりません');
  }
  Logger.log('');

  // 結論
  Logger.log('========================================');
  Logger.log('【診断結果】');
  Logger.log('');

  if (masterSheet) {
    const g2Formula = masterSheet.getRange('G2').getFormula();
    const g2Value = masterSheet.getRange('G2').getValue();

    if (typeof g2Value === 'number' && g2Value > 40000) {
      Logger.log('❌ 問題あり: G列に日付のシリアル値が表示されている');
      Logger.log('→ VLOOKUP数式の列番号が間違っている可能性が高い');
      Logger.log('→ 修正が必要です');
      Logger.log('→ SpreadsheetCompleteFix.gs の executeAllFixes() を実行してください');
    } else if (typeof g2Value === 'number' && g2Value >= 0 && g2Value <= 100) {
      Logger.log('✅ 正常: G列に正しいスコアが表示されている');
      Logger.log('→ 修正は不要です');
    } else if (g2Value === '' || g2Value === null) {
      Logger.log('⚠️ G列が空です');
      Logger.log('→ Candidate_Scoresにデータがない可能性があります');
    } else {
      Logger.log('⚠️ 不明: G列の値を確認してください');
      Logger.log(`   値: ${g2Value} (型: ${typeof g2Value})`);
    }
  }

  Logger.log('');
  Logger.log('確認完了');
  Logger.log('========================================');
}
