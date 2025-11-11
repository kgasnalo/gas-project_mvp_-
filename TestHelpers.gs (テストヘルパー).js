/**
 * テスト用ヘルパー関数
 */

/**
 * Candidates_Masterシートの列数とヘッダーを確認
 */
function checkCandidatesMasterColumns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);

    if (!sheet) {
      Logger.log('❌ Candidates_Masterシートが見つかりません');
      return;
    }

    // ヘッダー行を取得
    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    Logger.log(`📊 Candidates_Masterシートの情報:`);
    Logger.log(`   総列数: ${lastColumn}`);
    Logger.log(`   最終列: ${getColumnLetter(lastColumn)}`);
    Logger.log('');

    // 新しく追加された列を確認
    Logger.log('🆕 新規追加列の確認:');
    const newColumns = [
      { col: 43, name: 'AQ: メールアドレス' },
      { col: 44, name: 'AR: 初回面談日' },
      { col: 45, name: 'AS: 初回面談実施ステータス' },
      { col: 46, name: 'AT: 適性検査日' },
      { col: 47, name: 'AU: 適性検査実施ステータス' },
      { col: 48, name: 'AV: 1次面接日' },
      { col: 49, name: 'AW: 1次面接結果' },
      { col: 50, name: 'AX: 社員面談実施回数' },
      { col: 51, name: 'AY: 社員面談日（最終）' },
      { col: 52, name: 'AZ: 社員面談実施ステータス' },
      { col: 53, name: 'BA: 2次面接日' },
      { col: 54, name: 'BB: 2次面接実施ステータス' },
      { col: 55, name: 'BC: 最終面接日' },
      { col: 56, name: 'BD: 最終面接実施ステータス' }
    ];

    newColumns.forEach(item => {
      const actualHeader = headers[item.col - 1] || '(空)';
      const status = actualHeader !== '' && actualHeader !== '(空)' ? '✅' : '❌';
      Logger.log(`   ${status} ${item.name}: ${actualHeader}`);
    });

    // データ行の確認
    Logger.log('');
    Logger.log('📝 データ行の確認:');
    const dataRows = sheet.getLastRow() - 1;
    Logger.log(`   データ行数: ${dataRows}行`);

    if (dataRows > 0) {
      const firstDataRow = sheet.getRange(2, 1, 1, lastColumn).getValues()[0];
      const hasNewColumnData = firstDataRow[42] !== '' && firstDataRow[42] !== undefined;

      if (hasNewColumnData) {
        Logger.log(`   ✅ 2行目のAQ列にデータあり: ${firstDataRow[42]}`);
      } else {
        Logger.log(`   ⚠️ 2行目のAQ列にデータなし`);
        Logger.log(`   → insertSampleCandidateData()を実行してサンプルデータを投入してください`);
      }
    }

  } catch (error) {
    logError('checkCandidatesMasterColumns', error);
  }
}

/**
 * 列番号を列文字に変換（例: 1 → A, 27 → AA）
 */
function getColumnLetter(columnNumber) {
  let letter = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}

/**
 * Candidates_Masterシートの既存データをクリア（ヘッダー以外）
 */
function clearCandidatesMasterData() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '確認',
      'Candidates_Masterシートの全データ（ヘッダー以外）を削除しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      Logger.log('⚠️ キャンセルされました');
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);

    if (!sheet) {
      Logger.log('❌ Candidates_Masterシートが見つかりません');
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      Logger.log(`✅ ${lastRow - 1}行のデータを削除しました`);
    } else {
      Logger.log('ℹ️ 削除するデータがありません');
    }

  } catch (error) {
    logError('clearCandidatesMasterData', error);
  }
}

/**
 * 新しい列を使用したテストデータを1件追加
 */
function addTestCandidateWithNewColumns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);

    if (!sheet) {
      Logger.log('❌ Candidates_Masterシートが見つかりません');
      return;
    }

    const now = new Date();
    const today = new Date();

    function addDays(date, days) {
      const result = new Date(date);
      result.setDate(result.getDate() + days);
      return result;
    }

    // テストデータ（56列）
    const testData = [
      'TEST001', 'テスト太郎', '初回面談', now, '新卒', 'テスト担当', addDays(today, -5),
      0, 0, 0, 0, 0, 0, 0, 0,
      0, 0, 0, 0, 0, '', 0, 0, 0, '', '',
      '', '', '', '',
      '', 0, 0, 0, '',
      '', '', '', '', '',
      0, 0, 'test@example.com',  // AQ列（メールアドレス）
      addDays(today, -5), '実施済',  // AR-AS列（初回面談）
      '', '未実施',  // AT-AU列（適性検査）
      '', '未実施',  // AV-AW列（1次面接）
      0, '', '未実施',  // AX-AZ列（社員面談）
      '', '未実施',  // BA-BB列（2次面接）
      '', '未実施'  // BC-BD列（最終面接）
    ];

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, testData.length).setValues([testData]);

    Logger.log('✅ テストデータを追加しました');
    Logger.log(`   候補者ID: TEST001`);
    Logger.log(`   メールアドレス: test@example.com`);
    Logger.log(`   初回面談実施ステータス: 実施済`);
    Logger.log('');
    Logger.log('💡 AS列（初回面談実施ステータス）を「実施済」に変更すると、');
    Logger.log('   アンケート送信の確認ダイアログが表示されます');

  } catch (error) {
    logError('addTestCandidateWithNewColumns', error);
  }
}
