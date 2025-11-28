/**
 * ========================================
 * スプレッドシート再設計 完全実装
 * 作成日: 2025年11月27日
 * 対象: 【MVP_v1】候補者管理シート
 * ========================================
 *
 * 【重要な安全対策】
 * - Step 2実行前に必ずバックアップを作成
 * - エラーハンドリングを強化
 * - データ検証を徹底
 */

// ========================================
// Phase 0: 事前準備（必須）
// ========================================

/**
 * バックアップ作成関数
 * Step 2実行前に必ず実行してください
 */
function createBackupBeforeStep2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('Candidates_Master');

  if (!masterSheet) {
    throw new Error('Candidates_Masterが見つかりません');
  }

  // バックアップシートを作成
  const backupSheet = masterSheet.copyTo(ss);
  const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
  backupSheet.setName('Candidates_Master_BACKUP_' + timestamp);
  backupSheet.hideSheet(); // 非表示にする

  Logger.log('✅ バックアップ作成完了: ' + backupSheet.getName());
  return backupSheet.getName();
}

/**
 * Phase 0: 事前準備（必須）
 * 実行前に必ずこの関数を実行してください
 */
function phase0_preparation() {
  Logger.log('========================================');
  Logger.log('Phase 0: 事前準備');
  Logger.log('========================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 現在のシート構成を記録
  const sheets = ss.getSheets();
  Logger.log(`現在のシート数: ${sheets.length}`);
  sheets.forEach(sheet => {
    Logger.log(`- ${sheet.getName()} (${sheet.getLastColumn()}列)`);
  });

  // 2. Candidates_Masterの列数を確認
  const masterSheet = ss.getSheetByName('Candidates_Master');
  if (!masterSheet) {
    throw new Error('Candidates_Masterが見つかりません');
  }
  Logger.log(`Candidates_Master列数: ${masterSheet.getLastColumn()}列`);

  // 3. 必須列の存在確認
  const headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];

  // 現在の列数で新旧構造を判定
  const currentColumns = masterSheet.getLastColumn();
  const isOldStructure = currentColumns > 30; // 57列の旧構造
  const isNewStructure = currentColumns === 21; // 21列の新構造

  Logger.log('');
  Logger.log('構造判定:');
  Logger.log(`  列数: ${currentColumns}列`);

  let requiredColumns;

  if (isNewStructure) {
    Logger.log('  状態: ✅ 既に新構造（21列）に移行済み');
    Logger.log('');
    Logger.log('⚠️ 既に新構造に移行済みです。');
    Logger.log('   再度実行する場合は、以下を確認してください:');
    Logger.log('   1. バックアップシート（Candidates_Master_BACKUP_*）が存在するか');
    Logger.log('   2. Candidate_ScoresとCandidate_Insightsシートが既に存在する場合、削除されます');
    Logger.log('');

    // 新構造用の必須列チェック
    requiredColumns = [
      'candidate_id',
      '氏名',
      'メールアドレス',
      '現在ステータス',
      '最終更新日時',
      '最新_合格可能性',
      '最新_承諾可能性'
    ];
  } else if (isOldStructure) {
    Logger.log('  状態: 📝 旧構造（57列）- 移行が必要');

    // 旧構造用の必須列チェック
    requiredColumns = [
      'candidate_id',
      '氏名',
      '現在ステータス',
      '最終更新日時',
      '最新_合格可能性',
      '最新_承諾可能性（統合）',
      'コアモチベーション',
      '主要懸念事項'
    ];
  } else {
    Logger.log('  状態: ⚠️ 不明な構造');
    throw new Error(`想定外の列数です: ${currentColumns}列\n` +
      '旧構造（57列）または新構造（21列）である必要があります。');
  }

  Logger.log('');
  Logger.log('必須列の存在確認:');
  let allColumnsExist = true;
  requiredColumns.forEach(col => {
    const exists = headers.includes(col);
    Logger.log(`  ${exists ? '✅' : '❌'} ${col}`);
    if (!exists) allColumnsExist = false;
  });

  if (!allColumnsExist) {
    throw new Error('必須列が不足しています。上記のログを確認してください。');
  }

  // 4. バックアップ作成
  Logger.log('');
  const backupName = createBackupBeforeStep2();

  Logger.log('');
  Logger.log('====================================');
  Logger.log('✅ Phase 0完了');
  Logger.log('====================================');
  Logger.log('');
  Logger.log('次のステップ: phase1_execute() を実行してください');
  Logger.log('');
}

// ========================================
// Step 1: 新規シート作成（3シート）
// ========================================

/**
 * 1-1. Candidate_Scores シート作成
 * 目的: Candidates_Masterからスコア関連の列を分離
 */
function createCandidateScoresSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存シートを削除（再実行時）
  const existingSheet = ss.getSheetByName('Candidate_Scores');
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }

  // 新規シート作成
  const sheet = ss.insertSheet('Candidate_Scores');

  // タブの色を設定（青系）
  sheet.setTabColor('#4285f4');

  // ヘッダー行を作成
  const headers = [
    'candidate_id',                    // A
    '氏名',                            // B
    '最終更新日時',                    // C
    '最新_合格可能性',                 // D
    '前回_合格可能性',                 // E
    '合格可能性_増減',                 // F
    '最新_Philosophy',                 // G
    '最新_Strategy',                   // H
    '最新_Motivation',                 // I
    '最新_Execution',                  // J
    '最新_合計スコア',                 // K
    '最新_承諾可能性（AI予測）',       // L
    '最新_承諾可能性（人間の直感）',   // M
    '最新_承諾可能性（統合）',         // N
    '前回_承諾可能性',                 // O
    '承諾可能性_増減',                 // P
    '予測の信頼度',                    // Q
    '志望度スコア',                    // R
    '競合優位性スコア',                // S
    '懸念解消度スコア',                // T
    'アンケート回答速度スコア'         // U
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダーの書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');

  // 列幅を調整
  sheet.setColumnWidth(1, 120);  // candidate_id
  sheet.setColumnWidth(2, 150);  // 氏名
  sheet.setColumnWidth(3, 150);  // 最終更新日時
  for (let i = 4; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 150);
  }

  // 1行目を固定
  sheet.setFrozenRows(1);

  Logger.log('✅ Candidate_Scoresシート作成完了');
}

/**
 * 1-2. Candidate_Insights シート作成
 * 目的: AIインサイト（モチベーション、懸念、競合）を管理
 */
function createCandidateInsightsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存シートを削除（再実行時）
  const existingSheet = ss.getSheetByName('Candidate_Insights');
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }

  // 新規シート作成
  const sheet = ss.insertSheet('Candidate_Insights');

  // タブの色を設定（紫系）
  sheet.setTabColor('#9c27b0');

  // ヘッダー行を作成
  const headers = [
    'candidate_id',           // A
    '氏名',                   // B
    '最終更新日時',          // C
    'コアモチベーション',    // D
    '主要懸念事項',          // E
    '懸念カテゴリ',          // F
    '競合企業1',             // G
    '競合企業2',             // H
    '競合企業3',             // I
    '次推奨アクション',      // J
    'アクション期限'         // K
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダーの書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');

  // 列幅を調整
  sheet.setColumnWidth(1, 120);  // candidate_id
  sheet.setColumnWidth(2, 150);  // 氏名
  sheet.setColumnWidth(3, 150);  // 最終更新日時
  for (let i = 4; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 200);
  }

  // 1行目を固定
  sheet.setFrozenRows(1);

  Logger.log('✅ Candidate_Insightsシート作成完了');
}

/**
 * 1-3. Dify_Workflow_Log シート作成
 * 目的: Difyワークフローの実行履歴を記録
 */
function createDifyWorkflowLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存シートを削除（再実行時）
  const existingSheet = ss.getSheetByName('Dify_Workflow_Log');
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }

  // 新規シート作成
  const sheet = ss.insertSheet('Dify_Workflow_Log');

  // タブの色を設定（グレー系）
  sheet.setTabColor('#757575');

  // ヘッダー行を作成
  const headers = [
    'workflow_log_id',        // A
    'workflow_name',          // B
    'candidate_id',           // C
    'execution_date',         // D
    'status',                 // E
    'duration_seconds',       // F
    'input_summary',          // G
    'output_summary',         // H
    'error_message'           // I
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダーの書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#757575');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');

  // 列幅を調整
  sheet.setColumnWidth(1, 180);  // workflow_log_id
  sheet.setColumnWidth(2, 150);  // workflow_name
  sheet.setColumnWidth(3, 120);  // candidate_id
  sheet.setColumnWidth(4, 150);  // execution_date
  sheet.setColumnWidth(5, 100);  // status
  sheet.setColumnWidth(6, 120);  // duration_seconds
  sheet.setColumnWidth(7, 250);  // input_summary
  sheet.setColumnWidth(8, 250);  // output_summary
  sheet.setColumnWidth(9, 250);  // error_message

  // 1行目を固定
  sheet.setFrozenRows(1);

  // シートを非表示にする（デバッグ用）
  sheet.hideSheet();

  Logger.log('✅ Dify_Workflow_Logシート作成完了（非表示）');
}

/**
 * 1-4. Step 1 完了確認
 */
function executeStep1() {
  Logger.log('====================================');
  Logger.log('Step 1: 新規シート作成開始');
  Logger.log('====================================');

  createCandidateScoresSheet();
  createCandidateInsightsSheet();
  createDifyWorkflowLogSheet();

  Logger.log('====================================');
  Logger.log('✅ Step 1完了');
  Logger.log('====================================');
}

// ========================================
// Step 2: データ移行
// ========================================

/**
 * 2-2. データ移行スクリプト
 */
function migrateDataFromCandidatesMaster() {
  Logger.log('====================================');
  Logger.log('Step 2: データ移行開始');
  Logger.log('====================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('Candidates_Master');
  const scoresSheet = ss.getSheetByName('Candidate_Scores');
  const insightsSheet = ss.getSheetByName('Candidate_Insights');

  if (!masterSheet || !scoresSheet || !insightsSheet) {
    throw new Error('必要なシートが見つかりません');
  }

  // Candidates_Masterの全データを取得
  const masterData = masterSheet.getDataRange().getValues();
  const headers = masterData[0];

  // 列のインデックスを取得する関数（エラーハンドリング強化版）
  function getColumnIndex(headerName) {
    const index = headers.indexOf(headerName);
    if (index === -1) {
      throw new Error(`列が見つかりません: ${headerName}\n` +
        `利用可能な列: ${headers.join(', ')}`);
    }
    return index;
  }

  // 必要な列のインデックスを取得
  const colIndexes = {
    candidate_id: getColumnIndex('candidate_id'),
    氏名: getColumnIndex('氏名'),
    最終更新日時: getColumnIndex('最終更新日時'),
    最新_合格可能性: getColumnIndex('最新_合格可能性'),
    前回_合格可能性: getColumnIndex('前回_合格可能性'),
    合格可能性_増減: getColumnIndex('合格可能性_増減'),
    最新_Philosophy: getColumnIndex('最新_Philosophy'),
    最新_Strategy: getColumnIndex('最新_Strategy'),
    最新_Motivation: getColumnIndex('最新_Motivation'),
    最新_Execution: getColumnIndex('最新_Execution'),
    最新_合計スコア: getColumnIndex('最新_合計スコア'),
    最新_承諾可能性_AI: getColumnIndex('最新_承諾可能性（AI予測）'),
    最新_承諾可能性_人間: getColumnIndex('最新_承諾可能性（人間の直感）'),
    最新_承諾可能性_統合: getColumnIndex('最新_承諾可能性（統合）'),
    前回_承諾可能性: getColumnIndex('前回_承諾可能性'),
    承諾可能性_増減: getColumnIndex('承諾可能性_増減'),
    予測の信頼度: getColumnIndex('予測の信頼度'),
    志望度スコア: getColumnIndex('志望度スコア'),
    競合優位性スコア: getColumnIndex('競合優位性スコア'),
    懸念解消度スコア: getColumnIndex('懸念解消度スコア'),
    アンケート回答速度: getColumnIndex('アンケート回答速度スコア'),
    コアモチベーション: getColumnIndex('コアモチベーション'),
    主要懸念事項: getColumnIndex('主要懸念事項'),
    競合企業1: getColumnIndex('競合企業1'),
    競合企業2: getColumnIndex('競合企業2'),
    競合企業3: getColumnIndex('競合企業3'),
    次推奨アクション: getColumnIndex('次推奨アクション'),
    アクション期限: getColumnIndex('アクション期限')
  };

  // データ移行用の配列
  const scoresData = [];
  const insightsData = [];

  // 2行目以降（データ行）をループ
  for (let i = 1; i < masterData.length; i++) {
    const row = masterData[i];
    const candidateId = row[colIndexes.candidate_id];

    if (!candidateId) continue; // candidate_idがない行はスキップ

    // Candidate_Scores用のデータ
    scoresData.push([
      candidateId,
      row[colIndexes.氏名] || '',
      row[colIndexes.最終更新日時] || '',
      row[colIndexes.最新_合格可能性] || '',
      row[colIndexes.前回_合格可能性] || '',
      row[colIndexes.合格可能性_増減] || '',
      row[colIndexes.最新_Philosophy] || '',
      row[colIndexes.最新_Strategy] || '',
      row[colIndexes.最新_Motivation] || '',
      row[colIndexes.最新_Execution] || '',
      row[colIndexes.最新_合計スコア] || '',
      row[colIndexes.最新_承諾可能性_AI] || '',
      row[colIndexes.最新_承諾可能性_人間] || '',
      row[colIndexes.最新_承諾可能性_統合] || '',
      row[colIndexes.前回_承諾可能性] || '',
      row[colIndexes.承諾可能性_増減] || '',
      row[colIndexes.予測の信頼度] || '',
      row[colIndexes.志望度スコア] || '',
      row[colIndexes.競合優位性スコア] || '',
      row[colIndexes.懸念解消度スコア] || '',
      row[colIndexes.アンケート回答速度] || ''
    ]);

    // Candidate_Insights用のデータ
    insightsData.push([
      candidateId,
      row[colIndexes.氏名] || '',
      row[colIndexes.最終更新日時] || '',
      row[colIndexes.コアモチベーション] || '',
      row[colIndexes.主要懸念事項] || '',
      '',  // 懸念カテゴリ（後で追加予定）
      row[colIndexes.競合企業1] || '',
      row[colIndexes.競合企業2] || '',
      row[colIndexes.競合企業3] || '',
      row[colIndexes.次推奨アクション] || '',
      row[colIndexes.アクション期限] || ''
    ]);
  }

  // Candidate_Scoresにデータを書き込み
  if (scoresData.length > 0) {
    scoresSheet.getRange(2, 1, scoresData.length, scoresData[0].length)
      .setValues(scoresData);
    Logger.log(`✅ Candidate_Scoresに${scoresData.length}件のデータを移行`);
  }

  // Candidate_Insightsにデータを書き込み
  if (insightsData.length > 0) {
    insightsSheet.getRange(2, 1, insightsData.length, insightsData[0].length)
      .setValues(insightsData);
    Logger.log(`✅ Candidate_Insightsに${insightsData.length}件のデータを移行`);
  }

  Logger.log('====================================');
  Logger.log('✅ Step 2完了');
  Logger.log('====================================');
}

/**
 * 2-3. Candidates_Master の列削除と再構成
 * ⚠️ 重要: データ移行が完了してから実行してください
 * ⚠️ 安全対策: エラー時はバックアップから復旧可能
 */
function reconstructCandidatesMaster() {
  Logger.log('====================================');
  Logger.log('Candidates_Master再構成開始');
  Logger.log('====================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('Candidates_Master');

  if (!masterSheet) {
    throw new Error('Candidates_Masterが見つかりません');
  }

  // 現在のデータを全て取得
  const allData = masterSheet.getDataRange().getValues();
  const headers = allData[0];

  // 必要な列のインデックスを取得（エラーハンドリング強化版）
  function getColumnIndex(headerName) {
    const index = headers.indexOf(headerName);
    if (index === -1) {
      throw new Error(`列が見つかりません: ${headerName}\n` +
        `利用可能な列: ${headers.join(', ')}`);
    }
    return index;
  }

  // 残す列のインデックス（既存データから取得）
  const keepColumns = [
    getColumnIndex('candidate_id'),           // 1
    getColumnIndex('氏名'),                   // 2
    getColumnIndex('メールアドレス'),         // 3
    getColumnIndex('現在ステータス'),         // 4
    getColumnIndex('採用区分'),               // 5
    getColumnIndex('最終更新日時'),           // 6
    getColumnIndex('最新_合格可能性'),        // 7（参照）
    getColumnIndex('最新_承諾可能性（統合）'), // 8（参照）
    // 9: 総合ランク（新規追加）
    getColumnIndex('応募日'),                 // 10
    getColumnIndex('初回面談日'),             // 11
    getColumnIndex('担当面接官'),             // 12（初回面談担当者として使用）
    // 13: 面談出席（新規追加）
    getColumnIndex('1次面接日'),              // 14
    // 15: 1次面接合否（新規追加）
    getColumnIndex('2次面接日'),              // 16
    // 17: 2次面接合否（新規追加）
    getColumnIndex('最終面接日'),             // 18
    // 19: 最終面接合否（新規追加）
    // 20: 内定日（新規追加）
    // 21: 承諾日時（新規追加）
  ];

  // 新しいヘッダー（21列）
  const newHeaders = [
    'candidate_id',
    '氏名',
    'メールアドレス',
    '現在ステータス',
    '採用区分',
    '最終更新日時',
    '最新_合格可能性',
    '最新_承諾可能性',
    '総合ランク',
    '応募日',
    '初回面談日',
    '初回面談担当者',
    '面談出席',
    '1次面接日',
    '1次面接合否',
    '2次面接日',
    '2次面接合否',
    '最終面接日',
    '最終面接合否',
    '内定日',
    '承諾日時'
  ];

  // 新しいデータ配列を作成
  const newData = [newHeaders];

  // 2行目以降（データ行）
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const candidateId = row[keepColumns[0]];

    if (!candidateId) continue;

    // 21列のデータを構築
    const newRow = [
      row[keepColumns[0]] || '',  // 1. candidate_id
      row[keepColumns[1]] || '',  // 2. 氏名
      row[keepColumns[2]] || '',  // 3. メールアドレス
      row[keepColumns[3]] || '',  // 4. 現在ステータス
      row[keepColumns[4]] || '',  // 5. 採用区分
      row[keepColumns[5]] || '',  // 6. 最終更新日時
      row[keepColumns[6]] || '',  // 7. 最新_合格可能性（後で数式に置き換え）
      row[keepColumns[7]] || '',  // 8. 最新_承諾可能性（後で数式に置き換え）
      '',                         // 9. 総合ランク（後で計算）
      row[keepColumns[8]] || '',  // 10. 応募日
      row[keepColumns[9]] || '',  // 11. 初回面談日
      row[keepColumns[10]] || '', // 12. 初回面談担当者
      '',                         // 13. 面談出席（新規追加）
      row[keepColumns[11]] || '', // 14. 1次面接日
      '',                         // 15. 1次面接合否（新規追加）
      row[keepColumns[12]] || '', // 16. 2次面接日
      '',                         // 17. 2次面接合否（新規追加）
      row[keepColumns[13]] || '', // 18. 最終面接日
      '',                         // 19. 最終面接合否（新規追加）
      '',                         // 20. 内定日（新規追加）
      ''                          // 21. 承諾日時（新規追加）
    ];

    // 総合ランク（9列目）を計算
    const acceptanceRate = row[keepColumns[7]] || 0;
    let rank = 'E';
    if (acceptanceRate >= 80) rank = 'A';
    else if (acceptanceRate >= 70) rank = 'B';
    else if (acceptanceRate >= 60) rank = 'C';
    else if (acceptanceRate >= 50) rank = 'D';
    newRow[8] = rank;

    newData.push(newRow);
  }

  // データ検証
  if (newData.length < 2) {
    throw new Error('新しいデータが空です。処理を中断します。');
  }

  Logger.log(`新しいデータ: ${newData.length}行（ヘッダー含む）`);

  try {
    // 既存のシートをクリア
    masterSheet.clear();

    // 新しいデータを書き込み
    masterSheet.getRange(1, 1, newData.length, newData[0].length)
      .setValues(newData);

    Logger.log('✅ データの書き込み完了');

  } catch (error) {
    Logger.log('❌ エラー発生: データの書き込みに失敗しました');
    Logger.log('⚠️ バックアップシート（Candidates_Master_BACKUP_*）から手動で復旧してください');
    throw error;
  }

  // ヘッダーの書式設定
  const headerRange = masterSheet.getRange(1, 1, 1, newHeaders.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // G列とH列を数式に変更（Candidate_Scoresから参照）
  for (let i = 2; i <= newData.length; i++) {
    // G列: 最新_合格可能性（Candidate_Scoresから参照）
    masterSheet.getRange(i, 7).setFormula(
      `=IFERROR(VLOOKUP(A${i},Candidate_Scores!A:D,4,FALSE),"")`
    );

    // H列: 最新_承諾可能性（Candidate_Scoresから参照）
    masterSheet.getRange(i, 8).setFormula(
      `=IFERROR(VLOOKUP(A${i},Candidate_Scores!A:N,14,FALSE),"")`
    );
  }

  // データ検証（プルダウン）を設定
  const dataRowCount = newData.length - 1; // ヘッダーを除く
  if (dataRowCount > 0) {
    // M列（13列目）: 面談出席（出席、欠席）
    const attendanceRange = masterSheet.getRange(2, 13, dataRowCount, 1);
    const attendanceRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['出席', '欠席'], true)
      .setAllowInvalid(false)
      .build();
    attendanceRange.setDataValidation(attendanceRule);

    // O列（15列目）: 1次面接合否（合格、不合格、欠席）
    const interview1Range = masterSheet.getRange(2, 15, dataRowCount, 1);
    const interviewRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['合格', '不合格', '欠席'], true)
      .setAllowInvalid(false)
      .build();
    interview1Range.setDataValidation(interviewRule);

    // Q列（17列目）: 2次面接合否（合格、不合格、欠席）
    const interview2Range = masterSheet.getRange(2, 17, dataRowCount, 1);
    interview2Range.setDataValidation(interviewRule);

    // S列（19列目）: 最終面接合否（合格、不合格、欠席）
    const finalInterviewRange = masterSheet.getRange(2, 19, dataRowCount, 1);
    finalInterviewRange.setDataValidation(interviewRule);

    // T列（20列目）、U列（21列目）: 日付形式
    const offerDateRange = masterSheet.getRange(2, 20, dataRowCount, 1);
    const acceptDateRange = masterSheet.getRange(2, 21, dataRowCount, 1);
    const dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build();
    offerDateRange.setDataValidation(dateRule);
    acceptDateRange.setDataValidation(dateRule);
  }

  // 列幅を調整
  masterSheet.setColumnWidth(1, 120);  // candidate_id
  masterSheet.setColumnWidth(2, 150);  // 氏名
  masterSheet.setColumnWidth(3, 200);  // メールアドレス
  masterSheet.setColumnWidth(4, 130);  // 現在ステータス
  masterSheet.setColumnWidth(5, 100);  // 採用区分
  masterSheet.setColumnWidth(6, 150);  // 最終更新日時
  masterSheet.setColumnWidth(7, 130);  // 最新_合格可能性
  masterSheet.setColumnWidth(8, 130);  // 最新_承諾可能性
  masterSheet.setColumnWidth(9, 100);  // 総合ランク
  for (let i = 10; i <= newHeaders.length; i++) {
    masterSheet.setColumnWidth(i, 130);
  }

  // 1行目を固定
  masterSheet.setFrozenRows(1);

  Logger.log(`✅ Candidates_Masterを再構成: ${newHeaders.length}列`);
  Logger.log(`✅ データ検証（プルダウン）を設定完了`);
  Logger.log('====================================');
  Logger.log('✅ Candidates_Master再構成完了');
  Logger.log('====================================');
}

/**
 * 2-4. Step 2 完了確認
 */
function executeStep2() {
  Logger.log('====================================');
  Logger.log('Step 2: データ移行開始');
  Logger.log('====================================');

  // データ移行
  migrateDataFromCandidatesMaster();

  // Candidates_Master再構成
  reconstructCandidatesMaster();

  Logger.log('====================================');
  Logger.log('✅ Step 2完了');
  Logger.log('====================================');
  Logger.log('');
  Logger.log('⚠️ 次に進む前に、以下を確認してください:');
  Logger.log('1. Candidate_Scoresにデータが正しく移行されているか');
  Logger.log('2. Candidate_Insightsにデータが正しく移行されているか');
  Logger.log('3. Candidates_Masterが15列になっているか');
  Logger.log('4. Candidates_MasterのM列・N列が数式になっているか');
}

// ========================================
// Step 3: 既存シート拡張（Dify連携用）
// ========================================

/**
 * 3-1. Evaluation_Log の拡張
 */
function expandEvaluationLog() {
  Logger.log('====================================');
  Logger.log('Evaluation_Log拡張開始');
  Logger.log('====================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Evaluation_Log');

  if (!sheet) {
    throw new Error('Evaluation_Logが見つかりません');
  }

  // 現在の列数を取得
  const lastColumn = sheet.getLastColumn();

  // 新しい列のヘッダー
  const newHeaders = [
    '文字起こし本文',
    'バッティング企業',
    'Googleドキュメント評価レポートURL',
    '積極性スコア',
    'dify_workflow_id'
  ];

  // T列から追加
  const startColumn = lastColumn + 1;
  sheet.getRange(1, startColumn, 1, newHeaders.length)
    .setValues([newHeaders]);

  // ヘッダーの書式を既存列と合わせる
  const headerRange = sheet.getRange(1, startColumn, 1, newHeaders.length);
  const existingHeaderFormat = sheet.getRange(1, 1).getBackground();
  headerRange.setBackground(existingHeaderFormat);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅を調整
  sheet.setColumnWidth(startColumn, 400);      // 文字起こし本文
  sheet.setColumnWidth(startColumn + 1, 200);  // バッティング企業
  sheet.setColumnWidth(startColumn + 2, 300);  // URL
  sheet.setColumnWidth(startColumn + 3, 120);  // 積極性スコア
  sheet.setColumnWidth(startColumn + 4, 180);  // workflow_id

  Logger.log(`✅ Evaluation_Logに${newHeaders.length}列を追加`);
  Logger.log('====================================');
}

/**
 * 3-2. Acceptance_Story の拡張
 */
function expandAcceptanceStory() {
  Logger.log('====================================');
  Logger.log('Acceptance_Story拡張開始');
  Logger.log('====================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Acceptance_Story');

  if (!sheet) {
    throw new Error('Acceptance_Storyが見つかりません');
  }

  // 現在の列数を取得
  const lastColumn = sheet.getLastColumn();

  // 新しい列のヘッダー
  const newHeaders = [
    'AI信頼度',
    'Phase3_acceptance_rate',
    'AI_acceptance_rate',
    '競合状況分析',
    'リスク要因',
    '機会要因',
    'dify_workflow_id'
  ];

  // R列から追加（既存がQ列まで）
  const startColumn = lastColumn + 1;
  sheet.getRange(1, startColumn, 1, newHeaders.length)
    .setValues([newHeaders]);

  // ヘッダーの書式を既存列と合わせる
  const headerRange = sheet.getRange(1, startColumn, 1, newHeaders.length);
  const existingHeaderFormat = sheet.getRange(1, 1).getBackground();
  headerRange.setBackground(existingHeaderFormat);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅を調整
  sheet.setColumnWidth(startColumn, 100);      // AI信頼度
  sheet.setColumnWidth(startColumn + 1, 150);  // Phase3
  sheet.setColumnWidth(startColumn + 2, 150);  // AI再試算
  sheet.setColumnWidth(startColumn + 3, 300);  // 競合状況
  sheet.setColumnWidth(startColumn + 4, 300);  // リスク要因
  sheet.setColumnWidth(startColumn + 5, 300);  // 機会要因
  sheet.setColumnWidth(startColumn + 6, 180);  // workflow_id

  // S列（Phase3_acceptance_rate）に数式を追加
  const dataRows = sheet.getLastRow();
  if (dataRows > 1) {
    for (let i = 2; i <= dataRows; i++) {
      const candidateId = sheet.getRange(i, 1).getValue();
      if (candidateId) {
        // シングルクォートのエスケープ処理（SQL構文エラー防止）
        const escapedId = candidateId.toString().replace(/'/g, "''");
        // Phase3スコアをEngagement_Logから取得
        sheet.getRange(i, startColumn + 1).setFormula(
          `=IFERROR(QUERY(Engagement_Log!B:H,"SELECT MAX(H) WHERE B='${escapedId}' LABEL MAX(H) ''"),"")`
        );
      }
    }
  }

  Logger.log(`✅ Acceptance_Storyに${newHeaders.length}列を追加`);
  Logger.log('====================================');
}

/**
 * 3-3. NextQ の拡張
 */
function expandNextQ() {
  Logger.log('====================================');
  Logger.log('NextQ拡張開始');
  Logger.log('====================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('NextQ');

  if (!sheet) {
    throw new Error('NextQが見つかりません');
  }

  // 現在の列数を取得
  const lastColumn = sheet.getLastColumn();

  // 新しい列のヘッダー
  const newHeaders = [
    '使用ステータス',
    '使用日時',
    'dify_workflow_id'
  ];

  // I列から追加
  const startColumn = lastColumn + 1;
  sheet.getRange(1, startColumn, 1, newHeaders.length)
    .setValues([newHeaders]);

  // ヘッダーの書式を既存列と合わせる
  const headerRange = sheet.getRange(1, startColumn, 1, newHeaders.length);
  const existingHeaderFormat = sheet.getRange(1, 1).getBackground();
  headerRange.setBackground(existingHeaderFormat);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅を調整
  sheet.setColumnWidth(startColumn, 120);      // 使用ステータス
  sheet.setColumnWidth(startColumn + 1, 150);  // 使用日時
  sheet.setColumnWidth(startColumn + 2, 180);  // workflow_id

  // 既存データに初期値を設定
  const dataRows = sheet.getLastRow();
  if (dataRows > 1) {
    const defaultValues = Array(dataRows - 1).fill(['未使用', '', '']);
    sheet.getRange(2, startColumn, dataRows - 1, newHeaders.length)
      .setValues(defaultValues);
  }

  Logger.log(`✅ NextQに${newHeaders.length}列を追加`);
  Logger.log('====================================');
}

/**
 * 3-4. Survey_Send_Log の拡張
 */
function expandSurveySendLog() {
  Logger.log('====================================');
  Logger.log('Survey_Send_Log拡張開始');
  Logger.log('====================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Survey_Send_Log');

  if (!sheet) {
    throw new Error('Survey_Send_Logが見つかりません');
  }

  // 現在の列数を取得
  const lastColumn = sheet.getLastColumn();

  // 新しい列のヘッダー
  const newHeaders = ['dify_workflow_id'];

  // I列に追加
  const startColumn = lastColumn + 1;
  sheet.getRange(1, startColumn, 1, newHeaders.length)
    .setValues([newHeaders]);

  // ヘッダーの書式を既存列と合わせる
  const headerRange = sheet.getRange(1, startColumn, 1, newHeaders.length);
  const existingHeaderFormat = sheet.getRange(1, 1).getBackground();
  headerRange.setBackground(existingHeaderFormat);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅を調整
  sheet.setColumnWidth(startColumn, 180);  // workflow_id

  Logger.log(`✅ Survey_Send_Logに${newHeaders.length}列を追加`);
  Logger.log('====================================');
}

/**
 * 3-5. Contact_History の拡張
 */
function expandContactHistory() {
  Logger.log('====================================');
  Logger.log('Contact_History拡張開始');
  Logger.log('====================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Contact_History');

  if (!sheet) {
    throw new Error('Contact_Historyが見つかりません');
  }

  // 現在の列数を取得
  const lastColumn = sheet.getLastColumn();

  // 新しい列のヘッダー
  const newHeaders = ['contact_source'];

  // I列に追加
  const startColumn = lastColumn + 1;
  sheet.getRange(1, startColumn, 1, newHeaders.length)
    .setValues([newHeaders]);

  // ヘッダーの書式を既存列と合わせる
  const headerRange = sheet.getRange(1, startColumn, 1, newHeaders.length);
  const existingHeaderFormat = sheet.getRange(1, 1).getBackground();
  headerRange.setBackground(existingHeaderFormat);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅を調整
  sheet.setColumnWidth(startColumn, 120);  // contact_source

  // 既存データに初期値を設定（手動）
  const dataRows = sheet.getLastRow();
  if (dataRows > 1) {
    const defaultValues = Array(dataRows - 1).fill(['手動']);
    sheet.getRange(2, startColumn, dataRows - 1, 1).setValues(defaultValues);
  }

  Logger.log(`✅ Contact_Historyに${newHeaders.length}列を追加`);
  Logger.log('====================================');
}

/**
 * 3-6. Step 3 完了確認
 */
function executeStep3() {
  Logger.log('====================================');
  Logger.log('Step 3: 既存シート拡張開始');
  Logger.log('====================================');

  expandEvaluationLog();
  expandAcceptanceStory();
  expandNextQ();
  expandSurveySendLog();
  expandContactHistory();

  Logger.log('====================================');
  Logger.log('✅ Step 3完了');
  Logger.log('====================================');
}

// ========================================
// Step 4: 不要シート削除
// ========================================

/**
 * 4-2. 削除スクリプト
 */
function deleteUnnecessarySheets() {
  Logger.log('====================================');
  Logger.log('Step 4: 不要シート削除開始');
  Logger.log('====================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetsToDelete = [
    'Survey_Analysis',
    'Evidence',
    'Risk',
    'Survey_Response'
  ];

  sheetsToDelete.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
      Logger.log(`✅ ${sheetName}を削除`);
    } else {
      Logger.log(`⚠️ ${sheetName}が見つかりません（既に削除済み？）`);
    }
  });

  Logger.log('====================================');
  Logger.log('✅ Step 4完了');
  Logger.log('====================================');
}

/**
 * 4-3. Step 4 完了確認
 */
function executeStep4() {
  Logger.log('====================================');
  Logger.log('Step 4: 不要シート削除開始');
  Logger.log('====================================');

  deleteUnnecessarySheets();

  Logger.log('====================================');
  Logger.log('✅ Step 4完了');
  Logger.log('====================================');
}

// ========================================
// 完全実行スクリプト
// ========================================

/**
 * 全てのステップを順番に実行
 */
function executeAllSteps() {
  Logger.log('########################################');
  Logger.log('# スプレッドシート再設計 完全実行開始 #');
  Logger.log('########################################');
  Logger.log('');

  try {
    // Step 1: 新規シート作成
    executeStep1();
    Logger.log('');

    // Step 2: データ移行
    executeStep2();
    Logger.log('');
    Logger.log('⚠️ 【重要】Step 2完了後、データを確認してください');
    Logger.log('問題なければ、executeStep3AndStep4() を実行してください');
    Logger.log('');

  } catch (error) {
    Logger.log('');
    Logger.log('❌ エラー発生:');
    Logger.log(error.toString());
    Logger.log('');
    Logger.log('スタックトレース:');
    Logger.log(error.stack);
  }
}

/**
 * Step 3 & Step 4 実行
 */
function executeStep3AndStep4() {
  Logger.log('########################################');
  Logger.log('# Step 3 & Step 4 実行開始 #');
  Logger.log('########################################');
  Logger.log('');

  try {
    // Step 3: 既存シート拡張
    executeStep3();
    Logger.log('');

    // Step 4: 不要シート削除
    executeStep4();
    Logger.log('');

    Logger.log('########################################');
    Logger.log('# ✅ 全ステップ完了！ #');
    Logger.log('########################################');
    Logger.log('');
    Logger.log('📊 最終確認事項:');
    Logger.log('1. Candidates_Masterが21列になっているか');
    Logger.log('2. Candidate_ScoresとCandidate_Insightsに氏名列があるか');
    Logger.log('3. データ検証（プルダウン）が設定されているか');
    Logger.log('4. Dify_Workflow_Logが非表示になっているか');
    Logger.log('5. 4つのシートが削除されているか');
    Logger.log('6. 既存シートに列が追加されているか');

  } catch (error) {
    Logger.log('');
    Logger.log('❌ エラー発生:');
    Logger.log(error.toString());
    Logger.log('');
    Logger.log('スタックトレース:');
    Logger.log(error.stack);
  }
}

// ========================================
// 検証スクリプト
// ========================================

/**
 * データの整合性確認
 */
function verifyDataIntegrity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('Candidates_Master');
  const scoresSheet = ss.getSheetByName('Candidate_Scores');
  const insightsSheet = ss.getSheetByName('Candidate_Insights');

  const masterIds = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 1)
    .getValues().flat().filter(id => id);
  const scoresIds = scoresSheet.getRange(2, 1, scoresSheet.getLastRow() - 1, 1)
    .getValues().flat().filter(id => id);
  const insightsIds = insightsSheet.getRange(2, 1, insightsSheet.getLastRow() - 1, 1)
    .getValues().flat().filter(id => id);

  Logger.log('====================================');
  Logger.log('データ整合性チェック');
  Logger.log('====================================');
  Logger.log(`Candidates_Master: ${masterIds.length}件`);
  Logger.log(`Candidate_Scores: ${scoresIds.length}件`);
  Logger.log(`Candidate_Insights: ${insightsIds.length}件`);

  if (masterIds.length === scoresIds.length &&
      masterIds.length === insightsIds.length) {
    Logger.log('✅ データ件数が一致しています');
  } else {
    Logger.log('⚠️ データ件数が一致していません');
  }

  Logger.log('====================================');
}

/**
 * 数式の確認
 */
function verifyFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('Candidates_Master');

  Logger.log('====================================');
  Logger.log('数式チェック');
  Logger.log('====================================');

  // G列（最新_合格可能性）の数式を確認
  const formulaG = masterSheet.getRange(2, 7).getFormula();
  Logger.log('G列の数式: ' + formulaG);

  // H列（最新_承諾可能性）の数式を確認
  const formulaH = masterSheet.getRange(2, 8).getFormula();
  Logger.log('H列の数式: ' + formulaH);

  if (formulaG.includes('VLOOKUP') && formulaH.includes('VLOOKUP')) {
    Logger.log('✅ 数式が正しく設定されています');
  } else {
    Logger.log('⚠️ 数式が設定されていません');
  }

  Logger.log('====================================');
}

/**
 * 最終確認
 */
function finalVerification() {
  Logger.log('########################################');
  Logger.log('# 最終確認開始 #');
  Logger.log('########################################');
  Logger.log('');

  verifyDataIntegrity();
  Logger.log('');
  verifyFormulas();
  Logger.log('');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  Logger.log('====================================');
  Logger.log('全シート一覧');
  Logger.log('====================================');
  allSheets.forEach((sheet, index) => {
    const name = sheet.getName();
    const isHidden = sheet.isSheetHidden();
    const status = isHidden ? '🔒 非表示' : '✅ 表示';
    Logger.log(`${index + 1}. ${name} ${status}`);
  });

  Logger.log('====================================');
  Logger.log('');
  Logger.log('✅ 最終確認完了');
  Logger.log('');
}

/**
 * ロールバック手順
 */
function rollbackToBackup() {
  // 事前にスプレッドシートのコピーを作成しておいてください
  Logger.log('⚠️ ロールバックは手動で行ってください');
  Logger.log('1. スプレッドシートのバックアップから復元');
  Logger.log('2. または、Google Driveの「バージョン履歴」から復元');
}

// ========================================
// Phase別実行スクリプト（推奨）
// ========================================

/**
 * Phase 1: Step 1-2実行
 * 新規シート作成とデータ移行を実行します
 */
function phase1_execute() {
  Logger.log('========================================');
  Logger.log('Phase 1: Step 1-2実行');
  Logger.log('========================================');
  Logger.log('');

  try {
    executeAllSteps();

    Logger.log('');
    Logger.log('========================================');
    Logger.log('⚠️ データ確認してください:');
    Logger.log('========================================');
    Logger.log('1. Candidate_Scoresシートにデータがあるか');
    Logger.log('2. Candidate_Insightsシートにデータがあるか');
    Logger.log('3. Candidates_Masterが15列になっているか');
    Logger.log('4. Candidates_MasterのM列・N列が数式になっているか');
    Logger.log('');
    Logger.log('問題なければ phase2_execute() を実行してください');
    Logger.log('========================================');

  } catch (error) {
    Logger.log('');
    Logger.log('❌ エラー発生:');
    Logger.log(error.toString());
    Logger.log('');
    Logger.log('⚠️ バックアップシートから復旧してください');
  }
}

/**
 * Phase 2: Step 3-4実行
 * 既存シート拡張と不要シート削除を実行します
 */
function phase2_execute() {
  Logger.log('========================================');
  Logger.log('Phase 2: Step 3-4実行');
  Logger.log('========================================');
  Logger.log('');

  try {
    executeStep3AndStep4();

    Logger.log('');
    Logger.log('========================================');
    Logger.log('✅ 全ての実装が完了しました！');
    Logger.log('========================================');
    Logger.log('');
    Logger.log('最終確認: finalVerification() を実行してください');
    Logger.log('========================================');

  } catch (error) {
    Logger.log('');
    Logger.log('❌ エラー発生:');
    Logger.log(error.toString());
    Logger.log('');
  }
}

/**
 * 完全実行ガイド
 * この関数はガイドのみを表示します（実際の処理は行いません）
 */
function executionGuide() {
  Logger.log('########################################');
  Logger.log('# スプレッドシート再設計 実行ガイド #');
  Logger.log('########################################');
  Logger.log('');
  Logger.log('⚠️ 重要: 必ず以下の順序で実行してください');
  Logger.log('');
  Logger.log('【ステップ0】事前準備');
  Logger.log('  phase0_preparation()');
  Logger.log('  → バックアップ作成と列の存在確認');
  Logger.log('');
  Logger.log('【ステップ1】データ移行');
  Logger.log('  phase1_execute()');
  Logger.log('  → 新規シート作成 + データ移行');
  Logger.log('  → ログでデータを確認してください');
  Logger.log('');
  Logger.log('【ステップ2】既存シート拡張');
  Logger.log('  phase2_execute()');
  Logger.log('  → 既存シート拡張 + 不要シート削除');
  Logger.log('');
  Logger.log('【ステップ3】最終確認');
  Logger.log('  finalVerification()');
  Logger.log('  → データ整合性チェック');
  Logger.log('');
  Logger.log('########################################');
  Logger.log('');
  Logger.log('🚨 注意事項:');
  Logger.log('1. 必ずコピーしたスプレッドシートで実行');
  Logger.log('2. Phase 1実行後、必ずデータを確認');
  Logger.log('3. エラーが発生したらバックアップから復旧');
  Logger.log('');
  Logger.log('準備ができたら phase0_preparation() を実行');
  Logger.log('########################################');
}
