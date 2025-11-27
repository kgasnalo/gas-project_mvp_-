/**
 * ========================================
 * スプレッドシート再設計 完全実装
 * 作成日: 2025年11月27日
 * 対象: 【MVP_v1】候補者管理シート
 * ========================================
 */

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
    '最終更新日時',                    // B
    '最新_合格可能性',                 // C
    '前回_合格可能性',                 // D
    '合格可能性_増減',                 // E
    '最新_Philosophy',                 // F
    '最新_Strategy',                   // G
    '最新_Motivation',                 // H
    '最新_Execution',                  // I
    '最新_合計スコア',                 // J
    '最新_承諾可能性（AI予測）',       // K
    '最新_承諾可能性（人間の直感）',   // L
    '最新_承諾可能性（統合）',         // M
    '前回_承諾可能性',                 // N
    '承諾可能性_増減',                 // O
    '予測の信頼度',                    // P
    '志望度スコア',                    // Q
    '競合優位性スコア',                // R
    '懸念解消度スコア',                // S
    'アンケート回答速度スコア'         // T
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
  sheet.setColumnWidth(2, 150);  // 最終更新日時
  for (let i = 3; i <= headers.length; i++) {
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
    '最終更新日時',          // B
    'コアモチベーション',    // C
    '主要懸念事項',          // D
    '懸念カテゴリ',          // E
    '競合企業1',             // F
    '競合企業2',             // G
    '競合企業3',             // H
    '次推奨アクション',      // I
    'アクション期限'         // J
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
  sheet.setColumnWidth(2, 150);  // 最終更新日時
  for (let i = 3; i <= headers.length; i++) {
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

  // 列のインデックスを取得する関数
  function getColumnIndex(headerName) {
    const index = headers.indexOf(headerName);
    if (index === -1) {
      Logger.log(`⚠️ 列が見つかりません: ${headerName}`);
    }
    return index;
  }

  // 必要な列のインデックスを取得
  const colIndexes = {
    candidate_id: getColumnIndex('candidate_id'),
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

  // 必要な列のインデックスを取得
  function getColumnIndex(headerName) {
    return headers.indexOf(headerName);
  }

  // 残す列のインデックス
  const keepColumns = [
    getColumnIndex('candidate_id'),           // A
    getColumnIndex('氏名'),                   // B
    getColumnIndex('現在ステータス'),         // C
    getColumnIndex('最終更新日時'),           // D
    getColumnIndex('採用区分'),               // E
    getColumnIndex('担当面接官'),             // F
    getColumnIndex('応募日'),                 // G
    getColumnIndex('メールアドレス'),         // H
    getColumnIndex('初回面談日'),             // I
    getColumnIndex('1次面接日'),              // J
    getColumnIndex('2次面接日'),              // K
    getColumnIndex('最終面接日'),             // L
    getColumnIndex('最新_合格可能性'),        // M（参照）
    getColumnIndex('最新_承諾可能性（統合）'), // N（参照）
    // O列: 総合ランク（新規追加予定）
  ];

  // 新しいヘッダー
  const newHeaders = [
    'candidate_id',
    '氏名',
    '現在ステータス',
    '最終更新日時',
    '採用区分',
    '担当面接官',
    '応募日',
    'メールアドレス',
    '初回面談日',
    '1次面接日',
    '2次面接日',
    '最終面接日',
    '最新_合格可能性',
    '最新_承諾可能性',
    '総合ランク'
  ];

  // 新しいデータ配列を作成
  const newData = [newHeaders];

  // 2行目以降（データ行）
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const candidateId = row[keepColumns[0]];

    if (!candidateId) continue;

    const newRow = keepColumns.map(colIndex => row[colIndex] || '');

    // 総合ランク（新規列）を追加
    const acceptanceRate = row[keepColumns[13]] || 0;
    let rank = 'E';
    if (acceptanceRate >= 80) rank = 'A';
    else if (acceptanceRate >= 70) rank = 'B';
    else if (acceptanceRate >= 60) rank = 'C';
    else if (acceptanceRate >= 50) rank = 'D';

    newRow.push(rank);
    newData.push(newRow);
  }

  // 既存のシートをクリア
  masterSheet.clear();

  // 新しいデータを書き込み
  masterSheet.getRange(1, 1, newData.length, newData[0].length)
    .setValues(newData);

  // ヘッダーの書式設定
  const headerRange = masterSheet.getRange(1, 1, 1, newHeaders.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // M列とN列を数式に変更（参照）
  for (let i = 2; i <= newData.length; i++) {
    const candidateId = masterSheet.getRange(i, 1).getValue();

    // M列: 最新_合格可能性（Candidate_Scoresから参照）
    masterSheet.getRange(i, 13).setFormula(
      `=IFERROR(VLOOKUP(A${i},Candidate_Scores!A:C,3,FALSE),"")`
    );

    // N列: 最新_承諾可能性（Candidate_Scoresから参照）
    masterSheet.getRange(i, 14).setFormula(
      `=IFERROR(VLOOKUP(A${i},Candidate_Scores!A:M,13,FALSE),"")`
    );
  }

  // 列幅を調整
  masterSheet.setColumnWidth(1, 120);  // candidate_id
  masterSheet.setColumnWidth(2, 150);  // 氏名
  for (let i = 3; i <= newHeaders.length; i++) {
    masterSheet.setColumnWidth(i, 130);
  }

  // 1行目を固定
  masterSheet.setFrozenRows(1);

  Logger.log(`✅ Candidates_Masterを再構成: ${newHeaders.length}列`);
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
        // Phase3スコアをEngagement_Logから取得
        sheet.getRange(i, startColumn + 1).setFormula(
          `=IFERROR(QUERY(Engagement_Log!B:H,"SELECT MAX(H) WHERE B='${candidateId}' LABEL MAX(H) ''"),"")`
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
    Logger.log('1. Candidates_Masterが15列になっているか');
    Logger.log('2. Candidate_Scoresにデータがあるか');
    Logger.log('3. Candidate_Insightsにデータがあるか');
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

  // M列（最新_合格可能性）の数式を確認
  const formulaM = masterSheet.getRange(2, 13).getFormula();
  Logger.log('M列の数式: ' + formulaM);

  // N列（最新_承諾可能性）の数式を確認
  const formulaN = masterSheet.getRange(2, 14).getFormula();
  Logger.log('N列の数式: ' + formulaN);

  if (formulaM.includes('VLOOKUP') && formulaN.includes('VLOOKUP')) {
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
