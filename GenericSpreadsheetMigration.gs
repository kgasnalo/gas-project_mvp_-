// ========================================
// 汎用版スプレッドシート再設計 v2.0
// ========================================
// 作成日: 2025-11-28
// 目的: 販売可能・デモ可能・汎用性の高いMVP構築
// ========================================

// ========================================
// Phase 1: 汎用化基盤
// ========================================

/**
 * シートの構造を分析し、メタデータを返す
 * @param {string} sheetName - シート名
 * @return {Object} メタデータ
 */
function analyzeSheetStructure(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`シートが見つかりません: ${sheetName}`);
  }

  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    throw new Error(`シートが空です: ${sheetName}`);
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  const metadata = {
    sheetName: sheetName,
    totalColumns: headers.length,
    totalRows: sheet.getLastRow(),
    headers: headers,
    columnMap: {},
    dataTypes: {}
  };

  // 列名とインデックスのマッピング
  headers.forEach((header, index) => {
    if (header) {
      metadata.columnMap[header] = index;
    }
  });

  return metadata;
}

/**
 * 複数の候補から列を検索（別名対応）
 * @param {Object} sheetMetadata - analyzeSheetStructureの戻り値
 * @param {Array<string>} candidates - 候補となる列名
 * @return {number|null} 列インデックス（見つからない場合null）
 */
function findColumnByNames(sheetMetadata, candidates) {
  for (const candidate of candidates) {
    if (sheetMetadata.columnMap.hasOwnProperty(candidate)) {
      return sheetMetadata.columnMap[candidate];
    }
  }
  return null;
}

/**
 * 必須列の存在確認
 * @param {Object} sheetMetadata
 * @param {Object} requiredColumns - {論理名: [候補列名]}
 * @return {Object} {成功: bool, 不足列: [], 検出列: {}}
 */
function validateRequiredColumns(sheetMetadata, requiredColumns) {
  const missingColumns = [];
  const foundColumns = {};

  for (const [logicalName, candidates] of Object.entries(requiredColumns)) {
    const columnIndex = findColumnByNames(sheetMetadata, candidates);

    if (columnIndex === null) {
      missingColumns.push({
        logicalName: logicalName,
        candidates: candidates
      });
    } else {
      foundColumns[logicalName] = columnIndex;
    }
  }

  return {
    success: missingColumns.length === 0,
    missingColumns: missingColumns,
    foundColumns: foundColumns
  };
}

// ========================================
// Phase 2: データ移行
// ========================================

/**
 * バックアップシートを検索
 */
function findBackupSheet(spreadsheet) {
  const sheets = spreadsheet.getSheets();

  // Candidates_Master_BACKUP_* を探す
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (name.startsWith('Candidates_Master_BACKUP_')) {
      return sheet;
    }
  }

  return null;
}

/**
 * バックアップシートからデータを抽出
 * @return {Object} 抽出結果
 */
function extractDataFromBackup() {
  Logger.log('====================================');
  Logger.log('バックアップからデータ抽出開始');
  Logger.log('====================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // バックアップシートを検索
  const backupSheet = findBackupSheet(ss);
  if (!backupSheet) {
    throw new Error('バックアップシートが見つかりません');
  }

  Logger.log(`バックアップシート: ${backupSheet.getName()}`);

  // 構造を分析
  const backupMetadata = analyzeSheetStructure(backupSheet.getName());
  Logger.log(`列数: ${backupMetadata.totalColumns}`);
  Logger.log(`行数: ${backupMetadata.totalRows}`);
  Logger.log('');

  // 必須列の定義（複数の候補名に対応）
  const requiredColumns = {
    candidate_id: ['candidate_id', 'ID', '候補者ID'],
    name: ['氏名', '名前', 'name', 'Name'],
    email: ['メールアドレス', 'email', 'Email', 'Eメール'],
    status: ['現在ステータス', 'ステータス', 'status', 'Status'],
    updated_at: ['最終更新日時', '更新日時', 'updated_at'],

    // スコア系
    latest_pass_rate: ['最新_合格可能性', '合格可能性', 'pass_rate'],
    prev_pass_rate: ['前回_合格可能性'],
    prev2_pass_rate: ['前々回_合格可能性'],
    pass_diff1: ['スコア差分1'],
    pass_diff2: ['スコア差分2'],
    pass_diff3: ['スコア差分3'],
    pass_trend: ['スコア変動傾向'],

    // 承諾可能性
    latest_acceptance_integrated: ['最新_承諾可能性（統合）', '最新_承諾可能性(統合)'],
    prev_acceptance: ['前回_承諾可能性（統合）', '前回_承諾可能性(統合)'],
    prev2_acceptance: ['前々回_承諾可能性（統合）', '前々回_承諾可能性(統合)'],
    acceptance_diff1: ['スコア差分1（統合）', 'スコア差分1(統合)'],
    acceptance_diff2: ['スコア差分2（統合）', 'スコア差分2(統合)'],
    acceptance_diff3: ['スコア差分3（統合）', 'スコア差分3(統合)'],
    acceptance_trend: ['スコア変動傾向（統合）', 'スコア変動傾向(統合)'],

    // インサイト系
    core_motivation: ['コアモチベーション'],
    main_concern: ['主要懸念事項'],
    fit_assessment: ['フィット度総合判定'],
    recommended_action: ['推奨アクション'],
    focus_point: ['注目ポイント']
  };

  // 列の検証
  const validation = validateRequiredColumns(backupMetadata, requiredColumns);

  Logger.log('列検出結果:');
  Logger.log(`検出成功: ${Object.keys(validation.foundColumns).length}列`);
  Logger.log(`検出失敗: ${validation.missingColumns.length}列`);
  Logger.log('');

  if (validation.missingColumns.length > 0) {
    Logger.log('⚠️ 以下の列が見つかりません（オプション列として扱います）:');
    validation.missingColumns.forEach(missing => {
      Logger.log(`- ${missing.logicalName} (候補: ${missing.candidates.join(', ')})`);
    });
    Logger.log('');
  }

  // 基本列の必須チェック
  const essentialColumns = ['candidate_id', 'name', 'updated_at'];
  const missingEssential = essentialColumns.filter(col =>
    !validation.foundColumns.hasOwnProperty(col)
  );

  if (missingEssential.length > 0) {
    throw new Error(
      `必須列が見つかりません: ${missingEssential.join(', ')}\n` +
      `利用可能な列: ${backupMetadata.headers.join(', ')}`
    );
  }

  Logger.log('✅ 必須列を検出しました');
  Logger.log('');

  // データを抽出
  const dataRows = backupMetadata.totalRows - 1;
  if (dataRows === 0) {
    Logger.log('⚠️ バックアップシートにデータがありません');
    return {
      success: true,
      dataCount: 0,
      data: []
    };
  }

  const allData = backupSheet.getRange(2, 1, dataRows, backupMetadata.totalColumns).getValues();
  const extractedData = [];

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const candidateId = row[validation.foundColumns.candidate_id];

    if (!candidateId) continue;

    // 安全にデータを取得するヘルパー関数
    const getValue = (logicalName, defaultValue = '') => {
      const colIndex = validation.foundColumns[logicalName];
      return colIndex !== undefined ? (row[colIndex] || defaultValue) : defaultValue;
    };

    extractedData.push({
      // 基本情報
      candidate_id: candidateId,
      name: getValue('name'),
      email: getValue('email'),
      status: getValue('status'),
      updated_at: getValue('updated_at'),

      // スコア情報
      scores: {
        latest_pass_rate: getValue('latest_pass_rate'),
        prev_pass_rate: getValue('prev_pass_rate'),
        prev2_pass_rate: getValue('prev2_pass_rate'),
        pass_diff1: getValue('pass_diff1'),
        pass_diff2: getValue('pass_diff2'),
        pass_diff3: getValue('pass_diff3'),
        pass_trend: getValue('pass_trend'),
        latest_acceptance_integrated: getValue('latest_acceptance_integrated'),
        prev_acceptance: getValue('prev_acceptance'),
        prev2_acceptance: getValue('prev2_acceptance'),
        acceptance_diff1: getValue('acceptance_diff1'),
        acceptance_diff2: getValue('acceptance_diff2'),
        acceptance_diff3: getValue('acceptance_diff3'),
        acceptance_trend: getValue('acceptance_trend')
      },

      // インサイト情報
      insights: {
        core_motivation: getValue('core_motivation'),
        main_concern: getValue('main_concern'),
        fit_assessment: getValue('fit_assessment'),
        recommended_action: getValue('recommended_action'),
        focus_point: getValue('focus_point')
      }
    });
  }

  Logger.log(`✅ ${extractedData.length}件のデータを抽出`);
  Logger.log('====================================');
  Logger.log('');

  return {
    success: true,
    dataCount: extractedData.length,
    data: extractedData
  };
}

/**
 * 抽出データをCandidate_Scoresに投入
 * @param {Array} extractedData - extractDataFromBackup()の戻り値のdataプロパティ
 */
function populateCandidateScores(extractedData) {
  Logger.log('====================================');
  Logger.log('Candidate_Scoresにデータ投入開始');
  Logger.log('====================================');

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Candidate_Scores');

  if (!sheet) {
    throw new Error('Candidate_Scoresシートが見つかりません');
  }

  // データをクリア（ヘッダーは残す）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    Logger.log('既存データをクリアしました');
  }

  if (extractedData.length === 0) {
    Logger.log('⚠️ 投入するデータがありません');
    Logger.log('====================================');
    Logger.log('');
    return;
  }

  // データを準備（20列）
  const rowsData = extractedData.map(item => [
    item.candidate_id,                           // A: candidate_id
    item.name,                                    // B: 氏名
    item.updated_at,                              // C: 最終更新日時
    item.scores.latest_pass_rate,                // D: 最新_合格可能性
    item.scores.prev_pass_rate,                  // E: 前回_合格可能性
    item.scores.prev2_pass_rate,                 // F: 前々回_合格可能性
    item.scores.pass_diff1,                      // G: スコア差分1
    item.scores.pass_diff2,                      // H: スコア差分2
    item.scores.pass_diff3,                      // I: スコア差分3
    item.scores.pass_trend,                      // J: スコア変動傾向
    '',                                           // K: dify_workflow_id
    '',                                           // L: dify_run_id
    '',                                           // M: dify_updated_at
    item.scores.latest_acceptance_integrated,    // N: 最新_承諾可能性
    item.scores.prev_acceptance,                 // O: 前回_承諾可能性
    item.scores.prev2_acceptance,                // P: 前々回_承諾可能性
    item.scores.acceptance_diff1,                // Q: スコア差分1_承諾
    item.scores.acceptance_diff2,                // R: スコア差分2_承諾
    item.scores.acceptance_diff3,                // S: スコア差分3_承諾
    item.scores.acceptance_trend                 // T: スコア変動傾向_承諾
  ]);

  // データを書き込み
  sheet.getRange(2, 1, rowsData.length, rowsData[0].length)
    .setValues(rowsData);

  Logger.log(`✅ ${rowsData.length}件のデータを投入`);
  Logger.log('====================================');
  Logger.log('');
}

/**
 * 抽出データをCandidate_Insightsに投入
 * @param {Array} extractedData
 */
function populateCandidateInsights(extractedData) {
  Logger.log('====================================');
  Logger.log('Candidate_Insightsにデータ投入開始');
  Logger.log('====================================');

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Candidate_Insights');

  if (!sheet) {
    throw new Error('Candidate_Insightsシートが見つかりません');
  }

  // データをクリア
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    Logger.log('既存データをクリアしました');
  }

  if (extractedData.length === 0) {
    Logger.log('⚠️ 投入するデータがありません');
    Logger.log('====================================');
    Logger.log('');
    return;
  }

  // データを準備（10列）
  const rowsData = extractedData.map(item => [
    item.candidate_id,                    // A: candidate_id
    item.name,                            // B: 氏名
    item.updated_at,                      // C: 最終更新日時
    item.insights.core_motivation,        // D: コアモチベーション
    item.insights.main_concern,           // E: 主要懸念事項
    item.insights.fit_assessment,         // F: フィット度総合判定
    item.insights.recommended_action,     // G: 推奨アクション
    item.insights.focus_point,            // H: 注目ポイント
    '',                                   // I: dify_workflow_id
    ''                                    // J: dify_run_id
  ]);

  // データを書き込み
  sheet.getRange(2, 1, rowsData.length, rowsData[0].length)
    .setValues(rowsData);

  Logger.log(`✅ ${rowsData.length}件のデータを投入`);
  Logger.log('====================================');
  Logger.log('');
}

// ========================================
// Phase 3: 販売対応
// ========================================

/**
 * デモ用サンプルデータを投入
 */
function insertSampleData() {
  Logger.log('====================================');
  Logger.log('サンプルデータ投入開始');
  Logger.log('====================================');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = new Date();

  // Candidates_Masterにサンプル候補者を追加
  const masterSheet = ss.getSheetByName('Candidates_Master');
  if (!masterSheet) {
    throw new Error('Candidates_Masterシートが見つかりません');
  }

  const sampleCandidates = [
    [
      'DEMO_001',         // candidate_id
      '山田太郎',         // 氏名
      'yamada@example.com', // メールアドレス
      '一次面接完了',    // 現在ステータス
      '新卒',             // 採用区分
      timestamp,          // 最終更新日時
      75,                 // 最新_合格可能性（VLOOKUP）
      80,                 // 最新_承諾可能性（VLOOKUP）
      'B',                // 総合ランク
      new Date(2024, 0, 15), // 応募日
      new Date(2024, 1, 1),  // 初回面談日
      '鈴木面接官',      // 初回面談担当者
      '出席',             // 面談出席
      new Date(2024, 1, 15), // 1次面接日
      '合格',             // 1次面接合否
      '',                 // 2次面接日
      '',                 // 2次面接合否
      '',                 // 最終面接日
      '',                 // 最終面接合否
      '',                 // 内定日
      ''                  // 承諾日時
    ],
    [
      'DEMO_002',
      '佐藤花子',
      'sato@example.com',
      '書類選考中',
      '中途',
      timestamp,
      60,
      70,
      'C',
      new Date(2024, 0, 20),
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      ''
    ],
    [
      'DEMO_003',
      '鈴木一郎',
      'suzuki@example.com',
      '最終面接待ち',
      '新卒',
      timestamp,
      85,
      90,
      'A',
      new Date(2024, 0, 10),
      new Date(2024, 0, 25),
      '田中面接官',
      '出席',
      new Date(2024, 1, 5),
      '合格',
      new Date(2024, 1, 20),
      '合格',
      '',
      '',
      '',
      ''
    ]
  ];

  const lastRow = masterSheet.getLastRow();
  masterSheet.getRange(lastRow + 1, 1, sampleCandidates.length, sampleCandidates[0].length)
    .setValues(sampleCandidates);

  Logger.log(`✅ ${sampleCandidates.length}名のサンプル候補者を追加`);

  // Candidate_Scoresにもデータ追加
  const scoresSheet = ss.getSheetByName('Candidate_Scores');
  if (scoresSheet) {
    const sampleScores = [
      ['DEMO_001', '山田太郎', timestamp, 75, 70, 65, 5, 5, 5, '上昇傾向', '', '', '', 80, 75, 70, 5, 5, 5, '上昇傾向'],
      ['DEMO_002', '佐藤花子', timestamp, 60, 55, 50, 5, 5, 5, '上昇傾向', '', '', '', 70, 68, 65, 2, 3, 3, '緩やかな上昇'],
      ['DEMO_003', '鈴木一郎', timestamp, 85, 83, 80, 2, 3, 5, '上昇傾向', '', '', '', 90, 88, 85, 2, 3, 5, '上昇傾向']
    ];

    const scoresLastRow = scoresSheet.getLastRow();
    scoresSheet.getRange(scoresLastRow + 1, 1, sampleScores.length, sampleScores[0].length)
      .setValues(sampleScores);

    Logger.log('✅ サンプルスコアを追加');
  }

  // Candidate_Insightsにもデータ追加
  const insightsSheet = ss.getSheetByName('Candidate_Insights');
  if (insightsSheet) {
    const sampleInsights = [
      ['DEMO_001', '山田太郎', timestamp, '成長機会を重視', '給与水準', '高い適合性', '次回面接で詳細確認', '積極的な姿勢', '', ''],
      ['DEMO_002', '佐藤花子', timestamp, 'ワークライフバランス', '勤務地', '中程度の適合性', '条件面の調整が必要', '慎重に検討中', '', ''],
      ['DEMO_003', '鈴木一郎', timestamp, '技術力向上', '特になし', '非常に高い適合性', '早期にオファー提示', '高い志望度', '', '']
    ];

    const insightsLastRow = insightsSheet.getLastRow();
    insightsSheet.getRange(insightsLastRow + 1, 1, sampleInsights.length, sampleInsights[0].length)
      .setValues(sampleInsights);

    Logger.log('✅ サンプルインサイトを追加');
  }

  Logger.log('====================================');
  Logger.log('✅ サンプルデータ投入完了');
  Logger.log('====================================');
  Logger.log('');

  SpreadsheetApp.getUi().alert(
    '✅ サンプルデータ投入完了',
    `${sampleCandidates.length}名の候補者データを追加しました。\n` +
    'Candidates_Masterシートを確認してください。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 顧客向けREADMEシートを作成
 */
function createCustomerReadme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存のREADMEシートを削除
  const existingReadme = ss.getSheetByName('📖 README（必読）');
  if (existingReadme) {
    ss.deleteSheet(existingReadme);
  }

  // 新規シート作成
  const readmeSheet = ss.insertSheet('📖 README（必読）', 0); // 最初のシートに配置
  readmeSheet.setTabColor('#ff9900');

  // コンテンツを作成
  const content = [
    ['採用参謀AI - 候補者管理システム'],
    [''],
    ['■ はじめに'],
    ['このスプレッドシートは、候補者の選考プロセスを一元管理するためのシステムです。'],
    ['AI分析により、合格可能性や承諾可能性を自動で算出します。'],
    [''],
    ['■ 主要なシート'],
    ['📊 Candidates_Master - 候補者の基本情報・ステータス管理'],
    ['📈 Candidate_Scores - AIによるスコア・評価データ'],
    ['💡 Candidate_Insights - モチベーション・懸念事項の分析結果'],
    ['📝 Evaluation_Log - 面接評価の記録'],
    ['📧 Contact_History - 候補者との接点履歴'],
    [''],
    ['■ 初期設定（必ず実施）'],
    ['1. 拡張機能 → Apps Script を開く'],
    ['2. メニューバーから「📊 採用参謀AI」→「✅ 動作確認」を実行'],
    ['3. 権限の許可を求められたら「許可」をクリック'],
    ['4. 動作確認が完了したら利用開始可能です'],
    [''],
    ['■ 基本的な使い方'],
    ['【候補者の追加】'],
    ['1. Candidates_Masterシートを開く'],
    ['2. 最下行に新しい候補者情報を入力'],
    ['3. candidate_idは「会社名_YYYY_連番」形式で入力（例: ABC_2024_001）'],
    [''],
    ['【面接評価の入力】'],
    ['1. Evaluation_Logシートを開く'],
    ['2. 面接後、評価データを入力'],
    ['3. AIが自動でスコアを計算し、Candidate_Scoresに反映されます'],
    [''],
    ['■ デモモードについて'],
    ['メニューの「🎬 デモモードON」を実行すると、サンプルデータが表示されます。'],
    ['実際の運用前に、このデータで動作を確認することをお勧めします。'],
    [''],
    ['■ サポート'],
    ['ご不明な点がございましたら、弊社サポートまでお問い合わせください。'],
    ['Email: support@example.com']
  ];

  // データを書き込み
  readmeSheet.getRange(1, 1, content.length, 1).setValues(content);

  // 書式設定
  readmeSheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  readmeSheet.getRange('A3').setFontSize(14).setFontWeight('bold');
  readmeSheet.getRange('A7').setFontSize(14).setFontWeight('bold');
  readmeSheet.getRange('A13').setFontSize(14).setFontWeight('bold');
  readmeSheet.getRange('A19').setFontSize(14).setFontWeight('bold');
  readmeSheet.getRange('A27').setFontSize(14).setFontWeight('bold');
  readmeSheet.getRange('A31').setFontSize(14).setFontWeight('bold');

  // 列幅を調整
  readmeSheet.setColumnWidth(1, 800);

  Logger.log('✅ 顧客向けREADMEシートを作成');
}

// 続く...
