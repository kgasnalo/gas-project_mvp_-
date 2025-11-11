/**
 * 全データ検証を設定
 */
function setupAllDataValidation() {
  try {
    Logger.log('データ検証の設定を開始します...');
    
    setupCandidatesMasterValidation();
    setupEvaluationLogValidation();
    setupEngagementLogValidation();
    setupOtherSheetsValidation();
    
    Logger.log('✅ 全データ検証の設定が完了しました');
    
  } catch (error) {
    logError('setupAllDataValidation', error);
    throw error;
  }
}

/**
 * Candidates_Masterのデータ検証を設定
 */
function setupCandidatesMasterValidation() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);
  
  if (!sheet) return;
  
  // C列: 現在ステータス
  setDropdownValidation(
    sheet,
    'C2:C1000',
    CONFIG.VALIDATION_OPTIONS.STATUS,
    'ステータスを選択してください'
  );
  
  // E列: 採用区分
  setDropdownValidation(
    sheet,
    'E2:E1000',
    CONFIG.VALIDATION_OPTIONS.CATEGORY,
    '新卒または中途を選択してください'
  );

  // U列: 予測の信頼度 - 自動計算列のためデータ検証なし（Phase 1で関数を設定）
  // AL列: アクション緊急度 - 自動計算列のためデータ検証なし（Phase 1で関数を設定）

  // AM列: アクション実行状況
  setDropdownValidation(
    sheet,
    'AM2:AM1000',
    CONFIG.VALIDATION_OPTIONS.ACTION_STATUS,
    '実行状況を選択してください'
  );

  // 【新規追加】選考ステップ追跡列のデータ検証
  // AS列: 初回面談実施ステータス
  setDropdownValidation(
    sheet,
    'AS2:AS1000',
    CONFIG.VALIDATION_OPTIONS.IMPLEMENTATION_STATUS,
    '実施ステータスを選択してください'
  );

  // AU列: 適性検査実施ステータス
  setDropdownValidation(
    sheet,
    'AU2:AU1000',
    CONFIG.VALIDATION_OPTIONS.IMPLEMENTATION_STATUS,
    '実施ステータスを選択してください'
  );

  // AW列: 1次面接結果
  setDropdownValidation(
    sheet,
    'AW2:AW1000',
    CONFIG.VALIDATION_OPTIONS.TEST_RESULT,
    '面接結果を選択してください'
  );

  // AX列: 社員面談実施回数（チェックボックス形式ではなく数値入力）
  // 数値のみ許可（0以上の整数）
  const axRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false)
    .setHelpText('0以上の整数を入力してください')
    .build();
  sheet.getRange('AX2:AX1000').setDataValidation(axRule);

  // AZ列: 社員面談実施ステータス
  setDropdownValidation(
    sheet,
    'AZ2:AZ1000',
    CONFIG.VALIDATION_OPTIONS.IMPLEMENTATION_STATUS,
    '実施ステータスを選択してください'
  );

  // BB列: 2次面接結果
  setDropdownValidation(
    sheet,
    'BB2:BB1000',
    CONFIG.VALIDATION_OPTIONS.TEST_RESULT,
    '面接結果を選択してください'
  );

  // BD列: 最終面接結果
  setDropdownValidation(
    sheet,
    'BD2:BD1000',
    CONFIG.VALIDATION_OPTIONS.TEST_RESULT,
    '面接結果を選択してください'
  );

  Logger.log('✅ Candidates_Masterのデータ検証を設定しました');
}

/**
 * Evaluation_Logのデータ検証を設定
 */
function setupEvaluationLogValidation() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.EVALUATION_LOG);
  
  if (!sheet) return;
  
  // E列: 選考フェーズ（6項目に拡張）
  setDropdownValidation(
    sheet,
    'E2:E1000',
    ['初回面談', '社員面談', '1次面接', '2次面接', '最終面接', '内定後'],
    '選考フェーズを選択してください'
  );
  
  // P列: highest_risk_level
  setDropdownValidation(
    sheet,
    'P2:P1000',
    CONFIG.VALIDATION_OPTIONS.RISK_LEVEL,
    'リスクレベルを選択してください'
  );
  
  // R列: recommendation
  setDropdownValidation(
    sheet,
    'R2:R1000',
    CONFIG.VALIDATION_OPTIONS.RECOMMENDATION,
    '推奨を選択してください'
  );
  
  Logger.log('✅ Evaluation_Logのデータ検証を設定しました');
}

/**
 * Engagement_Logのデータ検証を設定
 */
function setupEngagementLogValidation() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.ENGAGEMENT_LOG);
  
  if (!sheet) return;
  
  // E列: contact_type
  setDropdownValidation(
    sheet,
    'E2:E1000',
    CONFIG.VALIDATION_OPTIONS.CONTACT_TYPE,
    '接点タイプを選択してください'
  );
  
  // I列: confidence_level
  setDropdownValidation(
    sheet,
    'I2:I1000',
    CONFIG.VALIDATION_OPTIONS.CONFIDENCE,
    '信頼度を選択してください'
  );
  
  // T列: action_priority
  setDropdownValidation(
    sheet,
    'T2:T1000',
    ['高', '中', '低'],
    '優先度を選択してください'
  );
  
  Logger.log('✅ Engagement_Logのデータ検証を設定しました');
}

/**
 * その他のシートのデータ検証を設定
 */
function setupOtherSheetsValidation() {
  // Risk
  const riskSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.RISK);
  
  if (riskSheet) {
    setDropdownValidation(
      riskSheet,
      'E2:E1000',
      CONFIG.VALIDATION_OPTIONS.RISK_LEVEL,
      'リスクレベルを選択してください'
    );
  }
  
  // Contact_History
  const contactSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.CONTACT_HISTORY);
  
  if (contactSheet) {
    setDropdownValidation(
      contactSheet,
      'E2:E1000',
      CONFIG.VALIDATION_OPTIONS.CONTACT_TYPE,
      '接点タイプを選択してください'
    );
  }
  
  Logger.log('✅ その他のシートのデータ検証を設定しました');
}