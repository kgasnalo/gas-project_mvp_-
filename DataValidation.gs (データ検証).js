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
  
  // U列: 予測の信頼度
  setDropdownValidation(
    sheet,
    'U2:U1000',
    CONFIG.VALIDATION_OPTIONS.CONFIDENCE,
    '信頼度を選択してください'
  );
  
  // AL列: アクション緊急度
  setDropdownValidation(
    sheet,
    'AL2:AL1000',
    CONFIG.VALIDATION_OPTIONS.URGENCY,
    '緊急度を選択してください'
  );
  
  // AM列: アクション実行状況
  setDropdownValidation(
    sheet,
    'AM2:AM1000',
    CONFIG.VALIDATION_OPTIONS.ACTION_STATUS,
    '実行状況を選択してください'
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
  
  // E列: 選考フェーズ
  setDropdownValidation(
    sheet,
    'E2:E1000',
    ['面談', '1次面接', '2次面接', '最終面接'],
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