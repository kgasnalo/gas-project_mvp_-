/**
 * 全条件付き書式を設定
 */
function setupAllConditionalFormatting() {
  try {
    Logger.log('条件付き書式の設定を開始します...');
    
    setupCandidatesMasterConditionalFormatting();
    setupEvaluationLogConditionalFormatting();
    setupEngagementLogConditionalFormatting();
    
    Logger.log('✅ 全条件付き書式の設定が完了しました');
    
  } catch (error) {
    logError('setupAllConditionalFormatting', error);
    throw error;
  }
}

/**
 * Candidates_Masterの条件付き書式を設定
 */
function setupCandidatesMasterConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);
  
  if (!sheet) return;
  
  // 既存の条件付き書式ルールをクリア（重複を避けるため）
  sheet.clearConditionalFormatRules();
  
  const rules = [];
  
  // H列: 最新_合格可能性（80%以上: 緑、60-79%: 黄、60%未満: 赤）
  rules.push(...createThreeColorRules(sheet, 'H2:H1000', 80, 60));
  
  // P列: 最新_承諾可能性（AI予測）
  rules.push(...createThreeColorRules(sheet, 'P2:P1000', 70, 50));
  
  // R列: 最新_承諾可能性（統合）
  rules.push(...createThreeColorRules(sheet, 'R2:R1000', 70, 50));
  
  // K-O列: 4軸スコア（0.8以上: 緑、0.6-0.79: 黄、0.6未満: 赤）
  rules.push(...createThreeColorRules(sheet, 'K2:O1000', 0.8, 0.6));
  
  // V-X列: 各種スコア（70以上: 緑、50-69: 黄、50未満: 赤）
  rules.push(...createThreeColorRules(sheet, 'V2:X1000', 70, 50));
  
  // AL列: アクション緊急度（色分け）
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('CRITICAL')
      .setBackground(CONFIG.COLORS.CRITICAL)
      .setFontColor('#ffffff')
      .setBold(true)
      .setRanges([sheet.getRange('AL2:AL1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('HIGH')
      .setBackground(CONFIG.COLORS.HIGH)
      .setRanges([sheet.getRange('AL2:AL1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('NORMAL')
      .setBackground(CONFIG.COLORS.NORMAL)
      .setFontColor('#ffffff')
      .setRanges([sheet.getRange('AL2:AL1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('LOW')
      .setBackground(CONFIG.COLORS.LOW)
      .setFontColor('#ffffff')
      .setRanges([sheet.getRange('AL2:AL1000')])
      .build()
  );
  
  // AO列: 注力度合いスコア（1.5以上: 緑、1.0-1.49: 黄、1.0未満: 赤）
  rules.push(...createThreeColorRules(sheet, 'AO2:AO1000', 1.5, 1.0));
  
  // C列: 現在ステータス（色分け）
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('承諾')
      .setBackground('#d9ead3')
      .setBold(true)
      .setRanges([sheet.getRange('C2:C1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('辞退')
      .setBackground('#f4cccc')
      .setRanges([sheet.getRange('C2:C1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('見送り')
      .setBackground('#efefef')
      .setRanges([sheet.getRange('C2:C1000')])
      .build()
  );
  
  // ルールを適用
  sheet.setConditionalFormatRules(rules);
  
  Logger.log('✅ Candidates_Masterの条件付き書式を設定しました');
}

/**
 * Evaluation_Logの条件付き書式を設定
 */
function setupEvaluationLogConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.EVALUATION_LOG);
  
  if (!sheet) return;
  
  sheet.clearConditionalFormatRules();
  
  const rules = [];
  
  // G列: overall_score（0.8以上: 緑、0.6-0.79: 黄、0.6未満: 赤）
  rules.push(...createThreeColorRules(sheet, 'G2:G1000', 0.8, 0.6));
  
  // H, J, L, N列: 各軸スコア
  rules.push(...createThreeColorRules(sheet, 'H2:H1000', 0.8, 0.6));
  rules.push(...createThreeColorRules(sheet, 'J2:J1000', 0.8, 0.6));
  rules.push(...createThreeColorRules(sheet, 'L2:L1000', 0.8, 0.6));
  rules.push(...createThreeColorRules(sheet, 'N2:N1000', 0.8, 0.6));
  
  // P列: highest_risk_level（色分け）
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Critical')
      .setBackground(CONFIG.COLORS.CRITICAL)
      .setFontColor('#ffffff')
      .setRanges([sheet.getRange('P2:P1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('High')
      .setBackground(CONFIG.COLORS.HIGH)
      .setRanges([sheet.getRange('P2:P1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Medium')
      .setBackground(CONFIG.COLORS.MEDIUM_SCORE)
      .setRanges([sheet.getRange('P2:P1000')])
      .build()
  );
  
  // Q列: pass_probability
  rules.push(...createThreeColorRules(sheet, 'Q2:Q1000', 80, 60));
  
  // R列: recommendation（色分け）
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pass')
      .setBackground('#d9ead3')
      .setBold(true)
      .setRanges([sheet.getRange('R2:R1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Fail')
      .setBackground('#f4cccc')
      .setRanges([sheet.getRange('R2:R1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Conditional')
      .setBackground('#fff2cc')
      .setRanges([sheet.getRange('R2:R1000')])
      .build()
  );
  
  sheet.setConditionalFormatRules(rules);
  
  Logger.log('✅ Evaluation_Logの条件付き書式を設定しました');
}

/**
 * Engagement_Logの条件付き書式を設定
 */
function setupEngagementLogConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.ENGAGEMENT_LOG);
  
  if (!sheet) return;
  
  sheet.clearConditionalFormatRules();
  
  const rules = [];
  
  // F-H列: 承諾可能性（70%以上: 緑、50-69%: 黄、50%未満: 赤）
  rules.push(...createThreeColorRules(sheet, 'F2:H1000', 70, 50));
  
  // J-L列: 各種スコア（70以上: 緑、50-69: 黄、50未満: 赤）
  rules.push(...createThreeColorRules(sheet, 'J2:L1000', 70, 50));
  
  // Q列: competitive_advantage（70以上: 緑、50-69: 黄、50未満: 赤）
  rules.push(...createThreeColorRules(sheet, 'Q2:Q1000', 70, 50));
  
  // I列: confidence_level（色分け）
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('高')
      .setBackground('#d9ead3')
      .setRanges([sheet.getRange('I2:I1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('中')
      .setBackground('#fff2cc')
      .setRanges([sheet.getRange('I2:I1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('低')
      .setBackground('#f4cccc')
      .setRanges([sheet.getRange('I2:I1000')])
      .build()
  );
  
  // T列: action_priority（色分け）
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('高')
      .setBackground(CONFIG.COLORS.CRITICAL)
      .setFontColor('#ffffff')
      .setRanges([sheet.getRange('T2:T1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('中')
      .setBackground(CONFIG.COLORS.HIGH)
      .setRanges([sheet.getRange('T2:T1000')])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('低')
      .setBackground(CONFIG.COLORS.LOW)
      .setFontColor('#ffffff')
      .setRanges([sheet.getRange('T2:T1000')])
      .build()
  );
  
  sheet.setConditionalFormatRules(rules);
  
  Logger.log('✅ Engagement_Logの条件付き書式を設定しました');
}

/**
 * 3段階の色分けルールを作成
 * @param {Sheet} sheet - シートオブジェクト
 * @param {string} rangeA1 - 範囲（A1形式）
 * @param {number} highThreshold - 高スコアの閾値
 * @param {number} lowThreshold - 低スコアの閾値
 * @return {Array} 条件付き書式ルールの配列
 */
function createThreeColorRules(sheet, rangeA1, highThreshold, lowThreshold) {
  const targetRange = sheet.getRange(rangeA1);
  
  const rules = [];
  
  // 高スコア（緑）
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(highThreshold)
      .setBackground(CONFIG.COLORS.HIGH_SCORE)
      .setRanges([targetRange])
      .build()
  );
  
  // 中スコア（黄）
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(lowThreshold, highThreshold - 0.01)
      .setBackground(CONFIG.COLORS.MEDIUM_SCORE)
      .setRanges([targetRange])
      .build()
  );
  
  // 低スコア（赤）
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(lowThreshold)
      .setBackground(CONFIG.COLORS.LOW_SCORE)
      .setRanges([targetRange])
      .build()
  );
  
  return rules;
}
