/**
 * 全シートを作成
 */
function createAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // シート名のリスト（表示順）
  const sheetNames = Object.values(CONFIG.SHEET_NAMES);
  
  // 各シートを作成
  sheetNames.forEach((name, index) => {
    getOrCreateSheet(name, index);
  });
  
  // READMEシートを追加
  getOrCreateSheet('README', sheetNames.length);
  
  Logger.log(`✅ 全${sheetNames.length + 1}シートの作成が完了しました`);
}

/**
 * Candidates_Masterシートを設定
 */
function setupCandidatesMaster() {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);
    
    // ヘッダー行の定義（AQ列まで）
    const headers = [
      // 基本情報 A-G列
      'candidate_id', '氏名', '現在ステータス', '最終更新日時', '採用区分', '担当面接官', '応募日',

      // 合格可能性 H-O列
      '最新_合格可能性', '前回_合格可能性', '合格可能性_増減', '最新_合計スコア',
      '最新_Philosophy', '最新_Strategy', '最新_Motivation', '最新_Execution',

      // 承諾可能性 P-Z列
      '最新_承諾可能性（AI予測）', '最新_承諾可能性（人間の直感）', '最新_承諾可能性（統合）',
      '前回_承諾可能性', '承諾可能性_増減', '予測の信頼度', '志望度スコア', '競合優位性スコア',
      '懸念解消度スコア', 'コアモチベーション', '主要懸念事項',

      // 競合情報 AA-AD列
      '競合企業1', '競合企業2', '競合企業3', '自社の志望順位',

      // 接点管理 AE-AI列
      '前回接点日', '前回接点からの空白期間', '接点回数', '平均接点間隔', '次回接点予定日',

      // アクション管理 AJ-AN列
      '次推奨アクション', 'アクション期限', 'アクション緊急度', 'アクション実行状況', 'アクション実行者',

      // 注力度合い AO-AP列
      '注力度合いスコア', '緊急度係数',

      // 連絡先 AQ列
      'メールアドレス',

      // 【新規追加】選考ステップ追跡 AR-BD列
      '初回面談日', '初回面談実施ステータス',
      '適性検査日', '適性検査実施ステータス',
      '1次面接日', '1次面接結果',
      '社員面談実施回数', '社員面談日（最終）', '社員面談実施ステータス',
      '2次面接日', '2次面接結果',
      '最終面接日', '最終面接結果'
    ];
    
    setHeaders(sheet, headers);
    
    // 列幅の詳細設定
    const columnWidths = {
      1: 100,   // A: candidate_id
      2: 120,   // B: 氏名
      3: 120,   // C: 現在ステータス
      4: 160,   // D: 最終更新日時
      5: 100,   // E: 採用区分
      6: 120,   // F: 担当面接官
      7: 120,   // G: 応募日
      8: 130,   // H: 最新_合格可能性
      9: 130,   // I: 前回_合格可能性
      10: 130,  // J: 合格可能性_増減
      11: 120,  // K: 最新_合計スコア
      12: 130,  // L: 最新_Philosophy
      13: 130,  // M: 最新_Strategy
      14: 130,  // N: 最新_Motivation
      15: 130,  // O: 最新_Execution
      16: 180,  // P: 最新_承諾可能性（AI予測）
      17: 200,  // Q: 最新_承諾可能性（人間の直感）
      18: 180,  // R: 最新_承諾可能性（統合）
      19: 150,  // S: 前回_承諾可能性
      20: 150,  // T: 承諾可能性_増減
      21: 120,  // U: 予測の信頼度
      22: 120,  // V: 志望度スコア
      23: 150,  // W: 競合優位性スコア
      24: 150,  // X: 懸念解消度スコア
      25: 200,  // Y: コアモチベーション
      26: 200,  // Z: 主要懸念事項
      27: 150,  // AA: 競合企業1
      28: 150,  // AB: 競合企業2
      29: 150,  // AC: 競合企業3
      30: 140,  // AD: 自社の志望順位
      31: 140,  // AE: 前回接点日
      32: 180,  // AF: 前回接点からの空白期間
      33: 100,  // AG: 接点回数
      34: 140,  // AH: 平均接点間隔
      35: 140,  // AI: 次回接点予定日
      36: 300,  // AJ: 次推奨アクション
      37: 140,  // AK: アクション期限
      38: 140,  // AL: アクション緊急度
      39: 150,  // AM: アクション実行状況
      40: 140,  // AN: アクション実行者
      41: 140,  // AO: 注力度合いスコア
      42: 120,  // AP: 緊急度係数
      43: 200,  // AQ: メールアドレス
      44: 140,  // AR: 初回面談日（新規）
      45: 150,  // AS: 初回面談実施ステータス（新規）
      46: 140,  // AT: 適性検査日（新規）
      47: 150,  // AU: 適性検査実施ステータス（新規）
      48: 140,  // AV: 1次面接日（新規）
      49: 140,  // AW: 1次面接結果（新規）
      50: 140,  // AX: 社員面談実施回数（新規）
      51: 140,  // AY: 社員面談日（最終）（新規）
      52: 150,  // AZ: 社員面談実施ステータス（新規）
      53: 140,  // BA: 2次面接日（新規）
      54: 150,  // BB: 2次面接実施ステータス（新規）
      55: 140,  // BC: 最終面接日（新規）
      56: 150   // BD: 最終面接実施ステータス（新規）
    };
    
    // 列幅を設定
    for (let col in columnWidths) {
      sheet.setColumnWidth(parseInt(col), columnWidths[col]);
    }
    
    // 数値列のフォーマット設定
    sheet.getRange('H2:J1000').setNumberFormat('0.00"%"');  // 合格可能性（%）
    sheet.getRange('K2:O1000').setNumberFormat('0.00');     // スコア（0-1）
    sheet.getRange('P2:T1000').setNumberFormat('0.00"%"');  // 承諾可能性（%）
    sheet.getRange('V2:X1000').setNumberFormat('0');        // スコア（0-100）
    sheet.getRange('AO2:AP1000').setNumberFormat('0.00');   // 注力度合い
    
    // 日付列のフォーマット設定
    sheet.getRange('D2:D1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange('G2:G1000').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('AE2:AE1000').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('AI2:AI1000').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('AK2:AK1000').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('AR2:AR1000').setNumberFormat('yyyy-mm-dd'); // 初回面談日
    sheet.getRange('AT2:AT1000').setNumberFormat('yyyy-mm-dd'); // 適性検査日
    sheet.getRange('AV2:AV1000').setNumberFormat('yyyy-mm-dd'); // 1次面接日
    sheet.getRange('AY2:AY1000').setNumberFormat('yyyy-mm-dd'); // 社員面談日（最終）
    sheet.getRange('BA2:BA1000').setNumberFormat('yyyy-mm-dd'); // 2次面接日
    sheet.getRange('BC2:BC1000').setNumberFormat('yyyy-mm-dd'); // 最終面接日

    // 整数列のフォーマット設定
    sheet.getRange('AX2:AX1000').setNumberFormat('0'); // 社員面談実施回数
    
    // テキストの折り返しを設定
    sheet.getRange('Y2:Z1000').setWrap(true);  // コアモチベーション、主要懸念事項
    sheet.getRange('AJ2:AJ1000').setWrap(true); // 次推奨アクション

    // ========== 自動計算列のデータ検証をクリア ==========
    // U列（予測の信頼度）とAL列（アクション緊急度）は自動計算列のため、
    // データ検証が設定されているとエラーになるのでクリアする
    sheet.getRange('U2:U1000').clearDataValidations();
    sheet.getRange('AL2:AL1000').clearDataValidations();

    // ========== 自動計算列の関数設定 ==========
    // ⚠️ Evaluation_Logの範囲をA:S（T列追加に対応）に修正

    // H列: 最新_合格可能性
    sheet.getRange('H2').setFormula('=IFERROR(QUERY(Evaluation_Log!A:S, "SELECT Q WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), 0)');

    // I列: 前回_合格可能性
    sheet.getRange('I2').setFormula('=IFERROR(QUERY(Evaluation_Log!A:S, "SELECT Q WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1 OFFSET 1", 0), 0)');

    // J列: 合格可能性_増減
    sheet.getRange('J2').setFormula('=IF(OR(H2="", I2=""), "", H2-I2)');

    // K列: 最新_合計スコア
    sheet.getRange('K2').setFormula('=IFERROR(QUERY(Evaluation_Log!A:S, "SELECT G WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), 0)');

    // L列: 最新_Philosophy
    sheet.getRange('L2').setFormula('=IFERROR(QUERY(Evaluation_Log!A:S, "SELECT H WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), 0)');

    // M列: 最新_Strategy
    sheet.getRange('M2').setFormula('=IFERROR(QUERY(Evaluation_Log!A:S, "SELECT J WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), 0)');

    // N列: 最新_Motivation
    sheet.getRange('N2').setFormula('=IFERROR(QUERY(Evaluation_Log!A:S, "SELECT L WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), 0)');

    // O列: 最新_Execution
    sheet.getRange('O2').setFormula('=IFERROR(QUERY(Evaluation_Log!A:S, "SELECT N WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), 0)');

    // P列: 最新_承諾可能性（AI予測）
    sheet.getRange('P2').setFormula('=IFERROR(QUERY(Engagement_Log!A:U, "SELECT G WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), 0)');

    // R列: 最新_承諾可能性（統合）
    sheet.getRange('R2').setFormula('=IF(OR(P2="", Q2=""), "", P2*0.7 + Q2*0.3)');

    // S列: 前回_承諾可能性
    sheet.getRange('S2').setFormula('=IFERROR(QUERY(Engagement_Log!A:U, "SELECT G WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1 OFFSET 1", 0), 0)');

    // T列: 承諾可能性_増減
    sheet.getRange('T2').setFormula('=IF(OR(P2="", S2=""), "", P2-S2)');

    // U列: 予測の信頼度
    sheet.getRange('U2').setFormula('=IF(P2="", "", IF(ABS(P2-IFERROR(QUERY(Engagement_Log!A:U, "SELECT F WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), 0)) < 10, "高", IF(ABS(P2-IFERROR(QUERY(Engagement_Log!A:U, "SELECT F WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), 0)) < 20, "中", "低")))');

    // V列: 志望度スコア
    sheet.getRange('V2').setFormula('=IFERROR(QUERY(Survey_Response!A:H, "SELECT E WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0)*10, 0)');

    // W列: 競合優位性スコア
    sheet.getRange('W2').setFormula('=IFERROR(QUERY(Engagement_Log!A:U, "SELECT Q WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), 0)');

    // X列: 懸念解消度スコア
    sheet.getRange('X2').setFormula('=IFERROR(QUERY(Engagement_Log!A:U, "SELECT L WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), 0)');

    // Y列: コアモチベーション
    sheet.getRange('Y2').setFormula('=IFERROR(QUERY(Engagement_Log!A:U, "SELECT M WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), "")');

    // Z列: 主要懸念事項
    sheet.getRange('Z2').setFormula('=IFERROR(QUERY(Engagement_Log!A:U, "SELECT N WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), "")');

    // AA列: 競合企業1
    sheet.getRange('AA2').setFormula('=IFERROR(INDEX(SPLIT(QUERY(Engagement_Log!A:U, "SELECT P WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), "、"), 1), "")');

    // AB列: 競合企業2
    sheet.getRange('AB2').setFormula('=IFERROR(INDEX(SPLIT(QUERY(Engagement_Log!A:U, "SELECT P WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), "、"), 2), "")');

    // AC列: 競合企業3
    sheet.getRange('AC2').setFormula('=IFERROR(INDEX(SPLIT(QUERY(Engagement_Log!A:U, "SELECT P WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), "、"), 3), "")');

    // AD列: 自社の志望順位
    sheet.getRange('AD2').setFormula('=IFERROR(QUERY(Survey_Response!A:H, "SELECT G WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), "不明")');

    // AE列: 前回接点日
    sheet.getRange('AE2').setFormula('=IFERROR(QUERY(Contact_History!A:H, "SELECT D WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), "")');

    // AF列: 前回接点からの空白期間
    sheet.getRange('AF2').setFormula('=IF(AE2="", "", TODAY()-AE2)');

    // AG列: 接点回数
    sheet.getRange('AG2').setFormula('=COUNTIF(Contact_History!B:B, A2)');

    // AH列: 平均接点間隔
    sheet.getRange('AH2').setFormula('=IF(AG2<=1, "", ROUND((TODAY()-G2)/AG2, 1))');

    // AJ列: 次推奨アクション
    sheet.getRange('AJ2').setFormula('=IFERROR(QUERY(Engagement_Log!A:U, "SELECT R WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), "")');

    // AK列: アクション期限
    sheet.getRange('AK2').setFormula('=IFERROR(QUERY(Engagement_Log!A:U, "SELECT S WHERE B=\'"&A2&"\' ORDER BY D DESC LIMIT 1", 0), "")');

    // AL列: アクション緊急度
    sheet.getRange('AL2').setFormula('=IF(C2="最終面接", "CRITICAL", IF(C2="2次面接", IF(AF2>=7, "HIGH", "NORMAL"), IF(C2="1次面接", IF(AF2>=14, "HIGH", "NORMAL"), "LOW")))');

    // AO列: 注力度合いスコア
    sheet.getRange('AO2').setFormula('=IF(OR(H2="", R2=""), "", H2/100 * R2/100 * AP2)');

    // AP列: 緊急度係数
    sheet.getRange('AP2').setFormula('=IF(C2="最終面接", 3.0, IF(C2="2次面接", 2.0, IF(C2="1次面接", 1.5, 1.0))) + IF(AF2>=14, 1.0, IF(AF2>=7, 0.5, 0))');

    Logger.log('✅ Candidates_Masterシートのセットアップが完了しました');
    Logger.log('✅ Candidates_Masterの自動計算列を設定しました');
    
  } catch (error) {
    logError('setupCandidatesMaster', error);
    throw error;
  }
}

/**
 * Evaluation_Logシートを設定
 */
function setupEvaluationLog() {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.EVALUATION_LOG);
    
    // ヘッダー行の定義（T列まで追加）
    const headers = [
      'log_id', 'candidate_id', '氏名', 'evaluation_date', '選考フェーズ', '面接官',
      'overall_score', 'philosophy_score', 'philosophy_reason',
      'strategy_score', 'strategy_reason', 'motivation_score', 'motivation_reason',
      'execution_score', 'execution_reason', 'highest_risk_level',
      'pass_probability', 'recommendation', 'doc_url',
      'proactivity_score' // 【新規追加】T列（20列目）
    ];
    
    setHeaders(sheet, headers);
    
    // 列幅の詳細設定
    const columnWidths = {
      1: 100,   // log_id
      2: 100,   // candidate_id
      3: 120,   // 氏名
      4: 160,   // evaluation_date
      5: 120,   // 選考フェーズ
      6: 120,   // 面接官
      7: 120,   // overall_score
      8: 130,   // philosophy_score
      9: 300,   // philosophy_reason
      10: 130,  // strategy_score
      11: 300,  // strategy_reason
      12: 130,  // motivation_score
      13: 300,  // motivation_reason
      14: 130,  // execution_score
      15: 300,  // execution_reason
      16: 150,  // highest_risk_level
      17: 140,  // pass_probability
      18: 140,  // recommendation
      19: 250,  // doc_url
      20: 130   // proactivity_score（新規）
    };
    
    for (let col in columnWidths) {
      sheet.setColumnWidth(parseInt(col), columnWidths[col]);
    }
    
    // フォーマット設定
    sheet.getRange('D2:D1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange('G2:G1000').setNumberFormat('0.00');
    sheet.getRange('H2:H1000').setNumberFormat('0.00');
    sheet.getRange('J2:J1000').setNumberFormat('0.00');
    sheet.getRange('L2:L1000').setNumberFormat('0.00');
    sheet.getRange('N2:N1000').setNumberFormat('0.00');
    sheet.getRange('Q2:Q1000').setNumberFormat('0.00"%"');
    sheet.getRange('T2:T1000').setNumberFormat('0'); // proactivity_score（0-100点）
    
    // テキストの折り返し
    sheet.getRange('I2:I1000').setWrap(true);
    sheet.getRange('K2:K1000').setWrap(true);
    sheet.getRange('M2:M1000').setWrap(true);
    sheet.getRange('O2:O1000').setWrap(true);
    
    Logger.log('✅ Evaluation_Logシートのセットアップが完了しました');
    
  } catch (error) {
    logError('setupEvaluationLog', error);
    throw error;
  }
}

/**
 * Engagement_Logシートを設定
 */
function setupEngagementLog() {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.ENGAGEMENT_LOG);
    
    const headers = [
      'log_id', 'candidate_id', '氏名', 'timestamp', 'contact_type',
      'acceptance_rate_rule', 'acceptance_rate_ai', 'acceptance_rate_final',
      'confidence_level', 'motivation_score', 'competitive_advantage_score',
      'concern_resolution_score', 'core_motivation', 'top_concern',
      'concern_category', 'competitors', 'competitive_advantage',
      'next_action', 'action_deadline', 'action_priority', 'doc_url'
    ];
    
    setHeaders(sheet, headers);
    
    // 列幅の詳細設定
    const columnWidths = {
      1: 100,   // log_id
      2: 100,   // candidate_id
      3: 120,   // 氏名
      4: 160,   // timestamp
      5: 120,   // contact_type
      6: 180,   // acceptance_rate_rule
      7: 160,   // acceptance_rate_ai
      8: 180,   // acceptance_rate_final
      9: 120,   // confidence_level
      10: 150,  // motivation_score
      11: 180,  // competitive_advantage_score
      12: 180,  // concern_resolution_score
      13: 250,  // core_motivation
      14: 250,  // top_concern
      15: 150,  // concern_category
      16: 250,  // competitors
      17: 180,  // competitive_advantage
      18: 300,  // next_action
      19: 140,  // action_deadline
      20: 140,  // action_priority
      21: 250   // doc_url
    };
    
    for (let col in columnWidths) {
      sheet.setColumnWidth(parseInt(col), columnWidths[col]);
    }
    
    // フォーマット設定
    sheet.getRange('D2:D1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange('F2:H1000').setNumberFormat('0.00"%"');
    sheet.getRange('J2:L1000').setNumberFormat('0');
    sheet.getRange('Q2:Q1000').setNumberFormat('0');
    sheet.getRange('S2:S1000').setNumberFormat('yyyy-mm-dd');
    
    // テキストの折り返し
    sheet.getRange('M2:N1000').setWrap(true);
    sheet.getRange('P2:P1000').setWrap(true);
    sheet.getRange('R2:R1000').setWrap(true);
    
    Logger.log('✅ Engagement_Logシートのセットアップが完了しました');
    
  } catch (error) {
    logError('setupEngagementLog', error);
    throw error;
  }
}

/**
 * Company_Assetsシートを設定
 */
function setupCompanyAssets() {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.COMPANY_ASSETS);
    
    const headers = [
      'asset_id', 'カテゴリ', 'アセット名', '詳細説明',
      '訴求ポイント', '推奨アクション', '関連資料URL', '最終更新日時'
    ];
    
    setHeaders(sheet, headers);
    
    // 列幅の詳細設定
    const columnWidths = {
      1: 100,   // asset_id
      2: 120,   // カテゴリ
      3: 200,   // アセット名
      4: 350,   // 詳細説明
      5: 350,   // 訴求ポイント
      6: 300,   // 推奨アクション
      7: 250,   // 関連資料URL
      8: 160    // 最終更新日時
    };
    
    for (let col in columnWidths) {
      sheet.setColumnWidth(parseInt(col), columnWidths[col]);
    }
    
    sheet.getRange('H2:H1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    
    // テキストの折り返し
    sheet.getRange('D2:F1000').setWrap(true);
    
    Logger.log('✅ Company_Assetsシートのセットアップが完了しました');
    
  } catch (error) {
    logError('setupCompanyAssets', error);
    throw error;
  }
}

/**
 * Competitor_DBシートを設定
 */
function setupCompetitorDB() {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.COMPETITOR_DB);
    
    const headers = [
      'competitor_id', '競合企業名', '業界', '企業規模',
      '強み1', '強み2', '強み3', '弱み1', '弱み2', '弱み3',
      '成長環境_詳細', '給与水準_詳細', '働き方_詳細', '社風_詳細',
      '自社との比較', '情報ソース', '最終更新日時'
    ];
    
    setHeaders(sheet, headers);
    
    // 列幅の詳細設定
    const columnWidths = {
      1: 120,   // competitor_id
      2: 180,   // 競合企業名
      3: 120,   // 業界
      4: 120,   // 企業規模
      5: 200,   // 強み1
      6: 200,   // 強み2
      7: 200,   // 強み3
      8: 200,   // 弱み1
      9: 200,   // 弱み2
      10: 200,  // 弱み3
      11: 300,  // 成長環境_詳細
      12: 300,  // 給与水準_詳細
      13: 300,  // 働き方_詳細
      14: 300,  // 社風_詳細
      15: 350,  // 自社との比較
      16: 250,  // 情報ソース
      17: 160   // 最終更新日時
    };
    
    for (let col in columnWidths) {
      sheet.setColumnWidth(parseInt(col), columnWidths[col]);
    }
    
    sheet.getRange('Q2:Q1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    
    // テキストの折り返し
    sheet.getRange('E2:O1000').setWrap(true);
    
    Logger.log('✅ Competitor_DBシートのセットアップが完了しました');
    
  } catch (error) {
    logError('setupCompetitorDB', error);
    throw error;
  }
}

/**
 * その他のシートを設定
 */
function setupOtherSheets() {
  try {
    // Evidence
    setupEvidenceSheet();
    
    // Risk
    setupRiskSheet();
    
    // NextQ
    setupNextQSheet();
    
    // Contact_History
    setupContactHistorySheet();
    
    // Survey_Response
    setupSurveyResponseSheet();
    
    // Acceptance_Story
    setupAcceptanceStorySheet();
    
    // Competitor_Comparison
    setupCompetitorComparisonSheet();
    
    // Archive
    setupArchiveSheet();

    // Survey_Send_Log（新規追加）
    setupSurveySendLog();

    // README（システム説明）
    setupReadmeSheet();

    Logger.log('✅ その他のシートのセットアップが完了しました');
    
  } catch (error) {
    logError('setupOtherSheets', error);
    throw error;
  }
}

/**
 * Evidenceシートを設定
 */
function setupEvidenceSheet() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.EVIDENCE);
  
  const headers = [
    'evidence_id', 'candidate_id', '氏名', 'evaluation_log_id',
    '評価軸', 'エビデンステキスト', '記録日時'
  ];
  
  setHeaders(sheet, headers);
  
  const columnWidths = {
    1: 120,   // evidence_id
    2: 100,   // candidate_id
    3: 120,   // 氏名
    4: 150,   // evaluation_log_id
    5: 150,   // 評価軸
    6: 500,   // エビデンステキスト
    7: 160    // 記録日時
  };
  
  for (let col in columnWidths) {
    sheet.setColumnWidth(parseInt(col), columnWidths[col]);
  }
  
  sheet.getRange('G2:G1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('F2:F1000').setWrap(true);
}

/**
 * Riskシートを設定
 */
function setupRiskSheet() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.RISK);
  
  const headers = [
    'risk_id', 'candidate_id', '氏名', 'evaluation_log_id',
    'リスクレベル', 'リスク内容', '対策', '記録日時'
  ];
  
  setHeaders(sheet, headers);
  
  const columnWidths = {
    1: 120,   // risk_id
    2: 100,   // candidate_id
    3: 120,   // 氏名
    4: 150,   // evaluation_log_id
    5: 120,   // リスクレベル
    6: 350,   // リスク内容
    7: 350,   // 対策
    8: 160    // 記録日時
  };
  
  for (let col in columnWidths) {
    sheet.setColumnWidth(parseInt(col), columnWidths[col]);
  }
  
  sheet.getRange('H2:H1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('F2:G1000').setWrap(true);
}

/**
 * NextQシートを設定
 */
function setupNextQSheet() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.NEXT_Q);
  
  const headers = [
    'question_id', 'candidate_id', '氏名', 'evaluation_log_id',
    '質問カテゴリ', '質問内容', '質問の意図', '記録日時'
  ];
  
  setHeaders(sheet, headers);
  
  const columnWidths = {
    1: 120,   // question_id
    2: 100,   // candidate_id
    3: 120,   // 氏名
    4: 150,   // evaluation_log_id
    5: 150,   // 質問カテゴリ
    6: 400,   // 質問内容
    7: 350,   // 質問の意図
    8: 160    // 記録日時
  };
  
  for (let col in columnWidths) {
    sheet.setColumnWidth(parseInt(col), columnWidths[col]);
  }
  
  sheet.getRange('H2:H1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('F2:G1000').setWrap(true);
}

/**
 * Contact_Historyシートを設定
 */
function setupContactHistorySheet() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.CONTACT_HISTORY);
  
  const headers = [
    'contact_id', 'candidate_id', '氏名', '接点日時',
    '接点タイプ', '担当者', '内容', '次回予定'
  ];
  
  setHeaders(sheet, headers);
  
  const columnWidths = {
    1: 120,   // contact_id
    2: 100,   // candidate_id
    3: 120,   // 氏名
    4: 160,   // 接点日時
    5: 120,   // 接点タイプ
    6: 120,   // 担当者
    7: 400,   // 内容
    8: 160    // 次回予定
  };
  
  for (let col in columnWidths) {
    sheet.setColumnWidth(parseInt(col), columnWidths[col]);
  }
  
  sheet.getRange('D2:D1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('H2:H1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('G2:G1000').setWrap(true);
}

/**
 * Survey_Responseシートを設定
 */
function setupSurveyResponseSheet() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.SURVEY_RESPONSE);
  
  const headers = [
    'response_id', 'candidate_id', '氏名', '回答日時',
    '志望度（1-10）', '懸念事項', '他社選考状況', 'その他コメント'
  ];
  
  setHeaders(sheet, headers);
  
  const columnWidths = {
    1: 120,   // response_id
    2: 100,   // candidate_id
    3: 120,   // 氏名
    4: 160,   // 回答日時
    5: 120,   // 志望度（1-10）
    6: 300,   // 懸念事項
    7: 300,   // 他社選考状況
    8: 300    // その他コメント
  };
  
  for (let col in columnWidths) {
    sheet.setColumnWidth(parseInt(col), columnWidths[col]);
  }
  
  sheet.getRange('D2:D1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('F2:H1000').setWrap(true);
}

/**
 * Acceptance_Storyシートを設定
 */
function setupAcceptanceStorySheet() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.ACCEPTANCE_STORY);
  
  const headers = [
    'candidate_id', '氏名', '現状分析', '承諾までのシナリオ',
    'ステップ1_アクション', 'ステップ1_期待効果', 'ステップ1_エビデンス',
    'ステップ2_アクション', 'ステップ2_期待効果', 'ステップ2_エビデンス',
    'ステップ3_アクション', 'ステップ3_期待効果', 'ステップ3_エビデンス',
    'リスクシナリオ', '推奨アクションの優先順位', '生成日時', '最終更新日時'
  ];
  
  setHeaders(sheet, headers);
  
  const columnWidths = {
    1: 100,   // candidate_id
    2: 120,   // 氏名
    3: 400,   // 現状分析
    4: 400,   // 承諾までのシナリオ
    5: 300,   // ステップ1_アクション
    6: 300,   // ステップ1_期待効果
    7: 300,   // ステップ1_エビデンス
    8: 300,   // ステップ2_アクション
    9: 300,   // ステップ2_期待効果
    10: 300,  // ステップ2_エビデンス
    11: 300,  // ステップ3_アクション
    12: 300,  // ステップ3_期待効果
    13: 300,  // ステップ3_エビデンス
    14: 400,  // リスクシナリオ
    15: 400,  // 推奨アクションの優先順位
    16: 160,  // 生成日時
    17: 160   // 最終更新日時
  };
  
  for (let col in columnWidths) {
    sheet.setColumnWidth(parseInt(col), columnWidths[col]);
  }
  
  sheet.getRange('P2:Q1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('C2:O1000').setWrap(true);
}

/**
 * Competitor_Comparisonシートを設定
 */
function setupCompetitorComparisonSheet() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.COMPETITOR_COMPARISON);
  
  const headers = [
    'comparison_id', 'candidate_id', '候補者名', '競合企業名',
    'ブランド力_自社', 'ブランド力_競合', 'ブランド力_優位性',
    '給与水準_自社', '給与水準_競合', '給与水準_優位性',
    '成長環境_自社', '成長環境_競合', '成長環境_優位性',
    '裁量権_自社', '裁量権_競合', '裁量権_優位性',
    '残業時間_自社', '残業時間_競合', '残業時間_優位性',
    '社風_自社', '社風_競合', '社風_優位性',
    '事業安定性_自社', '事業安定性_競合', '事業安定性_優位性',
    '総合的な優位性', '推奨する訴求ポイント', '生成日時', '最終更新日時'
  ];
  
  setHeaders(sheet, headers);
  
  const columnWidths = {
    1: 150,   // comparison_id
    2: 100,   // candidate_id
    3: 120,   // 候補者名
    4: 150,   // 競合企業名
    5: 200,   // ブランド力_自社
    6: 200,   // ブランド力_競合
    7: 120,   // ブランド力_優位性
    8: 200,   // 給与水準_自社
    9: 200,   // 給与水準_競合
    10: 120,  // 給与水準_優位性
    11: 200,  // 成長環境_自社
    12: 200,  // 成長環境_競合
    13: 120,  // 成長環境_優位性
    14: 200,  // 裁量権_自社
    15: 200,  // 裁量権_競合
    16: 120,  // 裁量権_優位性
    17: 200,  // 残業時間_自社
    18: 200,  // 残業時間_競合
    19: 120,  // 残業時間_優位性
    20: 200,  // 社風_自社
    21: 200,  // 社風_競合
    22: 120,  // 社風_優位性
    23: 200,  // 事業安定性_自社
    24: 200,  // 事業安定性_競合
    25: 120,  // 事業安定性_優位性
    26: 150,  // 総合的な優位性
    27: 400,  // 推奨する訴求ポイント
    28: 160,  // 生成日時
    29: 160   // 最終更新日時
  };
  
  for (let col in columnWidths) {
    sheet.setColumnWidth(parseInt(col), columnWidths[col]);
  }
  
  sheet.getRange('AB2:AC1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('E2:Z1000').setWrap(true);
}

/**
 * Archiveシートを設定
 */
function setupArchiveSheet() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.ARCHIVE);
  
  const headers = [
    'candidate_id', '氏名', '承諾日', '入社日', '配属部署',
    '最終_合格可能性', '最終_承諾可能性', '承諾ストーリー', 'アーカイブ日時'
  ];
  
  setHeaders(sheet, headers);
  
  const columnWidths = {
    1: 100,   // candidate_id
    2: 120,   // 氏名
    3: 120,   // 承諾日
    4: 120,   // 入社日
    5: 150,   // 配属部署
    6: 150,   // 最終_合格可能性
    7: 150,   // 最終_承諾可能性
    8: 500,   // 承諾ストーリー
    9: 160    // アーカイブ日時
  };
  
  for (let col in columnWidths) {
    sheet.setColumnWidth(parseInt(col), columnWidths[col]);
  }
  
  sheet.getRange('C2:D1000').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('F2:G1000').setNumberFormat('0.00"%"');
  sheet.getRange('I2:I1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange('H2:H1000').setWrap(true);
}

/**
 * READMEシート（システム説明）を設定
 */
function setupReadmeSheet() {
  const sheet = getOrCreateSheet('README');
  
  // タイトル
  sheet.getRange('A1').setValue('📖 採用参謀AI：システム説明');
  sheet.getRange('A1').setFontSize(20).setFontWeight('bold').setFontColor('#1a73e8');
  sheet.getRange('A1:F1').merge();
  
  // 概要
  sheet.getRange('A3').setValue('【概要】');
  sheet.getRange('A3').setFontSize(14).setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange('A4').setValue(
    'このスプレッドシートは、採用活動における「合格可能性」と「承諾可能性」を予測し、\n' +
    '「承諾ストーリー」を自動設計する統合型システムです。'
  );
  sheet.getRange('A4').setWrap(true);
  
  // データフロー
  sheet.getRange('A7').setValue('【データフロー】');
  sheet.getRange('A7').setFontSize(14).setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange('A8').setValue(
    '1. 面接実施 → Evaluation_Logにデータ入力\n' +
    '2. Difyで評価データ生成（4軸評価、リスク分析など）\n' +
    '3. Engagement_Logに承諾可能性データ入力\n' +
    '4. Candidates_Masterに最新データが集約\n' +
    '5. Dashboardで可視化'
  );
  sheet.getRange('A8').setWrap(true);
  
  // 各シートの役割
  sheet.getRange('A12').setValue('【各シートの役割】');
  sheet.getRange('A12').setFontSize(14).setFontWeight('bold').setBackground('#f3f3f3');
  
  const sheetDescriptions = [
    ['シート名', '役割', '更新頻度'],
    ['Dashboard', '候補者一覧、今週のアクション、KPIを表示', '自動更新'],
    ['Candidates_Master', '候補者の最新状態を1行で管理（マスターデータ）', '自動更新'],
    ['Evaluation_Log', '選考フェーズごとの評価履歴を記録', '面接後に手動入力'],
    ['Engagement_Log', '接点ごとの承諾可能性履歴を記録', 'Difyが自動生成'],
    ['Evidence', '評価の根拠となる発言を記録', 'Difyが自動生成'],
    ['Risk', 'リスク要因を記録', 'Difyが自動生成'],
    ['NextQ', '次回面接で聞くべき質問を記録', 'Difyが自動生成'],
    ['Contact_History', '接点履歴を記録', '手動入力'],
    ['Survey_Response', 'アンケート回答を記録', '候補者が回答'],
    ['Acceptance_Story', '承諾ストーリーを記録', 'Difyが自動生成'],
    ['Competitor_Comparison', '競合企業との比較表を記録', 'Difyが自動生成'],
    ['Company_Assets', '企業の強みをデータベース化', '初期設定済み'],
    ['Competitor_DB', '競合企業の情報をデータベース化', '初期設定済み'],
    ['Archive', '承諾した候補者のデータをアーカイブ', '承諾後に手動移動']
  ];
  
  sheet.getRange(13, 1, sheetDescriptions.length, sheetDescriptions[0].length)
    .setValues(sheetDescriptions);
  
  // ヘッダー行の書式設定
  sheet.getRange('A13:C13')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // 列幅設定
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 500);
  sheet.setColumnWidth(3, 150);
  
  // 罫線を追加
  sheet.getRange(13, 1, sheetDescriptions.length, sheetDescriptions[0].length)
    .setBorder(true, true, true, true, true, true);
  
  // 使い方
  sheet.getRange('A30').setValue('【使い方】');
  sheet.getRange('A30').setFontSize(14).setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange('A31').setValue(
    '1. 面接実施後、Evaluation_Logに評価データを入力\n' +
    '2. Difyが自動でデータを解析（Engagement_Log、Evidence、Riskなどに出力）\n' +
    '3. Candidates_Masterで候補者の最新状態を確認\n' +
    '4. Dashboardで今週のアクションを確認\n' +
    '5. 推奨アクションを実行'
  );
  sheet.getRange('A31').setWrap(true);
  
  // 注意事項
  sheet.getRange('A36').setValue('【注意事項】');
  sheet.getRange('A36').setFontSize(14).setFontWeight('bold').setBackground('#fff2cc');
  sheet.getRange('A37').setValue(
    '⚠️ Candidates_Masterの列は編集しないでください（関数が壊れます）\n' +
    '⚠️ Dashboardは自動生成されています。QUERY関数を削除しないでください\n' +
    '⚠️ Difyとの連携設定は別途必要です'
  );
  sheet.getRange('A37').setWrap(true).setFontColor('#d93025');
  
  Logger.log('✅ READMEシートを作成しました');
}

/**
 * Dashboard（完全版）を設定
 */
function setupDashboardComplete() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.DASHBOARD);
  
  // シートをクリア
  sheet.clear();
  
  // タイトル
  sheet.getRange('A1').setValue('📊 採用参謀AI - Dashboard');
  sheet.getRange('A1:L1').merge();
  sheet.getRange('A1')
    .setFontSize(20)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1a73e8')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 50);
  
  // 最終更新日時
  sheet.getRange('A2').setValue('最終更新: ' + formatDateTime(new Date()));
  sheet.getRange('A2:L2').merge();
  sheet.getRange('A2')
    .setFontSize(10)
    .setFontColor('#666666')
    .setHorizontalAlignment('right');
  
  // セクション1: 今週のアクション
  let currentRow = 4;
  
  sheet.getRange(`A${currentRow}`).setValue('【今週のアクション】');
  sheet.getRange(`A${currentRow}:L${currentRow}`).merge();
  sheet.getRange(`A${currentRow}`)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('left');
  
  currentRow++;
  
  // ヘッダー行
  const actionHeaders = ['緊急度', '候補者名', '現在ステータス', '承諾可能性', '次推奨アクション', 'アクション期限', '実行者'];
  sheet.getRange(`A${currentRow}:G${currentRow}`).setValues([actionHeaders]);
  sheet.getRange(`A${currentRow}:G${currentRow}`)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
  
  currentRow++;
  
  // QUERY関数でデータを抽出（修正版：ORDER BY を期限順のみに簡略化）
  const actionQuery = `=QUERY(Candidates_Master!A:AN,
    "SELECT AL, B, C, R, AJ, AK, AN
     WHERE C != '辞退' AND C != '見送り' AND C != '承諾'
     AND (AL = 'CRITICAL' OR AL = 'HIGH')
     ORDER BY AK ASC
     LIMIT 10", 0)`;
  
  sheet.getRange(`A${currentRow}`).setFormula(actionQuery);
  
  // データ範囲のフォーマット設定
  sheet.getRange(`A${currentRow}:G${currentRow + 9}`)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');
  
  sheet.getRange(`D${currentRow}:D${currentRow + 9}`).setNumberFormat('0.00"%"');
  sheet.getRange(`F${currentRow}:F${currentRow + 9}`).setNumberFormat('yyyy-mm-dd');
  
  currentRow += 12;
  
  // セクション2: 候補者一覧
  sheet.getRange(`A${currentRow}`).setValue('【候補者一覧】（承諾可能性順）');
  sheet.getRange(`A${currentRow}:L${currentRow}`).merge();
  sheet.getRange(`A${currentRow}`)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('left');
  
  currentRow++;
  
  // ヘッダー行
  const candidateHeaders = ['順位', '候補者名', '現在ステータス', '合格可能性', '承諾可能性', '注力度合い', '競合企業', '次推奨アクション'];
  sheet.getRange(`A${currentRow}:H${currentRow}`).setValues([candidateHeaders]);
  sheet.getRange(`A${currentRow}:H${currentRow}`)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
  
  currentRow++;
  
  // 順位列を追加
  const rankStartRow = currentRow;
  for (let i = 1; i <= 20; i++) {
    sheet.getRange(`A${rankStartRow + i - 1}`).setValue(i);
  }
  
  // QUERY関数でデータを抽出
  const candidateQuery = `=QUERY(Candidates_Master!A:AO, 
    "SELECT B, C, H, R, AO, AA, AJ 
     WHERE C != '辞退' AND C != '見送り' AND C != '承諾' 
     ORDER BY R DESC 
     LIMIT 20", 0)`;
  
  sheet.getRange(`B${currentRow}`).setFormula(candidateQuery);
  
  // データ範囲のフォーマット設定
  sheet.getRange(`A${currentRow}:H${currentRow + 19}`)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');
  
  sheet.getRange(`D${currentRow}:E${currentRow + 19}`).setNumberFormat('0.00"%"');
  sheet.getRange(`F${currentRow}:F${currentRow + 19}`).setNumberFormat('0.00');
  
  currentRow += 22;
  
  // セクション3: 今週のKPI
  sheet.getRange(`A${currentRow}`).setValue('【今週のKPI】');
  sheet.getRange(`A${currentRow}:L${currentRow}`).merge();
  sheet.getRange(`A${currentRow}`)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('left');
  
  currentRow++;
  
  // KPIテーブル
  const kpiData = [
    ['指標', '値'],
    ['選考中の候補者数', `=COUNTIF(Candidates_Master!C:C, "面談") + COUNTIF(Candidates_Master!C:C, "1次面接") + COUNTIF(Candidates_Master!C:C, "2次面接") + COUNTIF(Candidates_Master!C:C, "最終面接")`],
    ['平均合格可能性', `=AVERAGE(FILTER(Candidates_Master!H:H, Candidates_Master!C:C<>"辞退", Candidates_Master!C:C<>"見送り", Candidates_Master!C:C<>"承諾", Candidates_Master!H:H<>""))`],
    ['平均承諾可能性', `=AVERAGE(FILTER(Candidates_Master!R:R, Candidates_Master!C:C<>"辞退", Candidates_Master!C:C<>"見送り", Candidates_Master!C:C<>"承諾", Candidates_Master!R:R<>""))`],
    ['CRITICALアクション数', `=COUNTIF(Candidates_Master!AL:AL, "CRITICAL")`],
    ['HIGHアクション数', `=COUNTIF(Candidates_Master!AL:AL, "HIGH")`],
    ['アンケート回答率', `=IF(COUNTA(Survey_Response!A:A)-1>0, COUNTA(Survey_Response!A:A)-1 & "/" & COUNTA(Candidates_Master!A:A)-1 & " (" & TEXT((COUNTA(Survey_Response!A:A)-1)/(COUNTA(Candidates_Master!A:A)-1), "0%") & ")", "N/A")`]
  ];
  
  sheet.getRange(`A${currentRow}:B${currentRow + kpiData.length - 1}`).setValues(kpiData);
  
  // ヘッダー行のフォーマット
  sheet.getRange(`A${currentRow}:B${currentRow}`)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // データ行のフォーマット
  sheet.getRange(`A${currentRow}:B${currentRow + kpiData.length - 1}`)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');
  
  // 数値列のフォーマット
  sheet.getRange(`B${currentRow + 2}:B${currentRow + 3}`).setNumberFormat('0.00"%"');
  
  currentRow += kpiData.length + 2;
  
  // セクション4: 採用ファネル
  sheet.getRange(`A${currentRow}`).setValue('【採用ファネル】');
  sheet.getRange(`A${currentRow}:L${currentRow}`).merge();
  sheet.getRange(`A${currentRow}`)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('left');
  
  currentRow++;
  
  // ファネルテーブル
  const funnelData = [
    ['選考ステージ', '人数', '通過率'],
    ['面談', `=COUNTIF(Candidates_Master!C:C, "面談")`, '-'],
    ['1次面接', `=COUNTIF(Candidates_Master!C:C, "1次面接")`, `=IF(B${currentRow + 1}>0, TEXT(B${currentRow + 2}/B${currentRow + 1}, "0.0%"), "-")`],
    ['2次面接', `=COUNTIF(Candidates_Master!C:C, "2次面接")`, `=IF(B${currentRow + 2}>0, TEXT(B${currentRow + 3}/B${currentRow + 2}, "0.0%"), "-")`],
    ['最終面接', `=COUNTIF(Candidates_Master!C:C, "最終面接")`, `=IF(B${currentRow + 3}>0, TEXT(B${currentRow + 4}/B${currentRow + 3}, "0.0%"), "-")`],
    ['内定通知済', `=COUNTIF(Candidates_Master!C:C, "内定通知済")`, `=IF(B${currentRow + 4}>0, TEXT(B${currentRow + 5}/B${currentRow + 4}, "0.0%"), "-")`],
    ['承諾', `=COUNTIF(Candidates_Master!C:C, "承諾")`, `=IF(B${currentRow + 5}>0, TEXT(B${currentRow + 6}/B${currentRow + 5}, "0.0%"), "-")`]
  ];
  
  sheet.getRange(`A${currentRow}:C${currentRow + funnelData.length - 1}`).setValues(funnelData);
  
  // ヘッダー行のフォーマット
  sheet.getRange(`A${currentRow}:C${currentRow}`)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // データ行のフォーマット
  sheet.getRange(`A${currentRow}:C${currentRow + funnelData.length - 1}`)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle');
  
  // 人数列を中央揃え
  sheet.getRange(`B${currentRow + 1}:C${currentRow + funnelData.length - 1}`)
    .setHorizontalAlignment('center');
  
  // 列幅の設定
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 350);
  
  // 条件付き書式の設定
  const rules = [];
  
  // 今週のアクション: 緊急度の色分け
  const actionUrgencyRange = sheet.getRange('A6:A15');
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('CRITICAL')
      .setBackground(CONFIG.COLORS.CRITICAL)
      .setFontColor('#ffffff')
      .setBold(true)
      .setRanges([actionUrgencyRange])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('HIGH')
      .setBackground(CONFIG.COLORS.HIGH)
      .setRanges([actionUrgencyRange])
      .build()
  );
  
  // 候補者一覧: 承諾可能性の色分け
  const acceptanceRateRange = sheet.getRange('E19:E38');
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(70)
      .setBackground(CONFIG.COLORS.HIGH_SCORE)
      .setRanges([acceptanceRateRange])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(50, 69)
      .setBackground(CONFIG.COLORS.MEDIUM_SCORE)
      .setRanges([acceptanceRateRange])
      .build()
  );
  
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(50)
      .setBackground(CONFIG.COLORS.LOW_SCORE)
      .setRanges([acceptanceRateRange])
      .build()
  );
  
  sheet.setConditionalFormatRules(rules);

  Logger.log('✅ Dashboard（完全版）を作成しました');
}

/**
 * Survey_Send_Logシートを作成
 */
function setupSurveySendLog() {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);

    // ヘッダー行の定義
    const headers = [
      'send_id',
      'candidate_id',
      'candidate_name',
      'email',
      'phase',
      'send_time',
      'send_status',
      'error_message'
    ];

    setHeaders(sheet, headers);

    // 列幅を調整
    sheet.setColumnWidth(1, 150);  // send_id
    sheet.setColumnWidth(2, 150);  // candidate_id
    sheet.setColumnWidth(3, 120);  // candidate_name
    sheet.setColumnWidth(4, 200);  // email
    sheet.setColumnWidth(5, 100);  // phase
    sheet.setColumnWidth(6, 160);  // send_time
    sheet.setColumnWidth(7, 100);  // send_status
    sheet.setColumnWidth(8, 250);  // error_message

    // フォーマット設定
    sheet.getRange('F2:F1000').setNumberFormat('yyyy-mm-dd hh:mm:ss'); // send_time

    // データ検証（phase列）
    setDropdownValidation(
      sheet,
      'E2:E1000',
      CONFIG.VALIDATION_OPTIONS.SURVEY_PHASE,
      'アンケート種別を選択してください'
    );

    // データ検証（send_status列）
    setDropdownValidation(
      sheet,
      'G2:G1000',
      CONFIG.VALIDATION_OPTIONS.SEND_STATUS,
      '送信ステータスを選択してください'
    );

    // 条件付き書式（send_status列）
    const rules = [];

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('成功')
        .setBackground(CONFIG.COLORS.HIGH_SCORE)
        .setRanges([sheet.getRange('G2:G1000')])
        .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('失敗')
        .setBackground(CONFIG.COLORS.LOW_SCORE)
        .setRanges([sheet.getRange('G2:G1000')])
        .build()
    );

    sheet.setConditionalFormatRules(rules);

    Logger.log('✅ Survey_Send_Logシート作成完了');

  } catch (error) {
    logError('setupSurveySendLog', error);
    throw error;
  }
}

