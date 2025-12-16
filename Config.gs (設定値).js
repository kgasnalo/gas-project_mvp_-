/**
 * 設定値を管理
 */
const CONFIG = {
  // シート名の定義
  SHEET_NAMES: {
    DASHBOARD: 'Dashboard',
    DASHBOARD_DATA: 'Dashboard_Data', // Phase 4-1: ダッシュボード中間データ
    CANDIDATES_MASTER: 'Candidates_Master',
    COMPANY_ASSETS: 'Company_Assets',
    COMPETITOR_DB: 'Competitor_DB',
    EVALUATION_LOG: 'Evaluation_Log',
    EVALUATION_MASTER: 'Evaluation_Master', // Phase 4-2: 評価マスター
    ENGAGEMENT_LOG: 'Engagement_Log',
    EVIDENCE: 'Evidence',
    RISK: 'Risk',
    NEXT_Q: 'NextQ',
    CONTACT_HISTORY: 'Contact_History',
    SURVEY_RESPONSE: 'Survey_Response',
    ACCEPTANCE_STORY: 'Acceptance_Story',
    COMPETITOR_COMPARISON: 'Competitor_Comparison',
    ARCHIVE: 'Archive',
    SURVEY_SEND_LOG: 'Survey_Send_Log', // 新規追加
    SURVEY_ANALYSIS: 'Survey_Analysis', // Phase 2 Step 2で追加
    ANKETO_URL: 'アンケートURL' // 新規追加
  },
  
  // 色の定義
  COLORS: {
    HEADER_BG: '#4285f4',
    HEADER_TEXT: '#ffffff',
    HIGH_SCORE: '#d9ead3',
    MEDIUM_SCORE: '#fff2cc',
    LOW_SCORE: '#f4cccc',
    CRITICAL: '#ea4335',
    HIGH: '#fbbc04',
    NORMAL: '#34a853',
    LOW: '#4285f4'
  },
  
  // データ検証の選択肢
  VALIDATION_OPTIONS: {
    STATUS: ['応募', '初回面談', '適性検査', '1次面接', '社員面談', '2次面接', '最終面接', '内定通知済', '承諾', '辞退', '見送り'],
    CATEGORY: ['新卒', '中途'],
    URGENCY: ['CRITICAL', 'HIGH', 'NORMAL', 'LOW'],
    CONFIDENCE: ['低', '中', '高'],
    RISK_LEVEL: ['Critical', 'High', 'Medium', 'Low', 'None'],
    RECOMMENDATION: ['Pass', 'Conditional', 'Fail'],
    CONTACT_TYPE: ['メール', '電話', '面談', '面接', 'アンケート'],
    ACTION_STATUS: ['未実行', '実行中', '完了'],
    SURVEY_PHASE: ['初回面談', '社員面談', '2次面接', '内定後'],
    SEND_STATUS: ['成功', '失敗'],
    IMPLEMENTATION_STATUS: ['未実施', '実施済'], // 新規追加
    TEST_RESULT: ['合格', '不合格', '未実施'] // 新規追加
  },
  
  // 列幅の設定（ピクセル）
  COLUMN_WIDTHS: {
    ID: 100,
    NAME: 120,
    STATUS: 120,
    DATE: 140,
    SCORE: 80,
    PERCENTAGE: 100,
    TEXT_SHORT: 150,
    TEXT_LONG: 250,
    URL: 200,
    EMAIL: 200 // 新規追加
  },

  // アンケートフォームのURL
  SURVEY_URLS: {
    '初回面談': 'https://docs.google.com/forms/d/e/1FAIpQLSenk59dq_baS1ezizK43FgHkrGopRxw4hmQV9fhT-MNLnyAOw/viewform',
    '社員面談': 'https://docs.google.com/forms/d/e/1FAIpQLSc3kTG6eQVS5Ul6K0xIXlYQQ70T8H-P1ZmBjNDJd-IyKlYNCw/viewform',
    '2次面接': 'https://docs.google.com/forms/d/e/1FAIpQLSfiMPMiFc2dFYfTWqnQI1OvoiHfjt68oCrDudBsq6kjot9lhA/viewform',
    '内定後': 'https://docs.google.com/forms/d/e/1FAIpQLSe9-nE7mAsur_37z_ioSqy8BZF6QsfkOo5Th6znwhHlR-jMKg/viewform'
  },

  // ========== 【新規追加】列番号定数 ==========
  COLUMNS: {
    CANDIDATES_MASTER: {
      CANDIDATE_ID: 0,        // A列
      NAME: 1,                // B列
      STATUS: 2,              // C列
      LAST_UPDATE: 3,         // D列
      CATEGORY: 4,            // E列
      INTERVIEWER: 5,         // F列
      APPLICATION_DATE: 6,    // G列
      LATEST_PASS_PROB: 7,    // H列
      PREV_PASS_PROB: 8,      // I列
      PASS_PROB_CHANGE: 9,    // J列
      LATEST_TOTAL_SCORE: 10, // K列
      LATEST_PHILOSOPHY: 11,  // L列
      LATEST_STRATEGY: 12,    // M列
      LATEST_MOTIVATION: 13,  // N列
      LATEST_EXECUTION: 14,   // O列
      LATEST_ACCEPTANCE_AI: 15,     // P列
      LATEST_ACCEPTANCE_HUMAN: 16,  // Q列
      LATEST_ACCEPTANCE_INTEGRATED: 17, // R列
      PREV_ACCEPTANCE: 18,    // S列
      ACCEPTANCE_CHANGE: 19,  // T列
      CONFIDENCE: 20,         // U列
      ASPIRATION_SCORE: 21,   // V列
      COMPETITIVE_ADVANTAGE: 22, // W列
      CONCERN_RESOLUTION: 23, // X列
      CORE_MOTIVATION: 24,    // Y列
      MAIN_CONCERNS: 25,      // Z列
      COMPETITOR1: 26,        // AA列
      COMPETITOR2: 27,        // AB列
      COMPETITOR3: 28,        // AC列
      OUR_RANKING: 29,        // AD列
      LAST_CONTACT_DATE: 30,  // AE列
      DAYS_SINCE_CONTACT: 31, // AF列
      CONTACT_COUNT: 32,      // AG列
      AVG_CONTACT_INTERVAL: 33, // AH列
      NEXT_CONTACT_DATE: 34,  // AI列
      NEXT_ACTION: 35,        // AJ列
      ACTION_DEADLINE: 36,    // AK列
      ACTION_URGENCY: 37,     // AL列
      ACTION_STATUS: 38,      // AM列
      ACTION_OWNER: 39,       // AN列
      PRIORITY_SCORE: 40,     // AO列
      URGENCY_COEFFICIENT: 41, // AP列
      EMAIL: 42,              // AQ列
      FIRST_INTERVIEW_DATE: 43,    // AR列（新規）
      FIRST_INTERVIEW_STATUS: 44,  // AS列（新規）
      APTITUDE_TEST_DATE: 45,      // AT列（新規）
      APTITUDE_TEST_STATUS: 46,    // AU列（新規）
      FIRST_SELECTION_DATE: 47,    // AV列（新規）
      FIRST_SELECTION_RESULT: 48,  // AW列（新規）
      EMPLOYEE_INTERVIEW_COUNT: 49, // AX列（新規）
      EMPLOYEE_INTERVIEW_DATE: 50,  // AY列（新規）
      EMPLOYEE_INTERVIEW_STATUS: 51, // AZ列（新規）
      SECOND_INTERVIEW_DATE: 52,    // BA列（新規）
      SECOND_INTERVIEW_STATUS: 53,  // BB列（新規）
      FINAL_INTERVIEW_DATE: 54,     // BC列（新規）
      FINAL_INTERVIEW_STATUS: 55,   // BD列（新規）
      SURVEY_RESPONSE_SPEED_SCORE: 56 // BE列（Phase 2 Step 2追加）
    },
    EVALUATION_LOG: {
      LOG_ID: 0,              // A列
      CANDIDATE_ID: 1,        // B列
      NAME: 2,                // C列
      EVAL_DATE: 3,           // D列
      PHASE: 4,               // E列
      INTERVIEWER: 5,         // F列
      OVERALL_SCORE: 6,       // G列
      PHILOSOPHY_SCORE: 7,    // H列
      PHILOSOPHY_REASON: 8,   // I列
      STRATEGY_SCORE: 9,      // J列
      STRATEGY_REASON: 10,    // K列
      MOTIVATION_SCORE: 11,   // L列
      MOTIVATION_REASON: 12,  // M列
      EXECUTION_SCORE: 13,    // N列
      EXECUTION_REASON: 14,   // O列
      HIGHEST_RISK_LEVEL: 15, // P列
      PASS_PROBABILITY: 16,   // Q列
      RECOMMENDATION: 17,     // R列
      DOC_URL: 18,            // S列
      PROACTIVITY_SCORE: 19   // T列（新規）
    },
    SURVEY_SEND_LOG: {
      SEND_ID: 0,             // A列
      CANDIDATE_ID: 1,        // B列
      NAME: 2,                // C列
      EMAIL: 3,               // D列
      PHASE: 4,               // E列
      SEND_TIME: 5,           // F列
      STATUS: 6,              // G列
      ERROR_MSG: 7            // H列
    },

    // ========== 【新規追加】SURVEY_RESPONSE ==========
    SURVEY_RESPONSE: {
      RESPONSE_ID: 0,        // A列
      CANDIDATE_ID: 1,       // B列
      NAME: 2,               // C列
      RESPONSE_DATE: 3,      // D列
      ASPIRATION: 4,         // E列: 志望度
      CONCERNS: 5,           // F列: 懸念事項
      OTHER_COMPANIES: 6,    // G列: 他社選考状況
      COMMENTS: 7,           // H列: その他コメント
      PHASE: 8               // I列: アンケート種別（新規追加）
    },

    // ========== 【新規追加】SURVEY_ANALYSIS ==========
    SURVEY_ANALYSIS: {
      ANALYSIS_ID: 0,           // A列
      CANDIDATE_ID: 1,          // B列
      CANDIDATE_NAME: 2,        // C列
      PHASE: 3,                 // D列: アンケート種別
      SEND_TIME: 4,             // E列: 送信日時
      RESPONSE_TIME: 5,         // F列: 回答日時
      RESPONSE_SPEED_HOURS: 6,  // G列: 回答速度（時間）
      SPEED_SCORE: 7,           // H列: 回答速度スコア（0-100）
      CREATED_AT: 8             // I列: 作成日時
    }
  },

  // Gmail送信制限
  EMAIL: {
    DAILY_LIMIT: 90,          // 安全のため90通まで（実際は100通/日）
    RETRY_COUNT: 3,           // リトライ回数
    RETRY_DELAY: 2000         // リトライ間隔（ミリ秒）
  },

  // Dify API設定
  DIFY: {
    // API設定はスクリプトプロパティに保存
    // PropertiesService.getScriptProperties().getProperty('DIFY_API_URL')
    // PropertiesService.getScriptProperties().getProperty('DIFY_API_KEY')

    TIMEOUT: 30000, // タイムアウト（ミリ秒）
    RETRY_COUNT: 3,  // リトライ回数
    RETRY_DELAY: 1000 // リトライ間隔（ミリ秒）
  },

  // Webhook設定
  WEBHOOK: {
    VALID_TYPES: [
      'evaluation',
      'engagement',
      'evidence',
      'risk',
      'next_q',
      'acceptance_story',
      'competitor_comparison'
    ]
  }
};