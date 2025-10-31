/**
 * 設定値を管理
 */
const CONFIG = {
  // シート名の定義
  SHEET_NAMES: {
    DASHBOARD: 'Dashboard',
    CANDIDATES_MASTER: 'Candidates_Master',
    COMPANY_ASSETS: 'Company_Assets',
    COMPETITOR_DB: 'Competitor_DB',
    EVALUATION_LOG: 'Evaluation_Log',
    ENGAGEMENT_LOG: 'Engagement_Log',
    EVIDENCE: 'Evidence',
    RISK: 'Risk',
    NEXT_Q: 'NextQ',
    CONTACT_HISTORY: 'Contact_History',
    SURVEY_RESPONSE: 'Survey_Response',
    ACCEPTANCE_STORY: 'Acceptance_Story',
    COMPETITOR_COMPARISON: 'Competitor_Comparison',
    ARCHIVE: 'Archive'
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
    STATUS: ['面談', '1次面接', '2次面接', '最終面接', '内定通知済', '承諾', '辞退', '見送り'],
    CATEGORY: ['新卒', '中途'],
    URGENCY: ['CRITICAL', 'HIGH', 'NORMAL', 'LOW'],
    CONFIDENCE: ['低', '中', '高'],
    RISK_LEVEL: ['Critical', 'High', 'Medium', 'Low', 'None'],
    RECOMMENDATION: ['Pass', 'Conditional', 'Fail'],
    CONTACT_TYPE: ['メール', '電話', '面談', '面接', 'アンケート'],
    ACTION_STATUS: ['未実行', '実行中', '完了']
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
    URL: 200
  }
};