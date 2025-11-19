/**
 * AcceptanceCalculation.gs（新規ファイル）
 * アンケート回答速度スコアの計算と承諾可能性への反映
 *
 * 【主要機能】
 * - アンケート回答速度の計算
 * - Survey_Analysisシートへの記録
 * - Candidates_Masterのスコア更新
 * - 全候補者の一括計算
 *
 * @version 1.0
 * @date 2025-11-12
 */

/**
 * ========================================
 * Phase 3-1: 基礎要素スコア計算
 * ========================================
 */

// デフォルト値の定義
const DEFAULT_SCORES = {
  MOTIVATION: 50,           // 志望度スコア
  COMPETITIVE: 50,          // 競合優位性スコア
  CONCERN_RESOLUTION: 70,   // 懸念解消度スコア
  FOUNDATION: 50            // 基礎要素スコア
};

// 志望度変化係数のマッピング
const MOTIVATION_CHANGE_COEFFICIENT = {
  '大きく上がった': 1.2,
  '非常に高まった': 1.2,
  'やや上がった': 1.1,
  '上がった': 1.1,
  '高まった': 1.1,
  '変わらない': 1.0,
  '変化なし': 1.0,
  'やや下がった': 0.85,
  'やや低下した': 0.85,
  '下がった': 0.85,
  '大きく下がった': 0.7,
  '低下した': 0.7
};

// 懸念レベルのキーワード定義
const CONCERN_KEYWORDS = {
  CRITICAL: [
    '給与', '年収', '報酬', '待遇', 'salary', '給料',
    '勤務地', '転勤', '配属', 'location', '勤務先',
    '仕事内容', '業務内容', 'job', '職務', '業務',
    '契約', '雇用形態', 'contract', '雇用条件',
    '残業', '労働時間', '休日', '休暇', 'overtime'
  ],
  HIGH: [
    'キャリア', 'career', '昇進', '昇格', 'promotion',
    '成長', 'growth', 'スキル', 'skill', '能力開発',
    '社風', 'カルチャー', 'culture', '企業文化', '雰囲気',
    '評価', '査定', 'evaluation', '人事', '評価制度',
    '福利厚生', '研修', '教育', 'training', '制度'
  ],
  MEDIUM: [
    'オフィス', 'office', '環境', '設備', 'facilities',
    '通勤', 'commute', 'アクセス', '交通',
    'その他', 'other', '特になし'
  ]
};

// 競合状況スコアの係数
const COMPETITIVE_STATUS_POINTS = {
  APPLICATIONS: {
    ZERO: 20,      // 選考中0社
    LOW: 10,       // 1-2社
    MEDIUM: 0,     // 3-4社
    HIGH: -10      // 5社以上
  },
  OFFERS: {
    ZERO: 20,      // 内定0社
    ONE: 5,        // 内定1社
    MULTIPLE: -10  // 内定2社以上
  },
  TIMING: {
    NONE: 10,      // 他社最終面接なし
    FAR: 5,        // 1ヶ月以上先
    NEAR: -10      // 2週間以内
  }
};

// 懸念解消度のペナルティ
const CONCERN_PENALTY = {
  CRITICAL: 40,
  HIGH: 20,
  MEDIUM: 10
};

/**
 * アンケート回答速度を計算してシートを更新
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {boolean} 更新成功/失敗
 *
 * @example
 * calculateAndUpdateResponseSpeed('C001', '初回面談');
 */
function calculateAndUpdateResponseSpeed(candidateId, phase) {
  try {
    Logger.log(`📊 アンケート回答速度を計算: ${candidateId} (${phase})`);

    // DataGetters.gsの関数を使用して回答速度データを取得
    const speedData = getResponseSpeedData(candidateId, phase);

    if (!speedData) {
      Logger.log(`⚠️ 回答速度データが取得できません: ${candidateId} (${phase})`);
      return false;
    }

    Logger.log(`✅ 回答速度: ${speedData.hours}時間、スコア: ${speedData.score}点`);

    // 1. Survey_Analysisシートに記録
    const analysisSuccess = saveSurveyAnalysis(candidateId, phase, speedData);

    // 2. Candidates_MasterのBE列を更新
    const masterSuccess = updateCandidatesMasterSpeedScore(candidateId);

    if (analysisSuccess && masterSuccess) {
      Logger.log(`✅ アンケート回答速度の更新完了: ${candidateId} (${phase})`);
      return true;
    } else {
      Logger.log(`⚠️ 一部の更新に失敗: ${candidateId} (${phase})`);
      return false;
    }

  } catch (error) {
    Logger.log(`❌ calculateAndUpdateResponseSpeedエラー: ${error.message}`);
    Logger.log(error.stack);
    return false;
  }
}

/**
 * Survey_Analysisシートに分析データを保存
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @param {Object} speedData - 回答速度データ
 * @return {boolean} 保存成功/失敗
 */
function saveSurveyAnalysis(candidateId, phase, speedData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_ANALYSIS);

    if (!sheet) {
      Logger.log('⚠️ Survey_Analysisシートが見つかりません。作成します...');
      setupSurveyAnalysis();
      sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_ANALYSIS);
    }

    // 候補者名を取得
    const masterSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);
    const masterData = masterSheet.getDataRange().getValues();
    let candidateName = '';

    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID] === candidateId) {
        candidateName = masterData[i][CONFIG.COLUMNS.CANDIDATES_MASTER.NAME];
        break;
      }
    }

    // analysis_id を生成
    const analysisId = 'ANAL-' + candidateId + '-' + phase.replace(/\s/g, '') + '-' + new Date().getTime();

    // データを追加
    sheet.appendRow([
      analysisId,                    // A: analysis_id
      candidateId,                   // B: candidate_id
      candidateName,                 // C: candidate_name
      phase,                         // D: phase
      speedData.send_time,           // E: send_time
      speedData.response_time,       // F: response_time
      speedData.hours,               // G: response_speed_hours
      speedData.score,               // H: speed_score
      new Date()                     // I: created_at
    ]);

    Logger.log(`✅ Survey_Analysisに記録: ${candidateId} (${phase})`);
    return true;

  } catch (error) {
    Logger.log(`❌ saveSurveyAnalysisエラー: ${error.message}`);
    return false;
  }
}

/**
 * Candidates_MasterのBE列（アンケート回答速度スコア）を更新
 *
 * 全アンケートの平均スコアを計算してBE列に保存
 *
 * @param {string} candidateId - 候補者ID
 * @return {boolean} 更新成功/失敗
 */
function updateCandidatesMasterSpeedScore(candidateId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);
    const analysisSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_ANALYSIS);

    if (!masterSheet || !analysisSheet) {
      Logger.log('⚠️ 必要なシートが見つかりません');
      return false;
    }

    // Survey_Analysisから該当候補者の全スコアを取得
    const analysisData = analysisSheet.getDataRange().getValues();
    const scores = [];

    for (let i = 1; i < analysisData.length; i++) {
      if (analysisData[i][CONFIG.COLUMNS.SURVEY_ANALYSIS.CANDIDATE_ID] === candidateId) {
        const score = analysisData[i][CONFIG.COLUMNS.SURVEY_ANALYSIS.SPEED_SCORE];
        if (score !== '' && score !== null && !isNaN(score)) {
          scores.push(Number(score));
        }
      }
    }

    if (scores.length === 0) {
      Logger.log(`⚠️ スコアが見つかりません: ${candidateId}`);
      return false;
    }

    // 平均スコアを計算
    const avgScore = Math.round(scores.reduce((a, b) => a + b, 0) / scores.length);

    // Candidates_Masterの該当行を検索してBE列を更新
    const masterData = masterSheet.getDataRange().getValues();

    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID] === candidateId) {
        const beColumn = CONFIG.COLUMNS.CANDIDATES_MASTER.SURVEY_RESPONSE_SPEED_SCORE + 1; // 列番号は0始まりなので+1
        masterSheet.getRange(i + 1, beColumn).setValue(avgScore);

        Logger.log(`✅ Candidates_MasterのBE列を更新: ${candidateId} - ${avgScore}点`);
        return true;
      }
    }

    Logger.log(`⚠️ 候補者が見つかりません: ${candidateId}`);
    return false;

  } catch (error) {
    Logger.log(`❌ updateCandidatesMasterSpeedScoreエラー: ${error.message}`);
    return false;
  }
}

/**
 * 全候補者のアンケート回答速度を一括計算
 *
 * Survey_Send_LogとSurvey_Responseを突合して、
 * 全候補者・全アンケート種別の回答速度を計算
 *
 * @return {Object} { success: 成功数, failed: 失敗数, total: 合計数 }
 */
function calculateAllResponseSpeeds() {
  try {
    Logger.log('📊 全候補者のアンケート回答速度を一括計算開始...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sendLogSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);

    if (!sendLogSheet) {
      throw new Error('Survey_Send_Logシートが見つかりません');
    }

    const sendLogData = sendLogSheet.getDataRange().getValues();

    let successCount = 0;
    let failedCount = 0;
    let totalCount = 0;

    // 送信成功したログを全て処理
    for (let i = 1; i < sendLogData.length; i++) {
      const candidateId = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.CANDIDATE_ID];
      const phase = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.PHASE];
      const status = sendLogData[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.STATUS];

      // 送信成功したもののみ処理
      if (status === '成功') {
        totalCount++;

        // 回答速度を計算・更新
        const success = calculateAndUpdateResponseSpeed(candidateId, phase);

        if (success) {
          successCount++;
        } else {
          failedCount++;
        }

        // API制限を避けるため、少し待機
        if (totalCount % 10 === 0) {
          Utilities.sleep(1000); // 10件ごとに1秒待機
        }
      }
    }

    const result = {
      success: successCount,
      failed: failedCount,
      total: totalCount
    };

    Logger.log(`✅ 一括計算完了: 成功 ${successCount}件、失敗 ${failedCount}件、合計 ${totalCount}件`);

    // 結果をダイアログで表示
    SpreadsheetApp.getUi().alert(
      '一括計算完了',
      `アンケート回答速度の一括計算が完了しました。\n\n` +
      `成功: ${successCount}件\n` +
      `失敗: ${failedCount}件\n` +
      `合計: ${totalCount}件\n\n` +
      'Survey_Analysisシートで詳細を確認できます。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return result;

  } catch (error) {
    Logger.log(`❌ calculateAllResponseSpeedsエラー: ${error.message}`);
    Logger.log(error.stack);

    SpreadsheetApp.getUi().alert(
      'エラー',
      `一括計算中にエラーが発生しました:\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return { success: 0, failed: 0, total: 0 };
  }
}

/**
 * 特定の候補者の全アンケート回答速度を再計算
 *
 * @param {string} candidateId - 候補者ID
 * @return {number} 計算成功数
 */
function recalculateCandidateResponseSpeeds(candidateId) {
  try {
    Logger.log(`📊 候補者の回答速度を再計算: ${candidateId}`);

    const phases = ['初回面談', '社員面談', '2次面接', '内定後'];
    let successCount = 0;

    phases.forEach(phase => {
      const success = calculateAndUpdateResponseSpeed(candidateId, phase);
      if (success) {
        successCount++;
      }
    });

    Logger.log(`✅ 再計算完了: ${candidateId} - ${successCount}件成功`);
    return successCount;

  } catch (error) {
    Logger.log(`❌ recalculateCandidateResponseSpeedsエラー: ${error.message}`);
    return 0;
  }
}

/**
 * Survey_Analysisシートから候補者の平均回答速度スコアを取得
 *
 * @param {string} candidateId - 候補者ID
 * @return {number|null} 平均スコア（0-100）、データがない場合null
 */
function getAverageResponseSpeedScore(candidateId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_ANALYSIS);

    if (!sheet) {
      return null;
    }

    const data = sheet.getDataRange().getValues();
    const scores = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][CONFIG.COLUMNS.SURVEY_ANALYSIS.CANDIDATE_ID] === candidateId) {
        const score = data[i][CONFIG.COLUMNS.SURVEY_ANALYSIS.SPEED_SCORE];
        if (score !== '' && score !== null && !isNaN(score)) {
          scores.push(Number(score));
        }
      }
    }

    if (scores.length === 0) {
      return null;
    }

    return Math.round(scores.reduce((a, b) => a + b, 0) / scores.length);

  } catch (error) {
    Logger.log(`❌ getAverageResponseSpeedScoreエラー: ${error.message}`);
    return null;
  }
}

/**
 * 回答速度スコアの統計情報を取得
 *
 * @return {Object} { average, max, min, count }
 */
function getResponseSpeedStatistics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_ANALYSIS);

    if (!sheet) {
      return null;
    }

    const data = sheet.getDataRange().getValues();
    const scores = [];

    for (let i = 1; i < data.length; i++) {
      const score = data[i][CONFIG.COLUMNS.SURVEY_ANALYSIS.SPEED_SCORE];
      if (score !== '' && score !== null && !isNaN(score)) {
        scores.push(Number(score));
      }
    }

    if (scores.length === 0) {
      return null;
    }

    const average = Math.round(scores.reduce((a, b) => a + b, 0) / scores.length);
    const max = Math.max(...scores);
    const min = Math.min(...scores);

    return {
      average: average,
      max: max,
      min: min,
      count: scores.length
    };

  } catch (error) {
    Logger.log(`❌ getResponseSpeedStatisticsエラー: ${error.message}`);
    return null;
  }
}

/**
 * ========================================
 * Phase 3-1: 志望度スコア関数
 * ========================================
 */

/**
 * 志望度スコアの計算
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別（初回面談/社員面談/2次面接/内定後）
 * @return {number} 志望度スコア（0-100点）
 *
 * 計算式:
 * 志望度スコア = (アンケート回答値 × 10) × 志望度変化係数
 *
 * データ取得元:
 * - 初回面談: F列（Q5. 志望度 1-10）
 * - 社員面談: F列（Q5. 志望度 1-10）
 * - 2次面接: F列（Q5. 志望度 1-10）
 * - 内定後: H列（Q7. PIGNUSで働くことへの前向きさ 1-10）
 */
function calculateMotivationScore(candidateId, phase) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // フェーズに応じてシートと列を決定
    let sheetName, column;
    switch(phase) {
      case '初回面談':
        sheetName = 'アンケート_初回面談';
        column = 'F'; // Q5. 志望度
        break;
      case '社員面談':
        sheetName = 'アンケート_社員面談';
        column = 'F'; // Q5. 志望度
        break;
      case '2次面接':
        sheetName = 'アンケート_2次面接';
        column = 'F'; // Q5. 志望度
        break;
      case '内定後':
        sheetName = 'アンケート_内定';
        column = 'H'; // Q7. PIGNUSで働くことへの前向きさ（1-10）
        break;
      default:
        Logger.log(`⚠️ 不明なフェーズ: ${phase}`);
        return DEFAULT_SCORES.MOTIVATION;
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`❌ シートが見つかりません: ${sheetName}`);
      return DEFAULT_SCORES.MOTIVATION;
    }

    const data = sheet.getDataRange().getValues();

    // 候補者のメールアドレスを取得
    const email = getCandidateEmail(candidateId);
    if (!email) {
      Logger.log(`❌ メールアドレスが見つかりません: ${candidateId}`);
      return DEFAULT_SCORES.MOTIVATION;
    }

    // メールアドレスで候補者の回答を検索（C列）
    let motivationRaw = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === email) { // C列: メールアドレス
        const colIndex = column.charCodeAt(0) - 65;
        motivationRaw = data[i][colIndex];
        break;
      }
    }

    if (motivationRaw === null || motivationRaw === '' || motivationRaw === undefined) {
      Logger.log(`⚠️ 志望度データなし: ${candidateId}, ${phase}`);
      return DEFAULT_SCORES.MOTIVATION;
    }

    // 数値変換
    const motivationValue = Number(motivationRaw);
    if (isNaN(motivationValue)) {
      Logger.log(`⚠️ 志望度が数値ではありません: ${motivationRaw}`);
      return DEFAULT_SCORES.MOTIVATION;
    }

    // 志望度変化係数を取得
    const changeCoefficient = getMotivationChangeCoefficient(candidateId, phase);

    // 志望度スコア = (回答値 × 10) × 変化係数
    const motivationScore = motivationValue * 10 * changeCoefficient;

    Logger.log(`✅ 志望度スコア: ${motivationScore.toFixed(1)}（生値: ${motivationValue}, 係数: ${changeCoefficient}）`);

    // 0-100の範囲に収める
    return Math.min(Math.max(Math.round(motivationScore), 0), 100);

  } catch (error) {
    Logger.log(`❌ 志望度スコア計算エラー: ${error}`);
    return DEFAULT_SCORES.MOTIVATION;
  }
}

/**
 * 志望度変化係数の取得
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {number} 変化係数（0.7-1.2）
 *
 * データ取得元:
 * - 初回面談: E列（Q4. 興味の変化）
 * - 社員面談: E列（Q4. 志望度の変化）
 * - 2次面接: E列（Q4. 志望度の変化）
 * - 内定後: F列（Q5. 働きたい気持ちの変化）
 */
function getMotivationChangeCoefficient(candidateId, phase) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let sheetName, column;
    switch(phase) {
      case '初回面談':
        sheetName = 'アンケート_初回面談';
        column = 'E'; // Q4
        break;
      case '社員面談':
        sheetName = 'アンケート_社員面談';
        column = 'E'; // Q4
        break;
      case '2次面接':
        sheetName = 'アンケート_2次面接';
        column = 'E'; // Q4
        break;
      case '内定後':
        sheetName = 'アンケート_内定';
        column = 'F'; // Q5
        break;
      default:
        return 1.0;
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return 1.0;
    }

    const data = sheet.getDataRange().getValues();
    const email = getCandidateEmail(candidateId);

    if (!email) {
      return 1.0;
    }

    let changeText = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === email) {
        const colIndex = column.charCodeAt(0) - 65;
        changeText = data[i][colIndex];
        break;
      }
    }

    if (!changeText || changeText === '') {
      Logger.log(`⚠️ 志望度変化データなし: ${candidateId}, ${phase}`);
      return 1.0;
    }

    // マッピングオブジェクトで検索
    const changeStr = String(changeText);
    for (let key in MOTIVATION_CHANGE_COEFFICIENT) {
      if (changeStr.includes(key)) {
        const coefficient = MOTIVATION_CHANGE_COEFFICIENT[key];
        Logger.log(`✅ 志望度変化係数: ${coefficient} (「${changeStr}」)`);
        return coefficient;
      }
    }

    Logger.log(`⚠️ 変化テキストが未定義: ${changeStr} → デフォルト1.0`);
    return 1.0; // デフォルト

  } catch (error) {
    Logger.log(`❌ 変化係数取得エラー: ${error}`);
    return 1.0;
  }
}

/**
 * ========================================
 * Phase 3-1: 競合優位性スコア関数
 * ========================================
 */

/**
 * 競合優位性スコアの計算
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {number} 競合優位性スコア（0-100点）
 *
 * 計算式:
 * 競合優位性スコア = 50（基本点） + 競合状況スコア + 自社優位性スコア
 */
function calculateCompetitiveAdvantageScore(candidateId, phase) {
  try {
    let score = 50; // 基本点

    // 競合状況スコアを加算
    const competitiveStatusScore = calculateCompetitiveStatusScore(candidateId, phase);
    score += competitiveStatusScore;

    // 自社優位性スコア（内定後のみ）
    if (phase === '内定後') {
      const ownAdvantageScore = calculateOwnAdvantageScore(candidateId);
      score += ownAdvantageScore;
      Logger.log(`✅ 自社優位性スコア: ${ownAdvantageScore}`);
    }

    Logger.log(`✅ 競合優位性スコア: ${score}`);

    // 上限100、下限0
    return Math.min(Math.max(Math.round(score), 0), 100);

  } catch (error) {
    Logger.log(`❌ 競合優位性スコア計算エラー: ${error}`);
    return DEFAULT_SCORES.COMPETITIVE;
  }
}

/**
 * 競合状況スコアの計算
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {number} 競合状況スコア（-30 ~ +50）
 */
function calculateCompetitiveStatusScore(candidateId, phase) {
  let score = 0;

  // 選考中企業数
  const applicationsCount = getApplicationsCount(candidateId, phase);
  if (applicationsCount === 0) {
    score += COMPETITIVE_STATUS_POINTS.APPLICATIONS.ZERO;
  } else if (applicationsCount <= 2) {
    score += COMPETITIVE_STATUS_POINTS.APPLICATIONS.LOW;
  } else if (applicationsCount <= 4) {
    score += COMPETITIVE_STATUS_POINTS.APPLICATIONS.MEDIUM;
  } else {
    score += COMPETITIVE_STATUS_POINTS.APPLICATIONS.HIGH;
  }

  // 内定済企業数
  const offersCount = getOffersCount(candidateId, phase);
  if (offersCount === 0) {
    score += COMPETITIVE_STATUS_POINTS.OFFERS.ZERO;
  } else if (offersCount === 1) {
    score += COMPETITIVE_STATUS_POINTS.OFFERS.ONE;
  } else {
    score += COMPETITIVE_STATUS_POINTS.OFFERS.MULTIPLE;
  }

  // 他社最終面接の時期（2次面接・内定後のみ）
  if (phase === '2次面接' || phase === '内定後') {
    const finalInterviewTiming = getFinalInterviewTiming(candidateId, phase);
    if (finalInterviewTiming === 'なし') {
      score += COMPETITIVE_STATUS_POINTS.TIMING.NONE;
    } else if (finalInterviewTiming === '1ヶ月以上先') {
      score += COMPETITIVE_STATUS_POINTS.TIMING.FAR;
    } else if (finalInterviewTiming === '近い') {
      score += COMPETITIVE_STATUS_POINTS.TIMING.NEAR;
    }
  }

  Logger.log(`✅ 競合状況スコア: ${score}（選考中: ${applicationsCount}社, 内定: ${offersCount}社）`);

  return score;
}

/**
 * 選考中企業数を取得
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {number} 選考中企業数
 */
function getApplicationsCount(candidateId, phase) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const email = getCandidateEmail(candidateId);

  if (!email) return 0;

  let sheetName, column;
  switch(phase) {
    case '初回面談':
      sheetName = 'アンケート_初回面談';
      column = 'K'; // Q10-1
      break;
    case '社員面談':
      sheetName = 'アンケート_社員面談';
      column = 'M'; // Q13.1
      break;
    case '2次面接':
      sheetName = 'アンケート_2次面接';
      column = 'O'; // Q10-1
      break;
    case '内定後':
      sheetName = 'アンケート_内定';
      column = 'L'; // Q11
      break;
    default:
      return 0;
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === email) {
      const colIndex = column.charCodeAt(0) - 65;
      const value = data[i][colIndex];

      // 数値の場合はそのまま
      if (typeof value === 'number') {
        return value;
      }

      // テキストから数値を抽出（例: "2社選考中"→2）
      const match = String(value).match(/(\d+)/);
      return match ? parseInt(match[1]) : 0;
    }
  }

  return 0;
}

/**
 * 内定済企業数を取得
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {number} 内定済企業数
 */
function getOffersCount(candidateId, phase) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const email = getCandidateEmail(candidateId);

  if (!email) return 0;

  let sheetName, column;
  switch(phase) {
    case '初回面談':
      sheetName = 'アンケート_初回面談';
      column = 'L'; // Q10-2
      break;
    case '社員面談':
      sheetName = 'アンケート_社員面談';
      column = 'N'; // Q13-2
      break;
    case '2次面接':
      sheetName = 'アンケート_2次面接';
      column = 'P'; // Q10-2
      break;
    case '内定後':
      sheetName = 'アンケート_内定';
      column = 'K'; // Q10
      break;
    default:
      return 0;
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === email) {
      const colIndex = column.charCodeAt(0) - 65;
      const value = data[i][colIndex];

      if (typeof value === 'number') {
        return value;
      }

      const match = String(value).match(/(\d+)/);
      return match ? parseInt(match[1]) : 0;
    }
  }

  return 0;
}

/**
 * 他社最終面接の時期を取得
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {string} 'なし' | '1ヶ月以上先' | '近い'
 */
function getFinalInterviewTiming(candidateId, phase) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const email = getCandidateEmail(candidateId);

  if (!email) return 'なし';

  let sheetName, column;
  if (phase === '2次面接') {
    sheetName = 'アンケート_2次面接';
    column = 'Q'; // Q13-3
  } else if (phase === '内定後') {
    sheetName = 'アンケート_内定';
    column = 'J'; // Q9. 意思決定の期限
  } else {
    return 'なし';
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 'なし';

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === email) {
      const colIndex = column.charCodeAt(0) - 65;
      const value = String(data[i][colIndex]);

      if (!value || value === '' || value === 'undefined') {
        return 'なし';
      }

      // テキスト解析
      if (value.includes('なし') || value.includes('予定なし') || value.includes('特になし')) {
        return 'なし';
      }
      if (value.includes('1ヶ月以上') || value.includes('まだ先') || value.includes('ない')) {
        return '1ヶ月以上先';
      }
      if (value.includes('来週') || value.includes('近い') || value.includes('今週') || value.includes('2週間以内')) {
        return '近い';
      }

      return 'なし'; // デフォルト
    }
  }

  return 'なし';
}

/**
 * 自社優位性スコアの計算（内定後のみ）
 *
 * @param {string} candidateId - 候補者ID
 * @return {number} 自社優位性スコア（0-100点）
 *
 * データ取得元:
 * - 内定後: M列（Q12. PIGNUSの優位性, 1-10）
 */
function calculateOwnAdvantageScore(candidateId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('アンケート_内定');
  const email = getCandidateEmail(candidateId);

  if (!sheet || !email) return 0;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === email) {
      const advantageRaw = data[i][12]; // M列（Q12, 1-10）

      if (typeof advantageRaw === 'number') {
        // 1-10を0-100に変換
        return advantageRaw * 10;
      }

      return 0;
    }
  }

  return 0;
}

/**
 * ========================================
 * Phase 3-1: 懸念解消度スコア関数
 * ========================================
 */

/**
 * 懸念解消度スコアの計算
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {number} 懸念解消度スコア（0-100点）
 *
 * 計算式:
 * 懸念解消度スコア = 100 - (Critical懸念×40 + High懸念×20 + Medium懸念×10)
 *
 * データ取得元:
 * - 初回面談: J列（Q9. 不安・懸念事項）
 * - 社員面談: K列（Q11. まだ不安に感じている点）
 * - 2次面接: K列（Q10. 入社を決断するために必要な情報）
 * - 内定後: Q列（Q16. 入社を決断するために解消したい不安）
 */
function calculateConcernResolutionScore(candidateId, phase) {
  try {
    const concerns = getConcerns(candidateId, phase);

    if (!concerns || concerns.length === 0) {
      Logger.log(`✅ 懸念事項なし: ${candidateId} → 満点`);
      return 100; // 懸念なし = 満点
    }

    let criticalCount = 0;
    let highCount = 0;
    let mediumCount = 0;

    // 懸念事項を分類
    for (let concern of concerns) {
      const level = classifyConcernLevel(concern);
      if (level === 'Critical') {
        criticalCount++;
      } else if (level === 'High') {
        highCount++;
      } else if (level === 'Medium') {
        mediumCount++;
      }
    }

    // スコア計算
    let score = 100 - (
      criticalCount * CONCERN_PENALTY.CRITICAL +
      highCount * CONCERN_PENALTY.HIGH +
      mediumCount * CONCERN_PENALTY.MEDIUM
    );

    Logger.log(`✅ 懸念解消度スコア: ${score}（Critical: ${criticalCount}, High: ${highCount}, Medium: ${mediumCount}）`);

    // 下限0点
    return Math.max(0, Math.round(score));

  } catch (error) {
    Logger.log(`❌ 懸念解消度スコア計算エラー: ${error}`);
    return DEFAULT_SCORES.CONCERN_RESOLUTION;
  }
}

/**
 * 懸念事項を取得
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {Array<string>} 懸念事項の配列
 */
function getConcerns(candidateId, phase) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const email = getCandidateEmail(candidateId);

    if (!email) {
      Logger.log(`❌ メールアドレスなし: ${candidateId}`);
      return [];
    }

    let sheetName, column;
    switch(phase) {
      case '初回面談':
        sheetName = 'アンケート_初回面談';
        column = 'J'; // Q9
        break;
      case '社員面談':
        sheetName = 'アンケート_社員面談';
        column = 'K'; // Q11
        break;
      case '2次面接':
        sheetName = 'アンケート_2次面接';
        column = 'K'; // Q10
        break;
      case '内定後':
        sheetName = 'アンケート_内定';
        column = 'Q'; // Q16
        break;
      default:
        return [];
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`❌ シートなし: ${sheetName}`);
      return [];
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === email) { // C列: メールアドレス
        const colIndex = column.charCodeAt(0) - 65;
        const concernText = String(data[i][colIndex]);

        if (!concernText || concernText === '' || concernText === 'undefined' || concernText === 'null') {
          return [];
        }

        // 「特になし」などの場合は空配列を返す
        if (concernText.includes('特になし') || concernText.includes('なし') || concernText.includes('ない')) {
          return [];
        }

        // 複数選択の場合、カンマまたは改行で分割
        const concerns = concernText.split(/[,、\n]/)
          .map(c => c.trim())
          .filter(c => c !== '' && c !== 'なし' && c !== '特になし');

        Logger.log(`✅ 懸念事項: ${concerns.join(', ')}`);

        return concerns;
      }
    }

    return [];

  } catch (error) {
    Logger.log(`❌ 懸念事項取得エラー: ${error}`);
    return [];
  }
}

/**
 * 懸念レベルの分類
 *
 * @param {string} concernText - 懸念事項のテキスト
 * @return {string} 懸念レベル（'Critical' | 'High' | 'Medium'）
 *
 * 分類基準:
 * Critical: 給与水準、勤務地、仕事内容、契約条件、労働時間
 * High: キャリアパス、成長機会、社風、評価制度、福利厚生
 * Medium: オフィス環境、通勤、設備、その他
 */
function classifyConcernLevel(concernText) {
  // 正規化（小文字化）
  const text = concernText.toLowerCase();

  // Critical（重大）
  for (let keyword of CONCERN_KEYWORDS.CRITICAL) {
    if (text.includes(keyword.toLowerCase())) {
      Logger.log(`  → Critical: 「${concernText}」（キーワード: ${keyword}）`);
      return 'Critical';
    }
  }

  // High（高）
  for (let keyword of CONCERN_KEYWORDS.HIGH) {
    if (text.includes(keyword.toLowerCase())) {
      Logger.log(`  → High: 「${concernText}」（キーワード: ${keyword}）`);
      return 'High';
    }
  }

  // Medium（中）
  Logger.log(`  → Medium: 「${concernText}」`);
  return 'Medium';
}

/**
 * ========================================
 * Phase 3-1: 基礎要素スコア統合関数
 * ========================================
 */

/**
 * 基礎要素スコアの計算（統合）
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {number} 基礎要素スコア（0-100点）
 *
 * 計算式:
 * 基礎要素 = 志望度スコア(50%) + 競合優位性スコア(30%) + 懸念解消度スコア(20%)
 */
function calculateFoundationScore(candidateId, phase) {
  try {
    Logger.log(`\n========================================`);
    Logger.log(`基礎要素スコア計算開始: ${candidateId}, ${phase}`);
    Logger.log(`========================================`);

    // 各要素を計算
    const motivationScore = calculateMotivationScore(candidateId, phase);
    const competitiveScore = calculateCompetitiveAdvantageScore(candidateId, phase);
    const concernScore = calculateConcernResolutionScore(candidateId, phase);

    // 重み付け統合
    const foundationScore =
      motivationScore * 0.5 +
      competitiveScore * 0.3 +
      concernScore * 0.2;

    Logger.log(`\n--- 基礎要素スコア内訳 ---`);
    Logger.log(`志望度スコア: ${motivationScore} × 0.5 = ${(motivationScore * 0.5).toFixed(1)}`);
    Logger.log(`競合優位性スコア: ${competitiveScore} × 0.3 = ${(competitiveScore * 0.3).toFixed(1)}`);
    Logger.log(`懸念解消度スコア: ${concernScore} × 0.2 = ${(concernScore * 0.2).toFixed(1)}`);
    Logger.log(`基礎要素スコア（合計）: ${foundationScore.toFixed(1)}`);
    Logger.log(`========================================\n`);

    return Math.round(foundationScore); // 四捨五入

  } catch (error) {
    Logger.log(`❌ 基礎要素スコア計算エラー: ${error}`);
    return DEFAULT_SCORES.FOUNDATION;
  }
}

/**
 * ========================================
 * Phase 3-2: 関係性・行動シグナル要素スコア計算
 * ========================================
 */

// Phase 3-2のデフォルト値
const DEFAULT_SCORES_PHASE3_2 = {
  RELATIONSHIP: 70,          // 関係性要素スコア
  CONTACT_COUNT: 50,         // 接点回数スコア
  INTERVAL: 70,              // 接点間隔スコア
  QUALITY: 70,               // 接点の質スコア
  GAP: 70,                   // 空白期間スコア
  BEHAVIOR: 70,              // 行動シグナル要素スコア
  DESCRIPTION: 70,           // 自由記述スコア
  SELECTION_SPEED: 80,       // 選考スピードスコア
  RESPONSE_SPEED: 70,        // 回答速度スコア
  PROACTIVITY: 70            // 積極性スコア
};

// 接点回数のスコアリング基準
const CONTACT_COUNT_SCORING = {
  ZERO: 0,
  LOW: 30,       // 1-2回
  MEDIUM: 50,    // 3-4回
  HIGH: 70,      // 5-6回
  VERY_HIGH: 85, // 7-9回
  EXCELLENT: 100 // 10回以上
};

// 接点間隔のスコアリング基準（日数）
const INTERVAL_SCORING = {
  WEEKLY: { days: 7, score: 100 },       // 週1回以上
  BIWEEKLY: { days: 14, score: 80 },     // 2週間に1回
  TRIWEEKLY: { days: 21, score: 60 },    // 3週間に1回
  MONTHLY: { days: 30, score: 40 },      // 月1回
  RARE: { days: 31, score: 20 }          // 月1回未満
};

// 空白期間のスコアリング基準（日数）
const GAP_SCORING = {
  RECENT: { days: 7, score: 100 },       // 1週間以内
  FAIRLY_RECENT: { days: 14, score: 80 }, // 2週間以内
  MODERATE: { days: 21, score: 60 },     // 3週間以内
  OLD: { days: 30, score: 40 },          // 1ヶ月以内
  VERY_OLD: { days: 31, score: 20 }      // 1ヶ月以上
};

// 志望度変化のスコアリング基準
const MOTIVATION_CHANGE_SCORING = {
  LARGE_INCREASE: { change: 2, score: 100 },    // 大きく上昇
  INCREASE: { change: 1, score: 90 },           // 上昇
  STABLE: { change: 0, score: 80 },             // 維持
  SLIGHT_DECREASE: { change: -1, score: 60 },   // 微減
  DECREASE: { change: -2, score: 40 }           // 大きく下降
};

// 自由記述のスコアリング基準（文字数）
const DESCRIPTION_SCORING = {
  NONE: { length: 0, score: 0 },
  VERY_SHORT: { length: 20, score: 20 },   // ほとんど書いていない
  SHORT: { length: 50, score: 40 },        // 短い
  MEDIUM: { length: 100, score: 60 },      // 普通
  LONG: { length: 200, score: 80 },        // 詳しい
  VERY_LONG: { length: 201, score: 100 }   // 非常に詳しい
};

// 選考スピードのスコアリング基準（日数）
const SELECTION_SPEED_SCORING = {
  VERY_FAST: { days: 14, score: 100 },     // 2週間以内
  FAST: { days: 30, score: 90 },           // 1ヶ月以内
  NORMAL: { days: 60, score: 80 },         // 2ヶ月以内
  SLOW: { days: 90, score: 60 },           // 3ヶ月以内
  VERY_SLOW: { days: 91, score: 40 }       // 3ヶ月以上
};

/**
 * ========================================
 * Phase 3-2: 関係性要素の計算関数
 * ========================================
 */

/**
 * 関係性要素スコアの計算（統合）
 *
 * @param {string} candidateId - 候補者ID
 * @return {number} 関係性要素スコア（0-100点）
 *
 * 計算式:
 * 関係性要素 = 接点回数(30%) + 接点間隔(20%) + 接点の質(25%) + 空白期間(25%)
 */
function calculateRelationshipScore(candidateId) {
  try {
    Logger.log(`\n========================================`);
    Logger.log(`関係性要素スコア計算: ${candidateId}`);
    Logger.log(`========================================`);

    const contactCountScore = calculateContactCountScore(candidateId);
    const intervalScore = calculateIntervalScore(candidateId);
    const qualityScore = calculateQualityScore(candidateId);
    const gapScore = calculateGapScore(candidateId);

    const relationshipScore =
      contactCountScore * 0.3 +
      intervalScore * 0.2 +
      qualityScore * 0.25 +
      gapScore * 0.25;

    Logger.log(`\n--- 関係性要素スコア内訳 ---`);
    Logger.log(`接点回数スコア: ${contactCountScore} × 0.3 = ${(contactCountScore * 0.3).toFixed(1)}`);
    Logger.log(`接点間隔スコア: ${intervalScore} × 0.2 = ${(intervalScore * 0.2).toFixed(1)}`);
    Logger.log(`接点の質スコア: ${qualityScore} × 0.25 = ${(qualityScore * 0.25).toFixed(1)}`);
    Logger.log(`空白期間スコア: ${gapScore} × 0.25 = ${(gapScore * 0.25).toFixed(1)}`);
    Logger.log(`関係性要素スコア（合計）: ${relationshipScore.toFixed(1)}`);
    Logger.log(`========================================\n`);

    return Math.round(relationshipScore);

  } catch (error) {
    Logger.log(`❌ 関係性要素スコアエラー: ${error}`);
    return DEFAULT_SCORES_PHASE3_2.RELATIONSHIP;
  }
}

/**
 * 接点回数スコアの計算
 *
 * @param {string} candidateId - 候補者ID
 * @return {number} 接点回数スコア（0-100点）
 *
 * スコアリング基準:
 * 0回: 0点
 * 1-2回: 30点
 * 3-4回: 50点
 * 5-6回: 70点
 * 7-9回: 85点
 * 10回以上: 100点
 */
function calculateContactCountScore(candidateId) {
  try {
    const contacts = getContactHistory(candidateId);
    const count = contacts.length;

    let score;
    if (count === 0) {
      score = CONTACT_COUNT_SCORING.ZERO;
    } else if (count <= 2) {
      score = CONTACT_COUNT_SCORING.LOW;
    } else if (count <= 4) {
      score = CONTACT_COUNT_SCORING.MEDIUM;
    } else if (count <= 6) {
      score = CONTACT_COUNT_SCORING.HIGH;
    } else if (count <= 9) {
      score = CONTACT_COUNT_SCORING.VERY_HIGH;
    } else {
      score = CONTACT_COUNT_SCORING.EXCELLENT;
    }

    Logger.log(`✅ 接点回数: ${count}回 → スコア: ${score}点`);
    return score;

  } catch (error) {
    Logger.log(`❌ 接点回数スコアエラー: ${error}`);
    return DEFAULT_SCORES_PHASE3_2.CONTACT_COUNT;
  }
}

/**
 * 接点間隔スコアの計算
 *
 * @param {string} candidateId - 候補者ID
 * @return {number} 接点間隔スコア（0-100点）
 *
 * スコアリング基準（平均接点間隔）:
 * 0-7日: 100点（週1回以上）
 * 8-14日: 80点（2週間に1回）
 * 15-21日: 60点（3週間に1回）
 * 22-30日: 40点（月1回）
 * 31日以上: 20点（月1回未満）
 */
function calculateIntervalScore(candidateId) {
  try {
    const avgInterval = getAverageInterval(candidateId);

    if (avgInterval === 0) {
      Logger.log(`⚠️ 接点間隔データなし: ${candidateId} → デフォルト値`);
      return DEFAULT_SCORES_PHASE3_2.INTERVAL;
    }

    let score;
    if (avgInterval <= INTERVAL_SCORING.WEEKLY.days) {
      score = INTERVAL_SCORING.WEEKLY.score;
    } else if (avgInterval <= INTERVAL_SCORING.BIWEEKLY.days) {
      score = INTERVAL_SCORING.BIWEEKLY.score;
    } else if (avgInterval <= INTERVAL_SCORING.TRIWEEKLY.days) {
      score = INTERVAL_SCORING.TRIWEEKLY.score;
    } else if (avgInterval <= INTERVAL_SCORING.MONTHLY.days) {
      score = INTERVAL_SCORING.MONTHLY.score;
    } else {
      score = INTERVAL_SCORING.RARE.score;
    }

    Logger.log(`✅ 平均接点間隔: ${avgInterval.toFixed(1)}日 → スコア: ${score}点`);
    return score;

  } catch (error) {
    Logger.log(`❌ 接点間隔スコアエラー: ${error}`);
    return DEFAULT_SCORES_PHASE3_2.INTERVAL;
  }
}

/**
 * 接点の質スコアの計算
 *
 * @param {string} candidateId - 候補者ID
 * @return {number} 接点の質スコア（0-100点）
 *
 * 計算方法:
 * - 各アンケートの志望度の推移から計算
 * - 初回→社員→2次→内定の志望度変化を分析
 * - 上昇傾向: 高スコア
 * - 維持: 中スコア
 * - 下降傾向: 低スコア
 */
function calculateQualityScore(candidateId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const email = getCandidateEmail(candidateId);

    if (!email) {
      Logger.log(`❌ メールアドレスなし: ${candidateId}`);
      return DEFAULT_SCORES_PHASE3_2.QUALITY;
    }

    const phases = [
      { name: '初回面談', sheet: 'アンケート_初回面談', column: 'F' },
      { name: '社員面談', sheet: 'アンケート_社員面談', column: 'F' },
      { name: '2次面接', sheet: 'アンケート_2次面接', column: 'F' },
      { name: '内定後', sheet: 'アンケート_内定', column: 'H' }
    ];

    const motivationScores = [];

    for (let phase of phases) {
      const sheet = ss.getSheetByName(phase.sheet);
      if (!sheet) continue;

      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        if (data[i][2] === email) { // C列: メールアドレス
          const colIndex = phase.column.charCodeAt(0) - 65;
          const score = data[i][colIndex];

          if (score && typeof score === 'number') {
            motivationScores.push(score);
          }
          break;
        }
      }
    }

    if (motivationScores.length < 2) {
      Logger.log(`⚠️ 志望度データ不足: ${candidateId} → デフォルト値`);
      return DEFAULT_SCORES_PHASE3_2.QUALITY;
    }

    // 志望度の変化を計算
    let totalChange = 0;
    for (let i = 1; i < motivationScores.length; i++) {
      totalChange += (motivationScores[i] - motivationScores[i - 1]);
    }

    const avgChange = totalChange / (motivationScores.length - 1);

    // スコアリング
    let score;
    if (avgChange >= MOTIVATION_CHANGE_SCORING.LARGE_INCREASE.change) {
      score = MOTIVATION_CHANGE_SCORING.LARGE_INCREASE.score;
    } else if (avgChange >= MOTIVATION_CHANGE_SCORING.INCREASE.change) {
      score = MOTIVATION_CHANGE_SCORING.INCREASE.score;
    } else if (avgChange >= MOTIVATION_CHANGE_SCORING.STABLE.change) {
      score = MOTIVATION_CHANGE_SCORING.STABLE.score;
    } else if (avgChange >= MOTIVATION_CHANGE_SCORING.SLIGHT_DECREASE.change) {
      score = MOTIVATION_CHANGE_SCORING.SLIGHT_DECREASE.score;
    } else {
      score = MOTIVATION_CHANGE_SCORING.DECREASE.score;
    }

    Logger.log(`✅ 志望度推移: ${motivationScores.join(' → ')} (平均変化: ${avgChange.toFixed(1)}) → スコア: ${score}点`);
    return score;

  } catch (error) {
    Logger.log(`❌ 接点の質スコアエラー: ${error}`);
    return DEFAULT_SCORES_PHASE3_2.QUALITY;
  }
}

/**
 * 空白期間スコアの計算
 *
 * @param {string} candidateId - 候補者ID
 * @return {number} 空白期間スコア（0-100点）
 *
 * スコアリング基準（最新接点からの経過日数）:
 * 0-7日: 100点
 * 8-14日: 80点
 * 15-21日: 60点
 * 22-30日: 40点
 * 31日以上: 20点
 */
function calculateGapScore(candidateId) {
  try {
    const latestContact = getLatestContactDate(candidateId);

    if (!latestContact) {
      Logger.log(`⚠️ 接点履歴なし: ${candidateId} → デフォルト値`);
      return DEFAULT_SCORES_PHASE3_2.GAP;
    }

    const now = new Date();
    const gapDays = (now - latestContact) / (1000 * 60 * 60 * 24);

    let score;
    if (gapDays <= GAP_SCORING.RECENT.days) {
      score = GAP_SCORING.RECENT.score;
    } else if (gapDays <= GAP_SCORING.FAIRLY_RECENT.days) {
      score = GAP_SCORING.FAIRLY_RECENT.score;
    } else if (gapDays <= GAP_SCORING.MODERATE.days) {
      score = GAP_SCORING.MODERATE.score;
    } else if (gapDays <= GAP_SCORING.OLD.days) {
      score = GAP_SCORING.OLD.score;
    } else {
      score = GAP_SCORING.VERY_OLD.score;
    }

    Logger.log(`✅ 空白期間: ${gapDays.toFixed(1)}日 → スコア: ${score}点`);
    return score;

  } catch (error) {
    Logger.log(`❌ 空白期間スコアエラー: ${error}`);
    return DEFAULT_SCORES_PHASE3_2.GAP;
  }
}

/**
 * ========================================
 * Phase 3-2: 行動シグナル要素の計算関数
 * ========================================
 */

/**
 * 行動シグナル要素スコアの計算（統合）
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {number} 行動シグナル要素スコア（0-100点）
 *
 * 計算式:
 * 行動シグナル = 回答速度(30%) + 自由記述(20%) + 選考スピード(25%) + 積極性(25%)
 */
function calculateBehaviorScore(candidateId, phase) {
  try {
    Logger.log(`\n========================================`);
    Logger.log(`行動シグナル要素スコア計算: ${candidateId}, ${phase}`);
    Logger.log(`========================================`);

    // 回答速度スコア（既存のSurvey_Analysisから取得）
    const responseSpeedScore = getAverageResponseSpeedScore(candidateId) || DEFAULT_SCORES_PHASE3_2.RESPONSE_SPEED;

    // 自由記述スコア
    const descriptionScore = calculateDescriptionScore(candidateId, phase);

    // 選考スピードスコア
    const selectionSpeedScore = calculateSelectionSpeedScore(candidateId);

    // 積極性スコア（既存のEvaluation_LogのT列から取得）
    const proactivityScore = getProactivityScore(candidateId, phase) || DEFAULT_SCORES_PHASE3_2.PROACTIVITY;

    const behaviorScore =
      responseSpeedScore * 0.3 +
      descriptionScore * 0.2 +
      selectionSpeedScore * 0.25 +
      proactivityScore * 0.25;

    Logger.log(`\n--- 行動シグナル要素スコア内訳 ---`);
    Logger.log(`回答速度スコア: ${responseSpeedScore} × 0.3 = ${(responseSpeedScore * 0.3).toFixed(1)}`);
    Logger.log(`自由記述スコア: ${descriptionScore} × 0.2 = ${(descriptionScore * 0.2).toFixed(1)}`);
    Logger.log(`選考スピードスコア: ${selectionSpeedScore} × 0.25 = ${(selectionSpeedScore * 0.25).toFixed(1)}`);
    Logger.log(`積極性スコア: ${proactivityScore} × 0.25 = ${(proactivityScore * 0.25).toFixed(1)}`);
    Logger.log(`行動シグナル要素スコア（合計）: ${behaviorScore.toFixed(1)}`);
    Logger.log(`========================================\n`);

    return Math.round(behaviorScore);

  } catch (error) {
    Logger.log(`❌ 行動シグナル要素スコアエラー: ${error}`);
    return DEFAULT_SCORES_PHASE3_2.BEHAVIOR;
  }
}

/**
 * 自由記述スコアの計算
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {number} 自由記述スコア（0-100点）
 *
 * スコアリング基準（文字数）:
 * 0文字: 0点
 * 1-20文字: 20点（ほとんど書いていない）
 * 21-50文字: 40点（短い）
 * 51-100文字: 60点（普通）
 * 101-200文字: 80点（詳しい）
 * 201文字以上: 100点（非常に詳しい）
 */
function calculateDescriptionScore(candidateId, phase) {
  try {
    const freeText = getFreeTextResponses(candidateId, phase);
    const length = freeText.length;

    let score;
    if (length === DESCRIPTION_SCORING.NONE.length) {
      score = DESCRIPTION_SCORING.NONE.score;
    } else if (length <= DESCRIPTION_SCORING.VERY_SHORT.length) {
      score = DESCRIPTION_SCORING.VERY_SHORT.score;
    } else if (length <= DESCRIPTION_SCORING.SHORT.length) {
      score = DESCRIPTION_SCORING.SHORT.score;
    } else if (length <= DESCRIPTION_SCORING.MEDIUM.length) {
      score = DESCRIPTION_SCORING.MEDIUM.score;
    } else if (length <= DESCRIPTION_SCORING.LONG.length) {
      score = DESCRIPTION_SCORING.LONG.score;
    } else {
      score = DESCRIPTION_SCORING.VERY_LONG.score;
    }

    Logger.log(`✅ 自由記述: ${length}文字 → スコア: ${score}点`);
    return score;

  } catch (error) {
    Logger.log(`❌ 自由記述スコアエラー: ${error}`);
    return DEFAULT_SCORES_PHASE3_2.DESCRIPTION;
  }
}

/**
 * 選考スピードスコアの計算
 *
 * @param {string} candidateId - 候補者ID
 * @return {number} 選考スピードスコア（0-100点）
 *
 * スコアリング基準（応募から現在までの日数）:
 * 0-14日: 100点（非常に速い）
 * 15-30日: 90点（速い）
 * 31-60日: 80点（普通）
 * 61-90日: 60点（やや遅い）
 * 91日以上: 40点（遅い）
 */
function calculateSelectionSpeedScore(candidateId) {
  try {
    const duration = getSelectionDuration(candidateId);

    if (duration === 0) {
      Logger.log(`⚠️ 選考期間データなし: ${candidateId} → デフォルト値`);
      return DEFAULT_SCORES_PHASE3_2.SELECTION_SPEED;
    }

    let score;
    if (duration <= SELECTION_SPEED_SCORING.VERY_FAST.days) {
      score = SELECTION_SPEED_SCORING.VERY_FAST.score;
    } else if (duration <= SELECTION_SPEED_SCORING.FAST.days) {
      score = SELECTION_SPEED_SCORING.FAST.score;
    } else if (duration <= SELECTION_SPEED_SCORING.NORMAL.days) {
      score = SELECTION_SPEED_SCORING.NORMAL.score;
    } else if (duration <= SELECTION_SPEED_SCORING.SLOW.days) {
      score = SELECTION_SPEED_SCORING.SLOW.score;
    } else {
      score = SELECTION_SPEED_SCORING.VERY_SLOW.score;
    }

    Logger.log(`✅ 選考期間: ${duration.toFixed(1)}日 → スコア: ${score}点`);
    return score;

  } catch (error) {
    Logger.log(`❌ 選考スピードスコアエラー: ${error}`);
    return DEFAULT_SCORES_PHASE3_2.SELECTION_SPEED;
  }
}

/**
 * ========================================
 * Phase 3-3: 最終統合・Engagement_Log書き込み
 * ========================================
 */

// Phase 3-3のデフォルト値
const DEFAULT_SCORES_PHASE3_3 = {
  SELF_REPORT: 70,           // 自己申告要素スコア
  ACCEPTANCE_RATE: 50        // 承諾可能性（最終統合）
};

// 承諾可能性の重み付け（初回面談・社員面談）
const ACCEPTANCE_RATE_WEIGHTS_EARLY = {
  FOUNDATION: 0.4,           // 基礎要素 40%
  RELATIONSHIP: 0.3,         // 関係性要素 30%
  BEHAVIOR: 0.3              // 行動シグナル要素 30%
};

// 承諾可能性の重み付け（2次面接・内定後）
const ACCEPTANCE_RATE_WEIGHTS_LATE = {
  FOUNDATION: 0.4,           // 基礎要素 40%
  RELATIONSHIP: 0.3,         // 関係性要素 30%
  BEHAVIOR: 0.2,             // 行動シグナル要素 20%
  SELF_REPORT: 0.1           // 自己申告要素 10%
};

// Engagement_Logの列マッピング（appendRow用の配列インデックス）
const ENGAGEMENT_LOG_COLUMNS = {
  ENGAGEMENT_ID: 0,          // A列
  CANDIDATE_ID: 1,           // B列
  CANDIDATE_NAME: 2,         // C列
  ENGAGEMENT_DATE: 3,        // D列
  PHASE: 4,                  // E列
  AI_PREDICTION: 5,          // F列: AI予測_承諾可能性
  HUMAN_INTUITION: 6,        // G列: 人間の直感_承諾可能性
  INTEGRATED: 7,             // H列: 統合_承諾可能性
  CONFIDENCE: 8,             // I列: 信頼度
  MOTIVATION_SCORE: 9,       // J列: 志望度スコア
  COMPETITIVE_SCORE: 10,     // K列: 競合優位性スコア
  CONCERN_SCORE: 11,         // L列: 懸念解消度スコア
  CORE_MOTIVATION: 12,       // M列: コアモチベーション
  TOP_CONCERN: 13            // N列: 主要懸念事項
};

// Candidates_Masterの列番号（Y列・Z列）
const CANDIDATES_MASTER_EXTENDED_COLUMNS = {
  CORE_MOTIVATION: 24,       // Y列（A=0, Y=24）
  TOP_CONCERN: 25            // Z列（A=0, Z=25）
};

/**
 * 自己申告要素スコアの計算（2次面接・内定後のみ）
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {number|null} 自己申告要素スコア（0-100点）、または null（初回面談・社員面談）
 *
 * データ取得元:
 * - 2次面接: M列（Q12. 承諾可能性、1-10）
 * - 内定後: H列（Q7. PIGNUSで働くことへの前向きさ、1-10）
 *
 * スコアリング:
 * 自己申告スコア = (回答値 × 10)
 *
 * 戻り値:
 * - 初回面談・社員面談: null（自己申告要素なし）
 * - 2次面接・内定後: 0-100点
 * - データなし: 70点（デフォルト値）
 */
function calculateSelfReportScore(candidateId, phase) {
  try {
    // 初回面談・社員面談では自己申告要素なし
    if (phase === '初回面談' || phase === '社員面談') {
      Logger.log(`⚠️ 自己申告要素なし: ${phase} → null`);
      return null; // nullを返す
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const email = getCandidateEmail(candidateId);

    if (!email) {
      Logger.log(`❌ メールアドレスなし: ${candidateId}`);
      return DEFAULT_SCORES_PHASE3_3.SELF_REPORT;
    }

    let sheetName, column;
    if (phase === '2次面接') {
      sheetName = 'アンケート_2次面接';
      column = 'M'; // Q12. 承諾可能性（要確認）
    } else if (phase === '内定後') {
      sheetName = 'アンケート_内定';
      column = 'H'; // Q7. PIGNUSで働くことへの前向きさ
    } else {
      Logger.log(`⚠️ 不明なフェーズ: ${phase}`);
      return null;
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`❌ シートが見つかりません: ${sheetName}`);
      return DEFAULT_SCORES_PHASE3_3.SELF_REPORT;
    }

    const data = sheet.getDataRange().getValues();

    // ヘッダー行をスキップ（i=1から開始）
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === email) { // C列: メールアドレス
        const colIndex = column.charCodeAt(0) - 65;
        const selfReportRaw = data[i][colIndex];

        if (selfReportRaw && typeof selfReportRaw === 'number') {
          // 1-10を10-100に変換
          const score = selfReportRaw * 10;
          Logger.log(`✅ 自己申告スコア: ${score}点（生値: ${selfReportRaw}）`);
          return Math.min(Math.max(Math.round(score), 0), 100);
        }

        Logger.log(`⚠️ 自己申告データなし: ${candidateId}, ${phase} → デフォルト値`);
        return DEFAULT_SCORES_PHASE3_3.SELF_REPORT;
      }
    }

    Logger.log(`⚠️ アンケート回答なし: ${candidateId}, ${phase} → デフォルト値`);
    return DEFAULT_SCORES_PHASE3_3.SELF_REPORT;

  } catch (error) {
    Logger.log(`❌ 自己申告スコアエラー: ${error}`);
    return DEFAULT_SCORES_PHASE3_3.SELF_REPORT;
  }
}

/**
 * 承諾可能性の計算（最終統合）
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {number} 承諾可能性（0-100点）
 *
 * 計算式:
 * 【初回面談・社員面談】
 * 承諾可能性 = 基礎要素(40%) + 関係性要素(30%) + 行動シグナル要素(30%)
 *
 * 【2次面接・内定後】
 * 承諾可能性 = 基礎要素(40%) + 関係性要素(30%) + 行動シグナル要素(20%) + 自己申告要素(10%)
 */
function calculateAcceptanceRate(candidateId, phase) {
  try {
    Logger.log(`\n========================================`);
    Logger.log(`承諾可能性計算: ${candidateId}, ${phase}`);
    Logger.log(`========================================`);

    // Phase 3-1: 基礎要素（40%）
    const foundationScore = calculateFoundationScore(candidateId, phase);

    // Phase 3-2: 関係性要素（30%）
    const relationshipScore = calculateRelationshipScore(candidateId);

    // Phase 3-2: 行動シグナル要素
    const behaviorScore = calculateBehaviorScore(candidateId, phase);

    // Phase 3-3: 自己申告要素（2次面接・内定後のみ）
    const selfReportScore = calculateSelfReportScore(candidateId, phase);

    let acceptanceRate;

    if (selfReportScore === null) {
      // 初回面談・社員面談: 行動シグナル30%
      acceptanceRate =
        foundationScore * ACCEPTANCE_RATE_WEIGHTS_EARLY.FOUNDATION +
        relationshipScore * ACCEPTANCE_RATE_WEIGHTS_EARLY.RELATIONSHIP +
        behaviorScore * ACCEPTANCE_RATE_WEIGHTS_EARLY.BEHAVIOR;

      Logger.log(`\n--- 承諾可能性内訳（初回面談・社員面談）---`);
      Logger.log(`基礎要素: ${foundationScore} × 0.4 = ${(foundationScore * 0.4).toFixed(1)}`);
      Logger.log(`関係性要素: ${relationshipScore} × 0.3 = ${(relationshipScore * 0.3).toFixed(1)}`);
      Logger.log(`行動シグナル要素: ${behaviorScore} × 0.3 = ${(behaviorScore * 0.3).toFixed(1)}`);

    } else {
      // 2次面接・内定後: 行動シグナル20% + 自己申告10%
      acceptanceRate =
        foundationScore * ACCEPTANCE_RATE_WEIGHTS_LATE.FOUNDATION +
        relationshipScore * ACCEPTANCE_RATE_WEIGHTS_LATE.RELATIONSHIP +
        behaviorScore * ACCEPTANCE_RATE_WEIGHTS_LATE.BEHAVIOR +
        selfReportScore * ACCEPTANCE_RATE_WEIGHTS_LATE.SELF_REPORT;

      Logger.log(`\n--- 承諾可能性内訳（2次面接・内定後）---`);
      Logger.log(`基礎要素: ${foundationScore} × 0.4 = ${(foundationScore * 0.4).toFixed(1)}`);
      Logger.log(`関係性要素: ${relationshipScore} × 0.3 = ${(relationshipScore * 0.3).toFixed(1)}`);
      Logger.log(`行動シグナル要素: ${behaviorScore} × 0.2 = ${(behaviorScore * 0.2).toFixed(1)}`);
      Logger.log(`自己申告要素: ${selfReportScore} × 0.1 = ${(selfReportScore * 0.1).toFixed(1)}`);
    }

    Logger.log(`承諾可能性（合計）: ${acceptanceRate.toFixed(1)}`);
    Logger.log(`========================================\n`);

    return Math.round(acceptanceRate); // 四捨五入

  } catch (error) {
    Logger.log(`❌ 承諾可能性計算エラー: ${error}`);
    return DEFAULT_SCORES_PHASE3_3.ACCEPTANCE_RATE;
  }
}

/**
 * コアモチベーションを取得
 *
 * @param {string} candidateId - 候補者ID
 * @return {string} コアモチベーション
 *
 * データ取得元:
 * - Candidates_Master: Y列（コアモチベーション）
 *   インデックス: 24（A=0, Y=24）
 */
function getCoreMotivation(candidateId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);

    if (!master) {
      Logger.log('❌ Candidates_Masterシートが見つかりません');
      return '不明';
    }

    const data = master.getDataRange().getValues();

    // ヘッダー行をスキップ（i=1から開始）
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === candidateId) { // A列: candidate_id
        const coreMotivation = data[i][CANDIDATES_MASTER_EXTENDED_COLUMNS.CORE_MOTIVATION];

        if (coreMotivation && coreMotivation !== '' && coreMotivation !== 'undefined') {
          Logger.log(`✅ コアモチベーション: ${candidateId} → ${coreMotivation}`);
          return String(coreMotivation);
        }

        Logger.log(`⚠️ コアモチベーションが空: ${candidateId} → 不明`);
        return '不明';
      }
    }

    Logger.log(`❌ 候補者が見つかりません: ${candidateId}`);
    return '不明';

  } catch (error) {
    Logger.log(`❌ コアモチベーション取得エラー: ${error}`);
    return '不明';
  }
}

/**
 * 主要懸念事項を取得
 *
 * @param {string} candidateId - 候補者ID
 * @return {string} 主要懸念事項
 *
 * データ取得元:
 * - Candidates_Master: Z列（主要懸念事項）
 *   インデックス: 25（A=0, Z=25）
 */
function getTopConcern(candidateId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);

    if (!master) {
      Logger.log('❌ Candidates_Masterシートが見つかりません');
      return 'なし';
    }

    const data = master.getDataRange().getValues();

    // ヘッダー行をスキップ（i=1から開始）
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === candidateId) { // A列: candidate_id
        const topConcern = data[i][CANDIDATES_MASTER_EXTENDED_COLUMNS.TOP_CONCERN];

        if (topConcern && topConcern !== '' && topConcern !== 'undefined') {
          Logger.log(`✅ 主要懸念事項: ${candidateId} → ${topConcern}`);
          return String(topConcern);
        }

        Logger.log(`⚠️ 主要懸念事項が空: ${candidateId} → なし`);
        return 'なし';
      }
    }

    Logger.log(`❌ 候補者が見つかりません: ${candidateId}`);
    return 'なし';

  } catch (error) {
    Logger.log(`❌ 主要懸念事項取得エラー: ${error}`);
    return 'なし';
  }
}

/**
 * Engagement_Logに承諾可能性を書き込む
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {boolean} 成功/失敗
 *
 * Engagement_Logの列構成（appendRowで追加）:
 * A: engagement_id
 * B: candidate_id
 * C: 氏名
 * D: engagement_date
 * E: 選考フェーズ
 * F: AI予測_承諾可能性
 * G: 人間の直感_承諾可能性
 * H: 統合_承諾可能性
 * I: 信頼度
 * J: 志望度スコア
 * K: 競合優位性スコア
 * L: 懸念解消度スコア
 * M: コアモチベーション
 * N: 主要懸念事項
 */
function writeToEngagementLog(candidateId, phase) {
  try {
    Logger.log(`\n========================================`);
    Logger.log(`Engagement_Log書き込み: ${candidateId}, ${phase}`);
    Logger.log(`========================================`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const engagementSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ENGAGEMENT_LOG);

    if (!engagementSheet) {
      Logger.log('❌ Engagement_Logシートが見つかりません');
      Logger.log('⚠️ Engagement_Logシートを作成してください');
      return false;
    }

    // 承諾可能性を計算
    const acceptanceRate = calculateAcceptanceRate(candidateId, phase);

    // 候補者情報を取得
    const master = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);
    if (!master) {
      Logger.log('❌ Candidates_Masterシートが見つかりません');
      return false;
    }

    const masterData = master.getDataRange().getValues();
    let candidateName = '';

    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][0] === candidateId) {
        candidateName = masterData[i][1]; // B列: 氏名
        break;
      }
    }

    if (!candidateName) {
      Logger.log(`❌ 候補者名が見つかりません: ${candidateId}`);
      return false;
    }

    // 各要素スコアを取得
    const motivationScore = calculateMotivationScore(candidateId, phase);
    const competitiveScore = calculateCompetitiveAdvantageScore(candidateId, phase);
    const concernScore = calculateConcernResolutionScore(candidateId, phase);

    // コアモチベーションと主要懸念を取得
    const coreMotivation = getCoreMotivation(candidateId);
    const topConcern = getTopConcern(candidateId);

    // engagement_idを生成
    const timestamp = new Date().getTime();
    const engagementId = `ENG-${candidateId}-${timestamp}`;

    // 新規行を作成
    const newRow = [];
    newRow[ENGAGEMENT_LOG_COLUMNS.ENGAGEMENT_ID] = engagementId;
    newRow[ENGAGEMENT_LOG_COLUMNS.CANDIDATE_ID] = candidateId;
    newRow[ENGAGEMENT_LOG_COLUMNS.CANDIDATE_NAME] = candidateName;
    newRow[ENGAGEMENT_LOG_COLUMNS.ENGAGEMENT_DATE] = new Date();
    newRow[ENGAGEMENT_LOG_COLUMNS.PHASE] = phase;
    newRow[ENGAGEMENT_LOG_COLUMNS.AI_PREDICTION] = acceptanceRate;
    newRow[ENGAGEMENT_LOG_COLUMNS.HUMAN_INTUITION] = ''; // 空白
    newRow[ENGAGEMENT_LOG_COLUMNS.INTEGRATED] = acceptanceRate; // AI予測と同じ
    newRow[ENGAGEMENT_LOG_COLUMNS.CONFIDENCE] = '高'; // 固定値
    newRow[ENGAGEMENT_LOG_COLUMNS.MOTIVATION_SCORE] = motivationScore;
    newRow[ENGAGEMENT_LOG_COLUMNS.COMPETITIVE_SCORE] = competitiveScore;
    newRow[ENGAGEMENT_LOG_COLUMNS.CONCERN_SCORE] = concernScore;
    newRow[ENGAGEMENT_LOG_COLUMNS.CORE_MOTIVATION] = coreMotivation;
    newRow[ENGAGEMENT_LOG_COLUMNS.TOP_CONCERN] = topConcern;

    // 行を追加
    engagementSheet.appendRow(newRow);

    Logger.log(`✅ Engagement_Logに書き込み完了`);
    Logger.log(`  - engagement_id: ${engagementId}`);
    Logger.log(`  - 承諾可能性: ${acceptanceRate}点`);
    Logger.log(`  - コアモチベーション: ${coreMotivation}`);
    Logger.log(`  - 主要懸念事項: ${topConcern}`);
    Logger.log(`========================================\n`);

    return true;

  } catch (error) {
    Logger.log(`❌ Engagement_Log書き込みエラー: ${error}`);
    return false;
  }
}

/**
 * ========================================
 * Phase 3-3: テスト関数
 * ========================================
 */

/**
 * Phase 3-3の依存関数の確認
 */
function checkPhase33Dependencies() {
  Logger.log('\n=== Phase 3-3依存関数の確認 ===');

  try {
    // Phase 3-1の関数
    const foundation = calculateFoundationScore('C001', '初回面談');
    Logger.log(`✅ calculateFoundationScore(): ${foundation}`);

    // Phase 3-2の関数
    const relationship = calculateRelationshipScore('C001');
    Logger.log(`✅ calculateRelationshipScore(): ${relationship}`);

    const behavior = calculateBehaviorScore('C001', '初回面談');
    Logger.log(`✅ calculateBehaviorScore(): ${behavior}`);

    Logger.log('\n✅ 全ての依存関数が正常に動作しています');

  } catch (error) {
    Logger.log(`\n❌ エラー: ${error}`);
    Logger.log('⚠️ Phase 3-1またはPhase 3-2が正しく実装されていません');
  }
}

/**
 * Candidates_Masterの列構造確認
 */
function checkCandidatesMasterColumns() {
  Logger.log('\n=== Candidates_Master列構造の確認 ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);

  if (!master) {
    Logger.log('❌ Candidates_Masterシートが見つかりません');
    return;
  }

  const headers = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0];

  // Y列（インデックス24）を確認
  Logger.log(`Y列（インデックス24）: ${headers[24]}`);
  // 期待値: "コアモチベーション" または類似の列名

  // Z列（インデックス25）を確認
  Logger.log(`Z列（インデックス25）: ${headers[25]}`);
  // 期待値: "主要懸念事項" または類似の列名

  Logger.log('\n⚠️ 上記の列名を確認してください');
  Logger.log('⚠️ 列名が異なる場合、getCoreMotivation()とgetTopConcern()の列番号を修正してください');
}

/**
 * Engagement_Logの列構造確認
 */
function checkEngagementLogStructure() {
  Logger.log('\n=== Engagement_Log列構造の確認 ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const engagementLog = ss.getSheetByName(CONFIG.SHEET_NAMES.ENGAGEMENT_LOG);

  if (!engagementLog) {
    Logger.log('❌ Engagement_Logシートが見つかりません');
    Logger.log('⚠️ Engagement_Logシートを作成する必要があります');
    return;
  }

  const headers = engagementLog.getRange(1, 1, 1, engagementLog.getLastColumn()).getValues()[0];

  Logger.log('\nEngagement_Logの列構造:');
  for (let i = 0; i < Math.min(headers.length, 20); i++) {
    const columnLetter = String.fromCharCode(65 + i);
    Logger.log(`${columnLetter}列（インデックス${i}）: ${headers[i]}`);
  }

  Logger.log('\n⚠️ 上記の列構造を確認してください');
  Logger.log('⚠️ 列構造が異なる場合、writeToEngagementLog()の列マッピングを修正してください');
}

/**
 * アンケートシートの自己申告列確認
 */
function checkSurveyColumnsForSelfReport() {
  Logger.log('\n=== アンケートシートの自己申告列確認 ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 2次面接
  const interview2 = ss.getSheetByName('アンケート_2次面接');
  if (interview2) {
    const headers1 = interview2.getRange(1, 1, 1, interview2.getLastColumn()).getValues()[0];
    Logger.log(`\n2次面接 M列（インデックス12）: ${headers1[12]}`);
    // 期待値: Q12またはQ13（承諾可能性、1-10）
  }

  // 内定後
  const offer = ss.getSheetByName('アンケート_内定');
  if (offer) {
    const headers2 = offer.getRange(1, 1, 1, offer.getLastColumn()).getValues()[0];
    Logger.log(`内定後 H列（インデックス7）: ${headers2[7]}`);
    // 期待値: Q7（PIGNUSで働くことへの前向きさ、1-10）
  }

  Logger.log('\n⚠️ 上記の列名を確認してください');
  Logger.log('⚠️ 列名が異なる場合、calculateSelfReportScore()の列番号を修正してください');
}

/**
 * Phase 3-3の事前確認テスト
 */
function runPhase33PreChecks() {
  Logger.log('\n========================================');
  Logger.log('Phase 3-3 事前確認テスト');
  Logger.log('========================================\n');

  // 依存関数の確認
  checkPhase33Dependencies();

  // Candidates_Masterの列構造確認
  checkCandidatesMasterColumns();

  // Engagement_Logの列構造確認
  checkEngagementLogStructure();

  // アンケートシートの列構造確認
  checkSurveyColumnsForSelfReport();

  Logger.log('\n========================================');
  Logger.log('事前確認テスト完了');
  Logger.log('========================================\n');
}

/**
 * 自己申告要素スコアのテスト
 */
function testSelfReportScore() {
  Logger.log('\n=== 自己申告要素スコアのテスト ===');

  const candidates = ['C001', 'C002', 'C003'];

  for (let candidate of candidates) {
    Logger.log(`\n--- ${candidate} ---`);

    // 初回面談（自己申告なし）
    const score1 = calculateSelfReportScore(candidate, '初回面談');
    Logger.log(`初回面談: ${score1} (期待値: null)`);

    // 社員面談（自己申告なし）
    const score2 = calculateSelfReportScore(candidate, '社員面談');
    Logger.log(`社員面談: ${score2} (期待値: null)`);

    // 2次面接（自己申告あり）
    const score3 = calculateSelfReportScore(candidate, '2次面接');
    Logger.log(`2次面接: ${score3}点`);

    // 内定後（自己申告あり）
    const score4 = calculateSelfReportScore(candidate, '内定後');
    Logger.log(`内定後: ${score4}点`);
  }
}

/**
 * 承諾可能性計算のテスト
 */
function testAcceptanceRate() {
  Logger.log('\n=== 承諾可能性計算のテスト ===');

  const candidates = ['C001']; // まずはC001のみ
  const phases = ['初回面談', '社員面談', '2次面接', '内定後'];

  for (let candidate of candidates) {
    Logger.log(`\n--- ${candidate} ---`);
    for (let phase of phases) {
      const rate = calculateAcceptanceRate(candidate, phase);
      Logger.log(`${phase}: ${rate}点`);
    }
  }
}

/**
 * 補助関数のテスト
 */
function testHelperFunctionsPhase33() {
  Logger.log('\n=== 補助関数のテスト ===');

  const candidates = ['C001', 'C002', 'C003'];

  for (let candidate of candidates) {
    Logger.log(`\n--- ${candidate} ---`);

    const coreMotivation = getCoreMotivation(candidate);
    Logger.log(`コアモチベーション: ${coreMotivation}`);

    const topConcern = getTopConcern(candidate);
    Logger.log(`主要懸念事項: ${topConcern}`);
  }
}

/**
 * Engagement_Log書き込みのテスト
 */
function testWriteToEngagementLog() {
  Logger.log('\n=== Engagement_Log書き込みのテスト ===');

  // C001の初回面談を書き込み
  const result1 = writeToEngagementLog('C001', '初回面談');
  Logger.log(`\nC001-初回面談: ${result1 ? '✅ 成功' : '❌ 失敗'}`);

  // C001の2次面接を書き込み
  const result2 = writeToEngagementLog('C001', '2次面接');
  Logger.log(`C001-2次面接: ${result2 ? '✅ 成功' : '❌ 失敗'}`);

  // C001の内定後を書き込み
  const result3 = writeToEngagementLog('C001', '内定後');
  Logger.log(`C001-内定後: ${result3 ? '✅ 成功' : '❌ 失敗'}`);

  Logger.log('\n⚠️ Engagement_Logシートを確認して、3件のデータが追加されていることを確認してください');
}

/**
 * Phase 3-3の全テストを実行
 */
function runAllPhase33Tests() {
  Logger.log('\n========================================');
  Logger.log('Phase 3-3 全テスト実行');
  Logger.log('========================================\n');

  // 事前確認
  Logger.log('>>> 事前確認テスト');
  runPhase33PreChecks();

  // 単体テスト
  Logger.log('\n>>> 単体テスト');
  testSelfReportScore();
  testAcceptanceRate();
  testHelperFunctionsPhase33();

  // Engagement_Log書き込み
  Logger.log('\n>>> Engagement_Log書き込みテスト');
  testWriteToEngagementLog();

  Logger.log('\n========================================');
  Logger.log('Phase 3-3 全テスト完了');
  Logger.log('========================================\n');
}
