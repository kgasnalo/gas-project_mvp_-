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
