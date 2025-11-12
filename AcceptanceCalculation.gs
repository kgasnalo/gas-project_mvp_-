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
