/**
 * DataGetters.gs
 * データ取得用の共通関数
 *
 * 【主要機能】
 * - 各シートからデータを取得する汎用関数
 * - Phase 2以降の機能で使用される
 *
 * @version 1.0
 * @date 2025-11-12
 */

/**
 * 最新の評価ログを取得
 *
 * @param {string} candidateId - 候補者ID（例: 'C001'）
 * @return {Object|null} 評価ログオブジェクト、見つからない場合はnull
 *
 * @example
 * const evalLog = getLatestEvaluationLog('C001');
 * if (evalLog) {
 *   console.log(evalLog.overall_score);
 * }
 */
function getLatestEvaluationLog(candidateId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EVALUATION_LOG);

    if (!sheet) {
      Logger.log('⚠️ Evaluation_Logシートが見つかりません');
      return null;
    }

    const data = sheet.getDataRange().getValues();

    // 最新の評価ログを探す（日付降順）
    let latestLog = null;
    let latestDate = null;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCandidateId = row[CONFIG.COLUMNS.EVALUATION_LOG.CANDIDATE_ID];

      if (rowCandidateId === candidateId) {
        const evalDate = new Date(row[CONFIG.COLUMNS.EVALUATION_LOG.EVAL_DATE]);

        if (!latestDate || evalDate > latestDate) {
          latestDate = evalDate;
          latestLog = {
            log_id: row[CONFIG.COLUMNS.EVALUATION_LOG.LOG_ID],
            candidate_id: row[CONFIG.COLUMNS.EVALUATION_LOG.CANDIDATE_ID],
            name: row[CONFIG.COLUMNS.EVALUATION_LOG.NAME],
            eval_date: evalDate,
            phase: row[CONFIG.COLUMNS.EVALUATION_LOG.PHASE],
            interviewer: row[CONFIG.COLUMNS.EVALUATION_LOG.INTERVIEWER],
            overall_score: row[CONFIG.COLUMNS.EVALUATION_LOG.OVERALL_SCORE],
            philosophy_score: row[CONFIG.COLUMNS.EVALUATION_LOG.PHILOSOPHY_SCORE],
            strategy_score: row[CONFIG.COLUMNS.EVALUATION_LOG.STRATEGY_SCORE],
            motivation_score: row[CONFIG.COLUMNS.EVALUATION_LOG.MOTIVATION_SCORE],
            execution_score: row[CONFIG.COLUMNS.EVALUATION_LOG.EXECUTION_SCORE],
            pass_probability: row[CONFIG.COLUMNS.EVALUATION_LOG.PASS_PROBABILITY],
            recommendation: row[CONFIG.COLUMNS.EVALUATION_LOG.RECOMMENDATION],
            proactivity_score: row[CONFIG.COLUMNS.EVALUATION_LOG.PROACTIVITY_SCORE]
          };
        }
      }
    }

    return latestLog;

  } catch (error) {
    logError('getLatestEvaluationLog', error);
    return null;
  }
}

/**
 * 候補者一覧を取得
 *
 * @param {string|null} filterStatus - フィルタするステータス（nullの場合は全候補者）
 * @return {Array<Object>} 候補者オブジェクトの配列
 *
 * @example
 * // 全候補者を取得
 * const allCandidates = getAllCandidates(null);
 *
 * // 「1次面接」の候補者のみ取得
 * const interviewCandidates = getAllCandidates('1次面接');
 */
function getAllCandidates(filterStatus = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);

    if (!sheet) {
      Logger.log('⚠️ Candidates_Masterシートが見つかりません');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const candidates = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const candidateId = row[CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID];

      // 空行はスキップ
      if (!candidateId) continue;

      const status = row[CONFIG.COLUMNS.CANDIDATES_MASTER.STATUS];

      // ステータスフィルタ
      if (filterStatus && status !== filterStatus) {
        continue;
      }

      candidates.push({
        candidate_id: candidateId,
        name: row[CONFIG.COLUMNS.CANDIDATES_MASTER.NAME],
        status: status,
        category: row[CONFIG.COLUMNS.CANDIDATES_MASTER.CATEGORY],
        email: row[CONFIG.COLUMNS.CANDIDATES_MASTER.EMAIL],
        latest_pass_prob: row[CONFIG.COLUMNS.CANDIDATES_MASTER.LATEST_PASS_PROB],
        latest_acceptance_integrated: row[CONFIG.COLUMNS.CANDIDATES_MASTER.LATEST_ACCEPTANCE_INTEGRATED]
      });
    }

    return candidates;

  } catch (error) {
    logError('getAllCandidates', error);
    return [];
  }
}

/**
 * アンケート送信ログを取得
 *
 * @param {string} candidateId - 候補者ID
 * @param {string|null} phase - アンケート種別（nullの場合は全種別）
 * @return {Array<Object>} 送信ログオブジェクトの配列
 *
 * @example
 * // 特定の候補者の「初回面談」アンケート送信ログを取得
 * const logs = getSurveySendLog('C001', '初回面談');
 */
function getSurveySendLog(candidateId, phase = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);

    if (!sheet) {
      Logger.log('⚠️ Survey_Send_Logシートが見つかりません');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const logs = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCandidateId = row[CONFIG.COLUMNS.SURVEY_SEND_LOG.CANDIDATE_ID];
      const rowPhase = row[CONFIG.COLUMNS.SURVEY_SEND_LOG.PHASE];

      // 候補者IDでフィルタ
      if (rowCandidateId !== candidateId) continue;

      // アンケート種別でフィルタ（指定がある場合）
      if (phase && rowPhase !== phase) continue;

      logs.push({
        send_id: row[CONFIG.COLUMNS.SURVEY_SEND_LOG.SEND_ID],
        candidate_id: rowCandidateId,
        name: row[CONFIG.COLUMNS.SURVEY_SEND_LOG.NAME],
        email: row[CONFIG.COLUMNS.SURVEY_SEND_LOG.EMAIL],
        phase: rowPhase,
        send_time: new Date(row[CONFIG.COLUMNS.SURVEY_SEND_LOG.SEND_TIME]),
        send_status: row[CONFIG.COLUMNS.SURVEY_SEND_LOG.STATUS],  // ✅ status → send_status に修正
        error_message: row[CONFIG.COLUMNS.SURVEY_SEND_LOG.ERROR_MSG]
      });
    }

    return logs;

  } catch (error) {
    logError('getSurveySendLog', error);
    return [];
  }
}

/**
 * アンケート回答を取得
 *
 * @param {string} candidateId - 候補者ID
 * @param {string|null} phase - アンケート種別（nullの場合は全種別）
 * @return {Array<Object>} アンケート回答オブジェクトの配列
 *
 * @example
 * // 特定の候補者の「初回面談」アンケート回答を取得
 * const responses = getSurveyResponse('C001', '初回面談');
 */
function getSurveyResponse(candidateId, phase = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_RESPONSE);

    if (!sheet) {
      Logger.log('⚠️ Survey_Responseシートが見つかりません');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const responses = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCandidateId = row[CONFIG.COLUMNS.SURVEY_RESPONSE.CANDIDATE_ID];
      const rowPhase = row[CONFIG.COLUMNS.SURVEY_RESPONSE.PHASE];

      // 候補者IDでフィルタ
      if (rowCandidateId !== candidateId) continue;

      // アンケート種別でフィルタ（指定がある場合）
      if (phase && rowPhase !== phase) continue;

      responses.push({
        response_id: row[CONFIG.COLUMNS.SURVEY_RESPONSE.RESPONSE_ID],
        candidate_id: rowCandidateId,
        name: row[CONFIG.COLUMNS.SURVEY_RESPONSE.NAME],
        response_date: new Date(row[CONFIG.COLUMNS.SURVEY_RESPONSE.RESPONSE_DATE]),
        phase: rowPhase,
        aspiration: row[CONFIG.COLUMNS.SURVEY_RESPONSE.ASPIRATION],
        concerns: row[CONFIG.COLUMNS.SURVEY_RESPONSE.CONCERNS],
        other_companies: row[CONFIG.COLUMNS.SURVEY_RESPONSE.OTHER_COMPANIES],
        comments: row[CONFIG.COLUMNS.SURVEY_RESPONSE.COMMENTS]
      });
    }

    return responses;

  } catch (error) {
    logError('getSurveyResponse', error);
    return [];
  }
}

/**
 * 候補者IDから最新のアンケート回答速度を計算
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {Object|null} { send_time, response_time, hours, score }
 *
 * @example
 * const speed = getResponseSpeedData('C001', '初回面談');
 * console.log(speed.hours); // 2.5時間
 * console.log(speed.score); // 95点
 */
function getResponseSpeedData(candidateId, phase) {
  try {
    // 送信ログを取得
    const sendLogs = getSurveySendLog(candidateId, phase);
    if (sendLogs.length === 0) {
      Logger.log(`⚠️ 送信ログなし: ${candidateId} (${phase})`);
      return null;
    }

    // 最新の送信ログ（成功したもの）
    const successfulLog = sendLogs
      .filter(log => log.send_status === '成功')
      .sort((a, b) => new Date(b.send_time) - new Date(a.send_time))[0];

    if (!successfulLog) {
      Logger.log(`⚠️ 送信成功ログなし: ${candidateId} (${phase})`);
      return null;
    }

    // アンケート回答を取得
    const responses = getSurveyResponse(candidateId, phase);
    if (responses.length === 0) {
      Logger.log(`⚠️ アンケート回答なし: ${candidateId} (${phase})`);
      return null;
    }

    // 最新の回答
    const latestResponse = responses.sort((a, b) =>
      new Date(b.response_date) - new Date(a.response_date)
    )[0];

    // 送信日時と回答日時の差を計算（時間単位）
    const sendTime = new Date(successfulLog.send_time);
    const responseTime = new Date(latestResponse.response_date);
    const diffMs = responseTime - sendTime;
    const diffHours = diffMs / (1000 * 60 * 60);

    // 回答速度スコアを計算（0-100点）
    // 0-2時間: 100点
    // 2-6時間: 100 - 80点（線形減少）
    // 6-24時間: 80 - 50点（線形減少）
    // 24-48時間: 50 - 20点（線形減少）
    // 48時間以上: 20点以下
    let score = 0;
    if (diffHours < 0) {
      // 回答が送信前（データ不整合）
      score = 0;
    } else if (diffHours <= 2) {
      score = 100;
    } else if (diffHours <= 6) {
      score = 100 - ((diffHours - 2) / 4) * 20; // 100 → 80
    } else if (diffHours <= 24) {
      score = 80 - ((diffHours - 6) / 18) * 30; // 80 → 50
    } else if (diffHours <= 48) {
      score = 50 - ((diffHours - 24) / 24) * 30; // 50 → 20
    } else {
      score = Math.max(0, 20 - ((diffHours - 48) / 24) * 5); // 20 → 0
    }

    return {
      send_time: sendTime,
      response_time: responseTime,
      hours: Math.round(diffHours * 10) / 10, // 小数点1桁
      score: Math.round(score)
    };

  } catch (error) {
    logError('getResponseSpeedData', error);
    return null;
  }
}

/**
 * 候補者のメールアドレスを取得
 *
 * @param {string} candidateId - 候補者ID
 * @return {string|null} メールアドレス
 *
 * 使用例:
 * const email = getCandidateEmail('C001');
 * // → 'tanaka@example.com'
 */
function getCandidateEmail(candidateId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);

    if (!master) {
      Logger.log('❌ Candidates_Masterシートが見つかりません');
      return null;
    }

    const data = master.getDataRange().getValues();

    // ヘッダー行をスキップ（i=1から開始）
    for (let i = 1; i < data.length; i++) {
      if (data[i][CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID] === candidateId) {
        const email = data[i][CONFIG.COLUMNS.CANDIDATES_MASTER.EMAIL];

        if (!email || email === '') {
          Logger.log(`⚠️ メールアドレスが空です: ${candidateId}`);
          return null;
        }

        Logger.log(`✅ メールアドレス取得: ${candidateId} → ${email}`);
        return email;
      }
    }

    Logger.log(`❌ 候補者が見つかりません: ${candidateId}`);
    return null;

  } catch (error) {
    Logger.log(`❌ メールアドレス取得エラー: ${error}`);
    return null;
  }
}
