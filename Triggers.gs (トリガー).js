/**
 * スプレッドシート編集時のトリガー
 * Candidates_MasterのAS, AZ列が「実施済」、BB, BD列が「合格」または「不合格」に変更されたときにアンケートを自動送信
 * Survey_ResponseのD列（回答日時）またはI列（アンケート種別）が編集されたときに回答速度を計算
 */
function onEdit(e) {
  try {
    // イベントオブジェクトがない場合は終了
    if (!e) return;

    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const col = range.getColumn();
    const row = range.getRow();

    // ヘッダー行は無視
    if (row === 1) return;

    // ========== Candidates_Masterの処理 ==========
    if (sheet.getName() === CONFIG.SHEET_NAMES.CANDIDATES_MASTER) {
      // 新しい値を取得
      const newValue = e.value;
      if (!newValue) return;

      // どの列が編集されたかによって処理を分岐
      switch (col - 1) { // 0-indexed
        case CONFIG.COLUMNS.CANDIDATES_MASTER.FIRST_INTERVIEW_STATUS: // AS列
          if (newValue === '実施済') {
            handleFirstInterviewSurvey(sheet, row);
          }
          break;
        case CONFIG.COLUMNS.CANDIDATES_MASTER.EMPLOYEE_INTERVIEW_STATUS: // AZ列
          if (newValue === '実施済') {
            handleEmployeeInterviewSurvey(sheet, row);
          }
          break;
        case CONFIG.COLUMNS.CANDIDATES_MASTER.SECOND_INTERVIEW_STATUS: // BB列
          if (newValue === '合格' || newValue === '不合格') {
            handleSecondInterviewSurvey(sheet, row);
          }
          break;
        case CONFIG.COLUMNS.CANDIDATES_MASTER.FINAL_INTERVIEW_STATUS: // BD列
          if (newValue === '合格' || newValue === '不合格') {
            handleFinalInterviewSurvey(sheet, row);
          }
          break;
      }
    }

    // ========== 【Phase 2 Step 3追加】Survey_Responseの処理 ==========
    else if (sheet.getName() === CONFIG.SHEET_NAMES.SURVEY_RESPONSE) {
      // D列（回答日時）またはI列（アンケート種別）が編集された場合
      if (col === CONFIG.COLUMNS.SURVEY_RESPONSE.RESPONSE_DATE + 1 ||
          col === CONFIG.COLUMNS.SURVEY_RESPONSE.PHASE + 1) {

        const candidateId = sheet.getRange(row, CONFIG.COLUMNS.SURVEY_RESPONSE.CANDIDATE_ID + 1).getValue();
        const phase = sheet.getRange(row, CONFIG.COLUMNS.SURVEY_RESPONSE.PHASE + 1).getValue();

        if (candidateId && phase) {
          Logger.log(`📊 Survey_Response更新検知: ${candidateId} (${phase})`);

          // 回答速度を計算・更新（少し遅延させる）
          Utilities.sleep(1000); // 1秒待機（データ確定を待つ）
          calculateAndUpdateResponseSpeed(candidateId, phase);
        }
      }
    }

  } catch (error) {
    logError('onEdit', error);
  }
}

/**
 * 初回面談後のアンケート送信
 */
function handleFirstInterviewSurvey(sheet, row) {
  try {
    const candidateId = sheet.getRange(row, CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID + 1).getValue();
    const candidateName = sheet.getRange(row, CONFIG.COLUMNS.CANDIDATES_MASTER.NAME + 1).getValue();

    // 重複チェック
    if (isAlreadySent(candidateId, '初回面談')) {
      Logger.log(`⚠️ ${candidateName}（${candidateId}）の初回面談アンケートは既に送信済みです`);
      return;
    }

    // 確認ダイアログ
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '初回面談アンケート送信確認',
      `${candidateName}さんに初回面談後アンケートを送信しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      sendSurveyEmailSafe(candidateId, '初回面談');
    }

  } catch (error) {
    logError('handleFirstInterviewSurvey', error);
  }
}

/**
 * 社員面談後のアンケート送信
 * ※社員面談は複数回実施される可能性があるため、最終回のみ送信
 */
function handleEmployeeInterviewSurvey(sheet, row) {
  try {
    const candidateId = sheet.getRange(row, CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID + 1).getValue();
    const candidateName = sheet.getRange(row, CONFIG.COLUMNS.CANDIDATES_MASTER.NAME + 1).getValue();
    const employeeInterviewCount = sheet.getRange(row, CONFIG.COLUMNS.CANDIDATES_MASTER.EMPLOYEE_INTERVIEW_COUNT + 1).getValue();

    // 社員面談実施回数が2回以上の場合のみ送信（最終回と判断）
    if (employeeInterviewCount < 2) {
      Logger.log(`ℹ️ ${candidateName}（${candidateId}）の社員面談は${employeeInterviewCount}回目のため、アンケートは送信しません`);
      return;
    }

    // 重複チェック
    if (isAlreadySent(candidateId, '社員面談')) {
      Logger.log(`⚠️ ${candidateName}（${candidateId}）の社員面談アンケートは既に送信済みです`);
      return;
    }

    // 確認ダイアログ
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '社員面談アンケート送信確認',
      `${candidateName}さんに社員面談後アンケート（${employeeInterviewCount}回目）を送信しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      sendSurveyEmailSafe(candidateId, '社員面談');
    }

  } catch (error) {
    logError('handleEmployeeInterviewSurvey', error);
  }
}

/**
 * 2次面接後のアンケート送信
 */
function handleSecondInterviewSurvey(sheet, row) {
  try {
    const candidateId = sheet.getRange(row, CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID + 1).getValue();
    const candidateName = sheet.getRange(row, CONFIG.COLUMNS.CANDIDATES_MASTER.NAME + 1).getValue();

    // 重複チェック
    if (isAlreadySent(candidateId, '2次面接')) {
      Logger.log(`⚠️ ${candidateName}（${candidateId}）の2次面接アンケートは既に送信済みです`);
      return;
    }

    // 確認ダイアログ
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '2次面接アンケート送信確認',
      `${candidateName}さんに2次面接後アンケートを送信しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      sendSurveyEmailSafe(candidateId, '2次面接');
    }

  } catch (error) {
    logError('handleSecondInterviewSurvey', error);
  }
}

/**
 * 最終面接（内定後）アンケート送信
 */
function handleFinalInterviewSurvey(sheet, row) {
  try {
    const candidateId = sheet.getRange(row, CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID + 1).getValue();
    const candidateName = sheet.getRange(row, CONFIG.COLUMNS.CANDIDATES_MASTER.NAME + 1).getValue();

    // 重複チェック
    if (isAlreadySent(candidateId, '内定後')) {
      Logger.log(`⚠️ ${candidateName}（${candidateId}）の内定後アンケートは既に送信済みです`);
      return;
    }

    // 確認ダイアログ
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '内定後アンケート送信確認',
      `${candidateName}さんに内定後アンケートを送信しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      sendSurveyEmailSafe(candidateId, '内定後');
    }

  } catch (error) {
    logError('handleFinalInterviewSurvey', error);
  }
}

/**
 * 重複送信チェック
 * @param {string} candidateId - 候補者ID
 * @param {string} phase - アンケート種別
 * @return {boolean} 既に送信済みの場合はtrue
 */
function isAlreadySent(candidateId, phase) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);

    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();

    // ヘッダー行をスキップして検索
    for (let i = 1; i < data.length; i++) {
      const logCandidateId = data[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.CANDIDATE_ID];
      const logPhase = data[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.PHASE];
      const logStatus = data[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.STATUS];

      // 同じ候補者・同じフェーズで送信成功の記録があるかチェック
      if (logCandidateId === candidateId && logPhase === phase && logStatus === '成功') {
        return true;
      }
    }

    return false;

  } catch (error) {
    logError('isAlreadySent', error);
    return false;
  }
}

/**
 * 今日の送信状況を表示
 */
function showTodaySendCount() {
  const todayCount = getTodaySendCount(); // EmailSender.gsの関数を利用
  const limit = CONFIG.EMAIL.DAILY_LIMIT;
  const remaining = limit - todayCount;

  // リセット時刻（翌日0:00）
  const now = new Date();
  const tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  const hoursUntilReset = Math.floor((tomorrow - now) / (1000 * 60 * 60));
  const minutesUntilReset = Math.floor(((tomorrow - now) % (1000 * 60 * 60)) / (1000 * 60));

  // メッセージ作成
  let statusIcon = '';
  let statusMessage = '';

  if (remaining > 50) {
    statusIcon = '✅';
    statusMessage = '十分な送信枠があります';
  } else if (remaining > 20) {
    statusIcon = '⚠️';
    statusMessage = '送信枠が少なくなっています';
  } else if (remaining > 0) {
    statusIcon = '🚨';
    statusMessage = '送信枠がほぼ上限です';
  } else {
    statusIcon = '❌';
    statusMessage = '本日の送信制限に達しました';
  }

  const message =
    '【📧 今日のアンケート送信状況】\n\n' +
    `${statusIcon} ${statusMessage}\n\n` +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    `📤 本日の送信数: ${todayCount} / ${limit}通\n` +
    `📥 残り送信可能: ${remaining}通\n\n` +
    `🕐 制限リセット: 約${hoursUntilReset}時間${minutesUntilReset}分後\n` +
    '　 （翌日0:00にリセットされます）\n\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '💡 ヒント\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '・送信制限はGmailの仕様です\n' +
    '・安全のため90通/日に設定\n' +
    '・重要な送信を優先してください';

  SpreadsheetApp.getUi().alert(
    '送信状況',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 初回面談アンケート送信（手動）
 */
function showSendFirstInterviewSurvey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '初回面談アンケート送信',
    '候補者ID（例: C001）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const candidateId = response.getResponseText().trim();
    if (candidateId) {
      sendSurveyEmailSafe(candidateId, '初回面談');
    }
  }
}

/**
 * 社員面談アンケート送信（手動）
 */
function showSendEmployeeInterviewSurvey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '社員面談アンケート送信',
    '候補者ID（例: C001）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const candidateId = response.getResponseText().trim();
    if (candidateId) {
      sendSurveyEmailSafe(candidateId, '社員面談');
    }
  }
}

/**
 * 2次面接アンケート送信（手動）
 */
function showSendSecondInterviewSurvey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '2次面接アンケート送信',
    '候補者ID（例: C001）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const candidateId = response.getResponseText().trim();
    if (candidateId) {
      sendSurveyEmailSafe(candidateId, '2次面接');
    }
  }
}

/**
 * 内定後アンケート送信（手動）
 */
function showSendFinalInterviewSurvey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '内定後アンケート送信',
    '候補者ID（例: C001）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const candidateId = response.getResponseText().trim();
    if (candidateId) {
      sendSurveyEmailSafe(candidateId, '内定後');
    }
  }
}

/**
 * 送信履歴を表示
 */
function showSendHistory() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);

    if (!sheet) {
      SpreadsheetApp.getUi().alert('Survey_Send_Logシートが見つかりません');
      return;
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert('送信履歴', '送信履歴がありません', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // 最新10件を取得
    const recentLogs = [];
    for (let i = Math.max(1, data.length - 10); i < data.length; i++) {
      const log = data[i];
      const sendTime = Utilities.formatDate(
        new Date(log[CONFIG.COLUMNS.SURVEY_SEND_LOG.SEND_TIME]),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd HH:mm:ss'
      );

      recentLogs.push(
        `${log[CONFIG.COLUMNS.SURVEY_SEND_LOG.CANDIDATE_ID]} / ${log[CONFIG.COLUMNS.SURVEY_SEND_LOG.NAME]} / ${log[CONFIG.COLUMNS.SURVEY_SEND_LOG.PHASE]} / ${log[CONFIG.COLUMNS.SURVEY_SEND_LOG.STATUS]} / ${sendTime}`
      );
    }

    const message = '最新10件の送信履歴:\n\n' + recentLogs.reverse().join('\n');
    SpreadsheetApp.getUi().alert('送信履歴', message, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    logError('showSendHistory', error);
  }
}
