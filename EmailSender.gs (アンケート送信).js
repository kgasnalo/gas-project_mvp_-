/**
 * EmailSender.gs
 * アンケート送信機能（Gmail自動送信）
 *
 * 【主要機能】
 * - Gmail送信制限の事前チェック（90通/日）
 * - リトライ機構（3回まで、2秒間隔）
 * - 送信履歴の自動記録
 * - エラーハンドリング
 *
 * @version 1.0
 * @date 2025-11-06
 */

/**
 * アンケートを送信（安全版）
 *
 * @param {string} candidateId - 候補者ID（例: 'C001'）
 * @param {string} phase - アンケート種別（'初回面談'/'社員面談'/'2次面接'/'内定後'）
 * @return {boolean} 送信成功: true, 失敗: false
 *
 * @example
 * sendSurveyEmailSafe('C001', '初回面談');
 */
function sendSurveyEmailSafe(candidateId, phase) {
  let candidate = null; // catchブロックでも参照できるように外で定義

  try {
    // 【ステップ1】送信前に制限チェック
    const todayCount = getTodaySendCount();

    if (todayCount >= CONFIG.EMAIL.DAILY_LIMIT) {
      throw new Error(
        `本日の送信制限に達しました（${todayCount}/${CONFIG.EMAIL.DAILY_LIMIT}通）\n` +
        '明日以降に再送信してください。'
      );
    }

    // 【ステップ2】候補者情報を取得
    candidate = getCandidateInfo(candidateId);

    if (!candidate.email) {
      throw new Error('メールアドレスが登録されていません（AQ列を確認してください）');
    }

    // 【ステップ3】メール内容を構築
    const emailContent = buildEmailContent(phase, candidate);

    // 【ステップ4】リトライ付きで送信
    sendEmailWithRetry(
      candidate.email,
      emailContent.subject,
      emailContent.body,
      {
        htmlBody: emailContent.htmlBody,
        name: 'PIGNUS 採用担当'
      }
    );

    // 【ステップ5】送信ログを記録
    recordSendLog({
      candidateId: candidateId,
      candidateName: candidate.name,
      email: candidate.email,
      phase: phase,
      sendTime: new Date(),
      sendStatus: '成功',
      errorMessage: ''
    });

    Logger.log(`✅ アンケート送信成功: ${candidate.name} (${phase})`);

    // 成功メッセージを表示（UIが使える場合のみ）
    try {
      SpreadsheetApp.getUi().alert(
        'アンケート送信完了',
        `${candidate.name}様に${phase}アンケートを送信しました。\n\n` +
        `本日の送信数: ${todayCount + 1}/${CONFIG.EMAIL.DAILY_LIMIT}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiError) {
      // UIが使えないコンテキスト（エディタから直接実行など）の場合はログのみ
      Logger.log(`📧 送信完了メッセージ: ${candidate.name}様に${phase}アンケートを送信しました。本日の送信数: ${todayCount + 1}/${CONFIG.EMAIL.DAILY_LIMIT}`);
    }

    return true;

  } catch (error) {
    Logger.log(`❌ アンケート送信エラー: ${error.message}`);

    // エラーログを記録
    recordSendLog({
      candidateId: candidateId,
      candidateName: candidate ? candidate.name : '',
      email: candidate ? candidate.email : '',
      phase: phase,
      sendTime: new Date(),
      sendStatus: '失敗',
      errorMessage: error.toString()
    });

    // エラーメッセージを表示（UIが使える場合のみ）
    try {
      SpreadsheetApp.getUi().alert(
        'アンケート送信エラー',
        `送信に失敗しました。\n\n${error.message}\n\n` +
        'Survey_Send_Logシートでエラー詳細を確認できます。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiError) {
      // UIが使えないコンテキストの場合はログのみ
      Logger.log(`❌ エラーメッセージ: 送信に失敗しました。${error.message}`);
    }

    return false;
  }
}

/**
 * リトライ付きメール送信
 *
 * @param {string} to - 送信先メールアドレス
 * @param {string} subject - 件名
 * @param {string} body - 本文（テキスト）
 * @param {Object} options - オプション（htmlBody, nameなど）
 * @throws {Error} 最大リトライ回数後も失敗した場合
 */
function sendEmailWithRetry(to, subject, body, options) {
  const maxRetries = CONFIG.EMAIL.RETRY_COUNT || 3;
  const retryDelay = CONFIG.EMAIL.RETRY_DELAY || 2000;

  for (let i = 0; i < maxRetries; i++) {
    try {
      GmailApp.sendEmail(to, subject, body, options);
      Logger.log(`✅ メール送信成功（${i + 1}回目）`);
      return;

    } catch (error) {
      Logger.log(`⚠️ メール送信失敗（${i + 1}/${maxRetries}回目）: ${error.message}`);

      if (i === maxRetries - 1) {
        throw new Error(`メール送信に${maxRetries}回失敗しました: ${error.message}`);
      }

      // リトライ前に待機
      Utilities.sleep(retryDelay);
    }
  }
}

/**
 * 今日の送信数を取得
 *
 * @return {number} 本日の送信成功数
 */
function getTodaySendCount() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);

    if (!sheet) {
      Logger.log('⚠️ Survey_Send_Logシートが見つかりません。0を返します。');
      return 0;
    }

    const data = sheet.getDataRange().getValues();

    // 今日の日付（時刻を0:00:00にリセット）
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let count = 0;

    // 2行目から（ヘッダー除く）
    for (let i = 1; i < data.length; i++) {
      const sendTime = new Date(data[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.SEND_TIME]);
      sendTime.setHours(0, 0, 0, 0);

      const status = data[i][CONFIG.COLUMNS.SURVEY_SEND_LOG.STATUS];

      if (sendTime.getTime() === today.getTime() && status === '成功') {
        count++;
      }
    }

    Logger.log(`📊 本日の送信数: ${count}/${CONFIG.EMAIL.DAILY_LIMIT}`);
    return count;

  } catch (error) {
    Logger.log(`❌ 送信数取得エラー: ${error.message}`);
    return 0; // エラー時は0を返す（安全側）
  }
}

/**
 * 候補者情報を取得
 *
 * @param {string} candidateId - 候補者ID
 * @return {Object} { id, name, email }
 * @throws {Error} 候補者が見つからない場合
 */
function getCandidateInfo(candidateId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);
  const data = master.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID] === candidateId) {
      return {
        id: data[i][CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID],
        name: data[i][CONFIG.COLUMNS.CANDIDATES_MASTER.NAME],
        email: data[i][CONFIG.COLUMNS.CANDIDATES_MASTER.EMAIL] // AQ列
      };
    }
  }

  throw new Error(`候補者が見つかりません: ${candidateId}`);
}

/**
 * メール内容を構築
 *
 * @param {string} phase - アンケート種別
 * @param {Object} candidate - 候補者情報 { name, email }
 * @return {Object} { subject, body, htmlBody }
 * @throws {Error} URLまたはテンプレートが見つからない場合
 */
function buildEmailContent(phase, candidate) {
  // URLをCONFIG.gsから取得
  const surveyUrl = CONFIG.SURVEY_URLS[phase];

  if (!surveyUrl) {
    throw new Error(`アンケートURLが設定されていません: ${phase}\nConfig.gsのSURVEY_URLSを確認してください。`);
  }

  const templates = {
    '初回面談': {
      subject: '【PIGNUS】初回面談アンケートのお願い',
      getBody: (name, url) => `${name} 様

お世話になっております。
PIGNUS採用担当です。

先日は初回面談にお時間をいただき、誠にありがとうございました。

つきましては、今後の選考の参考とさせていただきたく、
簡単なアンケートへのご協力をお願いできますでしょうか。

▼アンケートURL
${url}

所要時間は3-5分程度です。
お忙しいところ恐縮ですが、ご協力のほどよろしくお願いいたします。

---
PIGNUS 採用担当`,
      getHtmlBody: (name, url) => `<p>${name} 様</p>
<p>お世話になっております。<br>PIGNUS採用担当です。</p>
<p>先日は初回面談にお時間をいただき、誠にありがとうございました。</p>
<p>つきましては、今後の選考の参考とさせていただきたく、<br>簡単なアンケートへのご協力をお願いできますでしょうか。</p>
<p><strong>▼アンケートURL</strong><br>
<a href="${url}">${url}</a></p>
<p>所要時間は3-5分程度です。<br>お忙しいところ恐縮ですが、ご協力のほどよろしくお願いいたします。</p>
<hr>
<p>PIGNUS 採用担当</p>`
    },

    '社員面談': {
      subject: '【PIGNUS】社員面談アンケートのお願い',
      getBody: (name, url) => `${name} 様

お世話になっております。
PIGNUS採用担当です。

先日は社員面談にご参加いただき、誠にありがとうございました。

つきましては、今後の選考の参考とさせていただきたく、
簡単なアンケートへのご協力をお願いできますでしょうか。

▼アンケートURL
${url}

所要時間は3-5分程度です。
お忙しいところ恐縮ですが、ご協力のほどよろしくお願いいたします。

---
PIGNUS 採用担当`,
      getHtmlBody: (name, url) => `<p>${name} 様</p>
<p>お世話になっております。<br>PIGNUS採用担当です。</p>
<p>先日は社員面談にご参加いただき、誠にありがとうございました。</p>
<p>つきましては、今後の選考の参考とさせていただきたく、<br>簡単なアンケートへのご協力をお願いできますでしょうか。</p>
<p><strong>▼アンケートURL</strong><br>
<a href="${url}">${url}</a></p>
<p>所要時間は3-5分程度です。<br>お忙しいところ恐縮ですが、ご協力のほどよろしくお願いいたします。</p>
<hr>
<p>PIGNUS 採用担当</p>`
    },

    '2次面接': {
      subject: '【PIGNUS】2次面接アンケートのお願い',
      getBody: (name, url) => `${name} 様

お世話になっております。
PIGNUS採用担当です。

先日は2次面接にお時間をいただき、誠にありがとうございました。

つきましては、今後の選考の参考とさせていただきたく、
簡単なアンケートへのご協力をお願いできますでしょうか。

▼アンケートURL
${url}

所要時間は5分程度です。
お忙しいところ恐縮ですが、ご協力のほどよろしくお願いいたします。

---
PIGNUS 採用担当`,
      getHtmlBody: (name, url) => `<p>${name} 様</p>
<p>お世話になっております。<br>PIGNUS採用担当です。</p>
<p>先日は2次面接にお時間をいただき、誠にありがとうございました。</p>
<p>つきましては、今後の選考の参考とさせていただきたく、<br>簡単なアンケートへのご協力をお願いできますでしょうか。</p>
<p><strong>▼アンケートURL</strong><br>
<a href="${url}">${url}</a></p>
<p>所要時間は5分程度です。<br>お忙しいところ恐縮ですが、ご協力のほどよろしくお願いいたします。</p>
<hr>
<p>PIGNUS 採用担当</p>`
    },

    '内定後': {
      subject: '【PIGNUS】内定通知後アンケートのお願い',
      getBody: (name, url) => `${name} 様

お世話になっております。
PIGNUS採用担当です。

この度は内定のご連絡をさせていただきました。

つきましては、今後のフォローアップの参考とさせていただきたく、
簡単なアンケートへのご協力をお願いできますでしょうか。

▼アンケートURL
${url}

所要時間は5分程度です。
お忙しいところ恐縮ですが、ご協力のほどよろしくお願いいたします。

---
PIGNUS 採用担当`,
      getHtmlBody: (name, url) => `<p>${name} 様</p>
<p>お世話になっております。<br>PIGNUS採用担当です。</p>
<p>この度は内定のご連絡をさせていただきました。</p>
<p>つきましては、今後のフォローアップの参考とさせていただきたく、<br>簡単なアンケートへのご協力をお願いできますでしょうか。</p>
<p><strong>▼アンケートURL</strong><br>
<a href="${url}">${url}</a></p>
<p>所要時間は5分程度です。<br>お忙しいところ恐縮ですが、ご協力のほどよろしくお願いいたします。</p>
<hr>
<p>PIGNUS 採用担当</p>`
    }
  };

  const template = templates[phase];

  if (!template) {
    throw new Error(`テンプレートが見つかりません: ${phase}`);
  }

  return {
    subject: template.subject,
    body: template.getBody(candidate.name, surveyUrl),
    htmlBody: template.getHtmlBody(candidate.name, surveyUrl)
  };
}

/**
 * 送信ログを記録
 *
 * @param {Object} data - 送信情報
 * @param {string} data.candidateId - 候補者ID
 * @param {string} data.candidateName - 候補者名
 * @param {string} data.email - メールアドレス
 * @param {string} data.phase - アンケート種別
 * @param {Date} data.sendTime - 送信日時
 * @param {string} data.sendStatus - 送信ステータス（'成功'/'失敗'）
 * @param {string} data.errorMessage - エラーメッセージ
 */
function recordSendLog(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);

  if (!sheet) {
    setupSurveySendLog(); // シートが存在しない場合は作成
    sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);
  }

  sheet.appendRow([
    'SEND-' + new Date().getTime(), // send_id
    data.candidateId,
    data.candidateName,
    data.email,
    data.phase,
    data.sendTime,
    data.sendStatus,
    data.errorMessage
  ]);

  Logger.log(`📝 送信ログを記録しました: ${data.candidateName} (${data.sendStatus})`);
}

// ========================================
// テスト関数
// ========================================

/**
 * テスト1: 基本的な送信テスト
 *
 * 実行方法:
 * 1. Candidates_MasterのAQ列にテスト用メールアドレスを入力
 * 2. GASエディターでこの関数を選択
 * 3. 実行ボタン（▶️）をクリック
 */
function testEmailSend() {
  // Candidates_Masterに存在する候補者IDを指定
  // AQ列にメールアドレスが入力されている必要があります
  sendSurveyEmailSafe('C001', '初回面談');
}

/**
 * テスト2: 送信数カウントのテスト
 *
 * 実行方法:
 * 1. GASエディターでこの関数を選択
 * 2. 実行ボタン（▶️）をクリック
 * 3. 「表示」→「ログ」で送信数を確認
 */
function testSendCount() {
  const count = getTodaySendCount();
  Logger.log(`本日の送信数: ${count}/${CONFIG.EMAIL.DAILY_LIMIT}`);
}

/**
 * テスト3: エラーケースのテスト
 *
 * 実行方法:
 * 1. GASエディターでこの関数を選択
 * 2. 実行ボタン（▶️）をクリック
 * 3. エラーダイアログが表示されることを確認
 */
function testEmailSendError() {
  // 存在しない候補者IDでテスト
  sendSurveyEmailSafe('C999', '初回面談');
}
