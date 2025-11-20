/**
 * ========================================
 * Phase 3.5: トリガー実装
 * ========================================
 *
 * 目的: アンケート送信時にEngagement_Logへ自動書き込み
 *
 * トリガーの種類:
 * - フォーム送信時トリガー（onFormSubmit）
 * - 各アンケートフォームに個別に設定
 */

/**
 * 初回面談アンケート送信時のトリガー
 *
 * @param {Object} e - イベントオブジェクト
 */
function onFormSubmit初回面談(e) {
  try {
    Logger.log('\n========================================');
    Logger.log('初回面談アンケート送信検知');
    Logger.log('========================================\n');

    calculateAndLogAcceptance(e, '初回面談');

  } catch (error) {
    logError('onFormSubmit初回面談', error);
  }
}

/**
 * 社員面談アンケート送信時のトリガー
 *
 * @param {Object} e - イベントオブジェクト
 */
function onFormSubmit社員面談(e) {
  try {
    Logger.log('\n========================================');
    Logger.log('社員面談アンケート送信検知');
    Logger.log('========================================\n');

    calculateAndLogAcceptance(e, '社員面談');

  } catch (error) {
    logError('onFormSubmit社員面談', error);
  }
}

/**
 * 2次面接アンケート送信時のトリガー
 *
 * @param {Object} e - イベントオブジェクト
 */
function onFormSubmit2次面接(e) {
  try {
    Logger.log('\n========================================');
    Logger.log('2次面接アンケート送信検知');
    Logger.log('========================================\n');

    calculateAndLogAcceptance(e, '2次面接');

  } catch (error) {
    logError('onFormSubmit2次面接', error);
  }
}

/**
 * 内定後アンケート送信時のトリガー
 *
 * @param {Object} e - イベントオブジェクト
 */
function onFormSubmit内定後(e) {
  try {
    Logger.log('\n========================================');
    Logger.log('内定後アンケート送信検知');
    Logger.log('========================================\n');

    calculateAndLogAcceptance(e, '内定後');

  } catch (error) {
    logError('onFormSubmit内定後', error);
  }
}

/**
 * 承諾可能性を計算してEngagement_Logに記録（共通処理）
 *
 * @param {Object} e - イベントオブジェクト
 * @param {string} phase - アンケート種別
 */
function calculateAndLogAcceptance(e, phase) {
  try {
    // イベントオブジェクトから回答値を取得
    const values = e.values;

    // メールアドレスを取得（C列: インデックス2）
    const email = values[2];

    if (!email) {
      Logger.log('❌ メールアドレスが取得できません');
      return;
    }

    Logger.log(`✅ メールアドレス取得: ${email}`);

    // メールアドレスから候補者IDを取得
    const candidateId = getCandidateIdByEmail(email);

    if (!candidateId) {
      Logger.log(`❌ 候補者IDが見つかりません: ${email}`);
      return;
    }

    Logger.log(`✅ 候補者ID取得: ${candidateId}`);

    // 承諾可能性を計算してEngagement_Logに書き込み
    const success = writeToEngagementLog(candidateId, phase);

    if (success) {
      Logger.log(`✅ Engagement_Logへの書き込み成功: ${candidateId}, ${phase}`);
    } else {
      Logger.log(`❌ Engagement_Logへの書き込み失敗: ${candidateId}, ${phase}`);
      sendErrorNotification(`Engagement_Log書き込み失敗: ${candidateId}, ${phase}`);
    }

  } catch (error) {
    Logger.log(`❌ 承諾可能性計算エラー: ${error}`);
    sendErrorNotification(`承諾可能性計算エラー: ${error}`);
  }
}

/**
 * メールアドレスから候補者IDを取得
 *
 * @param {string} email - メールアドレス
 * @return {string|null} 候補者ID
 */
function getCandidateIdByEmail(email) {
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
      const candidateEmail = data[i][CONFIG.COLUMNS.CANDIDATES_MASTER.EMAIL];

      if (candidateEmail === email) {
        const candidateId = data[i][0]; // A列: candidate_id
        Logger.log(`✅ 候補者ID発見: ${candidateId}`);
        return candidateId;
      }
    }

    Logger.log(`❌ 候補者が見つかりません: ${email}`);
    return null;

  } catch (error) {
    Logger.log(`❌ 候補者ID取得エラー: ${error}`);
    return null;
  }
}

/**
 * トリガーのセットアップ
 *
 * 注意: この関数は手動で1回だけ実行してください
 * 各Google Formに対してトリガーを設定します
 */
function setupTriggers() {
  Logger.log('\n========================================');
  Logger.log('トリガーセットアップ開始');
  Logger.log('========================================\n');

  Logger.log('⚠️ 重要: この関数は自動実行できません');
  Logger.log('⚠️ Google Apps Scriptのトリガー設定画面から、以下を手動で設定してください:\n');

  Logger.log('1. 初回面談アンケートフォーム:');
  Logger.log('   - イベントソース: フォームから');
  Logger.log('   - イベントタイプ: フォーム送信時');
  Logger.log('   - 実行する関数: onFormSubmit初回面談\n');

  Logger.log('2. 社員面談アンケートフォーム:');
  Logger.log('   - イベントソース: フォームから');
  Logger.log('   - イベントタイプ: フォーム送信時');
  Logger.log('   - 実行する関数: onFormSubmit社員面談\n');

  Logger.log('3. 2次面接アンケートフォーム:');
  Logger.log('   - イベントソース: フォームから');
  Logger.log('   - イベントタイプ: フォーム送信時');
  Logger.log('   - 実行する関数: onFormSubmit2次面接\n');

  Logger.log('4. 内定後アンケートフォーム:');
  Logger.log('   - イベントソース: フォームから');
  Logger.log('   - イベントタイプ: フォーム送信時');
  Logger.log('   - 実行する関数: onFormSubmit内定後\n');

  Logger.log('========================================');
  Logger.log('トリガーセットアップ手順完了');
  Logger.log('========================================\n');
}
