/**
 * ========================================
 * Phase 3.5: 時間ベース自動チェック
 * ========================================
 *
 * 目的: IMPORTRANGE経由で取得したアンケートデータを定期的にチェックし、
 *       新規回答があればEngagement_Logに自動記録
 *
 * トリガー設定:
 * - 時間主導型トリガー
 * - 推奨: 1時間ごとに実行
 * - 実行する関数: checkForNewSurveyResponses
 */

/**
 * 新規アンケート回答をチェックして処理（メイン関数）
 *
 * この関数を時間ベーストリガーで定期実行してください
 */
function checkForNewSurveyResponses() {
  try {
    Logger.log('\n========================================');
    Logger.log('自動チェック開始: ' + new Date());
    Logger.log('========================================\n');

    // 処理結果を記録
    const results = {
      processed: 0,
      errors: 0,
      skipped: 0
    };

    // 4つのアンケート種別を順番にチェック
    const phases = ['初回面談', '社員面談', '2次面接', '内定後'];

    for (let phase of phases) {
      Logger.log(`\n--- ${phase}アンケートをチェック中 ---`);
      const phaseResult = processSurveyPhase(phase);

      results.processed += phaseResult.processed;
      results.errors += phaseResult.errors;
      results.skipped += phaseResult.skipped;
    }

    // 結果サマリーをログ出力
    Logger.log('\n========================================');
    Logger.log('自動チェック完了');
    Logger.log(`処理済み: ${results.processed}件`);
    Logger.log(`スキップ: ${results.skipped}件`);
    Logger.log(`エラー: ${results.errors}件`);
    Logger.log('========================================\n');

    // エラーがあれば通知
    if (results.errors > 0) {
      sendErrorNotification(`自動チェックでエラー発生: ${results.errors}件`);
    }

    return results;

  } catch (error) {
    Logger.log(`❌ 自動チェックエラー: ${error}`);
    logError('checkForNewSurveyResponses', error);
    sendErrorNotification(`自動チェック失敗: ${error.message}`);
    throw error;
  }
}

/**
 * 特定フェーズのアンケートを処理
 *
 * @param {string} phase - アンケート種別
 * @return {Object} 処理結果 {processed, errors, skipped}
 */
function processSurveyPhase(phase) {
  const result = {
    processed: 0,
    errors: 0,
    skipped: 0
  };

  try {
    // アンケートシート名を取得
    const sheetName = getSheetNameByPhase(phase);

    if (!sheetName) {
      Logger.log(`❌ シート名が取得できません: ${phase}`);
      result.errors++;
      return result;
    }

    // スプレッドシートとシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`❌ シートが見つかりません: ${sheetName}`);
      result.errors++;
      return result;
    }

    // 最後に処理した行番号を取得
    const lastProcessedRow = getLastProcessedRow(phase);
    Logger.log(`最後に処理した行: ${lastProcessedRow}`);

    // シートのデータを取得
    const data = sheet.getDataRange().getValues();

    // ヘッダー行を除く
    const totalRows = data.length;

    if (totalRows <= 1) {
      Logger.log(`データなし（ヘッダーのみ）`);
      return result;
    }

    // 新規行のみ処理（lastProcessedRow + 1 から開始）
    let newRowsFound = false;

    for (let i = lastProcessedRow + 1; i < totalRows; i++) {
      const row = data[i];

      // メールアドレスを取得（C列: インデックス2）
      const email = row[2];

      if (!email || email === '') {
        Logger.log(`行${i + 1}: メールアドレスなし - スキップ`);
        result.skipped++;
        continue;
      }

      Logger.log(`\n行${i + 1}を処理中: ${email}`);

      // メールアドレスから候補者IDを取得
      const candidateId = getCandidateIdByEmail(email);

      if (!candidateId) {
        Logger.log(`❌ 候補者IDが見つかりません: ${email}`);
        result.errors++;
        continue;
      }

      // データ検証
      if (!validateData(candidateId, phase)) {
        Logger.log(`❌ データ検証失敗: ${candidateId}, ${phase}`);
        result.errors++;
        continue;
      }

      // Engagement_Logに書き込み
      try {
        const success = writeToEngagementLog(candidateId, phase);

        if (success) {
          Logger.log(`✅ 処理成功: ${candidateId}, ${phase}`);
          result.processed++;
          newRowsFound = true;

          // 処理した行番号を更新
          updateLastProcessedRow(phase, i);

        } else {
          Logger.log(`❌ 書き込み失敗: ${candidateId}, ${phase}`);
          result.errors++;
        }

      } catch (error) {
        Logger.log(`❌ 書き込みエラー: ${error}`);
        logError('processSurveyPhase', error);
        result.errors++;
      }
    }

    if (!newRowsFound && totalRows > lastProcessedRow + 1) {
      Logger.log(`新規データなし（最終行: ${totalRows}）`);
    }

  } catch (error) {
    Logger.log(`❌ フェーズ処理エラー: ${error}`);
    logError('processSurveyPhase', error);
    result.errors++;
  }

  return result;
}

/**
 * フェーズ名からシート名を取得
 *
 * @param {string} phase - アンケート種別
 * @return {string|null} シート名
 */
function getSheetNameByPhase(phase) {
  const mapping = {
    '初回面談': 'アンケート_初回面談',
    '社員面談': 'アンケート_社員面談',
    '2次面接': 'アンケート_2次面接',
    '内定後': 'アンケート_内定後'
  };

  return mapping[phase] || null;
}

/**
 * 最後に処理した行番号を取得
 *
 * @param {string} phase - アンケート種別
 * @return {number} 最後に処理した行番号（0-indexed、ヘッダー行は0）
 */
function getLastProcessedRow(phase) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let trackingSheet = ss.getSheetByName('Processing_Log');

    // Processing_Logシートが存在しない場合は作成
    if (!trackingSheet) {
      trackingSheet = createProcessingLogSheet();
    }

    // データを取得
    const data = trackingSheet.getDataRange().getValues();

    // ヘッダー行をスキップして検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === phase) {
        const lastRow = data[i][1];
        return lastRow || 0; // 数値に変換、空なら0
      }
    }

    // 該当フェーズが見つからない場合は新規行を追加
    trackingSheet.appendRow([phase, 0, new Date()]);
    return 0;

  } catch (error) {
    Logger.log(`❌ 最終処理行取得エラー: ${error}`);
    return 0; // エラー時は0から開始
  }
}

/**
 * 最後に処理した行番号を更新
 *
 * @param {string} phase - アンケート種別
 * @param {number} rowIndex - 処理した行番号（0-indexed）
 */
function updateLastProcessedRow(phase, rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let trackingSheet = ss.getSheetByName('Processing_Log');

    if (!trackingSheet) {
      trackingSheet = createProcessingLogSheet();
    }

    const data = trackingSheet.getDataRange().getValues();

    // 該当フェーズの行を探して更新
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === phase) {
        trackingSheet.getRange(i + 1, 2).setValue(rowIndex);
        trackingSheet.getRange(i + 1, 3).setValue(new Date());
        Logger.log(`📝 Processing_Log更新: ${phase}, 行${rowIndex}`);
        return;
      }
    }

    // 見つからない場合は新規追加
    trackingSheet.appendRow([phase, rowIndex, new Date()]);
    Logger.log(`📝 Processing_Log新規追加: ${phase}, 行${rowIndex}`);

  } catch (error) {
    Logger.log(`❌ 最終処理行更新エラー: ${error}`);
    logError('updateLastProcessedRow', error);
  }
}

/**
 * Processing_Logシートを作成
 *
 * @return {Sheet} 作成したシート
 */
function createProcessingLogSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet('Processing_Log');

    // ヘッダー行を追加
    sheet.getRange(1, 1, 1, 3).setValues([[
      'Survey Phase',
      'Last Processed Row',
      'Last Updated'
    ]]);

    // ヘッダー行の書式設定
    const headerRange = sheet.getRange(1, 1, 1, 3);
    headerRange.setBackground(CONFIG.COLORS.HEADER_BG);
    headerRange.setFontColor(CONFIG.COLORS.HEADER_TEXT);
    headerRange.setFontWeight('bold');

    // 4つのフェーズの初期データを挿入
    const phases = ['初回面談', '社員面談', '2次面接', '内定後'];
    for (let phase of phases) {
      sheet.appendRow([phase, 0, new Date()]);
    }

    Logger.log('✅ Processing_Logシート作成完了');

    return sheet;

  } catch (error) {
    Logger.log(`❌ Processing_Logシート作成エラー: ${error}`);
    logError('createProcessingLogSheet', error);
    throw error;
  }
}

/**
 * 処理状況をリセット（テスト用）
 *
 * 注意: この関数を実行すると、全てのアンケートを再処理します
 */
function resetProcessingLog() {
  try {
    Logger.log('\n========================================');
    Logger.log('Processing_Logをリセット中...');
    Logger.log('========================================\n');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const trackingSheet = ss.getSheetByName('Processing_Log');

    if (!trackingSheet) {
      Logger.log('Processing_Logシートが存在しません');
      return;
    }

    // 全ての行番号を0にリセット
    const data = trackingSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      trackingSheet.getRange(i + 1, 2).setValue(0);
      trackingSheet.getRange(i + 1, 3).setValue(new Date());
    }

    Logger.log('✅ Processing_Logリセット完了');
    Logger.log('⚠️ 次回の自動チェックで全データが再処理されます');

  } catch (error) {
    Logger.log(`❌ リセットエラー: ${error}`);
    logError('resetProcessingLog', error);
  }
}

/**
 * トリガーのセットアップ手順を表示
 */
function showTriggerSetupInstructions() {
  Logger.log('\n========================================');
  Logger.log('時間ベーストリガーのセットアップ手順');
  Logger.log('========================================\n');

  Logger.log('1. Google Apps Scriptエディタを開く');
  Logger.log('2. 左メニューから「トリガー」（時計アイコン）をクリック');
  Logger.log('3. 右下の「トリガーを追加」をクリック\n');

  Logger.log('4. 以下の設定を行う:');
  Logger.log('   - 実行する関数を選択: checkForNewSurveyResponses');
  Logger.log('   - 実行するデプロイを選択: Head');
  Logger.log('   - イベントのソースを選択: 時間主導型');
  Logger.log('   - 時間ベースのトリガーのタイプを選択: 時間ベースのタイマー');
  Logger.log('   - 時間の間隔を選択: 1時間おき（推奨）\n');

  Logger.log('5. 「保存」をクリック\n');

  Logger.log('========================================');
  Logger.log('セットアップ完了後、1時間ごとに自動実行されます');
  Logger.log('========================================\n');

  Logger.log('💡 Tips:');
  Logger.log('- 初回は手動で checkForNewSurveyResponses() を実行してテストしてください');
  Logger.log('- Processing_Logシートで処理状況を確認できます');
  Logger.log('- Error_Logシートでエラーを監視してください');
}
