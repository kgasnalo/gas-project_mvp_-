/**
 * InitialSetup.gs
 * Phase A 初期設定用ユーティリティ関数
 *
 * このファイルには、Phase A完了後の初期設定に必要な関数が含まれています。
 */

/**
 * 企業名をスクリプトプロパティに設定
 *
 * 実行手順:
 * 1. Apps Scriptエディタを開く
 * 2. 関数選択で「setCompanyName」を選択
 * 3. 実行ボタンをクリック
 *
 * @param {string} companyName - 企業名（デフォルト: "アマネク"）
 */
function setCompanyName(companyName = "アマネク") {
  Logger.log('=== 企業名設定開始 ===');

  try {
    // スクリプトプロパティに保存
    PropertiesService.getScriptProperties()
      .setProperty('COMPANY_NAME', companyName);

    Logger.log('✅ 企業名設定完了: ' + companyName);

    // 確認のため読み取り
    const savedName = PropertiesService.getScriptProperties()
      .getProperty('COMPANY_NAME');
    Logger.log('✅ 保存確認: ' + savedName);

    return {
      success: true,
      companyName: savedName
    };

  } catch (error) {
    Logger.log('❌ 企業名設定エラー: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Phase A 統合テスト
 *
 * テスト項目:
 * 1. スプレッドシート接続確認
 * 2. 必要なシート存在確認
 * 3. ドライブフォルダ確認
 * 4. 企業名設定確認
 * 5. 各種関数の動作確認
 *
 * 実行手順:
 * 1. Apps Scriptエディタを開く
 * 2. 関数選択で「testPhaseAComplete」を選択
 * 3. 実行ボタンをクリック
 * 4. ログを確認（表示 → ログ）
 */
function testPhaseAComplete() {
  Logger.log('');
  Logger.log('='.repeat(60));
  Logger.log('=== Phase A 統合テスト開始 ===');
  Logger.log('='.repeat(60));
  Logger.log('');

  const results = {
    spreadsheet: false,
    sheets: {},
    driveFolder: false,
    companyName: false,
    functions: {},
    errors: []
  };

  try {
    // ========================================
    // 1. スプレッドシート確認
    // ========================================
    Logger.log('--- 1. スプレッドシート確認 ---');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('✅ スプレッドシート名: ' + ss.getName());
    Logger.log('✅ スプレッドシートID: ' + ss.getId());
    results.spreadsheet = true;
    Logger.log('');

    // ========================================
    // 2. 必要なシート確認
    // ========================================
    Logger.log('--- 2. 必要なシート確認 ---');
    const requiredSheets = [
      'Candidates_Master',
      'Evaluation_Master',
      'Engagement_Log'
    ];

    requiredSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        Logger.log(`✅ ${sheetName}: 存在（${lastRow}行 × ${lastCol}列）`);
        results.sheets[sheetName] = true;
      } else {
        Logger.log(`❌ ${sheetName}: 存在しない`);
        results.sheets[sheetName] = false;
        results.errors.push(`シート「${sheetName}」が見つかりません`);
      }
    });
    Logger.log('');

    // ========================================
    // 3. ドライブフォルダ確認
    // ========================================
    Logger.log('--- 3. ドライブフォルダ確認 ---');
    try {
      const rootFolder = getOrCreateRootFolder();
      Logger.log('✅ ルートフォルダ名: ' + rootFolder.getName());
      Logger.log('✅ ルートフォルダURL: ' + rootFolder.getUrl());
      results.driveFolder = true;
    } catch (error) {
      Logger.log('❌ ドライブフォルダ確認エラー: ' + error.message);
      results.errors.push('ドライブフォルダの確認に失敗: ' + error.message);
    }
    Logger.log('');

    // ========================================
    // 4. 企業名確認
    // ========================================
    Logger.log('--- 4. 企業名確認 ---');
    const companyName = PropertiesService.getScriptProperties()
      .getProperty('COMPANY_NAME');

    if (companyName) {
      Logger.log('✅ 企業名: ' + companyName);
      results.companyName = true;
    } else {
      Logger.log('⚠️ 企業名が未設定です');
      Logger.log('   setCompanyName() を実行してください');
      results.errors.push('企業名が未設定（setCompanyName()を実行してください）');
    }
    Logger.log('');

    // ========================================
    // 5. 重要関数の動作確認
    // ========================================
    Logger.log('--- 5. 重要関数の動作確認 ---');

    // 5-1. Evaluation_Masterの列構造確認
    Logger.log('5-1. Evaluation_Master 列構造確認');
    try {
      const evalSheet = ss.getSheetByName('Evaluation_Master');
      if (evalSheet) {
        const headers = evalSheet.getRange(1, 1, 1, evalSheet.getLastColumn()).getValues()[0];
        const urlColumns = {
          32: headers[31], // AF列
          33: headers[32], // AG列
          35: headers[34], // AI列
          36: headers[35]  // AJ列
        };

        Logger.log('  AF列（32）: ' + urlColumns[32]);
        Logger.log('  AG列（33）: ' + urlColumns[33]);
        Logger.log('  AI列（35）: ' + urlColumns[35]);
        Logger.log('  AJ列（36）: ' + urlColumns[36]);

        if (urlColumns[35] === '評価レポートURL' && urlColumns[36] === '戦略レポートURL') {
          Logger.log('  ✅ URL列が正しい位置にあります（AI列35、AJ列36）');
          results.functions.columnStructure = true;
        } else {
          Logger.log('  ❌ URL列の位置が正しくありません');
          results.functions.columnStructure = false;
          results.errors.push('URL列の位置が正しくありません');
        }
      }
    } catch (error) {
      Logger.log('  ❌ 列構造確認エラー: ' + error.message);
      results.errors.push('列構造確認エラー: ' + error.message);
    }
    Logger.log('');

    // 5-2. getSummary関数の確認
    Logger.log('5-2. getSummary関数の確認');
    try {
      if (typeof getSummary === 'function') {
        Logger.log('  ✅ getSummary関数: 存在');
        results.functions.getSummary = true;
      } else {
        Logger.log('  ❌ getSummary関数: 存在しない');
        results.functions.getSummary = false;
      }
    } catch (error) {
      Logger.log('  ❌ getSummary確認エラー: ' + error.message);
      results.functions.getSummary = false;
    }
    Logger.log('');

    // 5-3. callDifyWorkflow関数の確認
    Logger.log('5-3. callDifyWorkflow関数の確認');
    try {
      if (typeof callDifyWorkflow === 'function') {
        Logger.log('  ✅ callDifyWorkflow関数: 存在');
        results.functions.callDifyWorkflow = true;
      } else {
        Logger.log('  ❌ callDifyWorkflow関数: 存在しない');
        results.functions.callDifyWorkflow = false;
      }
    } catch (error) {
      Logger.log('  ❌ callDifyWorkflow確認エラー: ' + error.message);
      results.functions.callDifyWorkflow = false;
    }
    Logger.log('');

    // 5-4. generateReportV2関数の確認
    Logger.log('5-4. generateReportV2関数の確認');
    try {
      if (typeof generateReportV2 === 'function') {
        Logger.log('  ✅ generateReportV2関数: 存在');
        results.functions.generateReportV2 = true;
      } else {
        Logger.log('  ❌ generateReportV2関数: 存在しない');
        results.functions.generateReportV2 = false;
      }
    } catch (error) {
      Logger.log('  ❌ generateReportV2確認エラー: ' + error.message);
      results.functions.generateReportV2 = false;
    }
    Logger.log('');

  } catch (error) {
    Logger.log('❌ テスト中にエラーが発生しました: ' + error.message);
    results.errors.push('テスト実行エラー: ' + error.message);
  }

  // ========================================
  // 6. テスト結果サマリー
  // ========================================
  Logger.log('');
  Logger.log('='.repeat(60));
  Logger.log('=== テスト結果サマリー ===');
  Logger.log('='.repeat(60));
  Logger.log('');

  Logger.log('【基本機能】');
  Logger.log('  スプレッドシート: ' + (results.spreadsheet ? '✅ OK' : '❌ NG'));
  Logger.log('  Candidates_Master: ' + (results.sheets.Candidates_Master ? '✅ OK' : '❌ NG'));
  Logger.log('  Evaluation_Master: ' + (results.sheets.Evaluation_Master ? '✅ OK' : '❌ NG'));
  Logger.log('  Engagement_Log: ' + (results.sheets.Engagement_Log ? '✅ OK' : '❌ NG'));
  Logger.log('  ドライブフォルダ: ' + (results.driveFolder ? '✅ OK' : '❌ NG'));
  Logger.log('  企業名設定: ' + (results.companyName ? '✅ OK' : '⚠️ 未設定'));
  Logger.log('');

  Logger.log('【Phase A2 機能】');
  Logger.log('  列構造: ' + (results.functions.columnStructure ? '✅ OK' : '❌ NG'));
  Logger.log('  getSummary: ' + (results.functions.getSummary ? '✅ OK' : '❌ NG'));
  Logger.log('  callDifyWorkflow: ' + (results.functions.callDifyWorkflow ? '✅ OK' : '❌ NG'));
  Logger.log('  generateReportV2: ' + (results.functions.generateReportV2 ? '✅ OK' : '❌ NG'));
  Logger.log('');

  if (results.errors.length > 0) {
    Logger.log('【エラー・警告】');
    results.errors.forEach((error, index) => {
      Logger.log(`  ${index + 1}. ${error}`);
    });
    Logger.log('');
  }

  // 総合判定
  const allPassed = results.spreadsheet &&
                   results.sheets.Candidates_Master &&
                   results.sheets.Evaluation_Master &&
                   results.sheets.Engagement_Log &&
                   results.driveFolder &&
                   results.functions.columnStructure &&
                   results.functions.getSummary &&
                   results.functions.callDifyWorkflow &&
                   results.functions.generateReportV2;

  Logger.log('='.repeat(60));
  if (allPassed && results.companyName) {
    Logger.log('🎉 Phase A 統合テスト: 完全合格');
    Logger.log('   次のステップ: デプロイ（Webhook URL取得）');
  } else if (allPassed && !results.companyName) {
    Logger.log('✅ Phase A 統合テスト: 合格（企業名未設定）');
    Logger.log('   次のステップ: setCompanyName() → デプロイ');
  } else {
    Logger.log('❌ Phase A 統合テスト: 不合格');
    Logger.log('   上記のエラーを修正してください');
  }
  Logger.log('='.repeat(60));
  Logger.log('');

  return results;
}

/**
 * 初期設定を一括実行
 *
 * 実行内容:
 * 1. ドライブ構造初期化
 * 2. 企業名設定
 * 3. 統合テスト
 *
 * 実行手順:
 * 1. Apps Scriptエディタを開く
 * 2. 関数選択で「runInitialSetup」を選択
 * 3. 実行ボタンをクリック
 * 4. ログを確認
 *
 * @param {string} companyName - 企業名（デフォルト: "アマネク"）
 */
function runInitialSetup(companyName = "アマネク") {
  Logger.log('');
  Logger.log('='.repeat(60));
  Logger.log('=== Phase A 初期設定 一括実行 ===');
  Logger.log('='.repeat(60));
  Logger.log('');

  try {
    // ステップ1: ドライブ構造初期化
    Logger.log('【ステップ1/3】ドライブ構造初期化');
    Logger.log('');
    const driveResult = initializeDriveStructure(companyName);
    Logger.log('');

    // 5秒待機（Drive APIのレート制限対策）
    Logger.log('⏳ 5秒待機中...');
    Utilities.sleep(5000);
    Logger.log('');

    // ステップ2: 企業名設定
    Logger.log('【ステップ2/3】企業名設定');
    Logger.log('');
    const nameResult = setCompanyName(companyName);
    Logger.log('');

    // ステップ3: 統合テスト
    Logger.log('【ステップ3/3】統合テスト');
    Logger.log('');
    const testResult = testPhaseAComplete();
    Logger.log('');

    // 最終結果
    Logger.log('='.repeat(60));
    Logger.log('=== 初期設定完了 ===');
    Logger.log('='.repeat(60));
    Logger.log('');
    Logger.log('✅ ステップ1: ドライブ構造初期化 → 完了');
    Logger.log('✅ ステップ2: 企業名設定 → 完了');
    Logger.log('✅ ステップ3: 統合テスト → 完了');
    Logger.log('');
    Logger.log('📝 次のステップ:');
    Logger.log('   1. Apps Scriptをデプロイ（ウェブアプリ）');
    Logger.log('   2. Webhook URLを取得');
    Logger.log('   3. DifyワークフローにWebhook URLを設定');
    Logger.log('');

    return {
      success: true,
      driveResult: driveResult,
      nameResult: nameResult,
      testResult: testResult
    };

  } catch (error) {
    Logger.log('');
    Logger.log('❌ 初期設定エラー: ' + error.message);
    Logger.log('');
    return {
      success: false,
      error: error.message
    };
  }
}
