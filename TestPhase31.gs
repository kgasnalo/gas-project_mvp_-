/**
 * TestPhase31.gs
 * Phase 3-1: 基礎要素スコア計算のテスト
 *
 * @version 1.0
 * @date 2025-11-17
 */

/**
 * テスト1: メールアドレス取得
 */
function testGetCandidateEmail() {
  Logger.log('\n=== メールアドレス取得のテスト ===');

  const email = getCandidateEmail('C001');
  Logger.log(`C001のメール: ${email}`);
  // 期待値: tanaka@example.com など
}

/**
 * テスト2: 志望度スコア
 */
function testMotivationScore() {
  Logger.log('\n=== 志望度スコアのテスト ===');

  const score = calculateMotivationScore('C001', '初回面談');
  Logger.log(`C001の志望度スコア（初回面談）: ${score}`);
  // 期待値: 70-100点
}

/**
 * テスト3: 競合優位性スコア
 */
function testCompetitiveScore() {
  Logger.log('\n=== 競合優位性スコアのテスト ===');

  const score = calculateCompetitiveAdvantageScore('C001', '初回面談');
  Logger.log(`C001の競合優位性スコア: ${score}`);
  // 期待値: 30-70点
}

/**
 * テスト4: 懸念解消度スコア
 */
function testConcernScore() {
  Logger.log('\n=== 懸念解消度スコアのテスト ===');

  const score = calculateConcernResolutionScore('C001', '初回面談');
  Logger.log(`C001の懸念解消度スコア: ${score}`);
  // 期待値: 50-100点
}

/**
 * テスト5: 基礎要素スコア（統合）
 */
function testFoundationScore() {
  Logger.log('\n=== 基礎要素スコアのテスト ===');

  const score = calculateFoundationScore('C001', '初回面談');
  Logger.log(`\nC001の基礎要素スコア: ${score}点`);
  // 期待値: 50-85点
}

/**
 * 全テストを実行
 */
function runAllPhase31Tests() {
  Logger.log('\n========================================');
  Logger.log('Phase 3-1: 全テスト実行開始');
  Logger.log('========================================\n');

  try {
    testGetCandidateEmail();
    testMotivationScore();
    testCompetitiveScore();
    testConcernScore();
    testFoundationScore();

    Logger.log('\n========================================');
    Logger.log('✅ すべてのテストが完了しました');
    Logger.log('========================================\n');
  } catch (error) {
    Logger.log(`\n❌ テスト実行エラー: ${error}`);
    Logger.log(error.stack);
  }
}

/**
 * 統合テスト（利用可能なデータのみ）
 */
function testAllFoundationScores() {
  Logger.log('\n========================================');
  Logger.log('全候補者の基礎要素スコアテスト');
  Logger.log('========================================\n');

  const candidates = ['C001']; // まずはC001のみ
  const phases = ['初回面談']; // まずは初回面談のみ

  const results = [];

  for (let candidate of candidates) {
    for (let phase of phases) {
      try {
        const score = calculateFoundationScore(candidate, phase);
        results.push({
          candidate: candidate,
          phase: phase,
          score: score,
          status: 'SUCCESS'
        });
      } catch (error) {
        results.push({
          candidate: candidate,
          phase: phase,
          score: null,
          status: 'ERROR',
          error: error.toString()
        });
      }
    }
  }

  // 結果を表示
  Logger.log('\n=== テスト結果サマリー ===');
  for (let result of results) {
    if (result.status === 'SUCCESS') {
      Logger.log(`✅ ${result.candidate} | ${result.phase} | ${result.score}点`);
    } else {
      Logger.log(`❌ ${result.candidate} | ${result.phase} | エラー: ${result.error}`);
    }
  }

  Logger.log(`========================================\n`);
}

/**
 * アンケートシートのデータ確認
 *
 * 各アンケートシートに存在するデータを確認する
 */
function checkSurveyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    'アンケート_初回面談',
    'アンケート_社員面談',
    'アンケート_2次面接',
    'アンケート_内定'
  ];

  Logger.log('\n========================================');
  Logger.log('アンケートシートのデータ確認');
  Logger.log('========================================\n');

  for (let sheetName of sheets) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`❌ ${sheetName}: シートが見つかりません`);
      continue;
    }

    const data = sheet.getDataRange().getValues();
    const rowCount = data.length - 1; // ヘッダー除く

    Logger.log(`\n📊 ${sheetName}`);
    Logger.log(`   データ件数: ${rowCount}件`);

    if (rowCount > 0) {
      Logger.log(`   データ一覧:`);
      for (let i = 1; i < data.length; i++) {
        const email = data[i][2]; // C列: メールアドレス
        const timestamp = data[i][0]; // A列: タイムスタンプ
        Logger.log(`     - ${i}行目: ${email} (${timestamp})`);
      }
    } else {
      Logger.log(`   ⚠️ データなし`);
    }
  }

  Logger.log('\n========================================\n');
}

/**
 * Contact_Historyシートのデータ構造を確認（デバッグ用）
 */
function debugContactHistory() {
  Logger.log('\n========================================');
  Logger.log('Contact_History シートの構造確認');
  Logger.log('========================================\n');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONTACT_HISTORY);

    if (!sheet) {
      Logger.log('❌ Contact_Historyシートが見つかりません');
      return;
    }

    // ヘッダー行を取得（getDisplayValues()を使用してデータバリデーションエラーを回避）
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    Logger.log('=== ヘッダー行 ===');
    headers.forEach((header, index) => {
      Logger.log(`${String.fromCharCode(65 + index)}列（インデックス${index}）: ${header}`);
    });

    // データ行数を確認
    const lastRow = sheet.getLastRow();
    Logger.log(`\n=== データ行数 ===`);
    Logger.log(`最終行: ${lastRow}行目（ヘッダー除くデータ: ${lastRow - 1}件）`);

    // 最初の10件のデータを表示（getDisplayValues()を使用）
    if (lastRow > 1) {
      Logger.log('\n=== 最初の10件のデータ ===');
      const dataRows = Math.min(10, lastRow - 1);
      const data = sheet.getRange(2, 1, dataRows, sheet.getLastColumn()).getDisplayValues();

      data.forEach((row, rowIndex) => {
        Logger.log(`\n${rowIndex + 2}行目:`);
        row.forEach((cell, colIndex) => {
          const cellValue = cell === null ? 'null' : cell === '' ? '(空文字)' : cell;
          Logger.log(`  ${String.fromCharCode(65 + colIndex)}列: ${cellValue}`);
        });
      });
    }

    Logger.log('\n========================================\n');

  } catch (error) {
    Logger.log(`❌ debugContactHistoryエラー: ${error.message}`);
    Logger.log(error.stack);
  }
}

/**
 * ========================================
 * Phase 3-2: 関係性・行動シグナルスコアのテスト
 * ========================================
 */

/**
 * テスト1: 依存関数の確認（Phase 3-1の関数）
 */
function testDependencyFunctions() {
  Logger.log('\n========================================');
  Logger.log('Phase 3-2: 依存関数の確認テスト');
  Logger.log('========================================\n');

  try {
    // getProactivityScore()のテスト
    Logger.log('=== getProactivityScore() のテスト ===');
    const proactivityScore = getProactivityScore('C001', '最終面接');
    Logger.log(`C001の積極性スコア（最終面接）: ${proactivityScore}点`);
    Logger.log(`期待値: 70-100点\n`);

    // getAverageResponseSpeedScore()のテスト
    Logger.log('=== getAverageResponseSpeedScore() のテスト ===');
    const avgResponseScore = getAverageResponseSpeedScore('C001', '初回面談');
    Logger.log(`C001の平均回答速度スコア（初回面談）: ${avgResponseScore}点`);
    Logger.log(`期待値: 0-100点\n`);

    Logger.log('✅ 依存関数の確認テストが完了しました\n');

  } catch (error) {
    Logger.log(`❌ 依存関数テストエラー: ${error.message}`);
    Logger.log(error.stack);
  }
}

/**
 * テスト2: ヘルパー関数のテスト
 */
function testHelperFunctions() {
  Logger.log('\n========================================');
  Logger.log('Phase 3-2: ヘルパー関数のテスト');
  Logger.log('========================================\n');

  const candidateId = 'C001';

  try {
    // 1. 接点履歴の取得
    Logger.log('=== 接点履歴の取得 ===');
    const contacts = getContactHistory(candidateId);
    Logger.log(`接点履歴件数: ${contacts.length}件`);
    Logger.log(`期待値: C001は8件\n`);

    if (contacts.length > 0) {
      Logger.log('最初の接点:');
      Logger.log(`  日付: ${contacts[0].date}`);
      Logger.log(`  タイプ: ${contacts[0].type}`);
      Logger.log(`  担当者: ${contacts[0].assignee}\n`);

      Logger.log('最後の接点:');
      Logger.log(`  日付: ${contacts[contacts.length - 1].date}`);
      Logger.log(`  タイプ: ${contacts[contacts.length - 1].type}`);
      Logger.log(`  担当者: ${contacts[contacts.length - 1].assignee}\n`);
    }

    // 2. 最新接点日の取得
    Logger.log('=== 最新接点日の取得 ===');
    const latestDate = getLatestContactDate(candidateId);
    Logger.log(`最新接点日: ${latestDate ? latestDate.toISOString().split('T')[0] : 'なし'}`);
    Logger.log(`期待値: 2025-11-10 (現在の3日前程度)\n`);

    // 3. 平均接点間隔の取得
    Logger.log('=== 平均接点間隔の取得 ===');
    const avgInterval = getAverageInterval(candidateId);
    Logger.log(`平均接点間隔: ${avgInterval.toFixed(1)}日`);
    Logger.log(`期待値: 7-10日\n`);

    // 4. 選考期間の取得
    Logger.log('=== 選考期間の取得 ===');
    const duration = getSelectionDuration(candidateId);
    Logger.log(`選考期間: ${duration.toFixed(1)}日`);
    Logger.log(`期待値: 応募日から現在までの日数\n`);

    Logger.log('✅ ヘルパー関数のテストが完了しました\n');

  } catch (error) {
    Logger.log(`❌ ヘルパー関数テストエラー: ${error.message}`);
    Logger.log(error.stack);
  }
}

/**
 * テスト3: 関係性要素スコアのテスト
 */
function testRelationshipScore() {
  Logger.log('\n========================================');
  Logger.log('Phase 3-2: 関係性要素スコアのテスト');
  Logger.log('========================================\n');

  const candidates = ['C001', 'C002', 'C003', 'C004', 'C005'];

  try {
    Logger.log('=== 各候補者の関係性要素スコア ===\n');

    for (let candidateId of candidates) {
      const score = calculateRelationshipScore(candidateId);
      Logger.log(`${candidateId}: ${score}点`);
    }

    Logger.log('\n期待値:');
    Logger.log('  C001: 85-95点（接点8件、間隔7-10日、空白3日）');
    Logger.log('  C002: 60-70点（接点5件、間隔12-16日、空白8日）');
    Logger.log('  C003: 95-100点（接点10件、間隔5-7日、空白1日）← 最高');
    Logger.log('  C004: 30-40点（接点3件、間隔20-25日、空白16日）← 最低');
    Logger.log('  C005: 70-80点（接点6件、間隔8-12日、空白5日）\n');

    Logger.log('✅ 関係性要素スコアのテストが完了しました\n');

  } catch (error) {
    Logger.log(`❌ 関係性要素スコアテストエラー: ${error.message}`);
    Logger.log(error.stack);
  }
}

/**
 * テスト4: 行動シグナル要素スコアのテスト
 */
function testBehaviorScore() {
  Logger.log('\n========================================');
  Logger.log('Phase 3-2: 行動シグナル要素スコアのテスト');
  Logger.log('========================================\n');

  const testCases = [
    { candidateId: 'C001', phase: '初回面談' },
    { candidateId: 'C002', phase: '社員面談' },
    { candidateId: 'C003', phase: '2次面接' }
  ];

  try {
    Logger.log('=== 各候補者の行動シグナルスコア ===\n');

    for (let testCase of testCases) {
      const score = calculateBehaviorScore(testCase.candidateId, testCase.phase);
      Logger.log(`${testCase.candidateId} (${testCase.phase}): ${score}点`);
    }

    Logger.log('\n期待値: アンケートデータに基づいて算出');
    Logger.log('  - 回答速度スコア: 0-100点');
    Logger.log('  - 積極性スコア: 0-100点');
    Logger.log('  - 記述充実度スコア: 0-100点\n');

    Logger.log('✅ 行動シグナル要素スコアのテストが完了しました\n');

  } catch (error) {
    Logger.log(`❌ 行動シグナル要素スコアテストエラー: ${error.message}`);
    Logger.log(error.stack);
  }
}

/**
 * テスト5: Phase 3 統合テスト（基礎+関係性+行動シグナル）
 */
function testPhase3IntegrationScore() {
  Logger.log('\n========================================');
  Logger.log('Phase 3 統合テスト: 全要素スコア計算');
  Logger.log('========================================\n');

  const candidates = ['C001', 'C002', 'C003', 'C004', 'C005'];
  const phases = ['初回面談', '社員面談', '2次面接', '内定後'];

  try {
    Logger.log('=== 各候補者×フェーズの統合スコア ===\n');
    Logger.log('候補者ID | フェーズ | 基礎(40%) | 関係性(30%) | 行動(30%) | 合計');
    Logger.log('---------|----------|-----------|-------------|-----------|-------');

    for (let candidateId of candidates) {
      for (let phase of phases) {
        try {
          // 基礎要素スコア（40%）
          const foundationScore = calculateFoundationScore(candidateId, phase);
          const foundationWeighted = foundationScore * 0.4;

          // 関係性要素スコア（30%）
          const relationshipScore = calculateRelationshipScore(candidateId);
          const relationshipWeighted = relationshipScore * 0.3;

          // 行動シグナル要素スコア（30%）
          const behaviorScore = calculateBehaviorScore(candidateId, phase);
          const behaviorWeighted = behaviorScore * 0.3;

          // 統合スコア
          const totalScore = foundationWeighted + relationshipWeighted + behaviorWeighted;

          Logger.log(`${candidateId} | ${phase.padEnd(10)} | ${foundationScore.toFixed(1).padStart(5)}点 | ${relationshipScore.toFixed(1).padStart(5)}点 | ${behaviorScore.toFixed(1).padStart(5)}点 | ${totalScore.toFixed(1)}点`);

        } catch (error) {
          Logger.log(`${candidateId} | ${phase.padEnd(10)} | エラー: ${error.message}`);
        }
      }
    }

    Logger.log('\n期待される統合スコア範囲: 50-85点');
    Logger.log('  - 基礎要素(40%): 志望度、競合優位性、懸念解消度');
    Logger.log('  - 関係性要素(30%): 接点回数、接点間隔、最新接点からの空白期間');
    Logger.log('  - 行動シグナル要素(30%): 回答速度、積極性、記述充実度\n');

    Logger.log('✅ Phase 3 統合テストが完了しました\n');

  } catch (error) {
    Logger.log(`❌ Phase 3 統合テストエラー: ${error.message}`);
    Logger.log(error.stack);
  }
}

/**
 * Phase 3-2 全テストを実行
 */
function runAllPhase32Tests() {
  Logger.log('\n========================================');
  Logger.log('Phase 3-2: 全テスト実行開始');
  Logger.log('========================================\n');

  try {
    // ステップ1: ダミーデータを投入
    Logger.log('=== ステップ1: ダミーデータ投入 ===');
    insertContactHistoryData();
    Logger.log('');

    // ステップ1.5: データ投入結果を確認（デバッグ）
    debugContactHistory();

    // ステップ2: 依存関数の確認
    testDependencyFunctions();

    // ステップ3: ヘルパー関数のテスト
    testHelperFunctions();

    // ステップ4: 関係性要素スコアのテスト
    testRelationshipScore();

    // ステップ5: 行動シグナル要素スコアのテスト
    testBehaviorScore();

    // ステップ6: 統合テスト
    testPhase3IntegrationScore();

    Logger.log('\n========================================');
    Logger.log('✅ すべてのPhase 3-2テストが完了しました');
    Logger.log('========================================\n');

  } catch (error) {
    Logger.log(`\n❌ テスト実行エラー: ${error.message}`);
    Logger.log(error.stack);
  }
}

/**
 * ========================================
 * Phase 3.5: 包括的テスト関数
 * ========================================
 */

/**
 * Phase 3.5のテストデータ投入テスト
 */
function testExpandedDataInsertion() {
  Logger.log('\n=== Phase 3.5テストデータ投入テスト ===');

  // データ投入実行
  runExpandedDataInsertion();

  // 投入結果の確認
  Logger.log('\n>>> 投入結果の確認');

  // 1. Candidates_Master
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);
  const masterData = master.getDataRange().getValues();

  for (let i = 1; i <= 5; i++) {
    const candidateId = `C00${i}`;
    const coreMotivation = masterData[i][24]; // Y列
    const topConcern = masterData[i][25]; // Z列
    Logger.log(`${candidateId}: ${coreMotivation} / ${topConcern}`);
  }

  // 2. アンケート回答数
  const survey初回 = ss.getSheetByName('アンケート_初回面談');
  Logger.log(`\n初回面談アンケート: ${survey初回.getLastRow() - 1}件`);

  const survey社員 = ss.getSheetByName('アンケート_社員面談');
  Logger.log(`社員面談アンケート: ${survey社員.getLastRow() - 1}件`);

  const survey2次 = ss.getSheetByName('アンケート_2次面接');
  Logger.log(`2次面接アンケート: ${survey2次.getLastRow() - 1}件`);

  const survey内定 = ss.getSheetByName('アンケート_内定');
  Logger.log(`内定後アンケート: ${survey内定.getLastRow() - 1}件`);

  // 3. 接点履歴
  const contactHistory = ss.getSheetByName(CONFIG.SHEET_NAMES.CONTACT_HISTORY);
  Logger.log(`\n接点履歴: ${contactHistory.getLastRow() - 1}件`);
}

/**
 * Phase 3.5包括的テスト
 *
 * 目標: 候補者5名 × 4フェーズ = 20パターンのテスト
 */
function runComprehensivePhase3Tests() {
  Logger.log('\n========================================');
  Logger.log('Phase 3.5 包括的テスト開始');
  Logger.log('========================================\n');

  const candidates = ['C001', 'C002', 'C003', 'C004', 'C005'];
  const phases = ['初回面談', '社員面談', '2次面接', '内定後'];

  const results = [];
  let successCount = 0;
  let errorCount = 0;

  for (let candidate of candidates) {
    Logger.log(`\n--- ${candidate} ---`);

    for (let phase of phases) {
      try {
        const acceptanceRate = calculateAcceptanceRate(candidate, phase);

        results.push({
          candidate: candidate,
          phase: phase,
          acceptanceRate: acceptanceRate,
          status: 'SUCCESS'
        });

        Logger.log(`✅ ${phase}: ${acceptanceRate.toFixed(1)}点`);
        successCount++;

      } catch (error) {
        results.push({
          candidate: candidate,
          phase: phase,
          status: 'ERROR',
          error: error.toString()
        });

        Logger.log(`❌ ${phase}: エラー - ${error.message}`);
        errorCount++;
      }
    }
  }

  // 結果サマリー
  Logger.log('\n========================================');
  Logger.log('テスト結果サマリー');
  Logger.log('========================================');
  Logger.log(`成功: ${successCount}/${successCount + errorCount}`);
  Logger.log(`失敗: ${errorCount}/${successCount + errorCount}`);

  if (errorCount === 0) {
    Logger.log('\n✅ 全テスト成功！');
  } else {
    Logger.log('\n❌ エラーが発生しました。Error_Logを確認してください。');
  }

  // 詳細結果
  Logger.log('\n--- 詳細結果 ---');
  for (let result of results) {
    if (result.status === 'SUCCESS') {
      Logger.log(`✅ ${result.candidate} | ${result.phase} | ${result.acceptanceRate.toFixed(1)}点`);
    } else {
      Logger.log(`❌ ${result.candidate} | ${result.phase} | ${result.error}`);
    }
  }

  Logger.log('\n========================================\n');

  return {
    successCount: successCount,
    errorCount: errorCount,
    results: results
  };
}

/**
 * 全候補者のEngagement_Log書き込みテスト
 */
function testWriteAllToEngagementLog() {
  Logger.log('\n=== 全候補者のEngagement_Log書き込みテスト ===');

  const candidates = ['C001', 'C002', 'C003', 'C004', 'C005'];
  const phases = ['初回面談', '社員面談', '2次面接', '内定後'];

  let successCount = 0;
  let errorCount = 0;

  for (let candidate of candidates) {
    for (let phase of phases) {
      const success = writeToEngagementLog(candidate, phase);

      if (success) {
        successCount++;
        Logger.log(`✅ ${candidate} - ${phase}: 書き込み成功`);
      } else {
        errorCount++;
        Logger.log(`❌ ${candidate} - ${phase}: 書き込み失敗`);
      }
    }
  }

  Logger.log(`\n成功: ${successCount}件, 失敗: ${errorCount}件`);
  Logger.log('\n⚠️ Engagement_Logシートを開いて、20件のデータが追加されていることを確認してください');
}

/**
 * トリガーのテスト（手動実行）
 *
 * 注意: 実際のアンケート送信をシミュレート
 */
function testTriggerManually() {
  Logger.log('\n=== トリガーのテスト（手動） ===');

  // シミュレーションイベントオブジェクト
  const mockEvent = {
    values: [
      new Date(), // タイムスタンプ
      '田中太郎', // 名前
      'tanaka@example.com', // メールアドレス
      8, // その他の回答...
    ]
  };

  // 初回面談トリガーをテスト
  try {
    onFormSubmit初回面談(mockEvent);
    Logger.log('✅ 初回面談トリガー: 成功');
  } catch (error) {
    Logger.log(`❌ 初回面談トリガー: 失敗 - ${error.message}`);
  }
}
