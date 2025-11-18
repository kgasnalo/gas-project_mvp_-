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
