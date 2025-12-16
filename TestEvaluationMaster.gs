/**
 * Evaluation_Master機能のテスト用関数
 */

/**
 * テスト用関数: Evaluation_Masterシート作成とデータ書き込み
 */
function testEvaluationMasterCreation() {
  try {
    // Step 1: シート作成
    Logger.log('========== Step 1: シート作成 ==========');
    createEvaluationMasterSheet();

    // Step 2: テストデータ書き込み
    Logger.log('========== Step 2: テストデータ書き込み ==========');
    const testData = {
      interview_datetime: new Date('2025-06-04 19:00'),
      candidate_id: 'CAND_001',
      candidate_name: '山田太郎',
      recruit_type: '新卒',
      selection_phase: '1次面接',
      dify_report_url: 'https://docs.google.com/document/d/sample',

      philosophy_rank: 'C',
      philosophy_score: 18,
      philosophy_reason: 'ポジティブな点: 社会貢献への意欲が見られる。\n改善点: 経営理念への深い理解が不足。',

      strategy_rank: 'B',
      strategy_score: 24,
      strategy_reason: 'ポジティブな点: 事業モデルの基本理解がある。\n改善点: 競合優位性の認識が浅い。',

      motivation_rank: 'D',
      motivation_score: 12,
      motivation_reason: 'ポジティブな点: 明確なキャリア目標がある。\n改善点: 当社固有の志望動機が弱い。',

      execution_rank: 'A',
      execution_score: 18,
      execution_reason: 'ポジティブな点: 学祭運営での高い実行力を発揮。\n改善点: ビジネス経験は未知数。',

      total_score: 72,
      total_rank: 'C',
      summary: '総合ランクC。実行力は高いが、志望動機の深掘りが必要。次回は当社への固有の関心を確認すべき。',

      transcript: '面接の文字起こし全文...',
      interview_memo: '面接官: 田中\n印象: 真面目で誠実な印象',
      concerns: '志望動機が曖昧、競合他社との比較が不十分',
      next_check_points: '当社への固有の志望理由、キャリアビジョンの詳細',
      workflow_id: 'workflow_12345'
    };

    const evaluationId = writeToEvaluationMaster(testData);
    Logger.log(`✅ Test completed: ${evaluationId}`);

    // 成功メッセージを表示（UIが利用可能な場合のみ）
    try {
      SpreadsheetApp.getUi().alert(
        '✅ テストが完了しました\n\n' +
        `評価ID: ${evaluationId}\n` +
        '候補者名: 山田太郎\n' +
        '総合スコア: 72点 (ランク: C)\n\n' +
        'Evaluation_Masterシートを確認してください。'
      );
    } catch (uiError) {
      Logger.log('✅ テスト完了（UI表示スキップ）: ' + evaluationId);
    }

  } catch (error) {
    Logger.log('❌ テストエラー: ' + error.message);
    Logger.log(error.stack);

    // エラーメッセージを表示（UIが利用可能な場合のみ）
    try {
      SpreadsheetApp.getUi().alert(
        '❌ テストに失敗しました\n\n' +
        'エラー: ' + error.message + '\n\n' +
        '詳細はログを確認してください（表示 → ログ）'
      );
    } catch (uiError) {
      // UI表示できない場合はログのみ
    }

    throw error;
  }
}

/**
 * 複数データテスト: 3件のテストデータを追加
 */
function testMultipleEvaluationData() {
  try {
    Logger.log('========== 複数データテスト開始 ==========');

    const testDataSet = [
      {
        interview_datetime: new Date('2025-06-05 10:00'),
        candidate_id: 'CAND_002',
        candidate_name: '佐藤花子',
        recruit_type: '新卒',
        selection_phase: '2次面接',
        dify_report_url: 'https://docs.google.com/document/d/sample2',

        philosophy_rank: 'A',
        philosophy_score: 28,
        philosophy_reason: 'ポジティブな点: 経営理念への共感が非常に強い。',

        strategy_rank: 'A',
        strategy_score: 27,
        strategy_reason: 'ポジティブな点: 戦略の理解が深い。',

        motivation_rank: 'B',
        motivation_score: 16,
        motivation_reason: 'ポジティブな点: 志望動機が明確。',

        execution_rank: 'B',
        execution_score: 15,
        execution_reason: 'ポジティブな点: チーム活動での実行力がある。',

        total_score: 86,
        total_rank: 'A',
        summary: '総合ランクA。優秀な候補者。次回は具体的なキャリアプランを確認。',

        transcript: '2次面接の文字起こし...',
        interview_memo: '面接官: 鈴木\n印象: 優秀で意欲的',
        concerns: '特になし',
        next_check_points: 'キャリアビジョンの詳細確認',
        workflow_id: 'workflow_12346'
      },
      {
        interview_datetime: new Date('2025-06-05 14:00'),
        candidate_id: 'CAND_003',
        candidate_name: '田中一郎',
        recruit_type: '中途',
        selection_phase: '1次面接',
        dify_report_url: 'https://docs.google.com/document/d/sample3',

        philosophy_rank: 'B',
        philosophy_score: 22,
        philosophy_reason: 'ポジティブな点: 理念への理解がある。',

        strategy_rank: 'C',
        strategy_score: 19,
        strategy_reason: 'ポジティブな点: 基本的な理解はある。\n改善点: 戦略の深掘りが必要。',

        motivation_rank: 'C',
        motivation_score: 13,
        motivation_reason: 'ポジティブな点: 転職理由は明確。\n改善点: 当社への志望動機が弱い。',

        execution_rank: 'B',
        execution_score: 16,
        execution_reason: 'ポジティブな点: 前職での実績がある。',

        total_score: 70,
        total_rank: 'C',
        summary: '総合ランクC。実績はあるが、志望動機の深掘りが必要。',

        transcript: '1次面接の文字起こし...',
        interview_memo: '面接官: 高橋\n印象: 経験豊富だが志望理由が弱い',
        concerns: '志望動機の明確化が必要',
        next_check_points: '当社への志望理由を深掘り',
        workflow_id: 'workflow_12347'
      }
    ];

    const evaluationIds = [];

    testDataSet.forEach((data, index) => {
      Logger.log(`\n----- データ ${index + 1}/${testDataSet.length} を書き込み中 -----`);
      const evaluationId = writeToEvaluationMaster(data);
      evaluationIds.push(evaluationId);
      Logger.log(`✅ ${data.candidate_name}: ${evaluationId}`);
    });

    Logger.log('\n========== 複数データテスト完了 ==========');
    Logger.log(`追加した評価ID: ${evaluationIds.join(', ')}`);

    // 成功メッセージを表示（UIが利用可能な場合のみ）
    try {
      SpreadsheetApp.getUi().alert(
        '✅ 複数データテストが完了しました\n\n' +
        `追加件数: ${evaluationIds.length}件\n\n` +
        '評価ID:\n' + evaluationIds.join('\n') + '\n\n' +
        'Evaluation_Masterシートを確認してください。'
      );
    } catch (uiError) {
      Logger.log('✅ 複数データテスト完了（UI表示スキップ）');
    }

  } catch (error) {
    Logger.log('❌ 複数データテストエラー: ' + error.message);
    Logger.log(error.stack);

    // エラーメッセージを表示（UIが利用可能な場合のみ）
    try {
      SpreadsheetApp.getUi().alert(
        '❌ テストに失敗しました\n\n' +
        'エラー: ' + error.message
      );
    } catch (uiError) {
      // UI表示できない場合はログのみ
    }

    throw error;
  }
}
