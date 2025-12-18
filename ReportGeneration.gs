/**
 * ReportGeneration.gs
 * Phase A: レポート生成機能
 */

// ========================================
// Task A1-1: Evaluation_Masterシート列拡張
// ========================================

/**
 * Evaluation_Masterシートに8列を追加
 * 追加する列:
 * - AC列: 次回質問1
 * - AD列: 次回質問2
 * - AE列: 次回質問3
 * - AF列: 次回質問4
 * - AG列: 次回質問5
 * - AH列: 競合企業分析
 * - AI列: 評価レポートURL
 * - AJ列: 戦略レポートURL
 */
function extendEvaluationMasterColumns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Evaluation_Master');

    if (!sheet) {
      throw new Error('Evaluation_Masterシートが見つかりません');
    }

    Logger.log('=== Evaluation_Master列拡張開始 ===');

    // 現在のヘッダー行を取得
    const currentLastCol = sheet.getLastColumn();
    Logger.log('現在の列数: ' + currentLastCol);

    if (currentLastCol === 0) {
      throw new Error('Evaluation_Masterにヘッダーがありません');
    }

    const headers = sheet.getRange(1, 1, 1, currentLastCol).getValues()[0];
    Logger.log('既存の最後の列: ' + headers[headers.length - 1]);

    // 新しいヘッダーを追加（AC列以降）
    const newHeaders = [
      '次回質問1',
      '次回質問2',
      '次回質問3',
      '次回質問4',
      '次回質問5',
      '競合企業分析',
      '評価レポートURL',
      '戦略レポートURL'
    ];

    const newStartCol = currentLastCol + 1;

    // ヘッダー行に新しい列を追加
    sheet.getRange(1, newStartCol, 1, newHeaders.length).setValues([newHeaders]);

    // ヘッダーの書式設定
    const newHeaderRange = sheet.getRange(1, newStartCol, 1, newHeaders.length);
    newHeaderRange.setBackground('#4A90E2');
    newHeaderRange.setFontColor('#FFFFFF');
    newHeaderRange.setFontWeight('bold');
    newHeaderRange.setHorizontalAlignment('center');
    newHeaderRange.setVerticalAlignment('middle');

    // 列幅の設定
    const columnWidths = {
      0: 300,  // AC: 次回質問1
      1: 300,  // AD: 次回質問2
      2: 300,  // AE: 次回質問3
      3: 300,  // AF: 次回質問4
      4: 300,  // AG: 次回質問5
      5: 400,  // AH: 競合企業分析
      6: 300,  // AI: 評価レポートURL
      7: 300   // AJ: 戦略レポートURL
    };

    for (let i = 0; i < newHeaders.length; i++) {
      sheet.setColumnWidth(newStartCol + i, columnWidths[i]);
    }

    // テキスト折り返し設定（質問列）
    const questionRange = sheet.getRange(2, newStartCol, 1000, 5);
    questionRange.setWrap(true);

    Logger.log('✅ 列追加完了: ' + newHeaders.join(', '));
    Logger.log('✅ 新しい列範囲: ' + String.fromCharCode(64 + newStartCol) + '-' +
               String.fromCharCode(64 + newStartCol + newHeaders.length - 1));
    Logger.log('=== Evaluation_Master列拡張完了 ===');

    return {
      success: true,
      added_columns: newHeaders.length,
      new_column_range: String.fromCharCode(64 + newStartCol) + '-' +
                        String.fromCharCode(64 + newStartCol + newHeaders.length - 1),
      headers: newHeaders
    };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    throw error;
  }
}

/**
 * 列拡張のテスト
 */
function testExtendEvaluationMasterColumns() {
  try {
    Logger.log('=== Evaluation_Master列拡張テスト開始 ===');

    const result = extendEvaluationMasterColumns();

    Logger.log('テスト結果:');
    Logger.log(JSON.stringify(result, null, 2));

    // 検証
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Evaluation_Master');
    const newLastCol = sheet.getLastColumn();

    Logger.log('✅ テスト成功');
    Logger.log('新しい列数: ' + newLastCol);

    // ヘッダーを表示
    const allHeaders = sheet.getRange(1, 1, 1, newLastCol).getValues()[0];
    Logger.log('全ヘッダー（最後の8列）:');
    for (let i = newLastCol - 8; i < newLastCol; i++) {
      Logger.log('  ' + String.fromCharCode(65 + i) + '列: ' + allHeaders[i]);
    }

  } catch (error) {
    Logger.log('❌ テスト失敗: ' + error.message);
  }
}

// ========================================
// Task A1-2: Report_Templatesシート作成
// ========================================

/**
 * Report_Templatesシートを作成
 */
function createReportTemplatesSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    Logger.log('=== Report_Templatesシート作成開始 ===');

    // 既存シートを削除（存在する場合）
    const existingSheet = ss.getSheetByName('Report_Templates');
    if (existingSheet) {
      Logger.log('既存のReport_Templatesシートを削除します');
      ss.deleteSheet(existingSheet);
    }

    // 新規シート作成
    const sheet = ss.insertSheet('Report_Templates');

    // ヘッダー設定
    const headers = [
      'template_id',
      'template_name',
      'template_type',
      'google_doc_template_id',
      'created_at',
      'updated_at',
      'is_active'
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // ヘッダーの書式設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');

    // 初期データ（テンプレートIDは後で設定）
    const initialData = [
      [
        'eval_report_v1',
        '面接評価レポート',
        '評価レポート',
        '',  // テンプレートID（後で設定）
        new Date(),
        new Date(),
        true
      ],
      [
        'strategy_report_v1',
        'エンゲージメント戦略レポート',
        '戦略レポート',
        '',  // テンプレートID（後で設定）
        new Date(),
        new Date(),
        true
      ]
    ];

    sheet.getRange(2, 1, initialData.length, headers.length).setValues(initialData);

    // 列幅設定
    sheet.setColumnWidth(1, 150);  // A: template_id
    sheet.setColumnWidth(2, 250);  // B: template_name
    sheet.setColumnWidth(3, 150);  // C: template_type
    sheet.setColumnWidth(4, 400);  // D: google_doc_template_id
    sheet.setColumnWidth(5, 160);  // E: created_at
    sheet.setColumnWidth(6, 160);  // F: updated_at
    sheet.setColumnWidth(7, 100);  // G: is_active

    // 日付フォーマット
    sheet.getRange('E2:F1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');

    // フォーマット
    sheet.setFrozenRows(1);

    Logger.log('✅ Report_Templatesシート作成完了');
    Logger.log('初期データ: ' + initialData.length + '行追加');
    Logger.log('=== Report_Templatesシート作成完了 ===');

    return {
      success: true,
      sheet_name: 'Report_Templates',
      initial_rows: initialData.length
    };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    throw error;
  }
}

/**
 * Report_Templatesシート作成のテスト
 */
function testCreateReportTemplatesSheet() {
  try {
    Logger.log('=== Report_Templatesシート作成テスト開始 ===');

    const result = createReportTemplatesSheet();

    Logger.log('テスト結果:');
    Logger.log(JSON.stringify(result, null, 2));

    // 検証
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Report_Templates');

    if (!sheet) {
      throw new Error('Report_Templatesシートが作成されませんでした');
    }

    const data = sheet.getDataRange().getValues();
    Logger.log('✅ テスト成功');
    Logger.log('シート名: Report_Templates');
    Logger.log('行数: ' + data.length + ' (ヘッダー含む)');
    Logger.log('列数: ' + data[0].length);

  } catch (error) {
    Logger.log('❌ テスト失敗: ' + error.message);
  }
}

// ========================================
// Task A2-1: Googleドキュメント生成関数
// ========================================

/**
 * 面接評価レポートを生成
 *
 * @param {Object} data - 評価データ
 * @return {string} 生成されたGoogleドキュメントのURL
 */
function generateEvaluationReport(data) {
  try {
    Logger.log('=== 面接評価レポート生成開始 ===');

    // 1. 新規Googleドキュメント作成
    const docTitle = `【面接評価】${data.candidate_name}様_${data.selection_phase}_${Utilities.formatDate(new Date(), 'JST', 'yyyy_MM_dd')}`;
    const doc = DocumentApp.create(docTitle);
    const body = doc.getBody();

    // 2. ヘッダー
    body.appendParagraph('面接評価レポート').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('').appendHorizontalRule();

    // 3. 基本情報
    body.appendParagraph('基本情報').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(`候補者名: ${data.candidate_name}`);
    body.appendParagraph(`選考フェーズ: ${data.selection_phase}`);
    body.appendParagraph(`面接日: ${data.interview_date || ''}`);
    body.appendParagraph(`面接官: ${data.interviewer || ''}`);
    body.appendParagraph('');

    // 4. 総合評価
    body.appendParagraph('総合評価').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(`総合ランク: ${data.total_rank || ''}`).setBold(true);
    body.appendParagraph(`総合スコア: ${data.total_score || 0}点`).setBold(true);
    body.appendParagraph('');
    body.appendParagraph('【総評】');
    body.appendParagraph(data.summary || '');
    body.appendParagraph('');

    // 5. 4軸評価
    body.appendParagraph('4軸評価').setHeading(DocumentApp.ParagraphHeading.HEADING2);

    const axes = [
      { name: 'Philosophy（理念共感）', rank: data.philosophy_rank, score: data.philosophy_score, reason: data.philosophy_reason },
      { name: 'Strategy（戦略理解）', rank: data.strategy_rank, score: data.strategy_score, reason: data.strategy_reason },
      { name: 'Motivation（動機）', rank: data.motivation_rank, score: data.motivation_score, reason: data.motivation_reason },
      { name: 'Execution（実行力）', rank: data.execution_rank, score: data.execution_score, reason: data.execution_reason }
    ];

    axes.forEach(axis => {
      body.appendParagraph(`■ ${axis.name}`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      body.appendParagraph(`ランク: ${axis.rank || ''} / スコア: ${axis.score || 0}点`).setBold(true);
      body.appendParagraph(`評価理由: ${axis.reason || ''}`);
      body.appendParagraph('');
    });

    // 6. 次回確認事項
    if (data.next_questions && data.next_questions.length > 0) {
      body.appendParagraph('次回確認事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      data.next_questions.forEach((q, index) => {
        if (q) {  // 空でない質問のみ
          body.appendParagraph(`${index + 1}. ${q}`);
        }
      });
      body.appendParagraph('');
    }

    // 7. 懸念事項
    if (data.concerns) {
      body.appendParagraph('懸念事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(data.concerns);
      body.appendParagraph('');
    }

    // 8. 次のチェックポイント
    if (data.next_check_points) {
      body.appendParagraph('次のチェックポイント').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(data.next_check_points);
    }

    // 9. ドキュメント保存
    doc.saveAndClose();

    const url = doc.getUrl();
    Logger.log('✅ 評価レポート生成成功: ' + url);

    return url;

  } catch (error) {
    Logger.log('❌ 評価レポート生成エラー: ' + error.message);
    throw error;
  }
}

/**
 * エンゲージメント戦略レポートを生成
 *
 * @param {Object} data - 戦略データ
 * @return {string} 生成されたGoogleドキュメントのURL
 */
function generateStrategyReport(data) {
  try {
    Logger.log('=== エンゲージメント戦略レポート生成開始 ===');

    // 1. 新規Googleドキュメント作成
    const docTitle = `【戦略】${data.candidate_name}様_エンゲージメント戦略_${Utilities.formatDate(new Date(), 'JST', 'yyyy_MM_dd')}`;
    const doc = DocumentApp.create(docTitle);
    const body = doc.getBody();

    // 2. ヘッダー
    body.appendParagraph('エンゲージメント戦略レポート').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('').appendHorizontalRule();

    // 3. 基本情報
    body.appendParagraph('候補者情報').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(`候補者名: ${data.candidate_name}`);
    body.appendParagraph(`現在フェーズ: ${data.current_phase || ''}`);
    body.appendParagraph('');

    // 4. 承諾可能性分析
    body.appendParagraph('承諾可能性分析').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(`現在の承諾可能性: ${data.acceptance_probability || 0}%`).setBold(true);
    body.appendParagraph(`信頼度: ${data.confidence_level || ''}`).setBold(true);
    body.appendParagraph('');

    // 5. ポジティブ要因
    if (data.positive_factors && data.positive_factors.length > 0) {
      body.appendParagraph('【ポジティブ要因】');
      data.positive_factors.forEach(factor => {
        if (factor) {
          body.appendParagraph(`✓ ${factor}`);
        }
      });
      body.appendParagraph('');
    }

    // 6. リスク要因
    if (data.risk_factors && data.risk_factors.length > 0) {
      body.appendParagraph('【リスク要因】');
      data.risk_factors.forEach(factor => {
        if (factor) {
          body.appendParagraph(`⚠ ${factor}`);
        }
      });
      body.appendParagraph('');
    }

    // 7. 承諾までのストーリー
    body.appendParagraph('承諾までのストーリー').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    if (data.acceptance_story && Array.isArray(data.acceptance_story)) {
      data.acceptance_story.forEach((step, index) => {
        if (step) {
          body.appendParagraph(`Step ${index + 1}: ${step}`);
        }
      });
    }
    body.appendParagraph('');

    // 8. エンゲージメント戦略
    body.appendParagraph('推奨アクション').setHeading(DocumentApp.ParagraphHeading.HEADING2);

    if (data.engagement_strategy) {
      if (data.engagement_strategy.immediate_action_24h) {
        body.appendParagraph('■ 24時間以内').setHeading(DocumentApp.ParagraphHeading.HEADING3);
        body.appendParagraph(data.engagement_strategy.immediate_action_24h);
        body.appendParagraph('');
      }

      if (data.engagement_strategy.followup_action_48h) {
        body.appendParagraph('■ 48時間以内').setHeading(DocumentApp.ParagraphHeading.HEADING3);
        body.appendParagraph(data.engagement_strategy.followup_action_48h);
        body.appendParagraph('');
      }

      if (data.engagement_strategy.longterm_action_72h) {
        body.appendParagraph('■ 72時間以内').setHeading(DocumentApp.ParagraphHeading.HEADING3);
        body.appendParagraph(data.engagement_strategy.longterm_action_72h);
        body.appendParagraph('');
      }
    }

    // 9. 競合状況
    if (data.competitor_analysis) {
      body.appendParagraph('競合状況').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(data.competitor_analysis);
      body.appendParagraph('');
    }

    // 10. 推奨事項
    if (data.recommendation) {
      body.appendParagraph('総合推奨事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(data.recommendation);
    }

    // 11. ドキュメント保存
    doc.saveAndClose();

    const url = doc.getUrl();
    Logger.log('✅ 戦略レポート生成成功: ' + url);

    return url;

  } catch (error) {
    Logger.log('❌ 戦略レポート生成エラー: ' + error.message);
    throw error;
  }
}

/**
 * レポート生成関数のテスト
 */
function testGenerateReports() {
  try {
    Logger.log('=== レポート生成テスト開始 ===');

    // テストデータ
    const evalData = {
      candidate_name: 'テスト太郎',
      selection_phase: '1次面接',
      interview_date: '2025/12/18',
      interviewer: '東',
      total_rank: 'A',
      total_score: 93,
      summary: '非常に優秀な候補者です。',
      philosophy_rank: 'A',
      philosophy_score: 29,
      philosophy_reason: '理念への深い共感が見られます。',
      strategy_rank: 'B',
      strategy_score: 25,
      strategy_reason: '戦略理解は十分です。',
      motivation_rank: 'A',
      motivation_score: 19,
      motivation_reason: '非常に高い志望度です。',
      execution_rank: 'A',
      execution_score: 20,
      execution_reason: '優れた実行力があります。',
      next_questions: [
        '前職での具体的な成果について',
        'チームマネジメントの経験',
        'キャリアビジョン'
      ],
      concerns: '待遇面への懸念',
      next_check_points: '次回面接で給与条件を明示'
    };

    const strategyData = {
      candidate_name: 'テスト太郎',
      current_phase: '1次面接',
      acceptance_probability: 75,
      confidence_level: 'HIGH',
      positive_factors: [
        '理念への強い共感',
        '高いスキルセット'
      ],
      risk_factors: [
        '給与への懸念',
        '競合他社の存在'
      ],
      acceptance_story: [
        '給与条件を明示',
        '社員面談を実施',
        'クロージング面談'
      ],
      engagement_strategy: {
        immediate_action_24h: '給与条件を明確に提示',
        followup_action_48h: '社員面談を設定',
        longterm_action_72h: 'クロージング面談実施'
      },
      competitor_analysis: 'リブコンサルティングとの競合',
      recommendation: '早期の内定提示を推奨'
    };

    // 評価レポート生成
    Logger.log('評価レポート生成中...');
    const evalUrl = generateEvaluationReport(evalData);
    Logger.log('評価レポートURL: ' + evalUrl);

    // 戦略レポート生成
    Logger.log('戦略レポート生成中...');
    const strategyUrl = generateStrategyReport(strategyData);
    Logger.log('戦略レポートURL: ' + strategyUrl);

    Logger.log('✅ レポート生成テスト成功');
    Logger.log('=== レポート生成テスト完了 ===');

    return {
      success: true,
      evaluation_report_url: evalUrl,
      strategy_report_url: strategyUrl
    };

  } catch (error) {
    Logger.log('❌ テスト失敗: ' + error.message);
    throw error;
  }
}
