/**
 * ReportGeneratorV2.gs
 *
 * 改善版レポート生成機能
 * - 3部構成（サマリー+詳細+議事録）
 * - 候補者フォルダに自動保存
 * - エビデンス引用付き
 * - リトライ機能
 */

/**
 * 評価レポート生成V2
 *
 * @param {Object} data - 評価データ
 * @param {string} recruitType - 採用区分
 * @param {string} phase - 選考フェーズ
 * @param {string} companyName - 企業名（デフォルト: "アマネク"）
 * @return {Object} 生成結果
 */
function generateEvaluationReportV2(data, recruitType, phase, companyName = "アマネク") {
  Logger.log('=== 評価レポートV2 生成開始 ===');

  return retryOperation(() => {
    try {
      // ドキュメント作成
      const dateStr = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd');
      const docTitle = `${dateStr}_${phase}_評価レポート`;
      const doc = DocumentApp.create(docTitle);
      const body = doc.getBody();

      // フォント設定
      const style = {};
      style[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
      style[DocumentApp.Attribute.FONT_SIZE] = 11;
      body.setAttributes(style);

      // Part 1: エグゼクティブサマリー
      addEvaluationSummary(body, data);

      // 改ページ
      body.appendPageBreak();

      // Part 2: 詳細評価
      addDetailedEvaluation(body, data);

      // 改ページ
      body.appendPageBreak();

      // Part 3: 議事録
      addTranscript(body, data);

      // ドキュメント保存
      doc.saveAndClose();

      // 候補者フォルダに移動
      const candidateFolder = getOrCreateCandidateFolder(
        recruitType,
        phase,
        data.candidate_id,
        data.candidate_name,
        companyName
      );

      const movedDoc = saveDocumentToFolder(doc, candidateFolder);

      // 共有設定（エラー時はスキップ）
      try {
        movedDoc.setSharing(
          DriveApp.Access.DOMAIN_WITH_LINK,
          DriveApp.Permission.VIEW
        );
      } catch (error) {
        Logger.log('⚠️ 共有権限設定をスキップ: ' + error.message);
      }

      const url = movedDoc.getUrl();
      Logger.log('✅ 評価レポート生成成功: ' + url);

      return {
        success: true,
        url: url,
        documentId: movedDoc.getId()
      };

    } catch (error) {
      Logger.log('❌ 評価レポート生成エラー: ' + error.message);
      throw error;
    }
  }, 3); // 最大3回リトライ
}

/**
 * エグゼクティブサマリー追加
 */
function addEvaluationSummary(body, data) {
  // ヘッダー
  const header = body.appendParagraph('【面接評価レポート】');
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const subtitle = body.appendParagraph(
    `${data.candidate_name} 様 - ${data.selection_phase} - ${data.interview_date}`
  );
  subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendHorizontalRule();

  // エグゼクティブサマリー見出し
  const summaryHeading = body.appendParagraph('📊 エグゼクティブサマリー');
  summaryHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  // 総合評価
  body.appendParagraph(`総合評価: ${data.total_rank}`)
    .setBold(true);
  body.appendParagraph(`採用推奨: ${data.recommendation || '要検討'}`)
    .setBold(true);

  body.appendParagraph(''); // 空行

  // 判断理由
  body.appendParagraph('【判断理由】').setBold(true);
  const reasons = data.summary_reasons || [];
  reasons.forEach(reason => {
    body.appendParagraph(`- ${reason}`);
  });

  body.appendParagraph(''); // 空行
  body.appendHorizontalRule();

  // 4軸評価
  const axesHeading = body.appendParagraph('🎯 4軸評価');
  axesHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  const axes = [
    { name: 'Philosophy（理念共感）', rank: data.philosophy_rank, score: data.philosophy_score, reason: data.philosophy_summary },
    { name: 'Strategy（戦略理解）', rank: data.strategy_rank, score: data.strategy_score, reason: data.strategy_summary },
    { name: 'Motivation（動機）', rank: data.motivation_rank, score: data.motivation_score, reason: data.motivation_summary },
    { name: 'Execution（実行力）', rank: data.execution_rank, score: data.execution_score, reason: data.execution_summary }
  ];

  axes.forEach(axis => {
    body.appendParagraph(`${axis.name}: ${axis.rank}（${axis.score}点）`)
      .setBold(true);
    body.appendParagraph(`→ ${axis.reason || '評価理由なし'}`);
    body.appendParagraph(''); // 空行
  });

  body.appendHorizontalRule();

  // 主要懸念
  const concernsHeading = body.appendParagraph('⚠️ 主要懸念（Critical/Highのみ）');
  concernsHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  const concerns = data.critical_concerns || [];
  if (concerns.length === 0) {
    body.appendParagraph('なし');
  } else {
    concerns.forEach(concern => {
      body.appendParagraph(`- ${concern}`);
    });
  }

  body.appendParagraph(''); // 空行
  body.appendHorizontalRule();

  // 次回確認ポイント
  const nextHeading = body.appendParagraph('🔍 次回確認ポイント');
  nextHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  const nextQuestions = data.next_questions || [];
  nextQuestions.forEach((q, index) => {
    if (q) {
      body.appendParagraph(`${index + 1}. ${q}`);
    }
  });

  body.appendParagraph(''); // 空行
  body.appendHorizontalRule();

  // 面接官コメント
  const commentHeading = body.appendParagraph('📝 面接官コメント');
  commentHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行
  body.appendParagraph(data.interviewer_comment || '（コメントなし）');

  body.appendParagraph(''); // 空行
  body.appendHorizontalRule();

  // 基本情報
  const infoHeading = body.appendParagraph('📄 基本情報');
  infoHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行
  body.appendParagraph(`作成日時: ${Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm')}`);
  body.appendParagraph(`面接官: ${data.interviewer}`);
  body.appendParagraph(`スプレッドシート: ${data.spreadsheet_url || '（リンクなし）'}`);
}

/**
 * 詳細評価追加
 */
function addDetailedEvaluation(body, data) {
  const heading = body.appendParagraph('Part 2: 詳細評価');
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendHorizontalRule();
  body.appendParagraph(''); // 空行

  const axes = [
    { name: 'Philosophy（理念共感）', rank: data.philosophy_rank, score: data.philosophy_score, reason: data.philosophy_reason, evidence: data.philosophy_evidence },
    { name: 'Strategy（戦略理解）', rank: data.strategy_rank, score: data.strategy_score, reason: data.strategy_reason, evidence: data.strategy_evidence },
    { name: 'Motivation（動機）', rank: data.motivation_rank, score: data.motivation_score, reason: data.motivation_reason, evidence: data.motivation_evidence },
    { name: 'Execution（実行力）', rank: data.execution_rank, score: data.execution_score, reason: data.execution_reason, evidence: data.execution_evidence }
  ];

  axes.forEach((axis, index) => {
    const axisHeading = body.appendParagraph(`【${axis.name}詳細】`);
    axisHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    body.appendParagraph(''); // 空行

    body.appendParagraph(`ランク: ${axis.rank}`).setBold(true);
    body.appendParagraph(`スコア: ${axis.score}/30点`).setBold(true);

    body.appendParagraph(''); // 空行

    body.appendParagraph('評価理由:').setBold(true);
    body.appendParagraph(axis.reason || '（評価理由なし）');

    body.appendParagraph(''); // 空行

    if (axis.evidence) {
      body.appendParagraph('主要エビデンス:').setBold(true);
      body.appendParagraph(`「${axis.evidence}」`).setItalic(true);
    }

    body.appendParagraph(''); // 空行

    if (index < axes.length - 1) {
      body.appendHorizontalRule();
      body.appendParagraph(''); // 空行
    }
  });
}

/**
 * 議事録追加
 */
function addTranscript(body, data) {
  const heading = body.appendParagraph('Part 3: 面接議事録');
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendHorizontalRule();
  body.appendParagraph(''); // 空行

  body.appendParagraph('【面接開始】').setBold(true);
  body.appendParagraph(''); // 空行

  // 文字起こし全文
  const transcript = data.transcript || '（文字起こしデータなし）';
  body.appendParagraph(transcript);

  body.appendParagraph(''); // 空行
  body.appendParagraph('【面接終了】').setBold(true);
}

/**
 * 戦略レポート生成V2
 *
 * @param {Object} data - 戦略データ
 * @param {string} recruitType - 採用区分
 * @param {string} phase - 選考フェーズ
 * @param {string} companyName - 企業名（デフォルト: "アマネク"）
 * @return {Object} 生成結果
 */
function generateStrategyReportV2(data, recruitType, phase, companyName = "アマネク") {
  Logger.log('=== 戦略レポートV2 生成開始 ===');

  return retryOperation(() => {
    try {
      // ドキュメント作成
      const dateStr = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd');
      const docTitle = `${dateStr}_${phase}_戦略レポート`;
      const doc = DocumentApp.create(docTitle);
      const body = doc.getBody();

      // フォント設定
      const style = {};
      style[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
      style[DocumentApp.Attribute.FONT_SIZE] = 11;
      body.setAttributes(style);

      // Part 1: エグゼクティブサマリー
      addStrategySummary(body, data);

      // 改ページ
      body.appendPageBreak();

      // Part 2: 詳細分析
      addDetailedStrategy(body, data);

      // ドキュメント保存
      doc.saveAndClose();

      // 候補者フォルダに移動
      const candidateFolder = getOrCreateCandidateFolder(
        recruitType,
        phase,
        data.candidate_id,
        data.candidate_name,
        companyName
      );

      const movedDoc = saveDocumentToFolder(doc, candidateFolder);

      // 共有設定（エラー時はスキップ）
      try {
        movedDoc.setSharing(
          DriveApp.Access.DOMAIN_WITH_LINK,
          DriveApp.Permission.VIEW
        );
      } catch (error) {
        Logger.log('⚠️ 共有権限設定をスキップ: ' + error.message);
      }

      const url = movedDoc.getUrl();
      Logger.log('✅ 戦略レポート生成成功: ' + url);

      return {
        success: true,
        url: url,
        documentId: movedDoc.getId()
      };

    } catch (error) {
      Logger.log('❌ 戦略レポート生成エラー: ' + error.message);
      throw error;
    }
  }, 3); // 最大3回リトライ
}

/**
 * 戦略サマリー追加
 */
function addStrategySummary(body, data) {
  // ヘッダー
  const header = body.appendParagraph('【エンゲージメント戦略レポート】');
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const subtitle = body.appendParagraph(
    `${data.candidate_name} 様 - ${data.current_phase}`
  );
  subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendHorizontalRule();

  // 承諾可能性分析
  const analysisHeading = body.appendParagraph('🎯 承諾可能性分析');
  analysisHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  body.appendParagraph(`現在の承諾可能性: ${data.acceptance_probability}%`)
    .setBold(true);
  body.appendParagraph(`信頼度: ${data.confidence_level}`)
    .setBold(true);

  body.appendParagraph(''); // 空行

  // 競合状況
  body.appendParagraph('競合状況:').setBold(true);
  if (data.competitor_probabilities && data.competitor_probabilities.length > 0) {
    data.competitor_probabilities.forEach(comp => {
      body.appendParagraph(`- ${comp.company}: ${comp.probability}%`);
    });
  } else {
    body.appendParagraph('- 競合情報なし');
  }

  body.appendParagraph(''); // 空行
  body.appendHorizontalRule();

  // 24時間アクション
  const actionHeading = body.appendParagraph('🚨 24時間以内の最優先アクション');
  actionHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  body.appendParagraph('【アクション】').setBold(true);
  body.appendParagraph(data.immediate_action_24h || '（アクションなし）');

  body.appendParagraph(''); // 空行

  body.appendParagraph('【理由】').setBold(true);
  body.appendParagraph(data.action_reason || '（理由なし）');

  body.appendParagraph(''); // 空行

  body.appendParagraph('【期待効果】').setBold(true);
  body.appendParagraph(data.expected_effect || '（効果不明）');

  body.appendParagraph(''); // 空行

  body.appendParagraph(`【担当】${data.interviewer}  【期限】（要設定）`);

  body.appendParagraph(''); // 空行
  body.appendHorizontalRule();

  // 主要リスク
  const riskHeading = body.appendParagraph('⚠️ 主要リスクと対策');
  riskHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  const risks = data.risk_factors || [];
  risks.forEach((risk, index) => {
    body.appendParagraph(`${index + 1}. リスク: ${risk.factor}`);
    body.appendParagraph(`   対策: ${risk.countermeasure}`);
    body.appendParagraph(''); // 空行
  });

  if (risks.length === 0) {
    body.appendParagraph('主要なリスクは検出されませんでした');
  }

  body.appendHorizontalRule();

  // 訴求すべき強み
  const strengthHeading = body.appendParagraph('✅ 訴求すべき強み');
  strengthHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  const strengths = data.our_strengths || [];
  strengths.forEach((strength, index) => {
    body.appendParagraph(`${index + 1}. ${strength}`);
  });

  body.appendParagraph(''); // 空行
  body.appendHorizontalRule();

  // 承諾ストーリー
  const storyHeading = body.appendParagraph('📈 承諾までのストーリー');
  storyHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  const story = data.acceptance_story || [];
  story.forEach(step => {
    body.appendParagraph(step);
  });

  body.appendParagraph(''); // 空行
  body.appendHorizontalRule();

  // 基本情報
  const infoHeading = body.appendParagraph('📄 基本情報');
  infoHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行
  body.appendParagraph(`作成日時: ${Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm')}`);
  body.appendParagraph(`面接官: ${data.interviewer}`);
  body.appendParagraph(`スプレッドシート: ${data.spreadsheet_url || '（リンクなし）'}`);
}

/**
 * 詳細戦略追加
 */
function addDetailedStrategy(body, data) {
  const heading = body.appendParagraph('Part 2: 詳細分析');
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendHorizontalRule();
  body.appendParagraph(''); // 空行

  // ポジティブ要因
  const positiveHeading = body.appendParagraph('【ポジティブ要因詳細】');
  positiveHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  const positiveFactors = data.positive_factors || [];
  positiveFactors.forEach((factor, index) => {
    body.appendParagraph(`${index + 1}. ${factor.factor}`);
    if (factor.evidence) {
      body.appendParagraph(`   エビデンス: 「${factor.evidence}」`).setItalic(true);
    }
    body.appendParagraph(''); // 空行
  });

  body.appendHorizontalRule();
  body.appendParagraph(''); // 空行

  // リスク要因詳細
  const riskDetailHeading = body.appendParagraph('【リスク要因詳細】');
  riskDetailHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  const riskDetails = data.risk_factors_detailed || [];
  riskDetails.forEach((risk, index) => {
    body.appendParagraph(`${index + 1}. ${risk.factor}`);
    body.appendParagraph(`   深刻度: ${risk.severity}`);
    body.appendParagraph(`   対策: ${risk.detailed_countermeasure}`);
    body.appendParagraph(''); // 空行
  });

  body.appendHorizontalRule();
  body.appendParagraph(''); // 空行

  // 競合分析
  const competitorHeading = body.appendParagraph('【競合分析】');
  competitorHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  const competitors = data.competitor_analysis || [];

  if (competitors.length === 0) {
    body.appendParagraph('競合分析データなし');
  } else {
    competitors.forEach((comp, index) => {
      // 企業名（見出し）
      body.appendParagraph(`## ${comp.company}`).setBold(true);
      body.appendParagraph('');

      // 給与・待遇
      if (comp.compensation) {
        body.appendParagraph('### 給与・待遇').setBold(true);

        if (comp.compensation.average_salary) {
          body.appendParagraph(`- 平均年収: ${comp.compensation.average_salary}`);
        }
        if (comp.compensation.starting_salary && comp.compensation.starting_salary !== '情報なし') {
          body.appendParagraph(`- 初任給: ${comp.compensation.starting_salary}`);
        }
        if (comp.compensation.details) {
          body.appendParagraph(`- 詳細: ${comp.compensation.details}`);
        }
        body.appendParagraph('');
      }

      // 企業文化・働き方
      if (comp.culture && (comp.culture.work_hours !== '情報なし' ||
          comp.culture.work_life_balance !== '情報なし' ||
          comp.culture.remote_work !== '情報なし')) {
        body.appendParagraph('### 企業文化・働き方').setBold(true);

        if (comp.culture.work_hours && comp.culture.work_hours !== '情報なし') {
          body.appendParagraph(`- 平均残業時間: ${comp.culture.work_hours}`);
        }
        if (comp.culture.work_life_balance && comp.culture.work_life_balance !== '情報なし') {
          body.appendParagraph(`- ワークライフバランス: ${comp.culture.work_life_balance}`);
        }
        if (comp.culture.remote_work && comp.culture.remote_work !== '情報なし') {
          body.appendParagraph(`- リモートワーク: ${comp.culture.remote_work}`);
        }
        body.appendParagraph('');
      }

      // 評判・口コミ
      if (comp.reputation && (comp.reputation.rating !== '情報なし' ||
          comp.reputation.positive !== '情報なし' ||
          comp.reputation.negative !== '情報なし')) {
        body.appendParagraph('### 評判・口コミ').setBold(true);

        if (comp.reputation.rating && comp.reputation.rating !== '情報なし') {
          body.appendParagraph(`- 総合評価: ${comp.reputation.rating}`);
        }
        if (comp.reputation.positive && comp.reputation.positive !== '情報なし') {
          body.appendParagraph(`- ポジティブ: ${comp.reputation.positive}`);
        }
        if (comp.reputation.negative && comp.reputation.negative !== '情報なし') {
          body.appendParagraph(`- ネガティブ: ${comp.reputation.negative}`);
        }
        body.appendParagraph('');
      }

      // 競合の強み
      if (comp.their_strengths) {
        body.appendParagraph('### 競合の強み').setBold(true);
        body.appendParagraph(comp.their_strengths);
        body.appendParagraph('');
      }

      // 競合の弱み
      if (comp.their_weaknesses) {
        body.appendParagraph('### 競合の弱み').setBold(true);
        body.appendParagraph(comp.their_weaknesses);
        body.appendParagraph('');
      }

      // 自社の対抗策
      if (comp.our_counter_strategy) {
        body.appendParagraph('### 自社の対抗策').setBold(true);
        body.appendParagraph(comp.our_counter_strategy);
        body.appendParagraph('');
      }

      // 承諾確率推定
      if (comp.estimated_probability !== undefined && comp.estimated_probability !== null) {
        body.appendParagraph('### 承諾確率推定').setBold(true);
        body.appendParagraph(`${comp.estimated_probability}%`);
        body.appendParagraph('');
      }

      // 区切り線（最後の企業以外）
      if (index < competitors.length - 1) {
        body.appendHorizontalRule();
        body.appendParagraph('');
      }
    });
  }

  body.appendHorizontalRule();
  body.appendParagraph(''); // 空行

  // 推奨施策
  const recommendationHeading = body.appendParagraph('【推奨エンゲージメント施策】');
  recommendationHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(''); // 空行

  const recommendations = data.engagement_recommendations || [];
  recommendations.forEach((rec, index) => {
    body.appendParagraph(`${index + 1}. ${rec}`);
  });
}

/**
 * ドキュメントをフォルダに移動
 *
 * @param {GoogleAppsScript.Document.Document} doc - Googleドキュメント
 * @param {GoogleAppsScript.Drive.Folder} folder - 移動先フォルダ
 * @return {GoogleAppsScript.Drive.File} 移動後のファイル
 */
function saveDocumentToFolder(doc, folder) {
  const docFile = DriveApp.getFileById(doc.getId());

  // 現在の親フォルダを削除
  const parents = docFile.getParents();
  while (parents.hasNext()) {
    const parent = parents.next();
    parent.removeFile(docFile);
  }

  // 新しいフォルダに追加
  folder.addFile(docFile);

  Logger.log(`✅ ドキュメント移動: ${folder.getName()}`);

  return docFile;
}

/**
 * リトライロジック
 *
 * @param {Function} operation - 実行する関数
 * @param {number} maxRetries - 最大リトライ回数（デフォルト: 3）
 * @return {*} operation()の戻り値
 */
function retryOperation(operation, maxRetries = 3) {
  let lastError;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      Logger.log(`試行 ${attempt}/${maxRetries}`);
      return operation();
    } catch (error) {
      lastError = error;
      Logger.log(`❌ 試行 ${attempt} 失敗: ${error.message}`);

      if (attempt < maxRetries) {
        const waitTime = attempt * 2; // 2秒, 4秒
        Logger.log(`${waitTime}秒待機後に再試行...`);
        Utilities.sleep(waitTime * 1000);
      }
    }
  }

  Logger.log(`❌ ${maxRetries}回試行後も失敗`);
  throw lastError;
}

/**
 * テスト関数
 */
function testReportGeneratorV2() {
  Logger.log('=== ReportGeneratorV2 テスト開始 ===');

  const companyName = 'テスト株式会社';  // テスト用企業名

  // テストデータ
  const evalData = {
    candidate_id: 'CAND_20251218_150000',
    candidate_name: 'テスト次郎',
    selection_phase: '1次面接',
    interview_date: '2025/12/18',
    interviewer: '東',
    total_rank: 'A',
    recommendation: '積極採用推奨',
    summary_reasons: [
      '理念への深い共感と高い志望度',
      '優れた戦略的思考力と実行力',
      '成長意欲が極めて高い'
    ],
    philosophy_rank: 'A',
    philosophy_score: 29,
    philosophy_summary: '理念への深い共感が見られる',
    philosophy_reason: '当社の理念に深く共感しており、自身のキャリアビジョンとも合致している',
    philosophy_evidence: '私も価値提供に本気な会社で働きたいと考えています',
    strategy_rank: 'B',
    strategy_score: 25,
    strategy_summary: '戦略理解は十分、実践経験で向上可',
    strategy_reason: '戦略的思考力は高いが、実践経験でさらに向上する余地がある',
    strategy_evidence: 'データに基づいた意思決定の重要性を理解しています',
    motivation_rank: 'A',
    motivation_score: 19,
    motivation_summary: '非常に高い志望度、成長意欲強',
    motivation_reason: '非常に高い志望度と成長意欲が見られる',
    motivation_evidence: '御社で成長し続けたいと強く思っています',
    execution_rank: 'A',
    execution_score: 20,
    execution_summary: '優れた実行力、実績あり',
    execution_reason: '優れた実行力があり、過去の実績も十分',
    execution_evidence: '前職でプロジェクトを成功に導いた経験があります',
    critical_concerns: [],
    next_questions: [
      '具体的な成果事例の深掘り',
      'チームマネジメント経験',
      'キャリアビジョンの明確化'
    ],
    interviewer_comment: '非常に優秀な候補者。理念への共感が深く、戦略的思考力も高い。積極的に次のステップへ。',
    transcript: '面接官: 本日はよろしくお願いします。\n候補者: よろしくお願いします。\n\n面接官: まず、弊社に応募いただいた理由を教えていただけますか？\n候補者: はい。御社の「価値提供に本気」という理念に深く共感しました。私も価値提供に本気な会社で働きたいと考えています。\n\n（以下、文字起こし全文）',
    spreadsheet_url: 'https://docs.google.com/spreadsheets/d/xxx'
  };

  // 評価レポート生成
  Logger.log('--- 評価レポート生成 ---');
  const evalResult = generateEvaluationReportV2(evalData, '新卒', '1次面接', companyName);
  Logger.log('結果: ' + JSON.stringify(evalResult, null, 2));

  // 戦略レポート生成
  Logger.log('--- 戦略レポート生成 ---');
  const strategyData = {
    candidate_id: 'CAND_20251218_150000',
    candidate_name: 'テスト次郎',
    current_phase: '1次面接',
    interviewer: '東',
    acceptance_probability: 75,
    confidence_level: 'HIGH',
    competitor_probabilities: [
      { company: '自社', probability: 53 },
      { company: 'リブコンサルティング', probability: 30 },
      { company: 'ベイカレント', probability: 17 }
    ],
    immediate_action_24h: '給与条件を明確に提示',
    action_reason: '待遇面への懸念が承諾率を下げている',
    expected_effect: '承諾可能性 75% → 85% (+10%)',
    risk_factors: [
      { factor: '給与への懸念', countermeasure: '次回面接で具体的な条件提示' },
      { factor: '競合他社の存在', countermeasure: '自社の差別化ポイント訴求' }
    ],
    our_strengths: [
      '成長機会の豊富さ',
      '理念への共感を活かせる環境',
      '早期裁量・挑戦機会'
    ],
    acceptance_story: [
      'Step 1: 給与明示 → +10% → 85%',
      'Step 2: 社員面談 → +8% → 93%',
      'Step 3: クロージング → +7% → 100%'
    ],
    positive_factors: [
      { factor: '理念への強い共感', evidence: '価値提供に本気な会社で働きたい' },
      { factor: '高いスキルセット', evidence: 'データ分析の経験が豊富' }
    ],
    risk_factors_detailed: [
      { factor: '給与への懸念', severity: 'HIGH', detailed_countermeasure: '次回面接で具体的な給与レンジと昇給制度を説明' }
    ],
    competitor_analysis: [
      { company: 'リブコンサルティング', strengths: '給与水準が高い', weaknesses: '残業が多い', counterstrategy: 'ワークライフバランスの良さを訴求' }
    ],
    engagement_recommendations: [
      '24時間以内に給与条件を提示',
      '48時間以内に社員面談を設定',
      '1週間以内にクロージング面談実施'
    ],
    spreadsheet_url: 'https://docs.google.com/spreadsheets/d/xxx'
  };

  const strategyResult = generateStrategyReportV2(strategyData, '新卒', '1次面接', companyName);
  Logger.log('結果: ' + JSON.stringify(strategyResult, null, 2));

  Logger.log('=== ReportGeneratorV2 テスト完了 ===');
}
