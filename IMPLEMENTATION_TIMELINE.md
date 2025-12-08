# Phase 4-2 実装タイムライン（戦略D+）

**作成日**: 2025-12-08
**所要時間**: 2-2.5日
**実装範囲**: 4種類完全実装 + 3種類親切なエラー

---

## 📋 目次

1. [全体スケジュール](#1-全体スケジュール)
2. [Day 1: コア機能実装](#2-day-1-コア機能実装)
3. [Day 2: 承諾戦略機能実装](#3-day-2-承諾戦略機能実装)
4. [Day 3（オプション）: テスト・ドキュメント](#4-day-3オプション-テストドキュメント)
5. [チェックポイント](#5-チェックポイント)
6. [タスク依存関係](#6-タスク依存関係)

---

## 1. 全体スケジュール

### 1.1 概要

```
Day 1: コア機能実装（6-8時間）
  Morning:   handleEvaluationData() 完全実装
  Afternoon: handleEngagementData() 完全実装

Day 2: 承諾戦略機能実装（6-8時間）
  Morning:   handleAcceptanceStoryData() + handleCompetitorComparisonData()
  Afternoon: 親切なエラーハンドラ + スプレッドシート準備 + デプロイ

Day 3（オプション）: テスト・ドキュメント（2-4時間）
  Morning: Dify連携テスト + エラーケーステスト
  ----

合計所要時間: 2-2.5日（Day 3はオプション）
```

### 1.2 マイルストーン

| マイルストーン | 完了時刻 | 成果物 |
|--------------|---------|--------|
| M1: evaluation実装完了 | Day 1 12:00 | handleEvaluationData() + ヘルパー関数 |
| M2: engagement実装完了 | Day 1 18:00 | handleEngagementData() + ヘルパー関数 |
| M3: acceptance_story実装完了 | Day 2 12:00 | handleAcceptanceStoryData() |
| M4: competitor_comparison実装完了 | Day 2 14:00 | handleCompetitorComparisonData() |
| M5: 親切なエラー実装完了 | Day 2 15:00 | 3種類のエラーハンドラ |
| M6: デプロイ完了 | Day 2 18:00 | Web App デプロイ + URL取得 |
| M7（オプション）: テスト完了 | Day 3 12:00 | 動作確認完了 |

---

## 2. Day 1: コア機能実装

### 2.1 Morning Session（09:00-12:00, 3時間）

#### タスク1.1: スプレッドシート構造の拡張（30分）
**所要時間**: 30分
**優先度**: 🔴 高

**作業内容**:
```
□ SpreadsheetRedesign.gsを修正
  - createCandidateInsightsSheet()を更新
  - 列を11列 → 16列に拡張
  - L列: 推奨承諾ストーリー
  - M列: キーメッセージ
  - N列: 差別化ポイント
  - O列: 競合分析結果
  - P列: 推奨戦略

□ phase0_preparation()を実行
  - バックアップ作成
  - 列の存在確認

□ phase1_execute()を実行
  - Candidate_Insightsシート作成（16列）
  - Candidate_Scoresシート作成
```

**成果物**:
- 修正されたSpreadsheetRedesign.gs
- 拡張されたCandidate_Insightsシート

**チェックポイント**:
- [ ] Candidate_Insightsシートが16列になっているか
- [ ] 列ヘッダーが正しく設定されているか
- [ ] バックアップシートが作成されているか

---

#### タスク1.2: ヘルパー関数の実装（60分）
**所要時間**: 60分
**優先度**: 🔴 高

**作業内容**:
```
DifyIntegration.gsに以下の関数を追加:

□ findCandidateRow(sheet, candidateId)
  - シート内でcandidate_idを検索
  - 該当行のインデックスを返す
  - 見つからない場合はnullを返す

□ addNewCandidateRow(sheet, candidateId)
  - 新規候補者の行を追加
  - 基本情報（candidate_id, 氏名）を設定

□ getCandidateName(candidateId)
  - Candidates_Masterから候補者の氏名を取得

□ calculateIntegratedAcceptance(aiPrediction, humanIntuition)
  - 統合承諾可能性を計算（AI 70% + 人間 30%）

□ formatComparisonResult(competitors)
  - 競合分析結果を整形

□ convertArrayToNewlineSeparated(array)
  - 配列を改行区切りの文字列に変換

□ convertISO8601ToJST(iso8601String)
  - ISO 8601 → 日本時間に変換
```

**実装例**:
```javascript
/**
 * シート内でcandidate_idを検索
 */
function findCandidateRow(sheet, candidateId) {
  const data = sheet.getDataRange().getValues();

  // ヘッダー行をスキップ（i=1から開始）
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === candidateId) { // A列: candidate_id
      return i + 1; // 行番号（1-indexed）
    }
  }

  return null; // 見つからない
}

/**
 * 新規候補者の行を追加
 */
function addNewCandidateRow(sheet, candidateId) {
  const candidateName = getCandidateName(candidateId);
  const timestamp = new Date().toISOString();

  // 基本情報のみを設定（他の列は空）
  const newRow = [candidateId, candidateName, timestamp];
  const columnsNeeded = sheet.getLastColumn();

  // 残りの列を空文字列で埋める
  for (let i = newRow.length; i < columnsNeeded; i++) {
    newRow.push('');
  }

  sheet.appendRow(newRow);

  Logger.log(`✅ 新規候補者の行を追加: ${candidateId}`);
}

// 他のヘルパー関数も同様に実装...
```

**成果物**:
- ヘルパー関数（6-7個）
- 単体テスト用のサンプルコード

**チェックポイント**:
- [ ] 各関数が正しく動作するか
- [ ] エラーハンドリングが適切か
- [ ] ログ出力が分かりやすいか

---

#### タスク1.3: handleEvaluationData()の実装（90分）
**所要時間**: 90分
**優先度**: 🔴 高

**作業内容**:
```
□ validateEvaluationData()の実装
  - 必須フィールドチェック
  - スコア範囲チェック（0-100）
  - 承諾可能性範囲チェック（0-1）

□ handleEvaluationData()の実装
  - データバリデーション
  - Candidate_Scoresシート更新
    * D列: 最新_合格可能性
    * G-J列: Philosophy, Strategy, Motivation, Execution
    * K列: 最新_合計スコア
    * L列: 最新_承諾可能性（AI予測）
  - Candidate_Insightsシート更新
    * D列: コアモチベーション
    * E列: 主要懸念事項
    * F列: 懸念カテゴリ
  - エラーハンドリング
  - ログ出力

□ 単体テスト実行
  - 正常ケース
  - エラーケース（必須フィールド欠損）
  - エラーケース（範囲外の値）
```

**実装骨組み**:
```javascript
function handleEvaluationData(data) {
  try {
    Logger.log('📊 Processing evaluation data for: ' + data.candidate_id);

    // 1. データバリデーション
    validateEvaluationData(data);

    // 2. スプレッドシート取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scoresSheet = ss.getSheetByName('Candidate_Scores');
    const insightsSheet = ss.getSheetByName('Candidate_Insights');

    if (!scoresSheet || !insightsSheet) {
      throw new Error('Required sheets not found');
    }

    // 3. 候補者の行を検索（なければ追加）
    let scoresRow = findCandidateRow(scoresSheet, data.candidate_id);
    if (!scoresRow) {
      addNewCandidateRow(scoresSheet, data.candidate_id);
      scoresRow = scoresSheet.getLastRow();
    }

    let insightsRow = findCandidateRow(insightsSheet, data.candidate_id);
    if (!insightsRow) {
      addNewCandidateRow(insightsSheet, data.candidate_id);
      insightsRow = insightsSheet.getLastRow();
    }

    // 4. Candidate_Scoresを更新
    const timestamp = convertISO8601ToJST(data.timestamp);
    scoresSheet.getRange(scoresRow, 3).setValue(timestamp);                    // C: 最終更新日時
    scoresSheet.getRange(scoresRow, 4).setValue(data.scores.overall);          // D: 最新_合格可能性
    scoresSheet.getRange(scoresRow, 7).setValue(data.scores.philosophy);       // G: Philosophy
    scoresSheet.getRange(scoresRow, 8).setValue(data.scores.strategy);         // H: Strategy
    scoresSheet.getRange(scoresRow, 9).setValue(data.scores.motivation);       // I: Motivation
    scoresSheet.getRange(scoresRow, 10).setValue(data.scores.execution);       // J: Execution
    scoresSheet.getRange(scoresRow, 11).setValue(data.scores.total);           // K: 合計スコア
    scoresSheet.getRange(scoresRow, 12).setValue(convertProbabilityToPercentage(data.scores.acceptance_probability)); // L: 承諾可能性

    // 5. Candidate_Insightsを更新
    insightsSheet.getRange(insightsRow, 3).setValue(timestamp);                      // C: 最終更新日時
    insightsSheet.getRange(insightsRow, 4).setValue(data.insights.core_motivation); // D: コアモチベーション
    insightsSheet.getRange(insightsRow, 5).setValue(data.insights.main_concern);    // E: 主要懸念事項
    insightsSheet.getRange(insightsRow, 6).setValue(data.insights.concern_category);// F: 懸念カテゴリ

    Logger.log('✅ Evaluation data processed successfully for: ' + data.candidate_id);

  } catch (error) {
    Logger.log('❌ Error in handleEvaluationData: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}
```

**成果物**:
- handleEvaluationData() 完全実装
- validateEvaluationData() 完全実装
- テスト結果ログ

**チェックポイント**:
- [ ] データがCandidate_Scoresに正しく書き込まれるか
- [ ] データがCandidate_Insightsに正しく書き込まれるか
- [ ] エラーケースで適切なエラーが発生するか
- [ ] ログが分かりやすいか

---

### 2.2 Afternoon Session（13:00-18:00, 5時間）

#### タスク1.4: handleEngagementData()の実装（120分）
**所要時間**: 120分
**優先度**: 🔴 高

**作業内容**:
```
□ validateEngagementData()の実装
  - 必須フィールドチェック
  - スコア範囲チェック（0-100）
  - 信頼度範囲チェック（0-1）

□ handleEngagementData()の実装
  - データバリデーション
  - engagement_id生成
  - 候補者名取得
  - 統合承諾可能性計算
  - Engagement_Logシートに新規行追加（appendRow）
  - エラーハンドリング
  - ログ出力

□ 単体テスト実行
  - 正常ケース
  - エラーケース（必須フィールド欠損）
  - エラーケース（範囲外の値）
  - human_intuitionがnullのケース
```

**実装骨組み**:
```javascript
function handleEngagementData(data) {
  try {
    Logger.log('🤝 Processing engagement data for: ' + data.candidate_id);

    // 1. データバリデーション
    validateEngagementData(data);

    // 2. スプレッドシート取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const engagementSheet = ss.getSheetByName('Engagement_Log');

    if (!engagementSheet) {
      throw new Error('Engagement_Log sheet not found');
    }

    // 3. 候補者の氏名を取得
    const candidateName = getCandidateName(data.candidate_id);

    // 4. engagement_idを生成
    const engagementId = 'ENG_' + new Date().getTime();

    // 5. 統合_承諾可能性を計算
    const integrated = calculateIntegratedAcceptance(
      data.engagement.ai_prediction,
      data.engagement.human_intuition
    );

    // 6. 信頼度を0-1 → 0-100に変換
    const confidencePercentage = convertProbabilityToPercentage(data.engagement.confidence);

    // 7. Engagement_Logに新規行を追加
    engagementSheet.appendRow([
      engagementId,                             // A: engagement_id
      data.candidate_id,                        // B: candidate_id
      candidateName,                            // C: 氏名
      data.engagement.date,                     // D: 接触日時
      data.engagement.phase,                    // E: フェーズ
      data.engagement.ai_prediction,            // F: AI予測
      data.engagement.human_intuition || '',    // G: 人間の直感
      integrated,                               // H: 統合
      confidencePercentage,                     // I: 信頼度
      data.engagement.motivation_score,         // J: 志望度
      data.engagement.competitive_score,        // K: 競合優位性
      data.engagement.concern_score,            // L: 懸念解消度
      data.engagement.core_motivation,          // M: コアモチベーション
      data.engagement.top_concern               // N: 主要懸念
    ]);

    Logger.log('✅ Engagement data processed successfully for: ' + data.candidate_id);

  } catch (error) {
    Logger.log('❌ Error in handleEngagementData: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}
```

**成果物**:
- handleEngagementData() 完全実装
- validateEngagementData() 完全実装
- テスト結果ログ

**チェックポイント**:
- [ ] データがEngagement_Logに正しく追記されるか
- [ ] engagement_idが正しく生成されるか
- [ ] 統合承諾可能性が正しく計算されるか
- [ ] human_intuitionがnullの場合に正しく処理されるか

---

#### タスク1.5: Day 1レビューとテスト（60分）
**所要時間**: 60分
**優先度**: 🟡 中

**作業内容**:
```
□ handleEvaluationData()のテスト
  - 正常ケース: 完全なデータ
  - エラーケース: 必須フィールド欠損
  - エラーケース: スコア範囲外
  - エラーケース: シートが見つからない

□ handleEngagementData()のテスト
  - 正常ケース: 完全なデータ
  - エラーケース: 必須フィールド欠損
  - エラーケース: 日付形式不正
  - 特殊ケース: human_intuitionがnull

□ スプレッドシートの確認
  - Candidate_Scoresにデータが正しく書き込まれているか
  - Candidate_Insightsにデータが正しく書き込まれているか
  - Engagement_Logにデータが正しく追記されているか

□ コードレビュー
  - エラーハンドリングは適切か
  - ログ出力は分かりやすいか
  - コメントは十分か
```

**成果物**:
- テスト結果レポート
- バグ修正リスト（あれば）

**チェックポイント**:
- [ ] すべてのテストケースがパスするか
- [ ] エラーメッセージが分かりやすいか
- [ ] ログ出力が適切か

---

### 2.3 Day 1 終了時の状態

**完了したもの**:
- ✅ Candidate_Insightsシートの拡張（16列）
- ✅ ヘルパー関数の実装（6-7個）
- ✅ handleEvaluationData() 完全実装
- ✅ handleEngagementData() 完全実装
- ✅ テスト実施とバグ修正

**未完了のもの**:
- ⏳ handleAcceptanceStoryData()
- ⏳ handleCompetitorComparisonData()
- ⏳ 親切なエラーハンドラ（3種類）
- ⏳ Web Appデプロイ

---

## 3. Day 2: 承諾戦略機能実装

### 3.1 Morning Session（09:00-12:00, 3時間）

#### タスク2.1: handleAcceptanceStoryData()の実装（90分）
**所要時間**: 90分
**優先度**: 🔴 高

**作業内容**:
```
□ validateAcceptanceStoryData()の実装
  - 必須フィールドチェック
  - 配列型チェック（key_messages, recommended_actions）

□ handleAcceptanceStoryData()の実装
  - データバリデーション
  - 配列データを改行区切りに変換
  - Candidate_Insightsシート更新（拡張列）
    * L列: 推奨承諾ストーリー
    * M列: キーメッセージ
    * J列: 次推奨アクション
    * K列: アクション期限
  - エラーハンドリング
  - ログ出力

□ 単体テスト実行
  - 正常ケース
  - エラーケース（必須フィールド欠損）
  - 配列データの変換テスト
```

**実装骨組み**:
```javascript
function handleAcceptanceStoryData(data) {
  try {
    Logger.log('📖 Processing acceptance story data for: ' + data.candidate_id);

    // 1. データバリデーション
    validateAcceptanceStoryData(data);

    // 2. スプレッドシート取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const insightsSheet = ss.getSheetByName('Candidate_Insights');

    if (!insightsSheet) {
      throw new Error('Candidate_Insights sheet not found');
    }

    // 3. 候補者の行を検索（なければ追加）
    let insightsRow = findCandidateRow(insightsSheet, data.candidate_id);
    if (!insightsRow) {
      addNewCandidateRow(insightsSheet, data.candidate_id);
      insightsRow = insightsSheet.getLastRow();
    }

    // 4. 配列データを改行区切りの文字列に変換
    const keyMessages = convertArrayToNewlineSeparated(data.story.key_messages);
    const recommendedActions = convertArrayToNewlineSeparated(data.story.recommended_actions);

    // 5. Candidate_Insightsを更新
    const timestamp = convertISO8601ToJST(data.timestamp);
    insightsSheet.getRange(insightsRow, 3).setValue(timestamp);                     // C: 最終更新日時
    insightsSheet.getRange(insightsRow, 12).setValue(data.story.recommended_story); // L: 推奨承諾ストーリー
    insightsSheet.getRange(insightsRow, 13).setValue(keyMessages);                  // M: キーメッセージ
    insightsSheet.getRange(insightsRow, 10).setValue(recommendedActions);           // J: 次推奨アクション
    insightsSheet.getRange(insightsRow, 11).setValue(data.story.deadline);          // K: アクション期限

    Logger.log('✅ Acceptance story data processed successfully for: ' + data.candidate_id);

  } catch (error) {
    Logger.log('❌ Error in handleAcceptanceStoryData: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}
```

**成果物**:
- handleAcceptanceStoryData() 完全実装
- validateAcceptanceStoryData() 完全実装
- テスト結果ログ

**チェックポイント**:
- [ ] データがCandidate_Insightsに正しく書き込まれるか
- [ ] 配列データが改行区切りで正しく変換されるか
- [ ] エラーケースで適切なエラーが発生するか

---

#### タスク2.2: handleCompetitorComparisonData()の実装（90分）
**所要時間**: 90分
**優先度**: 🔴 高

**作業内容**:
```
□ validateCompetitorComparisonData()の実装
  - 必須フィールドチェック
  - 配列型チェック（competitors, differentiation_points）

□ formatComparisonResult()の実装
  - competitors配列を整形

□ handleCompetitorComparisonData()の実装
  - データバリデーション
  - 競合企業名の抽出
  - 競合分析結果の整形
  - Candidate_Insightsシート更新（拡張列）
    * G-I列: 競合企業1-3
    * N列: 差別化ポイント
    * O列: 競合分析結果
    * P列: 推奨戦略
  - エラーハンドリング
  - ログ出力

□ 単体テスト実行
  - 正常ケース
  - エラーケース（必須フィールド欠損）
  - 競合企業が1社のみのケース
  - 競合企業が3社のケース
```

**実装骨組み**:
```javascript
function handleCompetitorComparisonData(data) {
  try {
    Logger.log('🏆 Processing competitor comparison data for: ' + data.candidate_id);

    // 1. データバリデーション
    validateCompetitorComparisonData(data);

    // 2. スプレッドシート取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const insightsSheet = ss.getSheetByName('Candidate_Insights');

    if (!insightsSheet) {
      throw new Error('Candidate_Insights sheet not found');
    }

    // 3. 候補者の行を検索（なければ追加）
    let insightsRow = findCandidateRow(insightsSheet, data.candidate_id);
    if (!insightsRow) {
      addNewCandidateRow(insightsSheet, data.candidate_id);
      insightsRow = insightsSheet.getLastRow();
    }

    // 4. 競合分析結果を整形
    const comparisonResult = formatComparisonResult(data.comparison.competitors);
    const differentiationPoints = convertArrayToNewlineSeparated(data.comparison.differentiation_points);

    // 5. Candidate_Insightsを更新
    const timestamp = convertISO8601ToJST(data.timestamp);

    // 競合企業名を設定（最大3社）
    if (data.comparison.competitors[0]) {
      insightsSheet.getRange(insightsRow, 7).setValue(data.comparison.competitors[0].company); // G: 競合企業1
    }
    if (data.comparison.competitors[1]) {
      insightsSheet.getRange(insightsRow, 8).setValue(data.comparison.competitors[1].company); // H: 競合企業2
    }
    if (data.comparison.competitors[2]) {
      insightsSheet.getRange(insightsRow, 9).setValue(data.comparison.competitors[2].company); // I: 競合企業3
    }

    // 拡張列を設定
    insightsSheet.getRange(insightsRow, 3).setValue(timestamp);                        // C: 最終更新日時
    insightsSheet.getRange(insightsRow, 14).setValue(differentiationPoints);           // N: 差別化ポイント
    insightsSheet.getRange(insightsRow, 15).setValue(comparisonResult);                // O: 競合分析結果
    insightsSheet.getRange(insightsRow, 16).setValue(data.comparison.recommended_strategy); // P: 推奨戦略

    Logger.log('✅ Competitor comparison data processed successfully for: ' + data.candidate_id);

  } catch (error) {
    Logger.log('❌ Error in handleCompetitorComparisonData: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}
```

**成果物**:
- handleCompetitorComparisonData() 完全実装
- validateCompetitorComparisonData() 完全実装
- formatComparisonResult() 実装
- テスト結果ログ

**チェックポイント**:
- [ ] データがCandidate_Insightsに正しく書き込まれるか
- [ ] 競合企業名が正しく抽出されるか
- [ ] 競合分析結果が読みやすく整形されるか
- [ ] 差別化ポイントが改行区切りで表示されるか

---

### 3.2 Afternoon Session（13:00-18:00, 5時間）

#### タスク2.3: 親切なエラーハンドラの実装（60分）
**所要時間**: 60分
**優先度**: 🟡 中

**作業内容**:
```
□ handleEvidenceData()の実装
  - 「未実装」のメッセージ返却
  - 利用可能な機能の案内
  - 受信データのログ記録

□ handleRiskData()の実装
  - 「未実装」のメッセージ返却
  - 利用可能な機能の案内
  - 受信データのログ記録

□ handleNextQData()の実装
  - 「未実装」のメッセージ返却
  - 利用可能な機能の案内
  - 受信データのログ記録
```

**実装例**:
```javascript
function handleEvidenceData(data) {
  const message = {
    status: 'not_implemented',
    message: 'Evidence機能は将来のバージョンで実装予定です',
    data_type: 'evidence',
    candidate_id: data.candidate_id,
    timestamp: new Date().toISOString(),
    currently_available: [
      'evaluation - 評価データ処理',
      'engagement - エンゲージメントデータ処理',
      'acceptance_story - 承諾ストーリー処理',
      'competitor_comparison - 競合比較処理'
    ],
    future_features: [
      'evidence - エビデンス記録（将来実装予定）',
      'risk - リスク分析（将来実装予定）',
      'nextq - 次の質問生成（将来実装予定）'
    ]
  };

  Logger.log('⚠️ Evidence機能は未実装です: ' + JSON.stringify(message));

  return ContentService.createTextOutput(
    JSON.stringify(message, null, 2)
  ).setMimeType(ContentService.MimeType.JSON);
}

// handleRiskData(), handleNextQData()も同様の構造
```

**成果物**:
- handleEvidenceData() 実装
- handleRiskData() 実装
- handleNextQData() 実装

**チェックポイント**:
- [ ] エラーメッセージが親切か
- [ ] 利用可能な機能が明示されているか
- [ ] 受信データがログに記録されるか

---

#### タスク2.4: doPost関数の更新（30分）
**所要時間**: 30分
**優先度**: 🔴 高

**作業内容**:
```
□ doPost関数の親切なエラーハンドラへの接続
  - handleEvidenceData()の返り値を正しく返す
  - handleRiskData()の返り値を正しく返す
  - handleNextQData()の返り値を正しく返す

□ エラーハンドリングの強化
  - 未知のdata typeへの対応
  - JSON parseエラーへの対応
```

**修正例**:
```javascript
function doPost(e) {
  try {
    const requestBody = e.postData.contents;
    Logger.log('Received webhook: ' + requestBody);

    const data = JSON.parse(requestBody);

    // データタイプに応じて処理を分岐
    switch (data.type) {
      case 'evaluation':
        handleEvaluationData(data);
        break;
      case 'engagement':
        handleEngagementData(data);
        break;
      case 'evidence':
        return handleEvidenceData(data);  // ← 返り値をreturn
      case 'risk':
        return handleRiskData(data);      // ← 返り値をreturn
      case 'next_q':
        return handleNextQData(data);     // ← 返り値をreturn
      case 'acceptance_story':
        handleAcceptanceStoryData(data);
        break;
      case 'competitor_comparison':
        handleCompetitorComparisonData(data);
        break;
      default:
        throw new Error('Unknown data type: ' + data.type);
    }

    // 成功レスポンスを返す
    return ContentService.createTextOutput(
      JSON.stringify({
        status: 'success',
        message: 'Data processed successfully',
        timestamp: new Date().toISOString()
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('Error in doPost: ' + error.message);
    Logger.log(error.stack);

    return ContentService.createTextOutput(
      JSON.stringify({
        status: 'error',
        message: error.message,
        timestamp: new Date().toISOString()
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
```

---

#### タスク2.5: Web Appデプロイ（60分）
**所要時間**: 60分
**優先度**: 🔴 高

**作業内容**:
```
□ デプロイ前の最終確認
  - すべての関数が実装されているか
  - エラーハンドリングが適切か
  - ログ出力が分かりやすいか

□ Web Appとしてデプロイ
  1. GASエディタで「デプロイ」→「新しいデプロイ」
  2. 種類を「ウェブアプリ」に設定
  3. 説明: "Phase 4-2 Dify統合API"
  4. 次のユーザーとして実行: 自分
  5. アクセスできるユーザー: 全員
  6. デプロイを実行

□ Webhook URLを取得
  - デプロイ後に表示されるWeb App URLをコピー
  - showWebhookUrl()を実行して確認

□ setupDifyApiSettings()を実行
  - Dify APIのエンドポイントURLを入力
  - Dify APIのAPI Keyを入力

□ testWebhook()で動作確認
  - evaluation, engagementのテストデータで確認
  - エラーが発生しないか確認
```

**成果物**:
- Web App デプロイ完了
- Webhook URL取得
- Dify API設定完了

**チェックポイント**:
- [ ] デプロイが成功したか
- [ ] Webhook URLが取得できたか
- [ ] testWebhook()が正常に動作するか

---

#### タスク2.6: 内部テスト（90分）
**所要時間**: 90分
**優先度**: 🔴 高

**作業内容**:
```
□ testWebhook()の拡張
  - evaluation, engagement, acceptance_story, competitor_comparisonの
    4種類すべてのテストデータを作成
  - 各データタイプでdoPost()を呼び出し
  - スプレッドシートにデータが正しく書き込まれるか確認

□ エラーケースのテスト
  - 必須フィールド欠損
  - データ型不正
  - 範囲外の値
  - 未知のdata type

□ 親切なエラーハンドラのテスト
  - evidence, risk, next_qのデータを送信
  - エラーメッセージが正しく返ってくるか確認

□ スプレッドシートの確認
  - Candidate_Scoresのデータ確認
  - Candidate_Insightsのデータ確認（16列すべて）
  - Engagement_Logのデータ確認

□ ログ確認
  - エラーログがないか
  - 警告ログの内容は適切か
```

**成果物**:
- テスト結果レポート
- バグ修正リスト（あれば）
- スプレッドシートのスクリーンショット

**チェックポイント**:
- [ ] すべてのテストケースがパスするか
- [ ] スプレッドシートにデータが正しく書き込まれるか
- [ ] エラーメッセージが分かりやすいか
- [ ] ログ出力が適切か

---

### 3.3 Day 2 終了時の状態

**完了したもの**:
- ✅ handleAcceptanceStoryData() 完全実装
- ✅ handleCompetitorComparisonData() 完全実装
- ✅ 親切なエラーハンドラ（3種類）実装
- ✅ doPost関数の更新
- ✅ Web Appデプロイ
- ✅ Webhook URL取得
- ✅ setupDifyApiSettings()実行
- ✅ 内部テスト完了

**未完了のもの（オプション）**:
- ⏳ Dify連携テスト
- ⏳ エラーケースの詳細テスト
- ⏳ ドキュメント作成

---

## 4. Day 3（オプション）: テスト・ドキュメント

### 4.1 Morning Session（09:00-12:00, 3時間）

#### タスク3.1: Dify連携テスト（120分）
**所要時間**: 120分
**優先度**: 🟡 中（オプション）

**作業内容**:
```
□ Difyワークフローの設定
  - WebhookノードをDifyワークフローに追加
  - GASのWebhook URLを設定
  - テストリクエストを送信

□ 4種類のデータタイプでテスト
  - evaluation
  - engagement
  - acceptance_story
  - competitor_comparison

□ エラーケースのテスト
  - 必須フィールド欠損
  - データ型不正
  - 親切なエラーハンドラ（evidence, risk, next_q）

□ 統合テスト
  - Difyワークフローを実行
  - GASがWebhookを正しく受信するか
  - スプレッドシートにデータが正しく書き込まれるか
  - エラーレスポンスが正しく返ってくるか
```

**成果物**:
- Dify連携テスト結果レポート
- バグ修正リスト（あれば）

**チェックポイント**:
- [ ] DifyからGASにWebhookが正しく送信されるか
- [ ] GASが正しくレスポンスを返すか
- [ ] スプレッドシートにデータが正しく書き込まれるか
- [ ] エラーケースで適切なエラーが返るか

---

#### タスク3.2: ドキュメント作成（60分）
**所要時間**: 60分
**優先度**: 🟢 低（オプション）

**作業内容**:
```
□ API仕様書の作成
  - エンドポイント: Webhook URL
  - HTTPメソッド: POST
  - リクエスト形式: JSON
  - レスポンス形式: JSON
  - エラーコード一覧

□ Dify連携ガイドの作成
  - Webhook設定手順
  - データ形式の説明
  - トラブルシューティング

□ 運用手順書の更新
  - setupDifyApiSettings()の使い方
  - testWebhook()の使い方
  - ログ確認方法
```

**成果物**:
- API仕様書
- Dify連携ガイド
- 運用手順書

---

### 4.2 Day 3 終了時の状態

**完了したもの**:
- ✅ Dify連携テスト完了
- ✅ ドキュメント作成完了
- ✅ すべてのタスク完了

---

## 5. チェックポイント

### 5.1 Day 1終了時のチェックポイント

```
□ Candidate_Insightsシートが16列に拡張されているか
□ ヘルパー関数（6-7個）が実装されているか
□ handleEvaluationData()が正常に動作するか
□ handleEngagementData()が正常に動作するか
□ Candidate_Scoresにデータが正しく書き込まれるか
□ Candidate_Insightsにデータが正しく書き込まれるか
□ Engagement_Logにデータが正しく追記されるか
□ エラーハンドリングが適切か
□ ログ出力が分かりやすいか
```

### 5.2 Day 2終了時のチェックポイント

```
□ handleAcceptanceStoryData()が正常に動作するか
□ handleCompetitorComparisonData()が正常に動作するか
□ 親切なエラーハンドラ（3種類）が実装されているか
□ doPost関数が正しく更新されているか
□ Web Appとしてデプロイされているか
□ Webhook URLが取得できているか
□ setupDifyApiSettings()が実行されているか
□ testWebhook()が正常に動作するか
□ すべてのデータタイプでテストが成功するか
□ スプレッドシートにデータが正しく書き込まれるか
```

### 5.3 最終チェックポイント

```
□ 4種類のデータ処理ハンドラが完全に実装されているか
□ 3種類の親切なエラーハンドラが実装されているか
□ すべてのバリデーション関数が実装されているか
□ すべてのヘルパー関数が実装されているか
□ Web Appとしてデプロイされているか
□ testWebhook()が成功するか
□ Dify連携テストが成功するか（オプション）
□ ドキュメントが作成されているか（オプション）
```

---

## 6. タスク依存関係

```
[Day 1]
タスク1.1: スプレッドシート構造の拡張
    ↓
タスク1.2: ヘルパー関数の実装
    ↓
タスク1.3: handleEvaluationData()の実装
    ↓
タスク1.4: handleEngagementData()の実装
    ↓
タスク1.5: Day 1レビューとテスト

[Day 2]
タスク2.1: handleAcceptanceStoryData()の実装
    ‖ (並行可能)
タスク2.2: handleCompetitorComparisonData()の実装
    ↓
タスク2.3: 親切なエラーハンドラの実装
    ↓
タスク2.4: doPost関数の更新
    ↓
タスク2.5: Web Appデプロイ
    ↓
タスク2.6: 内部テスト

[Day 3 - オプション]
タスク3.1: Dify連携テスト
    ↓
タスク3.2: ドキュメント作成
```

---

## 7. リスク管理

### 7.1 高リスク項目

| リスク | 影響 | 対策 |
|--------|------|------|
| スプレッドシート構造の変更失敗 | 🔴 高 | バックアップを作成（phase0_preparation()） |
| Web Appデプロイ失敗 | 🔴 高 | 権限設定を事前確認、手順書に従う |
| データバリデーションの漏れ | 🟡 中 | PHASE_4_2_IMPLEMENTATION_SPEC.mdを参照 |

### 7.2 中リスク項目

| リスク | 影響 | 対策 |
|--------|------|------|
| テストデータ不足 | 🟡 中 | DATA_STRUCTURE_SPEC.mdのサンプルを使用 |
| Dify連携テスト失敗 | 🟡 中 | 内部テスト（testWebhook()）を先に完了させる |
| ドキュメント作成の遅延 | 🟢 低 | Day 3をオプションとして設定 |

---

**以上、Phase 4-2 実装タイムライン（戦略D+）**
