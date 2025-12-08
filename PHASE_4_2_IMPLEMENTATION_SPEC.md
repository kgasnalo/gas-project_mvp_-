# Phase 4-2 実装仕様書（戦略D+）

**作成日**: 2025-12-08
**対象**: DifyIntegration.gs データ処理ハンドラの完全実装
**実装範囲**: 4種類の完全実装 + 3種類の親切なエラーメッセージ

---

## 📋 目次

1. [実装概要](#1-実装概要)
2. [シート構造の拡張](#2-シート構造の拡張)
3. [データ処理ハンドラ仕様](#3-データ処理ハンドラ仕様)
4. [エラーハンドリング方針](#4-エラーハンドリング方針)
5. [ログ出力形式](#5-ログ出力形式)
6. [バリデーションルール](#6-バリデーションルール)
7. [実装チェックリスト](#7-実装チェックリスト)

---

## 1. 実装概要

### 1.1 実装範囲

#### 完全実装（4種類）

1. **handleEvaluationData()** - 評価データ処理
   - 対象シート: Candidate_Scores, Candidate_Insights
   - 目的: Dify評価結果をスコアシートとインサイトシートに記録

2. **handleEngagementData()** - エンゲージメントデータ処理
   - 対象シート: Engagement_Log
   - 目的: 候補者との接触履歴を記録

3. **handleAcceptanceStoryData()** - 承諾ストーリーデータ処理
   - 対象シート: Candidate_Insights（拡張列）
   - 目的: 推奨承諾ストーリーとキーメッセージを記録

4. **handleCompetitorComparisonData()** - 競合比較データ処理
   - 対象シート: Candidate_Insights（拡張列）
   - 目的: 競合分析結果と差別化ポイントを記録

#### 親切なエラーメッセージ（3種類）

5. **handleEvidenceData()** - エビデンスデータ（未実装）
6. **handleRiskData()** - リスクデータ（未実装）
7. **handleNextQData()** - 次の質問データ（未実装）

### 1.2 設計原則

- **YAGNI原則**: Phase 4-2で必要な機能のみを実装
- **ユーザー体験**: 使える機能は完璧に、使えない機能は親切なエラー
- **保守性**: コードは読みやすく、ドキュメントは充実
- **セキュリティ**: 入力データの厳格なバリデーション

---

## 2. シート構造の拡張

### 2.1 Candidate_Insightsシートの列追加

**現在の構造（11列）**:
```
A: candidate_id
B: 氏名
C: 最終更新日時
D: コアモチベーション
E: 主要懸念事項
F: 懸念カテゴリ
G: 競合企業1
H: 競合企業2
I: 競合企業3
J: 次推奨アクション
K: アクション期限
```

**拡張後の構造（16列）**:
```
A: candidate_id
B: 氏名
C: 最終更新日時
D: コアモチベーション
E: 主要懸念事項
F: 懸念カテゴリ
G: 競合企業1
H: 競合企業2
I: 競合企業3
J: 次推奨アクション
K: アクション期限
L: 推奨承諾ストーリー         ← 新規追加（acceptance_story用）
M: キーメッセージ              ← 新規追加（acceptance_story用）
N: 差別化ポイント              ← 新規追加（competitor_comparison用）
O: 競合分析結果                ← 新規追加（competitor_comparison用）
P: 推奨戦略                    ← 新規追加（competitor_comparison用）
```

### 2.2 Engagement_Logシートの構造（既存）

**列構成（14列）**:
```
A: engagement_id
B: candidate_id
C: 氏名
D: 接触日時
E: フェーズ
F: AI予測_承諾可能性
G: 人間の直感_承諾可能性
H: 統合_承諾可能性
I: 信頼度
J: 志望度スコア
K: 競合優位性スコア
L: 懸念解消度スコア
M: コアモチベーション
N: 主要懸念事項
```

### 2.3 Candidate_Scoresシートの構造（既存）

**主要列**:
```
A: candidate_id
B: 氏名
C: 最終更新日時
D: 最新_合格可能性
E: 前回_合格可能性
F: 合格可能性_増減
G: 最新_Philosophy
H: 最新_Strategy
I: 最新_Motivation
J: 最新_Execution
K: 最新_合計スコア
L: 最新_承諾可能性（AI予測）
M: 最新_承諾可能性（人間の直感）
N: 最新_承諾可能性（統合）
O: 前回_承諾可能性
P: 承諾可能性_増減
Q: 予測の信頼度
R: 志望度スコア
S: 競合優位性スコア
T: 懸念解消度スコア
U: アンケート回答速度スコア
```

---

## 3. データ処理ハンドラ仕様

### 3.1 handleEvaluationData(data)

#### 目的
Difyからの評価データを受信し、Candidate_ScoresとCandidate_Insightsシートに書き込む

#### 受信データ形式
```javascript
{
  "type": "evaluation",
  "candidate_id": "C001",
  "timestamp": "2025-12-08T10:30:00Z",
  "scores": {
    "overall": 85,                    // 総合スコア（0-100）
    "philosophy": 90,                 // Philosophy スコア（0-100）
    "strategy": 85,                   // Strategy スコア（0-100）
    "motivation": 88,                 // Motivation スコア（0-100）
    "execution": 82,                  // Execution スコア（0-100）
    "acceptance_probability": 0.78    // 承諾可能性（0-1）
  },
  "insights": {
    "core_motivation": "技術的な挑戦機会と成長環境",
    "main_concern": "待遇面と将来のキャリアパス",
    "concern_category": "キャリア・成長"
  }
}
```

#### 処理フロー
```
1. データバリデーション
   ↓
2. candidate_idでスプレッドシート内の該当行を検索
   ↓
3. Candidate_Scoresシートを更新
   - D列: 最新_合格可能性 = scores.overall
   - G列: 最新_Philosophy = scores.philosophy
   - H列: 最新_Strategy = scores.strategy
   - I列: 最新_Motivation = scores.motivation
   - J列: 最新_Execution = scores.execution
   - K列: 最新_合計スコア = 計算値
   - L列: 最新_承諾可能性（AI予測） = scores.acceptance_probability * 100
   - C列: 最終更新日時 = timestamp
   ↓
4. Candidate_Insightsシートを更新
   - D列: コアモチベーション = insights.core_motivation
   - E列: 主要懸念事項 = insights.main_concern
   - F列: 懸念カテゴリ = insights.concern_category
   - C列: 最終更新日時 = timestamp
   ↓
5. 成功ログ出力
```

#### 実装コード骨組み
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

    // 3. 候補者の行を検索
    const candidateRow = findCandidateRow(scoresSheet, data.candidate_id);
    if (!candidateRow) {
      // 新規候補者の場合は行を追加
      addNewCandidateRow(scoresSheet, data.candidate_id);
      addNewCandidateRow(insightsSheet, data.candidate_id);
    }

    // 4. Candidate_Scoresを更新
    updateCandidateScores(scoresSheet, data);

    // 5. Candidate_Insightsを更新
    updateCandidateInsights(insightsSheet, data);

    Logger.log('✅ Evaluation data processed successfully for: ' + data.candidate_id);

  } catch (error) {
    Logger.log('❌ Error in handleEvaluationData: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}
```

#### バリデーションルール
```javascript
function validateEvaluationData(data) {
  // 必須フィールドチェック
  const requiredFields = ['candidate_id', 'timestamp', 'scores'];
  requiredFields.forEach(field => {
    if (!data[field]) {
      throw new Error(`Missing required field: ${field}`);
    }
  });

  // スコアの範囲チェック（0-100）
  const scores = data.scores;
  ['overall', 'philosophy', 'strategy', 'motivation', 'execution'].forEach(scoreType => {
    if (scores[scoreType] !== undefined) {
      if (scores[scoreType] < 0 || scores[scoreType] > 100) {
        throw new Error(`Invalid score range for ${scoreType}: ${scores[scoreType]}`);
      }
    }
  });

  // 承諾可能性の範囲チェック（0-1）
  if (scores.acceptance_probability !== undefined) {
    if (scores.acceptance_probability < 0 || scores.acceptance_probability > 1) {
      throw new Error(`Invalid acceptance_probability: ${scores.acceptance_probability}`);
    }
  }
}
```

---

### 3.2 handleEngagementData(data)

#### 目的
Difyからのエンゲージメントデータを受信し、Engagement_Logシートに書き込む

#### 受信データ形式
```javascript
{
  "type": "engagement",
  "candidate_id": "C001",
  "timestamp": "2025-12-08T14:30:00Z",
  "engagement": {
    "date": "2025-12-08",
    "phase": "初回面談",
    "ai_prediction": 78,              // AI予測_承諾可能性（0-100）
    "human_intuition": null,          // 人間の直感（0-100、未入力の場合null）
    "confidence": 0.85,               // 信頼度（0-1）
    "motivation_score": 75,           // 志望度スコア（0-100）
    "competitive_score": 60,          // 競合優位性スコア（0-100）
    "concern_score": 80,              // 懸念解消度スコア（0-100）
    "core_motivation": "成長機会",
    "top_concern": "給与水準"
  }
}
```

#### 処理フロー
```
1. データバリデーション
   ↓
2. Engagement_Logシートに新規行を追加（appendRow）
   - A列: engagement_id = 自動生成（"ENG_" + timestamp）
   - B列: candidate_id
   - C列: 氏名（Candidates_Masterから取得）
   - D列: 接触日時 = engagement.date
   - E列: フェーズ = engagement.phase
   - F列: AI予測_承諾可能性 = engagement.ai_prediction
   - G列: 人間の直感_承諾可能性 = engagement.human_intuition
   - H列: 統合_承諾可能性 = 計算値（AI予測 * 0.7 + 人間の直感 * 0.3）
   - I列: 信頼度 = engagement.confidence
   - J列: 志望度スコア = engagement.motivation_score
   - K列: 競合優位性スコア = engagement.competitive_score
   - L列: 懸念解消度スコア = engagement.concern_score
   - M列: コアモチベーション = engagement.core_motivation
   - N列: 主要懸念事項 = engagement.top_concern
   ↓
3. 成功ログ出力
```

#### 実装コード骨組み
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

    // 6. Engagement_Logに新規行を追加
    engagementSheet.appendRow([
      engagementId,                             // A: engagement_id
      data.candidate_id,                        // B: candidate_id
      candidateName,                            // C: 氏名
      data.engagement.date,                     // D: 接触日時
      data.engagement.phase,                    // E: フェーズ
      data.engagement.ai_prediction,            // F: AI予測
      data.engagement.human_intuition || '',    // G: 人間の直感
      integrated,                               // H: 統合
      data.engagement.confidence,               // I: 信頼度
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

---

### 3.3 handleAcceptanceStoryData(data)

#### 目的
Difyからの承諾ストーリーデータを受信し、Candidate_Insightsシートの拡張列に書き込む

#### 受信データ形式
```javascript
{
  "type": "acceptance_story",
  "candidate_id": "C001",
  "timestamp": "2025-12-08T15:00:00Z",
  "story": {
    "recommended_story": "技術的な挑戦を前面に出し、成長機会を強調する",
    "key_messages": [
      "最先端技術への取り組み",
      "自己成長の機会",
      "チームの雰囲気"
    ],
    "recommended_actions": [
      "CTOとの1on1を設定",
      "技術ブログを共有",
      "オフィス見学を提案"
    ],
    "priority": "高",
    "deadline": "2025-12-15"
  }
}
```

#### 処理フロー
```
1. データバリデーション
   ↓
2. candidate_idでCandidate_Insightsシート内の該当行を検索
   ↓
3. Candidate_Insightsシートを更新
   - L列: 推奨承諾ストーリー = story.recommended_story
   - M列: キーメッセージ = story.key_messages（配列を改行区切りで結合）
   - J列: 次推奨アクション = story.recommended_actions（配列を改行区切りで結合）
   - K列: アクション期限 = story.deadline
   - C列: 最終更新日時 = timestamp
   ↓
4. 成功ログ出力
```

#### 実装コード骨組み
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

    // 3. 候補者の行を検索
    const candidateRow = findCandidateRow(insightsSheet, data.candidate_id);
    if (!candidateRow) {
      // 新規候補者の場合は行を追加
      addNewCandidateRow(insightsSheet, data.candidate_id);
    }

    // 4. 配列データを改行区切りの文字列に変換
    const keyMessages = data.story.key_messages.join('\n');
    const recommendedActions = data.story.recommended_actions.join('\n');

    // 5. Candidate_Insightsを更新
    const rowIndex = candidateRow || getLastRow(insightsSheet) + 1;
    insightsSheet.getRange(rowIndex, 3).setValue(data.timestamp);           // C: 最終更新日時
    insightsSheet.getRange(rowIndex, 12).setValue(data.story.recommended_story); // L: 推奨承諾ストーリー
    insightsSheet.getRange(rowIndex, 13).setValue(keyMessages);             // M: キーメッセージ
    insightsSheet.getRange(rowIndex, 10).setValue(recommendedActions);      // J: 次推奨アクション
    insightsSheet.getRange(rowIndex, 11).setValue(data.story.deadline);     // K: アクション期限

    Logger.log('✅ Acceptance story data processed successfully for: ' + data.candidate_id);

  } catch (error) {
    Logger.log('❌ Error in handleAcceptanceStoryData: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}
```

---

### 3.4 handleCompetitorComparisonData(data)

#### 目的
Difyからの競合比較データを受信し、Candidate_Insightsシートの拡張列に書き込む

#### 受信データ形式
```javascript
{
  "type": "competitor_comparison",
  "candidate_id": "C001",
  "timestamp": "2025-12-08T16:00:00Z",
  "comparison": {
    "competitors": [
      {
        "company": "A社",
        "salary_comparison": "やや劣る",
        "tech_comparison": "優位",
        "work_life_balance": "優位"
      },
      {
        "company": "B社",
        "salary_comparison": "同等",
        "tech_comparison": "同等",
        "work_life_balance": "優位"
      }
    ],
    "differentiation_points": [
      "技術的な挑戦機会",
      "フレックス制度",
      "リモートワーク可"
    ],
    "recommended_strategy": "技術力とワークライフバランスを強調"
  }
}
```

#### 処理フロー
```
1. データバリデーション
   ↓
2. candidate_idでCandidate_Insightsシート内の該当行を検索
   ↓
3. 競合企業名を抽出
   - G列: 競合企業1 = competitors[0].company
   - H列: 競合企業2 = competitors[1].company（存在する場合）
   - I列: 競合企業3 = competitors[2].company（存在する場合）
   ↓
4. Candidate_Insightsシートの拡張列を更新
   - N列: 差別化ポイント = differentiation_points（配列を改行区切りで結合）
   - O列: 競合分析結果 = competitors情報を整形した文字列
   - P列: 推奨戦略 = recommended_strategy
   - C列: 最終更新日時 = timestamp
   ↓
5. 成功ログ出力
```

#### 実装コード骨組み
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

    // 3. 候補者の行を検索
    const candidateRow = findCandidateRow(insightsSheet, data.candidate_id);
    if (!candidateRow) {
      // 新規候補者の場合は行を追加
      addNewCandidateRow(insightsSheet, data.candidate_id);
    }

    // 4. 競合分析結果を整形
    const comparisonResult = formatComparisonResult(data.comparison.competitors);
    const differentiationPoints = data.comparison.differentiation_points.join('\n');

    // 5. Candidate_Insightsを更新
    const rowIndex = candidateRow || getLastRow(insightsSheet) + 1;

    // 競合企業名を設定
    if (data.comparison.competitors[0]) {
      insightsSheet.getRange(rowIndex, 7).setValue(data.comparison.competitors[0].company); // G: 競合企業1
    }
    if (data.comparison.competitors[1]) {
      insightsSheet.getRange(rowIndex, 8).setValue(data.comparison.competitors[1].company); // H: 競合企業2
    }
    if (data.comparison.competitors[2]) {
      insightsSheet.getRange(rowIndex, 9).setValue(data.comparison.competitors[2].company); // I: 競合企業3
    }

    // 拡張列を設定
    insightsSheet.getRange(rowIndex, 3).setValue(data.timestamp);                        // C: 最終更新日時
    insightsSheet.getRange(rowIndex, 14).setValue(differentiationPoints);                // N: 差別化ポイント
    insightsSheet.getRange(rowIndex, 15).setValue(comparisonResult);                     // O: 競合分析結果
    insightsSheet.getRange(rowIndex, 16).setValue(data.comparison.recommended_strategy); // P: 推奨戦略

    Logger.log('✅ Competitor comparison data processed successfully for: ' + data.candidate_id);

  } catch (error) {
    Logger.log('❌ Error in handleCompetitorComparisonData: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}

// ヘルパー関数: 競合分析結果を整形
function formatComparisonResult(competitors) {
  return competitors.map(comp => {
    return `${comp.company}: 給与${comp.salary_comparison}, 技術${comp.tech_comparison}, WLB${comp.work_life_balance}`;
  }).join('\n');
}
```

---

### 3.5 親切なエラーメッセージハンドラ（3種類）

#### handleEvidenceData(data)

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
```

#### handleRiskData(data)

```javascript
function handleRiskData(data) {
  const message = {
    status: 'not_implemented',
    message: 'Risk分析機能は将来のバージョンで実装予定です',
    data_type: 'risk',
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

  Logger.log('⚠️ Risk分析機能は未実装です: ' + JSON.stringify(message));

  return ContentService.createTextOutput(
    JSON.stringify(message, null, 2)
  ).setMimeType(ContentService.MimeType.JSON);
}
```

#### handleNextQData(data)

```javascript
function handleNextQData(data) {
  const message = {
    status: 'not_implemented',
    message: '次の質問生成機能は将来のバージョンで実装予定です',
    data_type: 'next_q',
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

  Logger.log('⚠️ 次の質問生成機能は未実装です: ' + JSON.stringify(message));

  return ContentService.createTextOutput(
    JSON.stringify(message, null, 2)
  ).setMimeType(ContentService.MimeType.JSON);
}
```

---

## 4. エラーハンドリング方針

### 4.1 エラーの分類

#### レベル1: データバリデーションエラー
- 必須フィールド欠損
- データ型不正
- 範囲外の値

**対応**: 即座にエラーレスポンスを返す

#### レベル2: シートアクセスエラー
- シートが見つからない
- 権限エラー

**対応**: エラーログを出力し、管理者に通知

#### レベル3: データ書き込みエラー
- スプレッドシートAPI制限
- 行/列範囲エラー

**対応**: リトライロジックまたはエラーレスポンス

### 4.2 エラーレスポンス形式

```javascript
{
  "status": "error",
  "error_code": "VALIDATION_ERROR",
  "message": "Missing required field: candidate_id",
  "timestamp": "2025-12-08T10:30:00Z",
  "data_type": "evaluation",
  "details": {
    "missing_fields": ["candidate_id"],
    "received_data": { /* ... */ }
  }
}
```

### 4.3 エラーコード一覧

```
VALIDATION_ERROR        - データバリデーションエラー
MISSING_FIELD           - 必須フィールド欠損
INVALID_DATA_TYPE       - データ型不正
OUT_OF_RANGE            - 値が範囲外
SHEET_NOT_FOUND         - シートが見つからない
CANDIDATE_NOT_FOUND     - 候補者が見つからない
PERMISSION_DENIED       - 権限エラー
WRITE_ERROR             - 書き込みエラー
UNKNOWN_ERROR           - 不明なエラー
```

---

## 5. ログ出力形式

### 5.1 成功ログ

```
✅ Evaluation data processed successfully for: C001
   - Updated Candidate_Scores: D, G, H, I, J, K, L columns
   - Updated Candidate_Insights: D, E, F columns
   - Timestamp: 2025-12-08T10:30:00Z
```

### 5.2 エラーログ

```
❌ Error in handleEvaluationData: Missing required field: candidate_id
   - Data type: evaluation
   - Received data: {...}
   - Stack trace: ...
```

### 5.3 警告ログ

```
⚠️ Evidence機能は未実装です
   - Candidate ID: C001
   - 利用可能な機能: evaluation, engagement, acceptance_story, competitor_comparison
```

---

## 6. バリデーションルール

### 6.1 共通バリデーション

```javascript
function validateCommonFields(data) {
  // candidate_idは必須かつ非空
  if (!data.candidate_id || data.candidate_id.trim() === '') {
    throw new Error('candidate_id is required and cannot be empty');
  }

  // timestampは有効な日付形式
  if (data.timestamp && !isValidTimestamp(data.timestamp)) {
    throw new Error('Invalid timestamp format: ' + data.timestamp);
  }
}

function isValidTimestamp(timestamp) {
  const date = new Date(timestamp);
  return date instanceof Date && !isNaN(date);
}
```

### 6.2 evaluationデータのバリデーション

```javascript
function validateEvaluationData(data) {
  validateCommonFields(data);

  // scoresオブジェクトは必須
  if (!data.scores) {
    throw new Error('scores object is required');
  }

  // スコアは0-100の範囲
  const scoreFields = ['overall', 'philosophy', 'strategy', 'motivation', 'execution'];
  scoreFields.forEach(field => {
    if (data.scores[field] !== undefined) {
      if (typeof data.scores[field] !== 'number' || data.scores[field] < 0 || data.scores[field] > 100) {
        throw new Error(`Invalid ${field} score: ${data.scores[field]} (must be 0-100)`);
      }
    }
  });

  // acceptance_probabilityは0-1の範囲
  if (data.scores.acceptance_probability !== undefined) {
    if (typeof data.scores.acceptance_probability !== 'number' ||
        data.scores.acceptance_probability < 0 ||
        data.scores.acceptance_probability > 1) {
      throw new Error(`Invalid acceptance_probability: ${data.scores.acceptance_probability} (must be 0-1)`);
    }
  }
}
```

### 6.3 engagementデータのバリデーション

```javascript
function validateEngagementData(data) {
  validateCommonFields(data);

  // engagementオブジェクトは必須
  if (!data.engagement) {
    throw new Error('engagement object is required');
  }

  // 日付は必須
  if (!data.engagement.date) {
    throw new Error('engagement.date is required');
  }

  // フェーズは必須
  if (!data.engagement.phase) {
    throw new Error('engagement.phase is required');
  }

  // 各スコアは0-100の範囲
  const scoreFields = ['ai_prediction', 'motivation_score', 'competitive_score', 'concern_score'];
  scoreFields.forEach(field => {
    if (data.engagement[field] !== undefined && data.engagement[field] !== null) {
      if (typeof data.engagement[field] !== 'number' ||
          data.engagement[field] < 0 ||
          data.engagement[field] > 100) {
        throw new Error(`Invalid ${field}: ${data.engagement[field]} (must be 0-100)`);
      }
    }
  });

  // 信頼度は0-1の範囲
  if (data.engagement.confidence !== undefined) {
    if (typeof data.engagement.confidence !== 'number' ||
        data.engagement.confidence < 0 ||
        data.engagement.confidence > 1) {
      throw new Error(`Invalid confidence: ${data.engagement.confidence} (must be 0-1)`);
    }
  }
}
```

### 6.4 acceptance_storyデータのバリデーション

```javascript
function validateAcceptanceStoryData(data) {
  validateCommonFields(data);

  // storyオブジェクトは必須
  if (!data.story) {
    throw new Error('story object is required');
  }

  // recommended_storyは必須
  if (!data.story.recommended_story) {
    throw new Error('story.recommended_story is required');
  }

  // key_messagesは配列
  if (data.story.key_messages && !Array.isArray(data.story.key_messages)) {
    throw new Error('story.key_messages must be an array');
  }

  // recommended_actionsは配列
  if (data.story.recommended_actions && !Array.isArray(data.story.recommended_actions)) {
    throw new Error('story.recommended_actions must be an array');
  }
}
```

### 6.5 competitor_comparisonデータのバリデーション

```javascript
function validateCompetitorComparisonData(data) {
  validateCommonFields(data);

  // comparisonオブジェクトは必須
  if (!data.comparison) {
    throw new Error('comparison object is required');
  }

  // competitorsは配列
  if (!Array.isArray(data.comparison.competitors)) {
    throw new Error('comparison.competitors must be an array');
  }

  // differentiation_pointsは配列
  if (data.comparison.differentiation_points &&
      !Array.isArray(data.comparison.differentiation_points)) {
    throw new Error('comparison.differentiation_points must be an array');
  }
}
```

---

## 7. 実装チェックリスト

### 7.1 事前準備

- [ ] SpreadsheetRedesign.gsを修正してCandidate_Insightsシートを16列に拡張
- [ ] phase0_preparation()を実行してバックアップ作成
- [ ] phase1_execute()を実行してシート構造を更新

### 7.2 ヘルパー関数の実装

- [ ] findCandidateRow() - 候補者の行を検索
- [ ] addNewCandidateRow() - 新規候補者の行を追加
- [ ] getCandidateName() - 候補者の氏名を取得
- [ ] calculateIntegratedAcceptance() - 統合承諾可能性を計算
- [ ] formatComparisonResult() - 競合分析結果を整形

### 7.3 バリデーション関数の実装

- [ ] validateCommonFields()
- [ ] validateEvaluationData()
- [ ] validateEngagementData()
- [ ] validateAcceptanceStoryData()
- [ ] validateCompetitorComparisonData()
- [ ] isValidTimestamp()

### 7.4 データ処理ハンドラの実装

- [ ] handleEvaluationData()
  - [ ] データバリデーション
  - [ ] Candidate_Scoresシート更新
  - [ ] Candidate_Insightsシート更新
  - [ ] エラーハンドリング
  - [ ] ログ出力

- [ ] handleEngagementData()
  - [ ] データバリデーション
  - [ ] Engagement_Logシート追記
  - [ ] エラーハンドリング
  - [ ] ログ出力

- [ ] handleAcceptanceStoryData()
  - [ ] データバリデーション
  - [ ] Candidate_Insightsシート更新（拡張列）
  - [ ] エラーハンドリング
  - [ ] ログ出力

- [ ] handleCompetitorComparisonData()
  - [ ] データバリデーション
  - [ ] Candidate_Insightsシート更新（拡張列）
  - [ ] エラーハンドリング
  - [ ] ログ出力

### 7.5 エラーメッセージハンドラの実装

- [ ] handleEvidenceData() - 親切なエラーメッセージ
- [ ] handleRiskData() - 親切なエラーメッセージ
- [ ] handleNextQData() - 親切なエラーメッセージ

### 7.6 テスト

- [ ] testWebhook() - 内部テスト実行
- [ ] 各データタイプのテストデータを作成
- [ ] エラーケースのテスト
- [ ] ログ確認

### 7.7 デプロイ

- [ ] Web Appとしてデプロイ
- [ ] Webhook URLを取得
- [ ] setupDifyApiSettings()を実行
- [ ] Dify連携テスト

---

## 8. 補足事項

### 8.1 パフォーマンス考慮事項

- スプレッドシートAPIの呼び出し回数を最小化
- バッチ更新の活用（複数セルを一度に更新）
- キャッシュの活用（候補者名の取得など）

### 8.2 セキュリティ考慮事項

- 入力データのサニタイゼーション
- SQL構文エラーの防止（シングルクォートのエスケープ）
- API制限への対応

### 8.3 拡張性の考慮

- 新しいデータタイプの追加が容易
- シート構造の変更に柔軟に対応
- ログ出力の一元管理

---

**以上、Phase 4-2 実装仕様書（戦略D+）**
