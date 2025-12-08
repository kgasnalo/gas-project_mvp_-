# データ構造仕様書（Phase 4-2）

**作成日**: 2025-12-08
**対象**: Dify → GAS Web API データフロー
**バージョン**: 1.0

---

## 📋 目次

1. [データフロー概要](#1-データフロー概要)
2. [evaluation データ構造](#2-evaluation-データ構造)
3. [engagement データ構造](#3-engagement-データ構造)
4. [acceptance_story データ構造](#4-acceptance_story-データ構造)
5. [competitor_comparison データ構造](#5-competitor_comparison-データ構造)
6. [スプレッドシートマッピング](#6-スプレッドシートマッピング)
7. [データ変換ルール](#7-データ変換ルール)

---

## 1. データフロー概要

### 1.1 全体フロー

```
Dify Workflow
    ↓ (POST request)
GAS Web API (doPost関数)
    ↓ (JSON parse)
データタイプ別の分岐
    ├→ evaluation → Candidate_Scores, Candidate_Insights
    ├→ engagement → Engagement_Log
    ├→ acceptance_story → Candidate_Insights（拡張列）
    └→ competitor_comparison → Candidate_Insights（拡張列）
```

### 1.2 共通フィールド

すべてのデータタイプに共通するフィールド:

```javascript
{
  "type": "データタイプ名",        // 必須: evaluation, engagement, etc.
  "candidate_id": "C001",         // 必須: 候補者ID
  "timestamp": "2025-12-08T10:30:00Z"  // 必須: ISO 8601形式のタイムスタンプ
}
```

---

## 2. evaluation データ構造

### 2.1 完全なデータ例

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
    "total": 345,                     // 合計スコア（philosophy + strategy + motivation + execution）
    "acceptance_probability": 0.78    // 承諾可能性（0-1）
  },
  "insights": {
    "core_motivation": "技術的な挑戦機会と成長環境を重視。最先端技術への興味が強い。",
    "main_concern": "待遇面と将来のキャリアパス。特に給与水準と昇進機会について懸念。",
    "concern_category": "キャリア・成長"
  }
}
```

### 2.2 フィールド仕様

#### scores オブジェクト

| フィールド名 | 型 | 必須 | 範囲 | 説明 |
|-------------|-----|------|------|------|
| overall | number | ✅ | 0-100 | 総合スコア |
| philosophy | number | ⚪ | 0-100 | 価値観の一致度 |
| strategy | number | ⚪ | 0-100 | 戦略的思考力 |
| motivation | number | ⚪ | 0-100 | モチベーション |
| execution | number | ⚪ | 0-100 | 実行力 |
| total | number | ⚪ | 0-400 | 合計スコア（自動計算可） |
| acceptance_probability | number | ⚪ | 0-1 | 承諾可能性（0.0-1.0） |

#### insights オブジェクト

| フィールド名 | 型 | 必須 | 最大長 | 説明 |
|-------------|-----|------|--------|------|
| core_motivation | string | ⚪ | 500 | コアモチベーション |
| main_concern | string | ⚪ | 500 | 主要懸念事項 |
| concern_category | string | ⚪ | 100 | 懸念カテゴリ |

### 2.3 スプレッドシートへのマッピング

#### Candidate_Scores シート

| 列 | 列名 | データソース | 変換ルール |
|----|------|--------------|------------|
| A | candidate_id | data.candidate_id | そのまま |
| B | 氏名 | Candidates_Masterから取得 | - |
| C | 最終更新日時 | data.timestamp | ISO 8601 → 日本時間 |
| D | 最新_合格可能性 | data.scores.overall | そのまま |
| G | 最新_Philosophy | data.scores.philosophy | そのまま |
| H | 最新_Strategy | data.scores.strategy | そのまま |
| I | 最新_Motivation | data.scores.motivation | そのまま |
| J | 最新_Execution | data.scores.execution | そのまま |
| K | 最新_合計スコア | data.scores.total | P+S+M+E |
| L | 最新_承諾可能性（AI予測） | data.scores.acceptance_probability | 0-1 → 0-100（×100） |

#### Candidate_Insights シート

| 列 | 列名 | データソース | 変換ルール |
|----|------|--------------|------------|
| A | candidate_id | data.candidate_id | そのまま |
| B | 氏名 | Candidates_Masterから取得 | - |
| C | 最終更新日時 | data.timestamp | ISO 8601 → 日本時間 |
| D | コアモチベーション | data.insights.core_motivation | そのまま |
| E | 主要懸念事項 | data.insights.main_concern | そのまま |
| F | 懸念カテゴリ | data.insights.concern_category | そのまま |

---

## 3. engagement データ構造

### 3.1 完全なデータ例

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
    "core_motivation": "成長機会を重視",
    "top_concern": "給与水準"
  }
}
```

### 3.2 フィールド仕様

#### engagement オブジェクト

| フィールド名 | 型 | 必須 | 範囲/形式 | 説明 |
|-------------|-----|------|-----------|------|
| date | string | ✅ | YYYY-MM-DD | 接触日時 |
| phase | string | ✅ | - | フェーズ（初回面談、社員面談、2次面接、内定後） |
| ai_prediction | number | ✅ | 0-100 | AI予測_承諾可能性 |
| human_intuition | number/null | ⚪ | 0-100 | 人間の直感_承諾可能性 |
| confidence | number | ⚪ | 0-1 | 信頼度 |
| motivation_score | number | ⚪ | 0-100 | 志望度スコア |
| competitive_score | number | ⚪ | 0-100 | 競合優位性スコア |
| concern_score | number | ⚪ | 0-100 | 懸念解消度スコア |
| core_motivation | string | ⚪ | 500 | コアモチベーション |
| top_concern | string | ⚪ | 500 | 主要懸念事項 |

### 3.3 スプレッドシートへのマッピング

#### Engagement_Log シート（appendRowで追記）

| 列 | 列名 | データソース | 変換ルール |
|----|------|--------------|------------|
| A | engagement_id | 自動生成 | "ENG_" + timestamp |
| B | candidate_id | data.candidate_id | そのまま |
| C | 氏名 | Candidates_Masterから取得 | - |
| D | 接触日時 | data.engagement.date | そのまま |
| E | フェーズ | data.engagement.phase | そのまま |
| F | AI予測_承諾可能性 | data.engagement.ai_prediction | そのまま |
| G | 人間の直感_承諾可能性 | data.engagement.human_intuition | null → 空文字列 |
| H | 統合_承諾可能性 | 計算値 | AI × 0.7 + 人間 × 0.3 |
| I | 信頼度 | data.engagement.confidence | 0-1 → 0-100（×100） |
| J | 志望度スコア | data.engagement.motivation_score | そのまま |
| K | 競合優位性スコア | data.engagement.competitive_score | そのまま |
| L | 懸念解消度スコア | data.engagement.concern_score | そのまま |
| M | コアモチベーション | data.engagement.core_motivation | そのまま |
| N | 主要懸念事項 | data.engagement.top_concern | そのまま |

### 3.4 統合承諾可能性の計算式

```javascript
function calculateIntegratedAcceptance(aiPrediction, humanIntuition) {
  // 人間の直感がnullまたは未入力の場合はAI予測のみを使用
  if (humanIntuition === null || humanIntuition === undefined || humanIntuition === '') {
    return aiPrediction;
  }

  // 重み付け平均: AI 70% + 人間 30%
  return Math.round(aiPrediction * 0.7 + humanIntuition * 0.3);
}
```

---

## 4. acceptance_story データ構造

### 4.1 完全なデータ例

```javascript
{
  "type": "acceptance_story",
  "candidate_id": "C001",
  "timestamp": "2025-12-08T15:00:00Z",
  "story": {
    "recommended_story": "技術的な挑戦を前面に出し、成長機会を強調する戦略が最適です。候補者は最先端技術への興味が強く、自己成長を重視する傾向があります。待遇面の懸念については、キャリアパスと昇進機会を明確に示すことで解消できます。",
    "key_messages": [
      "最先端技術（AI/ML）への取り組み",
      "技術的な自由度と裁量権",
      "自己成長の機会（研修・カンファレンス参加）",
      "明確なキャリアパス（tech lead → architect → CTO）"
    ],
    "recommended_actions": [
      "CTOとの1on1ミーティングを設定（技術的な議論）",
      "社内の技術ブログや最新プロジェクトを共有",
      "オフィス見学とチームメンバーとのランチ",
      "具体的なキャリアパスと昇進事例を提示",
      "研修制度と技術カンファレンス参加実績を説明"
    ],
    "priority": "高",
    "deadline": "2025-12-15",
    "notes": "候補者は他社からも内定を受けている可能性が高いため、早期のアクション実施が重要"
  }
}
```

### 4.2 フィールド仕様

#### story オブジェクト

| フィールド名 | 型 | 必須 | 最大長 | 説明 |
|-------------|-----|------|--------|------|
| recommended_story | string | ✅ | 2000 | 推奨承諾ストーリー（全体戦略） |
| key_messages | array[string] | ⚪ | - | キーメッセージ（3-5個推奨） |
| recommended_actions | array[string] | ⚪ | - | 推奨アクション（3-7個推奨） |
| priority | string | ⚪ | 10 | 優先度（高/中/低） |
| deadline | string | ⚪ | YYYY-MM-DD | アクション期限 |
| notes | string | ⚪ | 1000 | 備考・注意事項 |

### 4.3 スプレッドシートへのマッピング

#### Candidate_Insights シート（拡張列）

| 列 | 列名 | データソース | 変換ルール |
|----|------|--------------|------------|
| A | candidate_id | data.candidate_id | そのまま |
| B | 氏名 | Candidates_Masterから取得 | - |
| C | 最終更新日時 | data.timestamp | ISO 8601 → 日本時間 |
| J | 次推奨アクション | data.story.recommended_actions | 配列 → 改行区切り |
| K | アクション期限 | data.story.deadline | そのまま |
| L | 推奨承諾ストーリー | data.story.recommended_story | そのまま |
| M | キーメッセージ | data.story.key_messages | 配列 → 改行区切り |

### 4.4 配列データの変換

```javascript
// key_messages配列を改行区切りの文字列に変換
const keyMessages = data.story.key_messages.join('\n');
// 例:
// "最先端技術（AI/ML）への取り組み
// 技術的な自由度と裁量権
// 自己成長の機会（研修・カンファレンス参加）"

// recommended_actions配列を改行区切りの文字列に変換
const recommendedActions = data.story.recommended_actions.join('\n');
// 例:
// "CTOとの1on1ミーティングを設定（技術的な議論）
// 社内の技術ブログや最新プロジェクトを共有
// オフィス見学とチームメンバーとのランチ"
```

---

## 5. competitor_comparison データ構造

### 5.1 完全なデータ例

```javascript
{
  "type": "competitor_comparison",
  "candidate_id": "C001",
  "timestamp": "2025-12-08T16:00:00Z",
  "comparison": {
    "competitors": [
      {
        "company": "A社（大手IT企業）",
        "salary_comparison": "やや劣る（年収差: -50万円）",
        "tech_comparison": "優位（最新技術への取り組み）",
        "work_life_balance": "優位（フルリモート可）",
        "brand_comparison": "劣る（知名度）",
        "growth_comparison": "優位（昇進スピード）"
      },
      {
        "company": "B社（スタートアップ）",
        "salary_comparison": "同等",
        "tech_comparison": "同等",
        "work_life_balance": "優位（勤務時間の柔軟性）",
        "brand_comparison": "やや優位（業界内評価）",
        "growth_comparison": "優位（裁量権）"
      },
      {
        "company": "C社（コンサルティング）",
        "salary_comparison": "優位（年収差: +100万円）",
        "tech_comparison": "劣る（レガシー技術）",
        "work_life_balance": "劣る（長時間労働）",
        "brand_comparison": "同等",
        "growth_comparison": "劣る（技術的成長）"
      }
    ],
    "differentiation_points": [
      "技術的な挑戦機会（最先端AI/ML技術）",
      "フレックス制度とフルリモート勤務",
      "少数精鋭チームでの高い裁量権",
      "明確なキャリアパス（tech lead → CTO）",
      "技術カンファレンス参加支援"
    ],
    "recommended_strategy": "技術力とワークライフバランスを最大の強みとして強調。給与面ではキャリアパスと昇進機会で補完。大手企業との知名度差は、技術的な自由度と裁量権で差別化。",
    "risk_assessment": {
      "high_risk_competitors": ["C社（高給与）"],
      "moderate_risk_competitors": ["A社（ブランド）"],
      "low_risk_competitors": ["B社"],
      "overall_risk": "中"
    }
  }
}
```

### 5.2 フィールド仕様

#### comparison オブジェクト

| フィールド名 | 型 | 必須 | 説明 |
|-------------|-----|------|------|
| competitors | array[object] | ✅ | 競合企業リスト（1-3社推奨） |
| differentiation_points | array[string] | ⚪ | 差別化ポイント（3-5個推奨） |
| recommended_strategy | string | ⚪ | 推奨戦略（全体方針） |
| risk_assessment | object | ⚪ | リスク評価 |

#### competitors 配列の各オブジェクト

| フィールド名 | 型 | 必須 | 説明 |
|-------------|-----|------|------|
| company | string | ✅ | 競合企業名 |
| salary_comparison | string | ⚪ | 給与比較 |
| tech_comparison | string | ⚪ | 技術力比較 |
| work_life_balance | string | ⚪ | ワークライフバランス比較 |
| brand_comparison | string | ⚪ | ブランド比較 |
| growth_comparison | string | ⚪ | 成長機会比較 |

### 5.3 スプレッドシートへのマッピング

#### Candidate_Insights シート（拡張列）

| 列 | 列名 | データソース | 変換ルール |
|----|------|--------------|------------|
| A | candidate_id | data.candidate_id | そのまま |
| B | 氏名 | Candidates_Masterから取得 | - |
| C | 最終更新日時 | data.timestamp | ISO 8601 → 日本時間 |
| G | 競合企業1 | data.comparison.competitors[0].company | 企業名のみ抽出 |
| H | 競合企業2 | data.comparison.competitors[1].company | 企業名のみ抽出 |
| I | 競合企業3 | data.comparison.competitors[2].company | 企業名のみ抽出 |
| N | 差別化ポイント | data.comparison.differentiation_points | 配列 → 改行区切り |
| O | 競合分析結果 | data.comparison.competitors | オブジェクト → 整形文字列 |
| P | 推奨戦略 | data.comparison.recommended_strategy | そのまま |

### 5.4 競合分析結果の整形

```javascript
function formatComparisonResult(competitors) {
  return competitors.map(comp => {
    const lines = [
      `【${comp.company}】`,
      `給与: ${comp.salary_comparison}`,
      `技術: ${comp.tech_comparison}`,
      `WLB: ${comp.work_life_balance}`,
      `ブランド: ${comp.brand_comparison || '-'}`,
      `成長: ${comp.growth_comparison || '-'}`
    ];
    return lines.join('\n');
  }).join('\n\n');
}

// 出力例:
// 【A社（大手IT企業）】
// 給与: やや劣る（年収差: -50万円）
// 技術: 優位（最新技術への取り組み）
// WLB: 優位（フルリモート可）
// ブランド: 劣る（知名度）
// 成長: 優位（昇進スピード）
//
// 【B社（スタートアップ）】
// 給与: 同等
// 技術: 同等
// WLB: 優位（勤務時間の柔軟性）
// ブランド: やや優位（業界内評価）
// 成長: 優位（裁量権）
```

---

## 6. スプレッドシートマッピング一覧

### 6.1 Candidate_Scores シート（21列）

| 列 | 列名 | データソース | データタイプ |
|----|------|--------------|--------------|
| A | candidate_id | 全データタイプ | evaluation |
| B | 氏名 | Candidates_Master | - |
| C | 最終更新日時 | 全データタイプ | evaluation |
| D | 最新_合格可能性 | scores.overall | evaluation |
| E | 前回_合格可能性 | （手動入力） | - |
| F | 合格可能性_増減 | （計算式） | - |
| G | 最新_Philosophy | scores.philosophy | evaluation |
| H | 最新_Strategy | scores.strategy | evaluation |
| I | 最新_Motivation | scores.motivation | evaluation |
| J | 最新_Execution | scores.execution | evaluation |
| K | 最新_合計スコア | scores.total | evaluation |
| L | 最新_承諾可能性（AI予測） | scores.acceptance_probability | evaluation |
| M | 最新_承諾可能性（人間の直感） | （手動入力） | - |
| N | 最新_承諾可能性（統合） | （計算式） | - |
| O | 前回_承諾可能性 | （手動入力） | - |
| P | 承諾可能性_増減 | （計算式） | - |
| Q | 予測の信頼度 | （手動入力） | - |
| R | 志望度スコア | engagement.motivation_score | engagement |
| S | 競合優位性スコア | engagement.competitive_score | engagement |
| T | 懸念解消度スコア | engagement.concern_score | engagement |
| U | アンケート回答速度スコア | （自動計算） | - |

### 6.2 Candidate_Insights シート（16列 - 拡張後）

| 列 | 列名 | データソース | データタイプ |
|----|------|--------------|--------------|
| A | candidate_id | 全データタイプ | 全て |
| B | 氏名 | Candidates_Master | - |
| C | 最終更新日時 | 全データタイプ | 全て |
| D | コアモチベーション | insights.core_motivation | evaluation |
| E | 主要懸念事項 | insights.main_concern | evaluation |
| F | 懸念カテゴリ | insights.concern_category | evaluation |
| G | 競合企業1 | comparison.competitors[0].company | competitor_comparison |
| H | 競合企業2 | comparison.competitors[1].company | competitor_comparison |
| I | 競合企業3 | comparison.competitors[2].company | competitor_comparison |
| J | 次推奨アクション | story.recommended_actions | acceptance_story |
| K | アクション期限 | story.deadline | acceptance_story |
| L | 推奨承諾ストーリー | story.recommended_story | acceptance_story |
| M | キーメッセージ | story.key_messages | acceptance_story |
| N | 差別化ポイント | comparison.differentiation_points | competitor_comparison |
| O | 競合分析結果 | comparison.competitors | competitor_comparison |
| P | 推奨戦略 | comparison.recommended_strategy | competitor_comparison |

### 6.3 Engagement_Log シート（14列）

| 列 | 列名 | データソース | データタイプ |
|----|------|--------------|--------------|
| A | engagement_id | 自動生成 | engagement |
| B | candidate_id | candidate_id | engagement |
| C | 氏名 | Candidates_Master | - |
| D | 接触日時 | engagement.date | engagement |
| E | フェーズ | engagement.phase | engagement |
| F | AI予測_承諾可能性 | engagement.ai_prediction | engagement |
| G | 人間の直感_承諾可能性 | engagement.human_intuition | engagement |
| H | 統合_承諾可能性 | 計算値 | engagement |
| I | 信頼度 | engagement.confidence | engagement |
| J | 志望度スコア | engagement.motivation_score | engagement |
| K | 競合優位性スコア | engagement.competitive_score | engagement |
| L | 懸念解消度スコア | engagement.concern_score | engagement |
| M | コアモチベーション | engagement.core_motivation | engagement |
| N | 主要懸念事項 | engagement.top_concern | engagement |

---

## 7. データ変換ルール

### 7.1 数値変換

#### 0-1 → 0-100 変換

```javascript
// 承諾可能性: 0-1 → 0-100
function convertProbabilityToPercentage(probability) {
  if (probability === null || probability === undefined) {
    return '';
  }
  return Math.round(probability * 100);
}

// 例:
// 0.78 → 78
// 0.854 → 85
// 1.0 → 100
```

#### 信頼度: 0-1 → 0-100 変換

```javascript
function convertConfidenceToPercentage(confidence) {
  if (confidence === null || confidence === undefined) {
    return '';
  }
  return Math.round(confidence * 100);
}
```

### 7.2 日時変換

#### ISO 8601 → 日本時間

```javascript
function convertISO8601ToJST(iso8601String) {
  if (!iso8601String) {
    return '';
  }

  const date = new Date(iso8601String);

  // 日本時間に変換（UTC+9）
  const jstDate = new Date(date.getTime() + (9 * 60 * 60 * 1000));

  // "YYYY-MM-DD HH:mm:ss" 形式
  return Utilities.formatDate(jstDate, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

// 例:
// "2025-12-08T10:30:00Z" → "2025-12-08 19:30:00"
```

### 7.3 配列変換

#### 配列 → 改行区切り文字列

```javascript
function convertArrayToNewlineSeparated(array) {
  if (!array || !Array.isArray(array)) {
    return '';
  }

  return array.join('\n');
}

// 例:
// ["項目1", "項目2", "項目3"] → "項目1\n項目2\n項目3"
```

### 7.4 null/undefined処理

#### null → 空文字列

```javascript
function handleNullValue(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return value;
}
```

### 7.5 企業名抽出

#### 企業名から括弧内を除去（オプション）

```javascript
function extractCompanyName(companyWithDescription) {
  if (!companyWithDescription) {
    return '';
  }

  // "A社（大手IT企業）" → "A社"
  const match = companyWithDescription.match(/^([^（]+)/);
  return match ? match[1] : companyWithDescription;
}
```

---

## 8. データ検証チェックリスト

### 8.1 evaluation データ

- [ ] candidate_id が存在するか
- [ ] timestamp が有効なISO 8601形式か
- [ ] scores.overall が 0-100 の範囲内か
- [ ] scores.acceptance_probability が 0-1 の範囲内か
- [ ] insights.core_motivation が500文字以内か
- [ ] insights.main_concern が500文字以内か

### 8.2 engagement データ

- [ ] candidate_id が存在するか
- [ ] timestamp が有効なISO 8601形式か
- [ ] engagement.date が有効な日付形式（YYYY-MM-DD）か
- [ ] engagement.phase が有効なフェーズ名か
- [ ] engagement.ai_prediction が 0-100 の範囲内か
- [ ] engagement.confidence が 0-1 の範囲内か

### 8.3 acceptance_story データ

- [ ] candidate_id が存在するか
- [ ] timestamp が有効なISO 8601形式か
- [ ] story.recommended_story が存在するか
- [ ] story.key_messages が配列か
- [ ] story.recommended_actions が配列か
- [ ] story.deadline が有効な日付形式（YYYY-MM-DD）か

### 8.4 competitor_comparison データ

- [ ] candidate_id が存在するか
- [ ] timestamp が有効なISO 8601形式か
- [ ] comparison.competitors が配列か
- [ ] competitors[0].company が存在するか
- [ ] comparison.differentiation_points が配列か

---

**以上、データ構造仕様書（Phase 4-2）**
