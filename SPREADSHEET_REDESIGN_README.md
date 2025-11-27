# 📊 スプレッドシート再設計 完全実装

**作成日**: 2025年11月27日
**対象**: 【MVP_v1】候補者管理シート
**目的**: 企業販売可能なMVP_v1への最適化

## 🎯 実装の全体像

### Before（現状）
- **シート数**: 23シート
- **Candidates_Master**: 57列（肥大化）
- **重複シート**: 4シート存在
- **Dify連携**: 未対応

### After（目標）
- **シート数**: 21シート（実質17シート）
- **Candidates_Master**: 15列（最適化）
- **重複シート**: 削除完了
- **Dify連携**: 完全対応

## 📋 実装の段階

この実装は4つのステップに分かれています：

1. **Step 1**: 新規シート作成（3シート）
2. **Step 2**: データ移行（Candidates_Master → 3シート）
3. **Step 3**: 既存シート拡張（Dify連携用の列追加）
4. **Step 4**: 不要シート削除（4シート）

⚠️ **重要**: 必ずこの順序で実行してください

## 🚀 実行方法

### 方法1: 全ステップを一度に実行（推奨）

```javascript
// Step 1 & Step 2を実行
executeAllSteps();

// ログを確認してデータが正しく移行されているか確認

// Step 3 & Step 4を実行
executeStep3AndStep4();
```

### 方法2: 各ステップを個別に実行

```javascript
// Step 1: 新規シート作成
executeStep1();

// Step 2: データ移行
executeStep2();

// Step 3: 既存シート拡張
executeStep3();

// Step 4: 不要シート削除
executeStep4();
```

## 📊 実装内容の詳細

### Step 1: 新規シート作成（3シート）

#### 1. Candidate_Scores
- **目的**: Candidates_Masterからスコア関連の列を分離
- **列数**: 20列
- **主要列**:
  - candidate_id
  - 最新_合格可能性
  - 最新_承諾可能性（AI予測）
  - 最新_承諾可能性（人間の直感）
  - 最新_承諾可能性（統合）
  - 各種スコア（Philosophy、Strategy、Motivation、Execution）

#### 2. Candidate_Insights
- **目的**: AIインサイト（モチベーション、懸念、競合）を管理
- **列数**: 10列
- **主要列**:
  - candidate_id
  - コアモチベーション
  - 主要懸念事項
  - 懸念カテゴリ
  - 競合企業1〜3
  - 次推奨アクション

#### 3. Dify_Workflow_Log
- **目的**: Difyワークフローの実行履歴を記録
- **列数**: 9列
- **状態**: 非表示シート（デバッグ用）
- **主要列**:
  - workflow_log_id
  - workflow_name
  - candidate_id
  - status
  - error_message

### Step 2: データ移行

#### Candidates_Masterの再構成
- **57列 → 15列**に削減
- スコア関連の列を`Candidate_Scores`に移行
- インサイト関連の列を`Candidate_Insights`に移行
- 使用されていない列を削除

#### 新しいCandidates_Master（15列）
1. candidate_id
2. 氏名
3. 現在ステータス
4. 最終更新日時
5. 採用区分
6. 担当面接官
7. 応募日
8. メールアドレス
9. 初回面談日
10. 1次面接日
11. 2次面接日
12. 最終面接日
13. 最新_合格可能性（参照：VLOOKUP）
14. 最新_承諾可能性（参照：VLOOKUP）
15. 総合ランク（新規：A〜E）

### Step 3: 既存シート拡張（Dify連携用）

#### 1. Evaluation_Log
追加列（5列）:
- 文字起こし本文
- バッティング企業
- Googleドキュメント評価レポートURL
- 積極性スコア
- dify_workflow_id

#### 2. Acceptance_Story
追加列（7列）:
- AI信頼度
- Phase3_acceptance_rate
- AI_acceptance_rate
- 競合状況分析
- リスク要因
- 機会要因
- dify_workflow_id

#### 3. NextQ
追加列（3列）:
- 使用ステータス
- 使用日時
- dify_workflow_id

#### 4. Survey_Send_Log
追加列（1列）:
- dify_workflow_id

#### 5. Contact_History
追加列（1列）:
- contact_source（Dify/手動/自動）

### Step 4: 不要シート削除

以下の4シートを削除：
1. Survey_Analysis
2. Evidence
3. Risk
4. Survey_Response

## ✅ 検証方法

### データ整合性の確認

```javascript
verifyDataIntegrity();
```

このスクリプトは以下を確認します：
- Candidates_Masterのデータ件数
- Candidate_Scoresのデータ件数
- Candidate_Insightsのデータ件数
- 3つのシートでデータ件数が一致しているか

### 数式の確認

```javascript
verifyFormulas();
```

このスクリプトは以下を確認します：
- M列（最新_合格可能性）にVLOOKUP数式が設定されているか
- N列（最新_承諾可能性）にVLOOKUP数式が設定されているか

### 最終確認

```javascript
finalVerification();
```

このスクリプトは以下を確認します：
- データ整合性
- 数式の設定
- 全シート一覧の表示

## 📝 完了チェックリスト

### ✅ Step 1: 新規シート作成
- [ ] Candidate_Scoresシートが作成されている
- [ ] Candidate_Insightsシートが作成されている
- [ ] Dify_Workflow_Logシートが作成されている（非表示）
- [ ] 各シートにヘッダーが設定されている
- [ ] タブの色が設定されている

### ✅ Step 2: データ移行
- [ ] Candidate_Scoresに全候補者のデータが移行されている
- [ ] Candidate_Insightsに全候補者のデータが移行されている
- [ ] Candidates_Masterが15列に削減されている
- [ ] Candidates_MasterのM列・N列が数式になっている
- [ ] データの整合性が保たれている

### ✅ Step 3: 既存シート拡張
- [ ] Evaluation_Logに5列追加されている
- [ ] Acceptance_Storyに7列追加されている
- [ ] NextQに3列追加されている
- [ ] Survey_Send_Logに1列追加されている
- [ ] Contact_Historyに1列追加されている
- [ ] 各列のヘッダーが正しく設定されている

### ✅ Step 4: 不要シート削除
- [ ] Survey_Analysisシートが削除されている
- [ ] Evidenceシートが削除されている
- [ ] Riskシートが削除されている
- [ ] Survey_Responseシートが削除されている
- [ ] シート数が正しい（23 → 21シート）

## 🔧 トラブルシューティング

### エラー1: 「列が見つかりません」
**原因**: Candidates_Masterの列名が想定と異なる

**解決策**:
1. Candidates_Masterの1行目を確認
2. `getColumnIndex()`関数内の列名を実際の列名に合わせる

### エラー2: 「シートが見つかりません」
**原因**: 必要なシートが存在しない

**解決策**:
1. スプレッドシートに必要なシートが全て存在するか確認
2. シート名のスペルミスがないか確認

### エラー3: データが移行されない
**原因**: データの形式が想定と異なる

**解決策**:
1. `migrateDataFromCandidatesMaster()`内のログを確認
2. 各列のインデックスが正しいか確認
3. データが空でないか確認

## 🔄 ロールバック手順

万が一、問題が発生した場合：

```javascript
rollbackToBackup();
```

⚠️ **注意**: ロールバックは手動で行ってください
1. スプレッドシートのバックアップから復元
2. または、Google Driveの「バージョン履歴」から復元

## 📞 サポート

問題が発生した場合：
1. ログを確認してください（View → Logs）
2. エラーメッセージをコピーして共有してください
3. 必要に応じてロールバック手順を実行してください

## 🎯 実装後の効果

### データの見通しが良くなる
- Candidates_Masterが15列にスリム化
- スコアとインサイトが別シートで管理される
- 関連データが整理される

### Dify連携が可能になる
- 各シートに`dify_workflow_id`列が追加される
- ワークフローの実行履歴が記録される
- AIによる自動化が実現する

### 保守性が向上する
- 重複シートが削除される
- データ構造がシンプルになる
- 将来の拡張が容易になる
