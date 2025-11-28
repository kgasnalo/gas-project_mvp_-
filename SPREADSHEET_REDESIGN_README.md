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

### 🚨 実行前の必須確認事項

1. **スプレッドシート全体のコピーを作成**
   - Google Drive上で右クリック→「コピーを作成」
   - **本番データで直接実行しないこと**

2. **Apps Scriptの実行権限を確認**
   - スプレッドシートの編集権限があること
   - スクリプトの実行権限があること

3. **実行ガイドを確認**
   ```javascript
   executionGuide();
   ```

### 方法1: Phase別実行（推奨・最も安全）

```javascript
// Phase 0: 事前準備（必須）
phase0_preparation();
// → バックアップ作成と列の存在確認
// → ログを確認してください

// Phase 1: Step 1-2実行
phase1_execute();
// → 新規シート作成 + データ移行
// → ログでデータを確認してください

// Phase 2: Step 3-4実行
phase2_execute();
// → 既存シート拡張 + 不要シート削除

// 最終確認
finalVerification();
```

### 方法2: 全ステップを一度に実行（非推奨）

```javascript
// ⚠️ 警告: Phase 0を実行せずに進めるのは危険です
// Step 1 & Step 2を実行
executeAllSteps();

// ログを確認してデータが正しく移行されているか確認

// Step 3 & Step 4を実行
executeStep3AndStep4();
```

### 方法3: 各ステップを個別に実行（上級者向け）

```javascript
// ⚠️ Phase 0を必ず最初に実行してください
phase0_preparation();

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

## 🛡️ 安全対策

### 実装済みの安全機能

#### 1. 自動バックアップ機能
- `phase0_preparation()`実行時に自動でバックアップシートを作成
- バックアップ名: `Candidates_Master_BACKUP_yyyyMMdd_HHmmss`
- 非表示シートとして保存

#### 2. エラーハンドリング強化
- 列が見つからない場合は即座にエラーを投げる
- 利用可能な列名を表示してデバッグを容易に
- データが空の場合は処理を中断

#### 3. シングルクォートエスケープ
- QUERY関数内の文字列を自動エスケープ
- SQL構文エラーを防止

#### 4. データ検証
- 新しいデータが空でないことを確認
- 書き込み前にデータ件数をログ出力

### 🚨 万が一エラーが発生した場合

#### データ復旧手順

```javascript
// 1. バックアップシートを確認
// スプレッドシートから「Candidates_Master_BACKUP_*」を探す

// 2. バックアップシートをコピー
// バックアップシートを右クリック→「コピーを作成」

// 3. コピーしたシートの名前を変更
// 「Candidates_Master」に名前を変更

// 4. 元のシートを削除
// 破損したCandidates_Masterを削除
```

または

```javascript
// Google Driveのバージョン履歴から復元
// ファイル→バージョン履歴→以前のバージョンを復元
```

## 🔧 トラブルシューティング

### エラー1: 「列が見つかりません: XXX」
**原因**: Candidates_Masterに指定された列が存在しない

**解決策**:
1. エラーメッセージに表示される「利用可能な列」を確認
2. `phase0_preparation()`を実行して必須列が存在するか確認
3. 列名が完全一致しているか確認（スペース、全角/半角に注意）

**エラーメッセージ例**:
```
列が見つかりません: 最新_合格可能性
利用可能な列: candidate_id, 氏名, 現在ステータス, ...
```

### エラー2: 「シートが見つかりません」
**原因**: 必要なシートが存在しない

**解決策**:
1. スプレッドシートに必要なシートが全て存在するか確認
2. シート名のスペルミスがないか確認
3. `phase0_preparation()`を実行してシート一覧を確認

### エラー3: データが移行されない
**原因**: データの形式が想定と異なる

**解決策**:
1. `migrateDataFromCandidatesMaster()`内のログを確認
2. 各列のインデックスが正しいか確認
3. データが空でないか確認
4. `phase0_preparation()`で必須列が全て存在することを確認

### エラー4: 「新しいデータが空です」
**原因**: データ移行処理で候補者データが取得できなかった

**解決策**:
1. Candidates_Masterに候補者データが存在するか確認
2. `candidate_id`列にデータが入っているか確認
3. 2行目以降にデータがあるか確認（1行目はヘッダー）

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
