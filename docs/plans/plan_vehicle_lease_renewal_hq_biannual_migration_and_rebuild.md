# 実装計画書: 半期一括フローへの移行（不要機能削除 + 新機能実装）

- 作成日: 2026-02-11
- 区分: 次期要件基準（新フローの実装正本）
- 対象: `/Users/yushi/work/work/クラハシ`
- 主対象コード: `/Users/yushi/work/work/クラハシ/src/main.ts`
- 関連:
  - 業務定義: `/Users/yushi/work/work/クラハシ/docs/flows/business_flow_vehicle_lease_renewal_hq_biannual_confirmation.md`
  - 現行差分レビュー: `/Users/yushi/work/work/クラハシ/docs/reviews/review_vehicle_lease_renewal_hq_biannual_confirmation_implementation.md`
  - 不要機能棚卸し: `/Users/yushi/work/work/クラハシ/docs/reviews/review_obsolete_features_for_hq_biannual_migration.md`

## 1. 目的 / 成功条件

### 1.1 目的

旧フローの「部署単位依頼 + フォーム回答」を撤去し、  
新フローの「半期一括通知 + 確認用シート回答 + 専務承認 + 村田主任入力後のマスター反映」に移行する。

### 1.2 成功条件（受け入れ条件）

1. 3月便・9月便の固定レンジ抽出が再現できる  
2. 本部長・副本部長への同報通知が動く（部署マスタ非依存）  
3. `回答確認済み` 未完了時のみ、締切10日前に1回だけリマインドが送られる  
4. 専務は確認用シートで承認/差戻しを確定できる  
5. 村田主任入力を満たした行のみマスター反映される  
6. 更新は日付上書き、解約は行グレーアウトされる  
7. 旧フォーム系機能・旧部署マスタ依存の呼び出し導線が削除される  

## 2. 非目的（今回やらないこと）

- リース会社への連絡の自動化
- 既存データの意味を変える大規模データ移行（最小変換に留める）
- UIを別アプリ化する改修（スプレッドシート運用を継続）

## 3. 前提 / 制約

- GAS（V8）で実装し、`dist/` ビルド成果物を `clasp push` する
- スキーマ変更は `syncSchema` で管理し、既存シートを非破壊で移行する
- 本番運用に入るまで、旧導線と新導線の混在期間を最短化する

## 4. 現状（As-Is）

- 現行は部署マスタ依存、3シート統合依存、フォーム依存
- 承認は承認フォーム依存
- 日次トリガー（平日8時）で運用

根拠:
- `/Users/yushi/work/work/クラハシ/src/main.ts:19`
- `/Users/yushi/work/work/クラハシ/src/main.ts:637`
- `/Users/yushi/work/work/クラハシ/src/main.ts:3294`
- `/Users/yushi/work/work/クラハシ/src/main.ts:3362`
- `/Users/yushi/work/work/クラハシ/src/main.ts:2714`

## 5. To-Be（設計方針）

### 5.1 処理単位の変更

- 旧: 管理部門単位 `requestId`
- 新: 半期便単位 `batchId`（例: `2026H1`, `2026H2`）

### 5.2 回答手段の変更

- 旧: Googleフォーム送信
- 新: `◯月期確認用シート` 上の入力

### 5.3 承認手段の変更

- 旧: 承認フォーム送信
- 新: 確認用シートの `専務判断` / `専務コメント` 入力

### 5.4 反映手段の変更

- 旧: `回答` シートの値を書き戻し
- 新: 承認済み + 村田主任入力完了行をマスター反映

## 6. 不要機能の削除計画（網羅）

### 6.1 削除カテゴリA: 部署マスタ依存

- 削除対象:
  - `SHEET_NAMES.DEPT_MASTER`
  - `loadDeptMaster()`
  - `generateDeptTokens()`
  - 部署マスタ前提のメニュー項目
- 置換先:
  - 設定シートの全社固定通知先キー（本部長/副本部長）

### 6.2 削除カテゴリB: 3ソースシート統合

- 削除対象:
  - `SOURCE_SHEETS`
  - 3シートループ処理（同期/テスト/掃除）
- 置換先:
  - `車両一覧` 単独参照

### 6.3 削除カテゴリC: 一次回答フォーム

- 削除対象:
  - `createRequestForms`, `onRequestFormSubmit`
  - 回答復元補助（`extractAnswersFromFormResponse` 系）
  - フォーム再オープン/クローズ関連
- 置換先:
  - 確認用シートの回答列を読む処理

### 6.4 削除カテゴリD: 承認フォーム

- 削除対象:
  - `createOrUpdateApprovalForm`, `onApprovalFormSubmit`
  - 承認フォーム説明/トリガー関連
- 置換先:
  - 専務判断列の読み取りと状態遷移

### 6.5 削除カテゴリE: 廃止済みWeb回答残骸

- 削除対象:
  - `doGet`, `doPost`
  - `validateRequestAccess` など旧Web回答補助
- 置換先:
  - なし（完全撤去）

### 6.6 削除カテゴリF: 日次トリガー

- 削除対象:
  - `installDailyTriggers` の平日8時導線
- 置換先:
  - 半期実行トリガー（3/1・9/1）と手動補助コマンド

## 7. 新機能実装計画（網羅）

### 7.1 新スキーマ

1. 設定キー追加
   - `本部長副本部長_通知先To`
   - `専務_通知先To`
   - `専務_通知先Cc`
   - `半期送付日_3月`
   - `半期送付日_9月`
   - `回答期限_3月`
   - `回答期限_9月`
   - `リマインド_期限前日数`
2. 新シート追加
   - `通知バッチ`
3. 動的確認用シート
   - `本部長副本部長確認_YYYYMMDD`

### 7.2 新フロー関数（新規）

- `createBiannualBatch()`
  - 半期対象抽出と `通知バッチ` 起票
- `buildConfirmationSheet(batchId)`
  - 確認用シート生成と項目初期化
- `sendHqInitialEmail(batchId)`
  - 本部長/副本部長同報通知
- `sendHqReminderIfNeeded(batchId)`
  - 締切10日前・1回限定リマインド
- `sendSenmuApprovalRequestIfReady(batchId)`
  - `回答確認済み` 全件完了後に専務通知
- `applySenmuDecisionFromSheet(batchId)`
  - 専務判断の状態反映
- `applyMasterUpdates(batchId)`
  - 村田主任入力検証後に日付上書き/グレーアウト反映
- `protectSenmuColumns(sheetName)`
  - `専務判断` / `専務コメント` の列保護

### 7.3 既存関数の置換

- `createRequests` -> `createBiannualBatch` に置換
- `sendInitialEmails` -> `sendHqInitialEmail` に置換
- `sendReminderEmails` -> `sendHqReminderIfNeeded` に置換
- `sendApprovalRequestEmails` -> `sendSenmuApprovalRequestIfReady` に置換
- `applyApprovalDecisions` -> `applySenmuDecisionFromSheet` に置換
- `applyAnswers` -> `applyMasterUpdates` と役割分離して置換

## 8. 実装ステップ（削除と新規を統合）

### フェーズ1: 下地（スキーマと設定）

1. 新設定キー追加と `loadSettings` 拡張
2. `通知バッチ` シート追加
3. 確認用シートテンプレ項目定義
4. `document_status_and_timeline` 更新

### フェーズ2: 新フロー追加（旧フロー併存）

1. 半期抽出・確認用シート生成を実装
2. 本部長/副本部長通知を実装
3. `回答確認済み` 判定と10日前リマインドを実装
4. 専務依頼・専務判断反映を実装
5. 村田主任入力ゲートとマスター反映を実装

### フェーズ3: 旧フロー停止

1. メニューから旧導線を退避
2. 旧導線の呼び出し停止（フォーム系・部署マスタ系・日次系）
3. 新導線での運用確認

### フェーズ4: 旧機能削除

1. 旧関数群を削除
2. 旧設定キー/旧シート依存を削除
3. テストを新仕様へ更新

## 9. 影響範囲（変更ファイル）

- 主: `/Users/yushi/work/work/クラハシ/src/main.ts`
- 文書:
  - `/Users/yushi/work/work/クラハシ/docs/flows/business_flow_vehicle_lease_renewal_hq_biannual_confirmation.md`
  - `/Users/yushi/work/work/クラハシ/docs/operations/operation_manual_vehicle_lease_renewal.md`
  - `/Users/yushi/work/work/クラハシ/docs/guides/test_automation_vehicle_lease_renewal.md`
  - `/Users/yushi/work/work/クラハシ/docs/guides/document_status_and_timeline.md`

## 10. リスクと対策

- リスク: 旧導線削除が早すぎて運用停止
  - 対策: フェーズ3で短期併存・検証後に削除
- リスク: 半期抽出漏れ
  - 対策: 3月便/9月便の固定テストデータで自動検証
- リスク: 専務列保護が崩れる
  - 対策: シート生成時と判定実行時の2回で保護再適用
- リスク: 反映誤更新
  - 対策: 村田主任入力必須チェック + `反映済み` 二重防止

## 11. テスト計画

1. 3月便抽出（10月〜3月）検証
2. 9月便抽出（4月〜9月）検証
3. 初回通知の宛先検証（本部長/副本部長同報）
4. リマインド1回制約検証
5. `回答確認済み` 完了時の専務依頼検証
6. 専務承認/差戻し遷移検証
7. 村田主任入力不足時の反映ブロック検証
8. 更新時上書き・解約時グレーアウト検証
9. 旧関数呼び出しが残っていないことの静的検証（`rg`）

## 12. ロールバック

1. 新導線のエントリ呼び出しを停止
2. 旧導線のメニュー/トリガーを一時復帰
3. 旧関数削除前タグへ戻せるよう、フェーズ単位でアトミックコミット

## 13. 未解決 / 追加確認事項

1. `◯月期確認用シート` の命名規則（例: `2026年3月期確認` か `本部長副本部長確認_20260301` か）
2. 専務差戻し時に「全件差戻し」か「行単位差戻し」かの最終運用
3. マスター反映タイミング（都度反映か、バッチ確定反映か）
4. 既存 `更新依頼` / `回答` シートを履歴保管として残すか、完全撤去するか

