# E2Eテスト手順書: 車両リース更新通知（半期一括フロー）

- 作成日: 2026-02-12
- 区分: 実装基準（現行コードのE2E確認手順）
- 対象コード: `/Users/yushi/work/work/クラハシ/src/main.ts`
- 目的: 「理想の業務フロー」が実際に動作しているかを、モックデータで一連確認する

## 0. この手順書で確認すること

この手順書は、次を1本で確認します。

1. 半期バッチ起票（抽出レンジと件数）
2. 本部長・副本部長向け確認シート生成
3. 初回通知
4. 期限前リマインド（1回制約）
5. 専務依頼（全件確認後のみ）
6. 専務判断反映
7. マスター反映（更新は日付上書き、解約はグレーアウト）
8. 旧導線非到達（受け入れ条件7）

---

## 1. 用語（意味から）

- `E2Eテスト`
  - ここでは「設定・元台帳入力から、最終反映まで業務全体を通す確認」を指します。
- `モックデータ`
  - 実データの代わりに、意図した条件を再現するためのテスト用データです。
- `dry-run`
  - メールを実送信しない確認モード（`通知_メール送信=FALSE`）です。
- `full-run`
  - テスト用アドレスに実送信し、通知時刻や送信状態まで含めて確認するモードです。

---

## 2. 事前準備

### 2.1 作業対象

- 本番スプレッドシートを直接使わず、複製したテスト用シートを使う
- `.clasp.json` の接続先がテスト用スクリプトであることを確認する

### 2.2 ローカル反映

```bash
cd /Users/yushi/work/work/クラハシ
npm run build
npx clasp show-file-status
npx clasp push -f
```

### 2.3 `clasp run` の実行プロファイル

Spreadsheet権限エラー回避のため、実行は `runscope` を推奨します。

```bash
npx clasp -u runscope run ping
```

---

## 3. 初期化（毎回）

### 3.1 スキーマ同期

```bash
npx clasp -u runscope run syncSchema
```

### 3.2 旧テストデータ掃除

モック行は `TEST` 接頭辞で投入するため、掃除関数で削除できます。

```bash
npx clasp -u runscope run cleanupTestData
```

---

## 4. モックデータ準備

## 4.1 `設定` シート（必須キー）

`設定` タブで次を設定します。

- `通知_メール送信`
  - dry-run: `FALSE`
  - full-run: `TRUE`
- `本部長副本部長_通知先To`
  - full-run時はテスト用メアド（複数はカンマ区切り）
- `専務_通知先To`
  - full-run時はテスト用メアド
- `専務_通知先Cc`
  - 任意
- `半期送付日_3月`, `半期送付日_9月`
  - まずはデフォルトのままで可
- `回答期限_3月`, `回答期限_9月`
  - リマインドを即テストしたい場合は `今日` 相当の日付にする
  - 例: `2026-02-12`（実行日の絶対日付）
- `リマインド_期限前日数`
  - `10`

### 4.2 `車両一覧` シート（モック投入）

ヘッダは最低限、次を含めます。

- `登録番号`
- `車種`
- `車台番号`
- `契約開始日`
- `契約満了日`
- `管理部門`
- `管理担当者`
- `契約期間`
- `車検満了日`
- `リース料（税抜）`

次のデータをそのまま投入します（`TEST` 接頭辞を維持）。

| 登録番号 | 車種 | 車台番号 | 契約開始日 | 契約満了日 | 管理部門 | 管理担当者 | 契約期間 | 車検満了日 | リース料（税抜） |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| TEST-1001 | TEST_更新対象 | TEST-CH-001 | 2025-04-01 | `=TODAY()` | 本社総務 | テスト太郎 | 60ヶ月 | 2027-04-01 | 50000 |
| TEST-1002 | TEST_解約対象 | TEST-CH-002 | 2025-05-01 | `=TODAY()+1` | 本社総務 | テスト花子 | 60ヶ月 | 2027-05-01 | 52000 |
| TEST-9001 | TEST_対象外 | TEST-CH-003 | 2025-01-01 | `=TODAY()+200` | 本社総務 | テスト次郎 | 60ヶ月 | 2027-01-01 | 53000 |
| TEST-ERR1 | TEST_満了日欠損 | TEST-CH-004 | 2025-06-01 |  | 本社総務 | テスト欠損 | 60ヶ月 | 2027-06-01 | 54000 |

意図:

- 1行目: 更新反映を確認
- 2行目: 解約反映（グレーアウト）を確認
- 3行目: 半期抽出対象外を確認
- 4行目: `要入力` 検出を確認

---

## 5. 実行手順（E2E本線）

### 5.1 元台帳同期と抽出準備

```bash
npx clasp -u runscope run syncVehicles
npx clasp -u runscope run createBiannualBatch
npx clasp -u runscope run buildConfirmationSheetForLatestBatch
```

確認ポイント:

- `通知バッチ` に1行追加される
- `対象件数` が 2 以上（更新対象 + 解約対象）
- `確認用シート名` が設定される
- `要入力` に `契約満了日なし` が1件以上出る

### 5.2 初回通知

```bash
npx clasp -u runscope run sendHqInitialEmail
```

確認ポイント:

- full-run: `通知バッチ.初回通知送信日時` が入る
- dry-run: `通知ログ` に「通知_メール送信=FALSE のため送信をスキップ」が残る

### 5.3 リマインド（1回制約）

事前に確認用シートで以下を入力:

- 1行目: `本部回答=更新`, `回答確認済み=TRUE`
- 2行目: `本部回答=解約（満了）`, `回答確認済み=FALSE`（未確認を残す）

実行:

```bash
npx clasp -u runscope run sendHqReminderIfNeeded
npx clasp -u runscope run sendHqReminderIfNeeded
```

確認ポイント:

- 1回目で条件一致時のみ送信（またはスキップログ）
- 2回目で重複送信されない（`リマインド送信日時` が増えない）

### 5.4 専務依頼のゲート確認

まず未確認が残る状態で実行:

```bash
npx clasp -u runscope run sendSenmuApprovalRequestIfReady
```

期待:

- `通知ログ` に「未確認行あり」の未送信理由

次に確認用シートで 2行目の `回答確認済み=TRUE` に変更し、再実行:

```bash
npx clasp -u runscope run sendSenmuApprovalRequestIfReady
```

期待:

- full-run: `専務依頼送信日時` が入る
- dry-run: `通知_メール送信=FALSE` のスキップ理由が残る

### 5.5 専務判断反映

確認用シートで:

- 1行目 `専務判断=承認`
- 2行目 `専務判断=承認`（差戻しテスト時は `差戻し`）

実行:

```bash
npx clasp -u runscope run applySenmuDecisionFromSheet
```

確認ポイント:

- `通知バッチ.ステータス` が `専務承認済` または `専務差戻し`
- `通知ログ` に 承認/差戻し/保留/不正 の集計が残る

### 5.6 マスター反映

確認用シートで:

- 更新行: `新契約開始日` と `新契約満了日` を入力
- 解約行: `解約完了=TRUE`

実行:

```bash
npx clasp -u runscope run applyMasterUpdates
```

確認ポイント:

- 更新行: `車両（統合ビュー）` の `契約開始日`/`契約満了日` が上書きされる
- 解約行: `車両（統合ビュー）` 該当行がグレーアウトされる
- 確認用シート: `マスター反映済み=TRUE`, `反映日時` が入る

---

## 6. 受け入れ条件チェックリスト（理想フロー判定）

1. 3月便/9月便の固定レンジ抽出が再現できる
- `通知バッチ.対象開始日/対象終了日` が半期レンジ

2. 本部長・副本部長への同報通知
- full-runで `初回通知送信日時` 記録

3. 未完了時のみ締切前1回リマインド
- `sendHqReminderIfNeeded` 2回実行で重複しない

4. 専務承認/差戻し確定
- `applySenmuDecisionFromSheet` 後のステータス遷移

5. 村田主任入力済み行のみ反映
- 更新: 新契約日2項目必須
- 解約: `解約完了` 必須

6. 更新は日付上書き、解約はグレーアウト
- `車両（統合ビュー）` で視覚確認

7. 旧導線非到達

```bash
npx clasp -u runscope run verifyAcceptanceCondition7LegacyNonReachable
```

期待:

- `{ ok: true, remainingLegacyEntries: [], wrapperErrors: [] }`

---

## 7. 証跡回収（テスト報告に添付するもの）

### 7.1 `runTestSuite` ベースで回す場合

```bash
npx clasp -u runscope run runTestSuite
npx clasp -u runscope run exportTestResults --params "[50]"
```

補足:

- `runTestSuite` は `No response.` でも、`exportTestResults` 側で結果が取れれば成功扱い

### 7.2 最低限残す証跡

- `通知バッチ` の対象行（対象期間・件数・各送信時刻・ステータス）
- `通知ログ`（初回/リマインド/専務依頼/反映）
- 確認用シート（判断列と反映列）
- `車両（統合ビュー）` の更新/解約反映後状態
- `verifyAcceptanceCondition7LegacyNonReachable` の実行結果

---

## 8. 片付け（テスト後）

```bash
npx clasp -u runscope run cleanupTestData
```

注意:

- 本手順は `TEST` 接頭辞のモック行だけを掃除対象にしています
- テスト用スプレッドシートごと破棄する運用でも構いません
