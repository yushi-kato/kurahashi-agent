# 実装レビュー結果: 半期一括確認フローの実現性チェック

- 実施日: 2026-02-11
- 対象実装: `/Users/yushi/work/work/クラハシ/src/main.ts`
- 比較対象: `/Users/yushi/work/work/クラハシ/docs/flows/business_flow_vehicle_lease_renewal_hq_biannual_confirmation.md`
- レビュー方式: 静的コードレビュー（実行テストは未実施）

## 1. 結論サマリ

現行実装は、**部署マスタ + 部署単位依頼 + フォーム回答**を前提にした運用であり、  
今回定義した「本部長/副本部長への半期一括通知 + 確認用シート回答 + 村田主任の反映ゲート」は、**主要部分が未実装**。

実現度の目安:

- 実現済み: 2件
- 部分実現: 1件
- 未実現: 10件

## 2. 重要指摘（高優先度）

### Critical

1. 半期固定（3月便/9月便）での抽出・送付が未実装
   - 現状は `runDaily` を平日8時に実行する日次方式。
   - 根拠: `installDailyTriggers` は平日トリガーのみを作成。`/Users/yushi/work/work/クラハシ/src/main.ts:2714`
   - 根拠: `createRequests` は `本日` から `抽出_満了まで月数` で範囲抽出。`/Users/yushi/work/work/クラハシ/src/main.ts:562`

2. 通知先が本部長/副本部長固定ではなく、部署マスタ依存
   - `sendInitialEmails` と `sendReminderEmails` は `管理部門` ごとに `部署マスタ.通知先To/Cc` を参照。
   - 根拠: `/Users/yushi/work/work/クラハシ/src/main.ts:647`
   - 根拠: `/Users/yushi/work/work/クラハシ/src/main.ts:804`

3. 「確認用シートで回答」ではなく「Googleフォーム回答」
   - 一次回答は `createRequestForms` でフォームを作成し、`onRequestFormSubmit` で取り込み。
   - 根拠: `/Users/yushi/work/work/クラハシ/src/main.ts:3294`
   - 根拠: `/Users/yushi/work/work/クラハシ/src/main.ts:1060`

4. 「回答確認済みチェックボックス」および10日前1回リマインドが未実装
   - `回答確認済み` 列はスキーマに存在しない。
   - 根拠: `SCHEMA_DEFS` に該当ヘッダなし。`/Users/yushi/work/work/クラハシ/src/main.ts:65`
   - 根拠: リマインドは `初回から日数` / `間隔日数` / `最大回数` の繰り返し方式。
   - 根拠: `/Users/yushi/work/work/クラハシ/src/main.ts:878`

5. 専務承認は確認用シート入力ではなく承認フォーム方式
   - `sendApprovalRequestEmails` は承認フォームURLをメール送信。
   - 根拠: `/Users/yushi/work/work/クラハシ/src/main.ts:1619`
   - 根拠: `/Users/yushi/work/work/クラハシ/src/main.ts:1655`

6. 村田主任の必須入力ゲートとマスター反映（上書き/グレーアウト）が未実装
   - 承認後は「通知メール送信」までで、`新契約開始日/満了日` や `解約完了` の入力検証および反映処理は存在しない。
   - 根拠: `/Users/yushi/work/work/クラハシ/src/main.ts:1806`
   - 根拠: `applyAnswers` が行う反映は回答値の書き戻しのみ。`/Users/yushi/work/work/クラハシ/src/main.ts:1243`

### Important

1. 台帳一本化（車両一覧のみ）未実施
   - 3シート統合 `SOURCE_SHEETS` を前提に全処理が動く。
   - 根拠: `/Users/yushi/work/work/クラハシ/src/main.ts:19`
   - 根拠: `/Users/yushi/work/work/クラハシ/src/main.ts:423`

2. 承認フローが既定で無効
   - `承認フロー_有効` の既定値が `false`。設定ONしないと承認処理は動かない。
   - 根拠: `/Users/yushi/work/work/クラハシ/src/main.ts:217`
   - 根拠: `/Users/yushi/work/work/クラハシ/src/main.ts:1703`

## 3. 要件別チェック表

| ID | 要件 | 判定 | 根拠 |
|---|---|---|---|
| R1 | 本部長・副本部長へ集約通知（部署マスタ不要） | 未実現 | `loadDeptMaster` 依存で部署単位送信。`/Users/yushi/work/work/クラハシ/src/main.ts:3036`, `/Users/yushi/work/work/クラハシ/src/main.ts:647` |
| R2 | 半年に1回（3/1・9/1）送付 | 未実現 | 平日8時の日次トリガー。`/Users/yushi/work/work/クラハシ/src/main.ts:2714` |
| R3 | 3月便: 10月〜3月満了、9月便: 4月〜9月満了 | 未実現 | `本日〜Nか月` 抽出のみ。`/Users/yushi/work/work/クラハシ/src/main.ts:562` |
| R4 | `◯月期確認用シート` を生成して回答 | 未実現 | 回答はフォーム生成。`/Users/yushi/work/work/クラハシ/src/main.ts:3294` |
| R5 | 回答選択肢は ①更新 ②解約（入替）③解約（満了） | 実現済み | 選択肢定義が一致。`/Users/yushi/work/work/クラハシ/src/main.ts:46` |
| R6 | `回答確認済み` チェックボックス列 | 未実現 | スキーマに列なし。`/Users/yushi/work/work/クラハシ/src/main.ts:65` |
| R7 | 10日前に1回だけリマインド | 未実現 | 日数間隔 + 最大回数方式。`/Users/yushi/work/work/クラハシ/src/main.ts:878` |
| R8 | 全件確認後に専務へ承認依頼メール | 部分実現 | 「全回答後の専務依頼」はあるが基準が `回答確認済み` ではなくフォーム回答完了。`/Users/yushi/work/work/クラハシ/src/main.ts:1465`, `/Users/yushi/work/work/クラハシ/src/main.ts:1558` |
| R9 | 専務は承認/差戻しを判断 | 実現済み | 承認/差戻しの状態遷移あり。`/Users/yushi/work/work/クラハシ/src/main.ts:1784`, `/Users/yushi/work/work/クラハシ/src/main.ts:1854` |
| R10 | 村田主任が更新日付/解約完了を入力後、反映 | 未実現 | 入力項目・検証・反映処理なし。`/Users/yushi/work/work/クラハシ/src/main.ts:1806` |
| R11 | マスター反映: 更新は日付上書き、解約はグレーアウト | 未実現 | 回答の書き戻しのみで、日付上書き/行着色なし。`/Users/yushi/work/work/クラハシ/src/main.ts:1243` |
| R12 | 台帳は `車両一覧` 単独で運用 | 未実現 | 3シート前提。`/Users/yushi/work/work/クラハシ/src/main.ts:19` |
| R13 | 部署マスタ管理を廃止 | 未実現 | スキーマ・送信・テストが部署マスタ前提。`/Users/yushi/work/work/クラハシ/src/main.ts:72`, `/Users/yushi/work/work/クラハシ/src/main.ts:252`, `/Users/yushi/work/work/クラハシ/src/main.ts:2675` |

## 4. 既存実装の活用可能な要素

- 回答ラベル（更新/解約）定義は新要件と一致している
- 承認ステータス遷移（承認待ち/承認済/差戻し）の状態管理は流用可能
- 通知ログ基盤（`通知ログ`）はそのまま利用可能
- シートスキーマ同期基盤（`syncSchema`）は新タブ/新列追加の土台として有効

## 5. 総合判定

新業務フローに対する現行実装の適合度は、**低い（要件の中核が未実装）**。  
ただし、通知ログ・状態遷移・回答ラベルなど、土台として再利用できる部品は残っているため、全面作り直しではなく「通知単位・回答手段・反映手段」を入れ替える改修が現実的。

## 6. 次の実装着手優先度（提案）

1. 通知/抽出の単位を「部署」から「半期バッチ」へ変更
2. 確認用シート生成 + `回答確認済み` 判定 + 10日前1回リマインド実装
3. 専務承認をフォームから確認用シート入力へ切替
4. 村田主任入力ゲート + マスター反映（上書き/グレーアウト）実装
5. 部署マスタ依存の段階的撤去と `車両一覧` 一本化
