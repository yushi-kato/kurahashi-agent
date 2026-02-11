# 棚卸しレビュー: 半期一括フロー移行で不要になる機能

- 作成日: 2026-02-11
- 区分: 差分監査（次期要件への移行時に削除/置換すべき対象の棚卸し）
- 対象コード: `/Users/yushi/work/work/クラハシ/src/main.ts`
- 関連要件: `/Users/yushi/work/work/クラハシ/docs/flows/business_flow_vehicle_lease_renewal_hq_biannual_confirmation.md`

## 1. この文書の目的

業務フローが「部署単位通知」から「半期一括通知」へ変わると、  
機能の意味そのものが変わる箇所と、完全に不要になる箇所が混在する。

この文書では、次の3分類で整理する。

- 廃止（削除）:
  - 新業務では意味を持たず、残すと誤動作や運用混在を生むもの
- 置換（作り替え）:
  - 役割は必要だが、処理単位や判定基準を差し替えるもの
- 継続（流用）:
  - 新業務でもそのまま有効な基盤

## 2. 廃止（削除）対象

### 2.1 部署マスタ依存の通知ルート

- 背景の意味:
  - 旧フローは「管理部門ごとに通知先が異なる」前提だった。
  - 新フローは「本部長・副本部長へ全社一括通知」なので、部署別通知先管理が不要。
- 対象:
  - `SHEET_NAMES.DEPT_MASTER` 定義
  - `loadDeptMaster()`
  - `generateDeptTokens()`
  - `seedSettings` での部署マスタ前提補助
- 根拠:
  - `/Users/yushi/work/work/クラハシ/src/main.ts:8`
  - `/Users/yushi/work/work/クラハシ/src/main.ts:3036`
  - `/Users/yushi/work/work/クラハシ/src/main.ts:2675`

### 2.2 3ソースシート統合前提

- 背景の意味:
  - 旧フローは `車両一覧` / `車両一覧【ｹﾝｽｲ】` / `車両一覧【ﾈｸｽﾄ】` の統合が中心だった。
  - 新フローは台帳を `車両一覧` 一本に寄せるため、3元統合の前提が不要。
- 対象:
  - `SOURCE_SHEETS`
  - 3シートループ処理（`syncVehicles`, `applyAnswers`, `seedTestVehicles`, `diagnoseSourceSheets`, `cleanupTestData`, `runTestSuite` の各ループ）
- 根拠:
  - `/Users/yushi/work/work/クラハシ/src/main.ts:19`
  - `/Users/yushi/work/work/クラハシ/src/main.ts:423`
  - `/Users/yushi/work/work/クラハシ/src/main.ts:1243`

### 2.3 一次回答フォーム（作成・送信・回収）

- 背景の意味:
  - 旧フローの「回答」は Googleフォーム送信そのもの。
  - 新フローの「回答」は確認用シートでの車両別入力と `回答確認済み` 管理。
- 対象:
  - `createRequestForms`
  - `onRequestFormSubmit`
  - `extractAnswersFromFormResponse`
  - `extractAnswersFromGridItem`
  - `loadVehicleIdsForForm` などフォーム回答復元系
  - `closeRequestForms`, `reopenRequestForms`, `ensureFormSubmitTrigger`
- 根拠:
  - `/Users/yushi/work/work/クラハシ/src/main.ts:3294`
  - `/Users/yushi/work/work/クラハシ/src/main.ts:1060`
  - `/Users/yushi/work/work/クラハシ/src/main.ts:3866`

### 2.4 承認フォーム（作成・送信・回収）

- 背景の意味:
  - 旧フローの「承認」は承認フォーム入力。
  - 新フローの「承認」は確認用シートの承認列入力。
- 対象:
  - `createOrUpdateApprovalForm`
  - `onApprovalFormSubmit`
  - `extractApprovalDecisionFromFormResponse`
  - `ensureApprovalFormSubmitTrigger`
  - `closeApprovalFormByRequestRow`
- 根拠:
  - `/Users/yushi/work/work/クラハシ/src/main.ts:3362`
  - `/Users/yushi/work/work/クラハシ/src/main.ts:1111`

### 2.5 Web回答（廃止済み残骸）

- 背景の意味:
  - すでに Web回答は廃止済みで、フォームへ誘導するメッセージのみ。
  - 新フローではフォーム自体も外すため、関連コードは完全不要。
- 対象:
  - `doGet`, `doPost`
  - `validateRequestAccess`
  - `findRequestRow`, `getVehiclesByRequestId`, `loadAnswersForRequest`, `buildAnswerRowHtml`
- 根拠:
  - `/Users/yushi/work/work/クラハシ/src/main.ts:1050`
  - `/Users/yushi/work/work/クラハシ/src/main.ts:3964`

### 2.6 日次トリガー運用

- 背景の意味:
  - 新フローは 3月1日・9月1日を基準とした半期バッチ運用。
  - 平日毎朝の定時バッチは業務意図と不一致。
- 対象:
  - `installDailyTriggers`（平日8時固定）
- 根拠:
  - `/Users/yushi/work/work/クラハシ/src/main.ts:2714`

## 3. 置換（作り替え）対象

### 3.1 依頼作成

- 旧:
  - `createRequests` が「本日〜Nか月」かつ「部署単位」で依頼作成。
- 新:
  - 半期便（3月便/9月便）単位で対象抽出し、確認用シートを生成する方式へ置換。
- 根拠:
  - `/Users/yushi/work/work/クラハシ/src/main.ts:538`

### 3.2 初回通知

- 旧:
  - `sendInitialEmails` が部署ごとに送信。
- 新:
  - 本部長・副本部長へ同報（To）で半期便の確認依頼を送信。
- 根拠:
  - `/Users/yushi/work/work/クラハシ/src/main.ts:637`

### 3.3 リマインド通知

- 旧:
  - `リマインド_初回から日数` / `間隔日数` / `最大回数` の多段リマインド。
- 新:
  - 締切10日前に1回だけ、`回答確認済み` 未完了分がある場合に通知。
- 根拠:
  - `/Users/yushi/work/work/クラハシ/src/main.ts:792`

### 3.4 承認依頼・承認反映

- 旧:
  - 承認フォームURLをメール送信し、承認入力を取り込む。
- 新:
  - 確認用シートの `専務判断` / `専務コメント` を読み取って状態遷移させる。
- 根拠:
  - `/Users/yushi/work/work/クラハシ/src/main.ts:1558`
  - `/Users/yushi/work/work/クラハシ/src/main.ts:1697`

### 3.5 回答反映

- 旧:
  - `applyAnswers` が `回答` シート起点で元台帳へ書き戻す。
- 新:
  - 確認用シート上の一次回答・承認結果・村田主任入力を起点に、マスター反映へ置換。
- 根拠:
  - `/Users/yushi/work/work/クラハシ/src/main.ts:1192`

## 4. 継続（流用）対象

- `syncSchema` / `ensureHeaders`:
  - 新シート・新列追加の安全基盤として継続
- `appendNotificationLog`:
  - 通知実績監査として継続
- 回答ラベル定義:
  - `更新 / 解約（入替） / 解約（満了）` は新要件と整合
- ロック制御:
  - `LockService.getDocumentLock()` は並行実行保護として継続

## 5. 削除の進め方（推奨）

1. 先に新機能を「並行導線」で実装し、切替フラグで新旧共存期間を短期運用する  
2. 新導線の実運用確認後、旧導線呼び出しを止める  
3. 旧導線関数と旧設定キーを段階削除する  
4. 不要シートと不要メニューを最後に整理する  

## 6. この棚卸しの位置づけ

この文書は「削除対象の合意ベース」です。  
具体的な手順・順序・テスト項目は、実装計画書（`docs/plans/`）を正本として実施する。

