
# AI向け：Google Apps Script（GAS）安定開発フロー & 落とし穴対策ノウハウ集

## 0. 目的

あなた（AI）は、GASの仕様・制約を踏まえた**壊れにくい実装**を、**小さな差分**で反復しながら提供する。
「動いたっぽい」ではなく、**クォータ・実行時間・認可・トリガー・同時実行**を含めて破綻しないことをゴールにする。

---

## 1. 前提（これを破る提案はNG）

### 1.1 ランタイム/実行制限（最初に織り込む）

* **スクリプト実行時間は 1回あたり最大6分**（アカウント種別に関わらず “/execution” の制限として提示されている）ので、6分を超える可能性がある処理は**分割実行**が前提。([Google for Developers][1])
* **カスタム関数**（シートの `=MYFUNC()` 形式）は **30秒/回**。([Google for Developers][1])
* **シンプルトリガー（onEdit/onOpen等）**は **30秒を超えられない**前提で設計する。([Google for Developers][2])
* 1日クォータ（例：URL Fetch回数、Properties read/writeなど）と、トリガー合計実行時間（/day）も存在する。([Google for Developers][1])

### 1.2 V8ランタイムは “Nodeでもブラウザでもない”

GASのV8はモダンJSを使えるが、**標準APIが欠けている**。これを前提にコードを出す。([Google for Developers][3])

* **使えない**：`fetch`, `setTimeout`, `setInterval`, `window`, `process`, `Blob`, `URL` など ([Google for Developers][3])
* 代替：HTTPは **`UrlFetchApp.fetch`**、待機は **`Utilities.sleep`** ([Google for Developers][3])
* `import/export`（ES Modules）は **非対応**。ライブラリは「Apps Scriptのライブラリ機構」か、**バンドルして単一ファイル化**が必要。([Google for Developers][3])
* `async/await` は構文として使えるが、実行環境は基本同期で、I/Oはブロッキング。並列化したいときは `UrlFetchApp.fetchAll` を使う。([Google for Developers][3])

### 1.3 このリポジトリの開発・デプロイ方針（TypeScript + clasp）

**このプロジェクトは「TypeScriptでローカル開発 →（必要に応じてバンドル）→ JavaScriptとしてGASへ `clasp push`」が前提**。

* GASの実行環境は **TypeScriptを解釈しない**ため、`clasp push` の対象は **GAS互換のJavaScript（ビルド成果物）**にする。
* TypeScriptで `import/export` を使う場合、GAS側ではそのまま動かないため、**ビルドで単一ファイル化（またはimport/exportを除去）**する。
* **GASエディタでの直接編集は原則しない**（ローカルの差分レビュー/ロールバックが崩れる）。必要ならローカルへ反映→ビルド→`clasp push` の順で同期する。
* `appsscript.json`（マニフェスト）や `.clasp.json` / `.claspignore`（存在する場合）は、実行権限・反映対象を左右するため**Git管理の対象**として扱う。

### 1.4 （この案件）シートスキーマをコードで管理する方針

この案件では「スプレッドシートのタブ/ヘッダ構造」を **コードで再現可能**にする（いわゆる “Sheets migration”）。

* スキーマ管理ルール（AI向け）は `gas_for_lease_management/sheet_schema_management_rules.md` を参照する。
* 重要: スプレッドシートは人が触る前提のため、**列番号固定は禁止**（ヘッダ名で列解決）・**破壊的変更はmigrationとして明示**が基本。

### 1.5 TypeScript + ビルド + clasp 運用チェックリスト（つまずきポイント対策）

TypeScriptでローカル開発してGASへ反映するとき、事故りやすいのは **ツールチェーン差分 / モジュール・バンドル / 入口関数（トリガー・WebApp・`google.script.run`）/ V8非対応API**。
このリポジトリでは、壊れにくさ重視で **「事前ビルド → dist を push」** を原則とする。([Google for Developers][16]) ([GitHub][17])

#### 1) 「claspがTSをコンパイルしてくれる」前提にしない

* claspの挙動はバージョン/構成で差が出やすいので、安定運用として **`tsc` / `rollup` / `esbuild` 等で事前にJSへ変換**し、`dist/` のみを `clasp push` する。([GitHub][17])

#### 2) `clasp push` は “丸ごと置換” になりやすい（事故注意）

* `clasp push` は差分適用ではなく、ローカルの状態でリモートを更新する形になりやすい。([GitHub][17])
* 対策：
  * `.clasp.json` の `rootDir` を `dist` に固定して **push対象を明確化**する。([GitHub][17])
  * `.claspignore` を活用して **pushしてはいけないものを除外**する。([GitHub][17])
  * push前に `clasp show-file-status` で **何が送られるか確認**する。([GitHub][17])

#### 3) 出力（dist）の中に `import/export` を残さない（ESM禁止）

* Apps Script V8はES Modules非対応のため、最終成果物（dist）から `import/export` を無くす（IIFE/UMD等へバンドル）。([Google for Developers][3])

#### 4) 入口関数（トリガー/WebApp/HtmlService）をバンドルで消さない

* GASはトップレベルから呼べる関数が必要：
  * トリガー：`onOpen` / `onEdit` / 時間主導など
  * Webアプリ：`doGet` / `doPost`
  * HtmlService：`google.script.run.xxx()` の `xxx`
* バンドル/Tree-shaking/Minifyで入口が消える・名前が変わる事故がある。
  * 対策：入口は **トップレベルの `function` 宣言として残す**、または `globalThis.xxx = ...` のように明示エクスポートし、必要ならプラグイン（例：rollup-plugin-gas）で入口生成する。([GitHub][18])
  * **minifyはオフ**（少なくとも安定稼働するまで）。([GitHub][18])

#### 5) TSの言語機能で “GAS V8がパースできない” ものを混ぜない

* 例：`#private` フィールド、静的フィールド宣言などはApps Script V8で問題になるケースがあるため避ける（AIが混入させがち）。([Google for Developers][3])
* 対策：ビルド後の `dist` を目視し、「`import/export` や `#private` 等が残っていないか」を最低限確認する。

#### 6) Node/ブラウザ常識を持ち込まない（API差分）

* `fetch` / `setTimeout` / `URL` / `Blob` 等は無い前提。代替は `UrlFetchApp.fetch` と `Utilities.sleep`。([Google for Developers][3])
* npmパッケージは「バンドルできる」だけでなく「GASランタイム互換」かを必ず確認する。

#### 7) 型定義を入れる（ローカル補完・事故防止）

* ローカル開発では `@types/google-apps-script` を導入して型補完を効かせる。([Google for Developers][16]) ([npm][19])

#### 8) マニフェスト（`appsscript.json`）は dist に必ず同梱

* `appsscript.json` は runtime/scopes/timeZone 等の設定本体。dist運用では **distへコピーし忘れ**が事故要因になりやすい。([Google for Developers][9])

---

## 2. AIが勘違いしやすい落とし穴カタログ（誤り→回避策）

### 2.1 「Nodeのコード」をそのまま持ち込む

**誤り例**

* `fetch()` / `setTimeout()` / `import` を使う
* npmパッケージをrequireして当然のように動く前提
* TypeScript（またはESM）のまま `clasp push` して動く前提

**回避策（MUST）**

* HTTPは **UrlFetchApp**、時間待ち/リトライは **Utilities.sleep**、モジュールは **非対応**前提で設計する。([Google for Developers][3])
* TypeScriptは **ビルドしてJavaScript化**し、必要なら **バンドルして単一ファイル化**してから `clasp push` する。

---

### 2.2 「6分制限を伸ばせる」「Workspaceなら30分」などの思い込み

**誤り例**

* “プランを上げれば実行時間が伸びる” 前提で設計
* 6分超えを1回でやり切る実装

**回避策（MUST）**

* **6分/実行**は固定制約として扱い、長処理は必ず**チャンク分割**（例：1000行ずつ、次回トリガーで続き）に落とす。([Google for Developers][1])

---

### 2.3 SpreadsheetApp/DriveApp呼び出しをループで連打して遅死

**誤り例**

* 1セルごとに `getValue()` / `setValue()` / `setBackground()` などを呼ぶ
* read/write が交互に出る（キャッシュが効きにくい）

**回避策（MUST）**

* **“まとめて読む→配列で処理→まとめて書く”** を徹底する（batch）。公式ベストプラクティスの中核。([Google for Developers][4])

---

### 2.4 トリガーの種類と権限を混同する（動いたり動かなかったり）

**誤り例**

* `onEdit(e)` を書けば何でもできる前提
* “共有シートで他人が編集したら同じ権限で動く” 前提
* シンプルトリガーで30秒超・外部アクセス・権限要求を行う

**回避策（MUST）**

* **シンプルトリガー**の制限（30秒・権限やユーザー識別が状況依存等）を前提にする。([Google for Developers][2])
* 権限が必要な処理は、**インストール型トリガー**に寄せる（編集/オープン/変更など）。([Google for Developers][5])

---

### 2.5 OAuthスコープ/マニフェスト/高度なサービスを忘れて「実行時にコケる」

**誤り例**

* Drive API/Sheets APIのサンプルを貼って終わり（有効化してない）
* スコープ過大で「このアプリは未確認」系の問題を誘発

**回避策（MUST）**

* 高度なサービスは **“有効化が必要”**。手順/影響を必ず明記する。([Google for Developers][6])
* `appsscript.json` の `oauthScopes` を明示管理できる（＝過不足が出る）。スコープ最小化を設計レビュー項目に入れる。([Google for Developers][7])

---

### 2.6 UrlFetchAppの誤解（スコープ/ネットワーク/クォータ）

**誤り例**

* 外部HTTPの権限（スコープ）や、IP制限/allowlistを考えない
* 失敗時に即リトライ連打して “短時間の呼び出し過多” で落ちる

**回避策（MUST）**

* UrlFetchAppには外部リクエスト用スコープが必要。([Google for Developers][8])
* マニフェストでURL allowlist を設定できる（組織要件がある場合は特に）。([Google for Developers][9])
* クォータ（1日回数等）を前提に設計する。([Google for Developers][1])

---

### 2.7 同時実行・二重実行でデータが壊れる（地味に多い）

**誤り例**

* 同じシート/同じ行を複数実行が同時に触る
* “途中まで進んだ状態” を記録せず、再実行で二重反映

**回避策（MUST）**

* クリティカルセクションは **LockService** で排他する（script/document/user lock を目的に応じて選ぶ）。([Google for Developers][10])
* 進捗/再開点は **PropertiesService** 等で保持し、処理は**冪等**にする（同じ入力を2回処理しても結果が壊れない）。([Google for Developers][11])

---

### 2.8 HtmlService/UI系は “IFRAME前提” なのを忘れる

**誤り例**

* 古い sandbox モード前提のコード/記事をそのまま採用
* setSandboxMode が効く前提

**回避策（MUST）**

* HtmlServiceは **IFRAME以外のsandboxはサンセット**、`setSandboxMode` は実質効果なし。([Google for Developers][12])

---

### 2.9 Apps Script API（scripts.run）の勘違い（サービスアカウントで叩ける等）

**誤り例**

* “サービスアカウントで scripts.run してOK” と設計してしまう

**回避策（MUST）**

* **Apps Script APIはサービスアカウントで動かない**（警告が明記）。([Google for Developers][13])
* 逆に「Apps Script内から service account を使って他APIを呼ぶ」話と混同しない（別物）。([Google for Developers][14])

---

## 3. 安定した開発フロー（AIが守る手順）

### フェーズA：仕様の確定（実装前に必須）

AIは最初に、最低限これを確認できない場合は「不足情報」として列挙する（質問は最小限）。

* 入力：対象（Sheets/Drive/Gmail等）、シート名・列構造・キー、例データ
* 出力：どこに何を書き、既存データをどう扱うか
* 制約：実行トリガー種別、実行頻度、6分制限を超える可能性、外部API有無
* 失敗設計：再実行時の挙動、途中再開、ログ/通知先

### フェーズB：土台（開発環境/バージョン管理）

* **TypeScriptでローカル開発 → ビルド → clasp push**（差分レビュー・小刻みコミット・ロールバックを可能にする）。([Google for Developers][15])
* ビルド成果物がpush対象になっているか（`clasp status` 等で）を確認し、**「ローカルのTypeScript変更がGASへ反映される経路」**を常に維持する。

### フェーズC：実装（小さく・テスト可能に）

**MUST**

* “サービス呼び出し層” と “純粋ロジック” を分離する

  * ロジックは配列/オブジェクトを入力に取り配列/オブジェクトを返す（テスト容易）
* Spreadsheet/Drive等は **一括読み書き**（必要ならRangeListや連続範囲の工夫）([Google for Developers][4])
* 長処理は **チャンク化 + 進捗保存（PropertiesService） + 排他（LockService）** ([Google for Developers][11])

### フェーズD：性能/クォータ最適化（レビュー項目として固定）

* ループ内のサービス呼び出しが無いか
* read/write が交互になってないか
* UrlFetchAppを連打してないか（fetchAll/キャッシュ/バックオフ）
* トリガー実行時間や日次クォータに対して安全か ([Google for Developers][1])

### フェーズE：認可/デプロイ/運用

* `oauthScopes` を明示管理する場合は、必要スコープを漏れなく最小で ([Google for Developers][7])
* 高度なサービスは有効化が必要（使うなら手順も成果物に含める）([Google for Developers][6])
* 実行ログは “あとで追える形” に（V8では console ログが実行履歴に出る推奨がある）([Google for Developers][3])

---

## 4. AIの出力フォーマット（毎回これで返す）

AIはコードを出すとき、必ず以下をセットで返す：

1. **設計メモ**：トリガー/実行制限/冪等性/排他/進捗/クォータの方針
2. **変更点の要約**（差分レビューしやすく）
3. **コード**（必要ならファイル分割の意図も）
4. **手動テスト手順**（最短で再現できる）
5. **失敗時の確認ポイント**（ログの見方、どこを疑うか）

---

## 5. “AIレビュー”チェックリスト（実装後に必ず自己点検）

* [ ] V8非対応API（fetch/import/setTimeout等）を使っていない ([Google for Developers][3])
* [ ] Spreadsheet/Drive呼び出しがループ内にない（batch化できている）([Google for Developers][4])
* [ ] 6分制限を超えうる処理は分割され、進捗保存がある ([Google for Developers][1])
* [ ] 同時実行の排他がある（LockService）([Google for Developers][10])
* [ ] トリガー選定が正しい（simple 30秒制限・権限制約を考慮）([Google for Developers][2])
* [ ] UrlFetchAppはスコープ/allowlist/クォータを考慮している ([Google for Developers][8])
* [ ] Apps Script API（scripts.run）をサービスアカウントで叩く設計になっていない ([Google for Developers][13])
* [ ] HtmlServiceはIFRAME前提（古いsandbox前提が混ざってない）([Google for Developers][12])

---

## 6. 使い回し用：AIへの指示テンプレ（短いが効く）

### 実装依頼テンプレ

* 前提：Apps Script（V8）。6分/実行制限、simple triggerは30秒制限。
* 必須：Spreadsheet/Drive/UrlFetch等の呼び出しはバッチ化。冪等・排他（LockService）・進捗保存（PropertiesService）。
* 出力：設計メモ→差分要約→コード→手動テスト→失敗時の観点。

### バグ修正依頼テンプレ

* このコードを修正して：

  1. V8非対応API混入の有無
  2. ループ内サービス呼び出しの有無
  3. トリガー制限（30秒/権限）
  4. 冪等性・排他・進捗
     を順に潰して。修正は最小差分で。

---

## 7. 参考（一次資料：公式が最優先）

* 実行時間/クォータ（6分/実行、URL Fetch回数、trigger総実行時間等）([Google for Developers][1])
* パフォーマンスのベストプラクティス（サービス呼び出し最小化・バッチ）([Google for Developers][4])
* V8ランタイムの制約（fetch不可、import/export不可、fetchAll推奨等）([Google for Developers][3])
* トリガー（simple / installable）([Google for Developers][2])
* LockService / PropertiesService ([Google for Developers][10])
* clasp（ローカル開発）([Google for Developers][15])
* TypeScript（ローカル開発）([Google for Developers][16])
* clasp（GitHub）([GitHub][17])
* rollup-plugin-gas（入口関数/バンドル注意）([GitHub][18])
* `@types/google-apps-script`（型定義）([npm][19])
* Apps Script API（scripts.run）※サービスアカウント不可 ([Google for Developers][13])
* HtmlService sandbox（IFRAMEのみ）([Google for Developers][12])

---


[1]: https://developers.google.com/apps-script/guides/services/quotas "Quotas for Google Services  |  Apps Script  |  Google for Developers"
[2]: https://developers.google.com/apps-script/guides/triggers?utm_source=chatgpt.com "Simple Triggers | Apps Script"
[3]: https://developers.google.com/apps-script/guides/v8-runtime "V8 runtime overview  |  Apps Script  |  Google for Developers"
[4]: https://developers.google.com/apps-script/guides/support/best-practices "Best Practices  |  Apps Script  |  Google for Developers"
[5]: https://developers.google.com/apps-script/guides/triggers/installable?utm_source=chatgpt.com "Installable Triggers | Apps Script"
[6]: https://developers.google.com/apps-script/guides/services/advanced?utm_source=chatgpt.com "Advanced Google services | Apps Script"
[7]: https://developers.google.com/apps-script/concepts/scopes?utm_source=chatgpt.com "Authorization Scopes | Apps Script"
[8]: https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app?utm_source=chatgpt.com "Class UrlFetchApp | Apps Script"
[9]: https://developers.google.com/apps-script/manifest?utm_source=chatgpt.com "Manifest structure | Apps Script"
[10]: https://developers.google.com/apps-script/reference/lock?utm_source=chatgpt.com "Lock Service | Apps Script"
[11]: https://developers.google.com/apps-script/guides/properties?utm_source=chatgpt.com "Properties Service | Apps Script"
[12]: https://developers.google.com/apps-script/guides/html/restrictions?utm_source=chatgpt.com "HTML Service: Restrictions | Apps Script"
[13]: https://developers.google.com/apps-script/api/how-tos/execute?utm_source=chatgpt.com "Execute Functions with the Apps Script API"
[14]: https://developers.google.com/apps-script/guides/service-account?utm_source=chatgpt.com "Authenticate as an Apps Script project using service accounts"
[15]: https://developers.google.com/apps-script/guides/clasp?utm_source=chatgpt.com "Use the command line interface with clasp | Apps Script"
[16]: https://developers.google.com/apps-script/guides/typescript "Develop Apps Script using TypeScript | Google for Developers"
[17]: https://github.com/google/clasp "GitHub - google/clasp: Command Line Apps Script Projects"
[18]: https://github.com/mato533/rollup-plugin-gas "GitHub - mato533/rollup-plugin-gas: Rollup plugin for Google Apps Script"
[19]: https://www.npmjs.com/package/%40types/google-apps-script "types/google-apps-script"
