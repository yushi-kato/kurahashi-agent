# 文書ステータスと時系列（車両リース更新通知）

## 0. この文書の目的

時系列で文書が増えると、同じテーマの資料でも「現行運用を説明している文書」と「次の構想を説明している文書」が混在します。  
この文書は、次の2軸を分けて判断できるようにするための索引です。

- 要件としての最新（今後どうしたいか）
- 実装としての最新（いま実際に動いているか）

## 1. 用語（意味から整理）

- 正本（せいほん）
  - その時点で「判断の基準として使う主文書」のこと。複数ある場合は用途別に正本を分ける。
- 現行実装
  - `src/main.ts` の実際の挙動と一致している状態。
- 次期要件
  - まだ実装に反映しきれていないが、次に目指す業務ルール。
- 履歴参照
  - 決定理由や経緯を残すために読む文書。実装基準としては直接使わない。

## 2. 時系列（意思決定の流れ）

1. 2026-01-18: 基本方針と初期実装計画を作成（部署単位通知 + フォーム回答）
2. 2026-01-19〜2026-01-20: フォーム回答仕様とリマインド仕様を追加
3. 2026-02-02: 承認/差戻しフロー追加案を作成（既存方式の拡張）
4. 2026-02-06: 半期一括運用への転換案を確定（部署単位から半期バッチへ）
5. 2026-02-11: 実装レビューで「半期一括案は主要部分が未実装」と確認
6. 2026-02-11: 不要機能棚卸しと「削除 + 新規実装」を統合した移行計画を作成
7. 2026-02-11: 移行計画に基づき、フェーズ1（通知バッチ・半期設定キー・確認用シート生成）の実装着手
8. 2026-02-11: 同計画のフェーズ2〜4（通知/承認/反映導線・半期トリガー・旧導線停止）を実装し、半期一括フローへ切替
9. 2026-02-12: 半期一括フローのE2E手順書（モックデータ準備込み）を追加
10. 2026-02-12: 現行運用シートの意味と役割を整理したシート意味一覧を追加
11. 2026-02-12: 業務/システムのシーケンス図ドキュメントを追加
12. 2026-02-12: `車両（統合ビュー）` を廃止し、`車両一覧` 正本へ即時一本化（データ移行処理なし）を実装・文書反映

## 3. 現在の読み分け（重要）

### 3.1 要件としての最新（業務の目標）

- `docs/flows/business_flow_vehicle_lease_renewal_hq_biannual_confirmation.md`
- `docs/plans/plan_vehicle_lease_renewal_hq_biannual_confirmation.md`
- `docs/plans/plan_vehicle_lease_renewal_hq_biannual_migration_and_rebuild.md`

上記は「どこへ向かうか」の基準です。  
`docs/reviews/review_vehicle_lease_renewal_hq_biannual_confirmation_implementation.md` は 2026-02-11 午前時点の差分監査として扱い、同日夜の実装で半期一括導線へ切替済みです。

### 3.2 実装としての最新（現行運用）

- `docs/operations/operation_manual_vehicle_lease_renewal.md`
- `docs/guides/e2e_test_vehicle_lease_renewal_hq_biannual.md`
- `docs/guides/sheet_meanings_vehicle_lease_renewal_hq_biannual.md`
- `docs/guides/sequence_diagrams_vehicle_lease_renewal_hq_biannual.md`

これらは `車両一覧` 正本の現行実装と一致している文書です。  
旧フォーム系・部署単位の文書は履歴参照として扱います。

### 3.3 履歴参照（検討経緯）

- `docs/plans/plan_vehicle_lease_renewal_approval_flow.md`
- `docs/briefs/brief_vehicle_lease_renewal_changes_for_inoue.md`
- `docs/briefs/brief_vehicle_lease_renewal_changes_for_senmu.md`
- `docs/references/2026-01-27_vehicle-lease-automation_email/email_body.md`
- `docs/reviews/review_obsolete_features_for_hq_biannual_migration.md`

これらは「なぜその変更案が出たか」を理解するための文書です。  
単体で実装基準にせず、必ず現行実装文書または次期要件文書と併読します。

## 4. 文書ごとの位置づけ一覧

| ファイル | 主な内容 | 区分 | 現在の扱い |
| --- | --- | --- | --- |
| `docs/guides/about_this_project.md` | 案件の出発点 | 履歴参照 | 背景確認用 |
| `docs/guides/vehicle_management_structure.md` | 元台帳構造 | 共通基盤 | 現行/次期どちらでも参照 |
| `docs/guides/sheet_schema_management_rules.md` | スキーマ変更の安全ルール | 共通基盤 | 常時有効 |
| `docs/guides/development_tips_for_gas.md` | GAS開発の制約と実装上の注意 | 共通基盤 | 常時有効 |
| `docs/guides/test_automation_vehicle_lease_renewal.md` | テスト実行手順 | 共通基盤 | 常時有効 |
| `docs/guides/e2e_test_vehicle_lease_renewal_hq_biannual.md` | 半期一括フローのE2E手順（モックデータ準備込み） | 実装基準 | 常時有効 |
| `docs/guides/sheet_meanings_vehicle_lease_renewal_hq_biannual.md` | 現行運用シートの意味・役割・更新主体 | 実装基準 | 常時有効 |
| `docs/guides/sequence_diagrams_vehicle_lease_renewal_hq_biannual.md` | 半期一括フローの業務/システムシーケンス図 | 実装基準 | 常時有効 |
| `docs/plans/plan_vehicle_lease_renewal_gas.md` | 初期実装計画（部署単位） | 履歴参照 | 背景理解用 |
| `docs/flows/business_flow_vehicle_lease_renewal.md` | 初期業務フロー（部署単位） | 履歴参照 | 背景理解用 |
| `docs/specs/spec_vehicle_lease_renewal_answer_via_google_forms.md` | フォーム回答仕様 | 履歴参照 | 背景理解用 |
| `docs/specs/spec_vehicle_lease_renewal_reminder.md` | リマインド仕様 | 履歴参照 | 背景理解用 |
| `docs/operations/operation_manual_vehicle_lease_renewal.md` | 現行運用手順 | 実装基準 | 有効（車両一覧一本化を反映済み） |
| `docs/plans/plan_vehicle_lease_renewal_approval_flow.md` | 承認/差戻し追加案 | 中間検討 | 一部のみ採用、履歴扱い |
| `docs/briefs/brief_vehicle_lease_renewal_changes_for_inoue.md` | 井上さん向け要点 | 中間検討 | 説明履歴 |
| `docs/briefs/brief_vehicle_lease_renewal_changes_for_senmu.md` | 専務向け要点 | 中間検討 | 説明履歴 |
| `docs/plans/plan_vehicle_lease_renewal_hq_biannual_confirmation.md` | 半期一括の確定計画 | 次期要件基準 | 目標として最新 |
| `docs/plans/plan_vehicle_lease_renewal_hq_biannual_migration_and_rebuild.md` | 半期一括への移行計画（削除 + 新規実装統合） | 次期要件基準 | 実装済み項目の背景参照（車両一覧一本化まで反映済み） |
| `docs/flows/business_flow_vehicle_lease_renewal_hq_biannual_confirmation.md` | 半期一括の業務定義 | 次期要件基準 | 目標として最新 |
| `docs/reviews/review_vehicle_lease_renewal_hq_biannual_confirmation_implementation.md` | 実装との差分レビュー | 差分監査 | 2026-02-11時点の実装実態を示す最新 |
| `docs/reviews/review_obsolete_features_for_hq_biannual_migration.md` | 不要機能の棚卸し（削除候補一覧） | 差分監査 | 移行時の削除範囲を固定する基準 |
| `docs/references/2026-01-27_vehicle-lease-automation_email/README.md` | 一次情報の出典整理 | 履歴参照 | 背景根拠 |
| `docs/references/2026-01-27_vehicle-lease-automation_email/email_body.md` | 相談メール本文 | 履歴参照 | 背景根拠 |

## 5. 運用ルール（今後の追加時）

新しい文書を追加する場合は、先頭付近に次の4点を必ず書く。

1. 作成日（必要なら更新日）
2. 何の正本か（要件基準 / 実装基準 / 履歴参照）
3. 置き換える文書（置き換えがある場合）
4. 実装との一致状態（一致 / 部分一致 / 未実装）

これを揃えると、時系列で文書が増えても「どれを今読むべきか」が迷子になりにくくなります。
