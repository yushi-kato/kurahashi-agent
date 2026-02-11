# docs ディレクトリ構成

この `docs/` は、文書の「役割（何のための文書か）」ごとに分けています。

まず最初に、文書の新旧関係を把握するため `docs/guides/document_status_and_timeline.md` を参照してください。

- `docs/guides/`
  - 開発や運用の前提・ルール・手順をまとめたガイド
- `docs/flows/`
  - 業務フローや処理の流れを示す文書
- `docs/plans/`
  - 実装計画・対応方針・設計の進め方
- `docs/specs/`
  - 仕様（入力・出力・条件など、実装の基準）
- `docs/operations/`
  - 運用マニュアル（実運用での作業手順）
- `docs/briefs/`
  - 関係者向けの要約資料
- `docs/reviews/`
  - 実装レビューや検証結果
- `docs/references/`
  - 外部参照情報（メール原文、添付資料など）

## 補足

- 文書移動に合わせて、`docs/**/*.md` と `AGENTS.md` の参照パスを更新済みです。
- 時系列で追加された文書の関係（最新要件 / 現行実装 / 履歴）は `docs/guides/document_status_and_timeline.md` に集約しています。
