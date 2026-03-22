# シーケンス図: 車両リース更新通知（半期一括）

- 作成日: 2026-02-12
- 更新日: 2026-02-12
- 区分: 実装基準（現行コードの可視化）
- 対象コード: `/Users/yushi/work/work/クラハシ/src/main.ts`
- 目的: 「誰がどこを触るか」と「どの処理がどのシートを読み書きするか」を、業務フローとシステムフローの2視点で把握する

## 0. 背景（なぜ2種類の図に分けるか）

- 業務フロー図:
  - 人の役割と入力責任を確認するための図です。
- システムフロー図:
  - メニュー実行時に、どの関数がどのシートを読み書きするかを確認するための図です。

同じ処理でも、業務上の責任分界と、システム上の依存関係は見え方が違います。  
そのため、1枚に混ぜずに2枚へ分離しています。

## 1. 業務フロー図（誰がどこを触るか）

```mermaid
sequenceDiagram
    autonumber
    actor Ledger as "台帳管理担当"
    actor Ops as "運用担当"
    actor Timer as "定期トリガー"
    actor HQ as "本部長・副本部長"
    actor Senmu as "専務"
    actor Murata as "村田主任"
    participant GAS as "GAS（車両更新通知）"
    participant Set as "設定"
    participant List as "車両一覧"
    participant Need as "要入力"
    participant Batch as "通知バッチ"
    participant Conf as "本部長副本部長確認_YYYYMMDD"
    participant Log as "通知ログ"

    Ledger->>List: 元台帳を更新する
    Ops->>Set: 通知先・期限・送信ON/OFFを設定する

    Ops->>GAS: 初期整備として「車両一覧同期（要入力更新）」を実行
    GAS->>List: 台帳を読み取り、運用列を補完する
    GAS->>Need: 欠損/不備を出力する

    Ops->>GAS: 初回のみ「半期バッチ起票」を実行
    GAS->>List: 満了日レンジで対象抽出する
    GAS->>Batch: batchを起票する

    Timer->>GAS: 1時間ごとの syncVehicles を起動
    GAS->>List: 運用列を再補完する
    GAS->>Need: 不備ログを再生成する

    Timer->>GAS: 定期トリガーで自動進行を起動
    GAS->>Batch: 対象batchを確認する
    alt 確認用シート未作成
        GAS->>List: 対象車両を取得する
        GAS->>Conf: 確認用シートを作成する
        GAS->>Batch: 確認用シート名を記録する
    end
    alt 初回通知未送信
        GAS->>Set: 通知設定を参照する
        GAS->>Conf: 未回答件数を集計する
        GAS->>HQ: 初回依頼メールを送信する
        GAS->>Batch: 初回送信日時/状態を更新する
        GAS->>Log: 送信結果を記録する
    end

    HQ->>Conf: 本部回答を入力する
    HQ->>Conf: 回答確認済みをチェックする

    Timer->>GAS: 定期トリガーでリマインド判定
    GAS->>Conf: 未確認件数を判定する
    alt 未確認あり かつ 期限前日数到達 かつ 未送信
        GAS->>HQ: リマインドメールを送信する
        GAS->>Batch: リマインド送信日時/状態を更新する
        GAS->>Log: 送信結果を記録する
    end

    Timer->>GAS: 定期トリガーで専務依頼判定
    GAS->>Conf: 全件回答/全件確認済みを判定する
    alt 条件一致
        GAS->>Senmu: 承認依頼メールを送信する
        GAS->>Batch: 専務依頼送信日時/状態を更新する
        GAS->>Log: 送信結果を記録する
    end

    Senmu->>Conf: 専務判断/専務コメントを入力する

    HQ->>GAS: 確認用シート編集イベントで自動進行を起動
    Senmu->>GAS: 確認用シート編集イベントで自動進行を起動
    Murata->>GAS: 確認用シート編集イベントで自動進行を起動
    GAS->>Conf: 専務判断/反映条件を再評価する
    alt 専務判断反映が有効かつ条件一致
        GAS->>Batch: 専務判断ステータスを更新する
        GAS->>Log: 専務判断反映結果を記録する
    end
    alt 差戻しあり かつ 全行判断入力済み かつ 全行専務確認済みチェック済み
        GAS->>HQ: 差戻し通知メールを送信する
        GAS->>Conf: 差戻し行の本部回答/回答確認済み/村田主任確認済み/反映日時をクリアする（専務判断/コメント/確認済みは保持）
        GAS->>Batch: ステータスを初回通知送信済に戻す
        GAS->>Log: 差戻し通知結果を記録する
        Note over HQ,Conf: HQが差戻し行の本部回答を再入力 → 差戻し行の専務列が自動クリア → 専務依頼が再送信されるループ
    end
    alt 全件承認
        GAS->>Murata: 村田主任通知メールを送信する
        GAS->>Batch: 村田主任通知送信日時を記録する
        GAS->>Log: 村田主任通知結果を記録する
    end

    Murata->>Conf: 新契約日付 または 解約完了を入力する

    Ledger->>GAS: 車両一覧の編集イベントで同期を起動
    GAS->>List: onEditSourceSync で最小間隔ガード判定
    GAS->>Need: 条件一致時に要入力を更新

    alt マスター反映が有効かつ条件一致
        GAS->>List: 承認済みかつ入力完了行を反映する
        GAS->>Conf: 反映済み/反映日時を更新する
        GAS->>Batch: 最終ステータスを更新する
        GAS->>Log: 反映結果を記録する
    end

    Ops->>GAS: 例外時のみ「自動進行（最新バッチ）」を手動実行
```

## 2. システムフロー図（どの処理がどのシートで動くか）

```mermaid
sequenceDiagram
    autonumber
    actor Ops as "運用担当"
    participant Menu as "スプレッドシートメニュー"
    participant FSync as "syncVehicles"
    participant FBatch as "createBiannualBatch"
    participant FSheet as "buildConfirmationSheet"
    participant FInit as "sendHqInitialEmail"
    participant FRem as "sendHqReminderIfNeeded"
    participant FSenmu as "sendSenmuApprovalRequestIfReady"
    participant FDecision as "applySenmuDecisionFromSheet"
    participant FApply as "applyMasterUpdates"
    participant FAuto as "runAutoAdvance"
    participant FEdit as "onEditAutoAdvance"
    participant FSourceEdit as "onEditSourceSync"
    participant Set as "設定"
    participant List as "車両一覧"
    participant Need as "要入力"
    participant Batch as "通知バッチ"
    participant Conf as "本部長副本部長確認_YYYYMMDD"
    participant Log as "通知ログ"

    Ops->>Menu: 車両一覧同期（要入力更新）
    Menu->>FSync: 実行
    FSync->>List: READ（元台帳 + 運用列）
    FSync->>List: WRITE（vehicleId/登録番号_結合/一次回答 補完）
    FSync->>Need: WRITE（不備検出）

    Ops->>Menu: 半期バッチ起票
    Menu->>FBatch: 実行
    FBatch->>Set: READ（送付日/期限）
    FBatch->>List: READ（満了日レンジ抽出）
    FBatch->>Batch: READ（重複batch確認）
    FBatch->>Batch: WRITE（起票）

    Ops->>Menu: 確認用シート生成
    Menu->>FSheet: 実行
    FSheet->>Batch: READ（対象batch）
    FSheet->>List: READ（対象車両）
    FSheet->>Conf: WRITE（新規作成/入力ルール）
    FSheet->>Batch: WRITE（確認用シート名/対象件数）

    Ops->>Menu: 初回通知送信
    Menu->>FInit: 実行
    FInit->>Set: READ（通知先/送信ON-OFF）
    FInit->>Batch: READ（送信済み判定）
    FInit->>Conf: READ（件数集計）
    FInit->>Batch: WRITE（初回通知送信日時/状態）
    FInit->>Log: WRITE（結果）

    Ops->>Menu: リマインド送信
    Menu->>FRem: 実行
    FRem->>Set: READ（期限前日数/通知先）
    FRem->>Batch: READ（期限/送信済み）
    FRem->>Conf: READ（未確認件数）
    alt 送信条件一致
        FRem->>Batch: WRITE（リマインド送信日時/状態）
    end
    FRem->>Log: WRITE（結果またはスキップ理由）

    Ops->>Menu: 専務依頼送信
    Menu->>FSenmu: 実行
    FSenmu->>Set: READ（専務通知先）
    FSenmu->>Batch: READ（送信済み判定）
    FSenmu->>Conf: READ（全件確認済み判定）
    alt 送信条件一致
        FSenmu->>Batch: WRITE（専務依頼送信日時/状態）
    end
    FSenmu->>Log: WRITE（結果または未送信理由）

    Ops->>Menu: 専務判断反映
    Menu->>FDecision: 実行
    FDecision->>Batch: READ（対象batch）
    FDecision->>Conf: READ（専務判断）
    FDecision->>Batch: WRITE（ステータス更新）
    FDecision->>Log: WRITE（集計結果）

    Ops->>Menu: マスター反映
    Menu->>FApply: 実行
    FApply->>Batch: READ（対象batch）
    FApply->>Conf: READ（承認/必須入力）
    FApply->>List: READ（vehicleId対応）
    FApply->>List: WRITE（更新反映/解約グレーアウト）
    FApply->>Conf: WRITE（反映済み/反映日時）
    FApply->>Batch: WRITE（最終ステータス）
    FApply->>Log: WRITE（反映結果）

    Note over FAuto,FEdit: 自動進行モード（設定ON時）
    FAuto->>FSheet: 確認用シート未作成なら生成
    FAuto->>FInit: 初回通知未送信なら送信
    FAuto->>FRem: 条件一致時のみリマインド
    FAuto->>FSenmu: 全件確認済みなら専務依頼送信
    FEdit->>FDecision: 専務判断反映を自動実行
    FEdit->>FApply: 反映条件を満たす行を自動反映

    Note over FSourceEdit,FSync: 同期漏れ対策
    FSourceEdit->>FSync: 車両一覧のデータ行編集時に実行（最小間隔ガードあり）
    FAuto->>FSync: 1時間ごとの定期同期（保険）
```

## 3. 補足（読み方）

- 業務フロー図:
  - 人が入力責任を持つ箇所と、運用担当が実行責任を持つ箇所を確認します。
- システムフロー図:
  - 影響調査時に「このメニュー実行で、どのシートが更新されるか」を確認します。
- 確認用シートが複数ある場合:
  - 正本は `通知バッチ.確認用シート名` に記録されたシートです。
