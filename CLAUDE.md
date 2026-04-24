# 作業時間一元管理 (WorkTimeApp) プロジェクト運用ルール

## 社員マスター一元化ポリシー (重要、2026-04-22 決定)

**社員の基本登録は米倉社内管理ポータル (YonekuraPortal) が一元管理する**。WorkTimeApp では、社員の新規登録・基本情報の編集を行わない。

### 本アプリ (WorkTimeApp) が持つ項目 (アプリ独自列)
- `PinHash` — 認証用 PIN のハッシュ (scrypt)
  - 現状 PIN 登録 UI は WorkTimeApp にあるが、将来的に YonekuraPortal 側へ移す方向

### 本アプリで行ってはいけないこと
- 社員の新規作成・削除 UI の追加
- 氏名・フリガナ・部署・メール・住所・雇用情報など「基本情報・雇用情報」の編集 UI
- `dbo.Users` の `Name` / `DepartmentID` / `Email` / `FullNameKana` 等の更新
- 部署 (`dbo.Departments`) の新規作成・削除・編集

### 既存のユーザー管理画面の扱い
- 現在 WorkTimeApp に存在する「ユーザー管理 / ユーザー編集」機能について:
  - **基本情報 (氏名・部署・メール等) の編集は禁止**。既存の編集 UI は編集不可 (読み取り専用) に改修する方針
  - PIN 設定機能のみ残して OK (将来的に YonekuraPortal へ移管)
  - 氏名・部署等の編集導線は YonekuraPortal (`http://HOST:7050/users`) へのリンクに誘導

### DB スキーマ
- `dbo.Users` / `dbo.Departments` の基本列は YonekuraPortal 所有
- WorkTimeApp 独自列 (`PinHash` 等) は WorkTimeApp の migration で ALTER TABLE 追加
- 新しく独自列が必要な場合も ALTER TABLE (`IF NOT EXISTS` ガード付き) で対応

## 作業場所と本番運用場所

- **作業場所（ローカル）**: `c:\Users\Teppei\Claude作業フォルダ\作業時間一元管理\WorkTimeApp\`
- **本番運用場所（ネットワーク）**: `\\HOST\My Documents\各グループ\ソフト\WorkTimeApp\`

ローカルで変更したファイルは、上記の本番運用場所にも反映しないと実運用に乗らない。**機能変更のたびに、ローカル側の更新対象ファイルと合わせて「本番運用場所の同じファイルにコピー反映が必要」と必ずユーザーに伝えること**。

## 機能変更時の必須手順（毎回実施）

`WorkTimeApp/` 配下で機能を追加・変更・削除した場合、**応答の最後に必ず以下のチェックリストを提示**し、どの項目が更新対象かを明示すること。**「対象外」「要確認」は載せず、本当に更新が必要なものだけを列挙する**（ノイズを減らすため）。

### 更新候補チェックリスト

#### 1. アプリ本体
- [ ] [WorkTimeApp/server.js](WorkTimeApp/server.js) — APIエンドポイント・業務ロジック
- [ ] [WorkTimeApp/index.html](WorkTimeApp/index.html) — 画面・UI・クライアント側スクリプト
- [ ] [WorkTimeApp/daemon/](WorkTimeApp/daemon/) — 常駐プロセス関連
- [ ] [WorkTimeApp/package.json](WorkTimeApp/package.json) — 依存関係・スクリプト

#### 2. データベース
- [ ] [WorkTimeApp/setup_database.sql](WorkTimeApp/setup_database.sql) — スキーマ変更・初期データ
- [ ] [WorkTimeApp/db/](WorkTimeApp/db/) — マイグレーション・付随ファイル

#### 3. 開発・運用手順
- [ ] [WorkTimeApp/開発手順書.pptx](WorkTimeApp/開発手順書.pptx) — 開発フロー・ビルド手順
- [ ] [WorkTimeApp/start_server.bat](WorkTimeApp/start_server.bat) / [WorkTimeApp/install-service.js](WorkTimeApp/install-service.js) — 起動・サービス登録
- [ ] [WorkTimeApp/① タスクスケジューラを開く.txt](WorkTimeApp/①%20タスクスケジューラを開く.txt) — タスクスケジューラ手順

### 提示フォーマット（応答末尾に付ける）

```
【機能変更に伴う更新対象】
- <ファイル/手順> — <理由>
※ 更新が必要なものだけを列挙する。「対象外」「要確認」は書かない。

【本番運用場所への反映（忘れず実施）】
反映先: \\HOST\My Documents\各グループ\ソフト\WorkTimeApp\
- コピー対象: <上記で更新したファイルを列挙>
```

### 補足

- バグ修正で挙動が変わらない場合はドキュメント更新不要。その旨を明記する。
- SQL スキーマが変わったら `setup_database.sql` と関連する API ロジック（`server.js`）の両方を更新する。
- チェックリストに無いファイルを変更した場合は、このセクション自体の更新も候補に含めること。
