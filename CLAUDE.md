# 作業時間一元管理 プロジェクト運用ルール

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
