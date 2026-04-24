# YonekuraSystemDB リストア手順

`_tools\backup_db.bat` で取得した `.bak` ファイルから YonekuraSystemDB を復元する手順。

> **DB名変更について:** 本DBは以前 `WorkTimeDB` という名前でした。リネーム後のバックアップファイルは `YonekuraSystemDB_*.bak` という名前で保管されますが、リネーム以前に取得した `WorkTimeDB_*.bak` もバックアップディレクトリ `C:\WorkTimeDB_Backup\` にそのまま残っています。旧バックアップから復元する場合は手順 B の別名復元で中身確認後、名前を合わせて上書きしてください。

## 環境情報

| 項目 | 値 |
|---|---|
| SQL Server | `192.168.1.8\SQLEXPRESS` (Express Edition) |
| データベース名 | `YonekuraSystemDB` |
| バックアップ先 | `C:\WorkTimeDB_Backup\YonekuraSystemDB_YYYYMMDD_HHMMSS.bak` (ディレクトリ名は旧DB名のまま、将来的にリネーム予定) |
| 保持期間 | 30日（backup_db.bat が自動削除） |
| アプリ本体 | `D:\My Documents\各グループ\ソフト\WorkTimeApp` |
| ツール類 | `D:\My Documents\各グループ\ソフト\WorkTimeApp\_tools` |
| DB スクリプト | `D:\My Documents\各グループ\ソフト\WorkTimeApp\db` |

## SQL ログインと権限

| ログイン | 権限 | 用途 |
|---|---|---|
| `worktime_backup` | `db_backupoperator` | **バックアップ専用**（自動実行用。リストア権限なし） |
| `yonekura` | `sysadmin` | **リストア・マイグレーション・管理全般**（パスワードは管理者のみ保持) |

> ⚠ **リストアには `yonekura`（sysadmin）が必要**です。`worktime_backup` では権限不足でリストアできません。

---

## 前提作業（復元前に必ず実施）

### 1. server.js を停止する

現行 DB への接続が残っていると DB に排他ロックが掛からずリストアが失敗します。

```bat
REM HOST マシン上で
taskkill /IM node.exe /F
```

タスクスケジューラ経由や PM2 で動かしている場合はそれも停止する。

### 2. 作業前スナップショットを取得

本番 DB を壊す前に、現状バックアップを「事故前保険」として追加で 1 本取る。

```bat
cd /d D:\My Documents\各グループ\ソフト\WorkTimeApp\_tools
backup_db.bat
```

`C:\WorkTimeDB_Backup\backup.log` の末尾に `[OK] backup + checksum success` が出たことを確認。

### 3. 復元するファイルを特定

```bat
dir /O-D C:\WorkTimeDB_Backup\YonekuraSystemDB_*.bak
REM DB名変更前のバックアップを探す場合:
dir /O-D C:\WorkTimeDB_Backup\WorkTimeDB_*.bak
```

最新の .bak を復元するか、指定日時の .bak を使うかを決める。

---

## 手順 A: 同じ名前の DB に上書き復元（通常ケース）

> ⚠ この手順は現行 YonekuraSystemDB を **完全に置き換えます**。失敗すると元には戻せないので、必ず前提作業 2 の「事故前保険」を取ってから実行してください。

HOST マシンで管理者権限の cmd を開き、以下を1行ずつ実行します。`yonekura` のパスワードはコマンドごとにプロンプトで聞かれます。

```bat
set SERVER=192.168.1.8\SQLEXPRESS
set BAK=C:\WorkTimeDB_Backup\YonekuraSystemDB_20260101_020000.bak

REM 1) 接続を全て切る（他のセッションが残っていても強制切断する）
sqlcmd -S %SERVER% -U yonekura -I -Q "ALTER DATABASE [YonekuraSystemDB] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;"

REM 2) バックアップファイルの中身を確認（論理名とファイルパスを確認）
sqlcmd -S %SERVER% -U yonekura -I -Q "RESTORE FILELISTONLY FROM DISK = N'%BAK%';"

REM 3) リストア実行（REPLACE で既存 DB を上書き）
sqlcmd -S %SERVER% -U yonekura -I -Q "RESTORE DATABASE [YonekuraSystemDB] FROM DISK = N'%BAK%' WITH REPLACE, RECOVERY, STATS = 10;"

REM 4) マルチユーザーモードに戻す
sqlcmd -S %SERVER% -U yonekura -I -Q "ALTER DATABASE [YonekuraSystemDB] SET MULTI_USER;"

REM 5) 健全性チェック
sqlcmd -S %SERVER% -U yonekura -I -d YonekuraSystemDB -Q "SELECT COUNT(*) AS WorkLogsCount FROM WorkLogs; SELECT COUNT(*) AS ProjectsCount FROM Projects; SELECT COUNT(*) AS UsersCount FROM Users;"
```

> `-U yonekura` のみで `-P` を省略すると、sqlcmd がパスワードをプロンプトで聞きます。平文でコマンドに残さないためのやり方です。

---

## 手順 B: 別名で復元する（検証・比較用）

本番を触らずに中身を確認したいときは、別名の DB にリストアして読み比べる。

```bat
set SERVER=192.168.1.8\SQLEXPRESS
set BAK=C:\WorkTimeDB_Backup\YonekuraSystemDB_20260101_020000.bak

REM 1) バックアップの論理ファイル名と元の物理パスを確認
sqlcmd -S %SERVER% -U yonekura -I -Q "RESTORE FILELISTONLY FROM DISK = N'%BAK%';"
REM  → LogicalName 列を控える（DB名変更後も論理名は WorkTimeDB / WorkTimeDB_log のまま）
REM  → PhysicalName 列で元の配置先も確認できる

REM 2) 現行 SQL Server のデータパスを確認
sqlcmd -S %SERVER% -U yonekura -I -Q "SELECT SERVERPROPERTY('InstanceDefaultDataPath') AS DataPath, SERVERPROPERTY('InstanceDefaultLogPath') AS LogPath;"

REM 3) 別名 DB として、上で確認したパスに MOVE してリストア
REM    (下の C:\... 部分は手順 2 の結果に差し替える)
REM    LogicalName は バックアップファイル取得時点のDB名 (WorkTimeDB / WorkTimeDB_log) を指定する
sqlcmd -S %SERVER% -U yonekura -I -Q "RESTORE DATABASE [YonekuraSystemDB_Verify] FROM DISK = N'%BAK%' WITH MOVE 'WorkTimeDB' TO N'C:\Program Files\Microsoft SQL Server\MSSQL16.SQLEXPRESS\MSSQL\DATA\YonekuraSystemDB_Verify.mdf', MOVE 'WorkTimeDB_log' TO N'C:\Program Files\Microsoft SQL Server\MSSQL16.SQLEXPRESS\MSSQL\DATA\YonekuraSystemDB_Verify_log.ldf', RECOVERY;"

REM 4) 中身を確認（件数、最終日、スキーマバージョン等）
sqlcmd -S %SERVER% -U yonekura -I -d YonekuraSystemDB_Verify -Q "SELECT COUNT(*) AS WorkLogs, MAX(WorkDate) AS LastWorkDate FROM WorkLogs WHERE IsDeleted = 0; SELECT MigrationID, AppliedAt FROM SchemaMigrations ORDER BY AppliedAt;"

REM 5) 確認後は削除
sqlcmd -S %SERVER% -U yonekura -I -Q "DROP DATABASE [YonekuraSystemDB_Verify];"
```

---

## 手順 C: SSMS (GUI) でリストアする

コマンドが不安な場合は SSMS の GUI で行うのが安全です。

1. SSMS で `192.168.1.8\SQLEXPRESS` に `yonekura` (SQL 認証) で接続
2. 左ペインで **データベース** を右クリック → **データベースの復元**
3. ソースで **デバイス** を選び、`...` から `.bak` ファイルを追加
4. **[ オプション ]** タブで:
   - ☑ 既存のデータベースを上書きする (**REPLACE**)
   - ☑ 復元されたデータベースへの既存の接続を閉じる
   - リカバリ状態: **RESTORE WITH RECOVERY**
5. OK を押して完了
6. 完了後、**データベース** を右クリック → 更新 で一覧をリフレッシュ
7. ログイン時に server.js が起動してない状態になっているか確認（前提作業 1）

---

## リストア後の作業

### 1. マイグレーション状態の確認

古いバックアップを復元した場合、最新のマイグレーションが未適用になっている可能性があります。

```bat
sqlcmd -S 192.168.1.8\SQLEXPRESS -U yonekura -I -d YonekuraSystemDB -Q "SELECT MigrationID, AppliedAt FROM SchemaMigrations ORDER BY AppliedAt;"
```

期待される最新状態：

```
MigrationID                              AppliedAt
---------------------------------------  -------------------
001_worklogs_composite_indexes           2026-04-15 11:...
002_projects_auto_close_columns          2026-04-15 11:...
```

上記より古い場合は次のステップへ。

### 2. マイグレーションの再適用（必要時のみ）

```bat
cd /d D:\My Documents\各グループ\ソフト\WorkTimeApp\db
apply_migrations.bat
```

`SchemaMigrations` テーブルで既適用はスキップされるので、何度走らせても安全。

### 3. server.js を再起動

```bat
cd /d D:\My Documents\各グループ\ソフト\WorkTimeApp
node server.js
```

起動ログで以下が出ることを確認：
- `SQL Server プールを確立しました`
- `自動クローズ: ...` （#16 が動いている証拠）
- エラートレースがないこと

### 4. 動作確認

- ブラウザで `http://192.168.1.8:3000/` を開く
- ログインして、**一覧画面** で直近のログが見えること
- **サマリー画面** で集計が出ること
- **環境設定画面**（管理者）で注番リストが出ること
- ホーム画面上部の会計年度バッジ（第46期 後期 など）が表示されること

---

## トラブルシューティング

| 症状 | 原因 / 対応 |
|---|---|
| `Exclusive access could not be obtained because the database is in use` | 手順 A の 1)（`SET SINGLE_USER WITH ROLLBACK IMMEDIATE`）を実行してから再度リストア。それでもダメなら `taskkill /IM node.exe /F`、SSMS の他の接続を全て閉じる。 |
| `The backup set holds a backup of a database other than the existing` | 別 DB のバックアップを指定してしまっている。`RESTORE FILELISTONLY` で元の DB 名を確認する。DB名変更前のバックアップ (`WorkTimeDB_*.bak`) は中身のDB名が `WorkTimeDB` なので、手順 A で `RESTORE DATABASE [YonekuraSystemDB]` と指定しても復元自体は可能(`WITH REPLACE` 必須)。 |
| `Directory lookup for the file ... failed` | 物理ファイルのパスが元環境と違う。手順 B のように `WITH MOVE` で明示的に配置先を指定する。現行 SQL Server のデータパスは `SERVERPROPERTY('InstanceDefaultDataPath')` で確認できる。 |
| `ユーザー 'yonekura' はログインできませんでした` | SQL 認証失敗。パスワードを確認。SSMS で同じログインで接続できるかをまず試す。 |
| `CREATE INDEX ... QUOTED_IDENTIFIER` エラー | sqlcmd の `-I` オプションが抜けている。全てのコマンドに `-I` を付けて再実行する。 |
| リストア直後に 002 の ClosedAt / ClosedReason がない | 復元したバックアップが 002 適用前のもの。`apply_migrations.bat` を実行して 002 を再適用する。 |
| タスクスケジューラから `backup_db.bat` が失敗する | 「最上位の特権で実行する」にチェックが入っているか、「開始（オプション）」に `D:\My Documents\各グループ\ソフト\WorkTimeApp\_tools` が設定されているかを確認。`worktime_backup` が SQL 認証で生きているかを SSMS で確認。 |
| `backup.log` に書き込めない | HOST マシンの `C:\WorkTimeDB_Backup\` が存在するか、タスク実行アカウントに書き込み権限があるかを確認。 |

---

## 定期的に練習すること

> **リストアを練習していないバックアップは、バックアップではない。**

月に 1 回、**手順 B**（別名リストア）を実行して、以下を確認する：

1. `.bak` ファイルが破損していないこと
2. `WorkLogs` 件数が想定通りか（前月比で急激に減っていないか）
3. `MAX(WorkDate)` が直近であること（バックアップが最新かどうか）
4. `SchemaMigrations` テーブルのマイグレーション履歴が揃っていること

所要時間 5 分程度。別名 DB なので本番には影響しない。

---

## 補足: バックアップと権限の設計思想

- **バックアップは最小権限で自動化**: `worktime_backup` は `db_backupoperator` のみ。SQL サーバーが侵害されても、この認証情報からはデータ書き換えやリストアができない。
- **リストアと管理操作は最高権限を手動で**: `yonekura` は `sysadmin` だが、バックアップの自動化には使わない。緊急時の人間オペレーション専用。
- **パスワードの扱い**:
  - `worktime_backup` のパスワードは `backup_db.bat` に平文で書かれている。NTFS ACL で `Administrators` のみ読めるように制限する。
  - `yonekura` のパスワードはどのファイルにも書かない。管理者の頭の中にあるものをコマンド実行時に手入力する。
