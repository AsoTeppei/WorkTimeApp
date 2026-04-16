# 作業時間 一元管理システム — DB スクリプト運用

従来の `setup_database.sql` は **破壊的な初期化専用** なので、本番運用中に
スキーマ変更したい場合に使えなかった。そこで次の 2 系統に分離する:

```
db/
├── init/                      ← 新規 DB を 0 から作るスクリプト
│   ├── 01_schema.sql            DB/テーブル/インデックス/トリガー（DROP DATABASE あり）
│   ├── 02_masterdata.sql        部署・作業内容マスタ（IF NOT EXISTS で安全）
│   └── 03_sampledata.sql        サンプルユーザー・注番・作業ログ（デモ用）
│
├── migrations/                ← 本番 DB に差分だけ適用するスクリプト
│   ├── 001_worklogs_composite_indexes.sql
│   └── 002_projects_auto_close_columns.sql
│
├── run_init.bat               ← 新規 DB 初期化（本番では実行しない）
└── apply_migrations.bat       ← 本番 DB にマイグレーションを一括適用
```

## 初期構築（新規 DB）

```bat
cd WorkTimeApp\db
run_init.bat
apply_migrations.bat
```

`run_init.bat` は既存の WorkTimeDB を DROP するので、**本番環境では実行しないこと**。

## 本番 DB の更新

```bat
cd WorkTimeApp\db
apply_migrations.bat
```

`SchemaMigrations` テーブルで適用済みチェックを行うので、同じスクリプトを
再度流しても冪等（何度実行してもデータに影響なし）。

## マイグレーションの追加

- `db/migrations/NNN_説明.sql` を追加する（3桁連番）。
- 各スクリプトの先頭で `SchemaMigrations` を見てスキップ、末尾で INSERT する。
- 本番適用前に **必ずバックアップを取ること**（`_tools/backup_db.bat`）。

## 01_schema.sql の「内包済み」扱い

`01_schema.sql` は初期スキーマに既に `migrations/001,002` 相当のカラム・
インデックス定義を含む。新規 DB を立てるとき、これらのマイグレーションは
既に適用済み状態として `SchemaMigrations` テーブルに記録されるので、
`apply_migrations.bat` を後から流してもスキップされる。

## 旧 `setup_database.sql` の扱い

互換性のため `setup_database.sql` は残してあるが、内容は
`run_init.bat` を呼ぶ案内コメントのみ。今後の保守は `db/` 配下で行う。
