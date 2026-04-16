-- ============================================================
-- 作業時間 一元管理システム — 旧 setup_database.sql
--
-- ★★★ このファイルは非推奨です ★★★
--
-- DB 構築・マイグレーションは db\ フォルダに移行しました。
--
--   新規 DB の初期化        : db\run_init.bat
--   本番へのマイグレーション: db\apply_migrations.bat
--   詳細                    : db\README.md
--
-- 互換のため内容を残しておきたい場合は、db\init 配下を順に
-- 実行してください（01_schema.sql → 02_masterdata.sql →
-- 03_sampledata.sql）。01_schema.sql は既存 DB を DROP する
-- ので本番環境では絶対に実行しないこと。
-- ============================================================

PRINT N'このスクリプトは非推奨です。db\README.md を参照して run_init.bat / apply_migrations.bat を使用してください。';
GO
