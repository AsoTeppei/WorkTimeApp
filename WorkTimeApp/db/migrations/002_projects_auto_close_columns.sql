-- ============================================================
-- Migration 002: Projects auto-close columns
--
-- Purpose:
--   Support the "auto-lock projects with 2 months of no input"
--   feature by adding:
--     Projects.ClosedAt     DATETIME2 NULL
--     Projects.ClosedReason NVARCHAR(50) NULL
--       ('manual' | 'auto_inactive' | 'annual')
--
-- Safety:
--   - Idempotent: wrapped in IF EXISTS / ELSE so re-running is
--     a no-op. Each ALTER TABLE also has its own sys.columns
--     guard as a safety net.
--   - NULL-able ADD only. Existing rows are not affected.
--
-- NOTE on structure:
--   Same idempotence pattern as migration 001 - we avoid GO
--   inside the IF/ELSE block and avoid RETURN-based skipping.
--
-- NOTE on EXEC() for the backfill UPDATE:
--   SQL Server uses compile-time name resolution for direct
--   UPDATE statements. A direct UPDATE referencing ClosedAt /
--   ClosedReason in the same batch as the ALTER TABLE that
--   adds them would fail to parse. Wrapping the UPDATE in
--   EXEC() defers name resolution to execution time, by which
--   point the columns exist.
-- ============================================================

USE YonekuraSystemDB;
GO

IF EXISTS (SELECT 1 FROM SchemaMigrations WHERE MigrationID = N'002_projects_auto_close_columns')
BEGIN
    PRINT 'Migration 002 already applied (skipped).';
END
ELSE
BEGIN
    SET XACT_ABORT ON;
    BEGIN TRAN;

    IF NOT EXISTS (
        SELECT 1 FROM sys.columns
        WHERE object_id = OBJECT_ID('Projects') AND name = 'ClosedAt'
    )
        ALTER TABLE Projects ADD ClosedAt DATETIME2 NULL;

    IF NOT EXISTS (
        SELECT 1 FROM sys.columns
        WHERE object_id = OBJECT_ID('Projects') AND name = 'ClosedReason'
    )
        ALTER TABLE Projects ADD ClosedReason NVARCHAR(50) NULL;

    -- Backfill ClosedAt/ClosedReason for rows that are already IsClosed = 1.
    -- See the EXEC() note above.
    EXEC('UPDATE Projects
             SET ClosedAt     = ISNULL(ClosedAt, GETDATE()),
                 ClosedReason = ISNULL(ClosedReason, N''manual'')
           WHERE IsClosed = 1
             AND (ClosedAt IS NULL OR ClosedReason IS NULL);');

    INSERT INTO SchemaMigrations (MigrationID, Description)
    VALUES (N'002_projects_auto_close_columns', N'Projects.ClosedAt / ClosedReason columns (auto-close support)');

    COMMIT;
    PRINT 'Migration 002 applied.';
END
GO
