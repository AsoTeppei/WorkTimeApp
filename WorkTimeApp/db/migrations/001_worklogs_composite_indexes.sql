-- ============================================================
-- Migration 001: WorkLogs composite indexes
--
-- Purpose: Keep aggregation/filter queries fast as data grows.
--          Adds composite and filtered indexes for hot query
--          patterns.
--
-- Safety:
--   - Idempotent: wrapped in IF EXISTS / ELSE so re-running is
--     a no-op. Each CREATE INDEX also has its own sys.indexes
--     guard as a safety net.
--   - Existing single-column indexes are left in place.
--   - Offline. Runs in seconds to minutes even at scale.
--
-- NOTE on SET options:
--   Filtered indexes (WHERE ...) require QUOTED_IDENTIFIER ON
--   and ANSI_NULLS ON on the connection. sqlcmd defaults to OFF
--   unless invoked with -I, so we also set them explicitly.
--
-- NOTE on structure:
--   We deliberately avoid GO inside the IF/ELSE block and avoid
--   using RETURN as a skip pattern. RETURN only exits the current
--   GO-delimited batch, so a RETURN followed by another batch
--   would silently fall through on the second run and hit a PK
--   violation on SchemaMigrations.
-- ============================================================

SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
SET ANSI_PADDING ON;
SET ANSI_WARNINGS ON;
SET ARITHABORT ON;
SET CONCAT_NULL_YIELDS_NULL ON;
SET NUMERIC_ROUNDABORT OFF;
GO

USE YonekuraSystemDB;
GO

IF EXISTS (SELECT 1 FROM SchemaMigrations WHERE MigrationID = N'001_worklogs_composite_indexes')
BEGIN
    PRINT 'Migration 001 already applied (skipped).';
END
ELSE
BEGIN
    SET XACT_ABORT ON;
    BEGIN TRAN;

    -- (UserID, WorkDate DESC) - per-user timeline
    IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_WorkLogs_UserDate' AND object_id = OBJECT_ID('WorkLogs'))
        CREATE INDEX IX_WorkLogs_UserDate
            ON WorkLogs (UserID, WorkDate DESC)
            WHERE IsDeleted = 0;

    -- (DepartmentID, WorkDate DESC) - per-department timeline
    IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_WorkLogs_DeptDate' AND object_id = OBJECT_ID('WorkLogs'))
        CREATE INDEX IX_WorkLogs_DeptDate
            ON WorkLogs (DepartmentID, WorkDate DESC)
            WHERE IsDeleted = 0;

    -- (ProjectNo, WorkDate DESC) - per-project timeline
    IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_WorkLogs_ProjectDate' AND object_id = OBJECT_ID('WorkLogs'))
        CREATE INDEX IX_WorkLogs_ProjectDate
            ON WorkLogs (ProjectNo, WorkDate DESC)
            WHERE IsDeleted = 0 AND ProjectNo IS NOT NULL;

    -- (WorkDate DESC, LogID DESC) + INCLUDE - list pagination
    IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_WorkLogs_DateLog_Active' AND object_id = OBJECT_ID('WorkLogs'))
        CREATE INDEX IX_WorkLogs_DateLog_Active
            ON WorkLogs (WorkDate DESC, LogID DESC)
            INCLUDE (UserID, DepartmentID, ProjectNo, ContentName, WorkHours, WorkLocation, IsAfterShipment)
            WHERE IsDeleted = 0;

    -- ProjectTargets (UserID) - per-user target map
    IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_ProjectTargets_User' AND object_id = OBJECT_ID('ProjectTargets'))
        CREATE INDEX IX_ProjectTargets_User
            ON ProjectTargets (UserID)
            INCLUDE (ProjectNo, TargetHours);

    INSERT INTO SchemaMigrations (MigrationID, Description)
    VALUES (N'001_worklogs_composite_indexes', N'WorkLogs / ProjectTargets composite and filtered indexes');

    COMMIT;
    PRINT 'Migration 001 applied.';
END
GO
