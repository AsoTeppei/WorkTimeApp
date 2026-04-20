-- ============================================================
-- Migration 003: WorkLogs IsPaidLeave column (paid-leave support)
--
-- Purpose:
--   Add IsPaidLeave BIT NOT NULL DEFAULT 0 to WorkLogs,
--   and relax two CHECK constraints so that a paid-leave entry
--   can be saved as "the day is logged, but no actual work":
--     - CK_WorkLogs_Hours
--         before: WorkHours > 0 AND WorkHours <= 24
--         after : (IsPaidLeave=1 AND WorkHours=0)
--                 OR (IsPaidLeave=0 AND WorkHours>0 AND WorkHours<=24)
--     - CK_WorkLogs_NoProject_Details
--         before: ProjectNo IS NOT NULL OR Details non-empty
--         after : IsPaidLeave=1
--                 OR ProjectNo IS NOT NULL
--                 OR Details non-empty
--
-- Safety:
--   - Idempotent via SchemaMigrations + sys.columns /
--     sys.check_constraints guards.
--   - New column has DEFAULT 0, so existing rows keep normal
--     (non-paid-leave) semantics.
--   - No existing row currently has WorkHours=0, so the new
--     CHECK will not reject any existing row.
--
-- NOTE on EXEC() for ADD CONSTRAINT:
--   A direct ALTER TABLE ADD CONSTRAINT that references the new
--   IsPaidLeave column could fail to compile in the same batch
--   on some servers. Wrapping in EXEC() defers name resolution
--   to execution time. See migration 002 for the same pattern.
-- ============================================================

SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
GO

USE WorkTimeDB;
GO

IF EXISTS (SELECT 1 FROM SchemaMigrations WHERE MigrationID = N'003_worklogs_paid_leave')
BEGIN
    PRINT 'Migration 003 already applied (skipped).';
END
ELSE
BEGIN
    SET XACT_ABORT ON;
    BEGIN TRAN;

    -- 1. Add IsPaidLeave column (default 0 = normal work log).
    IF NOT EXISTS (
        SELECT 1 FROM sys.columns
        WHERE object_id = OBJECT_ID('WorkLogs') AND name = 'IsPaidLeave'
    )
        ALTER TABLE WorkLogs ADD IsPaidLeave BIT NOT NULL
            CONSTRAINT DF_WorkLogs_IsPaidLeave DEFAULT 0;

    -- 2. Replace CK_WorkLogs_Hours so that paid-leave rows may
    --    have WorkHours = 0.
    IF EXISTS (
        SELECT 1 FROM sys.check_constraints
        WHERE name = 'CK_WorkLogs_Hours' AND parent_object_id = OBJECT_ID('WorkLogs')
    )
        ALTER TABLE WorkLogs DROP CONSTRAINT CK_WorkLogs_Hours;

    EXEC('ALTER TABLE WorkLogs ADD CONSTRAINT CK_WorkLogs_Hours
          CHECK (
              (IsPaidLeave = 1 AND WorkHours = 0)
              OR
              (IsPaidLeave = 0 AND WorkHours > 0 AND WorkHours <= 24)
          );');

    -- 3. Replace CK_WorkLogs_NoProject_Details so that paid-leave
    --    rows do not need a project or details.
    IF EXISTS (
        SELECT 1 FROM sys.check_constraints
        WHERE name = 'CK_WorkLogs_NoProject_Details' AND parent_object_id = OBJECT_ID('WorkLogs')
    )
        ALTER TABLE WorkLogs DROP CONSTRAINT CK_WorkLogs_NoProject_Details;

    EXEC('ALTER TABLE WorkLogs ADD CONSTRAINT CK_WorkLogs_NoProject_Details
          CHECK (
              IsPaidLeave = 1
              OR ProjectNo IS NOT NULL
              OR (Details IS NOT NULL AND LEN(Details) > 0)
          );');

    INSERT INTO SchemaMigrations (MigrationID, Description)
    VALUES (N'003_worklogs_paid_leave',
            N'WorkLogs.IsPaidLeave column + relaxed CHECK constraints for paid-leave entries');

    COMMIT;
    PRINT 'Migration 003 applied.';
END
GO
