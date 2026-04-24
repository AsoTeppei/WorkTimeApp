-- ============================================================
-- Migration 004: CompanyCalendar table (admin-managed calendar)
--
-- Purpose:
--   Move company-specific calendar overrides out of hard-coded
--   Sets in server.js and into a table that admin users can edit
--   from the settings UI.
--
--   Each row represents a deliberate override of the default
--   rule ("weekend / JP holiday = not a business day, else a
--   business day"):
--     IsHoliday = 1  -> the date is a company holiday (even if
--                        it is a weekday)
--     IsHoliday = 0  -> the date is a business day (even if it
--                        falls on a Saturday / Sunday / JP holiday;
--                        e.g. a Saturday all-hands meeting)
--   If no row exists for a date, the default rule applies.
--
--   JP national holidays remain in a hard-coded list in server.js
--   (they do not change per company).
--
-- Safety:
--   - Idempotent via SchemaMigrations + sys.tables guard.
--   - Seed existing 2026 company holidays that were hard-coded in
--     server.js, so applying this migration does not change the
--     set of days users see.
-- ============================================================

SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
GO

USE YonekuraSystemDB;
GO

IF EXISTS (SELECT 1 FROM SchemaMigrations WHERE MigrationID = N'004_company_calendar')
BEGIN
    PRINT 'Migration 004 already applied (skipped).';
END
ELSE
BEGIN
    SET XACT_ABORT ON;
    BEGIN TRAN;

    IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'CompanyCalendar')
    BEGIN
        CREATE TABLE CompanyCalendar (
            CalendarDate     DATE           NOT NULL
                                            CONSTRAINT PK_CompanyCalendar PRIMARY KEY,
            IsHoliday        BIT            NOT NULL,
            Note             NVARCHAR(200)  NULL,
            UpdatedAt        DATETIME2      NOT NULL DEFAULT GETDATE(),
            UpdatedByUserID  INT            NULL,
            CONSTRAINT FK_CompanyCalendar_Users
                FOREIGN KEY (UpdatedByUserID) REFERENCES Users(UserID)
        );
    END;

    -- Seed 2026 company holidays (previously hard-coded in server.js)
    EXEC('
        IF NOT EXISTS (SELECT 1 FROM CompanyCalendar WHERE CalendarDate = ''2026-01-02'')
            INSERT INTO CompanyCalendar (CalendarDate, IsHoliday, Note) VALUES (''2026-01-02'', 1, N''Shogatsu holiday'');
        IF NOT EXISTS (SELECT 1 FROM CompanyCalendar WHERE CalendarDate = ''2026-05-01'')
            INSERT INTO CompanyCalendar (CalendarDate, IsHoliday, Note) VALUES (''2026-05-01'', 1, N''Company anniversary'');
        IF NOT EXISTS (SELECT 1 FROM CompanyCalendar WHERE CalendarDate = ''2026-08-10'')
            INSERT INTO CompanyCalendar (CalendarDate, IsHoliday, Note) VALUES (''2026-08-10'', 1, N''Obon holiday'');
        IF NOT EXISTS (SELECT 1 FROM CompanyCalendar WHERE CalendarDate = ''2026-08-14'')
            INSERT INTO CompanyCalendar (CalendarDate, IsHoliday, Note) VALUES (''2026-08-14'', 1, N''Obon holiday'');
        IF NOT EXISTS (SELECT 1 FROM CompanyCalendar WHERE CalendarDate = ''2026-11-13'')
            INSERT INTO CompanyCalendar (CalendarDate, IsHoliday, Note) VALUES (''2026-11-13'', 1, N''Company trip'');
        IF NOT EXISTS (SELECT 1 FROM CompanyCalendar WHERE CalendarDate = ''2026-12-30'')
            INSERT INTO CompanyCalendar (CalendarDate, IsHoliday, Note) VALUES (''2026-12-30'', 1, N''Year-end holiday'');
        IF NOT EXISTS (SELECT 1 FROM CompanyCalendar WHERE CalendarDate = ''2026-12-31'')
            INSERT INTO CompanyCalendar (CalendarDate, IsHoliday, Note) VALUES (''2026-12-31'', 1, N''Year-end holiday'');
    ');

    INSERT INTO SchemaMigrations (MigrationID, Description)
    VALUES (N'004_company_calendar',
            N'CompanyCalendar table for admin-managed calendar overrides + seed 2026 holidays');

    COMMIT;
    PRINT 'Migration 004 applied.';
END
GO
