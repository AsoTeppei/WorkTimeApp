-- ============================================================
-- 作業時間 一元管理システム — 初期化: スキーマ作成
--   DB・テーブル・インデックス・トリガーを新規作成する。
--   **既存 DB を削除** するため、本番環境では絶対に実行しないこと。
--   本番運用中のスキーマ変更は db\migrations\ 配下のマイグレーションで行う。
-- ============================================================

USE master;
GO

IF DB_ID(N'WorkTimeDB') IS NOT NULL
BEGIN
    ALTER DATABASE WorkTimeDB SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE WorkTimeDB;
    PRINT N'既存の WorkTimeDB を削除しました。';
END
GO

CREATE DATABASE WorkTimeDB
    COLLATE Japanese_CI_AS;
GO
PRINT N'データベース WorkTimeDB を作成しました。';
GO

USE WorkTimeDB;
GO

-- ============================================================
-- テーブル作成
-- ============================================================

-- 部署マスタ
CREATE TABLE Departments (
    DepartmentID   INT           IDENTITY(1,1) NOT NULL
                                 CONSTRAINT PK_Departments PRIMARY KEY,
    DepartmentName NVARCHAR(50)  NOT NULL
                                 CONSTRAINT UQ_Departments_Name UNIQUE,
    DepartmentCode NVARCHAR(20)  NULL,
    SortOrder      INT           NOT NULL DEFAULT 0,
    IsActive       BIT           NOT NULL DEFAULT 1,
    CreatedAt      DATETIME2     NOT NULL DEFAULT GETDATE()
);
PRINT N'テーブル Departments を作成しました。';
GO

-- ユーザーマスタ
-- Role: 'member' | 'leader' | 'admin'
-- PinHash: scrypt で salt 付きハッシュ化した PIN。NULL の場合は未設定（初回ログイン時に本人が設定）
-- TokenVersion: 認証トークンの版。PIN リセット/変更時にインクリメントして既存トークンを無効化
CREATE TABLE Users (
    UserID         INT           IDENTITY(1,1) NOT NULL
                                 CONSTRAINT PK_Users PRIMARY KEY,
    DepartmentID   INT           NOT NULL,
    Name           NVARCHAR(50)  NOT NULL,
    Email          NVARCHAR(100) NULL,
    Role           NVARCHAR(20)  NOT NULL DEFAULT N'member'
                                 CONSTRAINT CK_Users_Role
                                 CHECK (Role IN (N'member', N'leader', N'admin')),
    PinHash        NVARCHAR(255) NULL,
    PinSetAt       DATETIME2     NULL,
    TokenVersion   INT           NOT NULL DEFAULT 1,
    IsActive       BIT           NOT NULL DEFAULT 1,
    CreatedAt      DATETIME2     NOT NULL DEFAULT GETDATE(),
    CONSTRAINT FK_Users_Departments
        FOREIGN KEY (DepartmentID) REFERENCES Departments(DepartmentID)
);
PRINT N'テーブル Users を作成しました。';
GO

-- 作業内容マスタ
CREATE TABLE WorkTypes (
    WorkTypeID     INT           IDENTITY(1,1) NOT NULL
                                 CONSTRAINT PK_WorkTypes PRIMARY KEY,
    DepartmentID   INT           NOT NULL,
    TypeName       NVARCHAR(100) NOT NULL,
    SortOrder      INT           NOT NULL DEFAULT 0,
    IsActive       BIT           NOT NULL DEFAULT 1,
    CreatedAt      DATETIME2     NOT NULL DEFAULT GETDATE(),
    CONSTRAINT FK_WorkTypes_Departments
        FOREIGN KEY (DepartmentID) REFERENCES Departments(DepartmentID)
);
PRINT N'テーブル WorkTypes を作成しました。';
GO

-- 注番マスタ
CREATE TABLE Projects (
    ProjectID      INT           IDENTITY(1,1) NOT NULL
                                 CONSTRAINT PK_Projects PRIMARY KEY,
    ProjectNo      NVARCHAR(50)  NOT NULL
                                 CONSTRAINT UQ_Projects_No UNIQUE,
    ClientName     NVARCHAR(100) NOT NULL,
    Subject        NVARCHAR(200) NOT NULL,
    LastUsedDate   DATETIME2     NULL,
    IsClosed       BIT           NOT NULL DEFAULT 0,
    ClosedAt       DATETIME2     NULL,  -- クローズ日時（手動 or 自動クローズ時に記録）
    ClosedReason   NVARCHAR(50)  NULL,  -- 'manual' | 'auto_inactive' | 'annual'
    CreatedAt      DATETIME2     NOT NULL DEFAULT GETDATE()
);
PRINT N'テーブル Projects を作成しました。';
GO

-- 作業記録
-- ※ ProjectNo は NULL 許容（注番がない作業にも対応）
-- ※ DepartmentID は「記録時点の所属部署」を保持する（Users.DepartmentID のスナップショット）
CREATE TABLE WorkLogs (
    LogID           BIGINT        IDENTITY(1,1) NOT NULL
                                  CONSTRAINT PK_WorkLogs PRIMARY KEY,
    ProjectNo       NVARCHAR(50)  NULL,
    WorkDate        DATE          NOT NULL,
    UserID          INT           NOT NULL,
    DepartmentID    INT           NOT NULL,
    ContentName     NVARCHAR(100) NOT NULL,
    WorkHours       DECIMAL(5,2)  NOT NULL
                                  CONSTRAINT CK_WorkLogs_Hours CHECK (WorkHours > 0 AND WorkHours <= 24),
    WorkLocation    NVARCHAR(10)  NOT NULL DEFAULT N'社内'
                                  CONSTRAINT CK_WorkLogs_Location CHECK (WorkLocation IN (N'社内', N'社外')),
    IsAfterShipment BIT           NOT NULL DEFAULT 0,
    Details         NVARCHAR(MAX) NULL,
    IsDeleted       BIT           NOT NULL DEFAULT 0,
    CreatedAt       DATETIME2     NOT NULL DEFAULT GETDATE(),
    UpdatedAt       DATETIME2     NOT NULL DEFAULT GETDATE(),
    CONSTRAINT FK_WorkLogs_Users
        FOREIGN KEY (UserID)       REFERENCES Users(UserID),
    CONSTRAINT FK_WorkLogs_Departments
        FOREIGN KEY (DepartmentID) REFERENCES Departments(DepartmentID),
    CONSTRAINT FK_WorkLogs_Projects
        FOREIGN KEY (ProjectNo)    REFERENCES Projects(ProjectNo),
    CONSTRAINT CK_WorkLogs_NoProject_Details
        CHECK (ProjectNo IS NOT NULL OR (Details IS NOT NULL AND LEN(Details) > 0))
);
PRINT N'テーブル WorkLogs を作成しました。';
GO

-- 注番別 目標時間マスタ
CREATE TABLE ProjectTargets (
    TargetID       INT           IDENTITY(1,1) NOT NULL
                                 CONSTRAINT PK_ProjectTargets PRIMARY KEY,
    UserID         INT           NOT NULL,
    ProjectNo      NVARCHAR(50)  NOT NULL,
    TargetHours    DECIMAL(7,2)  NOT NULL
                                 CONSTRAINT CK_ProjectTargets_Hours CHECK (TargetHours > 0),
    CreatedAt      DATETIME2     NOT NULL DEFAULT GETDATE(),
    UpdatedAt      DATETIME2     NOT NULL DEFAULT GETDATE(),
    CONSTRAINT FK_ProjectTargets_Users
        FOREIGN KEY (UserID)    REFERENCES Users(UserID),
    CONSTRAINT FK_ProjectTargets_Projects
        FOREIGN KEY (ProjectNo) REFERENCES Projects(ProjectNo),
    CONSTRAINT UQ_ProjectTargets_UserProject
        UNIQUE (UserID, ProjectNo)
);
PRINT N'テーブル ProjectTargets を作成しました。';
GO

-- 目標時間 変更履歴（監査ログ）
CREATE TABLE TargetChangeLog (
    LogID         BIGINT        IDENTITY(1,1) NOT NULL
                                CONSTRAINT PK_TargetChangeLog PRIMARY KEY,
    TargetID      INT           NOT NULL,
    UserID        INT           NOT NULL,
    ProjectNo     NVARCHAR(50)  NOT NULL,
    OldValue      DECIMAL(7,2)  NULL,
    NewValue      DECIMAL(7,2)  NOT NULL,
    Reason        NVARCHAR(500) NOT NULL,
    ChangedBy     INT           NOT NULL,
    ChangedAt     DATETIME2     NOT NULL DEFAULT GETDATE(),
    CONSTRAINT FK_TCL_Target    FOREIGN KEY (TargetID)  REFERENCES ProjectTargets(TargetID),
    CONSTRAINT FK_TCL_User      FOREIGN KEY (UserID)    REFERENCES Users(UserID),
    CONSTRAINT FK_TCL_ChangedBy FOREIGN KEY (ChangedBy) REFERENCES Users(UserID)
);
CREATE INDEX IX_TargetChangeLog_Target ON TargetChangeLog (UserID, ProjectNo, ChangedAt DESC);
PRINT N'テーブル TargetChangeLog を作成しました。';
GO

-- スキーマバージョン管理（マイグレーション適用履歴）
CREATE TABLE SchemaMigrations (
    MigrationID  NVARCHAR(100) NOT NULL CONSTRAINT PK_SchemaMigrations PRIMARY KEY,
    AppliedAt    DATETIME2     NOT NULL DEFAULT GETDATE(),
    Description  NVARCHAR(500) NULL
);
PRINT N'テーブル SchemaMigrations を作成しました。';
GO

-- ============================================================
-- インデックス（基本セット）
-- ※ 追加インデックスは db\migrations\ で管理する
-- ============================================================

CREATE INDEX IX_WorkLogs_WorkDate       ON WorkLogs      (WorkDate    DESC);
CREATE INDEX IX_WorkLogs_UserID         ON WorkLogs      (UserID);
CREATE INDEX IX_WorkLogs_DepartmentID   ON WorkLogs      (DepartmentID);
CREATE INDEX IX_WorkLogs_ProjectNo      ON WorkLogs      (ProjectNo);
CREATE INDEX IX_WorkLogs_IsDeleted      ON WorkLogs      (IsDeleted);
CREATE INDEX IX_Projects_LastUsed       ON Projects      (LastUsedDate DESC);
CREATE INDEX IX_Projects_IsClosed       ON Projects      (IsClosed);
CREATE INDEX IX_ProjectTargets_Project  ON ProjectTargets(ProjectNo);
PRINT N'基本インデックスを作成しました。';
GO

-- ============================================================
-- UpdatedAt 自動更新トリガー
-- ============================================================

CREATE TRIGGER TR_WorkLogs_UpdatedAt
ON WorkLogs
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE WorkLogs
       SET UpdatedAt = GETDATE()
      FROM WorkLogs AS w
      INNER JOIN inserted AS i ON w.LogID = i.LogID;
END;
GO

CREATE TRIGGER TR_ProjectTargets_UpdatedAt
ON ProjectTargets
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE ProjectTargets
       SET UpdatedAt = GETDATE()
      FROM ProjectTargets AS t
      INNER JOIN inserted AS i ON t.TargetID = i.TargetID;
END;
GO
PRINT N'トリガーを作成しました。';
GO

-- ============================================================
-- 01_schema で追加したマイグレーション相当の内容を登録しておく
--   （既に schema に取り込み済みの場合は、そのマイグレーションを
--    改めて適用しないための印）
-- ============================================================
INSERT INTO SchemaMigrations (MigrationID, Description) VALUES
    (N'001_worklogs_composite_indexes',   N'WorkLogs の複合インデックス（初期スキーマ適用時にスキップ）'),
    (N'002_projects_auto_close_columns',  N'Projects.ClosedAt / ClosedReason カラム（初期スキーマに内包）');
GO
PRINT N'SchemaMigrations に初期適用済みエントリを登録しました。';
GO

PRINT N'';
PRINT N'===================================';
PRINT N'  01_schema 完了';
PRINT N'===================================';
GO
