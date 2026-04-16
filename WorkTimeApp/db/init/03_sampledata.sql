-- ============================================================
-- 作業時間 一元管理システム — 初期化: サンプルデータ投入
--   ユーザー37名、注番15件、作業記録約370件（過去約6ヶ月）、目標時間
--   動作確認・デモ環境向け。本番環境では実行しないこと。
-- ============================================================

USE WorkTimeDB;
GO

-- ユーザー 37名（admin:1 / leader:9 / member:27）
-- PIN はすべて NULL（初回ログイン時に本人が設定）
INSERT INTO Users (DepartmentID, Name, Email, Role)
SELECT d.DepartmentID, u.Name, u.Email, u.Role
FROM (VALUES
    (N'営業グループ',    N'松田 亜矢',   N'matsuda@caty-yonekura.co.jp',        N'member'),
    (N'営業グループ',    N'根本 以久武', N'nemoto@caty-yonekura.co.jp',         N'member'),
    (N'営業グループ',    N'加藤 侑',     N'kato@caty-yonekura.co.jp',           N'member'),
    (N'営業グループ',    N'後藤 信行',   N'gotoh@caty-yonekura.co.jp',          N'leader'),
    (N'総務',            N'尾村 優子',   N'omura@caty-yonekura.co.jp',          N'member'),
    (N'電気グループ',    N'堀之内 剛',   N'horinouchi@caty-yonekura.co.jp',     N'leader'),
    (N'電気グループ',    N'寒川 和則',   N'samukawa@caty-yonekura.co.jp',       N'member'),
    (N'高温グループ',    N'千葉 和雄',   N'chiba@caty-yonekura.co.jp',          N'member'),
    (N'ソフトグループ',  N'阿曽 哲兵',   N'aso@caty-yonekura.co.jp',            N'admin'),
    (N'製造グループ',    N'辻 啓一',     N'tsuji@caty-yonekura.co.jp',          N'member'),
    (N'ソフトグループ',  N'畠 好輝',     N'hata@caty-yonekura.co.jp',           N'member'),
    (N'電気グループ',    N'北東部 巧',   N'kitatobe@caty-yonekura.co.jp',       N'member'),
    (N'営業グループ',    N'猪狩 幸也',   N'ikari@caty-yonekura.co.jp',          N'leader'),
    (N'製造グループ',    N'服部 幸雄',   N'hattori@caty-yonekura.co.jp',        N'member'),
    (N'設計グループ',    N'田中 悠貴',   N'tanaka@caty-yonekura.co.jp',         N'member'),
    (N'製造グループ',    N'中田 憲吾',   N'nakata@caty-yonekura.co.jp',         N'leader'),
    (N'営業グループ',    N'玉岡 成美',   N'sales-support@caty-yonekura.co.jp',  N'member'),
    (N'製造グループ',    N'入江 文英',   N'irie@caty-yonekura.co.jp',           N'member'),
    (N'高温グループ',    N'福田 昌史',   N'fukuda@caty-yonekura.co.jp',         N'leader'),
    (N'設計グループ',    N'中山 明洋',   N'nakayama@caty-yonekura.co.jp',       N'leader'),
    (N'営業グループ',    N'降矢 正',     N'furuya@caty-yonekura.co.jp',         N'member'),
    (N'設計グループ',    N'澤 正治',     N'sawa@caty-yonekura.co.jp',           N'member'),
    (N'電気グループ',    N'大垣 武則',   N'oogaki@caty-yonekura.co.jp',         N'member'),
    (N'高温グループ',    N'山田 武史',   N'yamada@caty-yonekura.co.jp',         N'member'),
    (N'設計グループ',    N'相澤 洸',     N'aizawa@caty-yonekura.co.jp',         N'member'),
    (N'高温グループ',    N'渕上 浩孝',   N'fuchigami@caty-yonekura.co.jp',      N'leader'),
    (N'電気グループ',    N'野川 満徳',   N'nogawa@caty-yonekura.co.jp',         N'leader'),
    (N'ソフトグループ',  N'多上 益幸',   N'tagami@caty-yonekura.co.jp',         N'leader'),
    (N'製造グループ',    N'林田 直人',   N'hayashida@caty-yonekura.co.jp',      N'member'),
    (N'ソフトグループ',  N'伊東 真吾',   N'ito@caty-yonekura.co.jp',            N'member'),
    (N'高温グループ',    N'笹田 真愛',   N'sasada@caty-yonekura.co.jp',         N'member'),
    (N'設計グループ',    N'樋口 祐馬',   N'higuchi@caty-yonekura.co.jp',        N'member'),
    (N'製造グループ',    N'大野 純一',   N'ohno@caty-yonekura.co.jp',           N'leader'),
    (N'品質管理グループ',N'平畝 英司',   N'hiraune@caty-yonekura.co.jp',        N'leader'),
    (N'営業グループ',    N'山﨑 身旗',   N'sales-support2@caty-yonekura.co.jp', N'member'),
    (N'電気グループ',    N'石田 達廣',   N'ishida@caty-yonekura.co.jp',         N'member'),
    (N'営業グループ',    N'水野 邦惠',   N'pr@caty-yonekura.co.jp',             N'member')
) AS u (DeptName, Name, Email, Role)
JOIN Departments AS d ON d.DepartmentName = u.DeptName;
PRINT N'サンプルユーザー 37 名を投入しました。';
GO

-- 注番 15件
INSERT INTO Projects (ProjectNo, ClientName, Subject, LastUsedDate)
VALUES
    (N'P-2025-001', N'ABC工業',      N'制御盤製作',         GETDATE()),
    (N'P-2025-002', N'XYZ製作所',    N'高温炉設計',         GETDATE()),
    (N'P-2025-003', N'山田精機',     N'自動化ライン',       GETDATE()),
    (N'P-2025-004', N'東京電機',     N'計装システム',       GETDATE()),
    (N'P-2025-005', N'九州重工',     N'配管工事',           GETDATE()),
    (N'P-2025-006', N'信州メタル',   N'熱交換器製作',       GETDATE()),
    (N'P-2025-007', N'横浜精密',     N'検査装置更新',       GETDATE()),
    (N'P-2025-008', N'東海ケミカル', N'反応炉改造',         GETDATE()),
    (N'P-2025-010', N'北海道化学',   N'データ収集システム', GETDATE()),
    (N'P-2025-011', N'関西テクノ',   N'設備更新工事',       GETDATE()),
    (N'P-2025-012', N'四国工業',     N'温度監視システム',   GETDATE()),
    (N'P-2025-013', N'中国鉄鋼',     N'加熱炉制御盤',       GETDATE()),
    (N'P-2024-105', N'大阪産業',     N'ボイラー制御',       GETDATE()),
    (N'P-2024-106', N'名古屋工業',   N'熱処理装置',         GETDATE()),
    (N'P-2024-120', N'仙台精工',     N'乾燥炉更新',         GETDATE());
PRINT N'サンプル注番を投入しました（15件）。';
GO

-- サンプル作業記録（37ユーザー × 10件 = 約370件）
DECLARE @Today DATE = CAST(GETDATE() AS DATE);

;WITH
Nums AS (
    SELECT n FROM (VALUES (1),(2),(3),(4),(5),(6),(7),(8),(9),(10)) AS x(n)
),
RankedProjects AS (
    SELECT
        ProjectNo,
        ROW_NUMBER() OVER (ORDER BY ProjectID) - 1 AS idx,
        COUNT(*)    OVER ()                        AS cnt
    FROM Projects
),
RankedWorkTypes AS (
    SELECT
        DepartmentID,
        TypeName,
        ROW_NUMBER() OVER (PARTITION BY DepartmentID ORDER BY SortOrder, WorkTypeID) - 1 AS idx,
        COUNT(*)     OVER (PARTITION BY DepartmentID)                                   AS cnt
    FROM WorkTypes
),
LogSeed AS (
    SELECT
        u.UserID,
        u.DepartmentID,
        n.n,
        ((u.UserID * 3 + n.n * 2) % (SELECT cnt FROM RankedProjects WHERE idx = 0)) AS proj_idx,
        ((u.UserID + n.n * 3) %
            (SELECT TOP 1 cnt FROM RankedWorkTypes WHERE DepartmentID = u.DepartmentID)
        ) AS wt_idx,
        DATEADD(DAY, -((u.UserID * 11 + n.n * 17 + 3) % 180), @Today) AS WorkDate,
        CAST(2.0 + ((u.UserID * 2 + n.n * 3) % 13) * 0.5 AS DECIMAL(5,2)) AS WorkHours,
        CASE WHEN (u.UserID + n.n) % 5 = 0 THEN N'社外' ELSE N'社内' END AS WorkLocation,
        CASE WHEN (u.UserID * n.n + 1) % 17 = 0 THEN 1 ELSE 0 END AS IsAfterShipment
    FROM Users AS u
    CROSS JOIN Nums AS n
)
INSERT INTO WorkLogs (ProjectNo, WorkDate, UserID, DepartmentID, ContentName, WorkHours, WorkLocation, IsAfterShipment, Details)
SELECT
    rp.ProjectNo,
    ls.WorkDate,
    ls.UserID,
    ls.DepartmentID,
    rwt.TypeName,
    ls.WorkHours,
    ls.WorkLocation,
    ls.IsAfterShipment,
    NULL
FROM LogSeed AS ls
JOIN RankedProjects AS rp
  ON rp.idx = ls.proj_idx
JOIN RankedWorkTypes AS rwt
  ON rwt.DepartmentID = ls.DepartmentID
 AND rwt.idx          = ls.wt_idx;

DECLARE @LogCount INT = (SELECT COUNT(*) FROM WorkLogs);
PRINT N'サンプル作業記録を投入しました（' + CAST(@LogCount AS NVARCHAR(10)) + N' 件）。';
GO

-- 目標時間（各ユーザー × 注番の組合せごとに 1 件、20〜100h の疑似ランダム）
INSERT INTO ProjectTargets (UserID, ProjectNo, TargetHours)
SELECT DISTINCT
    w.UserID,
    w.ProjectNo,
    CAST(20 + (ABS(CHECKSUM(w.UserID, w.ProjectNo)) % 81) AS DECIMAL(7,2)) AS TargetHours
FROM WorkLogs AS w
WHERE w.ProjectNo IS NOT NULL;

DECLARE @TargetCount INT = (SELECT COUNT(*) FROM ProjectTargets);
PRINT N'サンプル目標時間を投入しました（' + CAST(@TargetCount AS NVARCHAR(10)) + N' 件）。';
GO

-- 目標時間の初回登録を TargetChangeLog にも投入
INSERT INTO TargetChangeLog (TargetID, UserID, ProjectNo, OldValue, NewValue, Reason, ChangedBy)
SELECT pt.TargetID, pt.UserID, pt.ProjectNo, NULL, pt.TargetHours, N'初回登録（サンプルデータ）', pt.UserID
FROM ProjectTargets AS pt;
PRINT N'サンプル目標時間の初回登録履歴を投入しました。';
GO

PRINT N'';
PRINT N'===================================';
PRINT N'  03_sampledata 完了';
PRINT N'===================================';

SELECT
    t.name AS テーブル名,
    p.rows AS 件数
FROM sys.tables     AS t
JOIN sys.partitions AS p
  ON  p.object_id = t.object_id
 AND  p.index_id IN (0, 1)
ORDER BY t.name;
GO
