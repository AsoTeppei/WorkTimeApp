// WorkTimeDB のフルバックアップを SQL Server 側 C:\Backup\ に取得する。
// 出力ファイル名: WorkTimeDB_YYYYMMDD_HHMMSS.bak
//
// リストア手順（万一の場合）:
//   USE master;
//   ALTER DATABASE WorkTimeDB SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
//   RESTORE DATABASE WorkTimeDB
//     FROM DISK = N'C:\Backup\WorkTimeDB_YYYYMMDD_HHMMSS.bak'
//     WITH REPLACE;
//   ALTER DATABASE WorkTimeDB SET MULTI_USER;

const sql = require('mssql');

const dbConfig = {
  user:'yonekura', password:'yone6066', server:'192.168.1.8', database:'master',
  options:{ encrypt:false, trustServerCertificate:true, instanceName:'SQLEXPRESS', enableArithAbort:true },
  connectionTimeout:15000, requestTimeout:600000
};

(async () => {
  const now = new Date();
  const ts = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}${String(now.getSeconds()).padStart(2,'0')}`;

  console.log(`=== WorkTimeDB バックアップ開始 ===`);
  await sql.connect(dbConfig);

  console.log('[1/2] SQL Server の既定バックアップディレクトリを取得');
  // SERVERPROPERTY が一番手軽だが、古い版では NULL なのでフォールバックを順に試す
  const dirRes = await sql.query(`
    DECLARE @p NVARCHAR(4000);
    SET @p = CAST(SERVERPROPERTY('InstanceDefaultBackupPath') AS NVARCHAR(4000));
    IF @p IS NULL
    BEGIN
      BEGIN TRY
        EXEC master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',
             N'Software\\Microsoft\\MSSQLServer\\MSSQLServer', N'BackupDirectory', @p OUTPUT;
      END TRY BEGIN CATCH END CATCH;
    END
    -- 最終フォールバック: master データファイルのディレクトリ
    IF @p IS NULL
      SELECT @p = LEFT(physical_name, LEN(physical_name) - CHARINDEX('\\', REVERSE(physical_name)) + 1)
      FROM sys.master_files WHERE database_id = 1 AND file_id = 1;
    SELECT @p AS BackupDirectory;
  `);
  let dir = dirRes.recordset[0]?.BackupDirectory;
  if (!dir) throw new Error('既定バックアップディレクトリが取得できませんでした');
  if (!dir.endsWith('\\')) dir += '\\';
  const filePath = `${dir}WorkTimeDB_${ts}.bak`;
  console.log(`  既定パス: ${dir}`);
  console.log(`  出力先: ${filePath}`);

  console.log('[2/2] BACKUP DATABASE 実行中（数分かかる場合あり）...');
  const start = Date.now();
  await new sql.Request()
    .input('path', sql.NVarChar, filePath)
    .query(`
      BACKUP DATABASE WorkTimeDB
        TO DISK = @path
        WITH FORMAT, INIT, NAME = N'WorkTimeDB-Full Backup', STATS = 25;
    `);
  const sec = ((Date.now() - start)/1000).toFixed(1);
  console.log(`完了 (${sec} 秒)`);
  console.log('');
  console.log(`★ バックアップファイル (192.168.1.8 上): ${filePath}`);
  console.log(`  リストア時はこのファイルから復元してください。`);

  await sql.close();
})().catch(e => { console.error('FATAL:', e.message || e); process.exit(1); });
