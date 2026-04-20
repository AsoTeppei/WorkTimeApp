// 移行元 Excel（ｿﾌﾄ作業時間SP.xls）2024-01-01 以降のデータを WorkTimeDB に投入する。
//
// 使い方:
//   node migrate_to_db.js              ← dry-run（トランザクションを ROLLBACK して結果のみ確認）
//   node migrate_to_db.js --commit     ← 実投入（COMMIT）
//
// 投入順:
//   1. Projects (新規 214 件)
//   2. WorkTypes (新規 6 件、ソフトグループ配下)
//   3. ProjectTargets (担当者×注番、目標=実績合計)
//   4. WorkLogs (約 1,901 件)
//
// データソース・ロジックは export_preview.js と完全一致。
// 出力した「確認用_DB追加データ_YYYYMMDD.xlsx」と件数が一致することを確認している。
//
// 前提:
//   - 担当者は SP の「姓のみ」を DB ソフトグループユーザーの氏名と前方一致
//   - (担当者×日付) が DB に既に存在する場合はスキップ
//   - 注番なし & 内容詳細が空 → 内容詳細を「内容選択」の値で埋める
//   - 半角/全角は NFKC で正規化して同一視
//   - 受注資料で特定できなかった注番に紐づく行は除外

const XLSX  = require('xlsx');
const sql   = require('mssql');
const fs    = require('fs');
const path  = require('path');

const COMMIT = process.argv.includes('--commit');

// ---------- 設定 ----------
const SRC_SP    = path.join(__dirname, '..', 'ｿﾌﾄ作業時間SP.xls');
const SHEET_SP  = 'ｿﾌﾄ作業時間SP';
const JUCHU_DIR = path.join(__dirname, '..');
const JUCHU_RE  = /^受注.*\.xlsx$/;
const DEPT_NAME = 'ソフトグループ';
const DATE_FROM = '2024-01-01';
const DATA_START = 5;
const COLS = { proj:0, place:1, date:2, hours:3, lodge:4, name:5, content:6, after:7, detail:8 };

const dbConfig = {
  user:'yonekura', password:'yone6066', server:'192.168.1.8', database:'WorkTimeDB',
  options:{ encrypt:false, trustServerCertificate:true, instanceName:'SQLEXPRESS', enableArithAbort:true },
  connectionTimeout:15000, requestTimeout:120000
};

// ---------- ユーティリティ ----------
const toJSTymd = (d) => {
  if (!(d instanceof Date)) return null;
  const t = new Date(d.getTime() + 9*60*60*1000);
  return `${t.getUTCFullYear()}-${String(t.getUTCMonth()+1).padStart(2,'0')}-${String(t.getUTCDate()).padStart(2,'0')}`;
};
const normalizeKey = (s) => s == null ? '' : String(s).normalize('NFKC').replace(/[\s　\t]/g, '');
const matchUserBySurname = (surname, users) => {
  const sn = normalizeKey(surname);
  if (!sn) return null;
  const exact = users.find(u => normalizeKey(u.Name) === sn);
  if (exact) return exact;
  const prefix = users.filter(u => normalizeKey(u.Name).startsWith(sn));
  if (prefix.length === 1) return prefix[0];
  if (prefix.length > 1) return { ambiguous:true, candidates:prefix };
  return null;
};

const loadJuchuuMap = () => {
  const files = fs.readdirSync(JUCHU_DIR).filter(f => JUCHU_RE.test(f)).sort();
  const map = new Map();
  const fillScore = (info, viaDevice) =>
    (info.client ? 2 : 0) + (info.subject ? 2 : 0) + (viaDevice ? 0 : 1);
  const setIfBetter = (key, info, viaDevice) => {
    if (!key) return;
    const cur = map.get(key);
    if (!cur) { map.set(key, { ...info, viaDevice }); return; }
    if (fillScore(info, viaDevice) >= fillScore(cur, cur.viaDevice)) {
      map.set(key, { ...info, viaDevice });
    }
  };
  for (const f of files) {
    const wb = XLSX.readFile(path.join(JUCHU_DIR, f), { cellDates: true });
    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      if (!ws) continue;
      const aoa = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' });
      const includesNorm = (cell, kw) => normalizeKey(cell).includes(kw);
      const PRIMARY_KW   = ['注文番号', 'クレームNo', '受注番号'];
      const SECONDARY_KW = ['装置No', '製品番号'];
      let headerRow = -1;
      for (let r = 0; r < Math.min(aoa.length, 10); r++) {
        if ((aoa[r] || []).some(v => PRIMARY_KW.some(k => includesNorm(v, k)))) { headerRow = r; break; }
      }
      if (headerRow === -1) continue;
      const header = aoa[headerRow];
      const findCols = (kws) => kws.map(kw => header.findIndex(v => includesNorm(v, kw))).filter(ix => ix >= 0);
      const primaryCols   = findCols(PRIMARY_KW);
      const secondaryCols = findCols(SECONDARY_KW);
      const ixClient  = header.findIndex(v => includesNorm(v, '顧客名'));
      const ixSubject = header.findIndex(v => includesNorm(v, '品名'));
      for (let r = headerRow + 1; r < aoa.length; r++) {
        const row = aoa[r] || [];
        const info = {
          client:  ixClient  >= 0 ? String(row[ixClient]  || '').trim() : '',
          subject: ixSubject >= 0 ? String(row[ixSubject] || '').trim() : '',
          src:     `${f}/${sheetName}`,
        };
        const seen = new Set();
        for (const ix of primaryCols) {
          const k = normalizeKey(row[ix]);
          if (k && !seen.has(k)) { setIfBetter(k, info, false); seen.add(k); }
        }
        for (const ix of secondaryCols) {
          const k = normalizeKey(row[ix]);
          if (k && !seen.has(k)) { setIfBetter(k, info, true); seen.add(k); }
        }
      }
    }
  }
  return map;
};

// ---------- データ構築 ----------
async function buildData() {
  console.log('[1/5] 受注資料を読み込み');
  const juchuuMap = loadJuchuuMap();
  console.log(`  注文番号ユニーク総数: ${juchuuMap.size}`);

  console.log('[2/5] 移行元 Excel を読み込み');
  const wb = XLSX.readFile(SRC_SP, { cellDates: true });
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[SHEET_SP], { header:1, raw:true, defval:null });

  console.log('[3/5] DB 接続して既存データを取得');
  await sql.connect(dbConfig);

  const deptRes = await sql.query`SELECT DepartmentID FROM Departments WHERE DepartmentName = ${DEPT_NAME}`;
  if (deptRes.recordset.length === 0) throw new Error(`部署 "${DEPT_NAME}" が存在しません`);
  const deptId = deptRes.recordset[0].DepartmentID;

  const usersRes = await sql.query`
    SELECT UserID, Name FROM Users WHERE DepartmentID = ${deptId}
  `;
  const dbUsers = usersRes.recordset;

  const wtRes = await sql.query`
    SELECT TypeName FROM WorkTypes WHERE DepartmentID = ${deptId}
  `;
  const dbWorkTypeNorms = new Set(wtRes.recordset.map(r => normalizeKey(r.TypeName)));

  const logsRes = await sql.query`
    SELECT DISTINCT wl.UserID, CONVERT(char(10), wl.WorkDate, 23) AS WorkDate
    FROM WorkLogs wl
    INNER JOIN Users u ON wl.UserID = u.UserID
    WHERE u.DepartmentID = ${deptId}
      AND wl.WorkDate >= ${DATE_FROM}
      AND wl.IsDeleted = 0
  `;
  const existingKeys = new Set(logsRes.recordset.map(r => `${r.UserID}__${r.WorkDate}`));

  const projRes = await sql.query`SELECT ProjectNo FROM Projects`;
  const dbProjects = new Set(projRes.recordset.map(r => normalizeKey(r.ProjectNo)));

  const ptRes = await sql.query`SELECT UserID, ProjectNo FROM ProjectTargets`;
  const existingTargets = new Set(ptRes.recordset.map(r => `${r.UserID}__${normalizeKey(r.ProjectNo)}`));

  console.log(`  部署 ${DEPT_NAME} (ID=${deptId}), ユーザー ${dbUsers.length}名, 既存日 ${existingKeys.size}組, 既存Project ${dbProjects.size}件, 既存ProjectTargets ${existingTargets.size}組`);

  console.log('[4/5] 行を分類');
  const toInsert       = [];
  const newWorkTypes   = new Map();
  const newProjectsMap = new Map();
  const targetMap      = new Map();
  let unmappedCount = 0, existingDayCount = 0;

  for (let i = DATA_START; i < aoa.length; i++) {
    const row = aoa[i];
    if (!row) continue;
    const d = row[COLS.date];
    if (!(d instanceof Date)) continue;
    const ymd = toJSTymd(d);
    if (!ymd || ymd < DATE_FROM) continue;

    const surname = row[COLS.name] ? String(row[COLS.name]).trim() : '';
    const m = matchUserBySurname(surname, dbUsers);
    if (!m || m.ambiguous) { unmappedCount++; continue; }

    const key = `${m.UserID}__${ymd}`;
    if (existingKeys.has(key)) { existingDayCount++; continue; }

    const hours    = row[COLS.hours];
    const location = row[COLS.place] ? String(row[COLS.place]).trim() : '社内';
    const content  = row[COLS.content] ? String(row[COLS.content]).trim() : '';
    const afterRaw = row[COLS.after] ? String(row[COLS.after]).trim() : '';
    const isAfter  = afterRaw === '該当';
    const projRaw  = row[COLS.proj];
    const projNo   = (projRaw == null || projRaw === '') ? null : normalizeKey(projRaw);
    let   detail   = row[COLS.detail] ? String(row[COLS.detail]).trim() : '';

    if (!projNo && !detail) detail = content || '（作業内容未記入）';

    const contentNorm = normalizeKey(content);
    const isReservedWT = contentNorm === 'その他';
    const isNewWT = !!content && !isReservedWT && !dbWorkTypeNorms.has(contentNorm);
    if (isNewWT) {
      if (!newWorkTypes.has(contentNorm)) newWorkTypes.set(contentNorm, { name: content, count: 0 });
      newWorkTypes.get(contentNorm).count++;
    }

    if (projNo && !dbProjects.has(projNo)) {
      if (!newProjectsMap.has(projNo)) {
        const j = juchuuMap.get(projNo);
        newProjectsMap.set(projNo, {
          projectNo: projNo,
          client:    j?.client  || '',
          subject:   j?.subject || '',
          matched:   !!j,
          viaDevice: !!j?.viaDevice,
        });
      }
    }

    const hoursNum = hours == null ? 0 : Number(hours);
    if (projNo && hoursNum > 0) {
      const tk = `${m.UserID}__${projNo}`;
      if (!targetMap.has(tk)) targetMap.set(tk, { userId: m.UserID, projectNo: projNo, hours: 0 });
      targetMap.get(tk).hours += hoursNum;
    }

    toInsert.push({
      excelRow: i + 1,
      date: ymd,
      userId: m.UserID,
      projectNo: projNo,
      content: content || '（作業内容未記入）',
      hours: hoursNum,
      location: (location === '社外') ? '社外' : '社内',
      isAfterShipment: isAfter,
      details: detail,
    });
  }

  // 注番未特定の行は除外
  const unspecifiedProjSet = new Set(
    [...newProjectsMap.values()].filter(p => !p.matched).map(p => p.projectNo)
  );
  const excludedByProj = toInsert.filter(r => r.projectNo && unspecifiedProjSet.has(r.projectNo));
  const toInsertFiltered = toInsert.filter(r => !(r.projectNo && unspecifiedProjSet.has(r.projectNo)));
  for (const k of [...targetMap.keys()]) {
    if (unspecifiedProjSet.has(targetMap.get(k).projectNo)) targetMap.delete(k);
  }

  // WorkLogs 制約 (WorkHours > 0 AND <= 24) を満たさない行を除外
  const invalidHours = toInsertFiltered.filter(r => !(r.hours > 0 && r.hours <= 24));
  const finalToInsert = toInsertFiltered.filter(r => r.hours > 0 && r.hours <= 24);

  // ProjectTargets: TargetHours > 0、かつ DB に既存 (UserID, ProjectNo) は除外
  let skippedTargetsExisting = 0;
  const finalTargets = [...targetMap.values()].filter(t => {
    if (!(t.hours > 0)) return false;
    if (existingTargets.has(`${t.userId}__${t.projectNo}`)) { skippedTargetsExisting++; return false; }
    return true;
  });

  // Projects: 特定済みのみ
  const finalProjects = [...newProjectsMap.values()].filter(p => p.matched);

  // WorkTypes: 全件
  const finalWorkTypes = [...newWorkTypes.values()];

  console.log(`  走査: 担当者未マッピング=${unmappedCount}, 既存日重複=${existingDayCount}`);
  console.log(`  注番未特定で除外: ${excludedByProj.length} 行`);
  console.log(`  WorkHours 不正で除外: ${invalidHours.length} 行`);
  console.log(`  ProjectTargets 既存重複で除外: ${skippedTargetsExisting} 組`);
  if (invalidHours.length > 0) {
    invalidHours.slice(0, 5).forEach(r => console.log(`    Excel行 ${r.excelRow} hours=${r.hours}`));
    if (invalidHours.length > 5) console.log(`    ...他 ${invalidHours.length - 5} 件`);
  }
  console.log('');
  console.log(`  ★ DB 投入対象`);
  console.log(`    Projects:        ${finalProjects.length} 件`);
  console.log(`    WorkTypes:       ${finalWorkTypes.length} 件`);
  console.log(`    ProjectTargets:  ${finalTargets.length} 組`);
  console.log(`    WorkLogs:        ${finalToInsert.length} 行`);

  return { deptId, finalProjects, finalWorkTypes, finalTargets, finalToInsert };
}

// ---------- DB 投入 ----------
async function insertAll(transaction, data) {
  const { deptId, finalProjects, finalWorkTypes, finalTargets, finalToInsert } = data;

  // 1. Projects
  console.log('  [Projects] 挿入中...');
  for (const p of finalProjects) {
    await new sql.Request(transaction)
      .input('projectNo',  sql.NVarChar(50),  p.projectNo)
      .input('clientName', sql.NVarChar(100), (p.client  || '').slice(0, 100))
      .input('subject',    sql.NVarChar(200), (p.subject || '').slice(0, 200))
      .query`INSERT INTO Projects (ProjectNo, ClientName, Subject) VALUES (@projectNo, @clientName, @subject)`;
  }
  console.log(`    ${finalProjects.length} 件 挿入完了`);

  // 2. WorkTypes
  console.log('  [WorkTypes] 挿入中...');
  for (const w of finalWorkTypes) {
    await new sql.Request(transaction)
      .input('deptId',   sql.Int,           deptId)
      .input('typeName', sql.NVarChar(100), (w.name || '').slice(0, 100))
      .query`INSERT INTO WorkTypes (DepartmentID, TypeName) VALUES (@deptId, @typeName)`;
  }
  console.log(`    ${finalWorkTypes.length} 件 挿入完了`);

  // 3. ProjectTargets
  console.log('  [ProjectTargets] 挿入中...');
  for (const t of finalTargets) {
    await new sql.Request(transaction)
      .input('userId',    sql.Int,          t.userId)
      .input('projectNo', sql.NVarChar(50), t.projectNo)
      .input('hours',     sql.Decimal(7,2), Number(t.hours.toFixed(2)))
      .query`INSERT INTO ProjectTargets (UserID, ProjectNo, TargetHours) VALUES (@userId, @projectNo, @hours)`;
  }
  console.log(`    ${finalTargets.length} 件 挿入完了`);

  // 4. WorkLogs
  console.log('  [WorkLogs] 挿入中...');
  let n = 0;
  for (const r of finalToInsert) {
    await new sql.Request(transaction)
      .input('projectNo',       sql.NVarChar(50),  r.projectNo)
      .input('workDate',        sql.Date,          r.date)
      .input('userId',          sql.Int,           r.userId)
      .input('deptId',          sql.Int,           deptId)
      .input('contentName',     sql.NVarChar(100), r.content.slice(0, 100))
      .input('workHours',       sql.Decimal(5,2),  Number(r.hours.toFixed(2)))
      .input('workLocation',    sql.NVarChar(10),  r.location)
      .input('isAfterShipment', sql.Bit,           r.isAfterShipment ? 1 : 0)
      .input('details',         sql.NVarChar(sql.MAX), r.details || null)
      .query`
        INSERT INTO WorkLogs (ProjectNo, WorkDate, UserID, DepartmentID, ContentName, WorkHours, WorkLocation, IsAfterShipment, Details)
        VALUES (@projectNo, @workDate, @userId, @deptId, @contentName, @workHours, @workLocation, @isAfterShipment, @details)
      `;
    n++;
    if (n % 200 === 0) console.log(`    ${n} / ${finalToInsert.length} 件...`);
  }
  console.log(`    ${finalToInsert.length} 件 挿入完了`);
}

// ---------- main ----------
(async () => {
  console.log(`=== DB 投入開始 (${COMMIT ? 'COMMIT モード' : 'DRY-RUN モード（最後に ROLLBACK）'}) ===`);
  const data = await buildData();

  console.log('');
  console.log('[5/5] トランザクション開始');
  const transaction = new sql.Transaction();
  await transaction.begin();
  try {
    await insertAll(transaction, data);
    if (COMMIT) {
      await transaction.commit();
      console.log('=== COMMIT しました ===');
    } else {
      await transaction.rollback();
      console.log('=== DRY-RUN なので ROLLBACK しました ===');
      console.log('=== 本投入する場合は: node migrate_to_db.js --commit ===');
    }
  } catch (e) {
    console.error('!!! エラー、ROLLBACK します:', e.message || e);
    try { await transaction.rollback(); } catch (_) {}
    throw e;
  } finally {
    await sql.close();
  }
})().catch(e => { console.error('FATAL:', e); process.exit(1); });
