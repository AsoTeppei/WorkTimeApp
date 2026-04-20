// 移行元 Excel（ｿﾌﾄ作業時間SP.xls）から 2024 年以降のデータを抽出し、
// DB に追加される予定のデータを確認用 Excel に出力する（DB への投入はまだ行わない）。
//
// ルール:
//   1) 対象期間: 2024-01-01 以降
//   2) 担当者マッピング: Excel の「姓のみ」を DB ソフトグループユーザーの氏名と前方一致
//   3) 既存日スキップ: (担当者 × 日付) が DB に既に存在する場合、その組は全スキップ
//   4) 注番なし & 内容詳細が空 → 内容詳細を「内容選択」の値で埋める
//   5) 作業内容が DB の WorkTypes（ソフトグループ）に無い場合は「追加候補」として別シートに出す
//   6) 新規注番 → 受注*.xlsx（一覧表シート）から 顧客名/品名 を引き当てる
//      見つからないものは仮置き「（移行データ・要確認）」。あとから環境設定の注番管理で修正する想定。
//   7) ProjectTargets → 担当者×注番 の実績合計時間を暫定目標として登録する想定（Option B）
//
// 出力: 確認用_DB追加データ_YYYYMMDD.xlsx
//   Sheet1 追加予定作業ログ
//   Sheet2 追加予定Projects（新規注番）
//   Sheet3 追加予定WorkTypes
//   Sheet4 追加予定ProjectTargets（目標=実績合計）
//   Sheet5 担当者未マッピング
//   Sheet6 内容詳細_自動補完ログ
//   Sheet7 サマリ

const XLSX   = require('xlsx');
const Excel  = require('exceljs');
const sql    = require('mssql');
const fs     = require('fs');
const path   = require('path');

// ---------- 設定 ----------
const SRC_SP     = path.join(__dirname, '..', 'ｿﾌﾄ作業時間SP.xls');
const SHEET_SP   = 'ｿﾌﾄ作業時間SP';
const JUCHU_DIR  = path.join(__dirname, '..');          // 受注*.xlsx を探すディレクトリ
const JUCHU_RE   = /^受注.*\.xlsx$/;                    // 全角数字を含む日本語ファイル名を想定
// 受注ファイルは「一覧表」+ 受注種別ごとのシート（例: S44, B44, HF44, BD44）に分かれており、
// 一覧表に出てこない注文番号もサブシートには載っているため、すべてのシートを横断的に走査する
const DEPT_NAME  = 'ソフトグループ';
const DATE_FROM  = '2024-01-01';
const DATA_START = 5;
const COLS = { proj:0, place:1, date:2, hours:3, lodge:4, name:5, content:6, after:7, detail:8 };
const PLACEHOLDER = '（移行データ・要確認）';

const dbConfig = {
  user:     'yonekura',
  password: 'yone6066',
  server:   '192.168.1.8',
  database: 'WorkTimeDB',
  options:  {
    encrypt:                false,
    trustServerCertificate: true,
    instanceName:           'SQLEXPRESS',
    enableArithAbort:       true
  },
  connectionTimeout: 15000,
  requestTimeout:    30000
};

// ---------- ユーティリティ ----------
const toJSTymd = (d) => {
  if (!(d instanceof Date)) return null;
  const t = new Date(d.getTime() + 9*60*60*1000);
  return `${t.getUTCFullYear()}-${String(t.getUTCMonth()+1).padStart(2,'0')}-${String(t.getUTCDate()).padStart(2,'0')}`;
};
// NFKC 正規化で 全角/半角 (英数字・カタカナ・記号) を統一し、空白を除去する。
// 例) "ＳＯＦＴ-０１２" → "SOFT-012"、"ｶﾝﾘ１" → "カンリ1"
const normalizeKey = (s) => {
  if (s == null) return '';
  return String(s).normalize('NFKC').replace(/[\s　\t]/g, '');
};
const normalizeName = normalizeKey;
const matchUserBySurname = (surname, users) => {
  const sn = normalizeKey(surname);
  if (!sn) return null;
  const exact = users.find(u => normalizeKey(u.Name) === sn);
  if (exact) return exact;
  const prefix = users.filter(u => normalizeKey(u.Name).startsWith(sn));
  if (prefix.length === 1) return prefix[0];
  if (prefix.length > 1) return { ambiguous: true, candidates: prefix };
  return null;
};

// 受注*.xlsx を自動収集して 注文番号/装置No → { client, subject, src } マップを構築。
// 各ファイルの全シートを走査する：
//   - 「一覧表」+ 受注種別シート（A43/B43/S44/HF44 等）はいずれも先頭付近に
//     「注文番号 / 装置No / 品名 / 顧客名」を含むヘッダー行を持つ。
//   - 注文番号 と 装置No の両方をキーとして登録する（SP.xls はどちらの形でも書きうる）。
//   - 同じキーに複数のレコードがマッチした場合は「より情報量が多いもの」を優先：
//       スコア = (顧客名あり ? 2 : 0) + (品名あり ? 2 : 0) + (注文番号由来 ? 1 : 0)
//     スコアが新規 >= 既存 なら上書き（同点は新しいファイル/シートが勝つ）。
const loadJuchuuMap = () => {
  const files = fs.readdirSync(JUCHU_DIR).filter(f => JUCHU_RE.test(f)).sort();
  console.log(`  対象: ${files.join(', ') || '(なし)'}`);
  const map = new Map();          // norm(key) -> { client, subject, src, viaDevice }
  const fillScore = (info, viaDevice) =>
    (info.client ? 2 : 0) + (info.subject ? 2 : 0) + (viaDevice ? 0 : 1);
  const setIfBetter = (key, info, viaDevice) => {
    if (!key) return false;
    const cur = map.get(key);
    if (!cur) { map.set(key, { ...info, viaDevice }); return true; }
    if (fillScore(info, viaDevice) >= fillScore(cur, cur.viaDevice)) {
      map.set(key, { ...info, viaDevice });
      return true;
    }
    return false;
  };

  for (const f of files) {
    const src = path.join(JUCHU_DIR, f);
    const wb = XLSX.readFile(src, { cellDates: true });
    let totalProj = 0, totalDev = 0, sheetsOK = 0, sheetsSkip = 0;
    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      if (!ws) { sheetsSkip++; continue; }
      const aoa = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' });
      // ヘッダー行を検出（先頭 10 行以内）。受注本体シートは「注文番号」を持つが、
      // クレーム管理シート（Z44 など）は「クレームNo」「受注番号」「製品番号」など
      // 異なる列名になっているため、いずれかを含む行をヘッダーとみなす。
      // ヘッダーは "品　　　　名" / "顧　客　名" のように全角空白を含むため、
      // 比較前に NFKC + 空白除去で正規化する。
      const includesNorm = (cell, kw) => normalizeKey(cell).includes(kw);
      const PRIMARY_KW   = ['注文番号', 'クレームNo', '受注番号'];
      const SECONDARY_KW = ['装置No', '製品番号'];
      let headerRow = -1;
      for (let r = 0; r < Math.min(aoa.length, 10); r++) {
        if ((aoa[r] || []).some(v => PRIMARY_KW.some(k => includesNorm(v, k)))) { headerRow = r; break; }
      }
      if (headerRow === -1) { sheetsSkip++; continue; }
      const header = aoa[headerRow];
      const findCols = (kws) => kws
        .map(kw => header.findIndex(v => includesNorm(v, kw)))
        .filter(ix => ix >= 0);
      const primaryCols   = findCols(PRIMARY_KW);
      const secondaryCols = findCols(SECONDARY_KW);
      const ixClient  = header.findIndex(v => includesNorm(v, '顧客名'));
      const ixSubject = header.findIndex(v => includesNorm(v, '品名'));
      let nP = 0, nD = 0;
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
          if (k && !seen.has(k)) { setIfBetter(k, info, false); seen.add(k); nP++; }
        }
        for (const ix of secondaryCols) {
          const k = normalizeKey(row[ix]);
          if (k && !seen.has(k)) { setIfBetter(k, info, true); seen.add(k); nD++; }
        }
      }
      sheetsOK++;
      totalProj += nP; totalDev += nD;
    }
    console.log(`  ${f}: 注番=${totalProj}件、装置No=${totalDev}件、対象シート=${sheetsOK}（ヘッダー無=${sheetsSkip}）`);
  }
  return map;
};

// ---------- main ----------
(async () => {
  console.log('=== 確認用 Excel 生成開始 ===');

  console.log(`[1/6] 受注資料を読み込み`);
  const juchuuMap = loadJuchuuMap();
  console.log(`  注文番号ユニーク総数: ${juchuuMap.size}`);

  console.log(`[2/6] 移行元 Excel を読み込み: ${path.basename(SRC_SP)}`);
  const wb = XLSX.readFile(SRC_SP, { cellDates: true });
  const ws = wb.Sheets[SHEET_SP];
  if (!ws) throw new Error(`シート "${SHEET_SP}" が見つかりません`);
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  console.log(`  総行数: ${aoa.length}`);

  console.log(`[3/6] DB 接続: ${dbConfig.server}\\${dbConfig.options.instanceName}/${dbConfig.database}`);
  await sql.connect(dbConfig);

  const usersRes = await sql.query`
    SELECT u.UserID, u.Name, u.Email, u.IsActive
    FROM Users u
    INNER JOIN Departments d ON u.DepartmentID = d.DepartmentID
    WHERE d.DepartmentName = ${DEPT_NAME}
  `;
  const dbUsers = usersRes.recordset;
  console.log(`  ${DEPT_NAME} のユーザー: ${dbUsers.length} 名`);

  const wtRes = await sql.query`
    SELECT wt.TypeName
    FROM WorkTypes wt
    INNER JOIN Departments d ON wt.DepartmentID = d.DepartmentID
    WHERE d.DepartmentName = ${DEPT_NAME}
  `;
  // 半角/全角や空白の差を吸収して比較（NFKC + 空白除去）
  const dbWorkTypeNorms = new Set(wtRes.recordset.map(r => normalizeKey(r.TypeName)));

  const logsRes = await sql.query`
    SELECT DISTINCT wl.UserID, CONVERT(char(10), wl.WorkDate, 23) AS WorkDate
    FROM WorkLogs wl
    INNER JOIN Users u ON wl.UserID = u.UserID
    INNER JOIN Departments d ON u.DepartmentID = d.DepartmentID
    WHERE d.DepartmentName = ${DEPT_NAME}
      AND wl.WorkDate >= ${DATE_FROM}
      AND wl.IsDeleted = 0
  `;
  const existingKeys = new Set(logsRes.recordset.map(r => `${r.UserID}__${r.WorkDate}`));
  console.log(`  既存 (担当者×日付) 組: ${existingKeys.size} 件`);

  const projRes = await sql.query`SELECT ProjectNo FROM Projects`;
  const dbProjects = new Set(projRes.recordset.map(r => normalizeKey(r.ProjectNo)));
  console.log(`  既存 Projects: ${dbProjects.size} 件`);

  await sql.close();

  console.log('[4/6] Excel 行を走査して分類');
  const toInsert       = [];
  const newWorkTypes   = new Map();   // norm -> { name (代表表記), count }
  const unmappedRows   = [];
  const filledDetails  = [];
  const newProjectsMap = new Map(); // ProjectNo -> { client, subject, src, occurrences }
  const targetMap      = new Map(); // `${userId}__${projectNo}` -> { userId, userName, projectNo, hours }

  const surnameStat = new Map();
  const getMapped = (surname) => {
    if (surnameStat.has(surname)) return surnameStat.get(surname);
    const m = matchUserBySurname(surname, dbUsers);
    let entry;
    if (!m) entry = { matched:false, surname, count:0, reason:'一致する DB ユーザーなし' };
    else if (m.ambiguous) entry = { matched:false, surname, count:0, reason:`曖昧（候補: ${m.candidates.map(c=>c.Name).join(' / ')}）` };
    else entry = { matched:true, surname, userId:m.UserID, userName:m.Name, count:0 };
    surnameStat.set(surname, entry);
    return entry;
  };

  let rowsInRange = 0;
  for (let i = DATA_START; i < aoa.length; i++) {
    const row = aoa[i];
    if (!row) continue;
    const d = row[COLS.date];
    if (!(d instanceof Date)) continue;
    const ymd = toJSTymd(d);
    if (!ymd || ymd < DATE_FROM) continue;
    rowsInRange++;

    const surname = row[COLS.name] ? String(row[COLS.name]).trim() : '';
    const mapped  = getMapped(surname);

    const hours    = row[COLS.hours];
    const location = row[COLS.place] ? String(row[COLS.place]).trim() : '社内';
    const content  = row[COLS.content] ? String(row[COLS.content]).trim() : '';
    const afterRaw = row[COLS.after] ? String(row[COLS.after]).trim() : '';
    const isAfter  = afterRaw === '該当';
    const projRaw  = row[COLS.proj];
    const projNo   = (projRaw == null || projRaw === '') ? null : normalizeKey(projRaw);
    let   detail   = row[COLS.detail] ? String(row[COLS.detail]).trim() : '';

    if (!mapped.matched) {
      unmappedRows.push({
        excelRow: i + 1,
        date: ymd, surname, reason: mapped.reason,
        projectNo: projNo, content, hours, detail, location, isAfter
      });
      mapped.count++;
      continue;
    }

    const key = `${mapped.userId}__${ymd}`;
    if (existingKeys.has(key)) { mapped.count++; continue; }

    let detailsFilled = false;
    if (!projNo && !detail) {
      detail = content || '（作業内容未記入）';
      detailsFilled = true;
      filledDetails.push({
        excelRow: i + 1, date: ymd, userId: mapped.userId, userName: mapped.userName,
        originalContent: content, filledDetail: detail
      });
    }

    const contentNorm = normalizeKey(content);
    // 「その他」は作業時間入力 UI に常時候補として出るため、WorkType 追加対象外。
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
          src:       j?.src     || '',
          matched:   !!j,
          viaDevice: !!j?.viaDevice,
          occurrences: 0
        });
      }
      newProjectsMap.get(projNo).occurrences++;
    }

    const hoursNum = hours == null ? 0 : Number(hours);
    if (projNo) {
      const tk = `${mapped.userId}__${projNo}`;
      if (!targetMap.has(tk)) targetMap.set(tk, {
        userId: mapped.userId, userName: mapped.userName, projectNo: projNo, hours: 0, occurrences: 0
      });
      const t = targetMap.get(tk);
      t.hours += hoursNum;
      t.occurrences++;
    }

    toInsert.push({
      excelRow: i + 1,
      date: ymd,
      userId: mapped.userId,
      userName: mapped.userName,
      department: DEPT_NAME,
      projectNo: projNo,
      content,
      hours: hoursNum || null,
      location: (location === '社外') ? '社外' : '社内',
      isAfterShipment: isAfter,
      details: detail,
      detailsFilled,
      isNewWorkType: isNewWT ? 'はい' : '',
      isNewProject:  projNo && !dbProjects.has(projNo) ? 'はい' : '',
    });
    mapped.count++;
  }

  // 新規注番のうち受注資料で引き当てできなかったものはプレースホルダを設定。
  // ※後段で「未特定の注番に紐づく作業ログ・Project・ProjectTargets」は除外する方針。
  for (const p of newProjectsMap.values()) {
    if (!p.matched) {
      p.client  = PLACEHOLDER;
      p.subject = PLACEHOLDER;
    }
  }

  // 受注資料で特定できなかった新規注番（HB0135 / P1697 / 受注前 など）は DB 投入対象から除外する
  const unspecifiedProjSet = new Set(
    [...newProjectsMap.values()].filter(p => !p.matched).map(p => p.projectNo)
  );
  const excludedRows = toInsert.filter(r => r.projectNo && unspecifiedProjSet.has(r.projectNo));
  const toInsertFiltered = toInsert.filter(r => !(r.projectNo && unspecifiedProjSet.has(r.projectNo)));
  for (const k of [...targetMap.keys()]) {
    if (unspecifiedProjSet.has(targetMap.get(k).projectNo)) targetMap.delete(k);
  }
  const newProjectsArrAll  = [...newProjectsMap.values()].sort((a,b) => a.projectNo.localeCompare(b.projectNo));
  const newProjectsArr     = newProjectsArrAll.filter(p => p.matched); // DB 追加対象は特定済みのみ
  const projMatched        = newProjectsArr.length;
  const projViaDevice      = newProjectsArr.filter(p => p.viaDevice).length;
  const projUnmatched      = newProjectsArrAll.length - projMatched;

  console.log(`  走査行（2024-01-01 以降）: ${rowsInRange}`);
  console.log(`  追加予定（除外後）: ${toInsertFiltered.length}（除外: ${excludedRows.length}）`);
  console.log(`  新規 Project（特定済みのみ）: ${projMatched}（うち装置No経由: ${projViaDevice}）`);
  console.log(`  受注資料で未特定（DB対象外）: ${projUnmatched}`);
  console.log(`  新規 WorkType: ${newWorkTypes.size}`);
  console.log(`  ProjectTargets（担当者×注番 組）: ${targetMap.size}`);
  console.log(`  担当者未マッピング行: ${unmappedRows.length}`);
  console.log(`  内容詳細を自動補完した行: ${filledDetails.length}`);

  console.log('[5/6] 確認用 Excel を書き出し');
  const ob = new Excel.Workbook();
  ob.created = new Date();
  const HDR = {
    font: { bold: true, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } },
    alignment: { vertical: 'middle', horizontal: 'center' }
  };
  const hdr2 = (fg) => ({
    font: { bold: true, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: fg } },
    alignment: { vertical: 'middle', horizontal: 'center' }
  });
  const addHeader = (sheet, cols, style = HDR) => {
    sheet.columns = cols;
    sheet.getRow(1).eachCell(c => Object.assign(c, style));
    sheet.getRow(1).height = 22;
    sheet.views = [{ state: 'frozen', ySplit: 1 }];
  };

  // Sheet1: 追加予定作業ログ
  const s1 = ob.addWorksheet('追加予定作業ログ');
  addHeader(s1, [
    { header:'Excel行',       key:'excelRow', width:8 },
    { header:'作業日',         key:'date', width:12 },
    { header:'UserID',         key:'userId', width:8 },
    { header:'担当者',         key:'userName', width:14 },
    { header:'部署',           key:'department', width:14 },
    { header:'注番',           key:'projectNo', width:14 },
    { header:'新規注番?',      key:'isNewProject', width:10 },
    { header:'客先名(参考)',   key:'clientName', width:24 },
    { header:'件名(参考)',     key:'subjectName', width:32 },
    { header:'作業内容',       key:'content', width:16 },
    { header:'新規WorkType?',  key:'isNewWorkType', width:12 },
    { header:'時間(h)',        key:'hours', width:8 },
    { header:'場所',           key:'location', width:8 },
    { header:'出荷後対応',     key:'isAfterShipment', width:10 },
    { header:'内容詳細',       key:'details', width:40 },
    { header:'詳細自動補完?',  key:'detailsFilled', width:12 }
  ]);
  toInsertFiltered.forEach(r => {
    const proj = r.projectNo ? newProjectsMap.get(r.projectNo) : null;
    s1.addRow({
      ...r,
      clientName:  proj ? proj.client  : '',
      subjectName: proj ? proj.subject : '',
      isAfterShipment: r.isAfterShipment ? '該当' : '',
      detailsFilled:   r.detailsFilled ? 'はい' : ''
    });
  });

  // Sheet2: 追加予定 Projects（新規注番、特定済みのみ）
  const s2 = ob.addWorksheet('追加予定Projects');
  addHeader(s2, [
    { header:'注番',              key:'projectNo', width:14 },
    { header:'客先名',            key:'client', width:30 },
    { header:'件名',              key:'subject', width:40 },
    { header:'受注ファイルでの特定', key:'matched', width:20 },
    { header:'参照元ファイル',     key:'src', width:18 },
    { header:'SP内の出現行数',     key:'occurrences', width:12 }
  ], hdr2('FF0F766E'));
  newProjectsArr.forEach(p => s2.addRow({
    ...p,
    matched: p.viaDevice ? '特定済み(装置No経由)' : '特定済み'
  }));

  // Sheet3: 追加予定 WorkTypes
  const s3 = ob.addWorksheet('追加予定WorkTypes');
  addHeader(s3, [
    { header:'追加する作業内容', key:'name', width:24 },
    { header:'部署',             key:'dept', width:14 },
    { header:'出現件数',         key:'count', width:10 }
  ], hdr2('FF9333EA'));
  [...newWorkTypes.values()].sort((a,b)=>b.count-a.count).forEach(v =>
    s3.addRow({ name: v.name, dept: DEPT_NAME, count: v.count })
  );

  // Sheet4: 追加予定 ProjectTargets
  const s4 = ob.addWorksheet('追加予定ProjectTargets');
  addHeader(s4, [
    { header:'UserID',       key:'userId', width:8 },
    { header:'担当者',       key:'userName', width:14 },
    { header:'注番',         key:'projectNo', width:14 },
    { header:'目標時間(h)',  key:'targetHours', width:12 },
    { header:'実績件数',     key:'occurrences', width:10 },
    { header:'備考',         key:'note', width:40 }
  ], hdr2('FFD97706'));
  [...targetMap.values()]
    .sort((a,b) => a.projectNo.localeCompare(b.projectNo) || a.userId - b.userId)
    .forEach(t => s4.addRow({
      userId: t.userId, userName: t.userName, projectNo: t.projectNo,
      targetHours: Number(t.hours.toFixed(2)),
      occurrences: t.occurrences,
      note: '移行時の実績合計を暫定目標として登録（担当者が後から変更可）'
    }));

  // Sheet5: 担当者未マッピング
  const s5 = ob.addWorksheet('担当者未マッピング');
  addHeader(s5, [
    { header:'Excel行',  key:'excelRow', width:8 },
    { header:'作業日',   key:'date', width:12 },
    { header:'担当(姓)', key:'surname', width:10 },
    { header:'理由',     key:'reason', width:40 },
    { header:'注番',     key:'projectNo', width:14 },
    { header:'作業内容', key:'content', width:16 },
    { header:'時間(h)',  key:'hours', width:8 },
    { header:'場所',     key:'location', width:8 },
    { header:'出荷後',   key:'isAfter', width:8 },
    { header:'内容詳細', key:'detail', width:40 }
  ], hdr2('FFB91C1C'));
  unmappedRows.forEach(r => s5.addRow({ ...r, isAfter: r.isAfter ? '該当' : '' }));
  if (unmappedRows.length === 0) s5.addRow({ excelRow: '—', reason: '（未マッピングはありません）' });

  // Sheet5b: 注番未特定で除外した行
  const s5b = ob.addWorksheet('注番未特定_除外');
  addHeader(s5b, [
    { header:'Excel行',     key:'excelRow', width:8 },
    { header:'作業日',      key:'date', width:12 },
    { header:'担当者',      key:'userName', width:14 },
    { header:'注番',        key:'projectNo', width:14 },
    { header:'作業内容',    key:'content', width:16 },
    { header:'時間(h)',     key:'hours', width:8 },
    { header:'場所',        key:'location', width:8 },
    { header:'出荷後',      key:'isAfter', width:8 },
    { header:'内容詳細',    key:'details', width:40 }
  ], hdr2('FFB45309'));
  excludedRows.forEach(r => s5b.addRow({
    ...r,
    isAfter: r.isAfterShipment ? '該当' : ''
  }));
  if (excludedRows.length === 0) s5b.addRow({ excelRow: '—', content: '（除外行はありません）' });

  // Sheet6: 内容詳細_自動補完ログ
  const s6 = ob.addWorksheet('内容詳細_自動補完ログ');
  addHeader(s6, [
    { header:'Excel行',          key:'excelRow', width:8 },
    { header:'作業日',           key:'date', width:12 },
    { header:'UserID',           key:'userId', width:8 },
    { header:'担当者',           key:'userName', width:14 },
    { header:'内容選択',         key:'originalContent', width:16 },
    { header:'補完後の内容詳細', key:'filledDetail', width:30 }
  ], hdr2('FF0EA5E9'));
  filledDetails.forEach(r => s6.addRow(r));
  if (filledDetails.length === 0) s6.addRow({ excelRow: '—', filledDetail: '（自動補完はありません）' });

  // Sheet7: サマリ
  const s7 = ob.addWorksheet('サマリ');
  addHeader(s7, [
    { header:'項目', key:'k', width:50 },
    { header:'値',   key:'v', width:60 }
  ], hdr2('FF334155'));
  s7.addRow({ k:'出力日時', v: new Date().toLocaleString('ja-JP') });
  s7.addRow({ k:'対象期間', v: `${DATE_FROM} 以降` });
  s7.addRow({ k:'走査した Excel 行数（期間内）', v: rowsInRange });
  s7.addRow({ k:'DB に追加予定の作業ログ行数', v: toInsertFiltered.length });
  s7.addRow({ k:'  注番未特定で除外した行数', v: excludedRows.length });
  s7.addRow({ k:'  担当者未マッピングでスキップした行数', v: unmappedRows.length });
  s7.addRow({ k:'  既存日重複でスキップした行数', v: rowsInRange - toInsertFiltered.length - excludedRows.length - unmappedRows.length });
  s7.addRow({ k:'内容詳細を自動補完した行数', v: filledDetails.length });
  s7.addRow({ k:'', v:'' });
  s7.addRow({ k:'追加予定 Project 数（特定済みのみ）', v: projMatched });
  s7.addRow({ k:'  うち 装置No経由で特定', v: projViaDevice });
  s7.addRow({ k:'受注資料で未特定（DB対象外）', v: projUnmatched });
  s7.addRow({ k:'追加予定 WorkType 数', v: newWorkTypes.size });
  s7.addRow({ k:'追加予定 ProjectTargets 組数（担当者×注番）', v: targetMap.size });
  s7.addRow({ k:'', v:'' });
  s7.addRow({ k:'--- 担当者マッピング ---', v:'' });
  for (const [sn, s] of surnameStat) {
    s7.addRow({ k: `  ${sn || '(空欄)'}`, v: s.matched ? `UserID=${s.userId} ${s.userName} (${s.count}行)` : `未マッピング: ${s.reason} (${s.count}行)` });
  }

  const stamp = new Date();
  const ymd = `${stamp.getFullYear()}${String(stamp.getMonth()+1).padStart(2,'0')}${String(stamp.getDate()).padStart(2,'0')}`;
  const outPath = path.join(__dirname, `確認用_DB追加データ_${ymd}.xlsx`);
  await ob.xlsx.writeFile(outPath);
  console.log(`[6/6] 出力完了: ${outPath}`);
  console.log('=== 完了 ===');
})().catch(err => { console.error('ERROR:', err); process.exit(1); });
