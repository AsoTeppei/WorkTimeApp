// 未特定の注番と受注資料の注番を文字単位で比較し、
// 残っている表記差（ハイフン種別など）を炙り出す。
const XLSX = require('xlsx');
const sql  = require('mssql');
const fs   = require('fs');
const path = require('path');

const SRC_SP    = path.join(__dirname, '..', 'ｿﾌﾄ作業時間SP.xls');
const JUCHU_DIR = path.join(__dirname, '..');
const JUCHU_RE  = /^受注.*\.xlsx$/;
const SHEET_SP  = 'ｿﾌﾄ作業時間SP';
const DATA_START = 5;
const COL_DATE = 2, COL_PROJ = 0;
const DATE_FROM = '2024-01-01';

const dbConfig = {
  user:'yonekura', password:'yone6066', server:'192.168.1.8', database:'YonekuraSystemDB',
  options:{ encrypt:false, trustServerCertificate:true, instanceName:'SQLEXPRESS', enableArithAbort:true },
  connectionTimeout:15000, requestTimeout:30000
};

const normalizeKey = (s) => s == null ? '' : String(s).normalize('NFKC').replace(/[\s　\t]/g, '');
const toJST = (d) => {
  if (!(d instanceof Date)) return null;
  const t = new Date(d.getTime() + 9*60*60*1000);
  return `${t.getUTCFullYear()}-${String(t.getUTCMonth()+1).padStart(2,'0')}-${String(t.getUTCDate()).padStart(2,'0')}`;
};
const cps = (s) => [...s].map(c => c + '(U+' + c.codePointAt(0).toString(16).toUpperCase().padStart(4,'0') + ')').join(' ');

(async () => {
  // 1) DB 既存
  await sql.connect(dbConfig);
  const projRes = await sql.query`SELECT ProjectNo FROM Projects`;
  const dbProjects = new Set(projRes.recordset.map(r => normalizeKey(r.ProjectNo)));
  await sql.close();

  // 2) SP の 2024+ 新規注番（normalized）と元表記
  const wb = XLSX.readFile(SRC_SP, { cellDates: true });
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[SHEET_SP], { header:1, raw:true, defval:null });
  const newProjects = new Map(); // norm -> raw
  for (let i = DATA_START; i < aoa.length; i++) {
    const d = aoa[i][COL_DATE];
    if (!(d instanceof Date)) continue;
    const ymd = toJST(d);
    if (!ymd || ymd < DATE_FROM) continue;
    const p = aoa[i][COL_PROJ];
    if (p == null || p === '') continue;
    const norm = normalizeKey(p);
    if (!norm) continue;
    if (dbProjects.has(norm)) continue;
    if (!newProjects.has(norm)) newProjects.set(norm, String(p));
  }

  // 3) 受注の注番（normalized）と元表記、ファイル別に
  const juchuuByNorm = new Map();   // norm -> [ {raw, file} ]
  const files = fs.readdirSync(JUCHU_DIR).filter(f => JUCHU_RE.test(f)).sort();
  for (const f of files) {
    const wb2 = XLSX.readFile(path.join(JUCHU_DIR, f), { cellDates:true });
    const ws = wb2.Sheets['一覧表'];
    if (!ws) continue;
    const aoa2 = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' });
    let hr = -1;
    for (let r = 0; r < Math.min(aoa2.length, 10); r++) {
      if ((aoa2[r] || []).some(v => String(v).includes('注文番号'))) { hr = r; break; }
    }
    if (hr === -1) continue;
    const ix = aoa2[hr].findIndex(v => String(v).includes('注文番号'));
    for (let r = hr + 1; r < aoa2.length; r++) {
      const raw = String((aoa2[r]||[])[ix] || '');
      const n = normalizeKey(raw);
      if (!n) continue;
      if (!juchuuByNorm.has(n)) juchuuByNorm.set(n, []);
      juchuuByNorm.get(n).push({ raw, file:f });
    }
  }

  // 4) SP 未特定 (norm が juchuu に無い) のサンプル + 「ハイフン違い等で近い候補があるか」もチェック
  const unmatched = [];
  for (const [norm, raw] of newProjects) {
    if (!juchuuByNorm.has(norm)) unmatched.push({ norm, raw });
  }
  console.log(`SP 2024+ 新規注番: ${newProjects.size} 件、 受注で未特定: ${unmatched.length} 件`);

  // 「ハイフン類だけを除いた英数字列」で近似マッチを試す
  const stripHyphenNorm = (s) => normalizeKey(s).replace(/[\-‐‑‒–—―ー－−ｰ_／/\s]/g, '');
  const juchuuByLoose = new Map();   // loose -> [ {norm, raw, file} ]
  for (const [norm, list] of juchuuByNorm) {
    const loose = stripHyphenNorm(norm);
    if (!juchuuByLoose.has(loose)) juchuuByLoose.set(loose, []);
    list.forEach(x => juchuuByLoose.get(loose).push({ norm, ...x }));
  }

  let nLooseMatch = 0;
  console.log(`\n--- 未特定 注番のうち「区切り文字を緩めれば一致するもの」---`);
  for (const u of unmatched) {
    const loose = stripHyphenNorm(u.norm);
    const cand = juchuuByLoose.get(loose) || [];
    if (cand.length > 0) {
      nLooseMatch++;
      if (nLooseMatch <= 20) {
        console.log(`  SP "${u.raw}" (norm="${u.norm}")`);
        cand.slice(0,3).forEach(c => console.log(`     ←→ 受注 "${c.raw}" (norm="${c.norm}", ${c.file})`));
      }
    }
  }
  console.log(`\n  区切り文字を緩めれば一致: ${nLooseMatch} / ${unmatched.length}`);

  console.log(`\n--- 完全に該当なし 注番 サンプル（先頭 20、文字コード付き）---`);
  let shown = 0;
  for (const u of unmatched) {
    const loose = stripHyphenNorm(u.norm);
    if ((juchuuByLoose.get(loose) || []).length > 0) continue;
    console.log(`  raw="${u.raw}"  norm="${u.norm}"  cps=[${cps(u.raw)}]`);
    if (++shown >= 20) break;
  }

  // 5) 担当者未マッピングの確認
  console.log(`\n--- 担当者未マッピング (姓集計) ---`);
  const surMap = new Map();
  for (let i = DATA_START; i < aoa.length; i++) {
    const d = aoa[i][COL_DATE];
    if (!(d instanceof Date)) continue;
    const ymd = toJST(d);
    if (!ymd || ymd < DATE_FROM) continue;
    const surnameRaw = aoa[i][5] == null ? '' : String(aoa[i][5]);
    const sn = normalizeKey(surnameRaw);
    const k = `${sn}__${surnameRaw}`;
    if (!surMap.has(k)) surMap.set(k, { surnameRaw, sn, count:0 });
    surMap.get(k).count++;
  }
  // DB
  await sql.connect(dbConfig);
  const u = await sql.query`
    SELECT u.UserID, u.Name FROM Users u
    INNER JOIN Departments d ON u.DepartmentID = d.DepartmentID
    WHERE d.DepartmentName = N'ソフトグループ'`;
  await sql.close();
  const dbUsers = u.recordset.map(r => ({ UserID:r.UserID, Name:r.Name, norm:normalizeKey(r.Name) }));
  console.log('  DB ユーザー:', dbUsers.map(x => `${x.Name}(norm=${x.norm})`).join(', '));
  for (const v of surMap.values()) {
    const exact = dbUsers.find(x => x.norm === v.sn);
    const prefix = dbUsers.filter(x => x.norm.startsWith(v.sn));
    let note;
    if (exact) note = `→ ${exact.Name} 完全一致`;
    else if (prefix.length === 1) note = `→ ${prefix[0].Name} 前方一致`;
    else if (prefix.length > 1) note = `→ 曖昧 ${prefix.map(x=>x.Name).join('/')}`;
    else note = '→ 該当なし';
    console.log(`  raw="${v.surnameRaw}" sn="${v.sn}" cps=[${cps(v.surnameRaw)}] cnt=${v.count} ${note}`);
  }
})().catch(e => { console.error('ERROR:', e); process.exit(1); });
