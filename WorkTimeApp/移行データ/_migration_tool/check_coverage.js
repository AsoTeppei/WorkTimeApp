// 受注４５.xlsx + 受注４６.xlsx の「一覧表」シートから
// 注文番号 → { 顧客名, 品名 } を抽出して、
// SP.xls の 2024+ 新規注番 217 件のうち何件カバーできるか確認する。

const XLSX = require('xlsx');
const sql  = require('mssql');
const path = require('path');

const SP_SRC = path.join(__dirname, '..', 'ｿﾌﾄ作業時間SP.xls');
const JU_SRCS = [
  path.join(__dirname, '..', '受注４５.xlsx'),
  path.join(__dirname, '..', '受注４６.xlsx'),
];
const DATE_FROM = '2024-01-01';
const DATA_START = 5;
const COL_DATE = 2, COL_PROJ = 0;

const toJSTymd = (d) => {
  if (!(d instanceof Date)) return null;
  const t = new Date(d.getTime() + 9*60*60*1000);
  return `${t.getUTCFullYear()}-${String(t.getUTCMonth()+1).padStart(2,'0')}-${String(t.getUTCDate()).padStart(2,'0')}`;
};

const dbConfig = {
  user: 'yonekura', password: 'yone6066', server: '192.168.1.8', database: 'YonekuraSystemDB',
  options: { encrypt:false, trustServerCertificate:true, instanceName:'SQLEXPRESS', enableArithAbort:true },
  connectionTimeout: 15000, requestTimeout: 30000
};

(async () => {
  // 1) DB の既存注番
  await sql.connect(dbConfig);
  const projRes = await sql.query`SELECT ProjectNo FROM Projects`;
  const dbProjects = new Set(projRes.recordset.map(r => r.ProjectNo));
  await sql.close();

  // 2) SP.xls の 2024+ 新規注番
  const spWb = XLSX.readFile(SP_SRC, { cellDates: true });
  const spAoa = XLSX.utils.sheet_to_json(spWb.Sheets['ｿﾌﾄ作業時間SP'], { header:1, raw:true, defval:null });
  const newProjects = new Set();
  for (let i = DATA_START; i < spAoa.length; i++) {
    const d = spAoa[i][COL_DATE];
    if (!(d instanceof Date)) continue;
    const ymd = toJSTymd(d);
    if (!ymd || ymd < DATE_FROM) continue;
    const p = spAoa[i][COL_PROJ];
    if (p == null || p === '') continue;
    const pn = String(p).trim();
    if (!dbProjects.has(pn)) newProjects.add(pn);
  }
  console.log(`SP.xls 2024+ 新規注番: ${newProjects.size} 件`);

  // 3) 受注資料を読む
  //   row2（1-based=row3）がヘッダー。行2以降がデータ
  const juMap = new Map(); // 注文番号 -> { 顧客名, 品名, src }
  let overwrites = 0;
  for (const src of JU_SRCS) {
    const name = path.basename(src);
    const wb = XLSX.readFile(src, { cellDates: true });
    const ws = wb.Sheets['一覧表'];
    if (!ws) { console.log(`  [警告] ${name} に「一覧表」シートなし`); continue; }
    const aoa = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' });
    // ヘッダー行を探す（"注文番号" を含む行）
    let headerRow = -1;
    for (let r = 0; r < Math.min(aoa.length, 10); r++) {
      if ((aoa[r] || []).some(v => String(v).includes('注文番号'))) { headerRow = r; break; }
    }
    if (headerRow === -1) { console.log(`  [警告] ${name} にヘッダー行が見つからない`); continue; }
    const header = aoa[headerRow];
    const idx = {
      proj: header.findIndex(v => String(v).includes('注文番号')),
      client: header.findIndex(v => String(v).includes('顧客名')),
      subject: header.findIndex(v => String(v).includes('品名')),
    };
    console.log(`  ${name}: headerRow=${headerRow}, proj=${idx.proj}, client=${idx.client}, subject=${idx.subject}`);
    let n = 0;
    for (let r = headerRow + 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const p = String(row[idx.proj] || '').trim();
      if (!p) continue;
      const c = String(row[idx.client]  || '').trim();
      const s = String(row[idx.subject] || '').trim();
      if (juMap.has(p)) overwrites++;
      juMap.set(p, { client:c, subject:s, src:name });
      n++;
    }
    console.log(`  ${name}: 読み込み ${n} 件`);
  }
  console.log(`受注資料 注文番号ユニーク: ${juMap.size} 件（重複上書き: ${overwrites}）`);

  // 4) カバー率
  let covered = 0, coveredEmpty = 0;
  const missing = [];
  const sampleCovered = [];
  for (const pn of newProjects) {
    const e = juMap.get(pn);
    if (!e) { missing.push(pn); continue; }
    covered++;
    if (!e.client && !e.subject) coveredEmpty++;
    if (sampleCovered.length < 10) sampleCovered.push([pn, e.client, e.subject, e.src]);
  }
  console.log(`\n=== カバー率 ===`);
  console.log(`  ${newProjects.size} 件中、受注資料で特定可能: ${covered} 件 (${(covered/newProjects.size*100).toFixed(1)}%)`);
  console.log(`  うち、顧客名・品名が両方空: ${coveredEmpty} 件`);
  console.log(`  未マッチ（受注資料に無い）: ${missing.length} 件`);

  console.log(`\n--- カバーできたサンプル 10 件 ---`);
  sampleCovered.forEach(([p,c,s,src]) => console.log(`  ${p.padEnd(14)} | ${(c||'(空)').padEnd(20)} | ${(s||'(空)').slice(0,30)} | ${src}`));

  if (missing.length > 0) {
    console.log(`\n--- 未マッチ（${missing.length} 件、先頭 30） ---`);
    missing.slice(0, 30).forEach(p => console.log(`  ${p}`));
  }
})().catch(e => { console.error('ERROR:', e); process.exit(1); });
