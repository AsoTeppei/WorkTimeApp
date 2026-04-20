// 出力された確認用 Excel の「追加予定Projects」シートを検証する。
// - 「特定済み」なのに 客先名 / 件名 が空欄のレコードがあれば一覧で表示する。
const Excel = require('exceljs');
const path  = require('path');

const stamp = new Date();
const ymd = `${stamp.getFullYear()}${String(stamp.getMonth()+1).padStart(2,'0')}${String(stamp.getDate()).padStart(2,'0')}`;
const file = path.join(__dirname, `確認用_DB追加データ_${ymd}.xlsx`);

(async () => {
  const wb = new Excel.Workbook();
  await wb.xlsx.readFile(file);
  const ws = wb.getWorksheet('追加予定Projects');
  const header = ws.getRow(1).values; // 1-based
  const ix = (n) => header.findIndex(v => v === n); // 1-based index
  const cProj = ix('注番'), cClient = ix('客先名'), cSubject = ix('件名'),
        cMatched = ix('受注ファイルでの特定'), cSrc = ix('参照元ファイル');

  const stat = { specified: 0, viaDevice: 0, unspecified: 0 };
  const emptySpecified = [];
  ws.eachRow((row, idx) => {
    if (idx === 1) return;
    const m = row.getCell(cMatched).value;
    const proj = row.getCell(cProj).value;
    const client = row.getCell(cClient).value || '';
    const subject = row.getCell(cSubject).value || '';
    const src = row.getCell(cSrc).value || '';
    if (m === '特定済み' || m === '特定済み(装置No経由)') {
      stat.specified++;
      if (m === '特定済み(装置No経由)') stat.viaDevice++;
      if (!String(client).trim() && !String(subject).trim()) {
        emptySpecified.push({ proj, client, subject, src, m });
      }
    } else {
      stat.unspecified++;
    }
  });

  console.log('合計:', stat);
  console.log('「特定済み」だが客先名・件名どちらも空欄のレコード:', emptySpecified.length);
  emptySpecified.slice(0, 30).forEach(r => console.log(`  ${r.proj} | ${r.m} | client="${r.client}" subject="${r.subject}" src=${r.src}`));
})().catch(e => { console.error(e); process.exit(1); });
