// 特定の受注ファイル/シート/行の中身を詳しく見る
const XLSX = require('xlsx');
const path = require('path');

const FILE = process.argv[2]; // 例: 受注４４.xlsx
const SHEET = process.argv[3]; // 例: B44
const ROW = parseInt(process.argv[4], 10); // 0-based

const wb = XLSX.readFile(path.join(__dirname, '..', FILE), { cellDates:true });
const ws = wb.Sheets[SHEET];
const aoa = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' });
console.log(`=== ${FILE} / ${SHEET} ===`);
console.log('header(row1):', aoa[1]);
console.log(`target row${ROW}:`, aoa[ROW]);
// 周辺 5 行
for (let r = Math.max(0, ROW-2); r <= Math.min(aoa.length-1, ROW+2); r++) {
  console.log(`row${r}:`, (aoa[r]||[]).slice(0, 10).map(v => String(v).slice(0,20)));
}
