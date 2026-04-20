// 移行元 Excel の構造を調べるための単純ツール
const XLSX = require('xlsx');
const path = require('path');

const SRC = path.join(__dirname, '..', 'ｿﾌﾄ作業時間SP.xls');
const wb = XLSX.readFile(SRC, { cellDates: true });

console.log('=== Sheet names ===');
console.log(wb.SheetNames);

for (const name of wb.SheetNames) {
  const ws = wb.Sheets[name];
  const ref = ws['!ref'];
  console.log(`\n--- Sheet: ${name} (range: ${ref}) ---`);
  // 先頭 8 行 × 12 列くらいを表示
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
  for (let r = 0; r < Math.min(aoa.length, 10); r++) {
    const row = aoa[r].slice(0, 14).map(v => v === undefined ? '' : String(v).slice(0, 25));
    console.log(`row${r}: `, row);
  }
  console.log(`(total rows: ${aoa.length})`);
}
