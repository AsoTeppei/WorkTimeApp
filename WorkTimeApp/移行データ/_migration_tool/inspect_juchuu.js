// 受注資料 (受注４５.xlsx / 受注４６.xlsx) の構造を調べる
const XLSX = require('xlsx');
const path = require('path');

for (const f of ['受注４５.xlsx', '受注４６.xlsx']) {
  console.log(`\n============================================================`);
  console.log(`=== ${f} ===`);
  console.log(`============================================================`);
  const SRC = path.join(__dirname, '..', f);
  const wb = XLSX.readFile(SRC, { cellDates: true });
  console.log('Sheet names:', wb.SheetNames);

  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const ref = ws['!ref'];
    console.log(`\n--- Sheet: ${name} (range: ${ref}) ---`);
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    for (let r = 0; r < Math.min(aoa.length, 12); r++) {
      const row = aoa[r].slice(0, 18).map(v => v === undefined ? '' : String(v).slice(0, 25));
      console.log(`row${r}: `, row);
    }
    console.log(`(total rows: ${aoa.length})`);
  }
}
