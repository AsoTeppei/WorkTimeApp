// 受注*.xlsx の全シートを一覧して、注文番号に相当する列があるかチェックする。
const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');

const DIR = path.join(__dirname, '..');
const files = fs.readdirSync(DIR).filter(f => /^受注.*\.xlsx$/.test(f)).sort();

const TARGET = process.argv[2]; // 任意: 特定の注番を渡すと探す

for (const f of files) {
  console.log(`\n========================================`);
  console.log(`=== ${f} ===`);
  console.log(`========================================`);
  const wb = XLSX.readFile(path.join(DIR, f), { cellDates: true });
  console.log('Sheets:', wb.SheetNames);
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const ref = ws['!ref'] || '(empty)';
    const aoa = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' });
    let hr = -1;
    for (let r = 0; r < Math.min(aoa.length, 10); r++) {
      if ((aoa[r] || []).some(v => String(v).includes('注文番号'))) { hr = r; break; }
    }
    let hasOrderCol = hr !== -1;
    let dataRows = 0;
    if (hasOrderCol) {
      const ix = aoa[hr].findIndex(v => String(v).includes('注文番号'));
      for (let r = hr + 1; r < aoa.length; r++) {
        const v = (aoa[r] || [])[ix];
        if (v != null && String(v).trim() !== '') dataRows++;
      }
    }
    console.log(`  [${name}] range=${ref}, rows=${aoa.length}, headerRow=${hr}, データ行(注番列)=${dataRows}`);
    // 先頭5行を表示
    if (aoa.length > 0) {
      for (let r = 0; r < Math.min(aoa.length, 4); r++) {
        const row = (aoa[r] || []).slice(0, 12).map(v => String(v).slice(0,15));
        console.log(`     row${r}: `, row);
      }
    }
    // TARGET 注番探し
    if (TARGET) {
      const norm = String(TARGET).normalize('NFKC').replace(/[\s　\t]/g, '');
      let found = 0;
      for (let r = 0; r < aoa.length; r++) {
        for (let c = 0; c < (aoa[r]||[]).length; c++) {
          const v = String((aoa[r]||[])[c] || '').normalize('NFKC').replace(/[\s　\t]/g,'');
          if (v === norm) {
            found++;
            if (found <= 3) console.log(`     ★ TARGET "${TARGET}" 発見: row=${r}, col=${c}`);
          }
        }
      }
      if (found > 0) console.log(`     ★ 合計 ${found} 箇所`);
    }
  }
}
