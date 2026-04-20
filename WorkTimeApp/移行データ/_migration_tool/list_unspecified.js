const Excel = require('exceljs');
const path  = require('path');
const stamp = new Date();
const ymd = `${stamp.getFullYear()}${String(stamp.getMonth()+1).padStart(2,'0')}${String(stamp.getDate()).padStart(2,'0')}`;
(async () => {
  const wb = new Excel.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, `確認用_DB追加データ_${ymd}.xlsx`));
  const ws = wb.getWorksheet('追加予定Projects');
  const header = ws.getRow(1).values;
  const ix = (n) => header.findIndex(v => v === n);
  const cProj = ix('注番'), cMatched = ix('受注ファイルでの特定'), cOcc = ix('SP内の出現行数');
  console.log('--- 未特定の注番 ---');
  ws.eachRow((row, i) => {
    if (i === 1) return;
    if (row.getCell(cMatched).value === '未特定(要確認)') {
      console.log(`  ${row.getCell(cProj).value}  (SP内出現: ${row.getCell(cOcc).value})`);
    }
  });
})();
