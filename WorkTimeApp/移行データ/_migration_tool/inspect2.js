// 年がどう記録されているか調査する。
// row3 が「自由記入/入力選択...」、row4 が見出し、row5 以降がデータ。
// 日付セルは「05/01」のように見えるが、Excel 上は serial number か文字列か？
const XLSX = require('xlsx');
const path = require('path');
const SRC = path.join(__dirname, '..', 'ｿﾌﾄ作業時間SP.xls');
const wb = XLSX.readFile(SRC, { cellDates: true });
const ws = wb.Sheets['ｿﾌﾄ作業時間SP'];

// セル C5（最初のデータ行の作業日）の生 cell を見てみる
function inspectCell(addr) {
  const c = ws[addr];
  if (!c) { console.log(addr, '(empty)'); return; }
  console.log(addr, JSON.stringify({ t: c.t, v: c.v, w: c.w, z: c.z }));
}

console.log('=== ヘッダー周辺 ===');
['A1','B1','C1','D1','E1','I1','J1','N1'].forEach(inspectCell);
console.log('\n=== 最初のデータ行 (row 6 in 1-based = row5 in 0-based) ===');
['A6','B6','C6','D6','E6','F6','G6','H6','I6'].forEach(inspectCell);
console.log('\n=== 何行か飛び飛びで日付セル ===');
[10,100,500,1000,2000,5000,10000,10028].forEach(r => inspectCell(`C${r}`));
console.log('\n=== 1, 2列目 (注番,場所) も同様に ===');
[10,100,500,1000,2000,5000,10000,10028].forEach(r => {
  inspectCell(`A${r}`); inspectCell(`B${r}`); inspectCell(`F${r}`);
});

// AoA 全体を読み、日付が「年が見える」形になっている行を探す
const aoa = XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:'' });
console.log('\n=== 行データの型サマリ（先頭30件） ===');
for (let i = 0; i < Math.min(aoa.length, 30); i++) {
  const c = aoa[i][2]; // C 列
  if (c === '' || c == null) continue;
  console.log(`row${i}: C=`, typeof c, JSON.stringify(c));
}

// 日付列が Date オブジェクトかどうかを統計
let nDate=0, nNumber=0, nString=0, nOther=0, nEmpty=0;
let yearSet = new Set();
for (let i = 5; i < aoa.length; i++) {
  const c = aoa[i][2];
  if (c === '' || c == null) { nEmpty++; continue; }
  if (c instanceof Date) { nDate++; yearSet.add(c.getFullYear()); }
  else if (typeof c === 'number') nNumber++;
  else if (typeof c === 'string') nString++;
  else nOther++;
}
console.log(`\n=== C列(作業日) 型分布 (5行目以降) ===`);
console.log({ nDate, nNumber, nString, nOther, nEmpty });
console.log('Years (if Date):', [...yearSet].sort());
