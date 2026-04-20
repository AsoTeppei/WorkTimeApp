// 列ごとの値分布（特に担当名・作業内容・出荷後対応）を調べる
const XLSX = require('xlsx');
const path = require('path');
const SRC = path.join(__dirname, '..', 'ｿﾌﾄ作業時間SP.xls');
const wb = XLSX.readFile(SRC, { cellDates: true });
const ws = wb.Sheets['ｿﾌﾄ作業時間SP'];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

// データは row 5 (0-based) から
const HEADER_ROW = 4;
const DATA_START = 5;
const COLS = { proj:0, place:1, date:2, hours:3, lodge:4, name:5, content:6, after:7, detail:8 };

const stat = (col) => {
  const m = new Map();
  for (let i = DATA_START; i < aoa.length; i++) {
    const v = aoa[i][col];
    if (v === null || v === undefined || v === '') continue;
    const k = String(v);
    m.set(k, (m.get(k) || 0) + 1);
  }
  return [...m.entries()].sort((a,b)=>b[1]-a[1]);
};

const toJST = (d) => {
  if (!(d instanceof Date)) return null;
  const t = new Date(d.getTime() + 9*60*60*1000);
  return `${t.getUTCFullYear()}-${String(t.getUTCMonth()+1).padStart(2,'0')}-${String(t.getUTCDate()).padStart(2,'0')}`;
};

console.log('=== 担当 (F列) 集計 ===');
stat(COLS.name).forEach(([k,c]) => console.log(`  ${k.padEnd(15)} ${c}`));

console.log('\n=== 作業場所 (B列) ===');
stat(COLS.place).forEach(([k,c]) => console.log(`  ${k.padEnd(10)} ${c}`));

console.log('\n=== 内容選択 (G列) ===');
stat(COLS.content).forEach(([k,c]) => console.log(`  ${k.padEnd(15)} ${c}`));

console.log('\n=== 出荷後対応 (H列) — 非空値 ===');
stat(COLS.after).forEach(([k,c]) => console.log(`  ${k.padEnd(15)} ${c}`));

console.log('\n=== 宿泊 (E列) ===');
stat(COLS.lodge).forEach(([k,c]) => console.log(`  ${k.padEnd(10)} ${c}`));

// 年別行数（2024 以降が対象）
const byYear = {};
for (let i = DATA_START; i < aoa.length; i++) {
  const d = aoa[i][COLS.date];
  if (!(d instanceof Date)) continue;
  const ymd = toJST(d);
  if (!ymd) continue;
  const y = ymd.slice(0,4);
  byYear[y] = (byYear[y]||0) + 1;
}
console.log('\n=== 年別件数 ===');
Object.keys(byYear).sort().forEach(y => console.log(`  ${y}: ${byYear[y]}`));

// 2024 以降の date / 担当 ペアのユニーク数
const targetUsers = new Set();
const targetDates = new Set();
const targetContents = new Set();
let n2024 = 0;
for (let i = DATA_START; i < aoa.length; i++) {
  const d = aoa[i][COLS.date];
  if (!(d instanceof Date)) continue;
  const ymd = toJST(d);
  if (ymd < '2024-01-01') continue;
  n2024++;
  const name = aoa[i][COLS.name];
  if (name) targetUsers.add(String(name).trim());
  targetDates.add(ymd);
  const c = aoa[i][COLS.content];
  if (c) targetContents.add(String(c).trim());
}
console.log(`\n=== 2024-01-01 以降 ===`);
console.log(`  行数: ${n2024}`);
console.log(`  担当ユニーク: ${[...targetUsers].sort().join(', ')}`);
console.log(`  日付ユニーク: ${targetDates.size}`);
console.log(`  作業内容ユニーク: ${targetContents.size}`);
console.log(`    内訳: ${[...targetContents].sort().join(' / ')}`);

// 注番が "-" のような特殊値を持つ件数
let nDash=0, nProjEmpty=0, nProjNormal=0;
for (let i = DATA_START; i < aoa.length; i++) {
  const d = aoa[i][COLS.date];
  if (!(d instanceof Date)) continue;
  if (toJST(d) < '2024-01-01') continue;
  const p = aoa[i][COLS.proj];
  if (p == null || p === '') nProjEmpty++;
  else if (String(p).trim() === '-' || String(p).trim() === 'ー') nDash++;
  else nProjNormal++;
}
console.log(`\n=== 2024-以降 注番分布 ===`);
console.log(`  通常: ${nProjNormal}, "-"系: ${nDash}, 空: ${nProjEmpty}`);
