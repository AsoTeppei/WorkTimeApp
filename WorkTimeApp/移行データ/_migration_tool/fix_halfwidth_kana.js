// 半角カタカナ → 全角カタカナ への一括変換スクリプト。
// 対象: WorkLogs.ContentName / WorkLogs.Details / WorkTypes.TypeName
//
// 使い方:
//   node fix_halfwidth_kana.js          ← ドライラン（変換対象を一覧表示するだけ）
//   node fix_halfwidth_kana.js --apply  ← 実際に DB を UPDATE
//
// 事前に backup_db.js を必ず実行しておくこと（万一の復旧用）。

const sql = require('mssql');

const dbConfig = {
  user:'yonekura', password:'yone6066', server:'192.168.1.8', database:'WorkTimeDB',
  options:{ encrypt:false, trustServerCertificate:true, instanceName:'SQLEXPRESS', enableArithAbort:true },
  connectionTimeout:15000, requestTimeout:600000
};

const APPLY = process.argv.includes('--apply');

// 半角カタカナ → 全角カタカナ 変換マップ
const HALF_TO_FULL = {
  'ｦ':'ヲ','ｧ':'ァ','ｨ':'ィ','ｩ':'ゥ','ｪ':'ェ','ｫ':'ォ','ｬ':'ャ','ｭ':'ュ','ｮ':'ョ','ｯ':'ッ',
  'ｰ':'ー',
  'ｱ':'ア','ｲ':'イ','ｳ':'ウ','ｴ':'エ','ｵ':'オ',
  'ｶ':'カ','ｷ':'キ','ｸ':'ク','ｹ':'ケ','ｺ':'コ',
  'ｻ':'サ','ｼ':'シ','ｽ':'ス','ｾ':'セ','ｿ':'ソ',
  'ﾀ':'タ','ﾁ':'チ','ﾂ':'ツ','ﾃ':'テ','ﾄ':'ト',
  'ﾅ':'ナ','ﾆ':'ニ','ﾇ':'ヌ','ﾈ':'ネ','ﾉ':'ノ',
  'ﾊ':'ハ','ﾋ':'ヒ','ﾌ':'フ','ﾍ':'ヘ','ﾎ':'ホ',
  'ﾏ':'マ','ﾐ':'ミ','ﾑ':'ム','ﾒ':'メ','ﾓ':'モ',
  'ﾔ':'ヤ','ﾕ':'ユ','ﾖ':'ヨ',
  'ﾗ':'ラ','ﾘ':'リ','ﾙ':'ル','ﾚ':'レ','ﾛ':'ロ',
  'ﾜ':'ワ','ﾝ':'ン',
  '｡':'。','｢':'「','｣':'」','､':'、','･':'・'
};
// 濁点・半濁点が付くペア（半角は2文字 → 全角1文字に合成）
const DAKU = {
  'ｶ':'ガ','ｷ':'ギ','ｸ':'グ','ｹ':'ゲ','ｺ':'ゴ',
  'ｻ':'ザ','ｼ':'ジ','ｽ':'ズ','ｾ':'ゼ','ｿ':'ゾ',
  'ﾀ':'ダ','ﾁ':'ヂ','ﾂ':'ヅ','ﾃ':'デ','ﾄ':'ド',
  'ﾊ':'バ','ﾋ':'ビ','ﾌ':'ブ','ﾍ':'ベ','ﾎ':'ボ',
  'ｳ':'ヴ'
};
const HANDAKU = {
  'ﾊ':'パ','ﾋ':'ピ','ﾌ':'プ','ﾍ':'ペ','ﾎ':'ポ'
};

function toFullKana(s) {
  if (s == null) return s;
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    const next = s[i+1];
    if (next === 'ﾞ' && DAKU[c])      { out += DAKU[c];    i++; continue; }
    if (next === 'ﾟ' && HANDAKU[c])   { out += HANDAKU[c]; i++; continue; }
    out += HALF_TO_FULL[c] != null ? HALF_TO_FULL[c] : c;
  }
  // 余った単独の濁点・半濁点も全角に
  out = out.replace(/ﾞ/g, '゛').replace(/ﾟ/g, '゜');
  return out;
}

// 半角カタカナ判定（U+FF61〜U+FF9F のいずれかを含むか）
const HW_KANA_RE = /[\uFF61-\uFF9F]/;

(async () => {
  console.log(`=== 半角カタカナ → 全角カタカナ 一括変換 (${APPLY ? 'APPLY モード' : 'DRY-RUN'}) ===`);
  await sql.connect(dbConfig);

  // ---------- WorkLogs.ContentName ----------
  console.log('\n[1/3] WorkLogs.ContentName をスキャン...');
  const r1 = await sql.query(`
    SELECT LogID, ContentName
    FROM WorkLogs
    WHERE IsDeleted = 0 AND ContentName IS NOT NULL
      AND ContentName COLLATE Latin1_General_BIN LIKE N'%[ｦ-ﾟ]%'
  `);
  const targets1 = r1.recordset
    .filter(r => HW_KANA_RE.test(r.ContentName))
    .map(r => ({ id: r.LogID, before: r.ContentName, after: toFullKana(r.ContentName) }))
    .filter(r => r.before !== r.after);
  console.log(`  対象: ${targets1.length} 件`);
  targets1.forEach(t => console.log(`    LogID=${t.id}: 「${t.before}」 → 「${t.after}」`));

  // ---------- WorkLogs.Details ----------
  console.log('\n[2/3] WorkLogs.Details をスキャン...');
  const r2 = await sql.query(`
    SELECT LogID, Details
    FROM WorkLogs
    WHERE IsDeleted = 0 AND Details IS NOT NULL
      AND Details COLLATE Latin1_General_BIN LIKE N'%[ｦ-ﾟ]%'
  `);
  const targets2 = r2.recordset
    .filter(r => HW_KANA_RE.test(r.Details))
    .map(r => ({ id: r.LogID, before: r.Details, after: toFullKana(r.Details) }))
    .filter(r => r.before !== r.after);
  console.log(`  対象: ${targets2.length} 件`);
  targets2.forEach(t => console.log(`    LogID=${t.id}: 「${t.before.slice(0,50)}${t.before.length>50?'...':''}」 → 「${t.after.slice(0,50)}${t.after.length>50?'...':''}」`));

  // ---------- WorkTypes.TypeName ----------
  console.log('\n[3/3] WorkTypes.TypeName をスキャン...');
  const r3 = await sql.query(`
    SELECT WorkTypeID, TypeName
    FROM WorkTypes
    WHERE IsActive = 1 AND TypeName IS NOT NULL
      AND TypeName COLLATE Latin1_General_BIN LIKE N'%[ｦ-ﾟ]%'
  `);
  const targets3 = r3.recordset
    .filter(r => HW_KANA_RE.test(r.TypeName))
    .map(r => ({ id: r.WorkTypeID, before: r.TypeName, after: toFullKana(r.TypeName) }))
    .filter(r => r.before !== r.after);
  console.log(`  対象: ${targets3.length} 件`);
  targets3.forEach(t => console.log(`    WorkTypeID=${t.id}: 「${t.before}」 → 「${t.after}」`));

  const total = targets1.length + targets2.length + targets3.length;
  console.log(`\n=== 合計対象: ${total} 件 ===`);

  if (!APPLY) {
    console.log('\n※ ドライランのみ。実際に書き換えるには --apply を付けて再実行してください。');
    await sql.close();
    return;
  }

  if (total === 0) {
    console.log('変換対象がないため UPDATE はスキップします。');
    await sql.close();
    return;
  }

  console.log('\n--- UPDATE 実行 ---');
  const tx = new sql.Transaction();
  await tx.begin();
  try {
    let n1=0, n2=0, n3=0;
    for (const t of targets1) {
      const req = new sql.Request(tx);
      req.input('id', sql.Int, t.id);
      req.input('v',  sql.NVarChar(100), t.after);
      const r = await req.query('UPDATE WorkLogs SET ContentName = @v, UpdatedAt = GETDATE() WHERE LogID = @id');
      n1 += r.rowsAffected[0] || 0;
    }
    for (const t of targets2) {
      const req = new sql.Request(tx);
      req.input('id', sql.Int, t.id);
      req.input('v',  sql.NVarChar(sql.MAX), t.after);
      const r = await req.query('UPDATE WorkLogs SET Details = @v, UpdatedAt = GETDATE() WHERE LogID = @id');
      n2 += r.rowsAffected[0] || 0;
    }
    for (const t of targets3) {
      const req = new sql.Request(tx);
      req.input('id', sql.Int, t.id);
      req.input('v',  sql.NVarChar(100), t.after);
      const r = await req.query('UPDATE WorkTypes SET TypeName = @v WHERE WorkTypeID = @id');
      n3 += r.rowsAffected[0] || 0;
    }
    await tx.commit();
    console.log(`  WorkLogs.ContentName : ${n1} 行更新`);
    console.log(`  WorkLogs.Details     : ${n2} 行更新`);
    console.log(`  WorkTypes.TypeName   : ${n3} 行更新`);
    console.log('完了。');
  } catch (e) {
    await tx.rollback();
    console.error('UPDATE 失敗、ロールバックしました:', e.message || e);
    process.exit(1);
  }

  await sql.close();
})().catch(e => { console.error('FATAL:', e.message || e); process.exit(1); });
