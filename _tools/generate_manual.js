// Node.js + SQL Server 社内ツール開発 環境構築・運用手順書
// 実行: node generate_manual.js
const pptxgen = require('pptxgenjs');
const path = require('path');

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'Claude';
pres.company = '社内ツール開発';
pres.title   = 'Node.js + SQL Server 環境構築・運用手順書';

// カラー（Midnight Executive ベース）
const C = {
  navy:    '1E2761',
  navy2:   '0F1640',
  blue:    '3B4BA3',
  blueL:   'CADCFC',
  ice:     'E8EEFB',
  white:   'FFFFFF',
  text:    '1F2937',
  gray:    '4B5563',
  gray2:   '9CA3AF',
  grayBg:  'F3F4F6',
  green:   '10B981',
  amber:   'F59E0B',
  red:     'EF4444',
  codeBg:  '0F172A',
  codeFg:  'E2E8F0',
  codeCm:  '94A3B8',
};

const FONT_TITLE = 'Georgia';
const FONT_JP    = 'Yu Gothic';
const FONT_JPB   = 'Yu Gothic UI';
const FONT_MONO  = 'Consolas';

// ----- スライドマスタ -----
pres.defineSlideMaster({
  title: 'CONTENT',
  background: { color: C.white },
  objects: [
    // 左サイドバー
    { rect: { x: 0, y: 0, w: 0.4, h: 7.5, fill: { color: C.navy } } },
    // 装飾のアクセント縦線
    { rect: { x: 0.4, y: 0, w: 0.04, h: 7.5, fill: { color: C.blue } } },
    // 左サイド縦書きタイトル的装飾
    { text: {
        text: 'DEV  MANUAL',
        options: {
          x: -0.1, y: 3.1, w: 0.6, h: 1.0,
          fontFace: FONT_TITLE, fontSize: 11, color: C.blueL, bold: true,
          align: 'center', valign: 'middle', rotate: 270
        }
    }},
    // フッタバー
    { rect: { x: 0.44, y: 7.18, w: 12.89, h: 0.32, fill: { color: C.navy } } },
    { text: {
        text: 'Node.js + SQL Server 環境構築・運用手順書',
        options: {
          x: 0.6, y: 7.18, w: 8, h: 0.32,
          fontFace: FONT_JP, fontSize: 9, color: C.blueL, valign: 'middle'
        }
    }},
  ],
  slideNumber: {
    x: 12.7, y: 7.18, w: 0.55, h: 0.32,
    fontFace: FONT_MONO, fontSize: 10, color: C.blueL, align: 'right', valign: 'middle'
  }
});

// ----- ヘルパー：コンテンツスライド -----
function addSlide(num, title, subtitle) {
  const s = pres.addSlide({ masterName: 'CONTENT' });
  // 番号バッジ
  s.addShape('ellipse', {
    x: 0.7, y: 0.45, w: 0.75, h: 0.75,
    fill: { color: C.blue }, line: { color: C.blue }
  });
  s.addText(String(num).padStart(2, '0'), {
    x: 0.7, y: 0.45, w: 0.75, h: 0.75,
    fontFace: FONT_TITLE, fontSize: 16, bold: true, color: C.white,
    align: 'center', valign: 'middle'
  });
  // タイトル
  s.addText(title, {
    x: 1.6, y: 0.4, w: 11.5, h: 0.55,
    fontFace: FONT_JPB, fontSize: 24, bold: true, color: C.navy, valign: 'middle'
  });
  if (subtitle) {
    s.addText(subtitle, {
      x: 1.6, y: 0.92, w: 11.5, h: 0.32,
      fontFace: FONT_JP, fontSize: 12, color: C.gray, valign: 'middle', italic: true
    });
  }
  // 区切りライン（短い）
  s.addShape('rect', {
    x: 1.6, y: 1.32, w: 0.5, h: 0.05,
    fill: { color: C.blue }, line: { color: C.blue }
  });
  return s;
}

// ----- ヘルパー：箇条書き -----
function bullets(s, items, opts = {}) {
  const x = opts.x ?? 0.7;
  const y = opts.y ?? 1.55;
  const w = opts.w ?? 12.1;
  const h = opts.h ?? 5.4;
  const fontSize = opts.fontSize ?? 13;

  const text = items.map(it => {
    if (typeof it === 'string') {
      return { text: it, options: { bullet: { code: '25A0' }, color: C.text, fontSize, paraSpaceAfter: 5, fontFace: FONT_JP } };
    }
    return {
      text: it.text,
      options: {
        bullet: it.sub ? { code: '25CB' } : (it.num ? { type: 'number' } : { code: '25A0' }),
        color: it.sub ? C.gray : C.navy,
        fontSize: it.sub ? (fontSize - 1) : fontSize,
        bold: !!it.bold,
        indentLevel: it.sub ? 1 : 0,
        paraSpaceAfter: it.sub ? 3 : 6,
        fontFace: FONT_JP
      }
    };
  });
  s.addText(text, { x, y, w, h, valign: 'top', margin: 0 });
}

// ----- ヘルパー：コマンドプロンプト風ブロック -----
function cmdBlock(s, title, lines, opts = {}) {
  const x = opts.x ?? 1.6;
  const y = opts.y ?? 1.6;
  const w = opts.w ?? 11.5;
  const h = opts.h ?? 4.0;
  const fontSize = opts.fontSize ?? 12;
  // 枠
  s.addShape('rect', {
    x, y, w, h,
    fill: { color: C.codeBg }, line: { color: C.codeBg }
  });
  // 上部バー（タイトル）
  s.addShape('rect', {
    x, y, w, h: 0.32,
    fill: { color: '334155' }, line: { color: '334155' }
  });
  // 信号ドット（装飾）
  ['F87171', 'FBBF24', '34D399'].forEach((col, i) => {
    s.addShape('ellipse', {
      x: x + 0.12 + i * 0.18, y: y + 0.09, w: 0.14, h: 0.14,
      fill: { color: col }, line: { color: col }
    });
  });
  s.addText(title || 'Command Prompt', {
    x: x + 0.8, y: y + 0.02, w: w - 1.6, h: 0.28,
    fontFace: FONT_MONO, fontSize: 10, color: C.codeFg, align: 'center', valign: 'middle'
  });
  // コンテンツ（各行ごとに色分け）
  const textRuns = [];
  lines.forEach((ln, i) => {
    if (typeof ln === 'string') {
      // コメント行 (# / :: / rem 始まり) かどうか判定
      const isComment = /^\s*(#|::|rem\s)/i.test(ln);
      textRuns.push({
        text: ln + (i < lines.length - 1 ? '\n' : ''),
        options: {
          color: isComment ? C.codeCm : C.codeFg,
          fontSize, fontFace: FONT_MONO,
          italic: isComment
        }
      });
    } else {
      textRuns.push({
        text: ln.text + (i < lines.length - 1 ? '\n' : ''),
        options: {
          color: ln.comment ? C.codeCm : (ln.hi ? '93C5FD' : C.codeFg),
          fontSize, fontFace: FONT_MONO,
          italic: !!ln.comment, bold: !!ln.hi
        }
      });
    }
  });
  s.addText(textRuns, {
    x: x + 0.2, y: y + 0.4, w: w - 0.4, h: h - 0.5,
    valign: 'top', margin: 0
  });
}

// ----- ヘルパー：情報ボックス -----
function infoBox(s, kind, title, body, opts = {}) {
  const styles = {
    info:    { bg: C.ice,     border: C.blue,  head: C.blue },
    tip:     { bg: 'ECFDF5',  border: C.green, head: C.green },
    warn:    { bg: 'FEF3C7',  border: C.amber, head: '92400E' },
    danger:  { bg: 'FEE2E2',  border: C.red,   head: '991B1B' },
  };
  const st = styles[kind] || styles.info;
  const x = opts.x ?? 1.6;
  const y = opts.y ?? 6.0;
  const w = opts.w ?? 11.5;
  const h = opts.h ?? 0.95;
  s.addShape('rect', { x, y, w: 0.1, h, fill: { color: st.border }, line: { color: st.border } });
  s.addShape('rect', { x: x + 0.1, y, w: w - 0.1, h, fill: { color: st.bg }, line: { color: st.bg } });
  s.addText([
    { text: title + '\n', options: { bold: true, color: st.head, fontSize: 12, fontFace: FONT_JPB } },
    { text: body, options: { color: C.text, fontSize: 11, fontFace: FONT_JP } },
  ], { x: x + 0.25, y: y + 0.08, w: w - 0.4, h: h - 0.16, valign: 'top', margin: 0 });
}

// ======================================================================
// 1. 表紙
// ======================================================================
{
  const s = pres.addSlide();
  s.background = { color: C.navy };
  // 大きな装飾円
  s.addShape('ellipse', {
    x: 9.5, y: -2.0, w: 7.0, h: 7.0,
    fill: { color: C.navy2 }, line: { color: C.navy2 }
  });
  s.addShape('ellipse', {
    x: 10.2, y: -1.3, w: 5.5, h: 5.5,
    fill: { color: '1A2457' }, line: { color: '1A2457' }
  });
  // アクセントライン
  s.addShape('rect', { x: 0, y: 3.15, w: 9, h: 0.08, fill: { color: C.blueL }, line: { color: C.blueL } });

  s.addText('DEV SETUP', {
    x: 0.8, y: 1.2, w: 10, h: 0.6,
    fontFace: FONT_TITLE, fontSize: 18, color: C.blueL, italic: true
  });
  s.addText('Node.js + SQL Server', {
    x: 0.8, y: 1.9, w: 11, h: 1.2,
    fontFace: FONT_TITLE, fontSize: 52, bold: true, color: C.white
  });
  s.addText('環境構築・運用手順書', {
    x: 0.8, y: 3.4, w: 11, h: 0.7,
    fontFace: FONT_JPB, fontSize: 28, bold: true, color: C.blueL
  });
  s.addText('— インストールから本番フォルダへのコピー運用まで —', {
    x: 0.8, y: 4.2, w: 11, h: 0.5,
    fontFace: FONT_JP, fontSize: 16, color: 'A5B4FC', italic: true
  });

  // 下部メタ情報
  s.addShape('rect', { x: 0.8, y: 5.6, w: 0.15, h: 1.3, fill: { color: C.blueL }, line: { color: C.blueL } });
  s.addText([
    { text: '対象 : ', options: { bold: true, color: C.white, fontSize: 13 } },
    { text: '社内ツールを Node.js + SQL Server で自作する担当者\n',  options: { color: 'CBD5E1', fontSize: 13 } },
    { text: 'ゴール : ', options: { bold: true, color: C.white, fontSize: 13 } },
    { text: '公式インストーラから本番サーバー常駐までを自分の手で完走できる\n', options: { color: 'CBD5E1', fontSize: 13 } },
    { text: '重点 : ', options: { bold: true, color: C.white, fontSize: 13 } },
    { text: 'コマンドプロンプト操作と robocopy による本番反映手順', options: { color: 'CBD5E1', fontSize: 13 } },
  ], {
    x: 1.05, y: 5.55, w: 11, h: 1.4,
    fontFace: FONT_JP, valign: 'top', paraSpaceAfter: 4, margin: 0
  });
}

// ======================================================================
// 2. 目次
// ======================================================================
{
  const s = addSlide(' ', '目次', 'インストール → 接続 → 開発 → 本番反映 の流れ');
  // 左に番号バッジ隠し
  const chapters = [
    ['Ⅰ', '事前準備と全体像',       '前提環境 / 作業の流れ'],
    ['Ⅱ', 'Node.js のインストール', '公式インストーラと PATH 確認'],
    ['Ⅲ', 'SQL Server / SSMS',     'インストール・TCP 有効化・DB 作成'],
    ['Ⅳ', 'バックエンド初期化',     'mkdir / npm init / npm install'],
    ['Ⅴ', 'msnodesqlv8 特別編',    'ビルド要件と導入・トラブル対処'],
    ['Ⅵ', 'server.js とサーバー起動', '最小テンプレートと動作確認'],
    ['Ⅶ', '作業フォルダ → 本番 Host フォルダ', '手動 / robocopy / deploy.bat'],
    ['Ⅷ', '常駐化と再起動',         'タスクスケジューラで常駐'],
    ['Ⅸ', 'トラブルシューティング', 'よくあるエラーと対処表'],
  ];
  chapters.forEach((c, i) => {
    const y = 1.55 + i * 0.6;
    // 行全体の薄い背景
    s.addShape('rect', {
      x: 1.6, y, w: 11.5, h: 0.52,
      fill: { color: i % 2 === 0 ? C.grayBg : C.white },
      line: { color: 'E5E7EB', width: 0.5 }
    });
    s.addText(c[0], {
      x: 1.7, y, w: 0.6, h: 0.52,
      fontFace: FONT_TITLE, fontSize: 18, bold: true, color: C.blue,
      align: 'center', valign: 'middle'
    });
    s.addText(c[1], {
      x: 2.35, y, w: 5.8, h: 0.52,
      fontFace: FONT_JPB, fontSize: 15, bold: true, color: C.navy, valign: 'middle'
    });
    s.addText(c[2], {
      x: 8.2, y, w: 4.8, h: 0.52,
      fontFace: FONT_JP, fontSize: 11, color: C.gray, valign: 'middle'
    });
  });
}

// ======================================================================
// 3. 事前準備 / 全体像
// ======================================================================
{
  const s = addSlide(1, '事前準備と全体像', 'これから何を入れて、どう繋げるのか');

  // 左：前提環境
  s.addShape('rect', { x: 1.6, y: 1.55, w: 5.4, h: 5.3, fill: { color: C.ice }, line: { color: C.blue, width: 1 } });
  s.addText('前提環境', {
    x: 1.75, y: 1.65, w: 5.1, h: 0.4,
    fontFace: FONT_JPB, fontSize: 15, bold: true, color: C.navy
  });
  s.addText([
    { text: 'OS', options: { bold: true, color: C.blue, fontSize: 12, fontFace: FONT_JPB } },
    { text: '  Windows 10 / 11（64bit）\n\n', options: { color: C.text, fontSize: 12, fontFace: FONT_JP } },
    { text: '権限', options: { bold: true, color: C.blue, fontSize: 12, fontFace: FONT_JPB } },
    { text: '  管理者権限（インストール時に必要）\n\n', options: { color: C.text, fontSize: 12, fontFace: FONT_JP } },
    { text: 'ネットワーク', options: { bold: true, color: C.blue, fontSize: 12, fontFace: FONT_JPB } },
    { text: '  インターネット接続（ダウンロード用）\n\n', options: { color: C.text, fontSize: 12, fontFace: FONT_JP } },
    { text: 'ディスク空き', options: { bold: true, color: C.blue, fontSize: 12, fontFace: FONT_JPB } },
    { text: '  約 10 GB（SQL Server 本体 + SSMS + ツール群）\n\n', options: { color: C.text, fontSize: 12, fontFace: FONT_JP } },
    { text: 'ツール', options: { bold: true, color: C.blue, fontSize: 12, fontFace: FONT_JPB } },
    { text: '  コマンドプロンプト（標準）\n  テキストエディタ（VS Code 推奨）', options: { color: C.text, fontSize: 12, fontFace: FONT_JP } },
  ], { x: 1.85, y: 2.1, w: 5.0, h: 4.7, valign: 'top', margin: 0 });

  // 右：インストール順
  s.addShape('rect', { x: 7.4, y: 1.55, w: 5.7, h: 5.3, fill: { color: C.white }, line: { color: C.blue, width: 1 } });
  s.addText('インストールする順序', {
    x: 7.55, y: 1.65, w: 5.4, h: 0.4,
    fontFace: FONT_JPB, fontSize: 15, bold: true, color: C.navy
  });
  const steps = [
    ['1', 'Node.js', '（LTS）'],
    ['2', 'Visual Studio Code', '（エディタ）'],
    ['3', 'SQL Server', '（Express 版で十分）'],
    ['4', 'SQL Server Management Studio', '（SSMS）'],
    ['5', 'Visual Studio Build Tools', '（msnodesqlv8 用）'],
    ['6', 'npm パッケージ', '（express / mssql / msnodesqlv8）'],
  ];
  steps.forEach((st, i) => {
    const y = 2.15 + i * 0.72;
    s.addShape('ellipse', {
      x: 7.6, y: y + 0.04, w: 0.45, h: 0.45,
      fill: { color: C.navy }, line: { color: C.navy }
    });
    s.addText(st[0], {
      x: 7.6, y: y + 0.04, w: 0.45, h: 0.45,
      fontFace: FONT_TITLE, fontSize: 14, bold: true, color: C.white,
      align: 'center', valign: 'middle'
    });
    s.addText(st[1], {
      x: 8.15, y, w: 4.8, h: 0.35,
      fontFace: FONT_JPB, fontSize: 13, bold: true, color: C.navy
    });
    s.addText(st[2], {
      x: 8.15, y: y + 0.3, w: 4.8, h: 0.3,
      fontFace: FONT_JP, fontSize: 10, color: C.gray
    });
  });
}

// ======================================================================
// 4. コマンドプロンプトの基礎
// ======================================================================
{
  const s = addSlide(2, 'コマンドプロンプトの基本操作', '以降の作業で頻繁に使うので先に押さえる');

  const tbl = [
    ['起動',                  'Win キー → 「cmd」と入力 → Enter',         '「ファイル名を指定して実行」から cmd でも可'],
    ['管理者として起動',       'Win キー → 「cmd」と入力 → 右クリック →', '「管理者として実行」'],
    ['現在のフォルダを確認',   'cd',                                     '引数なしで現在地を表示'],
    ['フォルダ移動',           'cd  フォルダ名',                          'cd ..  で 1 階層上へ'],
    ['ドライブ変更',           'D:',                                      '直接ドライブ文字を打つ'],
    ['ドライブ＆フォルダ同時', 'pushd  D:\\path\\to\\folder',             'ネットワークドライブも可'],
    ['戻る',                   'popd',                                    'pushd で移動した元の場所へ戻る'],
    ['ファイル一覧',           'dir',                                     'dir /b で名前だけ'],
    ['フォルダ作成',           'mkdir  フォルダ名',                        'md でも同じ'],
    ['コピー',                 'copy  src  dst',                          'フォルダごとは xcopy / robocopy'],
    ['クリア',                 'cls',                                     '画面クリア'],
    ['履歴',                   'F7 キー / ↑ キー',                        '過去に打ったコマンドを呼び出し'],
    ['終了',                   'exit',                                    'ウィンドウを閉じる'],
  ];
  const baseY = 1.55, rowH = 0.42;
  // ヘッダ
  s.addShape('rect', { x: 1.6, y: baseY, w: 11.5, h: rowH, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('操作',     { x: 1.7, y: baseY, w: 2.8, h: rowH, fontFace: FONT_JPB, fontSize: 12, bold: true, color: C.white, valign: 'middle' });
  s.addText('コマンド', { x: 4.5, y: baseY, w: 4.3, h: rowH, fontFace: FONT_JPB, fontSize: 12, bold: true, color: C.white, valign: 'middle' });
  s.addText('備考',     { x: 8.9, y: baseY, w: 4.1, h: rowH, fontFace: FONT_JPB, fontSize: 12, bold: true, color: C.white, valign: 'middle' });

  tbl.forEach((r, i) => {
    const y = baseY + (i + 1) * rowH;
    s.addShape('rect', { x: 1.6, y, w: 11.5, h: rowH,
      fill: { color: i % 2 === 0 ? C.grayBg : C.white }, line: { color: 'E5E7EB', width: 0.5 } });
    s.addText(r[0], { x: 1.7, y, w: 2.8, h: rowH, fontFace: FONT_JP, fontSize: 11, color: C.text, valign: 'middle' });
    s.addText(r[1], { x: 4.5, y, w: 4.3, h: rowH, fontFace: FONT_MONO, fontSize: 11, bold: true, color: C.blue, valign: 'middle' });
    s.addText(r[2], { x: 8.9, y, w: 4.1, h: rowH, fontFace: FONT_JP, fontSize: 10, color: C.gray, valign: 'middle' });
  });
}

// ======================================================================
// 5. Node.js ダウンロード
// ======================================================================
{
  const s = addSlide(3, 'Node.js インストール ①', '公式サイトから LTS 版インストーラを入手');
  bullets(s, [
    { text: '公式サイトにアクセス', bold: true },
    { text: 'https://nodejs.org/ja', sub: true },
    { text: '「LTS（推奨版）」を選ぶ', bold: true },
    { text: '左側にある「LTS」ボタンをクリック（最新版ではなく LTS を選ぶ）', sub: true },
    { text: 'LTS は長期サポート版で、業務ツールに向いている', sub: true },
    { text: 'Windows Installer (.msi) をダウンロード', bold: true },
    { text: '64bit 版（node-vXX.XX.X-x64.msi）をダウンロード', sub: true },
    { text: 'ダウンロード先はどこでも良いが、Downloads フォルダが無難', sub: true },
  ], { h: 2.7 });

  infoBox(s, 'info', 'LTS とは？',
    'Long Term Support（長期サポート版）の略。バグ修正とセキュリティパッチが長期間提供される安定版。' +
    '業務で使うツールを作るなら、新機能よりも安定性を優先して必ず LTS を選択する。',
    { y: 4.3, h: 0.9 }
  );
  infoBox(s, 'tip', 'バージョンを揃える',
    '開発 PC と本番サーバー PC で Node.js のメジャーバージョンを揃えると、動作差が出にくい。' +
    '両方で同じ .msi を使うのが最も確実。',
    { y: 5.35, h: 0.9 }
  );
  infoBox(s, 'warn', 'ファイル保存の落とし穴',
    'ブラウザによっては .msi を「安全でない可能性」として警告する。信頼できる公式サイトからのダウンロードであれば「継続」でよい。',
    { y: 6.4, h: 0.75 }
  );
}

// ======================================================================
// 6. Node.js インストーラ操作
// ======================================================================
{
  const s = addSlide(4, 'Node.js インストール ②', 'インストーラ画面の項目を順に解説');
  bullets(s, [
    { text: 'ダウンロードした .msi をダブルクリック', bold: true },
    { text: 'User Account Control（UAC）で「はい」をクリック', sub: true },
    { text: '各画面を順に進める', bold: true },
    { text: 'Welcome → Next', sub: true },
    { text: 'License → I accept にチェック → Next', sub: true },
    { text: 'Destination Folder → そのまま Next（既定: C:\\Program Files\\nodejs\\）', sub: true },
    { text: 'Custom Setup → すべてデフォルトで Next', sub: true },
    { text: '  └ Node.js runtime / npm package manager / Add to PATH は必須', sub: true },
    { text: 'Tools for Native Modules のチェック（重要）', bold: true },
    { text: '「Automatically install the necessary tools」にチェックを入れる', sub: true },
    { text: '→ 後で msnodesqlv8 をビルドするために Python や Build Tools が自動導入される', sub: true },
    { text: '→ これを飛ばすと msnodesqlv8 のインストールで失敗することがある', sub: true },
    { text: 'Install をクリック → 完了まで数分待つ', bold: true },
    { text: 'Finish で終了', sub: true },
  ]);
}

// ======================================================================
// 7. Node.js 動作確認
// ======================================================================
{
  const s = addSlide(5, 'Node.js インストール ③', 'コマンドプロンプトで動作確認');
  bullets(s, [
    { text: 'コマンドプロンプトを開く', bold: true },
    { text: 'インストール後は「新しいウィンドウ」を開くこと（PATH 反映のため）', sub: true },
  ], { h: 0.85 });

  cmdBlock(s, 'Command Prompt — 動作確認', [
    { text: 'rem Node.js のバージョン表示', comment: true },
    'node -v',
    { text: 'v20.17.0    ← こんな感じで表示されれば OK', comment: true },
    '',
    { text: 'rem npm（パッケージマネージャ）のバージョン', comment: true },
    'npm -v',
    { text: '10.8.2', comment: true },
    '',
    { text: 'rem 簡単な 1 行スクリプト実行', comment: true },
    'node -e "console.log(\'hello node\')"',
    { text: 'hello node', comment: true },
  ], { y: 2.4, h: 3.3, fontSize: 13 });

  infoBox(s, 'danger', 'node: コマンドが見つからない場合',
    'PATH にインストール先が追加されていない可能性が高い。' +
    'コマンドプロンプトを「新しく」開き直すとほぼ直る。直らない場合は一度サインアウト → 再ログインする。',
    { y: 5.9, h: 1.2 }
  );
}

// ======================================================================
// 8. VS Code
// ======================================================================
{
  const s = addSlide(6, '推奨エディタ: Visual Studio Code', '無償かつ定番。これを使っておけば間違いない');
  bullets(s, [
    { text: '入手先', bold: true },
    { text: 'https://code.visualstudio.com/', sub: true },
    { text: '「Download for Windows」→ ユーザー インストーラをダウンロード', sub: true },
    { text: 'インストール時の重要チェック', bold: true },
    { text: '「エクスプローラーのファイル コンテキスト メニューに [Code で開く] アクションを追加する」', sub: true },
    { text: '「エクスプローラーのディレクトリ コンテキスト メニューに [Code で開く] アクションを追加する」', sub: true },
    { text: '「PATH への追加」（既定でチェック済み、そのまま）', sub: true },
    { text: '入れておくと便利な拡張機能', bold: true },
    { text: 'Japanese Language Pack for Visual Studio Code（日本語化）', sub: true },
    { text: 'ESLint / Prettier（整形）', sub: true },
    { text: 'SQL Server (mssql)（SSMS の代わりに簡易クエリ実行）', sub: true },
  ], { h: 4.3 });

  cmdBlock(s, 'Command Prompt — VS Code でフォルダを開く', [
    { text: 'rem 現在のフォルダを VS Code で開く', comment: true },
    'code .',
    '',
    { text: 'rem 任意のフォルダを指定して開く', comment: true },
    'code D:\\projects\\myapp',
  ], { y: 5.95, h: 1.25, fontSize: 12 });
}

// ======================================================================
// 9. SQL Server ダウンロード
// ======================================================================
{
  const s = addSlide(7, 'SQL Server のインストール ①', 'Express 版（無償）を入手');
  bullets(s, [
    { text: 'エディションの選び方', bold: true },
    { text: 'Express … 無償・10 GB まで・社内ツール用途に十分', sub: true },
    { text: 'Developer … 無償・全機能だが本番利用不可', sub: true },
    { text: 'Standard / Enterprise … 有償', sub: true },
    { text: '→ 社内ツール自作なら Express で OK', sub: true },
    { text: 'ダウンロード', bold: true },
    { text: '「SQL Server Express ダウンロード」で検索', sub: true },
    { text: 'Microsoft 公式ページから SQL2022-SSEI-Expr.exe をダウンロード', sub: true },
    { text: 'インストーラを実行', bold: true },
    { text: 'インストールの種類で「基本」を選ぶ（カスタムは慣れてから）', sub: true },
    { text: '使用許諾契約に同意 → インストール先は既定のまま → インストール', sub: true },
    { text: '完了画面で「接続文字列」「SQL インスタンス名」が表示されるのでメモしておく', sub: true },
  ], { h: 5.0 });

  infoBox(s, 'info', 'メモしておくべき情報',
    'インスタンス名（既定: SQLEXPRESS）と接続文字列（例: Server=localhost\\SQLEXPRESS;）。\n' +
    'Node.js から繋ぐときにこの値を使う。',
    { y: 6.1, h: 1.0 }
  );
}

// ======================================================================
// 10. SSMS のインストール
// ======================================================================
{
  const s = addSlide(8, 'SQL Server のインストール ②', 'SSMS（管理ツール）を入れる');
  bullets(s, [
    { text: 'SSMS とは', bold: true },
    { text: 'SQL Server Management Studio。DB の作成・クエリ実行・バックアップを GUI で行う無償ツール', sub: true },
    { text: 'SQL Server のインストール完了画面から直接 SSMS インストーラへ進めるリンクがある', sub: true },
    { text: 'ダウンロード', bold: true },
    { text: '「SSMS ダウンロード」で検索 → Microsoft 公式の「SQL Server Management Studio」のページへ', sub: true },
    { text: '「無料ダウンロード」をクリック → SSMS-Setup-JPN.exe を入手', sub: true },
    { text: 'インストール', bold: true },
    { text: '.exe を起動 → Install → 完了を待つ（約 5〜10 分）', sub: true },
    { text: 'インストール後「コンピュータを再起動してください」と出たら再起動', sub: true },
    { text: '起動', bold: true },
    { text: 'スタートメニュー → 「SQL Server Management Studio」', sub: true },
    { text: '接続ダイアログが開く：サーバー名に localhost\\SQLEXPRESS を入力 → 接続', sub: true },
    { text: '認証は既定で「Windows 認証」でよい', sub: true },
  ]);
}

// ======================================================================
// 11. SQL Server ネットワーク設定
// ======================================================================
{
  const s = addSlide(9, 'SQL Server のインストール ③', 'TCP/IP を有効化する（Node.js から繋ぐために必要）');
  bullets(s, [
    { text: 'SQL Server 構成マネージャーを開く', bold: true },
    { text: 'スタート → 「SQL Server 構成マネージャー」で検索', sub: true },
    { text: 'TCP/IP を有効化', bold: true },
    { text: '左ペイン「SQL Server ネットワークの構成」→「SQLEXPRESS のプロトコル」', sub: true },
    { text: '右ペインの「TCP/IP」を右クリック →「有効化」', sub: true },
    { text: '警告が出たら OK（サービスの再起動が必要）', sub: true },
    { text: 'ポート番号の確認（任意）', bold: true },
    { text: '「TCP/IP」をダブルクリック →「IP アドレス」タブ → IPAll → TCP ポートに 1433 などを設定', sub: true },
    { text: '既定のインスタンスなら自動で 1433、名前付きは動的ポートになる', sub: true },
    { text: 'サービスを再起動', bold: true },
    { text: '左ペイン「SQL Server のサービス」→「SQL Server (SQLEXPRESS)」を右クリック →「再起動」', sub: true },
  ], { h: 4.4 });

  infoBox(s, 'warn', 'なぜ必要？',
    '既定では共有メモリでしか接続できないため、Node.js の mssql ライブラリや LAN 内の他 PC から接続できない。' +
    '社内運用で使うなら TCP/IP は必ず有効にする。',
    { y: 6.05, h: 1.1 }
  );
}

// ======================================================================
// 12. ファイアウォール
// ======================================================================
{
  const s = addSlide(10, 'SQL Server のインストール ④', 'Windows ファイアウォール設定と認証モード');
  bullets(s, [
    { text: 'ファイアウォール受信規則を追加（LAN 公開する場合のみ）', bold: true },
    { text: 'スタート → 「セキュリティが強化された Windows Defender ファイアウォール」', sub: true },
    { text: '受信の規則 → 新しい規則 → ポート → TCP → 1433 → 許可', sub: true },
    { text: '名前：SQL Server (TCP 1433)', sub: true },
    { text: '認証モードの選択', bold: true },
    { text: 'Windows 認証（推奨）：ログイン中の Windows アカウントで接続。パスワード不要', sub: true },
    { text: 'SQL Server 認証：ID / PW 方式。LAN 内公開や別ユーザーで動かす場合に必要', sub: true },
  ], { h: 3.0 });

  cmdBlock(s, 'Command Prompt — 接続テスト（sqlcmd）', [
    { text: 'rem Windows 認証で接続（-E オプション）', comment: true },
    'sqlcmd -S localhost\\SQLEXPRESS -E',
    '',
    { text: 'rem 接続できたら 1> プロンプトが出るので SELECT を試す', comment: true },
    '1> SELECT @@VERSION',
    '2> GO',
    '',
    { text: 'rem 終了', comment: true },
    '1> EXIT',
  ], { y: 4.6, h: 2.6, fontSize: 12 });
}

// ======================================================================
// 13. データベース作成
// ======================================================================
{
  const s = addSlide(11, 'データベースの作成', 'SSMS GUI と T-SQL 両方のやり方');
  bullets(s, [
    { text: '方法 A：SSMS の GUI で作る（初心者向け）', bold: true },
    { text: 'オブジェクト エクスプローラーで「データベース」を右クリック →「新しいデータベース」', sub: true },
    { text: '名前を入力（例：AppDB） → OK', sub: true },
    { text: '左ツリーに AppDB が現れれば完成', sub: true },
  ], { h: 1.6 });

  cmdBlock(s, 'SSMS クエリウィンドウ — T-SQL で作る方法', [
    { text: '-- 新しいクエリを開いて実行（Ctrl + N → 貼付け → F5）', comment: true },
    'CREATE DATABASE AppDB;',
    'GO',
    '',
    { text: '-- データベースを切り替え', comment: true },
    'USE AppDB;',
    'GO',
    '',
    { text: '-- 試しにテーブルを作る', comment: true },
    'CREATE TABLE TestTable (',
    '    Id   INT IDENTITY(1,1) PRIMARY KEY,',
    '    Name NVARCHAR(50) NOT NULL',
    ');',
    'GO',
    '',
    { text: '-- データを 1 行入れる', comment: true },
    "INSERT INTO TestTable (Name) VALUES (N'テスト');",
    'SELECT * FROM TestTable;',
  ], { y: 3.25, h: 3.9, fontSize: 11 });
}

// ======================================================================
// 14. 作業フォルダ作成
// ======================================================================
{
  const s = addSlide(12, 'バックエンド作業フォルダの作成', 'コマンドプロンプトでの操作を丁寧に');
  bullets(s, [
    { text: 'まず場所を決める', bold: true },
    { text: '開発中：自分のドキュメント配下（例：D:\\dev\\myapp）', sub: true },
    { text: '本番：共有サーバーのフォルダ（例：\\\\SERVER01\\apps\\myapp）', sub: true },
    { text: '日本語や空白を含むパスは引用符で囲む癖をつける', sub: true },
  ], { h: 1.65 });

  cmdBlock(s, 'Command Prompt — フォルダ作成と移動', [
    { text: 'rem ドライブを D: に変更', comment: true },
    'D:',
    '',
    { text: 'rem dev フォルダが無ければ作って移動', comment: true },
    'mkdir D:\\dev',
    'cd D:\\dev',
    '',
    { text: 'rem プロジェクト用サブフォルダを作成', comment: true },
    'mkdir myapp',
    'cd myapp',
    '',
    { text: 'rem 現在の場所を確認（プロンプトに D:\\dev\\myapp> と表示されていれば OK）', comment: true },
    'cd',
    '',
    { text: 'rem 空フォルダかどうか一覧で確認', comment: true },
    'dir',
    '',
    { text: 'rem 1 行でドライブ変更＋移動＋戻るもできる', comment: true },
    'pushd "D:\\dev\\myapp"',
    'popd',
  ], { y: 3.35, h: 3.8, fontSize: 11 });
}

// ======================================================================
// 15. npm init
// ======================================================================
{
  const s = addSlide(13, 'プロジェクト初期化 (npm init)', 'package.json を作る');
  bullets(s, [
    { text: 'npm init とは', bold: true },
    { text: 'プロジェクト情報を書いた package.json を作るコマンド', sub: true },
    { text: '以降インストールしたパッケージは package.json に記録される → 別 PC でも npm install だけで再現できる', sub: true },
  ], { h: 1.3 });

  cmdBlock(s, 'Command Prompt — package.json の生成', [
    { text: 'rem プロジェクトフォルダに居ることを確認', comment: true },
    'cd',
    { text: 'D:\\dev\\myapp', comment: true },
    '',
    { text: 'rem 対話なしで既定値の package.json を一発作成', comment: true },
    'npm init -y',
    '',
    { text: 'rem 出力例', comment: true },
    { text: 'Wrote to D:\\dev\\myapp\\package.json:', comment: true },
    { text: '{', comment: true },
    { text: '  "name": "myapp",', comment: true },
    { text: '  "version": "1.0.0",', comment: true },
    { text: '  "main": "index.js",', comment: true },
    { text: '  "scripts": { "test": "..." },', comment: true },
    { text: '  "license": "ISC"', comment: true },
    { text: '}', comment: true },
    '',
    { text: 'rem 生成できたか確認', comment: true },
    'dir package.json',
  ], { y: 3.0, h: 4.2, fontSize: 11 });
}

// ======================================================================
// 16. npm install 基本
// ======================================================================
{
  const s = addSlide(14, 'npm install でパッケージを入れる', 'express などの基本パッケージから');
  bullets(s, [
    { text: '--save は今は不要', bold: true },
    { text: 'npm 5 以降は npm install した時点で package.json に自動追加される', sub: true },
    { text: '--save-dev（省略形 -D）は開発時のみ使う依存', sub: true },
    { text: '代表的なパッケージ', bold: true },
    { text: 'express … 軽量 Web サーバー', sub: true },
    { text: 'mssql … Node から SQL Server に繋ぐ（ピュア JS ドライバ）', sub: true },
    { text: 'msnodesqlv8 … Windows 認証で繋ぎたい場合に使うドライバ（次ページで詳述）', sub: true },
    { text: 'dotenv … 環境変数を .env から読む', sub: true },
    { text: 'nodemon … 保存時に自動で再起動（開発用）', sub: true },
  ], { h: 3.4 });

  cmdBlock(s, 'Command Prompt — 基本パッケージの導入', [
    { text: 'rem 本番依存', comment: true },
    'npm install express',
    'npm install mssql',
    'npm install dotenv',
    '',
    { text: 'rem 開発用（-D は --save-dev の省略）', comment: true },
    'npm install -D nodemon',
    '',
    { text: 'rem 現在入っているパッケージを一覧', comment: true },
    'npm ls --depth=0',
  ], { y: 5.0, h: 2.1, fontSize: 12 });
}

// ======================================================================
// 17. msnodesqlv8 特別編
// ======================================================================
{
  const s = addSlide(15, 'msnodesqlv8 のインストール', 'Windows 認証で SQL Server に繋ぐためのドライバ（ビルドが必要）');
  bullets(s, [
    { text: 'なぜ mssql だけでは不十分？', bold: true },
    { text: 'ピュア JS 版の mssql 単体だと、SQL Server 認証（ID/PW）しか使えないケースがある', sub: true },
    { text: 'Windows 認証（trustedConnection）で繋ぎたいときは ネイティブドライバ msnodesqlv8 が必要', sub: true },
    { text: '前提：ネイティブビルド環境が必要', bold: true },
    { text: 'Python 3.x（Node.js インストーラで「Tools for Native Modules」にチェックしていれば済）', sub: true },
    { text: 'Visual Studio Build Tools（C++ コンパイラ）', sub: true },
    { text: 'いずれも既定の Node.js インストーラで自動導入される', sub: true },
  ], { h: 3.0 });

  cmdBlock(s, 'Command Prompt — msnodesqlv8 の導入', [
    { text: 'rem プロジェクトフォルダへ移動', comment: true },
    'cd /d D:\\dev\\myapp',
    '',
    { text: 'rem mssql の Windows ネイティブドライバ msnodesqlv8 を追加', comment: true },
    { text: 'rem （環境によっては Python / Build Tools のビルドが走る）', comment: true },
    'npm install msnodesqlv8',
    '',
    { text: 'rem 正常なら最後に「added N packages in Xs」と表示される', comment: true },
    { text: 'rem インストール結果の確認', comment: true },
    'npm ls msnodesqlv8',
  ], { y: 4.75, h: 2.4, fontSize: 12 });
}

// ======================================================================
// 18. msnodesqlv8 トラブルシュート
// ======================================================================
{
  const s = addSlide(16, 'msnodesqlv8 がうまく入らないとき', 'よくある失敗とその対処');
  const tbl = [
    ['症状 / エラー', '原因', '対処'],
    ['gyp ERR! find Python',      'Python が見つからない',                   'Node.js インストーラで「Tools for Native Modules」を付け直す'],
    ['MSBUILD : error MSB3428',   'Visual C++ Build Tools が無い',             'Visual Studio Build Tools をインストール（C++ ワークロード選択）'],
    ['ENOENT: ... node-gyp',      'node-gyp の設定ミス',                       'npm install --global node-gyp を一度実行してから再試行'],
    ['Permission denied / EPERM', 'フォルダに書込権限が無い or ウイルス対策',    '管理者権限で cmd を起動し直す。除外設定も確認'],
    ['binding.node が見つからない', 'ビルドは通ったが読込失敗',                 'node_modules\\msnodesqlv8 を削除して npm install をやり直す'],
    ['Python 2.x を参照している',  '古い Python が PATH にある',               'npm config set python "C:\\Path\\To\\python.exe" で明示'],
  ];
  const baseY = 1.55, rowH = 0.62;
  tbl.forEach((r, i) => {
    const y = baseY + i * rowH;
    const isHead = i === 0;
    s.addShape('rect', { x: 1.6, y, w: 11.5, h: rowH,
      fill: { color: isHead ? C.navy : (i % 2 === 0 ? C.grayBg : C.white) },
      line: { color: 'E5E7EB', width: 0.5 } });
    s.addText(r[0], { x: 1.7, y, w: 3.4, h: rowH, fontFace: FONT_MONO, fontSize: 10,
      bold: isHead, color: isHead ? C.white : C.red, valign: 'middle' });
    s.addText(r[1], { x: 5.15, y, w: 3.4, h: rowH, fontFace: FONT_JP, fontSize: 11,
      bold: isHead, color: isHead ? C.white : C.text, valign: 'middle' });
    s.addText(r[2], { x: 8.6, y, w: 4.45, h: rowH, fontFace: FONT_JP, fontSize: 11,
      bold: isHead, color: isHead ? C.white : C.text, valign: 'middle' });
  });

  infoBox(s, 'tip', '最終手段',
    'node_modules フォルダと package-lock.json を削除して npm cache clean --force → npm install を丸ごとやり直すと直ることが多い。',
    { y: 6.45, h: 0.75 }
  );
}

// ======================================================================
// 19. server.js 用意
// ======================================================================
{
  const s = addSlide(17, 'server.js の用意', 'まずは最小で動くサーバーを置く');
  bullets(s, [
    { text: 'VS Code でプロジェクトフォルダを開く', bold: true },
    { text: 'コマンドプロンプトから  code .  で現在フォルダを開く', sub: true },
    { text: '新規ファイル server.js を作成', bold: true },
    { text: 'エクスプローラーで右クリック → 新しいファイル → server.js', sub: true },
    { text: '以下のテンプレートを貼り付け → 保存（Ctrl + S）', sub: true },
  ], { h: 1.8 });

  cmdBlock(s, 'server.js — 最小テンプレート', [
    { text: '// 必要モジュールを読込', comment: true },
    "const express = require('express');",
    "const path    = require('path');",
    '',
    'const app = express();',
    'app.use(express.json());',
    '',
    { text: '// ヘルスチェック', comment: true },
    "app.get('/api/ping', (req, res) => {",
    "    res.json({ ok: true, time: new Date() });",
    '});',
    '',
    { text: '// 静的ファイル（index.html など）', comment: true },
    "app.use(express.static(path.join(__dirname)));",
    '',
    'const PORT = 3000;',
    'app.listen(PORT, () => {',
    "    console.log('Server running on http://localhost:' + PORT);",
    '});',
  ], { y: 3.6, h: 3.55, fontSize: 11 });
}

// ======================================================================
// 20. server.js + SQL 接続
// ======================================================================
{
  const s = addSlide(18, 'server.js — SQL Server 接続部分', 'msnodesqlv8 で Windows 認証接続する例');

  cmdBlock(s, 'server.js — DB 接続とテストエンドポイント', [
    { text: '// mssql + msnodesqlv8 ドライバを使う', comment: true },
    "const sql = require('mssql/msnodesqlv8');",
    '',
    { text: '// 接続設定（Windows 認証）', comment: true },
    'const dbConfig = {',
    "    server:   'localhost\\\\SQLEXPRESS',",
    "    database: 'AppDB',",
    "    driver:   'msnodesqlv8',",
    '    options: {',
    '        trustedConnection:      true,',
    '        trustServerCertificate: true,',
    '        enableArithAbort:       true',
    '    }',
    '};',
    '',
    { text: '// 接続プールを 1 つだけ作って使い回す', comment: true },
    'let pool;',
    'async function getPool() {',
    '    if (pool) return pool;',
    '    pool = await sql.connect(dbConfig);',
    '    return pool;',
    '}',
    '',
    { text: '// テスト用エンドポイント', comment: true },
    "app.get('/api/db-check', async (req, res) => {",
    '    try {',
    '        const p = await getPool();',
    "        const r = await p.request().query('SELECT GETDATE() AS now');",
    '        res.json({ ok: true, now: r.recordset[0].now });',
    '    } catch (e) {',
    '        res.status(500).json({ ok: false, error: e.message });',
    '    }',
    '});',
  ], { y: 1.55, h: 5.6, fontSize: 10 });
}

// ======================================================================
// 21. サーバー起動
// ======================================================================
{
  const s = addSlide(19, 'サーバー起動と停止', 'コマンドプロンプトから実行する');

  cmdBlock(s, 'Command Prompt — 起動 → 停止', [
    { text: 'rem プロジェクトフォルダへ移動', comment: true },
    'cd /d D:\\dev\\myapp',
    '',
    { text: 'rem サーバー起動（フォアグラウンド）', comment: true },
    'node server.js',
    { text: 'Server running on http://localhost:3000', comment: true },
    '',
    { text: 'rem 停止したいとき', comment: true },
    { text: 'rem  → コマンドプロンプトで Ctrl + C を押す', comment: true },
    { text: 'rem  → 「バッチ ジョブを終了しますか (Y/N)?」と出たら Y', comment: true },
    '',
    { text: 'rem nodemon 経由で起動（ファイル保存時に自動再起動）', comment: true },
    'npx nodemon server.js',
    '',
    { text: 'rem package.json に書いておくと npm start で起動できる', comment: true },
    'npm start',
  ], { y: 1.55, h: 4.1, fontSize: 12 });

  infoBox(s, 'tip', 'package.json への scripts 追加',
    '"scripts": { "start": "node server.js", "dev": "nodemon server.js" } と書くと、npm start / npm run dev で起動できる。',
    { y: 5.85, h: 0.85 }
  );
  infoBox(s, 'warn', 'ポート競合',
    'Error: listen EADDRINUSE :::3000 が出たら別プロセスが 3000 を使っている。 ' +
    'netstat -ano | findstr :3000 でプロセス特定 → taskkill /PID <番号> /F で停止。',
    { y: 6.8, h: 0.4 }
  );
}

// ======================================================================
// 22. 動作確認
// ======================================================================
{
  const s = addSlide(20, 'ブラウザと curl で動作確認', 'サーバーが本当に動いているかチェック');
  bullets(s, [
    { text: 'ブラウザで確認', bold: true },
    { text: 'http://localhost:3000/api/ping を開く', sub: true },
    { text: '{"ok":true,"time":"..."}   のような JSON が表示されれば成功', sub: true },
    { text: 'http://localhost:3000/api/db-check で SQL Server 接続も確認', sub: true },
  ], { h: 1.7 });

  cmdBlock(s, 'Command Prompt — curl で叩く（ブラウザなしで確認）', [
    { text: 'rem 別のコマンドプロンプトを開いて実行', comment: true },
    'curl http://localhost:3000/api/ping',
    { text: '{"ok":true,"time":"2026-04-14T10:00:00.000Z"}', comment: true },
    '',
    { text: 'rem JSON を整形して見たいとき（Node の JSON で整形）', comment: true },
    'curl -s http://localhost:3000/api/ping | node -e "let d=\'\';process.stdin.on(\'data\',x=>d+=x).on(\'end\',()=>console.log(JSON.stringify(JSON.parse(d),null,2)))"',
    '',
    { text: 'rem DB 接続テスト', comment: true },
    'curl http://localhost:3000/api/db-check',
  ], { y: 3.55, h: 3.0, fontSize: 11 });

  infoBox(s, 'info', 'LAN 内の別 PC から繋ぐ',
    'http://<サーバー PC 名 or IP>:3000/api/ping に置き換える。つながらない場合は Windows ファイアウォールで 3000 番受信許可を追加する。',
    { y: 6.7, h: 0.5 }
  );
}

// ======================================================================
// 23. 開発と本番の使い分け
// ======================================================================
{
  const s = addSlide(21, '開発フォルダと本番フォルダを分ける', '「作業用」と「実際に動かす場所」を混ぜないこと');

  // 左：開発フォルダ
  s.addShape('rect', { x: 1.6, y: 1.55, w: 5.4, h: 5.3, fill: { color: C.ice }, line: { color: C.blue, width: 1 } });
  s.addText('開発（作業）フォルダ', {
    x: 1.75, y: 1.65, w: 5.1, h: 0.4,
    fontFace: FONT_JPB, fontSize: 15, bold: true, color: C.navy
  });
  s.addText([
    { text: '場所の例\n', options: { bold: true, color: C.blue, fontSize: 12, fontFace: FONT_JPB } },
    { text: 'D:\\dev\\myapp\n\n', options: { color: C.text, fontSize: 12, fontFace: FONT_MONO } },
    { text: '役割\n', options: { bold: true, color: C.blue, fontSize: 12, fontFace: FONT_JPB } },
    { text: '・ソースコードを書く / 試す\n', options: { color: C.text, fontSize: 11, fontFace: FONT_JP } },
    { text: '・nodemon で動作確認\n', options: { color: C.text, fontSize: 11, fontFace: FONT_JP } },
    { text: '・git 管理する (任意)\n\n', options: { color: C.text, fontSize: 11, fontFace: FONT_JP } },
    { text: '特徴\n', options: { bold: true, color: C.blue, fontSize: 12, fontFace: FONT_JPB } },
    { text: '・node_modules は丸ごとある\n', options: { color: C.text, fontSize: 11, fontFace: FONT_JP } },
    { text: '・途中のテストファイルもある\n', options: { color: C.text, fontSize: 11, fontFace: FONT_JP } },
    { text: '・壊してもよい', options: { color: C.text, fontSize: 11, fontFace: FONT_JP } },
  ], { x: 1.85, y: 2.1, w: 5.0, h: 4.7, valign: 'top', margin: 0, paraSpaceAfter: 2 });

  // 矢印
  s.addShape('rightTriangle', {
    x: 7.1, y: 3.8, w: 0.4, h: 0.8,
    fill: { color: C.blue }, line: { color: C.blue }, rotate: 90
  });
  s.addText('コピー', {
    x: 6.95, y: 4.7, w: 0.8, h: 0.3,
    fontFace: FONT_JP, fontSize: 10, color: C.blue, align: 'center', bold: true
  });

  // 右：本番フォルダ
  s.addShape('rect', { x: 7.7, y: 1.55, w: 5.4, h: 5.3, fill: { color: 'ECFDF5' }, line: { color: C.green, width: 1 } });
  s.addText('本番（Host）フォルダ', {
    x: 7.85, y: 1.65, w: 5.1, h: 0.4,
    fontFace: FONT_JPB, fontSize: 15, bold: true, color: '065F46'
  });
  s.addText([
    { text: '場所の例\n', options: { bold: true, color: C.green, fontSize: 12, fontFace: FONT_JPB } },
    { text: 'C:\\apps\\myapp\n', options: { color: C.text, fontSize: 12, fontFace: FONT_MONO } },
    { text: '\\\\SERVER01\\apps\\myapp\n\n', options: { color: C.text, fontSize: 12, fontFace: FONT_MONO } },
    { text: '役割\n', options: { bold: true, color: C.green, fontSize: 12, fontFace: FONT_JPB } },
    { text: '・タスクスケジューラが指すフォルダ\n', options: { color: C.text, fontSize: 11, fontFace: FONT_JP } },
    { text: '・実ユーザーが利用している\n\n', options: { color: C.text, fontSize: 11, fontFace: FONT_JP } },
    { text: '特徴\n', options: { bold: true, color: C.green, fontSize: 12, fontFace: FONT_JPB } },
    { text: '・安定版だけを置く\n', options: { color: C.text, fontSize: 11, fontFace: FONT_JP } },
    { text: '・勝手に書き換えない\n', options: { color: C.text, fontSize: 11, fontFace: FONT_JP } },
    { text: '・更新前に必ずバックアップ', options: { color: C.text, fontSize: 11, fontFace: FONT_JP } },
  ], { x: 7.95, y: 2.1, w: 5.0, h: 4.7, valign: 'top', margin: 0, paraSpaceAfter: 2 });
}

// ======================================================================
// 24. コピー手順 - 手動
// ======================================================================
{
  const s = addSlide(22, 'ブラッシュアップ時のコピー手順 ①', '手動コピー（エクスプローラー + cmd）');
  bullets(s, [
    { text: '基本の流れ', bold: true, num: true },
    { text: '① 本番サーバーのサービス（タスクスケジューラ / node プロセス）を止める', sub: true },
    { text: '② 本番フォルダをバックアップ（後述）', sub: true },
    { text: '③ 作業フォルダから必要ファイルをコピー', sub: true },
    { text: '④ 必要なら本番で npm install', sub: true },
    { text: '⑤ サービスを再起動 → 動作確認', sub: true },
    { text: 'コピーしてよいもの', bold: true },
    { text: '・server.js / index.html / 他の .js / .html / 画像', sub: true },
    { text: '・package.json / package-lock.json', sub: true },
    { text: 'コピーしてはいけないもの', bold: true },
    { text: '・.env（接続情報などローカル設定が入っている）', sub: true },
    { text: '・node_modules（環境ごとに再生成すべき。LAN 越しコピーは非常に遅い）', sub: true },
    { text: '・ログファイル / 一時ファイル', sub: true },
  ]);
}

// ======================================================================
// 25. コピー手順 - cmd 手動
// ======================================================================
{
  const s = addSlide(23, 'ブラッシュアップ時のコピー手順 ②', 'cmd での手動コピー（少数ファイル向け）');

  cmdBlock(s, 'Command Prompt — copy / xcopy で個別コピー', [
    { text: 'rem 1) 本番フォルダに移動してバックアップを取る', comment: true },
    'cd /d C:\\apps\\myapp',
    'xcopy /E /I /Y . ..\\myapp_backup_2026-04-14',
    '',
    { text: 'rem 2) 単一ファイルのコピー（作業 → 本番）', comment: true },
    'copy /Y D:\\dev\\myapp\\server.js   C:\\apps\\myapp\\server.js',
    'copy /Y D:\\dev\\myapp\\index.html  C:\\apps\\myapp\\index.html',
    '',
    { text: 'rem 3) フォルダごとコピー（xcopy）', comment: true },
    { text: 'rem   /E  … サブフォルダ含む（空でも）', comment: true },
    { text: 'rem   /I  … 対象がフォルダと解釈', comment: true },
    { text: 'rem   /Y  … 上書き確認なし', comment: true },
    { text: 'rem   /EXCLUDE:list.txt … 除外ファイル一覧を使う', comment: true },
    'xcopy /E /I /Y D:\\dev\\myapp\\public C:\\apps\\myapp\\public',
    '',
    { text: 'rem 4) 結果確認', comment: true },
    'dir C:\\apps\\myapp',
  ], { y: 1.55, h: 5.6, fontSize: 11 });
}

// ======================================================================
// 26. robocopy
// ======================================================================
{
  const s = addSlide(24, 'ブラッシュアップ時のコピー手順 ③', 'robocopy で一括反映（推奨）');
  bullets(s, [
    { text: 'robocopy とは', bold: true },
    { text: 'Windows 標準の堅牢なコピーツール（Robust File Copy）', sub: true },
    { text: 'xcopy より速く、差分コピー・除外指定・ログ出力・中断再開に強い', sub: true },
    { text: '社内ツールの本番反映はこれ一択と言ってよい', sub: true },
  ], { h: 1.8 });

  cmdBlock(s, 'Command Prompt — robocopy 実行例', [
    { text: 'rem 差分コピーで本番フォルダへ反映', comment: true },
    'robocopy  D:\\dev\\myapp  C:\\apps\\myapp  /MIR /Z /R:2 /W:3 ^',
    '          /XD node_modules .git .vscode logs ^',
    '          /XF .env *.log *.tmp ^',
    '          /NFL /NDL /NP /TEE /LOG+:C:\\apps\\deploy.log',
    '',
    { text: 'rem 成功時の終了コード 0〜7 は正常、8 以上はエラー', comment: true },
    'echo %errorlevel%',
  ], { y: 3.5, h: 2.1, fontSize: 11 });

  infoBox(s, 'tip', '初回は /MIR を使わない',
    '/MIR はミラー（宛先にしか無いファイルを削除）。初回反映や本番の既存ファイルを残したいときは /E（サブフォルダ含めてコピー）に置き換える。',
    { y: 5.8, h: 1.3 }
  );
}

// ======================================================================
// 27. robocopy オプション詳細
// ======================================================================
{
  const s = addSlide(25, 'robocopy のオプションまとめ', '覚えておきたい主要オプション');

  const tbl = [
    ['オプション',       '意味'],
    ['/MIR',             'ミラーリング（宛先を source と同一にする、削除も含む）'],
    ['/E',               '空フォルダを含むサブフォルダ全体をコピー（削除はしない）'],
    ['/XD  <dir...>',    '指定フォルダを除外（例: node_modules .git）'],
    ['/XF  <file...>',   '指定ファイルを除外（例: *.log .env）'],
    ['/Z',               '再開可能モード（ネットワーク越しで途切れたとき有効）'],
    ['/R:n',             'エラー時のリトライ回数（既定 100 万回 → 2〜3 に下げると便利）'],
    ['/W:n',             'リトライの間隔（秒、既定 30）'],
    ['/NFL /NDL /NP',    'ファイル名・フォルダ名・進捗を出力しない（ログを軽くする）'],
    ['/TEE',             '画面とログの両方に出力'],
    ['/LOG+:path',       'ログファイルに追記（:path でなく /LOG:path なら上書き）'],
    ['/MT[:n]',          'マルチスレッドコピー（既定 8、高速化）'],
    ['/DCOPY:DAT',       'ファイル属性を含めてコピー（Data / Attributes / Timestamps）'],
    ['/L',               '実際にコピーせずに「コピーされる予定の一覧」だけ表示（下見用）'],
  ];
  const baseY = 1.55, rowH = 0.42;
  tbl.forEach((r, i) => {
    const y = baseY + i * rowH;
    const isHead = i === 0;
    s.addShape('rect', { x: 1.6, y, w: 11.5, h: rowH,
      fill: { color: isHead ? C.navy : (i % 2 === 0 ? C.grayBg : C.white) },
      line: { color: 'E5E7EB', width: 0.5 } });
    s.addText(r[0], { x: 1.7, y, w: 3.2, h: rowH, fontFace: FONT_MONO, fontSize: 11,
      bold: true, color: isHead ? C.white : C.blue, valign: 'middle' });
    s.addText(r[1], { x: 5.0, y, w: 8.0, h: rowH, fontFace: FONT_JP, fontSize: 11,
      bold: isHead, color: isHead ? C.white : C.text, valign: 'middle' });
  });
}

// ======================================================================
// 28. deploy.bat
// ======================================================================
{
  const s = addSlide(26, 'ブラッシュアップ時のコピー手順 ④', 'バッチファイル化（deploy.bat）で毎回同じ手順を再現');
  bullets(s, [
    { text: 'バッチファイルにすると何が嬉しい？', bold: true },
    { text: '毎回同じ robocopy / バックアップ手順を 1 クリックで実行できる', sub: true },
    { text: 'ミス（コピー漏れ・除外忘れ）を減らせる', sub: true },
    { text: '引き継ぎが楽', sub: true },
  ], { h: 1.6 });

  cmdBlock(s, 'deploy.bat — 本番反映スクリプト', [
    '@echo off',
    'chcp 65001 > nul',
    'setlocal',
    '',
    'set SRC=D:\\dev\\myapp',
    'set DST=C:\\apps\\myapp',
    'set BAK=C:\\apps\\backup\\myapp_%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%',
    '',
    'echo [1/4] 本番バックアップ: %BAK%',
    'robocopy "%DST%" "%BAK%" /E /NFL /NDL /NP > nul',
    '',
    'echo [2/4] 作業フォルダから本番へ反映',
    'robocopy "%SRC%" "%DST%" /MIR /Z /R:2 /W:3 ^',
    '    /XD node_modules .git .vscode logs ^',
    '    /XF .env *.log *.tmp ^',
    '    /NFL /NDL /NP /TEE /LOG+:C:\\apps\\deploy.log',
    '',
    'echo [3/4] 本番で npm install',
    'pushd "%DST%"',
    'call npm install --omit=dev',
    'popd',
    '',
    'echo [4/4] 完了。タスクスケジューラからサーバーを再起動してください。',
    'pause',
  ], { y: 3.15, h: 4.0, fontSize: 11 });
}

// ======================================================================
// 29. バックアップ / 再起動
// ======================================================================
{
  const s = addSlide(27, 'デプロイ後の再起動手順', '止める → 反映 → 起動 → 確認');
  bullets(s, [
    { text: '手順全体', bold: true },
    { text: '1. 利用者にアナウンス（数分だけ止まります）', sub: true },
    { text: '2. 本番サーバーにリモート接続 or 直接操作', sub: true },
    { text: '3. タスクスケジューラからタスクを停止', sub: true },
    { text: '4. deploy.bat を実行（または robocopy を手打ち）', sub: true },
    { text: '5. タスクを再度「実行」', sub: true },
    { text: '6. ブラウザで /api/ping を叩いて動作確認', sub: true },
  ], { h: 2.6 });

  cmdBlock(s, 'Command Prompt — 停止・起動のコマンド操作', [
    { text: 'rem 動いている node プロセスを確認', comment: true },
    'tasklist | findstr node',
    '',
    { text: 'rem 特定 PID のプロセスを停止', comment: true },
    'taskkill /PID 12345 /F',
    '',
    { text: 'rem すべての node プロセスを停止（注意）', comment: true },
    'taskkill /IM node.exe /F',
    '',
    { text: 'rem タスクスケジューラからの起動・停止（名前指定）', comment: true },
    'schtasks /End   /TN "MyAppServer"',
    'schtasks /Run   /TN "MyAppServer"',
    'schtasks /Query /TN "MyAppServer"',
  ], { y: 4.3, h: 2.85, fontSize: 11 });
}

// ======================================================================
// 30. トラブルシュート
// ======================================================================
{
  const s = addSlide(28, 'トラブルシューティング', 'よく出るエラーとコマンド対処集');

  const tbl = [
    ['現象', '確認コマンド / 対処'],
    ["'node' は内部コマンドまたは外部コマンド...",             'コマンドプロンプトを開き直す / 再ログインで PATH を再読込'],
    ['npm install が Permission denied で失敗',                '管理者で cmd を起動し直す。ウイルス対策ソフトの除外に dev フォルダを追加'],
    ['msnodesqlv8 の gyp / MSBuild エラー',                    'Visual Studio Build Tools を入れ直す。次に node_modules を削除して再インストール'],
    ['SQL Server に繋がらない（timeout）',                      'SQL Server 構成マネージャーで TCP/IP 有効化 → サービス再起動 → ファイアウォール確認'],
    ['Login failed for user ...',                                'Windows 認証で server.js を動かしているか確認。SQL 認証なら Login とパスワードを再確認'],
    ['listen EADDRINUSE :::3000',                                'netstat -ano | findstr :3000 → taskkill /PID <num> /F でプロセス停止'],
    ['robocopy が 16 を返して終了',                              '実行権限 / パスの存在を確認。宛先フォルダを手動で作成してからリトライ'],
    ['本番で動かない（画面真っ白）',                               'F12 DevTools の Console / Network を見る。静的ファイルのパスと apiUrl の書き換え忘れに注意'],
    ['反映したのに古い画面が出る',                                 'ブラウザキャッシュ。Ctrl + F5 で強制再読込 / 運用手順として案内する'],
  ];
  const baseY = 1.55, rowH = 0.57;
  tbl.forEach((r, i) => {
    const y = baseY + i * rowH;
    const isHead = i === 0;
    s.addShape('rect', { x: 1.6, y, w: 11.5, h: rowH,
      fill: { color: isHead ? C.navy : (i % 2 === 0 ? C.grayBg : C.white) },
      line: { color: 'E5E7EB', width: 0.5 } });
    s.addText(r[0], { x: 1.7, y, w: 4.6, h: rowH,
      fontFace: FONT_JP, fontSize: 11,
      bold: isHead, color: isHead ? C.white : C.red, valign: 'middle' });
    s.addText(r[1], { x: 6.35, y, w: 6.7, h: rowH,
      fontFace: FONT_JP, fontSize: 11,
      bold: isHead, color: isHead ? C.white : C.text, valign: 'middle' });
  });
}

// ======================================================================
// 31. まとめ
// ======================================================================
{
  const s = addSlide(29, 'まとめ', '手順を身体で覚えたら、迷わず運用できる');
  bullets(s, [
    { text: '① インストール順', bold: true },
    { text: 'Node.js（LTS＋Native Tools）→ VS Code → SQL Server Express → SSMS → Build Tools', sub: true },
    { text: '② SQL Server は TCP/IP とファイアウォールを開けないと繋がらない', bold: true },
    { text: 'SQL Server 構成マネージャーで有効化 → サービス再起動', sub: true },
    { text: '③ プロジェクトは npm init → npm install の順', bold: true },
    { text: 'msnodesqlv8 は特別扱い。Build Tools が必要なので失敗したら落ち着いて対処', sub: true },
    { text: '④ 開発フォルダと本番フォルダは必ず分ける', bold: true },
    { text: 'D:\\dev\\myapp ⇄ C:\\apps\\myapp（あるいは共有フォルダ）', sub: true },
    { text: '⑤ 反映は robocopy + deploy.bat で自動化', bold: true },
    { text: '毎回同じ手順を踏めばコピーミスが起きない', sub: true },
    { text: '⑥ 止める・反映・起動・確認をセットで覚える', bold: true },
    { text: 'schtasks または taskkill + 再起動 + ping で最短ルート', sub: true },
  ]);
}

// ======================================================================
// 32. クロージング
// ======================================================================
{
  const s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape('ellipse', { x: -2.0, y: 4.0, w: 8.0, h: 8.0, fill: { color: C.navy2 }, line: { color: C.navy2 } });
  s.addShape('ellipse', { x: 8.0, y: -3.0, w: 8.0, h: 8.0, fill: { color: '1A2457' }, line: { color: '1A2457' } });
  s.addShape('rect', { x: 0, y: 3.1, w: 13.333, h: 0.08, fill: { color: C.blueL }, line: { color: C.blueL } });

  s.addText('Ready to Deploy', {
    x: 0.8, y: 1.8, w: 12, h: 0.8,
    fontFace: FONT_TITLE, fontSize: 40, bold: true, color: C.white, italic: true
  });
  s.addText('あとは手を動かすだけ。', {
    x: 0.8, y: 3.3, w: 12, h: 0.6,
    fontFace: FONT_JPB, fontSize: 24, color: C.blueL
  });
  s.addText('この手順書を「読む」のではなく「使いながら覚える」のが一番の近道です。', {
    x: 0.8, y: 4.0, w: 12, h: 0.5,
    fontFace: FONT_JP, fontSize: 14, color: 'CBD5E1'
  });
  s.addText('— Happy Setup —', {
    x: 0.8, y: 5.5, w: 12, h: 0.5,
    fontFace: FONT_TITLE, fontSize: 16, color: C.blueL, italic: true
  });
}

// =========================================================
// 保存
// =========================================================
const outPath = path.resolve('C:\\Users\\Teppei\\Claude作業フォルダ\\作業時間一元管理\\WorkTimeApp\\開発手順書.pptx');
pres.writeFile({ fileName: outPath })
  .then(f => console.log('Saved:', f))
  .catch(err => { console.error(err); process.exit(1); });
