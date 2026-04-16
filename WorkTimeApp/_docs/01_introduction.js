// ============================================================
// 01_introduction.js — 導入告知資料（プレリリース案内）
//
// 目的: 社員向けに「作業時間一元管理システム」の導入を告知し、
//       前向きに受け入れてもらうためのメッセージ資料。
// ============================================================

const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10" × 5.625"
pres.author = "作業時間一元管理システム";
pres.title  = "作業時間 一元管理システム ご案内";

// ---- カラーパレット (Teal Trust + warm accents) ----
const C = {
    primary:   "028090", // deep teal
    primary2:  "00A896", // seafoam
    primary3:  "02C39A", // mint
    dark:      "0F3B3E", // very dark teal
    darkMid:   "164E52",
    light:     "F5FBFA", // near white with teal tint
    accent:    "F4A261", // warm orange (for highlights)
    accent2:   "E76F51", // terracotta
    textDark:  "1E293B",
    textMid:   "475569",
    textLight: "94A3B8",
    white:     "FFFFFF",
};

const FONT_HEADER = "Georgia";
const FONT_BODY   = "Calibri";

// ---- 共通ヘルパー ----
function addFooter(slide, pageNum, totalPages) {
    slide.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 5.4, w: 10, h: 0.225, fill: { color: C.primary }, line: { color: C.primary }
    });
    slide.addText("作業時間 一元管理システム  ご案内", {
        x: 0.4, y: 5.4, w: 6, h: 0.225,
        fontSize: 9, fontFace: FONT_BODY, color: C.white, valign: "middle", margin: 0
    });
    slide.addText(`${pageNum} / ${totalPages}`, {
        x: 9.0, y: 5.4, w: 0.8, h: 0.225,
        fontSize: 9, fontFace: FONT_BODY, color: C.white, valign: "middle", align: "right", margin: 0
    });
}

function addSectionLabel(slide, text, y = 0.55) {
    slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: y, w: 0.15, h: 0.35, fill: { color: C.accent }, line: { color: C.accent }
    });
    slide.addText(text, {
        x: 0.75, y: y, w: 9, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.primary,
        valign: "middle", charSpacing: 2, margin: 0
    });
}

function addTitle(slide, text, y = 0.95) {
    slide.addText(text, {
        x: 0.5, y: y, w: 9, h: 0.7,
        fontSize: 32, fontFace: FONT_HEADER, bold: true, color: C.dark,
        valign: "middle", margin: 0
    });
}

const TOTAL = 11;

// ============================================================
// Slide 1: タイトル
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // 背景の装飾: 大きな円（半透明）
    s.addShape(pres.shapes.OVAL, {
        x: -2, y: -2, w: 6, h: 6,
        fill: { color: C.primary2, transparency: 70 },
        line: { color: C.primary2, transparency: 70 }
    });
    s.addShape(pres.shapes.OVAL, {
        x: 7, y: 3, w: 5, h: 5,
        fill: { color: C.primary3, transparency: 80 },
        line: { color: C.primary3, transparency: 80 }
    });

    // ラベル
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 0.5, w: 0.15, h: 0.4, fill: { color: C.accent }, line: { color: C.accent }
    });
    s.addText("新システム導入のお知らせ", {
        x: 0.75, y: 0.5, w: 9, h: 0.4,
        fontSize: 13, fontFace: FONT_BODY, bold: true, color: C.accent,
        valign: "middle", charSpacing: 2, margin: 0
    });

    // メインタイトル
    s.addText("作業時間", {
        x: 0.5, y: 1.5, w: 9, h: 1.0,
        fontSize: 54, fontFace: FONT_HEADER, bold: true, color: C.white, margin: 0
    });
    s.addText("一元管理システム", {
        x: 0.5, y: 2.4, w: 9, h: 1.0,
        fontSize: 54, fontFace: FONT_HEADER, bold: true, color: C.primary3, margin: 0
    });

    // サブタイトル
    s.addText("みんなで豊かに、もっと効率よく働くために", {
        x: 0.5, y: 3.6, w: 9, h: 0.5,
        fontSize: 18, fontFace: FONT_BODY, italic: true, color: C.primary2, margin: 0
    });

    // 下部バー
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 5.2, w: 10, h: 0.425, fill: { color: C.primary }, line: { color: C.primary }
    });
    s.addText("社員のみなさまへ", {
        x: 0.5, y: 5.2, w: 9, h: 0.425,
        fontSize: 12, fontFace: FONT_BODY, color: C.white, valign: "middle", margin: 0
    });
}

// ============================================================
// Slide 2: はじめに
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addSectionLabel(s, "はじめに", 0.55);
    addTitle(s, "「作業時間 一元管理システム」とは？", 0.95);

    // メイン説明
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 1.95, w: 9, h: 1.4,
        fill: { color: C.white }, line: { color: C.primary3, width: 0 },
        shadow: { type: "outer", blur: 8, offset: 2, angle: 90, color: "000000", opacity: 0.08 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 1.95, w: 0.1, h: 1.4, fill: { color: C.primary }, line: { color: C.primary }
    });
    s.addText("日々の作業時間を、全部署で同じ形式で記録・集計できる社内アプリです。", {
        x: 0.8, y: 2.05, w: 8.6, h: 0.45,
        fontSize: 16, fontFace: FONT_BODY, bold: true, color: C.textDark, margin: 0
    });
    s.addText([
        { text: "これまでは部署ごとに Excel やノートで管理していた作業時間を、", options: { breakLine: true } },
        { text: "ひとつのシステムに集約することで「会社として見える化」します。" }
    ], {
        x: 0.8, y: 2.55, w: 8.6, h: 0.75,
        fontSize: 13, fontFace: FONT_BODY, color: C.textMid, margin: 0
    });

    // 3つのポイント
    const points = [
        { title: "スマホ・PCで", desc: "ブラウザから 1 分で入力" },
        { title: "部署を超えて", desc: "全社で同じ形式で集計" },
        { title: "あなたの時間を", desc: "自分で確認・分析できる" },
    ];
    const boxW = 2.9, boxH = 1.55, gap = 0.15, startX = 0.5;
    const startY = 3.55;
    points.forEach((p, i) => {
        const x = startX + i * (boxW + gap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: startY, w: boxW, h: boxH,
            fill: { color: C.white }, line: { color: C.primary2, width: 1.5 }
        });
        s.addShape(pres.shapes.OVAL, {
            x: x + boxW/2 - 0.25, y: startY + 0.2, w: 0.5, h: 0.5,
            fill: { color: C.primary }, line: { color: C.primary }
        });
        s.addText(String(i + 1), {
            x: x + boxW/2 - 0.25, y: startY + 0.2, w: 0.5, h: 0.5,
            fontSize: 18, fontFace: FONT_HEADER, bold: true, color: C.white,
            align: "center", valign: "middle", margin: 0
        });
        s.addText(p.title, {
            x: x + 0.15, y: startY + 0.8, w: boxW - 0.3, h: 0.3,
            fontSize: 13, fontFace: FONT_BODY, bold: true, color: C.primary,
            align: "center", margin: 0
        });
        s.addText(p.desc, {
            x: x + 0.15, y: startY + 1.1, w: boxW - 0.3, h: 0.35,
            fontSize: 12, fontFace: FONT_BODY, color: C.textMid,
            align: "center", margin: 0
        });
    });

    addFooter(s, 2, TOTAL);
}

// ============================================================
// Slide 3: なぜ導入するのか（メインメッセージ）
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // 装飾
    s.addShape(pres.shapes.OVAL, {
        x: 7, y: -1, w: 6, h: 6,
        fill: { color: C.primary2, transparency: 85 },
        line: { color: C.primary2, transparency: 85 }
    });

    // ラベル
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 0.55, w: 0.15, h: 0.35, fill: { color: C.accent }, line: { color: C.accent }
    });
    s.addText("なぜ導入するのか", {
        x: 0.75, y: 0.55, w: 9, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.accent,
        valign: "middle", charSpacing: 2, margin: 0
    });

    // メインメッセージ
    s.addText("会社も、個人も、", {
        x: 0.5, y: 1.1, w: 9, h: 0.9,
        fontSize: 40, fontFace: FONT_HEADER, bold: true, color: C.white, margin: 0
    });
    s.addText("もっと豊かになるために。", {
        x: 0.5, y: 1.95, w: 9, h: 0.9,
        fontSize: 40, fontFace: FONT_HEADER, bold: true, color: C.primary3, margin: 0
    });

    // サブメッセージ
    s.addShape(pres.shapes.LINE, {
        x: 0.5, y: 3.15, w: 0.8, h: 0,
        line: { color: C.accent, width: 2 }
    });
    s.addText([
        { text: "時間に追われるのではなく、", options: { breakLine: true } },
        { text: "時間を味方にする働き方へ。" }
    ], {
        x: 0.5, y: 3.3, w: 9, h: 1.0,
        fontSize: 18, fontFace: FONT_BODY, italic: true, color: C.primary2, margin: 0
    });

    // タグ
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 4.55, w: 3.2, h: 0.5,
        fill: { color: C.primary }, line: { color: C.primary }
    });
    s.addText("生産性UP = みんなの余裕", {
        x: 0.5, y: 4.55, w: 3.2, h: 0.5,
        fontSize: 13, fontFace: FONT_BODY, bold: true, color: C.white,
        align: "center", valign: "middle", margin: 0
    });

    addFooter(s, 3, TOTAL);
}

// ============================================================
// Slide 4: 4つのメリット（一覧）
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addSectionLabel(s, "このシステムでできること", 0.55);
    addTitle(s, "4つのメリット", 0.95);

    const benefits = [
        { no: "01", title: "見える化", sub: "時間の使い方が", sub2: "はっきり分かる", color: C.primary },
        { no: "02", title: "自分で目標設定", sub: "ノルマではなく", sub2: "自分のための指標", color: C.primary2 },
        { no: "03", title: "見積もり精度UP", sub: "過去の実績から", sub2: "適正な見積もりへ", color: C.primary3 },
        { no: "04", title: "生産性UP", sub: "ゆとりある時間を", sub2: "自分のものに", color: C.accent },
    ];
    const boxW = 2.2, boxH = 2.7, gap = 0.15;
    const startX = (10 - (boxW * 4 + gap * 3)) / 2;
    const startY = 2.1;
    benefits.forEach((b, i) => {
        const x = startX + i * (boxW + gap);
        // Card
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: startY, w: boxW, h: boxH,
            fill: { color: C.white }, line: { color: "E2E8F0", width: 1 },
            shadow: { type: "outer", blur: 10, offset: 3, angle: 90, color: "000000", opacity: 0.06 }
        });
        // Colored top bar
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: startY, w: boxW, h: 0.15,
            fill: { color: b.color }, line: { color: b.color }
        });
        // Number
        s.addText(b.no, {
            x: x, y: startY + 0.3, w: boxW, h: 0.55,
            fontSize: 32, fontFace: FONT_HEADER, bold: true, color: b.color,
            align: "center", margin: 0
        });
        // Title
        s.addText(b.title, {
            x: x, y: startY + 0.95, w: boxW, h: 0.5,
            fontSize: 18, fontFace: FONT_BODY, bold: true, color: C.textDark,
            align: "center", margin: 0
        });
        // Divider
        s.addShape(pres.shapes.LINE, {
            x: x + boxW/2 - 0.3, y: startY + 1.55, w: 0.6, h: 0,
            line: { color: b.color, width: 1.5 }
        });
        // Sub
        s.addText(b.sub, {
            x: x + 0.1, y: startY + 1.75, w: boxW - 0.2, h: 0.35,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid,
            align: "center", margin: 0
        });
        s.addText(b.sub2, {
            x: x + 0.1, y: startY + 2.1, w: boxW - 0.2, h: 0.35,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid,
            align: "center", margin: 0
        });
    });

    // 下部メッセージ
    s.addText("→ 次ページから 1 つずつご紹介します", {
        x: 0.5, y: 5.0, w: 9, h: 0.3,
        fontSize: 12, fontFace: FONT_BODY, italic: true, color: C.textMid,
        align: "center", margin: 0
    });

    addFooter(s, 4, TOTAL);
}

// ============================================================
// Slide 5: メリット1 - 見える化
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addSectionLabel(s, "メリット 01", 0.55);
    addTitle(s, "時間の使い方が「見える」ようになる", 0.95);

    // 左: 説明
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.0, w: 4.6, h: 3.1,
        fill: { color: C.white }, line: { color: "E2E8F0", width: 1 },
        shadow: { type: "outer", blur: 8, offset: 2, angle: 90, color: "000000", opacity: 0.08 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.0, w: 0.12, h: 3.1, fill: { color: C.primary }, line: { color: C.primary }
    });

    s.addText("「忙しい」を数字で見る", {
        x: 0.8, y: 2.15, w: 4.2, h: 0.4,
        fontSize: 16, fontFace: FONT_BODY, bold: true, color: C.primary, margin: 0
    });

    const leftItems = [
        { t: "どの仕事に", d: "何時間使ったか一目瞭然" },
        { t: "会議や雑務に", d: "どれだけ取られているか" },
        { t: "本来の業務は", d: "どれだけ進んでいるか" },
    ];
    leftItems.forEach((it, i) => {
        const y = 2.65 + i * 0.75;
        s.addShape(pres.shapes.OVAL, {
            x: 0.8, y: y + 0.08, w: 0.22, h: 0.22,
            fill: { color: C.primary3 }, line: { color: C.primary3 }
        });
        s.addText(it.t, {
            x: 1.1, y: y, w: 3.9, h: 0.3,
            fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.textDark, margin: 0
        });
        s.addText(it.d, {
            x: 1.1, y: y + 0.3, w: 3.9, h: 0.3,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    // 右: 大きな統計ビジュアル
    s.addShape(pres.shapes.RECTANGLE, {
        x: 5.4, y: 2.0, w: 4.1, h: 3.1,
        fill: { color: C.primary }, line: { color: C.primary },
        shadow: { type: "outer", blur: 8, offset: 2, angle: 90, color: "000000", opacity: 0.15 }
    });
    s.addText("自分の時間を", {
        x: 5.4, y: 2.2, w: 4.1, h: 0.4,
        fontSize: 13, fontFace: FONT_BODY, color: C.primary3,
        align: "center", margin: 0
    });
    s.addText("取り戻そう", {
        x: 5.4, y: 2.6, w: 4.1, h: 0.6,
        fontSize: 22, fontFace: FONT_HEADER, bold: true, color: C.white,
        align: "center", margin: 0
    });
    s.addShape(pres.shapes.LINE, {
        x: 6.9, y: 3.35, w: 1.1, h: 0,
        line: { color: C.accent, width: 2 }
    });
    s.addText("1日", {
        x: 5.4, y: 3.5, w: 4.1, h: 0.4,
        fontSize: 12, fontFace: FONT_BODY, color: C.primary3,
        align: "center", margin: 0
    });
    s.addText("8時間", {
        x: 5.4, y: 3.85, w: 4.1, h: 0.8,
        fontSize: 48, fontFace: FONT_HEADER, bold: true, color: C.white,
        align: "center", margin: 0
    });
    s.addText("あなたの大切な時間", {
        x: 5.4, y: 4.7, w: 4.1, h: 0.3,
        fontSize: 11, fontFace: FONT_BODY, italic: true, color: C.primary3,
        align: "center", margin: 0
    });

    addFooter(s, 5, TOTAL);
}

// ============================================================
// Slide 6: メリット2 - 自分で目標時間を決める
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addSectionLabel(s, "メリット 02", 0.55);
    addTitle(s, "目標は「自分で」決める", 0.95);

    // リード文
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 1.85, w: 9, h: 0.7,
        fill: { color: C.primary2 }, line: { color: C.primary2 }
    });
    s.addText("上からノルマが降ってくるのではなく、自分の裁量で目標時間を設定できます", {
        x: 0.5, y: 1.85, w: 9, h: 0.7,
        fontSize: 15, fontFace: FONT_BODY, bold: true, color: C.white,
        align: "center", valign: "middle", margin: 0
    });

    // 左右比較: 従来 vs これから
    const leftX = 0.5, rightX = 5.2, boxW = 4.3, boxY = 2.8, boxH = 2.3;

    // 左: 従来
    s.addShape(pres.shapes.RECTANGLE, {
        x: leftX, y: boxY, w: boxW, h: boxH,
        fill: { color: "F1F5F9" }, line: { color: "CBD5E1", width: 1 }
    });
    s.addText("これまで", {
        x: leftX + 0.2, y: boxY + 0.2, w: boxW - 0.4, h: 0.35,
        fontSize: 11, fontFace: FONT_BODY, bold: true, color: C.textLight,
        charSpacing: 2, margin: 0
    });
    s.addText("上からのノルマ", {
        x: leftX + 0.2, y: boxY + 0.55, w: boxW - 0.4, h: 0.45,
        fontSize: 18, fontFace: FONT_HEADER, bold: true, color: C.textMid, margin: 0
    });
    const oldItems = [
        "誰かに決められた時間",
        "達成のプレッシャー",
        "やらされ感",
    ];
    oldItems.forEach((it, i) => {
        s.addText("×", {
            x: leftX + 0.25, y: boxY + 1.15 + i * 0.35, w: 0.3, h: 0.3,
            fontSize: 14, fontFace: FONT_BODY, bold: true, color: C.textLight, margin: 0
        });
        s.addText(it, {
            x: leftX + 0.55, y: boxY + 1.15 + i * 0.35, w: boxW - 0.75, h: 0.3,
            fontSize: 12, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    // 右: これから
    s.addShape(pres.shapes.RECTANGLE, {
        x: rightX, y: boxY, w: boxW, h: boxH,
        fill: { color: C.white }, line: { color: C.primary, width: 2 },
        shadow: { type: "outer", blur: 8, offset: 2, angle: 90, color: "000000", opacity: 0.10 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
        x: rightX, y: boxY, w: boxW, h: 0.15,
        fill: { color: C.primary }, line: { color: C.primary }
    });
    s.addText("これから", {
        x: rightX + 0.2, y: boxY + 0.25, w: boxW - 0.4, h: 0.35,
        fontSize: 11, fontFace: FONT_BODY, bold: true, color: C.primary,
        charSpacing: 2, margin: 0
    });
    s.addText("自分のための指標", {
        x: rightX + 0.2, y: boxY + 0.6, w: boxW - 0.4, h: 0.45,
        fontSize: 18, fontFace: FONT_HEADER, bold: true, color: C.textDark, margin: 0
    });
    const newItems = [
        "自分で決めた目標",
        "効率よく仕事を進める指針",
        "達成したときの達成感",
    ];
    newItems.forEach((it, i) => {
        s.addText("✓", {
            x: rightX + 0.25, y: boxY + 1.2 + i * 0.35, w: 0.3, h: 0.3,
            fontSize: 14, fontFace: FONT_BODY, bold: true, color: C.primary3, margin: 0
        });
        s.addText(it, {
            x: rightX + 0.55, y: boxY + 1.2 + i * 0.35, w: boxW - 0.75, h: 0.3,
            fontSize: 12, fontFace: FONT_BODY, color: C.textDark, margin: 0
        });
    });

    addFooter(s, 6, TOTAL);
}

// ============================================================
// Slide 7: メリット3 - 見積もり精度UP
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addSectionLabel(s, "メリット 03", 0.55);
    addTitle(s, "次の見積もりが「根拠を持って」出せる", 0.95);

    // リード
    s.addText("過去の実績データが溜まるほど、見積もりの精度は上がります。", {
        x: 0.5, y: 1.9, w: 9, h: 0.4,
        fontSize: 14, fontFace: FONT_BODY, color: C.textMid, margin: 0
    });

    // フロー図
    const flowY = 2.7, flowH = 2.2;
    const steps = [
        { label: "記録する", desc: "毎日の作業時間", color: C.primary },
        { label: "集計する", desc: "注番・期別に", color: C.primary2 },
        { label: "分析する", desc: "傾向が見える", color: C.primary3 },
        { label: "見積もる", desc: "根拠ある数字を", color: C.accent },
    ];
    const stepW = 1.9, gap = 0.25;
    const totalW = stepW * 4 + gap * 3;
    const startXFlow = (10 - totalW) / 2;
    steps.forEach((st, i) => {
        const x = startXFlow + i * (stepW + gap);
        // Box
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: flowY, w: stepW, h: flowH,
            fill: { color: C.white }, line: { color: st.color, width: 2 },
            shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.08 }
        });
        // Number circle
        s.addShape(pres.shapes.OVAL, {
            x: x + stepW/2 - 0.3, y: flowY + 0.25, w: 0.6, h: 0.6,
            fill: { color: st.color }, line: { color: st.color }
        });
        s.addText(String(i + 1), {
            x: x + stepW/2 - 0.3, y: flowY + 0.25, w: 0.6, h: 0.6,
            fontSize: 22, fontFace: FONT_HEADER, bold: true, color: C.white,
            align: "center", valign: "middle", margin: 0
        });
        // Label
        s.addText(st.label, {
            x: x, y: flowY + 1.0, w: stepW, h: 0.4,
            fontSize: 15, fontFace: FONT_BODY, bold: true, color: C.textDark,
            align: "center", margin: 0
        });
        // Desc
        s.addText(st.desc, {
            x: x + 0.1, y: flowY + 1.45, w: stepW - 0.2, h: 0.6,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid,
            align: "center", margin: 0
        });

        // Arrow between steps
        if (i < steps.length - 1) {
            const arrowX = x + stepW + 0.02;
            s.addText("→", {
                x: arrowX, y: flowY + 0.7, w: gap + 0.1, h: 0.6,
                fontSize: 24, fontFace: FONT_BODY, bold: true, color: C.textLight,
                align: "center", valign: "middle", margin: 0
            });
        }
    });

    addFooter(s, 7, TOTAL);
}

// ============================================================
// Slide 8: メリット4 - 生産性UP
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addSectionLabel(s, "メリット 04", 0.55);
    addTitle(s, "時間を意識する = ゆとりを生み出す", 0.95);

    // 大きな数式風ビジュアル
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 1.9, w: 9, h: 1.3,
        fill: { color: C.primary }, line: { color: C.primary },
        shadow: { type: "outer", blur: 10, offset: 3, angle: 90, color: "000000", opacity: 0.15 }
    });

    s.addText("時間を意識する", {
        x: 0.6, y: 2.05, w: 2.8, h: 1.0,
        fontSize: 18, fontFace: FONT_BODY, bold: true, color: C.white,
        align: "center", valign: "middle", margin: 0
    });
    s.addText("+", {
        x: 3.4, y: 2.05, w: 0.5, h: 1.0,
        fontSize: 32, fontFace: FONT_HEADER, bold: true, color: C.primary3,
        align: "center", valign: "middle", margin: 0
    });
    s.addText("効率よく働く", {
        x: 3.9, y: 2.05, w: 2.8, h: 1.0,
        fontSize: 18, fontFace: FONT_BODY, bold: true, color: C.white,
        align: "center", valign: "middle", margin: 0
    });
    s.addText("=", {
        x: 6.7, y: 2.05, w: 0.5, h: 1.0,
        fontSize: 32, fontFace: FONT_HEADER, bold: true, color: C.primary3,
        align: "center", valign: "middle", margin: 0
    });
    s.addText("余白 & 笑顔", {
        x: 7.2, y: 2.05, w: 2.3, h: 1.0,
        fontSize: 20, fontFace: FONT_HEADER, bold: true, color: C.accent,
        align: "center", valign: "middle", margin: 0
    });

    // 下部: 余白 = 何に使う？
    s.addText("その余白は、あなたのものです。", {
        x: 0.5, y: 3.55, w: 9, h: 0.4,
        fontSize: 15, fontFace: FONT_BODY, italic: true, color: C.primary, margin: 0
    });

    const uses = [
        { t: "家族", d: "との時間" },
        { t: "学び", d: "新しいスキル" },
        { t: "休息", d: "しっかり休む" },
        { t: "挑戦", d: "新しい企画" },
    ];
    const useW = 2.05, useGap = 0.2;
    const useStartX = (10 - (useW * 4 + useGap * 3)) / 2;
    uses.forEach((u, i) => {
        const x = useStartX + i * (useW + useGap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: 4.1, w: useW, h: 1.0,
            fill: { color: C.white }, line: { color: C.primary2, width: 1.5 }
        });
        s.addText(u.t, {
            x: x, y: 4.2, w: useW, h: 0.4,
            fontSize: 18, fontFace: FONT_HEADER, bold: true, color: C.primary,
            align: "center", margin: 0
        });
        s.addText(u.d, {
            x: x, y: 4.65, w: useW, h: 0.35,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid,
            align: "center", margin: 0
        });
    });

    addFooter(s, 8, TOTAL);
}

// ============================================================
// Slide 9: よくある心配への回答
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addSectionLabel(s, "みなさまの心配ごと", 0.55);
    addTitle(s, "「こう思うかもしれません…」にお答えします", 0.95);

    const qas = [
        {
            q: "監視されているみたいで嫌…",
            a: "個人の時間を責めるためのものではありません。会社全体の傾向を掴んで、無理のない仕事配分をするためのツールです。"
        },
        {
            q: "入力が面倒くさそう…",
            a: "1 日の入力は 30 秒〜 1 分程度。前日のコピペも可能です。忘れても翌日のホーム画面でリマインドします。"
        },
        {
            q: "評価に使われるの？",
            a: "直接的な人事評価には使いません。「頑張っているのに時間が足りない」を可視化して、業務を調整するためのデータです。"
        },
    ];
    const boxY = 1.95, boxH = 1.0, gap = 0.15;
    qas.forEach((qa, i) => {
        const y = boxY + i * (boxH + gap);

        // Q マーク
        s.addShape(pres.shapes.OVAL, {
            x: 0.5, y: y + 0.1, w: 0.55, h: 0.55,
            fill: { color: C.accent }, line: { color: C.accent }
        });
        s.addText("Q", {
            x: 0.5, y: y + 0.1, w: 0.55, h: 0.55,
            fontSize: 18, fontFace: FONT_HEADER, bold: true, color: C.white,
            align: "center", valign: "middle", margin: 0
        });

        // Background card
        s.addShape(pres.shapes.RECTANGLE, {
            x: 1.2, y: y, w: 8.3, h: boxH,
            fill: { color: C.white }, line: { color: "E2E8F0", width: 1 },
            shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.06 }
        });

        // Question
        s.addText(qa.q, {
            x: 1.35, y: y + 0.1, w: 8.05, h: 0.3,
            fontSize: 13, fontFace: FONT_BODY, bold: true, color: C.textDark, margin: 0
        });

        // A marker
        s.addText("A.", {
            x: 1.35, y: y + 0.45, w: 0.4, h: 0.45,
            fontSize: 13, fontFace: FONT_HEADER, bold: true, color: C.primary, margin: 0
        });

        // Answer
        s.addText(qa.a, {
            x: 1.75, y: y + 0.45, w: 7.65, h: 0.5,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    addFooter(s, 9, TOTAL);
}

// ============================================================
// Slide 10: 運用スケジュール
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addSectionLabel(s, "運用について", 0.55);
    addTitle(s, "スケジュールとサポート", 0.95);

    // タイムライン
    s.addShape(pres.shapes.LINE, {
        x: 1.2, y: 2.6, w: 7.6, h: 0,
        line: { color: C.primary2, width: 3 }
    });

    const milestones = [
        { when: "プレリリース",  what: "一部ユーザーで試用",  color: C.primary2 },
        { when: "本格リリース",   what: "全社で利用開始",      color: C.primary },
        { when: "運用開始",       what: "毎日の入力習慣化",    color: C.primary3 },
        { when: "3ヶ月後",        what: "データ活用スタート",  color: C.accent },
    ];
    milestones.forEach((m, i) => {
        const x = 1.2 + (7.6 / 3) * i;
        // Dot
        s.addShape(pres.shapes.OVAL, {
            x: x - 0.2, y: 2.4, w: 0.4, h: 0.4,
            fill: { color: m.color }, line: { color: C.white, width: 2 }
        });
        // Label above
        s.addText(m.when, {
            x: x - 1.3, y: 1.9, w: 2.6, h: 0.35,
            fontSize: 12, fontFace: FONT_BODY, bold: true, color: m.color,
            align: "center", margin: 0
        });
        // Desc below
        s.addText(m.what, {
            x: x - 1.3, y: 2.95, w: 2.6, h: 0.35,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid,
            align: "center", margin: 0
        });
    });

    // サポート窓口
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 3.8, w: 9, h: 1.3,
        fill: { color: C.white }, line: { color: C.primary3, width: 1 },
        shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.06 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 3.8, w: 0.15, h: 1.3,
        fill: { color: C.primary3 }, line: { color: C.primary3 }
    });
    s.addText("困ったときは", {
        x: 0.8, y: 3.9, w: 8.5, h: 0.35,
        fontSize: 11, fontFace: FONT_BODY, bold: true, color: C.primary,
        charSpacing: 2, margin: 0
    });
    s.addText("いつでも管理者・IT担当者にご相談ください", {
        x: 0.8, y: 4.25, w: 8.5, h: 0.4,
        fontSize: 16, fontFace: FONT_BODY, bold: true, color: C.textDark, margin: 0
    });
    s.addText("操作マニュアルは別途配布します。分からないことは、遠慮なく聞いてください。", {
        x: 0.8, y: 4.65, w: 8.5, h: 0.35,
        fontSize: 11, fontFace: FONT_BODY, color: C.textMid, margin: 0
    });

    addFooter(s, 10, TOTAL);
}

// ============================================================
// Slide 11: クロージング
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // 装飾
    s.addShape(pres.shapes.OVAL, {
        x: -3, y: 2, w: 7, h: 7,
        fill: { color: C.primary, transparency: 80 },
        line: { color: C.primary, transparency: 80 }
    });
    s.addShape(pres.shapes.OVAL, {
        x: 7, y: -2, w: 6, h: 6,
        fill: { color: C.primary3, transparency: 85 },
        line: { color: C.primary3, transparency: 85 }
    });

    // ラベル
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 0.55, w: 0.15, h: 0.35, fill: { color: C.accent }, line: { color: C.accent }
    });
    s.addText("最後に", {
        x: 0.75, y: 0.55, w: 9, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.accent,
        valign: "middle", charSpacing: 2, margin: 0
    });

    // メインメッセージ
    s.addText("一緒に、", {
        x: 0.5, y: 1.3, w: 9, h: 0.9,
        fontSize: 44, fontFace: FONT_HEADER, bold: true, color: C.white, margin: 0
    });
    s.addText("いい働き方を", {
        x: 0.5, y: 2.2, w: 9, h: 0.9,
        fontSize: 44, fontFace: FONT_HEADER, bold: true, color: C.primary3, margin: 0
    });
    s.addText("作っていきましょう。", {
        x: 0.5, y: 3.1, w: 9, h: 0.9,
        fontSize: 44, fontFace: FONT_HEADER, bold: true, color: C.primary3, margin: 0
    });

    // ハイライト
    s.addShape(pres.shapes.LINE, {
        x: 0.5, y: 4.3, w: 0.8, h: 0,
        line: { color: C.accent, width: 2 }
    });
    s.addText("効率よく、豊かに、笑顔で。", {
        x: 0.5, y: 4.45, w: 9, h: 0.4,
        fontSize: 16, fontFace: FONT_BODY, italic: true, color: C.primary2, margin: 0
    });

    // 下部
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 5.2, w: 10, h: 0.425, fill: { color: C.primary }, line: { color: C.primary }
    });
    s.addText("作業時間 一元管理システム  —  社員のみなさまへ", {
        x: 0.5, y: 5.2, w: 9, h: 0.425,
        fontSize: 11, fontFace: FONT_BODY, color: C.white, valign: "middle", margin: 0
    });
}

// ============================================================
// 書き出し
// ============================================================
pres.writeFile({ fileName: "01_導入告知資料.pptx" })
    .then(fileName => console.log(`✓ ${fileName} を作成しました`))
    .catch(err => { console.error(err); process.exit(1); });
