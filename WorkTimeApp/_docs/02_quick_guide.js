// ============================================================
// 02_quick_guide.js — 簡易操作手順書
//
// 目的: 社員が「これだけ覚えれば使える」最低限の操作を学べる
//       8ページの簡潔なガイド。
// ============================================================

const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10" × 5.625"
pres.author = "作業時間一元管理システム";
pres.title  = "作業時間 一元管理システム 簡易操作手順書";

// ---- カラーパレット (Ocean Gradient) ----
const C = {
    primary:   "065A82", // deep blue
    primary2:  "1C7293", // teal-blue
    primary3:  "9EB3C2", // light blue-gray
    dark:      "21295C", // midnight
    darker:    "0B1340",
    light:     "F8FAFC", // near white
    lightBg:   "E8F1F8", // very light blue
    accent:    "F59E0B", // amber
    accent2:   "10B981", // green (for success/check)
    textDark:  "1E293B",
    textMid:   "475569",
    textLight: "94A3B8",
    white:     "FFFFFF",
};

const FONT_HEADER = "Calibri";
const FONT_BODY   = "Calibri";

const TOTAL = 8;

// ---- 共通ヘルパー ----
function addFooter(slide, pageNum) {
    slide.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 5.4, w: 10, h: 0.225, fill: { color: C.primary }, line: { color: C.primary }
    });
    slide.addText("簡易操作手順書  —  作業時間 一元管理システム", {
        x: 0.4, y: 5.4, w: 6, h: 0.225,
        fontSize: 9, fontFace: FONT_BODY, color: C.white, valign: "middle", margin: 0
    });
    slide.addText(`${pageNum} / ${TOTAL}`, {
        x: 9.0, y: 5.4, w: 0.8, h: 0.225,
        fontSize: 9, fontFace: FONT_BODY, color: C.white, valign: "middle", align: "right", margin: 0
    });
}

function addStepHeader(slide, stepNum, stepTitle) {
    // 上部帯
    slide.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 0, w: 10, h: 1.3, fill: { color: C.primary }, line: { color: C.primary }
    });
    // ステップバッジ
    slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 0.3, w: 1.4, h: 0.7,
        fill: { color: C.accent }, line: { color: C.accent }
    });
    slide.addText(`STEP ${stepNum}`, {
        x: 0.5, y: 0.3, w: 1.4, h: 0.7,
        fontSize: 18, fontFace: FONT_HEADER, bold: true, color: C.white,
        align: "center", valign: "middle", margin: 0
    });
    // タイトル
    slide.addText(stepTitle, {
        x: 2.1, y: 0.3, w: 7.5, h: 0.7,
        fontSize: 24, fontFace: FONT_HEADER, bold: true, color: C.white,
        valign: "middle", margin: 0
    });
}

// ============================================================
// Slide 1: タイトル
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // 装飾
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 4.4, w: 10, h: 1.225,
        fill: { color: C.primary }, line: { color: C.primary }
    });
    s.addShape(pres.shapes.OVAL, {
        x: 6.5, y: -1, w: 5, h: 5,
        fill: { color: C.primary2, transparency: 80 },
        line: { color: C.primary2, transparency: 80 }
    });
    s.addShape(pres.shapes.OVAL, {
        x: -2, y: 2, w: 4, h: 4,
        fill: { color: C.primary, transparency: 70 },
        line: { color: C.primary, transparency: 70 }
    });

    // ラベル
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 0.5, w: 0.15, h: 0.4,
        fill: { color: C.accent }, line: { color: C.accent }
    });
    s.addText("QUICK GUIDE", {
        x: 0.75, y: 0.5, w: 9, h: 0.4,
        fontSize: 14, fontFace: FONT_HEADER, bold: true, color: C.accent,
        valign: "middle", charSpacing: 3, margin: 0
    });

    // タイトル
    s.addText("簡易操作手順書", {
        x: 0.5, y: 1.5, w: 9, h: 1.1,
        fontSize: 52, fontFace: FONT_HEADER, bold: true, color: C.white, margin: 0
    });
    s.addText("作業時間 一元管理システム", {
        x: 0.5, y: 2.65, w: 9, h: 0.5,
        fontSize: 20, fontFace: FONT_BODY, color: C.primary3, margin: 0
    });

    // サブ情報
    s.addShape(pres.shapes.LINE, {
        x: 0.5, y: 3.5, w: 0.8, h: 0,
        line: { color: C.accent, width: 2 }
    });
    s.addText("4 STEP、5分で覚えられます", {
        x: 0.5, y: 3.6, w: 9, h: 0.4,
        fontSize: 16, fontFace: FONT_BODY, italic: true, color: C.primary3, margin: 0
    });

    // 下部ラベル
    s.addText("これだけ覚えれば、毎日使えます", {
        x: 0.5, y: 4.65, w: 9, h: 0.7,
        fontSize: 14, fontFace: FONT_BODY, bold: true, color: C.white,
        valign: "middle", margin: 0
    });
}

// ============================================================
// Slide 2: STEP 1 — ログイン
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addStepHeader(s, 1, "ログインする");

    // 左: 操作手順
    s.addText("操作手順", {
        x: 0.5, y: 1.6, w: 4.5, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.primary,
        charSpacing: 2, margin: 0
    });

    const steps = [
        { n: "1", t: "ブラウザで URL を開く", d: "http://192.168.1.8:3000/" },
        { n: "2", t: "自分の名前を選ぶ", d: "ユーザー一覧から" },
        { n: "3", t: "4 桁の PIN を入力", d: "初回は自分で設定" },
        { n: "4", t: "ログイン完了", d: "ホーム画面が表示" },
    ];
    steps.forEach((st, i) => {
        const y = 2.0 + i * 0.7;
        // Number circle
        s.addShape(pres.shapes.OVAL, {
            x: 0.5, y: y, w: 0.5, h: 0.5,
            fill: { color: C.primary }, line: { color: C.primary }
        });
        s.addText(st.n, {
            x: 0.5, y: y, w: 0.5, h: 0.5,
            fontSize: 18, fontFace: FONT_HEADER, bold: true, color: C.white,
            align: "center", valign: "middle", margin: 0
        });
        // Title
        s.addText(st.t, {
            x: 1.15, y: y + 0.02, w: 3.5, h: 0.3,
            fontSize: 13, fontFace: FONT_BODY, bold: true, color: C.textDark, margin: 0
        });
        // Desc
        s.addText(st.d, {
            x: 1.15, y: y + 0.28, w: 3.5, h: 0.25,
            fontSize: 10, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    // 右: 注意ボックス
    s.addShape(pres.shapes.RECTANGLE, {
        x: 5.3, y: 1.6, w: 4.2, h: 3.5,
        fill: { color: C.white }, line: { color: C.primary2, width: 1.5 },
        shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.08 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
        x: 5.3, y: 1.6, w: 4.2, h: 0.5,
        fill: { color: C.primary2 }, line: { color: C.primary2 }
    });
    s.addText("ポイント", {
        x: 5.5, y: 1.6, w: 4, h: 0.5,
        fontSize: 13, fontFace: FONT_BODY, bold: true, color: C.white,
        valign: "middle", margin: 0
    });

    const points = [
        { t: "PIN は覚えやすい 4 桁", d: "誕生日・電話番号下 4 桁など" },
        { t: "忘れたときは管理者へ", d: "リセットすれば再設定できます" },
        { t: "1 度ログインすれば OK", d: "次回から自動でログイン状態が継続" },
    ];
    points.forEach((p, i) => {
        const y = 2.25 + i * 0.95;
        s.addText("✓", {
            x: 5.5, y: y, w: 0.3, h: 0.3,
            fontSize: 16, fontFace: FONT_BODY, bold: true, color: C.accent2, margin: 0
        });
        s.addText(p.t, {
            x: 5.85, y: y, w: 3.55, h: 0.3,
            fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.textDark, margin: 0
        });
        s.addText(p.d, {
            x: 5.85, y: y + 0.3, w: 3.55, h: 0.35,
            fontSize: 10, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    addFooter(s, 2);
}

// ============================================================
// Slide 3: STEP 2 — 毎日の入力
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addStepHeader(s, 2, "その日の作業時間を入力する");

    // リード
    s.addText("ナビの「入力画面」タブから、毎日の作業を記録します", {
        x: 0.5, y: 1.55, w: 9, h: 0.35,
        fontSize: 13, fontFace: FONT_BODY, italic: true, color: C.textMid, margin: 0
    });

    // 必須項目リスト
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.0, w: 4.5, h: 3.15,
        fill: { color: C.white }, line: { color: "E2E8F0", width: 1 },
        shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.06 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.0, w: 0.12, h: 3.15,
        fill: { color: C.primary }, line: { color: C.primary }
    });
    s.addText("入力する項目", {
        x: 0.75, y: 2.1, w: 4.1, h: 0.35,
        fontSize: 13, fontFace: FONT_BODY, bold: true, color: C.primary, margin: 0
    });

    const items = [
        { label: "日付",       desc: "基本は本日（変更可）" },
        { label: "注番",       desc: "あり / なし を選択" },
        { label: "作業内容",   desc: "プルダウンから選択" },
        { label: "時間",       desc: "0.5時間単位で入力" },
        { label: "場所",       desc: "社内 / 社外" },
        { label: "詳細",       desc: "注番なしの時は必須" },
    ];
    items.forEach((it, i) => {
        const y = 2.5 + i * 0.42;
        s.addShape(pres.shapes.OVAL, {
            x: 0.85, y: y + 0.1, w: 0.16, h: 0.16,
            fill: { color: C.accent }, line: { color: C.accent }
        });
        s.addText(it.label, {
            x: 1.1, y: y, w: 1.3, h: 0.35,
            fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.textDark, margin: 0
        });
        s.addText(it.desc, {
            x: 2.3, y: y, w: 2.6, h: 0.35,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    // 右: コツ
    s.addShape(pres.shapes.RECTANGLE, {
        x: 5.3, y: 2.0, w: 4.2, h: 3.15,
        fill: { color: C.primary }, line: { color: C.primary },
        shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.12 }
    });
    s.addText("時短のコツ", {
        x: 5.5, y: 2.1, w: 4, h: 0.35,
        fontSize: 13, fontFace: FONT_BODY, bold: true, color: C.accent,
        charSpacing: 2, margin: 0
    });
    s.addText("30 秒で入力完了", {
        x: 5.5, y: 2.45, w: 4, h: 0.5,
        fontSize: 20, fontFace: FONT_HEADER, bold: true, color: C.white, margin: 0
    });
    s.addShape(pres.shapes.LINE, {
        x: 5.5, y: 3.05, w: 3.8, h: 0,
        line: { color: C.primary2, width: 1 }
    });
    const tips = [
        "前回の入力をワンクリックで再利用",
        "注番は部分一致で素早くサジェスト",
        "同じ注番の続きの作業なら 10 秒",
    ];
    tips.forEach((t, i) => {
        const y = 3.2 + i * 0.55;
        s.addText("→", {
            x: 5.5, y: y, w: 0.3, h: 0.3,
            fontSize: 14, fontFace: FONT_BODY, bold: true, color: C.accent, margin: 0
        });
        s.addText(t, {
            x: 5.85, y: y, w: 3.55, h: 0.4,
            fontSize: 12, fontFace: FONT_BODY, color: C.white, margin: 0
        });
    });

    addFooter(s, 3);
}

// ============================================================
// Slide 4: STEP 3 — 入力忘れチェック
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addStepHeader(s, 3, "入力忘れをチェックする");

    // リード
    s.addText("ホーム画面を開けば、直近 7 営業日の入力状況が一目で分かります", {
        x: 0.5, y: 1.55, w: 9, h: 0.35,
        fontSize: 13, fontFace: FONT_BODY, italic: true, color: C.textMid, margin: 0
    });

    // 疑似パネル
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.0, w: 9, h: 2.4,
        fill: { color: C.white }, line: { color: C.primary2, width: 1 },
        shadow: { type: "outer", blur: 8, offset: 3, angle: 90, color: "000000", opacity: 0.08 }
    });
    // ヘッダー帯
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.0, w: 9, h: 0.45,
        fill: { color: C.lightBg }, line: { color: C.lightBg }
    });
    s.addText("直近の入力状況（営業日ベース）", {
        x: 0.75, y: 2.0, w: 6, h: 0.45,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.primary,
        valign: "middle", margin: 0
    });
    s.addText("5 / 7 日 入力済み", {
        x: 7.5, y: 2.0, w: 1.85, h: 0.45,
        fontSize: 11, fontFace: FONT_BODY, bold: true, color: C.textMid,
        valign: "middle", align: "right", margin: 0
    });

    // チップ列
    const chips = [
        { d: "4/7",  dow: "火", state: "ok", label: "8h" },
        { d: "4/8",  dow: "水", state: "ok", label: "8h" },
        { d: "4/9",  dow: "木", state: "miss", label: "未入力" },
        { d: "4/10", dow: "金", state: "ok", label: "8h" },
        { d: "4/13", dow: "月", state: "miss", label: "未入力" },
        { d: "4/14", dow: "火", state: "ok", label: "8h" },
        { d: "4/15", dow: "水", state: "today", label: "本日" },
    ];
    const chipW = 1.0, chipH = 1.2, chipGap = 0.17;
    const totalChipW = chipW * 7 + chipGap * 6;
    const chipStartX = (10 - totalChipW) / 2;
    chips.forEach((c, i) => {
        const x = chipStartX + i * (chipW + chipGap);
        const y = 2.65;
        let fill, border, textColor, icon;
        if (c.state === "ok") {
            fill = "D1FAE5"; border = "10B981"; textColor = "065F46"; icon = "●";
        } else if (c.state === "miss") {
            fill = "FEF3C7"; border = "F59E0B"; textColor = "92400E"; icon = "⚠";
        } else {
            fill = "E5E7EB"; border = "94A3B8"; textColor = "334155"; icon = "○";
        }
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: y, w: chipW, h: chipH,
            fill: { color: fill }, line: { color: border, width: 1.5 }
        });
        s.addText(c.d, {
            x: x, y: y + 0.1, w: chipW, h: 0.25,
            fontSize: 11, fontFace: FONT_BODY, bold: true, color: textColor,
            align: "center", margin: 0
        });
        s.addText(`(${c.dow})`, {
            x: x, y: y + 0.32, w: chipW, h: 0.2,
            fontSize: 9, fontFace: FONT_BODY, color: textColor,
            align: "center", margin: 0
        });
        s.addText(icon, {
            x: x, y: y + 0.52, w: chipW, h: 0.3,
            fontSize: 16, fontFace: FONT_BODY, bold: true, color: border,
            align: "center", margin: 0
        });
        s.addText(c.label, {
            x: x, y: y + 0.85, w: chipW, h: 0.25,
            fontSize: 10, fontFace: FONT_BODY, bold: true, color: textColor,
            align: "center", margin: 0
        });
    });

    // 凡例
    const legends = [
        { color: "10B981", icon: "●", text: "入力済み" },
        { color: "F59E0B", icon: "⚠", text: "未入力（警告）" },
        { color: "94A3B8", icon: "○", text: "今日まだ" },
    ];
    legends.forEach((lg, i) => {
        const x = 1.0 + i * 3.0;
        const y = 4.7;
        s.addText(lg.icon, {
            x: x, y: y, w: 0.3, h: 0.3,
            fontSize: 14, fontFace: FONT_BODY, bold: true, color: lg.color, margin: 0
        });
        s.addText(lg.text, {
            x: x + 0.3, y: y, w: 2.5, h: 0.3,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    addFooter(s, 4);
}

// ============================================================
// Slide 5: STEP 4 — 過去の記録を見る
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addStepHeader(s, 4, "過去の記録を見る / 修正する");

    // リード
    s.addText("ナビの「データ解析画面」から、過去の記録が全て確認できます", {
        x: 0.5, y: 1.55, w: 9, h: 0.35,
        fontSize: 13, fontFace: FONT_BODY, italic: true, color: C.textMid, margin: 0
    });

    // 2列: データ一覧 / 集計画面
    const cards = [
        {
            title: "データ一覧",
            subtitle: "記録を一覧で見る",
            color: C.primary,
            items: [
                "全員の記録をリスト表示",
                "日付・担当者・注番で絞り込み",
                "自分の記録は編集・削除も可能",
                "過去 +1/+3/+6/+12ヶ月 に範囲拡張",
            ]
        },
        {
            title: "集計画面",
            subtitle: "期別・注番別に集計",
            color: C.primary2,
            items: [
                "期（半期・通年）で切替",
                "注番別・担当者別に時間合計",
                "前年同期との比較表示",
                "CSV ダウンロードで Excel 活用",
            ]
        },
    ];
    const cardW = 4.4, cardH = 3.2, cardGap = 0.2;
    const cardStartX = (10 - (cardW * 2 + cardGap)) / 2;
    cards.forEach((c, i) => {
        const x = cardStartX + i * (cardW + cardGap);
        const y = 2.0;
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: y, w: cardW, h: cardH,
            fill: { color: C.white }, line: { color: "E2E8F0", width: 1 },
            shadow: { type: "outer", blur: 8, offset: 2, angle: 90, color: "000000", opacity: 0.08 }
        });
        // Top color bar
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: y, w: cardW, h: 0.15,
            fill: { color: c.color }, line: { color: c.color }
        });
        s.addText(c.title, {
            x: x + 0.3, y: y + 0.3, w: cardW - 0.6, h: 0.5,
            fontSize: 22, fontFace: FONT_HEADER, bold: true, color: c.color, margin: 0
        });
        s.addText(c.subtitle, {
            x: x + 0.3, y: y + 0.8, w: cardW - 0.6, h: 0.35,
            fontSize: 12, fontFace: FONT_BODY, italic: true, color: C.textMid, margin: 0
        });
        // Divider
        s.addShape(pres.shapes.LINE, {
            x: x + 0.3, y: y + 1.25, w: cardW - 0.6, h: 0,
            line: { color: "E2E8F0", width: 1 }
        });
        // Items
        c.items.forEach((it, j) => {
            const iy = y + 1.4 + j * 0.42;
            s.addText("→", {
                x: x + 0.3, y: iy, w: 0.25, h: 0.3,
                fontSize: 12, fontFace: FONT_BODY, bold: true, color: c.color, margin: 0
            });
            s.addText(it, {
                x: x + 0.6, y: iy, w: cardW - 0.85, h: 0.35,
                fontSize: 11, fontFace: FONT_BODY, color: C.textDark, margin: 0
            });
        });
    });

    addFooter(s, 5);
}

// ============================================================
// Slide 6: ワンポイント — 便利な使い方
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    // ヘッダー
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 0, w: 10, h: 1.3, fill: { color: C.primary }, line: { color: C.primary }
    });
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 0.3, w: 1.4, h: 0.7,
        fill: { color: C.accent }, line: { color: C.accent }
    });
    s.addText("TIPS", {
        x: 0.5, y: 0.3, w: 1.4, h: 0.7,
        fontSize: 18, fontFace: FONT_HEADER, bold: true, color: C.white,
        align: "center", valign: "middle", margin: 0
    });
    s.addText("便利な使い方", {
        x: 2.1, y: 0.3, w: 7.5, h: 0.7,
        fontSize: 24, fontFace: FONT_HEADER, bold: true, color: C.white,
        valign: "middle", margin: 0
    });

    // 4 tips grid
    const tips = [
        {
            title: "前日の続きは 10 秒",
            desc: "入力画面の「前回を再利用」ボタンで、直前の記録がコピーされます。同じ注番の続きなら時間だけ変更して保存。",
            color: C.primary,
        },
        {
            title: "注番は部分一致で検索",
            desc: "番号の一部だけ入力すると候補が出ます。頭の数字を覚えていなくても大丈夫。",
            color: C.primary2,
        },
        {
            title: "目標時間は自分で設定",
            desc: "初回のみ各注番に目標時間を入れられます。達成率が自動計算され、進捗が一目で分かります。",
            color: C.accent,
        },
        {
            title: "忘れた日もクリック 1 つで",
            desc: "ホーム画面の未入力チップをクリックすると、その日付が入力画面に自動で入ります。",
            color: C.accent2,
        },
    ];
    const tw = 4.4, th = 1.75, tgap = 0.2;
    const tsx = (10 - (tw * 2 + tgap)) / 2;
    const tsy = 1.6;
    tips.forEach((t, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const x = tsx + col * (tw + tgap);
        const y = tsy + row * (th + tgap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: y, w: tw, h: th,
            fill: { color: C.white }, line: { color: "E2E8F0", width: 1 },
            shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.06 }
        });
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: y, w: 0.1, h: th,
            fill: { color: t.color }, line: { color: t.color }
        });
        s.addShape(pres.shapes.OVAL, {
            x: x + 0.3, y: y + 0.2, w: 0.5, h: 0.5,
            fill: { color: t.color }, line: { color: t.color }
        });
        s.addText(String(i + 1), {
            x: x + 0.3, y: y + 0.2, w: 0.5, h: 0.5,
            fontSize: 16, fontFace: FONT_HEADER, bold: true, color: C.white,
            align: "center", valign: "middle", margin: 0
        });
        s.addText(t.title, {
            x: x + 0.9, y: y + 0.25, w: tw - 1.05, h: 0.4,
            fontSize: 14, fontFace: FONT_BODY, bold: true, color: C.textDark, margin: 0
        });
        s.addText(t.desc, {
            x: x + 0.3, y: y + 0.8, w: tw - 0.5, h: 0.85,
            fontSize: 10, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    addFooter(s, 6);
}

// ============================================================
// Slide 7: 毎日の基本フロー
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    // ヘッダー
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 0, w: 10, h: 1.3, fill: { color: C.primary }, line: { color: C.primary }
    });
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 0.3, w: 1.7, h: 0.7,
        fill: { color: C.accent }, line: { color: C.accent }
    });
    s.addText("DAILY", {
        x: 0.5, y: 0.3, w: 1.7, h: 0.7,
        fontSize: 18, fontFace: FONT_HEADER, bold: true, color: C.white,
        align: "center", valign: "middle", margin: 0
    });
    s.addText("毎日の基本フロー", {
        x: 2.4, y: 0.3, w: 7.2, h: 0.7,
        fontSize: 24, fontFace: FONT_HEADER, bold: true, color: C.white,
        valign: "middle", margin: 0
    });

    // 横フロー
    const steps = [
        { time: "朝",   what: "ホーム画面を開く",     desc: "前日の入力確認" },
        { time: "昼",   what: "午前分を入力",         desc: "記憶が新しいうちに" },
        { time: "夕",   what: "午後分を入力",         desc: "1 日分まとめて" },
        { time: "退勤", what: "忘れてないか確認",     desc: "明日の朝もチェック" },
    ];
    const sw = 2.15, sh = 2.3, sgap = 0.15;
    const ssx = (10 - (sw * 4 + sgap * 3)) / 2;
    const ssy = 1.7;
    steps.forEach((st, i) => {
        const x = ssx + i * (sw + sgap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: ssy, w: sw, h: sh,
            fill: { color: C.white }, line: { color: C.primary2, width: 1.5 },
            shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.08 }
        });
        // Time badge
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: ssy, w: sw, h: 0.5,
            fill: { color: C.primary }, line: { color: C.primary }
        });
        s.addText(st.time, {
            x: x, y: ssy, w: sw, h: 0.5,
            fontSize: 16, fontFace: FONT_HEADER, bold: true, color: C.white,
            align: "center", valign: "middle", margin: 0
        });
        // What
        s.addText(st.what, {
            x: x + 0.15, y: ssy + 0.7, w: sw - 0.3, h: 0.5,
            fontSize: 14, fontFace: FONT_BODY, bold: true, color: C.textDark,
            align: "center", margin: 0
        });
        // Divider
        s.addShape(pres.shapes.LINE, {
            x: x + sw/2 - 0.3, y: ssy + 1.3, w: 0.6, h: 0,
            line: { color: C.primary2, width: 1 }
        });
        // Desc
        s.addText(st.desc, {
            x: x + 0.15, y: ssy + 1.4, w: sw - 0.3, h: 0.8,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid,
            align: "center", margin: 0
        });
    });

    // 下部メッセージ
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 4.3, w: 9, h: 0.8,
        fill: { color: C.primary3 }, line: { color: C.primary3 }
    });
    s.addText("覚えるのは「朝チェック・昼夕入力」だけ。習慣にすれば負担ゼロです。", {
        x: 0.5, y: 4.3, w: 9, h: 0.8,
        fontSize: 14, fontFace: FONT_BODY, italic: true, color: C.dark,
        align: "center", valign: "middle", margin: 0
    });

    addFooter(s, 7);
}

// ============================================================
// Slide 8: 困ったときは / 最後に
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.dark };

    // 装飾
    s.addShape(pres.shapes.OVAL, {
        x: 6, y: -2, w: 6, h: 6,
        fill: { color: C.primary, transparency: 75 },
        line: { color: C.primary, transparency: 75 }
    });

    // ラベル
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 0.55, w: 0.15, h: 0.35, fill: { color: C.accent }, line: { color: C.accent }
    });
    s.addText("困ったときは", {
        x: 0.75, y: 0.55, w: 9, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.accent,
        valign: "middle", charSpacing: 2, margin: 0
    });

    // タイトル
    s.addText("遠慮なく、聞いてください。", {
        x: 0.5, y: 1.2, w: 9, h: 0.8,
        fontSize: 34, fontFace: FONT_HEADER, bold: true, color: C.white, margin: 0
    });

    // 3 つのサポート
    const supports = [
        { title: "操作が分からない", desc: "完全マニュアルを参照",      color: C.primary2 },
        { title: "エラーが出た",     desc: "IT 担当者に連絡",           color: C.accent },
        { title: "PIN を忘れた",     desc: "管理者にリセット依頼",      color: C.accent2 },
    ];
    const supW = 2.9, supH = 1.35, supGap = 0.15;
    const supSx = (10 - (supW * 3 + supGap * 2)) / 2;
    const supSy = 2.3;
    supports.forEach((sp, i) => {
        const x = supSx + i * (supW + supGap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: supSy, w: supW, h: supH,
            fill: { color: C.white }, line: { color: sp.color, width: 2 }
        });
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: supSy, w: supW, h: 0.12,
            fill: { color: sp.color }, line: { color: sp.color }
        });
        s.addText(sp.title, {
            x: x + 0.15, y: supSy + 0.3, w: supW - 0.3, h: 0.4,
            fontSize: 14, fontFace: FONT_BODY, bold: true, color: C.dark,
            align: "center", margin: 0
        });
        s.addShape(pres.shapes.LINE, {
            x: x + supW/2 - 0.25, y: supSy + 0.8, w: 0.5, h: 0,
            line: { color: sp.color, width: 1 }
        });
        s.addText(sp.desc, {
            x: x + 0.15, y: supSy + 0.9, w: supW - 0.3, h: 0.4,
            fontSize: 12, fontFace: FONT_BODY, color: C.textMid,
            align: "center", margin: 0
        });
    });

    // 最後のメッセージ
    s.addShape(pres.shapes.LINE, {
        x: 0.5, y: 4.15, w: 0.6, h: 0,
        line: { color: C.accent, width: 2 }
    });
    s.addText("使ってみて、気づいたことがあれば教えてください。", {
        x: 0.5, y: 4.3, w: 9, h: 0.5,
        fontSize: 15, fontFace: FONT_BODY, italic: true, color: C.primary3, margin: 0
    });
    s.addText("みんなで育てていく、そういうシステムです。", {
        x: 0.5, y: 4.75, w: 9, h: 0.4,
        fontSize: 13, fontFace: FONT_BODY, color: C.primary3, margin: 0
    });

    // 下部バー
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 5.4, w: 10, h: 0.225, fill: { color: C.primary }, line: { color: C.primary }
    });
    s.addText("簡易操作手順書  —  作業時間 一元管理システム", {
        x: 0.4, y: 5.4, w: 6, h: 0.225,
        fontSize: 9, fontFace: FONT_BODY, color: C.white, valign: "middle", margin: 0
    });
    s.addText(`${TOTAL} / ${TOTAL}`, {
        x: 9.0, y: 5.4, w: 0.8, h: 0.225,
        fontSize: 9, fontFace: FONT_BODY, color: C.white, valign: "middle", align: "right", margin: 0
    });
}

// ============================================================
// 書き出し
// ============================================================
pres.writeFile({ fileName: "02_簡易操作手順書.pptx" })
    .then(fileName => console.log(`✓ ${fileName} を作成しました`))
    .catch(err => { console.error(err); process.exit(1); });
