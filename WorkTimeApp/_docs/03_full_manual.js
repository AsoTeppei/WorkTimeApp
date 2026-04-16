// ============================================================
// 03_full_manual.js — 完全マニュアル
//
// 目的: 全機能を網羅した詳細マニュアル。
// ============================================================

const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "作業時間一元管理システム";
pres.title  = "作業時間 一元管理システム 完全マニュアル";

// ---- Midnight Executive パレット ----
const C = {
    primary:   "1E2761", // navy
    primary2:  "2D3E7C",
    primary3:  "5468A8", // lighter navy
    light:     "F8FAFC",
    lightBg:   "EEF2F9",
    border:    "E2E8F0",
    accent:    "D4AF37", // gold
    accent2:   "4FC3F7", // light blue
    success:   "10B981",
    warning:   "F59E0B",
    danger:    "EF4444",
    textDark:  "1E293B",
    textMid:   "475569",
    textLight: "94A3B8",
    white:     "FFFFFF",
};

const FONT_HEADER = "Georgia";
const FONT_BODY   = "Calibri";

const TOTAL = 18;

// ---- 共通ヘルパー ----
function addFooter(slide, pageNum) {
    slide.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 5.4, w: 10, h: 0.225, fill: { color: C.primary }, line: { color: C.primary }
    });
    slide.addText("完全マニュアル  —  作業時間 一元管理システム", {
        x: 0.4, y: 5.4, w: 6, h: 0.225,
        fontSize: 9, fontFace: FONT_BODY, color: C.white, valign: "middle", margin: 0
    });
    slide.addText(`${pageNum} / ${TOTAL}`, {
        x: 9.0, y: 5.4, w: 0.8, h: 0.225,
        fontSize: 9, fontFace: FONT_BODY, color: C.white, valign: "middle", align: "right", margin: 0
    });
}

function addHeader(slide, chapter, title) {
    // 上部薄い帯
    slide.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 0, w: 10, h: 0.1, fill: { color: C.accent }, line: { color: C.accent }
    });
    // 章ラベル
    slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 0.35, w: 0.15, h: 0.35, fill: { color: C.accent }, line: { color: C.accent }
    });
    slide.addText(chapter, {
        x: 0.75, y: 0.35, w: 9, h: 0.35,
        fontSize: 11, fontFace: FONT_BODY, bold: true, color: C.primary3,
        valign: "middle", charSpacing: 2, margin: 0
    });
    // タイトル
    slide.addText(title, {
        x: 0.5, y: 0.75, w: 9, h: 0.65,
        fontSize: 28, fontFace: FONT_HEADER, bold: true, color: C.primary,
        valign: "middle", margin: 0
    });
    // 下線
    slide.addShape(pres.shapes.LINE, {
        x: 0.5, y: 1.45, w: 1.2, h: 0,
        line: { color: C.accent, width: 2 }
    });
}

function addInfoBox(slide, x, y, w, h, title, items, color) {
    slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w, h,
        fill: { color: C.white }, line: { color: C.border, width: 1 },
        shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.06 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w, h: 0.4, fill: { color: color }, line: { color: color }
    });
    slide.addText(title, {
        x: x + 0.2, y: y, w: w - 0.4, h: 0.4,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.white,
        valign: "middle", margin: 0
    });
    items.forEach((it, i) => {
        const iy = y + 0.55 + i * 0.45;
        slide.addShape(pres.shapes.OVAL, {
            x: x + 0.25, y: iy + 0.12, w: 0.15, h: 0.15,
            fill: { color: color }, line: { color: color }
        });
        if (typeof it === "string") {
            slide.addText(it, {
                x: x + 0.5, y: iy, w: w - 0.7, h: 0.4,
                fontSize: 11, fontFace: FONT_BODY, color: C.textDark, margin: 0
            });
        } else {
            slide.addText(it.label, {
                x: x + 0.5, y: iy, w: w - 0.7, h: 0.3,
                fontSize: 11, fontFace: FONT_BODY, bold: true, color: C.textDark, margin: 0
            });
            slide.addText(it.desc, {
                x: x + 0.5, y: iy + 0.22, w: w - 0.7, h: 0.3,
                fontSize: 10, fontFace: FONT_BODY, color: C.textMid, margin: 0
            });
        }
    });
}

// ============================================================
// Slide 1: タイトル
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.primary };

    // 装飾
    s.addShape(pres.shapes.OVAL, {
        x: -3, y: 2, w: 7, h: 7,
        fill: { color: C.primary2, transparency: 70 },
        line: { color: C.primary2, transparency: 70 }
    });
    s.addShape(pres.shapes.OVAL, {
        x: 7, y: -2, w: 5, h: 5,
        fill: { color: C.primary3, transparency: 85 },
        line: { color: C.primary3, transparency: 85 }
    });

    // 上部アクセント
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 0, w: 10, h: 0.15, fill: { color: C.accent }, line: { color: C.accent }
    });

    // ラベル
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 0.6, w: 0.15, h: 0.4, fill: { color: C.accent }, line: { color: C.accent }
    });
    s.addText("COMPLETE MANUAL", {
        x: 0.75, y: 0.6, w: 9, h: 0.4,
        fontSize: 13, fontFace: FONT_HEADER, bold: true, color: C.accent,
        valign: "middle", charSpacing: 3, margin: 0
    });

    // タイトル
    s.addText("完全マニュアル", {
        x: 0.5, y: 1.5, w: 9, h: 1.0,
        fontSize: 54, fontFace: FONT_HEADER, bold: true, color: C.white, margin: 0
    });
    s.addText("作業時間 一元管理システム", {
        x: 0.5, y: 2.6, w: 9, h: 0.6,
        fontSize: 22, fontFace: FONT_BODY, color: C.accent2, margin: 0
    });

    // 下部情報
    s.addShape(pres.shapes.LINE, {
        x: 0.5, y: 3.7, w: 1, h: 0,
        line: { color: C.accent, width: 2 }
    });
    s.addText("全機能リファレンス", {
        x: 0.5, y: 3.85, w: 9, h: 0.4,
        fontSize: 16, fontFace: FONT_BODY, italic: true, color: C.primary3, margin: 0
    });
    s.addText("ログインから管理者機能まで、全ての操作を網羅", {
        x: 0.5, y: 4.25, w: 9, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, color: C.primary3, margin: 0
    });

    // 下部バー
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 5.2, w: 10, h: 0.425, fill: { color: C.accent }, line: { color: C.accent }
    });
    s.addText("Version 1.0  —  2026年4月", {
        x: 0.5, y: 5.2, w: 9, h: 0.425,
        fontSize: 11, fontFace: FONT_BODY, bold: true, color: C.primary,
        valign: "middle", margin: 0
    });
}

// ============================================================
// Slide 2: 目次
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "CONTENTS", "目次");

    const toc = [
        { num: "1", title: "画面構成の全体像",       page: "3" },
        { num: "2", title: "ログイン（初回 / 2 回目以降）", page: "4" },
        { num: "3", title: "ホーム画面の見方",       page: "5" },
        { num: "4", title: "入力画面 — 基本操作",    page: "6" },
        { num: "5", title: "入力画面 — 注番ありの作業", page: "7" },
        { num: "6", title: "入力画面 — 注番なしの作業", page: "8" },
        { num: "7", title: "目標時間の設定",         page: "9" },
        { num: "8", title: "データ解析 — データ一覧", page: "10" },
        { num: "9", title: "データ解析 — 集計画面",   page: "11" },
        { num: "10", title: "データ解析 — 会計年度集計", page: "12" },
        { num: "11", title: "データ解析 — 前年比較",  page: "13" },
        { num: "12", title: "データ解析 — CSV 出力",  page: "14" },
        { num: "13", title: "環境設定 — ユーザー管理（管理者）", page: "15" },
        { num: "14", title: "環境設定 — グループ・注番管理", page: "16" },
        { num: "15", title: "トラブルシューティング", page: "17" },
        { num: "16", title: "Q&A / お問い合わせ",     page: "18" },
    ];

    // 2列配置
    const col1 = toc.slice(0, 8);
    const col2 = toc.slice(8);
    const lineH = 0.41;
    const startY = 1.75;

    [col1, col2].forEach((col, colIdx) => {
        const colX = 0.5 + colIdx * 4.75;
        col.forEach((it, i) => {
            const y = startY + i * lineH;
            s.addShape(pres.shapes.RECTANGLE, {
                x: colX, y: y + 0.02, w: 0.45, h: 0.35,
                fill: { color: C.primary }, line: { color: C.primary }
            });
            s.addText(it.num, {
                x: colX, y: y + 0.02, w: 0.45, h: 0.35,
                fontSize: 11, fontFace: FONT_HEADER, bold: true, color: C.white,
                align: "center", valign: "middle", margin: 0
            });
            s.addText(it.title, {
                x: colX + 0.55, y: y + 0.02, w: 3.6, h: 0.35,
                fontSize: 11, fontFace: FONT_BODY, color: C.textDark,
                valign: "middle", margin: 0
            });
            s.addText(`p.${it.page}`, {
                x: colX + 4.15, y: y + 0.02, w: 0.55, h: 0.35,
                fontSize: 10, fontFace: FONT_BODY, color: C.textLight,
                valign: "middle", align: "right", margin: 0
            });
        });
    });

    addFooter(s, 2);
}

// ============================================================
// Slide 3: 画面構成の全体像
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 1 章", "画面構成の全体像");

    s.addText("システムは 4 つの画面で構成されています。ナビゲーションタブから切り替えます。", {
        x: 0.5, y: 1.6, w: 9, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, color: C.textMid, margin: 0
    });

    const screens = [
        { no: "01", title: "ホーム画面",       desc: "直近の入力状況、今月の統計、メニューへのリンク", who: "全員" },
        { no: "02", title: "入力画面",         desc: "作業時間の記録・保存",                         who: "全員" },
        { no: "03", title: "データ解析画面",   desc: "過去データの一覧・集計・CSV 出力",             who: "全員" },
        { no: "04", title: "環境設定画面",     desc: "ユーザー・グループ・注番の管理",               who: "管理者のみ" },
    ];
    const sw = 4.4, sh = 1.5, sgap = 0.2;
    const ssx = 0.5;
    const ssy = 2.1;
    screens.forEach((sc, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const x = ssx + col * (sw + sgap);
        const y = ssy + row * (sh + sgap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: y, w: sw, h: sh,
            fill: { color: C.white }, line: { color: C.border, width: 1 },
            shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.06 }
        });
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: y, w: 0.12, h: sh,
            fill: { color: C.primary }, line: { color: C.primary }
        });
        s.addText(sc.no, {
            x: x + 0.25, y: y + 0.15, w: 0.8, h: 0.4,
            fontSize: 20, fontFace: FONT_HEADER, bold: true, color: C.accent, margin: 0
        });
        s.addText(sc.title, {
            x: x + 1.05, y: y + 0.2, w: sw - 1.3, h: 0.4,
            fontSize: 17, fontFace: FONT_BODY, bold: true, color: C.primary, margin: 0
        });
        s.addText(sc.desc, {
            x: x + 0.25, y: y + 0.75, w: sw - 0.5, h: 0.4,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
        // 権限バッジ
        const isAdminOnly = sc.who === "管理者のみ";
        s.addShape(pres.shapes.RECTANGLE, {
            x: x + 0.25, y: y + 1.1, w: 1.2, h: 0.3,
            fill: { color: isAdminOnly ? C.warning : C.success },
            line: { color: isAdminOnly ? C.warning : C.success }
        });
        s.addText(sc.who, {
            x: x + 0.25, y: y + 1.1, w: 1.2, h: 0.3,
            fontSize: 9, fontFace: FONT_BODY, bold: true, color: C.white,
            align: "center", valign: "middle", margin: 0
        });
    });

    addFooter(s, 3);
}

// ============================================================
// Slide 4: ログイン
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 2 章", "ログイン");

    // 左: 初回ログイン
    addInfoBox(s, 0.5, 1.8, 4.5, 3.3, "初回ログイン", [
        { label: "1. URL を開く", desc: "http://192.168.1.8:3000/" },
        { label: "2. 一覧から自分を選択", desc: "名前・部署で検索可" },
        { label: "3. PIN を自分で設定", desc: "4 桁の半角数字" },
        { label: "4. 再入力で確認", desc: "同じ PIN を 2 回入力" },
        { label: "5. ログイン完了", desc: "ホーム画面に遷移" },
    ], C.primary);

    // 右: 2回目以降
    addInfoBox(s, 5.1, 1.8, 4.4, 3.3, "2 回目以降", [
        { label: "自動ログイン", desc: "1 度ログインすれば 10 年維持" },
        { label: "PIN 再入力は不要", desc: "ブラウザを閉じても保持" },
        { label: "別 PC では再ログイン", desc: "端末ごとに初回操作必要" },
        { label: "PIN を忘れたら", desc: "管理者にリセット依頼" },
        { label: "強制ログアウト", desc: "PIN 変更時に他端末から離脱" },
    ], C.primary2);

    addFooter(s, 4);
}

// ============================================================
// Slide 5: ホーム画面の見方
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 3 章", "ホーム画面の見方");

    const sections = [
        { title: "会計年度バッジ",       desc: "現在の期（第○期 前期/後期）が画面上部に表示" },
        { title: "入力忘れパネル",       desc: "直近 7 営業日の入力状況をチップで表示。未入力日があれば警告" },
        { title: "今月のサマリー",       desc: "登録ユーザー数、進行中注番数、今月の総時間など" },
        { title: "部署別 月次集計",     desc: "今月の時間を部署ごとに棒グラフで比較" },
        { title: "注番ランキング",       desc: "今月の時間が多い注番 TOP 5 を表示" },
        { title: "担当者ランキング",     desc: "今月の時間が多い担当者 TOP 5 を表示" },
    ];
    const sw = 4.4, sh = 1.0, sgap = 0.15;
    const ssx = 0.5;
    const ssy = 1.85;
    sections.forEach((sc, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const x = ssx + col * (sw + sgap);
        const y = ssy + row * (sh + sgap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: y, w: sw, h: sh,
            fill: { color: C.white }, line: { color: C.border, width: 1 }
        });
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: y, w: 0.1, h: sh,
            fill: { color: C.accent }, line: { color: C.accent }
        });
        s.addText(sc.title, {
            x: x + 0.25, y: y + 0.15, w: sw - 0.4, h: 0.35,
            fontSize: 13, fontFace: FONT_BODY, bold: true, color: C.primary, margin: 0
        });
        s.addText(sc.desc, {
            x: x + 0.25, y: y + 0.5, w: sw - 0.4, h: 0.45,
            fontSize: 10, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    // 下部ヒント
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 5.0, w: 9, h: 0.3,
        fill: { color: C.lightBg }, line: { color: C.lightBg }
    });
    s.addText("💡 入力忘れパネルの未入力チップをクリックすると、その日付が自動で入力画面に入ります。", {
        x: 0.5, y: 5.0, w: 9, h: 0.3,
        fontSize: 10, fontFace: FONT_BODY, italic: true, color: C.primary,
        align: "center", valign: "middle", margin: 0
    });

    addFooter(s, 5);
}

// ============================================================
// Slide 6: 入力画面 — 基本操作
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 4 章", "入力画面 — 基本操作");

    // 入力項目テーブル
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 1.8, w: 9, h: 3.4,
        fill: { color: C.white }, line: { color: C.border, width: 1 },
        shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.06 }
    });
    // テーブルヘッダー
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 1.8, w: 9, h: 0.45,
        fill: { color: C.primary }, line: { color: C.primary }
    });
    s.addText("項目", {
        x: 0.7, y: 1.8, w: 1.6, h: 0.45,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.white,
        valign: "middle", margin: 0
    });
    s.addText("必須", {
        x: 2.3, y: 1.8, w: 0.8, h: 0.45,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.white,
        align: "center", valign: "middle", margin: 0
    });
    s.addText("説明", {
        x: 3.15, y: 1.8, w: 6.3, h: 0.45,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.white,
        valign: "middle", margin: 0
    });

    const rows = [
        { item: "日付",       req: "○", desc: "デフォルト本日。カレンダーで変更可能。" },
        { item: "担当者",     req: "固定", desc: "ログインユーザー固定（変更不可）" },
        { item: "注番",       req: "条件", desc: "「注番あり」モードのみ必須。部分一致で検索" },
        { item: "作業内容",   req: "○", desc: "部署ごとのプルダウンから選択" },
        { item: "時間",       req: "○", desc: "0.5 時間単位で入力（0.5 〜 24 時間）" },
        { item: "場所",       req: "○", desc: "社内 / 社外 を選択" },
        { item: "出荷後対応", req: "—", desc: "チェックボックス。必要時のみ" },
        { item: "詳細",       req: "条件", desc: "注番なし作業の場合のみ必須" },
    ];
    rows.forEach((r, i) => {
        const y = 2.3 + i * 0.36;
        // 交互色
        if (i % 2 === 1) {
            s.addShape(pres.shapes.RECTANGLE, {
                x: 0.5, y: y, w: 9, h: 0.36,
                fill: { color: C.lightBg }, line: { color: C.lightBg }
            });
        }
        s.addText(r.item, {
            x: 0.7, y: y, w: 1.6, h: 0.36,
            fontSize: 11, fontFace: FONT_BODY, bold: true, color: C.textDark,
            valign: "middle", margin: 0
        });
        const reqColor = r.req === "○" ? C.danger : (r.req === "条件" ? C.warning : C.textLight);
        s.addText(r.req, {
            x: 2.3, y: y, w: 0.8, h: 0.36,
            fontSize: 11, fontFace: FONT_BODY, bold: true, color: reqColor,
            align: "center", valign: "middle", margin: 0
        });
        s.addText(r.desc, {
            x: 3.15, y: y, w: 6.3, h: 0.36,
            fontSize: 10, fontFace: FONT_BODY, color: C.textMid,
            valign: "middle", margin: 0
        });
    });

    addFooter(s, 6);
}

// ============================================================
// Slide 7: 入力画面 — 注番ありの作業
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 5 章", "入力画面 — 注番ありの作業");

    // 手順
    const steps = [
        { n: "1", t: "「注番あり」を選択", d: "トグルで切り替え" },
        { n: "2", t: "注番を入力 or 選択", d: "部分一致で候補表示" },
        { n: "3", t: "新規注番なら登録", d: "得意先名・件名を入力" },
        { n: "4", t: "作業内容・時間を記入", d: "プルダウンから選択" },
        { n: "5", t: "登録ボタンをクリック", d: "完了通知が表示" },
    ];

    // 左: 手順
    steps.forEach((st, i) => {
        const y = 1.85 + i * 0.6;
        s.addShape(pres.shapes.OVAL, {
            x: 0.5, y: y, w: 0.5, h: 0.5,
            fill: { color: C.primary }, line: { color: C.primary }
        });
        s.addText(st.n, {
            x: 0.5, y: y, w: 0.5, h: 0.5,
            fontSize: 18, fontFace: FONT_HEADER, bold: true, color: C.white,
            align: "center", valign: "middle", margin: 0
        });
        s.addText(st.t, {
            x: 1.15, y: y + 0.02, w: 4, h: 0.3,
            fontSize: 13, fontFace: FONT_BODY, bold: true, color: C.textDark, margin: 0
        });
        s.addText(st.d, {
            x: 1.15, y: y + 0.28, w: 4, h: 0.25,
            fontSize: 10, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    // 右: 注意事項
    addInfoBox(s, 5.3, 1.85, 4.2, 3.2, "注意事項", [
        { label: "クローズ済み注番", desc: "入力不可（管理者が解除)" },
        { label: "2ヶ月入力なし", desc: "自動でクローズされる" },
        { label: "注番のコピー入力", desc: "前回の再利用で時短" },
        { label: "注番サジェスト順", desc: "自分の最近の注番が上位に" },
    ], C.warning);

    addFooter(s, 7);
}

// ============================================================
// Slide 8: 入力画面 — 注番なしの作業
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 6 章", "入力画面 — 注番なしの作業");

    // リード
    s.addText("会議・事務作業・社内業務など、特定の注番に紐付かない作業を記録します。", {
        x: 0.5, y: 1.65, w: 9, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, color: C.textMid, margin: 0
    });

    // 該当例
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.1, w: 4.5, h: 3.0,
        fill: { color: C.white }, line: { color: C.border, width: 1 },
        shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.06 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.1, w: 4.5, h: 0.4,
        fill: { color: C.primary2 }, line: { color: C.primary2 }
    });
    s.addText("該当例", {
        x: 0.7, y: 2.1, w: 4.3, h: 0.4,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.white,
        valign: "middle", margin: 0
    });
    const examples = [
        "朝礼・定例会議",
        "社内勉強会・教育",
        "資料整理・事務処理",
        "5S 活動・清掃",
        "健康診断・面談",
        "出張移動時間",
    ];
    examples.forEach((ex, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const x = 0.75 + col * 2.1;
        const y = 2.7 + row * 0.75;
        s.addShape(pres.shapes.OVAL, {
            x: x, y: y + 0.08, w: 0.18, h: 0.18,
            fill: { color: C.primary2 }, line: { color: C.primary2 }
        });
        s.addText(ex, {
            x: x + 0.25, y: y, w: 1.8, h: 0.4,
            fontSize: 11, fontFace: FONT_BODY, color: C.textDark, margin: 0
        });
    });

    // 右: 入力ルール
    addInfoBox(s, 5.3, 2.1, 4.2, 3.0, "入力ルール", [
        { label: "詳細欄は必須", desc: "具体的な作業内容を記載" },
        { label: "作業内容はプルダウン", desc: "部署マスタから選択" },
        { label: "時間は 0.5 単位", desc: "最小 0.5h、最大 24h" },
        { label: "注番フィールドは空", desc: "入力不要" },
    ], C.accent2);

    addFooter(s, 8);
}

// ============================================================
// Slide 9: 目標時間の設定
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 7 章", "目標時間の設定");

    // リード
    s.addText("各注番に対して、自分が何時間で完了させたいかの目標を設定できます。", {
        x: 0.5, y: 1.65, w: 9, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, color: C.textMid, margin: 0
    });

    // 3カラム
    const cols = [
        {
            title: "初回設定",
            color: C.primary,
            items: [
                "その注番で初めて入力するとき",
                "入力画面で目標時間の入力欄が表示",
                "スキップも可能（後から設定できる)",
            ]
        },
        {
            title: "変更方法",
            color: C.accent,
            items: [
                "集計画面の注番別サマリーから変更",
                "変更時は理由の記録が必要",
                "変更履歴は監査ログに残る",
            ]
        },
        {
            title: "活用方法",
            color: C.success,
            items: [
                "達成率が自動計算される",
                "超過した場合は警告表示",
                "次回の見積もり精度が上がる",
            ]
        },
    ];
    const cw = 2.95, ch = 3.0, cgap = 0.15;
    const csx = (10 - (cw * 3 + cgap * 2)) / 2;
    const csy = 2.0;
    cols.forEach((c, i) => {
        const x = csx + i * (cw + cgap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: csy, w: cw, h: ch,
            fill: { color: C.white }, line: { color: C.border, width: 1 },
            shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.06 }
        });
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: csy, w: cw, h: 0.45,
            fill: { color: c.color }, line: { color: c.color }
        });
        s.addText(c.title, {
            x: x, y: csy, w: cw, h: 0.45,
            fontSize: 14, fontFace: FONT_BODY, bold: true, color: C.white,
            align: "center", valign: "middle", margin: 0
        });
        c.items.forEach((it, j) => {
            const iy = csy + 0.7 + j * 0.7;
            s.addShape(pres.shapes.OVAL, {
                x: x + 0.2, y: iy + 0.1, w: 0.18, h: 0.18,
                fill: { color: c.color }, line: { color: c.color }
            });
            s.addText(it, {
                x: x + 0.45, y: iy, w: cw - 0.6, h: 0.6,
                fontSize: 11, fontFace: FONT_BODY, color: C.textDark, margin: 0
            });
        });
    });

    addFooter(s, 9);
}

// ============================================================
// Slide 10: データ解析 — データ一覧
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 8 章", "データ解析 — データ一覧画面");

    const features = [
        { title: "記録の一覧表示",   desc: "全員の作業記録を日付降順で表示" },
        { title: "絞り込み",         desc: "日付・担当者・部署・注番で絞り込み" },
        { title: "ソート",           desc: "列ヘッダークリックで並び替え" },
        { title: "編集・削除",       desc: "自分の記録は編集・削除が可能" },
        { title: "範囲拡張",         desc: "+1ヶ月 / +3ヶ月 / +6ヶ月 / +12ヶ月 のボタンで過去を遡る" },
        { title: "ページネーション", desc: "大量データは自動で分割ロード" },
    ];
    const fw = 4.4, fh = 1.0, fgap = 0.15;
    const fsx = 0.5;
    const fsy = 1.8;
    features.forEach((f, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const x = fsx + col * (fw + fgap);
        const y = fsy + row * (fh + fgap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: y, w: fw, h: fh,
            fill: { color: C.white }, line: { color: C.border, width: 1 }
        });
        s.addShape(pres.shapes.OVAL, {
            x: x + 0.2, y: y + 0.3, w: 0.4, h: 0.4,
            fill: { color: C.primary }, line: { color: C.primary }
        });
        s.addText(String(i + 1), {
            x: x + 0.2, y: y + 0.3, w: 0.4, h: 0.4,
            fontSize: 14, fontFace: FONT_HEADER, bold: true, color: C.white,
            align: "center", valign: "middle", margin: 0
        });
        s.addText(f.title, {
            x: x + 0.7, y: y + 0.15, w: fw - 0.85, h: 0.4,
            fontSize: 13, fontFace: FONT_BODY, bold: true, color: C.primary, margin: 0
        });
        s.addText(f.desc, {
            x: x + 0.7, y: y + 0.5, w: fw - 0.85, h: 0.4,
            fontSize: 10, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    // 下部ヒント
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 5.0, w: 9, h: 0.3,
        fill: { color: C.lightBg }, line: { color: C.lightBg }
    });
    s.addText("💡 他の人の記録は閲覧のみ。編集・削除できるのは自分が入力した記録だけです。", {
        x: 0.5, y: 5.0, w: 9, h: 0.3,
        fontSize: 10, fontFace: FONT_BODY, italic: true, color: C.primary,
        align: "center", valign: "middle", margin: 0
    });

    addFooter(s, 10);
}

// ============================================================
// Slide 11: データ解析 — 集計画面
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 9 章", "データ解析 — 集計画面");

    s.addText("フィルターを組み合わせて、欲しい切り口で集計結果を見ることができます。", {
        x: 0.5, y: 1.65, w: 9, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, color: C.textMid, margin: 0
    });

    // 左: 集計軸
    addInfoBox(s, 0.5, 2.05, 4.5, 3.05, "集計軸", [
        { label: "期間", desc: "日付範囲 or 会計年度で絞り込み" },
        { label: "部署", desc: "記録時点の所属部署で集計" },
        { label: "担当者", desc: "個人別の時間合計" },
        { label: "注番", desc: "注番ごとの時間合計・目標達成率" },
        { label: "場所", desc: "社内 / 社外" },
    ], C.primary);

    // 右: 表示内容
    addInfoBox(s, 5.1, 2.05, 4.4, 3.05, "表示内容", [
        { label: "合計時間", desc: "フィルター条件での総計" },
        { label: "割合", desc: "カテゴリ別の構成比" },
        { label: "目標達成率", desc: "注番別に達成状況を色分け" },
        { label: "前年比較", desc: "ワンクリックで呼び出し" },
        { label: "CSV ダウンロード", desc: "Excel でさらに加工可能" },
    ], C.accent);

    addFooter(s, 11);
}

// ============================================================
// Slide 12: データ解析 — 会計年度集計
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 10 章", "データ解析 — 会計年度集計");

    s.addText("会計年度（9月〜翌8月）単位で集計します。現在は第 46 期です。", {
        x: 0.5, y: 1.65, w: 9, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, color: C.textMid, margin: 0
    });

    // 3モード
    const modes = [
        {
            title: "手動",
            subtitle: "カスタム期間",
            desc: "日付範囲を自由に指定",
            detail: "下のフィルターで dateFrom / dateTo を設定",
            color: C.primary3,
        },
        {
            title: "半期",
            subtitle: "前期 / 後期",
            desc: "6 ヶ月単位で集計",
            detail: "前期: 9〜2月  /  後期: 3〜8月",
            color: C.accent,
        },
        {
            title: "通年",
            subtitle: "期全体",
            desc: "12 ヶ月単位で集計",
            detail: "第○期の 9/1 〜 翌 8/31",
            color: C.primary,
        },
    ];
    const mw = 2.95, mh = 3.0, mgap = 0.15;
    const msx = (10 - (mw * 3 + mgap * 2)) / 2;
    const msy = 2.0;
    modes.forEach((m, i) => {
        const x = msx + i * (mw + mgap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: msy, w: mw, h: mh,
            fill: { color: C.white }, line: { color: m.color, width: 2 },
            shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.08 }
        });
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: msy, w: mw, h: 0.15,
            fill: { color: m.color }, line: { color: m.color }
        });
        s.addText(m.title, {
            x: x, y: msy + 0.3, w: mw, h: 0.6,
            fontSize: 30, fontFace: FONT_HEADER, bold: true, color: m.color,
            align: "center", margin: 0
        });
        s.addText(m.subtitle, {
            x: x, y: msy + 0.95, w: mw, h: 0.35,
            fontSize: 12, fontFace: FONT_BODY, italic: true, color: C.textMid,
            align: "center", margin: 0
        });
        s.addShape(pres.shapes.LINE, {
            x: x + mw/2 - 0.4, y: msy + 1.4, w: 0.8, h: 0,
            line: { color: m.color, width: 1 }
        });
        s.addText(m.desc, {
            x: x + 0.2, y: msy + 1.55, w: mw - 0.4, h: 0.5,
            fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.textDark,
            align: "center", margin: 0
        });
        s.addText(m.detail, {
            x: x + 0.2, y: msy + 2.1, w: mw - 0.4, h: 0.7,
            fontSize: 10, fontFace: FONT_BODY, color: C.textMid,
            align: "center", margin: 0
        });
    });

    addFooter(s, 12);
}

// ============================================================
// Slide 13: データ解析 — 前年比較
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 11 章", "データ解析 — 前年同期比較");

    s.addText("集計画面の「前年同期比較」ボタンで、前期と今期を部署別に比較できます。", {
        x: 0.5, y: 1.65, w: 9, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, color: C.textMid, margin: 0
    });

    // 比較イメージ
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.1, w: 9, h: 2.1,
        fill: { color: C.white }, line: { color: C.border, width: 1 },
        shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.06 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 2.1, w: 9, h: 0.45,
        fill: { color: C.primary }, line: { color: C.primary }
    });
    s.addText("前年同期比較（イメージ）", {
        x: 0.7, y: 2.1, w: 9, h: 0.45,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.white,
        valign: "middle", margin: 0
    });

    const compareData = [
        { dept: "営業グループ",    last: 1200, cur: 1350, diff: "+12.5%" },
        { dept: "技術グループ",    last: 1800, cur: 1720, diff: "-4.4%"  },
        { dept: "品質管理グループ", last: 900,  cur: 950,  diff: "+5.6%"  },
    ];
    compareData.forEach((d, i) => {
        const y = 2.7 + i * 0.45;
        s.addText(d.dept, {
            x: 0.8, y: y, w: 3, h: 0.4,
            fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.textDark,
            valign: "middle", margin: 0
        });
        s.addText(`45期: ${d.last}h`, {
            x: 4.0, y: y, w: 1.8, h: 0.4,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid,
            valign: "middle", margin: 0
        });
        s.addText("→", {
            x: 5.8, y: y, w: 0.3, h: 0.4,
            fontSize: 14, fontFace: FONT_BODY, color: C.textLight,
            align: "center", valign: "middle", margin: 0
        });
        s.addText(`46期: ${d.cur}h`, {
            x: 6.1, y: y, w: 1.8, h: 0.4,
            fontSize: 11, fontFace: FONT_BODY, bold: true, color: C.primary,
            valign: "middle", margin: 0
        });
        const diffColor = d.diff.startsWith("+") ? C.success : C.danger;
        s.addText(d.diff, {
            x: 7.9, y: y, w: 1.5, h: 0.4,
            fontSize: 12, fontFace: FONT_BODY, bold: true, color: diffColor,
            align: "right", valign: "middle", margin: 0
        });
    });

    // ヒント
    s.addText("💡 何に時間がかかっているか、前年と比べて分かれば、来期の改善方針が立てやすくなります。", {
        x: 0.5, y: 4.5, w: 9, h: 0.4,
        fontSize: 11, fontFace: FONT_BODY, italic: true, color: C.primary,
        align: "center", margin: 0
    });

    addFooter(s, 13);
}

// ============================================================
// Slide 14: CSV 出力
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 12 章", "データ解析 — CSV 出力");

    s.addText("集計画面で表示中のデータを CSV ファイルとしてダウンロードできます。", {
        x: 0.5, y: 1.65, w: 9, h: 0.35,
        fontSize: 12, fontFace: FONT_BODY, color: C.textMid, margin: 0
    });

    // 2モード
    const modes = [
        {
            title: "基本モード",
            desc: "見たまま出力",
            color: C.primary,
            items: [
                "画面上の集計結果をそのまま",
                "部署別・担当者別・注番別",
                "合計時間と割合を含む",
                "そのまま Excel で開ける",
            ]
        },
        {
            title: "詳細モード",
            desc: "元データ出力",
            color: C.accent,
            items: [
                "各作業記録を 1 行ずつ出力",
                "日付・担当者・注番・内容・時間",
                "ピボットテーブルで自由に分析",
                "他システムへのインポート用",
            ]
        },
    ];
    const mw = 4.4, mh = 2.95, mgap = 0.2;
    const msx = (10 - (mw * 2 + mgap)) / 2;
    const msy = 2.1;
    modes.forEach((m, i) => {
        const x = msx + i * (mw + mgap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: msy, w: mw, h: mh,
            fill: { color: C.white }, line: { color: C.border, width: 1 },
            shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.08 }
        });
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: msy, w: mw, h: 0.15,
            fill: { color: m.color }, line: { color: m.color }
        });
        s.addText(m.title, {
            x: x + 0.3, y: msy + 0.3, w: mw - 0.6, h: 0.45,
            fontSize: 20, fontFace: FONT_HEADER, bold: true, color: m.color, margin: 0
        });
        s.addText(m.desc, {
            x: x + 0.3, y: msy + 0.8, w: mw - 0.6, h: 0.3,
            fontSize: 11, fontFace: FONT_BODY, italic: true, color: C.textMid, margin: 0
        });
        s.addShape(pres.shapes.LINE, {
            x: x + 0.3, y: msy + 1.2, w: mw - 0.6, h: 0,
            line: { color: C.border, width: 1 }
        });
        m.items.forEach((it, j) => {
            const iy = msy + 1.35 + j * 0.38;
            s.addShape(pres.shapes.OVAL, {
                x: x + 0.3, y: iy + 0.12, w: 0.13, h: 0.13,
                fill: { color: m.color }, line: { color: m.color }
            });
            s.addText(it, {
                x: x + 0.5, y: iy, w: mw - 0.7, h: 0.35,
                fontSize: 11, fontFace: FONT_BODY, color: C.textDark, margin: 0
            });
        });
    });

    addFooter(s, 14);
}

// ============================================================
// Slide 15: 環境設定 — ユーザー管理
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 13 章", "環境設定 — ユーザー管理（管理者）");

    // 警告バー
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 1.65, w: 9, h: 0.35,
        fill: { color: C.warning }, line: { color: C.warning }
    });
    s.addText("⚠ 管理者権限を持つユーザーのみアクセス可能", {
        x: 0.5, y: 1.65, w: 9, h: 0.35,
        fontSize: 11, fontFace: FONT_BODY, bold: true, color: C.white,
        align: "center", valign: "middle", margin: 0
    });

    const features = [
        { title: "新規ユーザー登録",  desc: "氏名・部署・メール・権限を設定" },
        { title: "ユーザー編集",      desc: "既存ユーザーの情報を変更" },
        { title: "ユーザー削除",      desc: "退職者などを無効化（データ保持）" },
        { title: "グループフィルター", desc: "ユーザー一覧を部署で絞り込み" },
        { title: "PIN リセット",      desc: "パスワード忘れ時の対応" },
        { title: "権限変更",          desc: "member / leader / admin の切替" },
    ];
    const fw = 4.4, fh = 0.95, fgap = 0.12;
    const fsx = 0.5;
    const fsy = 2.2;
    features.forEach((f, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const x = fsx + col * (fw + fgap);
        const y = fsy + row * (fh + fgap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: y, w: fw, h: fh,
            fill: { color: C.white }, line: { color: C.border, width: 1 }
        });
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: y, w: 0.1, h: fh,
            fill: { color: C.primary }, line: { color: C.primary }
        });
        s.addText(f.title, {
            x: x + 0.25, y: y + 0.12, w: fw - 0.4, h: 0.35,
            fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.primary, margin: 0
        });
        s.addText(f.desc, {
            x: x + 0.25, y: y + 0.48, w: fw - 0.4, h: 0.4,
            fontSize: 10, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    addFooter(s, 15);
}

// ============================================================
// Slide 16: 環境設定 — グループ・注番管理
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 14 章", "環境設定 — グループ・注番管理");

    // 3カラム
    const cols = [
        {
            title: "グループ",
            desc: "部署マスタの管理",
            items: ["新規追加", "名前変更", "削除（所属 0 名時のみ）"],
            color: C.primary,
        },
        {
            title: "作業内容",
            desc: "プルダウン項目の管理",
            items: ["グループ別に項目管理", "追加 / 並び替え / 削除", "ドラッグ＆ドロップ"],
            color: C.accent,
        },
        {
            title: "注番管理",
            desc: "進行中注番の管理",
            items: ["検索・絞り込み", "得意先・件名の編集", "手動クローズ / 再開"],
            color: C.accent2,
        },
    ];
    const cw = 2.95, ch = 3.0, cgap = 0.15;
    const csx = (10 - (cw * 3 + cgap * 2)) / 2;
    const csy = 1.9;
    cols.forEach((c, i) => {
        const x = csx + i * (cw + cgap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: csy, w: cw, h: ch,
            fill: { color: C.white }, line: { color: C.border, width: 1 },
            shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.08 }
        });
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: csy, w: cw, h: 0.6,
            fill: { color: c.color }, line: { color: c.color }
        });
        s.addText(c.title, {
            x: x + 0.2, y: csy + 0.08, w: cw - 0.4, h: 0.3,
            fontSize: 16, fontFace: FONT_HEADER, bold: true, color: C.white, margin: 0
        });
        s.addText(c.desc, {
            x: x + 0.2, y: csy + 0.35, w: cw - 0.4, h: 0.25,
            fontSize: 10, fontFace: FONT_BODY, italic: true, color: C.white, margin: 0
        });
        c.items.forEach((it, j) => {
            const iy = csy + 0.85 + j * 0.55;
            s.addShape(pres.shapes.OVAL, {
                x: x + 0.25, y: iy + 0.1, w: 0.18, h: 0.18,
                fill: { color: c.color }, line: { color: c.color }
            });
            s.addText(it, {
                x: x + 0.5, y: iy, w: cw - 0.7, h: 0.4,
                fontSize: 11, fontFace: FONT_BODY, color: C.textDark, margin: 0
            });
        });
    });

    // 下部注意
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 5.0, w: 9, h: 0.3,
        fill: { color: C.lightBg }, line: { color: C.lightBg }
    });
    s.addText("💡 2 ヶ月入力のない注番は自動でクローズされます。管理者は手動で解除も可能です。", {
        x: 0.5, y: 5.0, w: 9, h: 0.3,
        fontSize: 10, fontFace: FONT_BODY, italic: true, color: C.primary,
        align: "center", valign: "middle", margin: 0
    });

    addFooter(s, 16);
}

// ============================================================
// Slide 17: トラブルシューティング
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.light };

    addHeader(s, "第 15 章", "トラブルシューティング");

    const troubles = [
        {
            q: "ログインできない",
            a: "PIN を 3 回以上間違えていないか確認。忘れた場合は管理者にリセット依頼。",
        },
        {
            q: "「通信エラー」と出る",
            a: "サーバーが停止している可能性。IT 担当者に連絡してください。",
        },
        {
            q: "注番が候補に出てこない",
            a: "2 ヶ月入力のない注番は自動でクローズされています。管理者に解除依頼。",
        },
        {
            q: "保存ボタンを押しても反応がない",
            a: "必須項目（時間・作業内容など）が空白になっていないか確認してください。",
        },
        {
            q: "過去のデータが表示されない",
            a: "データ一覧画面の「+3ヶ月」「+6ヶ月」ボタンで読み込み範囲を拡張できます。",
        },
    ];
    const ty = 1.75, th = 0.68, tgap = 0.05;
    troubles.forEach((t, i) => {
        const y = ty + i * (th + tgap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: 0.5, y: y, w: 9, h: th,
            fill: { color: C.white }, line: { color: C.border, width: 1 }
        });
        // Q badge
        s.addShape(pres.shapes.OVAL, {
            x: 0.65, y: y + 0.15, w: 0.38, h: 0.38,
            fill: { color: C.danger }, line: { color: C.danger }
        });
        s.addText("Q", {
            x: 0.65, y: y + 0.15, w: 0.38, h: 0.38,
            fontSize: 13, fontFace: FONT_HEADER, bold: true, color: C.white,
            align: "center", valign: "middle", margin: 0
        });
        s.addText(t.q, {
            x: 1.15, y: y + 0.08, w: 8.2, h: 0.3,
            fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.textDark, margin: 0
        });
        s.addText(t.a, {
            x: 1.15, y: y + 0.35, w: 8.2, h: 0.35,
            fontSize: 10, fontFace: FONT_BODY, color: C.textMid, margin: 0
        });
    });

    addFooter(s, 17);
}

// ============================================================
// Slide 18: Q&A / お問い合わせ
// ============================================================
{
    const s = pres.addSlide();
    s.background = { color: C.primary };

    // 装飾
    s.addShape(pres.shapes.OVAL, {
        x: 6, y: -2, w: 6, h: 6,
        fill: { color: C.primary2, transparency: 75 },
        line: { color: C.primary2, transparency: 75 }
    });

    // 上部アクセント
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 0, w: 10, h: 0.15, fill: { color: C.accent }, line: { color: C.accent }
    });

    // ラベル
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: 0.6, w: 0.15, h: 0.4, fill: { color: C.accent }, line: { color: C.accent }
    });
    s.addText("お問い合わせ", {
        x: 0.75, y: 0.6, w: 9, h: 0.4,
        fontSize: 12, fontFace: FONT_BODY, bold: true, color: C.accent,
        valign: "middle", charSpacing: 2, margin: 0
    });

    // タイトル
    s.addText("分からないことがあれば", {
        x: 0.5, y: 1.3, w: 9, h: 0.7,
        fontSize: 32, fontFace: FONT_HEADER, bold: true, color: C.white, margin: 0
    });
    s.addText("遠慮なく管理者まで。", {
        x: 0.5, y: 2.05, w: 9, h: 0.7,
        fontSize: 32, fontFace: FONT_HEADER, bold: true, color: C.accent2, margin: 0
    });

    // 連絡先カード
    const contacts = [
        { title: "操作について", who: "まずは管理者へ",   icon: "①" },
        { title: "エラー・不具合", who: "IT 担当者へ",     icon: "②" },
        { title: "改善要望",      who: "何でも歓迎",      icon: "③" },
    ];
    const cw = 2.9, ch = 1.5, cgap = 0.15;
    const csx = (10 - (cw * 3 + cgap * 2)) / 2;
    const csy = 3.2;
    contacts.forEach((c, i) => {
        const x = csx + i * (cw + cgap);
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: csy, w: cw, h: ch,
            fill: { color: C.white }, line: { color: C.accent, width: 1.5 }
        });
        s.addShape(pres.shapes.RECTANGLE, {
            x: x, y: csy, w: cw, h: 0.1,
            fill: { color: C.accent }, line: { color: C.accent }
        });
        s.addText(c.icon, {
            x: x, y: csy + 0.2, w: cw, h: 0.4,
            fontSize: 20, fontFace: FONT_HEADER, bold: true, color: C.accent,
            align: "center", margin: 0
        });
        s.addText(c.title, {
            x: x, y: csy + 0.65, w: cw, h: 0.35,
            fontSize: 13, fontFace: FONT_BODY, bold: true, color: C.primary,
            align: "center", margin: 0
        });
        s.addText(c.who, {
            x: x, y: csy + 0.98, w: cw, h: 0.35,
            fontSize: 11, fontFace: FONT_BODY, color: C.textMid,
            align: "center", margin: 0
        });
    });

    // 下部バー
    s.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: 5.2, w: 10, h: 0.425, fill: { color: C.accent }, line: { color: C.accent }
    });
    s.addText("完全マニュアル  —  作業時間 一元管理システム  Version 1.0", {
        x: 0.5, y: 5.2, w: 9, h: 0.425,
        fontSize: 11, fontFace: FONT_BODY, bold: true, color: C.primary,
        valign: "middle", margin: 0
    });
}

// ============================================================
// 書き出し
// ============================================================
pres.writeFile({ fileName: "03_完全マニュアル.pptx" })
    .then(fileName => console.log(`✓ ${fileName} を作成しました`))
    .catch(err => { console.error(err); process.exit(1); });
