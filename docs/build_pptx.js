// 最強の公式LINE設計ガイド - PowerPoint generator
const path = require("path");
const globalNodeModules = "C:\\Users\\ogaig\\AppData\\Roaming\\npm\\node_modules";
const pptxgen = require(path.join(globalNodeModules, "pptxgenjs"));

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
pres.author = "キャリサポ";
pres.title = "最強の公式LINE設計ガイド";

// ============ Color Palette ============
const C = {
  navy: "0A2540",       // primary dark
  deepTeal: "0D4F4C",   // brand
  teal: "1A8896",       // mid
  mint: "16C79A",       // bright accent
  line: "06C755",       // LINE green
  gold: "FFC857",       // warm accent
  cream: "FDFAF6",      // light bg
  paper: "F4F1EA",      // washi paper
  white: "FFFFFF",
  text: "1E293B",       // dark text
  sub: "64748B",        // muted
  red: "E63946",        // emphasis
};

const FONT_H = "Yu Gothic UI";
const FONT_B = "Yu Gothic UI";

// ============ Helper Functions ============
function addBg(slide, color) {
  slide.background = { color: color };
}

function addHeader(slide, num, title, accent = C.mint) {
  // Top bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 13.3, h: 0.55,
    fill: { color: C.navy }, line: { color: C.navy, width: 0 }
  });
  // Page number circle
  slide.addShape(pres.shapes.OVAL, {
    x: 0.4, y: 0.1, w: 0.35, h: 0.35,
    fill: { color: accent }, line: { color: accent, width: 0 }
  });
  slide.addText(String(num).padStart(2, "0"), {
    x: 0.4, y: 0.1, w: 0.35, h: 0.35,
    fontSize: 11, bold: true, color: C.navy, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });
  // Title text
  slide.addText(title, {
    x: 0.95, y: 0.1, w: 11, h: 0.35,
    fontSize: 14, bold: true, color: C.white, fontFace: FONT_H, margin: 0,
    valign: "middle"
  });
  // Brand
  slide.addText("CARRE SUPPORT", {
    x: 11.8, y: 0.1, w: 1.4, h: 0.35,
    fontSize: 9, color: C.gold, align: "right", valign: "middle",
    fontFace: FONT_H, charSpacing: 2, margin: 0
  });
}

function addFooter(slide) {
  slide.addShape(pres.shapes.LINE, {
    x: 0.5, y: 7.15, w: 12.3, h: 0,
    line: { color: C.sub, width: 0.5 }
  });
  slide.addText("最強の公式LINE設計ガイド | キャリサポ式LINE運用術", {
    x: 0.5, y: 7.2, w: 12.3, h: 0.25,
    fontSize: 9, color: C.sub, fontFace: FONT_B, margin: 0
  });
}

function makeShadow() {
  return { type: "outer", color: "000000", blur: 8, offset: 2, angle: 90, opacity: 0.12 };
}

// =====================================================
// SLIDE 1 — TITLE
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.navy);

  // Decorative gradient blocks
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 13.3, h: 0.15,
    fill: { color: C.mint }, line: { width: 0 }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 7.35, w: 13.3, h: 0.15,
    fill: { color: C.gold }, line: { width: 0 }
  });

  // Side accent
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.25, h: 7.5,
    fill: { color: C.mint }, line: { width: 0 }
  });

  // Eyebrow
  s.addText("OFFICIAL LINE DESIGN GUIDE", {
    x: 0.9, y: 1.3, w: 12, h: 0.4,
    fontSize: 14, color: C.gold, fontFace: FONT_H, charSpacing: 6, bold: true
  });

  // Main title
  s.addText("最強の公式LINE設計ガイド", {
    x: 0.9, y: 1.85, w: 12, h: 1.3,
    fontSize: 54, bold: true, color: C.white, fontFace: FONT_H
  });

  // Sub
  s.addText("エルメで作る、キャリサポ式LINE運用術", {
    x: 0.9, y: 3.2, w: 12, h: 0.7,
    fontSize: 24, color: C.mint, fontFace: FONT_H
  });

  // Divider
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.9, y: 4.1, w: 1.2, h: 0.06,
    fill: { color: C.gold }, line: { width: 0 }
  });

  // Description
  s.addText("DM → LINE登録 → ステップ配信 → 個別相談 → 成約\n問い合わせを最大化する自動化フロー、完全公開。", {
    x: 0.9, y: 4.3, w: 12, h: 1.3,
    fontSize: 16, color: C.cream, fontFace: FONT_B, paraSpaceAfter: 6
  });

  // Bottom info bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.9, y: 6.2, w: 11.5, h: 0.7,
    fill: { color: C.deepTeal }, line: { width: 0 }
  });
  s.addText([
    { text: "📘 全20スライド", options: { color: C.white, bold: true } },
    { text: "    |    ", options: { color: C.mint } },
    { text: "🎯 初心者OK", options: { color: C.white, bold: true } },
    { text: "    |    ", options: { color: C.mint } },
    { text: "💰 月額0円スタート可", options: { color: C.white, bold: true } },
  ], {
    x: 0.9, y: 6.2, w: 11.5, h: 0.7,
    fontSize: 14, fontFace: FONT_H, align: "center", valign: "middle", margin: 0
  });
}

// =====================================================
// SLIDE 2 — AGENDA
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 2, "AGENDA / 目次");
  addFooter(s);

  s.addText("このガイドで学べること", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.6,
    fontSize: 32, bold: true, color: C.navy, fontFace: FONT_H
  });
  s.addText("LINE登録者をファンに変え、自動で問い合わせを生む仕組み。", {
    x: 0.5, y: 1.5, w: 12.3, h: 0.4,
    fontSize: 14, color: C.sub, fontFace: FONT_B
  });

  // 4 chapters in cards
  const chapters = [
    { num: "01", title: "なぜ公式LINEなのか", desc: "SNSやメルマガを超える理由", color: C.teal },
    { num: "02", title: "エルメ完全入門", desc: "無料で始める拡張ツール", color: C.mint },
    { num: "03", title: "最強の構築フロー", desc: "ウェルカム〜ステップ配信", color: C.gold },
    { num: "04", title: "成果を出すコツ", desc: "KPI・改善・実践ロードマップ", color: C.line },
  ];

  chapters.forEach((c, i) => {
    const x = 0.5 + (i % 2) * 6.4;
    const y = 2.2 + Math.floor(i / 2) * 2.3;

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 6, h: 2,
      fill: { color: C.white }, line: { color: "E2E8F0", width: 0.5 },
      shadow: makeShadow()
    });
    // Left accent
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 0.12, h: 2,
      fill: { color: c.color }, line: { width: 0 }
    });
    // Number
    s.addText(c.num, {
      x: x + 0.35, y: y + 0.25, w: 1.2, h: 0.7,
      fontSize: 36, bold: true, color: c.color, fontFace: FONT_H, margin: 0
    });
    // Title
    s.addText(c.title, {
      x: x + 1.6, y: y + 0.35, w: 4.2, h: 0.6,
      fontSize: 20, bold: true, color: C.navy, fontFace: FONT_H, margin: 0
    });
    // Desc
    s.addText(c.desc, {
      x: x + 1.6, y: y + 1.0, w: 4.2, h: 0.6,
      fontSize: 13, color: C.sub, fontFace: FONT_B, margin: 0
    });
  });
}

// =====================================================
// SLIDE 3 — WHY LINE
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 3, "Chapter 1 — なぜ公式LINEなのか");
  addFooter(s);

  s.addText("公式LINE = 最強の顧客接点", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.7,
    fontSize: 32, bold: true, color: C.navy, fontFace: FONT_H
  });

  // Big stats
  const stats = [
    { num: "98%", label: "LINEメッセージ開封率", sub: "メルマガ平均20%の約5倍" },
    { num: "60%", label: "24時間以内既読率", sub: "他SNSと比べ圧倒的に高速" },
    { num: "9,500万人", label: "国内アクティブユーザー", sub: "ほぼ全世代が使用中" },
  ];

  stats.forEach((st, i) => {
    const x = 0.5 + i * 4.3;
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 1.8, w: 4, h: 2.5,
      fill: { color: C.navy }, line: { width: 0 }, shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 1.8, w: 4, h: 0.1,
      fill: { color: C.line }, line: { width: 0 }
    });
    s.addText(st.num, {
      x: x, y: 2.0, w: 4, h: 1.1,
      fontSize: 48, bold: true, color: C.line, align: "center", fontFace: FONT_H, margin: 0
    });
    s.addText(st.label, {
      x: x, y: 3.15, w: 4, h: 0.4,
      fontSize: 14, bold: true, color: C.white, align: "center", fontFace: FONT_H, margin: 0
    });
    s.addText(st.sub, {
      x: x + 0.2, y: 3.6, w: 3.6, h: 0.5,
      fontSize: 11, color: C.mint, align: "center", fontFace: FONT_B, margin: 0
    });
  });

  // Bottom message
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.7, w: 12.3, h: 2.2,
    fill: { color: C.white }, line: { color: "E2E8F0", width: 0.5 }, shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.7, w: 0.12, h: 2.2,
    fill: { color: C.gold }, line: { width: 0 }
  });
  s.addText("つまり、LINEは「届く・読まれる・反応される」唯一のチャネル。", {
    x: 0.85, y: 4.85, w: 11.8, h: 0.6,
    fontSize: 18, bold: true, color: C.navy, fontFace: FONT_H, margin: 0
  });
  s.addText([
    { text: "✓ ", options: { color: C.line, bold: true } },
    { text: "Threads/Instagramの投稿はタイムラインで流れていくが、LINEは1対1で必ず届く\n", options: { color: C.text } },
    { text: "✓ ", options: { color: C.line, bold: true } },
    { text: "メルマガと違い、開封されないリスクが圧倒的に低い\n", options: { color: C.text } },
    { text: "✓ ", options: { color: C.line, bold: true } },
    { text: "リッチメニュー・自動応答・タグ管理まで全部できる「ミニWebサイト」になる", options: { color: C.text } },
  ], {
    x: 0.85, y: 5.5, w: 11.8, h: 1.4,
    fontSize: 13, fontFace: FONT_B, paraSpaceAfter: 4, margin: 0
  });
}

// =====================================================
// SLIDE 4 — Common Mistakes
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 4, "現状の失敗パターン");
  addFooter(s);

  s.addText("こんな運用、していませんか?", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.7,
    fontSize: 32, bold: true, color: C.navy, fontFace: FONT_H
  });
  s.addText("登録者は集めても、ほとんどが活かされていない現実。", {
    x: 0.5, y: 1.55, w: 12.3, h: 0.4,
    fontSize: 14, color: C.sub, fontFace: FONT_B
  });

  const mistakes = [
    { icon: "❌", title: "ウェルカム1通だけ", desc: "登録直後の関係構築チャンスを逃している" },
    { icon: "❌", title: "ステップ配信ゼロ", desc: "登録者を放置 → 興味が冷めて離脱" },
    { icon: "❌", title: "情報を一気に詰め込み", desc: "1通が長すぎて読まれない" },
    { icon: "❌", title: "CTAが弱い・ない", desc: "「で、何をすればいいの?」状態" },
    { icon: "❌", title: "リッチメニューが装飾だけ", desc: "タップ後の動線が設計されていない" },
    { icon: "❌", title: "全員に同じ配信", desc: "属性別の最適化ができていない" },
  ];

  mistakes.forEach((m, i) => {
    const x = 0.5 + (i % 3) * 4.3;
    const y = 2.15 + Math.floor(i / 3) * 2.4;

    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 4, h: 2.1,
      fill: { color: C.white }, line: { color: "FFD6D6", width: 1 }, shadow: makeShadow()
    });
    s.addText(m.icon, {
      x: x + 0.2, y: y + 0.2, w: 0.6, h: 0.6,
      fontSize: 28, fontFace: FONT_H, margin: 0
    });
    s.addText(m.title, {
      x: x + 0.85, y: y + 0.25, w: 3.0, h: 0.5,
      fontSize: 16, bold: true, color: C.red, fontFace: FONT_H, margin: 0
    });
    s.addText(m.desc, {
      x: x + 0.2, y: y + 0.95, w: 3.65, h: 1.05,
      fontSize: 12, color: C.text, fontFace: FONT_B, margin: 0
    });
  });
}

// =====================================================
// SLIDE 5 — What is エルメ
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 5, "Chapter 2 — エルメ完全入門");
  addFooter(s);

  s.addText("エルメ(L-Message)とは?", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.7,
    fontSize: 32, bold: true, color: C.navy, fontFace: FONT_H
  });

  // Left: explanation
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.8, w: 6.2, h: 5.2,
    fill: { color: C.white }, line: { color: "E2E8F0", width: 0.5 }, shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.8, w: 6.2, h: 0.6,
    fill: { color: C.line }, line: { width: 0 }
  });
  s.addText("公式LINEの拡張ツール", {
    x: 0.5, y: 1.8, w: 6.2, h: 0.6,
    fontSize: 16, bold: true, color: C.white, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });
  s.addText([
    { text: "公式LINE本体だけでは、単発配信や簡単なあいさつしかできません。\n", options: { breakLine: true } },
    { text: "\n", options: { breakLine: true } },
    { text: "エルメを連携することで...\n", options: { bold: true, color: C.navy, breakLine: true } },
    { text: "\n", options: { breakLine: true } },
    { text: "✓ 時間差配信(30秒後/1分後/3日後)\n", options: { color: C.line, breakLine: true } },
    { text: "✓ 7日間ステップ配信\n", options: { color: C.line, breakLine: true } },
    { text: "✓ ユーザーへのタグ付け\n", options: { color: C.line, breakLine: true } },
    { text: "✓ タグ別シナリオ分岐\n", options: { color: C.line, breakLine: true } },
    { text: "✓ 予約システム・フォーム\n", options: { color: C.line, breakLine: true } },
    { text: "✓ 流入経路分析\n", options: { color: C.line, breakLine: true } },
  ], {
    x: 0.8, y: 2.6, w: 5.7, h: 4.3,
    fontSize: 13, color: C.text, fontFace: FONT_B, margin: 0
  });

  // Right: visual diagram
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.0, y: 1.8, w: 5.8, h: 5.2,
    fill: { color: C.navy }, line: { width: 0 }, shadow: makeShadow()
  });
  s.addText("構造イメージ", {
    x: 7.0, y: 2.0, w: 5.8, h: 0.4,
    fontSize: 14, bold: true, color: C.gold, align: "center", fontFace: FONT_H, margin: 0
  });

  // LINE box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.6, y: 2.6, w: 4.6, h: 1.0,
    fill: { color: C.line }, line: { width: 0 }
  });
  s.addText("公式LINE(本体・無料)", {
    x: 7.6, y: 2.6, w: 4.6, h: 0.5,
    fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });
  s.addText("友だち・配信機能・リッチメニュー", {
    x: 7.6, y: 3.05, w: 4.6, h: 0.5,
    fontSize: 10, color: C.white, align: "center", valign: "middle",
    fontFace: FONT_B, margin: 0
  });

  // Plus
  s.addText("+", {
    x: 7.6, y: 3.7, w: 4.6, h: 0.4,
    fontSize: 24, bold: true, color: C.gold, align: "center", fontFace: FONT_H, margin: 0
  });

  // Erume box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.6, y: 4.2, w: 4.6, h: 1.0,
    fill: { color: C.mint }, line: { width: 0 }
  });
  s.addText("エルメ(拡張・無料〜)", {
    x: 7.6, y: 4.2, w: 4.6, h: 0.5,
    fontSize: 14, bold: true, color: C.navy, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });
  s.addText("ステップ配信・タグ・分岐・予約", {
    x: 7.6, y: 4.65, w: 4.6, h: 0.5,
    fontSize: 10, color: C.navy, align: "center", valign: "middle",
    fontFace: FONT_B, margin: 0
  });

  // Result
  s.addText("↓", {
    x: 7.6, y: 5.3, w: 4.6, h: 0.4,
    fontSize: 24, color: C.gold, align: "center", fontFace: FONT_H, margin: 0
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.6, y: 5.8, w: 4.6, h: 0.9,
    fill: { color: C.gold }, line: { width: 0 }
  });
  s.addText("最強の自動化LINE", {
    x: 7.6, y: 5.8, w: 4.6, h: 0.9,
    fontSize: 16, bold: true, color: C.navy, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });
}

// =====================================================
// SLIDE 6 — エルメ vs Lステップ vs 公式LINE
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 6, "ツール比較 — どれを選ぶべきか");
  addFooter(s);

  s.addText("公式LINE単体 vs エルメ vs Lステップ", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.7,
    fontSize: 30, bold: true, color: C.navy, fontFace: FONT_H
  });

  // Comparison table
  const tableData = [
    [
      { text: "機能", options: { fill: { color: C.navy }, color: C.white, bold: true, align: "center", valign: "middle", fontFace: FONT_H } },
      { text: "公式LINE単体", options: { fill: { color: C.navy }, color: C.white, bold: true, align: "center", valign: "middle", fontFace: FONT_H } },
      { text: "エルメ(無料〜)", options: { fill: { color: C.line }, color: C.white, bold: true, align: "center", valign: "middle", fontFace: FONT_H } },
      { text: "Lステップ", options: { fill: { color: C.navy }, color: C.white, bold: true, align: "center", valign: "middle", fontFace: FONT_H } },
    ],
    [
      { text: "月額", options: { bold: true, fontFace: FONT_H, valign: "middle" } },
      { text: "0円〜", options: { align: "center", valign: "middle", fontFace: FONT_B } },
      { text: "0円〜", options: { align: "center", valign: "middle", fontFace: FONT_B, color: C.line, bold: true } },
      { text: "2,980円〜", options: { align: "center", valign: "middle", fontFace: FONT_B } },
    ],
    [
      { text: "ステップ配信", options: { bold: true, fontFace: FONT_H, valign: "middle" } },
      { text: "△ 簡易のみ", options: { align: "center", valign: "middle", fontFace: FONT_B } },
      { text: "◎", options: { align: "center", valign: "middle", fontSize: 18, color: C.line, bold: true, fontFace: FONT_H } },
      { text: "◎", options: { align: "center", valign: "middle", fontSize: 18, fontFace: FONT_H } },
    ],
    [
      { text: "時間差配信", options: { bold: true, fontFace: FONT_H, valign: "middle" } },
      { text: "✗", options: { align: "center", valign: "middle", color: C.red, fontFace: FONT_H } },
      { text: "◎", options: { align: "center", valign: "middle", fontSize: 18, color: C.line, bold: true, fontFace: FONT_H } },
      { text: "◎", options: { align: "center", valign: "middle", fontSize: 18, fontFace: FONT_H } },
    ],
    [
      { text: "タグ管理・分岐", options: { bold: true, fontFace: FONT_H, valign: "middle" } },
      { text: "✗", options: { align: "center", valign: "middle", color: C.red, fontFace: FONT_H } },
      { text: "◎", options: { align: "center", valign: "middle", fontSize: 18, color: C.line, bold: true, fontFace: FONT_H } },
      { text: "◎", options: { align: "center", valign: "middle", fontSize: 18, fontFace: FONT_H } },
    ],
    [
      { text: "予約システム", options: { bold: true, fontFace: FONT_H, valign: "middle" } },
      { text: "✗", options: { align: "center", valign: "middle", color: C.red, fontFace: FONT_H } },
      { text: "◎", options: { align: "center", valign: "middle", fontSize: 18, color: C.line, bold: true, fontFace: FONT_H } },
      { text: "◎", options: { align: "center", valign: "middle", fontSize: 18, fontFace: FONT_H } },
    ],
    [
      { text: "操作の難易度", options: { bold: true, fontFace: FONT_H, valign: "middle" } },
      { text: "簡単", options: { align: "center", valign: "middle", fontFace: FONT_B } },
      { text: "簡単(直感的)", options: { align: "center", valign: "middle", fontFace: FONT_B, color: C.line, bold: true } },
      { text: "やや難しい", options: { align: "center", valign: "middle", fontFace: FONT_B } },
    ],
    [
      { text: "おすすめ規模", options: { bold: true, fontFace: FONT_H, valign: "middle" } },
      { text: "〜50人", options: { align: "center", valign: "middle", fontFace: FONT_B } },
      { text: "50〜1500人", options: { align: "center", valign: "middle", fontFace: FONT_B, color: C.line, bold: true } },
      { text: "1500人〜", options: { align: "center", valign: "middle", fontFace: FONT_B } },
    ],
  ];

  s.addTable(tableData, {
    x: 0.5, y: 1.7, w: 12.3, h: 4.6,
    colW: [3.0, 3.1, 3.1, 3.1],
    rowH: 0.55,
    border: { pt: 0.5, color: "E2E8F0" },
    fontSize: 13,
    color: C.text,
  });

  // Conclusion
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 6.45, w: 12.3, h: 0.6,
    fill: { color: C.line }, line: { width: 0 }
  });
  s.addText("👉 初心者・登録者〜1500人なら エルメ(無料〜) が圧倒的におすすめ", {
    x: 0.5, y: 6.45, w: 12.3, h: 0.6,
    fontSize: 16, bold: true, color: C.white, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });
}

// =====================================================
// SLIDE 7 — エルメで何ができる
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 7, "エルメの主要機能");
  addFooter(s);

  s.addText("エルメで自動化できる8つのこと", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.7,
    fontSize: 30, bold: true, color: C.navy, fontFace: FONT_H
  });

  const features = [
    { icon: "⏱", title: "時間差配信", desc: "30秒・1分・1時間後の自動送信" },
    { icon: "📅", title: "ステップ配信", desc: "登録◯日後に自動でメッセージ" },
    { icon: "🏷", title: "タグ付け", desc: "ユーザーを属性別に自動分類" },
    { icon: "🔀", title: "シナリオ分岐", desc: "タグ別に違うメッセージを配信" },
    { icon: "📋", title: "フォーム", desc: "アンケート・申込フォーム作成" },
    { icon: "📆", title: "予約システム", desc: "面談予約のカレンダー連携" },
    { icon: "📊", title: "流入分析", desc: "どこから来たかを計測" },
    { icon: "🎯", title: "セグメント配信", desc: "条件に合う人だけに送信" },
  ];

  features.forEach((f, i) => {
    const x = 0.5 + (i % 4) * 3.18;
    const y = 1.85 + Math.floor(i / 4) * 2.55;

    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 2.95, h: 2.35,
      fill: { color: C.white }, line: { color: "E2E8F0", width: 0.5 }, shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 2.95, h: 0.08,
      fill: { color: C.mint }, line: { width: 0 }
    });

    s.addText(f.icon, {
      x: x, y: y + 0.3, w: 2.95, h: 0.7,
      fontSize: 36, align: "center", fontFace: FONT_H, margin: 0
    });
    s.addText(f.title, {
      x: x, y: y + 1.05, w: 2.95, h: 0.45,
      fontSize: 16, bold: true, color: C.navy, align: "center", fontFace: FONT_H, margin: 0
    });
    s.addText(f.desc, {
      x: x + 0.15, y: y + 1.55, w: 2.65, h: 0.7,
      fontSize: 11, color: C.sub, align: "center", fontFace: FONT_B, margin: 0
    });
  });
}

// =====================================================
// SLIDE 8 — Pricing
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 8, "エルメ料金プラン");
  addFooter(s);

  s.addText("無料から始められる、明朗な料金体系", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.7,
    fontSize: 30, bold: true, color: C.navy, fontFace: FONT_H
  });

  const plans = [
    {
      name: "フリー", price: "0円", limit: "月1,000通",
      target: "登録者〜100人", color: C.line, popular: true,
      features: ["基本機能すべて", "ステップ配信", "タグ・分岐", "リッチメニュー連携"]
    },
    {
      name: "スタンダード", price: "10,780円", limit: "月15,000通",
      target: "登録者〜1,500人", color: C.teal, popular: false,
      features: ["フリーの全機能", "高度な分析", "API連携", "優先サポート"]
    },
    {
      name: "プロ", price: "33,000円", limit: "無制限",
      target: "登録者〜5,000人", color: C.navy, popular: false,
      features: ["スタンダードの全機能", "無制限配信", "専任サポート", "カスタム対応"]
    },
  ];

  plans.forEach((p, i) => {
    const x = 0.5 + i * 4.3;
    const y = 1.85;

    if (p.popular) {
      s.addShape(pres.shapes.RECTANGLE, {
        x: x + 1.0, y: y - 0.2, w: 2.0, h: 0.35,
        fill: { color: C.gold }, line: { width: 0 }
      });
      s.addText("⭐ おすすめ", {
        x: x + 1.0, y: y - 0.2, w: 2.0, h: 0.35,
        fontSize: 11, bold: true, color: C.navy, align: "center", valign: "middle",
        fontFace: FONT_H, margin: 0
      });
    }

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 4, h: 4.9,
      fill: { color: C.white }, line: { color: p.color, width: p.popular ? 3 : 1 },
      shadow: makeShadow()
    });
    // Top color block
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 4, h: 1.0,
      fill: { color: p.color }, line: { width: 0 }
    });
    s.addText(p.name, {
      x: x, y: y + 0.2, w: 4, h: 0.6,
      fontSize: 22, bold: true, color: C.white, align: "center",
      fontFace: FONT_H, margin: 0
    });

    // Price
    s.addText(p.price, {
      x: x, y: y + 1.2, w: 4, h: 0.8,
      fontSize: 36, bold: true, color: p.color, align: "center",
      fontFace: FONT_H, margin: 0
    });
    s.addText("/ 月", {
      x: x, y: y + 1.95, w: 4, h: 0.3,
      fontSize: 11, color: C.sub, align: "center", fontFace: FONT_B, margin: 0
    });

    // Limit
    s.addText(p.limit, {
      x: x, y: y + 2.35, w: 4, h: 0.35,
      fontSize: 14, bold: true, color: C.text, align: "center",
      fontFace: FONT_H, margin: 0
    });
    s.addText(p.target, {
      x: x, y: y + 2.7, w: 4, h: 0.3,
      fontSize: 11, color: C.sub, align: "center", fontFace: FONT_B, margin: 0
    });

    // Divider
    s.addShape(pres.shapes.LINE, {
      x: x + 0.5, y: y + 3.15, w: 3, h: 0,
      line: { color: "E2E8F0", width: 1 }
    });

    // Features
    const featRuns = [];
    p.features.forEach((f, j) => {
      featRuns.push({ text: "✓ ", options: { color: p.color, bold: true } });
      featRuns.push({ text: f, options: { color: C.text, breakLine: j < p.features.length - 1 } });
    });
    s.addText(featRuns, {
      x: x + 0.4, y: y + 3.3, w: 3.2, h: 1.5,
      fontSize: 12, fontFace: FONT_B, paraSpaceAfter: 4, margin: 0
    });
  });

  // Bottom note
  s.addText("※ 登録者39名のキャリサポは「フリープラン」で十分。月額0円で全機能使えます。", {
    x: 0.5, y: 6.95, w: 12.3, h: 0.3,
    fontSize: 12, color: C.sub, align: "center", italic: true, fontFace: FONT_B
  });
}

// =====================================================
// SLIDE 9 — Setup Steps
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 9, "エルメ導入ステップ");
  addFooter(s);

  s.addText("30分で完了する導入手順", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.7,
    fontSize: 30, bold: true, color: C.navy, fontFace: FONT_H
  });

  const steps = [
    { num: "1", time: "5分", title: "エルメ無料登録", desc: "lme.jp にアクセス → メールアドレスで登録" },
    { num: "2", time: "5分", title: "公式LINEと連携", desc: "LINE Developersで認証ボタンをポチッと押すだけ" },
    { num: "3", time: "10分", title: "管理画面を触ってみる", desc: "壊れないので、まず触って慣れる(チュートリアル動画あり)" },
    { num: "4", time: "10分", title: "ウェルカム配信を作る", desc: "ステップ配信メニューから3通分割を設定" },
    { num: "5", time: "テスト送信", title: "自分で受信確認", desc: "本番有効化前に必ずテスト送信して動作チェック" },
  ];

  steps.forEach((st, i) => {
    const y = 1.9 + i * 1.0;

    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: 0.7, y: y, w: 0.85, h: 0.85,
      fill: { color: C.line }, line: { width: 0 }, shadow: makeShadow()
    });
    s.addText(st.num, {
      x: 0.7, y: y, w: 0.85, h: 0.85,
      fontSize: 30, bold: true, color: C.white, align: "center", valign: "middle",
      fontFace: FONT_H, margin: 0
    });

    // Connector line
    if (i < steps.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: 1.125, y: y + 0.85, w: 0, h: 0.15,
        line: { color: C.line, width: 2, dashType: "dash" }
      });
    }

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: 1.85, y: y, w: 10.95, h: 0.85,
      fill: { color: C.white }, line: { color: "E2E8F0", width: 0.5 }, shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 1.85, y: y, w: 0.1, h: 0.85,
      fill: { color: C.line }, line: { width: 0 }
    });

    // Time badge
    s.addShape(pres.shapes.RECTANGLE, {
      x: 2.1, y: y + 0.18, w: 1.1, h: 0.45,
      fill: { color: C.mint }, line: { width: 0 }
    });
    s.addText(st.time, {
      x: 2.1, y: y + 0.18, w: 1.1, h: 0.45,
      fontSize: 12, bold: true, color: C.navy, align: "center", valign: "middle",
      fontFace: FONT_H, margin: 0
    });

    // Title
    s.addText(st.title, {
      x: 3.4, y: y + 0.1, w: 4, h: 0.35,
      fontSize: 16, bold: true, color: C.navy, fontFace: FONT_H, margin: 0, valign: "middle"
    });
    // Desc
    s.addText(st.desc, {
      x: 3.4, y: y + 0.45, w: 9.3, h: 0.35,
      fontSize: 12, color: C.text, fontFace: FONT_B, margin: 0, valign: "middle"
    });
  });
}

// =====================================================
// SLIDE 10 — Full Build Flow
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 10, "Chapter 3 — 最強の構築フロー全体図");
  addFooter(s);

  s.addText("集客 → 登録 → 教育 → 成約までの完全フロー", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.6,
    fontSize: 28, bold: true, color: C.navy, fontFace: FONT_H
  });

  const flow = [
    { label: "①集客", color: C.teal, items: ["Threads発信", "DM送信"] },
    { label: "②登録", color: C.line, items: ["LINE友だち追加", "100万円特典"] },
    { label: "③教育", color: C.mint, items: ["ウェルカム3通", "7日間ステップ"] },
    { label: "④CV", color: C.gold, items: ["無料診断", "面談予約"] },
    { label: "⑤成約", color: C.navy, items: ["Zoom相談", "サービス契約"] },
  ];

  flow.forEach((f, i) => {
    const x = 0.5 + i * 2.56;

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 2.0, w: 2.36, h: 3.0,
      fill: { color: C.white }, line: { color: f.color, width: 2 }, shadow: makeShadow()
    });
    // Header bar
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: 2.0, w: 2.36, h: 0.7,
      fill: { color: f.color }, line: { width: 0 }
    });
    s.addText(f.label, {
      x: x, y: 2.0, w: 2.36, h: 0.7,
      fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle",
      fontFace: FONT_H, margin: 0
    });

    // Items
    const runs = [];
    f.items.forEach((it, j) => {
      runs.push({ text: "● ", options: { color: f.color, bold: true } });
      runs.push({ text: it, options: { color: C.text, breakLine: j < f.items.length - 1 } });
    });
    s.addText(runs, {
      x: x + 0.2, y: 2.95, w: 2.0, h: 1.9,
      fontSize: 12, fontFace: FONT_B, paraSpaceAfter: 6, valign: "top", margin: 0
    });

    // Arrow
    if (i < flow.length - 1) {
      s.addText("▶", {
        x: x + 2.36, y: 3.3, w: 0.2, h: 0.4,
        fontSize: 16, bold: true, color: C.gold, align: "center", valign: "middle",
        fontFace: FONT_H, margin: 0
      });
    }
  });

  // Bottom callout
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 5.4, w: 12.3, h: 1.6,
    fill: { color: C.navy }, line: { width: 0 }, shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 5.4, w: 0.12, h: 1.6,
    fill: { color: C.gold }, line: { width: 0 }
  });
  s.addText("🎯 各ステップでKPIを設定し、ボトルネックを特定して改善し続けるのが成功の鍵", {
    x: 0.85, y: 5.55, w: 11.8, h: 0.4,
    fontSize: 15, bold: true, color: C.gold, fontFace: FONT_H, margin: 0
  });
  s.addText([
    { text: "集客→登録 ", options: { color: C.mint, bold: true } },
    { text: "10〜15%   |   ", options: { color: C.white } },
    { text: "登録→面談 ", options: { color: C.mint, bold: true } },
    { text: "15〜20%   |   ", options: { color: C.white } },
    { text: "面談→成約 ", options: { color: C.mint, bold: true } },
    { text: "30〜40%", options: { color: C.white } },
  ], {
    x: 0.85, y: 6.05, w: 11.8, h: 0.4,
    fontSize: 13, fontFace: FONT_B, margin: 0
  });
  s.addText("これが上位プレイヤーの目安数値。下回るステップが改善ポイント。", {
    x: 0.85, y: 6.5, w: 11.8, h: 0.35,
    fontSize: 11, color: C.cream, italic: true, fontFace: FONT_B, margin: 0
  });
}

// =====================================================
// SLIDE 11 — Welcome Messages
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 11, "ウェルカム配信の設計");
  addFooter(s);

  s.addText("登録直後の3分割配信がCV率を3倍にする", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.6,
    fontSize: 28, bold: true, color: C.navy, fontFace: FONT_H
  });
  s.addText("情報を1通に詰めず、時間差で「読みやすく・行動しやすく」分割するのがポイント", {
    x: 0.5, y: 1.5, w: 12.3, h: 0.4,
    fontSize: 13, color: C.sub, fontFace: FONT_B
  });

  const messages = [
    {
      time: "登録直後",
      title: "歓迎+特典",
      role: "返報性のフック",
      content: "✓ 個人名で挨拶\n✓ 特典の即受け渡し\n✓ リッチメニュー誘導",
      color: C.line
    },
    {
      time: "30秒後",
      title: "サービス紹介",
      role: "教育(価値提示)",
      content: "✓ 4サービスの簡潔説明\n✓ 473万円実績\n✓ ベネフィット中心",
      color: C.mint
    },
    {
      time: "1分後",
      title: "1タップCTA",
      role: "コンバージョン",
      content: "✓ 「無料診断」ボタン\n✓ 1タップで完結\n✓ 押しすぎない一言",
      color: C.gold
    },
  ];

  messages.forEach((m, i) => {
    const x = 0.5 + i * 4.3;
    const y = 2.1;

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 4, h: 4.9,
      fill: { color: C.white }, line: { color: "E2E8F0", width: 0.5 }, shadow: makeShadow()
    });
    // Top
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 4, h: 1.1,
      fill: { color: m.color }, line: { width: 0 }
    });
    s.addText(m.time, {
      x: x, y: y + 0.15, w: 4, h: 0.4,
      fontSize: 13, bold: true, color: C.white, align: "center",
      fontFace: FONT_H, margin: 0
    });
    s.addText("MESSAGE " + (i + 1), {
      x: x, y: y + 0.55, w: 4, h: 0.5,
      fontSize: 22, bold: true, color: C.white, align: "center",
      fontFace: FONT_H, charSpacing: 2, margin: 0
    });

    // Title
    s.addText(m.title, {
      x: x, y: y + 1.3, w: 4, h: 0.5,
      fontSize: 20, bold: true, color: C.navy, align: "center",
      fontFace: FONT_H, margin: 0
    });
    s.addText(m.role, {
      x: x, y: y + 1.8, w: 4, h: 0.35,
      fontSize: 12, color: m.color, align: "center", italic: true,
      fontFace: FONT_B, margin: 0
    });

    // Divider
    s.addShape(pres.shapes.LINE, {
      x: x + 0.5, y: y + 2.3, w: 3, h: 0,
      line: { color: "E2E8F0", width: 1 }
    });

    // Content
    s.addText(m.content, {
      x: x + 0.4, y: y + 2.5, w: 3.2, h: 2.2,
      fontSize: 13, color: C.text, fontFace: FONT_B, paraSpaceAfter: 6, margin: 0
    });
  });
}

// =====================================================
// SLIDE 12 — 7-day Step
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 12, "7日間ステップ配信シナリオ");
  addFooter(s);

  s.addText("登録者を「ファン」に育てる7日間", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.6,
    fontSize: 28, bold: true, color: C.navy, fontFace: FONT_H
  });
  s.addText("放置されがちな登録者を、配信で温め続けて自然な相談申込に導く", {
    x: 0.5, y: 1.5, w: 12.3, h: 0.35,
    fontSize: 12, color: C.sub, fontFace: FONT_B
  });

  const days = [
    { day: "Day 0", title: "ウェルカム+特典", desc: "歓迎・100万円手順書を即渡す", color: C.line },
    { day: "Day 1", title: "事例ストーリー", desc: "「473万円受給した30代の話」", color: C.mint },
    { day: "Day 2", title: "退職サポート紹介", desc: "失業保険の落とし穴を解説", color: C.teal },
    { day: "Day 3", title: "学習サポート紹介", desc: "月3万もらいながら学ぶ仕組み", color: C.gold },
    { day: "Day 4", title: "転職サポート紹介", desc: "実績アドバイザー+30万お祝い金", color: C.line },
    { day: "Day 5", title: "AIリスキリング紹介", desc: "48万円が0円になる理由", color: C.mint },
    { day: "Day 6", title: "🎯 無料相談案内", desc: "最重要CTA・面談予約フォーム", color: C.red },
    { day: "Day 7", title: "FAQ+再CTA", desc: "よくある質問+ラスト後押し", color: C.navy },
  ];

  days.forEach((d, i) => {
    const x = 0.5 + (i % 4) * 3.18;
    const y = 2.0 + Math.floor(i / 4) * 2.4;

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 2.95, h: 2.2,
      fill: { color: C.white }, line: { color: "E2E8F0", width: 0.5 }, shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 2.95, h: 0.5,
      fill: { color: d.color }, line: { width: 0 }
    });
    s.addText(d.day, {
      x: x, y: y, w: 2.95, h: 0.5,
      fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle",
      fontFace: FONT_H, margin: 0
    });
    s.addText(d.title, {
      x: x + 0.15, y: y + 0.7, w: 2.65, h: 0.6,
      fontSize: 14, bold: true, color: C.navy, align: "center",
      fontFace: FONT_H, margin: 0
    });
    s.addText(d.desc, {
      x: x + 0.15, y: y + 1.3, w: 2.65, h: 0.85,
      fontSize: 11, color: C.text, align: "center",
      fontFace: FONT_B, margin: 0
    });
  });
}

// =====================================================
// SLIDE 13 — Rich Menu
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 13, "リッチメニュー設計の鉄則");
  addFooter(s);

  s.addText("リッチメニュー = LINEのトップページ", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.6,
    fontSize: 28, bold: true, color: C.navy, fontFace: FONT_H
  });
  s.addText("登録者が毎回見る場所。「6つの導線」で迷わせない設計を。", {
    x: 0.5, y: 1.5, w: 12.3, h: 0.35,
    fontSize: 12, color: C.sub, fontFace: FONT_B
  });

  // Left: 6-grid menu mockup
  const mx = 0.7, my = 2.1, mw = 5.2, mh = 4.6;
  s.addShape(pres.shapes.RECTANGLE, {
    x: mx - 0.1, y: my - 0.1, w: mw + 0.2, h: mh + 0.5,
    fill: { color: C.navy }, line: { width: 0 }
  });
  s.addText("理想のリッチメニュー(6枠)", {
    x: mx, y: my - 0.1, w: mw, h: 0.35,
    fontSize: 12, bold: true, color: C.gold, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });

  const menuItems = [
    { label: "キャリサポ\nとは", color: C.teal },
    { label: "100万円\n手順書", color: C.gold },
    { label: "無料診断", color: C.line },
    { label: "面談予約", color: C.mint },
    { label: "よくある\n質問", color: C.deepTeal },
    { label: "お客様の声", color: C.red },
  ];
  const cellW = mw / 3;
  const cellH = (mh - 0.3) / 2;
  menuItems.forEach((mi, i) => {
    const cx = mx + (i % 3) * cellW + 0.05;
    const cy = my + 0.3 + Math.floor(i / 3) * cellH + 0.05;
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cellW - 0.1, h: cellH - 0.1,
      fill: { color: mi.color }, line: { width: 0 }
    });
    s.addText(mi.label, {
      x: cx, y: cy, w: cellW - 0.1, h: cellH - 0.1,
      fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle",
      fontFace: FONT_H, margin: 0
    });
  });

  // Right: rules
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.4, y: 2.1, w: 6.4, h: 5.0,
    fill: { color: C.white }, line: { color: "E2E8F0", width: 0.5 }, shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.4, y: 2.1, w: 6.4, h: 0.55,
    fill: { color: C.mint }, line: { width: 0 }
  });
  s.addText("✅ 設計の5原則", {
    x: 6.4, y: 2.1, w: 6.4, h: 0.55,
    fontSize: 16, bold: true, color: C.navy, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });

  const rules = [
    { n: "1", t: "「面談」より「診断・相談」", d: "言葉を柔らかくして予約率UP" },
    { n: "2", t: "特典は常に1枠確保", d: "100万円手順書など登録動機を可視化" },
    { n: "3", t: "お客様の声を必ず入れる", d: "信頼性が一気に上がる" },
    { n: "4", t: "デザインに統一感", d: "ブランドカラーで揃える" },
    { n: "5", t: "月1回の更新", d: "新特典・季節キャンペーンで飽きさせない" },
  ];

  rules.forEach((r, i) => {
    const ry = 2.85 + i * 0.85;
    s.addShape(pres.shapes.OVAL, {
      x: 6.6, y: ry + 0.05, w: 0.45, h: 0.45,
      fill: { color: C.line }, line: { width: 0 }
    });
    s.addText(r.n, {
      x: 6.6, y: ry + 0.05, w: 0.45, h: 0.45,
      fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle",
      fontFace: FONT_H, margin: 0
    });
    s.addText(r.t, {
      x: 7.2, y: ry, w: 5.55, h: 0.4,
      fontSize: 14, bold: true, color: C.navy, fontFace: FONT_H, margin: 0, valign: "middle"
    });
    s.addText(r.d, {
      x: 7.2, y: ry + 0.4, w: 5.55, h: 0.4,
      fontSize: 11, color: C.sub, fontFace: FONT_B, margin: 0, valign: "middle"
    });
  });
}

// =====================================================
// SLIDE 14 — Tag branching
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 14, "タグ設計でシナリオ分岐");
  addFooter(s);

  s.addText("属性別に違うメッセージを送る = 反応率が劇的に変わる", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.6,
    fontSize: 26, bold: true, color: C.navy, fontFace: FONT_H
  });
  s.addText("登録者全員に同じ配信は卒業。タグで分けて、その人に刺さる内容を届ける", {
    x: 0.5, y: 1.5, w: 12.3, h: 0.35,
    fontSize: 12, color: C.sub, fontFace: FONT_B
  });

  // Top-level entry
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.4, y: 2.1, w: 2.5, h: 0.7,
    fill: { color: C.navy }, line: { width: 0 }, shadow: makeShadow()
  });
  s.addText("LINE登録", {
    x: 5.4, y: 2.1, w: 2.5, h: 0.7,
    fontSize: 16, bold: true, color: C.white, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });

  // Question
  s.addShape(pres.shapes.RECTANGLE, {
    x: 4.4, y: 3.05, w: 4.5, h: 0.6,
    fill: { color: C.gold }, line: { width: 0 }
  });
  s.addText("Q. 今のご状況を教えてください(タグ取得)", {
    x: 4.4, y: 3.05, w: 4.5, h: 0.6,
    fontSize: 12, bold: true, color: C.navy, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });

  // Branches
  const branches = [
    { tag: "会社員", color: C.teal, scenario: "退職検討中シナリオ\n→ ①退職サポート訴求" },
    { tag: "退職済み", color: C.line, scenario: "失業保険シナリオ\n→ ①+②フルコース" },
    { tag: "主婦", color: C.mint, scenario: "学習支援シナリオ\n→ ②+③訴求" },
    { tag: "転職希望", color: C.gold, scenario: "転職シナリオ\n→ ③+④訴求" },
  ];

  branches.forEach((b, i) => {
    const x = 0.5 + i * 3.2;
    const y = 4.2;

    // Connector
    s.addShape(pres.shapes.LINE, {
      x: 6.65, y: 3.65, w: x + 1.3 - 6.65, h: 0.55,
      line: { color: b.color, width: 1.5 }
    });

    // Tag box
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 2.7, h: 0.55,
      fill: { color: b.color }, line: { width: 0 }
    });
    s.addText("# " + b.tag, {
      x: x, y: y, w: 2.7, h: 0.55,
      fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle",
      fontFace: FONT_H, margin: 0
    });

    // Scenario
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y + 0.65, w: 2.7, h: 1.3,
      fill: { color: C.white }, line: { color: b.color, width: 1 }, shadow: makeShadow()
    });
    s.addText(b.scenario, {
      x: x + 0.1, y: y + 0.7, w: 2.5, h: 1.2,
      fontSize: 11, color: C.text, align: "center", valign: "middle",
      fontFace: FONT_B, margin: 0
    });
  });

  // Bottom impact
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 6.55, w: 12.3, h: 0.55,
    fill: { color: C.line }, line: { width: 0 }
  });
  s.addText("📊 タグ分岐ありの配信は、なし配信に比べて反応率が約2〜3倍に上がるデータあり", {
    x: 0.5, y: 6.55, w: 12.3, h: 0.55,
    fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });
}

// =====================================================
// SLIDE 15 — Tips
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 15, "Chapter 4 — 反応を上げる7つのコツ");
  addFooter(s);

  s.addText("実践者が知っている、成果を出す秘訣", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.6,
    fontSize: 28, bold: true, color: C.navy, fontFace: FONT_H
  });

  const tips = [
    { num: "01", t: "1メッセージ1テーマ", d: "詰め込まない。短く、読みやすく。" },
    { num: "02", t: "数字を冒頭に出す", d: "「473万円」など具体性で離脱防止" },
    { num: "03", t: "1タップCTA", d: "操作は最小限。3ステップは離脱の元" },
    { num: "04", t: "個人名で呼びかけ", d: "「〇〇様、こんにちは」で開封率UP" },
    { num: "05", t: "絵文字は控えめに", d: "多用はチープ感。1メッセに2〜3個" },
    { num: "06", t: "配信時間を最適化", d: "12時/19時/21時が反応率高い" },
    { num: "07", t: "テスト送信を必ず", d: "本番前に自分のLINEで確認" },
  ];

  tips.forEach((tip, i) => {
    const x = 0.5 + (i % 4) * 3.18;
    const y = 1.85 + Math.floor(i / 4) * 2.6;

    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 2.95, h: 2.4,
      fill: { color: C.white }, line: { color: "E2E8F0", width: 0.5 }, shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 0.1, h: 2.4,
      fill: { color: C.gold }, line: { width: 0 }
    });

    s.addText(tip.num, {
      x: x + 0.25, y: y + 0.25, w: 1.2, h: 0.6,
      fontSize: 28, bold: true, color: C.gold, fontFace: FONT_H, margin: 0
    });
    s.addText(tip.t, {
      x: x + 0.25, y: y + 0.95, w: 2.5, h: 0.5,
      fontSize: 15, bold: true, color: C.navy, fontFace: FONT_H, margin: 0
    });
    s.addText(tip.d, {
      x: x + 0.25, y: y + 1.5, w: 2.55, h: 0.85,
      fontSize: 11, color: C.text, fontFace: FONT_B, margin: 0
    });
  });
}

// =====================================================
// SLIDE 16 — KPIs
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 16, "追うべきKPI");
  addFooter(s);

  s.addText("数字を見て、改善し続ける", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.6,
    fontSize: 28, bold: true, color: C.navy, fontFace: FONT_H
  });
  s.addText("感覚ではなく数値で判断。ボトルネックの特定と改善が成果を最大化する。", {
    x: 0.5, y: 1.5, w: 12.3, h: 0.35,
    fontSize: 12, color: C.sub, fontFace: FONT_B
  });

  const kpis = [
    { stage: "STAGE 1", title: "登録率", target: "10〜15%", desc: "DM/投稿閲覧 → LINE登録", color: C.teal },
    { stage: "STAGE 2", title: "開封率", target: "70%以上", desc: "配信メッセージの開封割合", color: C.line },
    { stage: "STAGE 3", title: "クリック率", target: "20〜30%", desc: "CTAボタンのタップ率", color: C.mint },
    { stage: "STAGE 4", title: "面談予約率", target: "15〜20%", desc: "登録者→無料相談申込", color: C.gold },
    { stage: "STAGE 5", title: "成約率", target: "30〜40%", desc: "面談→契約", color: C.red },
  ];

  kpis.forEach((k, i) => {
    const x = 0.5 + i * 2.56;
    const y = 2.2;

    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 2.36, h: 4.4,
      fill: { color: C.white }, line: { color: "E2E8F0", width: 0.5 }, shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 2.36, h: 0.5,
      fill: { color: k.color }, line: { width: 0 }
    });
    s.addText(k.stage, {
      x: x, y: y, w: 2.36, h: 0.5,
      fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle",
      fontFace: FONT_H, charSpacing: 1.5, margin: 0
    });
    s.addText(k.title, {
      x: x, y: y + 0.7, w: 2.36, h: 0.5,
      fontSize: 18, bold: true, color: C.navy, align: "center",
      fontFace: FONT_H, margin: 0
    });
    s.addText(k.target, {
      x: x, y: y + 1.3, w: 2.36, h: 1.3,
      fontSize: 32, bold: true, color: k.color, align: "center", valign: "middle",
      fontFace: FONT_H, margin: 0
    });
    s.addShape(pres.shapes.LINE, {
      x: x + 0.4, y: y + 2.85, w: 1.55, h: 0,
      line: { color: "E2E8F0", width: 1 }
    });
    s.addText(k.desc, {
      x: x + 0.15, y: y + 3.0, w: 2.05, h: 1.3,
      fontSize: 11, color: C.text, align: "center",
      fontFace: FONT_B, margin: 0
    });
  });

  // Footer note
  s.addText("💡 1週間ごとに数字を確認し、最も低いステージを集中改善するのが最短ルート", {
    x: 0.5, y: 6.85, w: 12.3, h: 0.35,
    fontSize: 12, color: C.navy, italic: true, bold: true, align: "center", fontFace: FONT_B
  });
}

// =====================================================
// SLIDE 17 — NG Examples
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 17, "やってはいけないNG例");
  addFooter(s);

  s.addText("これをすると一発でブロックされる", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.6,
    fontSize: 28, bold: true, color: C.navy, fontFace: FONT_H
  });

  // 2 columns: NG vs OK
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.85, w: 6.1, h: 5.2,
    fill: { color: C.white }, line: { color: "FFD6D6", width: 1 }, shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.85, w: 6.1, h: 0.6,
    fill: { color: C.red }, line: { width: 0 }
  });
  s.addText("❌ NG例", {
    x: 0.5, y: 1.85, w: 6.1, h: 0.6,
    fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });

  const ngList = [
    "1日に何通も連投する(スパム判定)",
    "毎回同じ売り込み文面を送る",
    "相手の返信を無視して配信を続ける",
    "深夜・早朝に配信する",
    "「絶対稼げる」など断定表現を使う",
    "情報量が多すぎて読まれない",
    "リンクだらけで信頼性が下がる",
  ];

  const ngRuns = [];
  ngList.forEach((n, i) => {
    ngRuns.push({ text: "✗ ", options: { color: C.red, bold: true } });
    ngRuns.push({ text: n, options: { color: C.text, breakLine: i < ngList.length - 1 } });
  });
  s.addText(ngRuns, {
    x: 0.85, y: 2.65, w: 5.4, h: 4.2,
    fontSize: 13, fontFace: FONT_B, paraSpaceAfter: 8, margin: 0
  });

  // OK column
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.7, y: 1.85, w: 6.1, h: 5.2,
    fill: { color: C.white }, line: { color: "C8F0DC", width: 1 }, shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.7, y: 1.85, w: 6.1, h: 0.6,
    fill: { color: C.line }, line: { width: 0 }
  });
  s.addText("✅ OK例", {
    x: 6.7, y: 1.85, w: 6.1, h: 0.6,
    fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });

  const okList = [
    "1日1通以下、ステップ配信で計画的に",
    "テーマを変え、価値ある情報を発信",
    "返信があった人には個別丁寧対応",
    "12時/19時/21時のゴールデンタイム",
    "「〇〇な可能性があります」と謙虚に",
    "1メッセ1テーマ、簡潔に",
    "リンクは1〜2個、必要なものだけ",
  ];
  const okRuns = [];
  okList.forEach((o, i) => {
    okRuns.push({ text: "✓ ", options: { color: C.line, bold: true } });
    okRuns.push({ text: o, options: { color: C.text, breakLine: i < okList.length - 1 } });
  });
  s.addText(okRuns, {
    x: 7.05, y: 2.65, w: 5.4, h: 4.2,
    fontSize: 13, fontFace: FONT_B, paraSpaceAfter: 8, margin: 0
  });
}

// =====================================================
// SLIDE 18 — Roadmap
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 18, "30日実践ロードマップ");
  addFooter(s);

  s.addText("導入から運用安定までの30日プラン", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.6,
    fontSize: 28, bold: true, color: C.navy, fontFace: FONT_H
  });

  const weeks = [
    {
      week: "Week 1",
      title: "セットアップ",
      color: C.line,
      items: [
        "エルメ無料登録+公式LINE連携",
        "管理画面の使い方を理解",
        "ウェルカム3分割配信を作成",
        "テスト送信で動作確認"
      ]
    },
    {
      week: "Week 2",
      title: "ステップ配信構築",
      color: C.mint,
      items: [
        "7日間ステップ配信を設計",
        "各日のメッセージを執筆",
        "事例・FAQを盛り込む",
        "本番配信スタート"
      ]
    },
    {
      week: "Week 3",
      title: "リッチメニュー改善",
      color: C.gold,
      items: [
        "リッチメニューを6枠に再設計",
        "面談予約フォーム作成",
        "「お客様の声」セクション追加",
        "LP→LINE導線を強化"
      ]
    },
    {
      week: "Week 4",
      title: "改善と最適化",
      color: C.teal,
      items: [
        "KPIを計測+ボトルネック特定",
        "タグ分岐シナリオを追加",
        "反応の良い配信を分析",
        "次月の改善計画を立てる"
      ]
    },
  ];

  weeks.forEach((w, i) => {
    const x = 0.5 + i * 3.18;
    const y = 1.85;

    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 2.95, h: 5.0,
      fill: { color: C.white }, line: { color: "E2E8F0", width: 0.5 }, shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 2.95, h: 1.0,
      fill: { color: w.color }, line: { width: 0 }
    });
    s.addText(w.week, {
      x: x, y: y + 0.15, w: 2.95, h: 0.4,
      fontSize: 13, bold: true, color: C.white, align: "center",
      fontFace: FONT_H, charSpacing: 1.5, margin: 0
    });
    s.addText(w.title, {
      x: x, y: y + 0.5, w: 2.95, h: 0.5,
      fontSize: 18, bold: true, color: C.white, align: "center",
      fontFace: FONT_H, margin: 0
    });

    const itemRuns = [];
    w.items.forEach((it, j) => {
      itemRuns.push({ text: "▸ ", options: { color: w.color, bold: true } });
      itemRuns.push({ text: it, options: { color: C.text, breakLine: j < w.items.length - 1 } });
    });
    s.addText(itemRuns, {
      x: x + 0.25, y: y + 1.2, w: 2.55, h: 3.7,
      fontSize: 11, fontFace: FONT_B, paraSpaceAfter: 8, valign: "top", margin: 0
    });
  });
}

// =====================================================
// SLIDE 19 — Summary
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.cream);
  addHeader(s, 19, "まとめ");
  addFooter(s);

  s.addText("最強の公式LINE = 6つの要素の掛け算", {
    x: 0.5, y: 0.85, w: 12.3, h: 0.6,
    fontSize: 30, bold: true, color: C.navy, fontFace: FONT_H
  });

  const summary = [
    { num: "1", t: "エルメ導入で時間差・ステップ配信を可能に", color: C.line },
    { num: "2", t: "ウェルカム3分割で読みやすく行動しやすく", color: C.mint },
    { num: "3", t: "7日間ステップ配信で関係構築", color: C.teal },
    { num: "4", t: "リッチメニュー6枠で迷わせない", color: C.gold },
    { num: "5", t: "タグ分岐で個別最適化", color: C.deepTeal },
    { num: "6", t: "KPI計測で改善し続ける", color: C.red },
  ];

  summary.forEach((it, i) => {
    const y = 1.85 + i * 0.85;

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 12.3, h: 0.7,
      fill: { color: C.white }, line: { color: "E2E8F0", width: 0.5 }, shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 0.12, h: 0.7,
      fill: { color: it.color }, line: { width: 0 }
    });
    s.addShape(pres.shapes.OVAL, {
      x: 0.85, y: y + 0.13, w: 0.45, h: 0.45,
      fill: { color: it.color }, line: { width: 0 }
    });
    s.addText(it.num, {
      x: 0.85, y: y + 0.13, w: 0.45, h: 0.45,
      fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle",
      fontFace: FONT_H, margin: 0
    });
    s.addText(it.t, {
      x: 1.5, y: y, w: 11.2, h: 0.7,
      fontSize: 16, bold: true, color: C.navy, valign: "middle",
      fontFace: FONT_H, margin: 0
    });
  });
}

// =====================================================
// SLIDE 20 — Final CTA / Closing
// =====================================================
{
  const s = pres.addSlide();
  addBg(s, C.navy);

  // Decorative
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 13.3, h: 0.15,
    fill: { color: C.gold }, line: { width: 0 }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 7.35, w: 13.3, h: 0.15,
    fill: { color: C.mint }, line: { width: 0 }
  });

  // Eyebrow
  s.addText("READY TO START?", {
    x: 0.5, y: 1.2, w: 12.3, h: 0.4,
    fontSize: 14, color: C.gold, fontFace: FONT_H, charSpacing: 6, bold: true,
    align: "center"
  });

  // Title
  s.addText("今日、最初の30分を", {
    x: 0.5, y: 1.7, w: 12.3, h: 0.9,
    fontSize: 44, bold: true, color: C.white, align: "center", fontFace: FONT_H
  });
  s.addText("エルメに投資してください", {
    x: 0.5, y: 2.6, w: 12.3, h: 0.9,
    fontSize: 44, bold: true, color: C.mint, align: "center", fontFace: FONT_H
  });

  // Divider
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.05, y: 3.75, w: 1.2, h: 0.06,
    fill: { color: C.gold }, line: { width: 0 }
  });

  s.addText("無料で始められて、登録者を確実にファン化できる、唯一の方法です。", {
    x: 0.5, y: 3.95, w: 12.3, h: 0.5,
    fontSize: 17, color: C.cream, align: "center", fontFace: FONT_B
  });

  // CTA buttons
  s.addShape(pres.shapes.RECTANGLE, {
    x: 2.0, y: 4.85, w: 4.3, h: 1.0,
    fill: { color: C.line }, line: { width: 0 }, shadow: makeShadow()
  });
  s.addText("👉 エルメ公式サイト", {
    x: 2.0, y: 4.85, w: 4.3, h: 0.5,
    fontSize: 16, bold: true, color: C.white, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });
  s.addText("https://lme.jp/", {
    x: 2.0, y: 5.3, w: 4.3, h: 0.5,
    fontSize: 12, color: C.white, align: "center", valign: "middle",
    fontFace: FONT_B, margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.0, y: 4.85, w: 4.3, h: 1.0,
    fill: { color: C.gold }, line: { width: 0 }, shadow: makeShadow()
  });
  s.addText("📘 キャリサポ公式LINE", {
    x: 7.0, y: 4.85, w: 4.3, h: 0.5,
    fontSize: 16, bold: true, color: C.navy, align: "center", valign: "middle",
    fontFace: FONT_H, margin: 0
  });
  s.addText("https://lin.ee/tpWHOuh", {
    x: 7.0, y: 5.3, w: 4.3, h: 0.5,
    fontSize: 12, color: C.navy, align: "center", valign: "middle",
    fontFace: FONT_B, margin: 0
  });

  // Bottom bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 6.5, w: 12.3, h: 0.6,
    fill: { color: C.deepTeal }, line: { width: 0 }
  });
  s.addText("最強の公式LINE設計ガイド | © キャリサポ", {
    x: 0.5, y: 6.5, w: 12.3, h: 0.6,
    fontSize: 12, color: C.gold, align: "center", valign: "middle",
    fontFace: FONT_H, charSpacing: 2, margin: 0
  });
}

// ============ Write file ============
const outPath = "C:\\Users\\ogaig\\OneDrive\\デスクトップ\\キャリサポ\\docs\\最強の公式LINE設計ガイド.pptx";
pres.writeFile({ fileName: outPath })
  .then(fileName => console.log("✅ Saved:", fileName))
  .catch(err => { console.error("❌ Error:", err); process.exit(1); });
