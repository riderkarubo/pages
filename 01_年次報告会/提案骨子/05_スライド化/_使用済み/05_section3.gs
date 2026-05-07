/**
 * MCCM提案スライド化 - §3: 2026年度KPI目標設定（6スライド）
 *
 * マスター情報源: https://riderkarubo.github.io/pages/mccm-proposal/
 * （これ以外の場所からは情報を拾わない）
 *
 * 構成:
 *  S7:  2026年度KPI（協賛配信受注 年間48件・2,400万円）
 *  S8:  売上構造の分解（KPIツリー）
 *  S9:  マイルストーン①：メディアパワー（4指標）
 *  S10: 競合ベンチマーク（コスメカテゴリ・60分換算）
 *  S11: メディアパワーの因数分解（3カード）
 *  S12: マイルストーン②：リピーター数値（3指標・追加提案）
 *
 * 実行手順:
 *  1. 02_helpers.gs / 04_section2.gs を貼り付け済みであること
 *  2. このコードを Apps Script に貼り付け
 *  3. insertSection3() を実行 → 末尾に6枚追加される
 */

function insertSection3() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  Logger.log("§3 開始 - 現在: " + pres.getSlides().length + "枚");

  insertS7_kpiHeadline(pres);   Logger.log("S7 完了");
  insertS8_revenueTree(pres);   Logger.log("S8 完了");
  insertS9_milestone1(pres);    Logger.log("S9 完了");
  insertS10_benchmark(pres);    Logger.log("S10 完了");
  insertS11_factoring(pres);    Logger.log("S11 完了");
  insertS12_milestone2(pres);   Logger.log("S12 完了");

  Logger.log("§3 完了 - 現在: " + pres.getSlides().length + "枚");
}

// ============================================================
// S7: 2026年度KPI（協賛配信受注）
// ============================================================
function insertS7_kpiHeadline(pres) {
  var s = newSlide(pres, "2026年度 主要KPI（暫定目標）");

  // メイン訴求カード（中央大型）
  rect(s, 60, 100, 600, 140, "#0F3460");
  txt(s, "2026年度 主要KPI", 60, 110, 600, 18, {
    size: 11, color: "#FFD166", bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "協賛配信 受注件数", 60, 134, 600, 22, {
    size: 13, color: C_WHITE,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "月平均4件 / 年間48件", 60, 158, 600, 36, {
    size: 26, color: C_WHITE, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "売上換算：2,400万円（1件あたり50万円換算）", 60, 200, 600, 24, {
    size: 14, color: C_WHITE, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 進捗カード（横2分割）
  rect(s, 60, 260, 290, 80, C_WHITE, C_BORDER, 1);
  txt(s, "2026年度進捗（※4/30更新予定）", 70, 268, 270, 14, {
    size: 9, color: C_GRAY, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "11 件", 70, 286, 270, 38, {
    size: 26, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "※5月以降配信含む", 70, 322, 270, 12, {
    size: 8, color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  rect(s, 370, 260, 290, 80, C_WHITE, C_BORDER, 1);
  txt(s, "2026年度進捗 売上（※先方共有待ち）", 380, 268, 270, 14, {
    size: 9, color: C_GRAY, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "◯ 万円", 380, 286, 270, 38, {
    size: 26, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // フッター注釈
  txt(s, "※本目標は現時点の暫定値です。本資料の競合ベンチマーク分析を踏まえ、最適なKPI体系を改めてご提案します。",
    25, 360, 671, 18, {
      size: 9, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S8: 売上構造の分解（KPIツリー）
// ============================================================
function insertS8_revenueTree(pres) {
  var s = newSlide(pres, "協賛配信の売上構造（KPIツリー）");

  // L0: 売上（最上段中央）
  rectTxt(s, "売上", 305, 90, 110, 32, "#0F3460", {
    size: 14, color: C_WHITE, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // L0→L1 接続線
  rect(s, 359, 122, 2, 16, "#0F3460");  // 縦
  rect(s, 200, 138, 320, 2, "#0F3460");  // 横
  rect(s, 200, 138, 2, 14, "#0F3460");  // 左下
  rect(s, 518, 138, 2, 14, "#0F3460");  // 右下

  // ×記号
  rectTxt(s, "×", 350, 132, 20, 14, "#E5E7EB", {
    size: 9, color: "#555", bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // L1左: 受注件数
  rect(s, 130, 152, 140, 36, C_WHITE, "#0F3460", 2);
  txt(s, "受注件数", 130, 152, 140, 36, {
    size: 12, color: "#0F3460", bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // L1右: 受注単価
  rect(s, 450, 152, 140, 36, C_WHITE, "#0F3460", 2);
  txt(s, "受注単価", 450, 152, 140, 36, {
    size: 12, color: "#0F3460", bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // L1→L2 接続線
  rect(s, 199, 188, 2, 12, "#999999");  // 受注件数下
  rect(s, 519, 188, 2, 12, "#0F3460");  // 受注単価下

  // L2左: 商談件数 × 受注率
  rect(s, 75, 220, 2, 12, "#999999");  // 縦
  rect(s, 235, 220, 2, 12, "#999999");
  rect(s, 75, 200, 162, 2, "#999999");
  rectTxt(s, "×", 145, 195, 16, 12, "#E5E7EB", {
    size: 8, color: "#555", bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 商談件数（緑）
  rect(s, 35, 232, 130, 56, "#F0FAF8", "#1B998B", 1);
  txt(s, "KPI", 35, 234, 130, 12, {
    size: 8, color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.START
  });
  txt(s, "商談件数", 35, 246, 130, 16, {
    size: 11, color: "#1B998B", bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "貴社セールス設計", 35, 264, 130, 20, {
    size: 8, color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 受注率（赤）
  rect(s, 175, 232, 130, 56, "#FFF5F7", "#E94560", 1);
  txt(s, "KPI", 175, 234, 130, 12, {
    size: 8, color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.START
  });
  txt(s, "受注率", 175, 250, 130, 18, {
    size: 11, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 受注率→Fireworkサポートボックス
  rect(s, 175, 296, 130, 60, "#FFF0F3", C_ACCENT, 1);
  txt(s, "▲ コンテンツ品質向上\n▲ メディアパワー向上", 175, 298, 130, 32, {
    size: 9, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "Firework がサポート", 175, 332, 130, 18, {
    size: 8, color: C_ACCENT,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // L2右: 受注単価ロジック
  rect(s, 415, 200, 210, 60, C_WHITE, C_BORDER, 1);
  txt(s, "エンゲージメントが高く\n「ここに広告出稿したい」と\n思わせるメディアであれば\n単価を引き上げられる",
    420, 200, 200, 60, {
      size: 9, color: "#555",
      align: SlidesApp.ParagraphAlignment.CENTER,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  rect(s, 415, 270, 210, 60, "#FFF0F3", C_ACCENT, 1);
  txt(s, "▲ コンテンツ品質向上\n▲ メディアパワー向上", 415, 273, 210, 32, {
    size: 9, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "Firework がサポート", 415, 305, 210, 18, {
    size: 8, color: C_ACCENT,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 注釈
  txt(s, "💡 商談件数の拡大は貴社で設計いただき、Fireworkは受注率と受注単価を押し上げるためのコンテンツ・メディア強化に集中します。",
    25, 372, 671, 16, {
      size: 9, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S9: マイルストーン①：メディアパワー（4指標）
// ============================================================
function insertS9_milestone1(pres) {
  var s = newSlide(pres, "マイルストーン①：メディアパワー（全視聴者ベース）");

  // 4カード（横2×縦2）
  var cardW = 320, cardH = 130, gap = 16;
  var startX = (SW - 2 * cardW - gap) / 2;  // 32
  var startY = 96;

  drawMilestoneCard(s, startX, startY, cardW, cardH, {
    no: "①", border: "#0F3460",
    name: "ライブ視聴者数（リーチ指標）",
    nowLabel: "直近3ヶ月平均（2026.1-3）", nowValue: "5,084", nowUnit: "人",
    targetValue: "10,000", targetUnit: "人",
    gap: "+4,916人",
    note: "業界上位の高水準。5,000人超を安定維持しつつ、最大10,000人を目指す。"
  });

  drawMilestoneCard(s, startX + cardW + gap, startY, cardW, cardH, {
    no: "②", border: C_ACCENT,
    name: "平均視聴分数（リテンション指標）",
    nowLabel: "直近3ヶ月平均", nowValue: "3.7", nowUnit: "分",
    targetValue: "7", targetUnit: "分",
    gap: "+3.3分",
    note: "リピーター中心設計でエンゲージメント改善余地が大きい。"
  });

  drawMilestoneCard(s, startX, startY + cardH + gap, cardW, cardH, {
    no: "③", border: "#E07A5F",
    name: "商品クリック率 / CTR（中央値ベース）",
    nowLabel: "2025年平均", nowValue: "1.13", nowUnit: "%",
    targetValue: "2.2", targetUnit: "%",
    gap: "+1.07pt",
    note: "クーポン・限定訴求・CTR誘導設計で改善余地あり。"
  });

  drawMilestoneCard(s, startX + cardW + gap, startY + cardH + gap, cardW, cardH, {
    no: "④", border: C_ACCENT2,
    name: "コメント参加率（中央値ベース）",
    nowLabel: "2025年平均", nowValue: "0.68", nowUnit: "%",
    targetValue: "1.0", targetUnit: "%",
    gap: "+0.32pt",
    note: "参加型コンテンツ強化で改善余地あり。"
  });
}

/** マイルストーンカード（現状→目標の対比） */
function drawMilestoneCard(slide, x, y, w, h, p) {
  rect(slide, x, y, w, h, C_WHITE, C_BORDER, 1);
  rect(slide, x, y, 4, h, p.border);  // 左ボーダー

  // 番号
  txt(slide, p.no + " " + p.name, x + 14, y + 6, w - 28, 18, {
    size: 11, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 現状（左）
  txt(slide, p.nowLabel, x + 14, y + 32, (w - 28) / 2 - 8, 12, {
    size: 8, color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(slide, p.nowValue, x + 14, y + 46, 80, 30, {
    size: 22, color: C_TEXT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(slide, p.nowUnit, x + 96, y + 56, 30, 20, {
    size: 11, color: C_TEXT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.BOTTOM
  });

  // 区切り線
  rect(slide, x + (w / 2), y + 32, 1, 50, C_BORDER);

  // 目標（右）
  txt(slide, "2026年度目標", x + (w / 2) + 10, y + 32, (w - 28) / 2 - 8, 12, {
    size: 8, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(slide, p.targetValue, x + (w / 2) + 10, y + 46, 80, 30, {
    size: 22, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(slide, p.targetUnit, x + (w / 2) + 92, y + 56, 30, 20, {
    size: 11, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.BOTTOM
  });

  // ギャップ＋注釈
  txt(slide, "ギャップ：" + p.gap, x + 14, y + 86, w - 28, 14, {
    size: 9, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(slide, p.note, x + 14, y + 102, w - 28, 24, {
    size: 8, color: "#555",
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
}

// ============================================================
// S10: 競合ベンチマーク（コスメカテゴリ）
// ============================================================
function insertS10_benchmark(pres) {
  var s = newSlide(pres, "競合ベンチマーク（コスメカテゴリ・60分換算）");

  var colDefs = [
    { label: "ブランド", align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "視聴者数", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "平均視聴分数", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "商品CTR", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "コメント率", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "リピーター比率", align: SlidesApp.ParagraphAlignment.CENTER }
  ];
  var rows = [
    ["マツキヨココカラライブ\nコスメ限定（2026.1-3）", "5,228人", "3.40分", "1.24%", "0.94%", "5.0%"],
    ["A社（リテール）",      "2,494人", "24.3分", "29.4%", "33.4%", "50.5%"],
    ["B社（コスメブランド）", "6,958人", "11.8分", "8.2%",  "5.1%",  "17.3%"],
    ["C社(コスメブランド)",  "1,458人", "28.2分", "11.7%", "9.1%",  "43.1%"],
    ["D社（メーカー）",      "871人",   "9.9分",  "5.3%",  "4.4%",  "16.5%"]
  ];

  makeTable(s, 25, 96, 671, colDefs, rows, {
    rowH: 32,
    hdrH: 28,
    headerSize: 9,
    bodySize: 9,
    highlightRows: [0]  // マツキヨ行ハイライト
  });

  // 結論ボックス
  rect(s, 25, 290, 671, 80, "#FFF5F5", "#C9302C", 1);
  txt(s, "コスメに絞っても競合コスメブランドにビハインド", 35, 296, 651, 22, {
    size: 13, color: "#C9302C", bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "視聴者数は業界2位ポジションを維持。一方で視聴分数・CTR・コメント率はB社の約1/3〜1/9、C社の約1/8〜1/10。", 35, 320, 651, 18, {
    size: 10, color: C_TEXT,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "→ 視聴者数（リーチ）の強みは活きているが、見続けさせる・行動させる・参加させる仕掛けに改善余地が大きい。", 35, 342, 651, 18, {
    size: 10, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
}

// ============================================================
// S11: メディアパワーの因数分解
// ============================================================
function insertS11_factoring(pres) {
  var s = newSlide(pres, "メディアパワーの因数分解 ― リピーターが最大レバー");

  // 数式バー（中央上部）
  rect(s, 25, 90, 671, 70, "#F8FAFC", "#D0D8E8", 1);
  txt(s, "メディアパワーの因数分解", 25, 96, 671, 14, {
    size: 9, color: "#0F3460", bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "メディアパワー = 集める力 × 引き留める力 × 再訪する力", 25, 114, 671, 28, {
    size: 18, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "（= 視聴者数 × エンゲージメント深度［視聴分数・CTR・コメント率］ × リピーター比率）",
    25, 142, 671, 14, {
      size: 9, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.CENTER,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 3カード（横並び）
  var cardW = 215, cardH = 140, gap = 13;
  var startX = (SW - 3 * cardW - 2 * gap) / 2;
  var startY = 180;

  drawFactorCard(s, startX, startY, cardW, cardH, {
    border: "#0F3460",
    label: "① 集める力",
    headline: "業界上位・頭打ち",
    detail: "視聴者数 平均5,000人超を安定維持。競合と比較しても高水準で、ここからの倍増は困難。"
  });

  drawFactorCard(s, startX + (cardW + gap), startY, cardW, cardH, {
    border: "#C0392B",
    label: "② 引き留める力",
    headline: "競合と大差・伸び代大",
    detail: "視聴分数・CTR・コメント率すべてで競合トップの約2〜11%水準。ここを埋めれば集客力を増やさずにメディア価値を数倍化できる。"
  });

  drawFactorCard(s, startX + 2 * (cardW + gap), startY, cardW, cardH, {
    border: C_ACCENT2,
    label: "③ 再訪する力",
    headline: "②を底上げするレバー",
    detail: "リピーターはCTR・コメント率・視聴分数が非リピーターの数倍。再訪する視聴者を増やすことが、②の平均値を引き上げる最短経路。"
  });

  // 結論ボックス
  rect(s, 25, 335, 671, 50, "#FFF8E1", "#F0C040", 1);
  txt(s, "集める力は業界上位。次は引き留める力と再訪する力を作る年。", 35, 340, 651, 22, {
    size: 12, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "リピーターを増やすことが、メディアパワー全体を底上げし、協賛メーカーに選ばれ続けるための最短経路。",
    35, 360, 651, 18, {
      size: 10, color: "#C07000",
      align: SlidesApp.ParagraphAlignment.CENTER,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
}

/** 因数分解カード */
function drawFactorCard(slide, x, y, w, h, p) {
  rect(slide, x, y, w, h, C_WHITE, C_BORDER, 1);
  rect(slide, x, y, 4, h, p.border);

  txt(slide, p.label, x + 14, y + 8, w - 28, 14, {
    size: 9, color: p.border, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(slide, p.headline, x + 14, y + 28, w - 28, 22, {
    size: 13, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(slide, p.detail, x + 14, y + 56, w - 28, h - 64, {
    size: 9, color: "#555",
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.START
  });
}

// ============================================================
// S12: マイルストーン②：リピーター数値（追加提案）
// ============================================================
function insertS12_milestone2(pres) {
  var s = newSlide(pres, "マイルストーン②：リピーター数値（追加提案）");

  // ヘッダーボックス（提案趣旨）
  rect(s, 25, 90, 671, 56, "#F0F4FF", "#0F3460", 2);
  txt(s, "マイルストーン①（メディアパワー）を底上げするためにリピーター3指標を追加提案",
    35, 96, 651, 20, {
      size: 12, color: "#0F3460", bold: true,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "リピーターはCTRで全視聴者の約3.7倍、コメント率で約5.3倍のエンゲージメントを持つ高価値層。",
    35, 118, 651, 22, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 3カード（横並び）
  var cardW = 215, cardH = 200, gap = 13;
  var startX = (SW - 3 * cardW - 2 * gap) / 2;
  var startY = 160;

  drawRepeaterCard(s, startX, startY, cardW, cardH, {
    label: "マイルストーン② A",
    name: "リピーター比率",
    border: C_ACCENT,
    nowVal: "4.88", nowUnit: "%",
    targetVal: "8.0", targetUnit: "%",
    sub: "視聴者5,000人換算：244人→400人",
    note: "自然成長トレンド（+約2pt/年）に出演者SNSフォロワー増による「推し再訪」施策を加えて目標設定。"
  });

  drawRepeaterCard(s, startX + (cardW + gap), startY, cardW, cardH, {
    label: "マイルストーン② B",
    name: "リピーターCTR",
    border: "#E07A5F",
    nowVal: "3.87", nowUnit: "%",
    targetVal: "8.0", targetUnit: "%",
    sub: "全視聴者CTRの約3.7倍の高価値層",
    note: "リピーター向け限定特典の導入＋出演者ファン育成で「この人がすすめるなら買いたい」層を増やし目標設定。"
  });

  drawRepeaterCard(s, startX + 2 * (cardW + gap), startY, cardW, cardH, {
    label: "マイルストーン② C",
    name: "リピーターコメント率",
    border: C_ACCENT2,
    nowVal: "3.78", nowUnit: "%",
    targetVal: "7.5", targetUnit: "%",
    sub: "全視聴者コメント率の約5.3倍の高エンゲージ層",
    note: "出演者がInstagram等で視聴者に話しかける配信外コミュニティ形成を新施策として目標設定。"
  });

  // フッター結論
  rect(s, 25, 372, 671, 22, C_LIGHT);
  txt(s, "📌 マイルストーン② 達成 → マイルストーン① 底上げ → KPI（協賛配信受注）達成 という最短経路",
    35, 372, 651, 22, {
      size: 10, color: C_ACCENT2, bold: true,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
}

/** リピーターカード（現状→目標） */
function drawRepeaterCard(slide, x, y, w, h, p) {
  rect(slide, x, y, w, h, C_WHITE, C_BORDER, 1);
  rect(slide, x, y, w, 4, p.border);  // 上ボーダー

  txt(slide, p.label, x + 12, y + 10, w - 24, 12, {
    size: 8, color: C_GRAY, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(slide, p.name, x + 12, y + 26, w - 24, 18, {
    size: 12, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 現状
  txt(slide, "現状", x + 12, y + 50, w - 24, 12, {
    size: 8, color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(slide, p.nowVal + p.nowUnit, x + 12, y + 64, w - 24, 26, {
    size: 18, color: "#555", bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 目標
  txt(slide, "2026年度目標", x + 12, y + 96, w - 24, 12, {
    size: 8, color: p.border, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(slide, p.targetVal + p.targetUnit, x + 12, y + 110, w - 24, 32, {
    size: 24, color: p.border, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // サブ
  if (p.sub) {
    txt(slide, p.sub, x + 12, y + 144, w - 24, 16, {
      size: 8, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.CENTER,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  }

  // 注釈
  txt(slide, p.note, x + 12, y + 162, w - 24, h - 170, {
    size: 8, color: "#555",
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.START
  });
}
