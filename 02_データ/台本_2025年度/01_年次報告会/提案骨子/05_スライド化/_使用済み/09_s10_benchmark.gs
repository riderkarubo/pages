/**
 * MCCM提案スライド化 - S10 単独生成コマンド
 *
 * マスター情報源「コスメカテゴリに絞った場合の比較
 * （マツキヨココカラライブ 2025年度 × 競合4社）」テーブルを完全踏襲。
 *
 * ★既存の S10 (insertS10v2_benchmark) は使わず、こちらが正本。
 *
 * データ:
 *  - マツキヨ コスメ限定（2026.1-3）  ← クリーム強調
 *  - マツキヨ 全体（2026.1-3）
 *  - A社（リテール）/ B社・C社（コスメブランド）/ D社（メーカー）
 *
 * 列構成（7列）: ブランド / 配信数 / 平均配信時間 / 視聴者数 / 視聴分数 / CTR / コメント率
 *
 * 実行手順:
 *  1. 02_helpers.gs を最新版で貼り付け済み
 *  2. このファイルを Apps Script に追加
 *  3. runS10Benchmark() を選んで実行 → 末尾に1枚追加
 */

function runS10Benchmark() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  Logger.log("S10 開始 - 現在: " + pres.getSlides().length + "枚");
  insertS10_cosmeBenchmark(pres);
  Logger.log("S10 完了 - 現在: " + pres.getSlides().length + "枚");
}

function insertS10_cosmeBenchmark(pres) {
  var s = newSlide(pres, "コスメカテゴリ限定比較（マツキヨ × 競合4社・2026.1-3）");

  // 結論メッセージ（強調キーワード「業界2位」）
  drawConclusionMessage(s, 88,
    "コスメ限定でも視聴者数は業界2位の高水準。エンゲージメント指標で競合と大差・伸び代大",
    "業界2位の高水準");

  // 7列テーブル（マスターHTML完全踏襲）
  var colDefs = [
    { label: "ブランド / セグメント",     align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "配信数",                    align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "平均\n配信時間",             align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "視聴者数\n(60分換算)",      align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "視聴分数\n(60分換算)",      align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "商品CTR\n(60分換算)",       align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "コメント率\n(60分換算)",    align: SlidesApp.ParagraphAlignment.CENTER }
  ];
  var rows = [
    ["マツキヨココカラライブ\nコスメ限定（2026.1-3）", "7回",  "61.8分", "5,228人",  "3.40分",  "1.24%",  "0.94%"],
    ["マツキヨココカラライブ\n全体（2026.1-3）",       "10回", "64.2分", "5,084人",  "3.1分",   "1.0%",   "0.8%"],
    ["A社(リテール)",                                 "13回", "74.8分", "2,494人",  "24.3分",  "29.4%",  "33.4%"],
    ["B社(コスメブランド)",                           "23回", "81.1分", "6,958人",  "11.8分",  "8.2%",   "5.1%"],
    ["C社(コスメブランド)",                           "4回",  "63.0分", "1,458人",  "28.2分",  "11.7%",  "9.1%"],
    ["D社(メーカー)",                                 "3回",  "51.5分", "871人",    "9.9分",   "5.3%",   "4.4%"]
  ];

  makeColHighlightTable(s, 25, 128, 671, colDefs, rows, {
    rowH: 28,
    hdrH: 36,
    bodySize: 9,
    headerSize: 9,
    highlightRow: 0  // マツキヨ コスメ限定 行をクリーム強調
  });

  // 注釈
  txt(s, "※ コスメ＝メイク＋スキンケア(UV含む)。ヘアケア・日用品等は除外。全ブランド2026.1-3期間で統一。社名はA〜D表記。",
    25, 322, 671, 16, {
      size: 9, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 結論注釈
  txt(s, "→ カテゴリ補正しても エンゲージメント系（視聴分数・CTR・コメント率）のギャップ構造は変わらない",
    25, 348, 671, 22, {
      size: 11, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
}
