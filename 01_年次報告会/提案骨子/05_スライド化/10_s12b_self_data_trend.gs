/**
 * MCCM提案スライド化 - S12b: 自社データ(コメント参加率・リピーター比率)
 *
 * 最終更新: 2026-04-28 01:38
 * マスター情報源: https://riderkarubo.github.io/pages/mccm-proposal/
 *  「③ 自社データで確認」テーブルから 2行(コメント参加率・リピーター比率) を抽出
 *
 * 構成:
 *  - タイトル黒帯のみ（テキストボックス・結論メッセージ等は配置しない）
 *  - 4列テーブル: 指標 / 2025年度(n=47) / 直近 2026.1-3(n=10) / トレンド
 *
 * 実行手順:
 *  1. 02_helpers.gs を最新版で貼り付け済み
 *  2. このファイルを Apps Script に新規追加
 *  3. runS12bOnly() を実行 → 末尾に1枚追加
 */

function runS12bOnly() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  Logger.log("S12b 開始 - 現在: " + pres.getSlides().length + "枚");
  insertS12b_selfDataTrend(pres);
  Logger.log("S12b 完了 - 現在: " + pres.getSlides().length + "枚");
}

function insertS12b_selfDataTrend(pres) {
  var s = newSlide(pres, "自社データ：コメント参加率・リピーター比率の推移");

  // 4列テーブル: 指標 / 2025年度(n=47) / 直近 2026.1-3(n=10) / トレンド
  var colDefs = [
    { label: "指標",                 align: SlidesApp.ParagraphAlignment.START,  nameCol: true },
    { label: "2025年度\n(n=47)",     align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "直近 2026.1-3\n(n=10)", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "トレンド",             align: SlidesApp.ParagraphAlignment.START }
  ];
  var rows = [
    ["コメント参加率（%）",  "平均 0.52",   "平均 0.71",   "↗ 直近で改善傾向"],
    ["リピーター比率（%）",  "平均 4.36%",  "平均 4.88%",  "↗ 明確な上昇トレンド"]
  ];

  // リピーター比率行(index=1)をクリーム背景でハイライト
  makeColHighlightTable(s, 25, 130, 671, colDefs, rows, {
    rowH: 50,
    hdrH: 36,
    bodySize: 12,
    headerSize: 10,
    highlightRow: 1
  });

  // 注釈（※マークの注釈は8pt統一・MCCM標準）
  txt(s, "※ 2025年度（n=47）と直近2026.1-3（n=10）の比較。",
    25, 250, 671, 16, {
      size: 8, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
}
