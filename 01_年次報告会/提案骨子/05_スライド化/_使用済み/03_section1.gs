/**
 * MCCM提案スライド化 - §1: エグゼクティブサマリ（1スライド）
 *
 * P4（ユーザー再修正版）を完全踏襲。
 * 各オブジェクトの座標・サイズ・色・フォント・配置を
 * inspectSlide4()ログ実測値そのままで再現する。
 *
 * 実行手順:
 *  1. 02_helpers.gs を最新版で貼り付け済みであること
 *  2. このコードを最新版で貼り付け
 *  3. 既存P4はそのまま残す
 *  4. insertSection1() を実行 → 新しいスライドが追加される
 *  5. 既存P4と完全に揃っていれば、既存P4を消して新スライドを残す
 */

function insertSection1() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  Logger.log("§1 開始 - 現在: " + pres.getSlides().length + "枚");

  insertS1_executiveSummary(pres);
  Logger.log("S1 完了");

  Logger.log("§1 完了 - 現在: " + pres.getSlides().length + "枚");
}

// ============================================================
// S1: エグゼクティブサマリ（P4完全再現）
// ============================================================
function insertS1_executiveSummary(pres) {
  // newSlide() がタイトル黒帯（TITLE_X=9, Y=20, W=671, H=45）に
  // 白文字を自動配置する。全スライド共通。
  var s = newSlide(pres, "エグゼクティブサマリ");

  // ===== [1] メインメッセージ =====
  // x=41, y=88, w=639, h=24 / size=14, bold, #434343, CENTER, MIDDLE
  txt(s, "〔貴社とFireworkが目指すべきビジョンにアラインするメッセージ〕",
    41, 88, 639, 24, {
      size: 14, bold: true, color: "#434343",
      align: SlidesApp.ParagraphAlignment.CENTER,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // ===== [2] サブメッセージ =====
  // x=41, y=121, w=639, h=18 / size=10, regular, #434343, CENTER, MIDDLE
  txt(s, "1回の配信が1ヶ月後に2〜3倍のリーチを生み続けるストック型資産へ",
    41, 121, 639, 18, {
      size: 10, color: "#434343",
      align: SlidesApp.ParagraphAlignment.CENTER,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // ===== KPIカード 4枚 =====
  // 各KPIは7要素で構成: カード本体 / 左ボーダー / ラベル / 数字 / 単位 / 注釈
  // 左上：累計視聴（赤 #E94560）
  drawP4KpiCard(s, {
    cardX: 25, cardY: 164,
    label: "累計視聴（ライブ＋アーカイブ）",
    valueX: 118, valueY: 190, valueW: 108, valueH: 36, value: "200万",
    unitX: 228, unitY: 202, unitW: 43, unitH: 26, unit: "人超",
    note: "2024年度 578,567人 ＋ 2025年度 1,440,000人",
    accent: "#E94560"
  });

  // 右上：累計協賛売上（緑 #1B998B）
  drawP4KpiCard(s, {
    cardX: 367, cardY: 164,
    label: "累計 協賛売上（2024〜2025年度）",
    labelX: 379,
    valueX: 460, valueY: 190, valueW: 107, valueH: 36, value: "2,935",
    unitX: 570, unitY: 202, unitW: 40, unitH: 26, unit: "万円",
    noteX: 379,
    note: "2024年度 1,432万円 ＋ 2025年度 1,487万円",
    accent: "#1B998B"
  });

  // 左下：配信回数（黒 #1A1A2E）
  drawP4KpiCard(s, {
    cardX: 25, cardY: 264,
    label: "配信回数(2024〜2025年度 累計)",
    valueX: 148, valueY: 290, valueW: 60, valueH: 36, value: "91",
    unitX: 210, unitY: 302, unitW: 22, unitH: 26, unit: "回",
    note: "2024年度開始〜2025年度までの累計配信数",
    accent: "#1A1A2E"
  });

  // 右下：平均視聴者数（黒 #1A1A2E）
  drawP4KpiCard(s, {
    cardX: 367, cardY: 264,
    label: "平均視聴者数（2025年度）",
    labelX: 379,
    valueX: 457, valueY: 290, valueW: 107, valueH: 36, value: "5,860",
    unitX: 567, unitY: 302, unitW: 50, unitH: 26, unit: "人/回",
    noteX: 379,
    note: "前年比 +52%（前年 3,857人）",
    accent: "#1A1A2E"
  });
}

/**
 * P4実測値完全準拠のKPIカード描画
 *  カード本体: 330×92 #FFFFFF / 左ボーダー: 4×92 accent
 *  ラベル: y+6 (+6=170-164), w=306, h=14, size=9, bold, #434343, START, MIDDLE
 *  数字: 個別座標, size=28, bold, accent, END, MIDDLE
 *  単位: 個別座標, size=12, bold, accent, START, BOTTOM
 *  注釈: y+74 (+74=238-164), w=306, h=14, size=8, regular, #666666, START, MIDDLE
 *
 *  右側カードは labelX/noteX を 379 に指定（cardX=367+12 ではなく)。
 *  左側カードは labelX/noteX 未指定で 37（cardX=25+12）にする。
 */
function drawP4KpiCard(s, p) {
  // [3/9/15/21] カード本体
  rect(s, p.cardX, p.cardY, 330, 92, "#FFFFFF");
  // [4/10/16/22] 左ボーダー（4pt幅）
  rect(s, p.cardX, p.cardY, 4, 92, p.accent);

  var labelX = p.labelX || (p.cardX + 12);  // 25→37, 367→379
  var noteX = p.noteX || (p.cardX + 12);

  // [5/11/17/23] ラベル
  txt(s, p.label, labelX, p.cardY + 6, 306, 14, {
    size: 9, bold: true, color: "#434343",
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // [6/12/18/24] 数字（END alignment, MIDDLE va）
  txt(s, p.value, p.valueX, p.valueY, p.valueW, p.valueH, {
    size: 28, bold: true, color: p.accent,
    align: SlidesApp.ParagraphAlignment.END,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // [7/13/19/25] 単位（START alignment, BOTTOM va）
  txt(s, p.unit, p.unitX, p.unitY, p.unitW, p.unitH, {
    size: 12, bold: true, color: p.accent,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.BOTTOM
  });

  // [8/14/20/26] 注釈
  txt(s, p.note, noteX, p.cardY + 74, 306, 14, {
    size: 8, color: "#666666",
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
}
