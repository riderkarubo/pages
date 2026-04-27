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
/**
 * MCCM提案スライド化 - §2: 2025年度総括（5スライド）
 *
 * マスター情報源: https://riderkarubo.github.io/pages/mccm-proposal/
 * （これ以外の場所からは情報を拾わない）
 *
 * 構成:
 *  S2: 章扉「2025年度総括 ― 1年間で積み上げた『資産価値』」
 *  S3: 視聴者数の急成長（テーブル5列）
 *  S4: 協賛実績の推移（テーブル4列）
 *  S5: マイルストーン（時系列9項目）
 *  S6: 7つの経営資産（7カード）
 *
 * 実行手順:
 *  1. 02_helpers.gs を貼り付け済みであること
 *  2. このコードを Apps Script に貼り付け
 *  3. insertSection2() を実行 → 末尾に5枚追加される
 *  4. ブラウザで確認 → 修正指示
 */

function insertSection2() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  Logger.log("§2 開始 - 現在: " + pres.getSlides().length + "枚");

  insertS2_chapterCover(pres);
  Logger.log("S2 完了");

  insertS3_viewerGrowth(pres);
  Logger.log("S3 完了");

  insertS4_sponsorRecord(pres);
  Logger.log("S4 完了");

  insertS5_milestones(pres);
  Logger.log("S5 完了");

  insertS6_sevenAssets(pres);
  Logger.log("S6 完了");

  Logger.log("§2 完了 - 現在: " + pres.getSlides().length + "枚");
}

// ============================================================
// 単独実行用ラッパー（Apps Scriptから関数単体を呼ぶ用）
// ============================================================
function runS2Only() { insertS2_chapterCover(SlidesApp.openById(PRESENTATION_ID)); }
function runS3Only() { insertS3_viewerGrowth(SlidesApp.openById(PRESENTATION_ID)); }
function runS4Only() { insertS4_sponsorRecord(SlidesApp.openById(PRESENTATION_ID)); }
function runS5Only() { insertS5_milestones(SlidesApp.openById(PRESENTATION_ID)); }
function runS6Only() { insertS6_sevenAssets(SlidesApp.openById(PRESENTATION_ID)); }

// ============================================================
// S2: 章扉「2025年度総括 ― 1年間で積み上げた資産価値」
// ============================================================
function insertS2_chapterCover(pres) {
  var s = newSlide(pres, "2025年度総括 ― 1年間で積み上げた「資産価値」");

  // サブタイトル（マスター情報源から）
  txt(s, "出典：マスターデータ・協賛実績xlsx・SNS数値CSV",
    25, 76, 671, 24, {
      size: 11, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 中央訴求ブロック2点（横並び大型カード）
  // 視聴者数+149%
  drawHeadlineCard(s, {
    x: 60, y: 150, w: 285, h: 180,
    accent: C_ACCENT,  // 赤
    label: "年度別 視聴者数",
    value: "+149%",
    breakdown: [
      "ライブ累計   +407%（169,725人 → 860,000人）",
      "アーカイブ   +42%  （408,842人 → 580,000人）",
      "1回平均ライブ +52% （3,857人 → 5,860人）"
    ]
  });

  // 協賛+44%
  drawHeadlineCard(s, {
    x: 375, y: 150, w: 285, h: 180,
    accent: C_ACCENT2,  // 緑
    label: "年度別 協賛実施回数",
    value: "+44%",
    breakdown: [
      "実施回数   18回 → 26回（+44%）",
      "売上合計   1,432万円 → 1,487万円（+4%）",
      "2024年度は単価300万×2件で売上を積み上げた年",
      "2025年度は件数ベースで安定化"
    ]
  });

  // フッター
  txt(s, "次章で各指標の詳細を解説します",
    25, 360, 671, 18, {
      size: 10, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.CENTER,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
}

/** 章扉用ヘッドラインカード（大きな数字＋ラベル＋内訳箇条書き） */
function drawHeadlineCard(slide, p) {
  // カード本体（白）
  rect(slide, p.x, p.y, p.w, p.h, C_WHITE, C_BORDER, 1);
  // 上部アクセントバー
  rect(slide, p.x, p.y, p.w, 4, p.accent);

  // ラベル
  txt(slide, p.label, p.x + 16, p.y + 14, p.w - 32, 16, {
    size: 10, color: C_GRAY, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 大きな数字
  txt(slide, p.value, p.x + 16, p.y + 36, p.w - 32, 56, {
    size: 44, color: p.accent, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 内訳箇条書き
  var bulletText = p.breakdown.map(function(b) { return "・" + b; }).join("\n");
  txt(slide, bulletText, p.x + 16, p.y + 100, p.w - 32, p.h - 110, {
    size: 9, color: C_TEXT,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.START
  });
}

// ============================================================
// S3: 視聴者数の急成長
// ============================================================
function insertS3_viewerGrowth(pres) {
  var s = newSlide(pres, "視聴者数 +149%（年度比）");

  // テーブル4列: 指標 / 2024年度 / 2025年度 / 前年比
  var colDefs = [
    { label: "指標", align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "2024年度（2024.4〜2025.3）", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "2025年度（2025.4〜2026.3）", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "前年比", align: SlidesApp.ParagraphAlignment.CENTER }
  ];
  var rows = [
    ["ライブ累計",       "169,725人",   "860,000人",   "+407%"],
    ["アーカイブ累計",   "408,842人",   "580,000人",   "+42%"],
    ["合計視聴",         "578,567人",   "1,440,000人", "+149%"],
    ["1回平均(ライブ)",  "3,857人/回",  "5,860人/回",  "+52%"]
  ];

  makeTable(s, 25, 110, 671, colDefs, rows, {
    rowH: 32,
    hdrH: 30,
    headerSize: 10,
    bodySize: 10,
    highlightRows: [2]  // 「合計視聴」行をハイライト
  });

  // フッター注釈
  txt(s, "※ 前年比は2024年度実績との比較。",
    25, 270, 671, 16, {
      size: 9, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 主要メッセージ（中央下部に大文字訴求）
  rect(s, 25, 295, 671, 80, C_LIGHT, C_BORDER, 1);
  txt(s, "ライブ視聴者は5倍へ拡大、アーカイブを加えた年間1,440,000人接触",
    35, 304, 651, 24, {
      size: 13, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "1回平均5,860人は競合4社中B社（6,958人）に次ぐ業界2位ポジション",
    35, 330, 651, 20, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.CENTER,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "2024年8月のフローティング表示導入で5,000人超の安定維持を実現",
    35, 352, 651, 20, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.CENTER,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S4: 協賛実績の推移
// ============================================================
function insertS4_sponsorRecord(pres) {
  var s = newSlide(pres, "協賛実績の推移");

  // テーブル4列: 期間 / 協賛回数 / 売上合計 / 備考
  var colDefs = [
    { label: "期間", align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "協賛実施回数", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "売上合計", align: SlidesApp.ParagraphAlignment.END },
    { label: "備考", align: SlidesApp.ParagraphAlignment.START }
  ];
  var rows = [
    ["2024年度\n（2024.4〜2025.3）", "18回", "14,320,290円", "単価300万円を2件受注（ブランドジャック型）／基本単価50万円"],
    ["2025年度\n（2025.4〜2026.3）", "26回\n(+44%)", "14,868,840円\n(+4%)", "基本単価50万円（+従量課金）"]
  ];

  makeTable(s, 25, 100, 671, colDefs, rows, {
    rowH: 60,
    hdrH: 30,
    headerSize: 10,
    bodySize: 9,
    highlightRows: [1]
  });

  // ポイント注釈ボックス
  rect(s, 25, 240, 671, 130, "#F0F7FF", "#3B82F6", 1);
  txt(s, "💡 ポイント",
    35, 250, 651, 18, {
      size: 11, color: "#1E3A8A", bold: true,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "2024年度は高単価案件2件（300万円／件）で売上を積み上げた年。",
    35, 272, 651, 22, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "2025年度は件数ベース（26回）で受注を安定化させ、売上はほぼ横ばい維持。",
    35, 295, 651, 22, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "→ 2026年度は件数KPI確定（48件/年）＋単価見直しでトップライン成長を目指す。",
    35, 318, 651, 22, {
      size: 11, color: C_ACCENT, bold: true,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S5: マイルストーン
// ============================================================
function insertS5_milestones(pres) {
  var s = newSlide(pres, "マイルストーン");

  // 縦タイムライン9項目
  // レイアウト: 左に日付（90pt幅）/ 中央に丸ドット/縦線 / 右に内容
  var items = [
    { date: "2024年4月",      text: "本格稼働スタート", milestone: true },
    { date: "2024年6月5日",   text: "1社協賛 初回配信実施（melt）", milestone: true },
    { date: "2024年8月",      text: "オンラインサイトのフローティング表示導入により視聴者数が急増（5月平均約300人→8月以降5,000人超）", milestone: true },
    { date: "2024年8月19日",  text: "ブランド横断型 初回配信実施（売れ筋ナプキン徹底比較／ユニ・チャーム協賛）", milestone: true },
    { date: "2024年11月",     text: "毎月メーカー協賛を安定獲得する体制を確立", milestone: false },
    { date: "2025年10月",     text: "第2期出演者オーディションにより新メンバー3名加入（1期生含め計10名）", milestone: true },
    { date: "2025年10月28日", text: "LIPSコラボ 初回配信実施（KATE）", milestone: true },
    { date: "2026年1月〜5月", text: "内製化レクチャー開始（全5回）", milestone: false },
    { date: "2026年度 目標",  text: "アプリ内ライブLP導線の改修により視聴者数増加予定", milestone: true }
  ];

  var startY = 86;
  var rowH = 32;
  var dateX = 30;
  var dateW = 110;
  var dotX = 152;
  var dotSize = 10;
  var lineX = dotX + dotSize / 2 - 1;
  var contentX = 175;
  var contentW = 520;

  // 縦線（最初のドットから最後のドットまで）
  rect(s, lineX, startY + 12, 2, rowH * (items.length - 1), C_BORDER);

  for (var i = 0; i < items.length; i++) {
    var y = startY + i * rowH;
    var item = items[i];

    // 日付（左）
    txt(s, item.date, dateX, y + 4, dateW, 20, {
      size: 9, color: C_GRAY, bold: true,
      align: SlidesApp.ParagraphAlignment.END,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

    // ドット（中央・マイルストーンは赤、それ以外はグレー）
    var dotColor = item.milestone ? C_ACCENT : C_GRAY;
    ellipse(s, dotX, y + 8, dotSize, dotSize, dotColor);

    // 内容（右）
    txt(s, item.text, contentX, y + 4, contentW, 22, {
      size: 9.5, color: item.milestone ? C_DARK : C_TEXT, bold: item.milestone,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  }
}

// ============================================================
// S6: 7つの経営資産
// ============================================================
function insertS6_sevenAssets(pres) {
  var s = newSlide(pres, "7つの経営資産");

  // 7カードを配置（4列×2行で7枚 = 4+3 でも、横7列でも、3+4 でもOK）
  // ここでは: 上段4枚（160pt幅）+ 下段3枚（中央寄せ）
  var cardW = 160;
  var cardH = 110;
  var gapX = 12;
  var rowGap = 14;
  var topY = 90;
  var bottomY = topY + cardH + rowGap;

  // 上段4枚 (x: 27, 199, 371, 543)
  // 色は資産の意味でグルーピング（赤=リーチ／橙=ストック／緑=売上／紺=基盤／紫=知／金=人／シアン=外部）
  var topRow = [
    { no: "資産 01", name: "視聴者資産",       value: "累計 約144万人",          note: "アーカイブが3.3倍に拡張",            accent: "#E94560" },
    { no: "資産 02", name: "コンテンツ資産",   value: "91回",                     note: "配信後もアーカイブとして蓄積",        accent: "#F59E0B" },
    { no: "資産 03", name: "メーカー関係資産", value: "約19社（のべ）\n累計44回", note: "月4〜5回 広告販売できる体制",        accent: "#1B998B" },
    { no: "資産 04", name: "データ資産",       value: "CDP上に蓄積中",            note: "視聴者ID×会員ID 突合",              accent: "#0F3460" }
  ];

  // 下段3枚（中央寄せ: 720幅から3カードを中央配置 → 開始x = (720 - 3*160 - 2*12)/2 = 102）
  var bottomStartX = (SW - 3 * cardW - 2 * gapX) / 2;
  var bottomRow = [
    { no: "資産 05", name: "ノウハウ資産",   value: "内製化レクチャー\n全5回", note: "2026年1月〜5月実施",                  accent: "#8B5CF6" },
    { no: "資産 06", name: "出演者資産",     value: "計10名",                  note: "第2期オーディションで3名追加",        accent: "#EAB308" },
    { no: "資産 07", name: "パートナー資産", value: "LIPS連携",                note: "リテールメディア強化（メニュー化済）", accent: "#06B6D4" }
  ];

  // 上段の開始x（中央寄せ）
  var topStartX = (SW - 4 * cardW - 3 * gapX) / 2;

  // 上段描画
  for (var i = 0; i < topRow.length; i++) {
    var x = topStartX + i * (cardW + gapX);
    drawAssetCard(s, x, topY, cardW, cardH, topRow[i]);
  }

  // 下段描画
  for (var j = 0; j < bottomRow.length; j++) {
    var x2 = bottomStartX + j * (cardW + gapX);
    drawAssetCard(s, x2, bottomY, cardW, cardH, bottomRow[j]);
  }
}

/** 経営資産カード（左ボーダー付き・アクセント色は資産ごと可変） */
function drawAssetCard(slide, x, y, w, h, p) {
  var accent = p.accent || C_ACCENT;

  // カード本体（白）
  rect(slide, x, y, w, h, C_WHITE, C_BORDER, 1);
  // 左ボーダー（資産ごとのアクセント色）
  rect(slide, x, y, 3, h, accent);

  // 資産番号（小さくグレー）
  txt(slide, p.no, x + 10, y + 6, w - 20, 12, {
    size: 8, color: C_GRAY, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 名前（中央太字）
  txt(slide, p.name, x + 10, y + 22, w - 20, 16, {
    size: 11, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 値（強調・資産ごとのアクセント色・複数行対応）
  txt(slide, p.value, x + 10, y + 44, w - 20, 36, {
    size: 12, color: accent, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 注釈
  txt(slide, p.note, x + 10, y + 84, w - 20, 22, {
    size: 8, color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.START
  });
}
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
/**
 * MCCM提案スライド化 - §4: 2026年度の2本柱（7スライド）
 *
 * マスター情報源: https://riderkarubo.github.io/pages/mccm-proposal/
 * （これ以外の場所からは情報を拾わない）
 *
 * 構成:
 *  S13: 2026年度の2本柱（章扉）
 *  S14: 柱① 現状課題（価格・粗利率）
 *  S15: 柱① 提案①：単価是正（A社比較のみ・グレーアウト部分非掲載）
 *  S16: 柱① 体制改善：現状課題（工数膨張・薬機法CK詰まり）
 *  S17: 柱① 体制改善 上半期①：レギュレーション設定
 *  S18: 柱① 体制改善：薬機法CK + AI活用（統合）
 *  S19: 柱② 社員インフルエンサー強化
 *
 * 実行手順:
 *  1. 02_helpers.gs / 04_section2.gs / 05_section3.gs を貼り付け済みであること
 *  2. このコードを Apps Script に貼り付け
 *  3. insertSection4() を実行 → 末尾に7枚追加される
 */

function insertSection4() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  Logger.log("§4 開始 - 現在: " + pres.getSlides().length + "枚");

  insertS13_pillarsCover(pres);   Logger.log("S13 完了");
  insertS14_pillar1Issues(pres);  Logger.log("S14 完了");
  insertS15_priceRevision(pres);  Logger.log("S15 完了");
  insertS16_opsIssues(pres);      Logger.log("S16 完了");
  insertS17_regulation(pres);     Logger.log("S17 完了");
  insertS18_yakkihoAI(pres);      Logger.log("S18 完了");
  insertS19_pillar2(pres);        Logger.log("S19 完了");

  Logger.log("§4 完了 - 現在: " + pres.getSlides().length + "枚");
}

// ============================================================
// S13: 2026年度の2本柱（章扉）
// ============================================================
function insertS13_pillarsCover(pres) {
  var s = newSlide(pres, "2026年度の2本柱");

  // サブタイトル
  txt(s, "KPI（協賛配信受注 年間48件・2,400万円）達成に向けた2つの重点施策",
    25, 76, 671, 24, {
      size: 11, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 2本柱カード（横並び大型）
  var cardW = 320, cardH = 200, gap = 24;
  var startX = (SW - 2 * cardW - gap) / 2;
  var startY = 130;

  // 柱①
  rect(s, startX, startY, cardW, cardH, "#FFF5F7", C_ACCENT, 2);
  txt(s, "柱 ①", startX + 14, startY + 14, cardW - 28, 18, {
    size: 11, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "協賛配信のメニュー・体制改善", startX + 14, startY + 38, cardW - 28, 28, {
    size: 17, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "受注後の運営体制を標準化し、月4回以上の安定受注を実現する。",
    startX + 14, startY + 76, cardW - 28, 36, {
      size: 11, color: "#555",
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
  txt(s, "・提案①：フルサポートプラン単価是正（A社比較）\n・体制改善：レギュレーション設定\n・体制改善：薬機法CKフロー改善 + AI活用",
    startX + 14, startY + 116, cardW - 28, 70, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });

  // 柱②
  rect(s, startX + cardW + gap, startY, cardW, cardH, "#F0FAF8", C_ACCENT2, 2);
  txt(s, "柱 ②", startX + cardW + gap + 14, startY + 14, cardW - 28, 18, {
    size: 11, color: C_ACCENT2, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "社員インフルエンサー強化", startX + cardW + gap + 14, startY + 38, cardW - 28, 28, {
    size: 17, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "出演者を育成・サポートし、リピーター比率の底上げを実現する。",
    startX + cardW + gap + 14, startY + 76, cardW - 28, 36, {
      size: 11, color: "#555",
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
  txt(s, "・SNSフォロワー上位スタッフを優先発掘・教育\n・既存出演者のSNS開設・強化（&ST連動）\n・社員インフルエンサーコンテンツの自社メディア活用\n・出演者への薬事知識研修",
    startX + cardW + gap + 14, startY + 116, cardW - 28, 80, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
}

// ============================================================
// S14: 柱① 現状課題
// ============================================================
function insertS14_pillar1Issues(pres) {
  var s = newSlide(pres, "柱① 協賛配信のメニュー：現状課題");

  // 2カード（縦並び・厚め）
  var cardW = 671, cardH = 130, gap = 16;
  var startX = 25;
  var startY = 96;

  // 課題①
  rect(s, startX, startY, cardW, cardH, "#FFF5F7", C_ACCENT, 1);
  rect(s, startX, startY, 4, cardH, C_ACCENT);
  txt(s, "現状課題 ①", startX + 16, startY + 12, cardW - 32, 16, {
    size: 10, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "競合他社と比べて提供価値に対する価格が低い", startX + 16, startY + 32, cardW - 32, 24, {
    size: 14, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "フルサポート50万円のみ。A社の15分プラン（45万円）と同水準の価格で60分配信＋データ連動を提供しており、市場価格との乖離がある。",
    startX + 16, startY + 60, cardW - 32, 60, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });

  // 課題②
  var y2 = startY + cardH + gap;
  rect(s, startX, y2, cardW, cardH, "#FFF5F7", C_ACCENT, 1);
  rect(s, startX, y2, 4, cardH, C_ACCENT);
  txt(s, "現状課題 ②", startX + 16, y2 + 12, cardW - 32, 16, {
    size: 10, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "貴社の他メニューと比較して粗利率が低い（69.3%）", startX + 16, y2 + 32, cardW - 32, 24, {
    size: 14, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "ショート動画83.3%・Xフォロリポ81.6%と比べて低い。制作外注費の原価（台本作成、現場ディレクション、クリエイティブ制作）が50万円に内包されている構造による。",
    startX + 16, y2 + 60, cardW - 32, 60, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
}

// ============================================================
// S15: 柱① 提案①：単価是正（A社比較）
// ============================================================
function insertS15_priceRevision(pres) {
  var s = newSlide(pres, "提案①：フルサポートプランの単価是正（A社比較）");

  // 説明文
  txt(s, "A社の競合プランと比較し、マツキヨココカラライブの提供価値（60分・データ連動・5,000人）と価格の乖離を整理。",
    25, 96, 671, 22, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // テーブル
  var colDefs = [
    { label: "プラン", align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "価格", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "内容", align: SlidesApp.ParagraphAlignment.START }
  ];
  var rows = [
    ["A社 スタンダード",                "45万円", "15分・最大3商品"],
    ["A社 ポップアップ中継",            "70万円", "生中継"],
    ["マツキヨココカラライブ（現行）",   "50万円", "60分・データ連動・5,000人"]
  ];

  makeTable(s, 25, 130, 671, colDefs, rows, {
    rowH: 36,
    hdrH: 30,
    headerSize: 11,
    bodySize: 11,
    highlightRows: [2]
  });

  // 結論ボックス
  rect(s, 25, 280, 671, 90, "#FFF5F5", C_ACCENT, 1);
  txt(s, "結論", 35, 286, 651, 18, {
    size: 10, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "同水準の価格で60分配信＋データ連動を提供しており、市場価格との乖離がある。",
    35, 308, 651, 22, {
      size: 13, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "次期分科会にて、本提案に対する議論・粗利率インパクトを別途ご提示します。",
    35, 334, 651, 22, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S16: 柱① 体制改善 現状課題
// ============================================================
function insertS16_opsIssues(pres) {
  var s = newSlide(pres, "協賛配信の運営体制改善 ― 現状課題");

  // リード文
  txt(s, "月5回以上の安定受注はできている。次の課題は、受注後の運営フローを2社間で最適化し、貴社の作業工数負荷を下げること。",
    25, 96, 671, 30, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 2カード（縦並び）
  var cardW = 671, cardH = 110, gap = 14;
  var startX = 25;
  var startY = 132;

  rect(s, startX, startY, cardW, cardH, "#FFF5F7", C_ACCENT, 1);
  rect(s, startX, startY, 4, cardH, C_ACCENT);
  txt(s, "現状課題 ①  受け入れ無制限による工数膨張", startX + 16, startY + 10, cardW - 32, 22, {
    size: 13, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "台本初稿後の商品変更・追加に上限がなく、台本量・薬機法チェック工数が読めない。競合（A社）が約2週間で準備を完了するのに対し、現状の準備期間は1.5〜2ヶ月。実例として初稿後の商品変更により台本が6稿に達したケースも。",
    startX + 16, startY + 36, cardW - 32, 70, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });

  var y2 = startY + cardH + gap;
  rect(s, startX, y2, cardW, cardH, "#FFF5F7", C_ACCENT, 1);
  rect(s, startX, y2, 4, cardH, C_ACCENT);
  txt(s, "現状課題 ②  薬機法チェックが詰まる構造的問題", startX + 16, y2 + 10, cardW - 32, 22, {
    size: 13, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "修正稿ベースの後ろ倒し運用に加え、貴社法務からのNGの代替表現が提示されず現場が毎回詰まる。台本作成者・貴社担当者ともに「どこまでOKか」の共通基準がなく、都度確認が発生している。",
    startX + 16, y2 + 36, cardW - 32, 70, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
}

// ============================================================
// S17: 柱① 体制改善 上半期①：レギュレーション設定
// ============================================================
function insertS17_regulation(pres) {
  var s = newSlide(pres, "上半期 ① レギュレーション設定");

  // バッジ＋リード文
  rect(s, 25, 88, 100, 22, C_ACCENT);
  txt(s, "上半期（〜2026年9月）", 25, 88, 100, 22, {
    size: 9, color: C_WHITE, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  txt(s, "配信ごとの商品変更・追加・修正の無制限受け入れが、台本・チェック工数を押し上げている。\n受注前にメーカー側と共通ルールを定義することで、工数負荷を構造的に下げる。",
    25, 118, 671, 40, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // テーブル
  var colDefs = [
    { label: "項目", align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "内容", align: SlidesApp.ParagraphAlignment.START }
  ];
  var rows = [
    ["商品確定締め切り", "配信X週間前（要協議）"],
    ["主要紹介アイテム数", "◯アイテムまで（カラバリ・香りなどのバリエーションは除く、要協議）"],
    ["台本修正回数", "◯回まで（◯回目以降は別途修正費用発生。誤植除く）"]
  ];

  makeTable(s, 25, 170, 671, colDefs, rows, {
    rowH: 50,
    hdrH: 30,
    headerSize: 11,
    bodySize: 11
  });

  // フッター
  txt(s, "※ 具体的な数値（締切週数・アイテム数・修正回数）は別途協議",
    25, 360, 671, 18, {
      size: 9, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S18: 柱① 体制改善：薬機法CK + AI活用（統合）
// ============================================================
function insertS18_yakkihoAI(pres) {
  var s = newSlide(pres, "体制改善 ② 薬機法CKフロー改善 ＋ ③ AI活用");

  // 上半期バッジ
  rect(s, 25, 88, 100, 20, C_ACCENT);
  txt(s, "上半期 ②", 25, 88, 100, 20, {
    size: 9, color: C_WHITE, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "薬機法CKフロー改善", 130, 88, 565, 20, {
    size: 12, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 上段2カード（横並び）
  var cardW = 330, cardH = 110, gap = 11;
  var startX = 25;
  var topY = 116;

  rect(s, startX, topY, cardW, cardH, "#FFF5F7");
  rect(s, startX, topY, 3, cardH, C_ACCENT);
  txt(s, "ライブ配信向け薬事ルールの制定", startX + 14, topY + 10, cardW - 28, 20, {
    size: 12, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "ライブ配信に特化した判断基準・ルールを設ける。参考：A社は薬機法を熟知した美容部員教育担当が出演者を兼任しており、法務の薬事CKを実施していない。今後、社員インフルエンサーを育成していく際には、薬事知識のインプットも併せて進められるとよい。",
    startX + 14, topY + 32, cardW - 28, 74, {
      size: 9, color: "#555",
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });

  rect(s, startX + cardW + gap, topY, cardW, cardH, "#FFF5F7");
  rect(s, startX + cardW + gap, topY, 3, cardH, C_ACCENT);
  txt(s, "GoogleAI を活用した薬事CKフロー構築", startX + cardW + gap + 14, topY + 10, cardW - 28, 20, {
    size: 12, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "別途、Fireworkから具体的なフローをご提案。法務確認の前段でAIによる一次チェックを導入し、現場が詰まる構造を解消する。",
    startX + cardW + gap + 14, topY + 32, cardW - 28, 74, {
      size: 9, color: "#555",
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });

  // 下半期バッジ
  var bottomY = 246;
  rect(s, 25, bottomY, 100, 20, C_ACCENT2);
  txt(s, "下半期 ③", 25, bottomY, 100, 20, {
    size: 9, color: C_WHITE, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "AI活用による業務効率化", 130, bottomY, 565, 20, {
    size: 12, color: C_ACCENT2, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 下段ブロック
  rect(s, 25, bottomY + 28, 671, 110, "#F0FAF8");
  rect(s, 25, bottomY + 28, 3, 110, C_ACCENT2);
  txt(s, "Google AI（Gemini・NotebookLM・AI Studio）を実績レポートに活用",
    39, bottomY + 36, 657, 22, {
      size: 12, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "貴社が活用中のGoogle AI（Gemini・NotebookLM・AI Studio）を、配信後の実績レポート作成などの定型業務に活用。\n貴社セキュリティポリシーに則ったかたちでFireworkが具体的な活用法をご提案。",
    39, bottomY + 62, 657, 64, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
}

// ============================================================
// S19: 柱② 社員インフルエンサー強化
// ============================================================
function insertS19_pillar2(pres) {
  var s = newSlide(pres, "柱② 社員インフルエンサー強化");

  // メイン訴求ボックス
  rect(s, 25, 88, 671, 64, "#F0FAF8", C_ACCENT2, 2);
  txt(s, "「推しの出演者が出るから観る」という再訪動機を作り、リピーター比率を底上げ",
    35, 96, 651, 24, {
      size: 13, color: C_ACCENT2, bold: true,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "出演者（美容部員・店舗スタッフ）のライブ配信スキルを育成・サポートする。SNSでの発信は個人の自由な活動を基本とし、ライブ配信の適性・志望がある人材を社員・パートを問わず発掘・育成するスタンスで推進。",
    35, 122, 651, 28, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });

  // 具体的アクション（4枚カード・横2×縦2）
  txt(s, "具体的アクション", 25, 162, 671, 18, {
    size: 11, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  var cardW = 330, cardH = 60, gap = 11, vgap = 10;
  var startX = 25, startY = 184;

  drawActionCard(s, startX, startY, cardW, cardH, {
    no: "①",
    title: "SNSフォロワー上位スタッフを優先発掘・教育",
    detail: "ライブ配信出演者として教育・サポートする"
  });
  drawActionCard(s, startX + cardW + gap, startY, cardW, cardH, {
    no: "②",
    title: "既存出演者のSNS開設・強化（&ST連動）",
    detail: "STさま施策との連動"
  });
  drawActionCard(s, startX, startY + cardH + vgap, cardW, cardH, {
    no: "③",
    title: "社員インフルエンサーコンテンツの自社メディア活用",
    detail: "ECサイト、アプリなどでの再活用"
  });
  drawActionCard(s, startX + cardW + gap, startY + cardH + vgap, cardW, cardH, {
    no: "④",
    title: "出演者への薬事知識研修（A社事例）",
    detail: "A社はライブ配信台本において法務CK不要な体制を実現"
  });

  // 根拠フッター
  rect(s, 25, 318, 671, 56, "#FFF5F5", C_ACCENT, 1);
  txt(s, "根拠", 35, 322, 651, 16, {
    size: 9, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "競合B社は個人SNSフォロワー数万〜16万超のBAが複数在籍しファン流入→高エンゲージを実現。\nマツキヨココカラライブはリピーター比率5.0%（業界最下位）であり、出演者ファンベース構築が最重点課題。",
    35, 340, 651, 32, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
}

/** アクションカード（番号＋タイトル＋詳細） */
function drawActionCard(slide, x, y, w, h, p) {
  rect(slide, x, y, w, h, C_WHITE, C_BORDER, 1);
  rect(slide, x, y, 4, h, C_ACCENT2);

  // 番号（左寄せ大きめ）
  txt(slide, p.no, x + 10, y + 6, 22, h - 12, {
    size: 18, color: C_ACCENT2, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // タイトル
  txt(slide, p.title, x + 36, y + 8, w - 46, 20, {
    size: 10.5, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 詳細
  txt(slide, p.detail, x + 36, y + 30, w - 46, h - 36, {
    size: 9, color: "#555",
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.START
  });
}
