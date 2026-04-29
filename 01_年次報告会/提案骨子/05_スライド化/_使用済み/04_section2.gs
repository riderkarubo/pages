/**
 * MCCM提案スライド化 - §2: 2025年度総括（5スライド）
 *
 * 最終更新: 2026-04-29 18:53
 *
 * マスター情報源: https://riderkarubo.github.io/pages/mccm-proposal/
 * （これ以外の場所からは情報を拾わない）
 *
 * 構成:
 *  S2: 章扉「2025年度総括 ― 1年間で積み上げた『資産価値』」
 *  S3: 視聴者数の急成長（テーブル5列）
 *  S4: 協賛実績の推移（テーブル4列）
 *  S5: マイルストーン（時系列9項目）
 *  S6: 7つの経営資産（7カード + 一言メッセージ）
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

/**
 * S6単独再生成用ラッパー
 * 使い方:
 *  1. Slides上で既存のS6スライド（タイトル「7つの経営資産」）を手動で削除
 *  2. Apps Scriptで rebuildS6_sevenAssets() を実行 → 末尾に新S6が追加される
 *  3. Slides上で新S6を正しい位置（S5の直後）にドラッグ移動
 */
function rebuildS6_sevenAssets() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  Logger.log("S6 再生成 開始 - 現在: " + pres.getSlides().length + "枚");
  insertS6_sevenAssets(pres);
  Logger.log("S6 再生成 完了 - 現在: " + pres.getSlides().length + "枚");
}

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

  // テーブル5列: 期間 / ライブ累計 / アーカイブ累計 / 合計 / 1回平均
  var colDefs = [
    { label: "期間", align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "ライブ累計", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "アーカイブ累計", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "合計", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "1回平均（ライブ）", align: SlidesApp.ParagraphAlignment.CENTER }
  ];
  var rows = [
    ["2024年度（2024.4〜2025.3）", "169,725人", "408,842人", "578,567人", "3,857人"],
    ["2025年度（2025.4〜2026.3）", "860,000人\n(+407%)", "580,000人\n(+42%)", "1,440,000人\n(+149%)", "5,860人\n(+52%)"]
  ];

  makeTable(s, 25, 100, 671, colDefs, rows, {
    rowH: 60,
    hdrH: 30,
    headerSize: 10,
    bodySize: 10,
    highlightRows: [1]  // 2025年度行をハイライト
  });

  // フッター注釈
  txt(s, "※ 前年比は2024年度実績との比較。",
    25, 240, 671, 16, {
      size: 9, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 主要メッセージ（中央下部に大文字訴求）
  rect(s, 25, 270, 671, 100, C_LIGHT, C_BORDER, 1);
  txt(s, "ライブ視聴者は5倍へ拡大、アーカイブを加えた年間1,440,000人接触",
    35, 280, 651, 26, {
      size: 13, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "1回平均5,860人は競合4社中B社（6,958人）に次ぐ業界2位ポジション",
    35, 310, 651, 22, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.CENTER,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "2024年8月のフローティング表示導入で5,000人超の安定維持を実現",
    35, 336, 651, 22, {
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

  // タイトル直下の一言メッセージ（タイトル黒帯 y=20-65 の直下）
  // 「毎週協賛配信が成立するリテールメディア」を C_ACCENT で部分強調
  var msgFull = "2年間で積み上げた7つの資産が、毎週協賛配信が成立するリテールメディアへと成長させることができた。";
  var msgBox = txt(s, msgFull, 27, 72, 666, 18, {
    size: 11, color: C_DARK, bold: false,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  // 部分強調：「毎週協賛配信が成立するリテールメディア」だけ C_ACCENT + Bold
  var keyword = "毎週協賛配信が成立するリテールメディア";
  var idx = msgFull.indexOf(keyword);
  if (idx >= 0) {
    var range = msgBox.getText().getRange(idx, idx + keyword.length);
    range.getTextStyle().setForegroundColor(C_ACCENT).setBold(true);
  }

  // 7カードを配置（4列×2行で7枚 = 4+3 でも、横7列でも、3+4 でもOK）
  // ここでは: 上段4枚（160pt幅）+ 下段3枚（中央寄せ）
  // 一言メッセージ用に y=90 → y=110 に20pt下シフト
  var cardW = 160;
  var cardH = 110;
  var gapX = 12;
  var rowGap = 14;
  var topY = 110;
  var bottomY = topY + cardH + rowGap;

  // 上段4枚 (x: 27, 199, 371, 543)
  var topRow = [
    { no: "資産 01", name: "視聴者資産",       value: "累計 約144万人", note: "アーカイブが3.3倍に拡張" },
    { no: "資産 02", name: "コンテンツ資産",   value: "91回",          note: "配信後もアーカイブとして蓄積" },
    { no: "資産 03", name: "メーカー関係資産", value: "約19社（のべ）\n累計44回", note: "月4〜5回 広告販売できる体制" },
    { no: "資産 04", name: "データ資産",       value: "CDP上に蓄積中", note: "視聴者ID×会員ID 突合" }
  ];

  // 下段3枚（中央寄せ: 720幅から3カードを中央配置 → 開始x = (720 - 3*160 - 2*12)/2 = 102）
  var bottomStartX = (SW - 3 * cardW - 2 * gapX) / 2;
  var bottomRow = [
    { no: "資産 05", name: "ノウハウ資産",   value: "内製化レクチャー\n全5回", note: "2026年1月〜5月実施" },
    { no: "資産 06", name: "出演者資産",     value: "計10名",       note: "第2期オーディションで3名追加" },
    { no: "資産 07", name: "パートナー資産", value: "LIPS連携",      note: "リテールメディア強化（メニュー化済）" }
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

/** 経営資産カード（左ボーダー付き） */
function drawAssetCard(slide, x, y, w, h, p) {
  // カード本体（白）
  rect(slide, x, y, w, h, C_WHITE, C_BORDER, 1);
  // 左ボーダー（赤）
  rect(slide, x, y, 3, h, C_ACCENT);

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

  // 値（強調・複数行対応）
  txt(slide, p.value, x + 10, y + 44, w - 20, 36, {
    size: 12, color: C_ACCENT, bold: true,
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
