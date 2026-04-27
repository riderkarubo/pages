/**
 * MCCM提案スライド化 - §3: 2026年度KPI目標設定（7スライド・v2 実測テンプレ準拠版）
 *
 * マスター情報源: https://riderkarubo.github.io/pages/mccm-proposal/
 * （これ以外の場所からは情報を拾わない）
 *
 * Issy手直し解析パターン（P4-P7 inspect 実測値ベース）:
 *  - タイトル黒帯: x=9, y=20, w=671, h=45（newSlide()自動配置）
 *  - 結論メッセージ: drawConclusionMessage() で y≈90, 16pt bold #434343 + 強調語だけ #FA006D
 *  - テーブル列強調: makeColHighlightTable() で2025年度や目標列を #FFF2CC クリーム背景
 *  - ポイントBOX: drawTipsBox() で薄水色 + 💡見出し + 結論行は #FA006D 太字
 *  - タイムライン: drawTimelineRow() / drawTimelineAxis() で #FA006D丸=重要 グレー丸=中継ぎ
 *  - 資産カード: drawAssetCardP7() でアクセント色を意味別に分ける
 *  - 強調色: 必ず C_ACCENT (#FA006D マゼンタ) 経由（ハードコード禁止）
 *
 * 構成:
 *  S7:  2026年度KPI（協賛配信受注 年間48件・2,400万円）
 *  S8:  売上構造の分解（KPIツリー）— 画像差し替え方式
 *  S9:  マイルストーン①：メディアパワー（4指標）
 *  S10: コスメカテゴリ限定比較（マツキヨ × 競合4社・2026.1-3）
 *  S10b: 競合比較：他社の強み（4観点）
 *  S11: メディアパワーの因数分解（3カード）
 *  S12: マイルストーン②：リピーター数値（3指標・追加提案）
 *
 * 実行手順:
 *  1. 02_helpers.gs を最新版で貼り付け済み
 *  2. このファイルを Apps Script に新規追加
 *  3. insertSection3V2() を実行 → 末尾に7枚追加される
 *
 * 単独実行用ラッパー:
 *  runS7Only() / runS8Only() / runS9Only() / runS10Only() / runS10bOnly() / runS11Only() / runS12Only()
 */

function insertSection3V2() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  Logger.log("§3 v2 開始 - 現在: " + pres.getSlides().length + "枚");

  insertS7v2_kpiHeadline(pres);          Logger.log("S7 完了");
  insertS8v2_revenueTree(pres);          Logger.log("S8 完了");
  insertS9v2_milestone1(pres);           Logger.log("S9 完了");
  insertS10v2_benchmark(pres);           Logger.log("S10 完了");
  insertS10bv2_competitorStrength(pres); Logger.log("S10b 完了");
  insertS11v2_factoring(pres);           Logger.log("S11 完了");
  insertS12v2_milestone2(pres);          Logger.log("S12 完了");

  Logger.log("§3 v2 完了 - 現在: " + pres.getSlides().length + "枚");
}

function runS7Only()   { insertS7v2_kpiHeadline(SlidesApp.openById(PRESENTATION_ID)); }
function runS8Only()   { insertS8v2_revenueTree(SlidesApp.openById(PRESENTATION_ID)); }
function runS9Only()   { insertS9v2_milestone1(SlidesApp.openById(PRESENTATION_ID)); }
function runS10Only()  { insertS10v2_benchmark(SlidesApp.openById(PRESENTATION_ID)); }
function runS10bOnly() { insertS10bv2_competitorStrength(SlidesApp.openById(PRESENTATION_ID)); }
function runS11Only()  { insertS11v2_factoring(SlidesApp.openById(PRESENTATION_ID)); }
function runS12Only()  { insertS12v2_milestone2(SlidesApp.openById(PRESENTATION_ID)); }

// ============================================================
// S8 画像差し替え用ファイルID
// ============================================================
// パワポ等で作図したKPIツリーPNGをGoogle Driveにアップ → ファイルIDをここに貼る
// ファイルIDの取り方: 共有リンク https://drive.google.com/file/d/{ここがID}/view
// 一時的に空文字でもエラーにならないようガードあり（空ならテキスト案内のみ表示）
var S8_KPI_TREE_IMAGE_ID = "";  // 例: "1AbCdEfGhIjKlMnOpQrStUvWxYz"

// ============================================================
// S7: 2026年度KPI（協賛配信受注 年間48件・2,400万円）
// ============================================================
function insertS7v2_kpiHeadline(pres) {
  var s = newSlide(pres, "2026年度 主要KPI（暫定目標）");

  // 結論メッセージ（テーブル上）— 強調キーワード「年間48件・2,400万円」だけ #FA006D
  drawConclusionMessage(s, 88,
    "協賛配信 受注：月平均4件 / 年間48件・売上2,400万円",
    "年間48件・売上2,400万円");

  // 大きいKPI数字を中央2カードで（#FA006D=件数 / 緑=売上）
  // カード1: 件数
  rect(s, 80, 145, 260, 130, C_WHITE, C_BORDER, 1);
  rect(s, 80, 145, 260, 5, C_ACCENT);  // 上ライン
  txt(s, "受注件数（暫定目標）", 90, 158, 240, 16, {
    size: 11, color: C_GRAY, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "48", 90, 184, 180, 56, {
    size: 48, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.END, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "件 / 年", 270, 200, 70, 28, {
    size: 14, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "月平均 4件", 90, 245, 240, 18, {
    size: 11, color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
  });

  // カード2: 売上
  rect(s, 380, 145, 260, 130, C_WHITE, C_BORDER, 1);
  rect(s, 380, 145, 260, 5, C_ACCENT2);
  txt(s, "売上換算（暫定目標）", 390, 158, 240, 16, {
    size: 11, color: C_GRAY, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "2,400", 390, 184, 180, 56, {
    size: 48, color: C_ACCENT2, bold: true,
    align: SlidesApp.ParagraphAlignment.END, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "万円", 570, 200, 70, 28, {
    size: 14, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "1件あたり50万円換算", 390, 245, 240, 18, {
    size: 11, color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 注釈
  txt(s, "※ 本目標は現時点の暫定値。本資料の競合ベンチマーク分析を踏まえ、最適なKPI体系を改めてご提案します。",
    25, 305, 671, 22, {
      size: 10, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S8: 売上構造の分解（KPIツリー）— 画像差し替え方式
// ============================================================
// 運用:
//  1. パワポ・Figma等でKPIツリーを作図 → PNGエクスポート（推奨: 1200x480px / 16:9互換）
//  2. Google DriveにアップロードしファイルIDを取得
//  3. ファイル冒頭の `S8_KPI_TREE_IMAGE_ID` に設定
//  4. runS8Only() で再生成
//
// 画像配置エリア: x=60, y=130, w=600, h=240（タイトル下〜下端余白）
// ※ 実際の画像縦横比に応じて自動調整される
function insertS8v2_revenueTree(pres) {
  var s = newSlide(pres, "売上構造の分解（KPIツリー）");

  // 結論メッセージ（タイトル直下）
  drawConclusionMessage(s, 88,
    "売上＝商談件数 × 受注率 × 受注単価。Fireworkは受注率と単価を押し上げる",
    "受注率と単価");

  if (S8_KPI_TREE_IMAGE_ID && S8_KPI_TREE_IMAGE_ID.length > 0) {
    try {
      var blob = DriveApp.getFileById(S8_KPI_TREE_IMAGE_ID).getBlob();
      var img = s.insertImage(blob);
      var areaX = 60, areaY = 130, areaW = 600, areaH = 240;
      var origW = img.getWidth();
      var origH = img.getHeight();
      var scale = Math.min(areaW / origW, areaH / origH);
      var newW = origW * scale;
      var newH = origH * scale;
      img.setWidth(newW).setHeight(newH);
      img.setLeft(areaX + (areaW - newW) / 2);
      img.setTop(areaY + (areaH - newH) / 2);
    } catch (e) {
      Logger.log("S8 画像読み込みエラー: " + e.message);
      _drawS8Placeholder(s, "画像読み込みに失敗しました: " + e.message);
    }
  } else {
    _drawS8Placeholder(s,
      "▼ パワポ等で作図したKPIツリー画像をここに配置\n" +
      "S8_KPI_TREE_IMAGE_ID にDriveのファイルIDを設定してください");
  }
}

/** S8用: 画像未設定時のプレースホルダー */
function _drawS8Placeholder(slide, message) {
  rect(slide, 60, 130, 600, 240, C_LIGHT, C_BORDER, 2);
  txt(slide, message || "KPIツリー画像をここに", 80, 220, 560, 60, {
    size: 12, color: C_GRAY, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
  });
}

// ============================================================
// S9: マイルストーン①：メディアパワー（4指標）
// ============================================================
function insertS9v2_milestone1(pres) {
  var s = newSlide(pres, "マイルストーン①：メディアパワー（全視聴者ベース）");

  // 4列テーブル: 指標 / 現状 / 2026年度目標 / ギャップ
  // 「2026年度目標」列を #FFF2CC でハイライト
  var colDefs = [
    { label: "指標",            align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "現状",            align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "2026年度目標",    align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "ギャップ",        align: SlidesApp.ParagraphAlignment.CENTER }
  ];
  var rows = [
    ["① ライブ視聴者数（リーチ）",       "5,084人",     "10,000人",   "+4,916人"],
    ["② 平均視聴分数（リテンション）",   "3.7分",       "7分",        "+3.3分"],
    ["③ 商品CTR（アクション）",          "1.13%",       "2.2%",       "+1.07pt"],
    ["④ コメント参加率（エンゲージ）",   "0.68%",       "1.0%",       "+0.32pt"]
  ];

  makeColHighlightTable(s, 25, 100, 671, colDefs, rows, {
    rowH: 36,
    hdrH: 30,
    bodySize: 11,
    highlightCol: 2  // 2026年度目標列をクリーム強調
  });

  // 注釈
  txt(s, "※ 直近3ヶ月平均=2026.1-3。視聴者数は最大値、CTR・コメント率は中央値ベース。",
    25, 295, 671, 16, {
      size: 9, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });

  // ポイントBOX（KPIツリー上の位置づけを再確認）
  drawTipsBox(s, 80, 320, 560, 65, [
    { text: "メディアパワー向上 = KPI（協賛配信受注）達成のための中間目標", conclusion: false },
    { text: "→ 受注率・単価を押し上げるためのコンテンツ・メディア強化に集中", conclusion: true }
  ]);
}

// ============================================================
// S10: コスメカテゴリ限定比較（マツキヨ × 競合4社・2026.1-3）
// マスター: 「コスメカテゴリに絞った場合の比較」
// ============================================================
function insertS10v2_benchmark(pres) {
  var s = newSlide(pres, "コスメカテゴリ限定比較（マツキヨ × 競合4社・2026.1-3）");

  // 結論メッセージ（強調キーワード「業界2位の高水準」）
  drawConclusionMessage(s, 88,
    "コスメ限定でも視聴者数は業界2位の高水準。エンゲージメント指標で競合と大差・伸び代大",
    "業界2位の高水準");

  // 7列テーブル（マスターHTML完全踏襲）
  var colDefs = [
    { label: "ブランド / セグメント",   align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "配信数",                  align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "平均\n配信時間",           align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "視聴者数\n(60分換算)",    align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "視聴分数\n(60分換算)",    align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "商品CTR\n(60分換算)",     align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "コメント率\n(60分換算)",  align: SlidesApp.ParagraphAlignment.CENTER }
  ];
  var rows = [
    ["マツキヨココカラライブ\nコスメ限定（2026.1-3）", "7回",  "61.8分", "5,228人",  "3.40分",  "1.24%",  "0.94%"],
    ["マツキヨココカラライブ\n全体（2026.1-3）",       "10回", "64.2分", "5,084人",  "3.1分",   "1.0%",   "0.8%"],
    ["A社（リテール）",                              "13回", "74.8分", "2,494人",  "24.3分",  "29.4%",  "33.4%"],
    ["B社(コスメブランド)",                          "23回", "81.1分", "6,958人",  "11.8分",  "8.2%",   "5.1%"],
    ["C社(コスメブランド)",                          "4回",  "63.0分", "1,458人",  "28.2分",  "11.7%",  "9.1%"],
    ["D社(メーカー)",                                "3回",  "51.5分", "871人",    "9.9分",   "5.3%",   "4.4%"]
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

// ============================================================
// S10b: 競合比較：他社の強み（4観点）
// マスター: index.html L897-927「競合比較：他社の強み」
// ============================================================
function insertS10bv2_competitorStrength(pres) {
  var s = newSlide(pres, "競合比較：他社の強み");

  // 結論メッセージ（強調語「推し」と「自分ごと化」）
  drawConclusionMessage(s, 88,
    "他社は「推し」と「自分ごと化」で視聴者を捕まえている",
    "「推し」と「自分ごと化」");

  // 2列テーブル: 観点 / 他社の強み
  // マスター実測踏襲: ヘッダ紺・データ行 薄ピンク #FFF5F7
  var colDefs = [
    { label: "観点",         align: SlidesApp.ParagraphAlignment.CENTER, nameCol: true },
    { label: "他社の強み",   align: SlidesApp.ParagraphAlignment.START }
  ];
  var rows = [
    [
      "出演者",
      "SNSフォロワー数万〜16万超のBAが個人ファンを配信に連れてくる(B社)。\nベテランBA固定でブランドファンがコミュニティ化(C社)。\n→ 推しが紹介するから買いたい・コミュ取りたいという動機が強い。"
    ],
    [
      "購買誘導",
      "B社：視聴者限定500円クーポン＋限定性訴求でCTR 8.2%。\nC社：クーポン不使用、純粋なコンテンツ力でCTR 11.7%。"
    ],
    [
      "コンテンツ力",
      "A社：セルフチェック・原因解説など「自分ごと化」設計でCTR 49.4%・視聴24分。\nC社：クーポンなしで純粋なコンテンツ力でCTR 11.7%・視聴28分。"
    ],
    [
      "リーチ",
      "B社が視聴者数6,958人で首位。\nファンが多いInstagramで配信して認知を広げ、ECサイトでのライブ配信に流入させている。"
    ]
  ];

  // テーブル幅: 観点列 134(20%)・強み列 537(80%)。makeColHighlightTableが等幅なので個別矩形+テキストで実装
  // → だが既存ヘルパーで頑張る場合、列幅指定がないため個別実装する
  _drawCompetitorStrengthTable(s, 25, 110, 671, colDefs, rows);

  // 注釈
  txt(s, "※1: B社CTR出典は公式IR資料／※2: C社CTR出典は公式ブランドサイト2025発表値。",
    25, 360, 671, 16, {
      size: 9, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
}

/**
 * S10b専用: 2列固定テーブル(観点20% / 他社の強み80%)
 * makeColHighlightTable は等幅前提なので、列幅可変が必要なここだけ個別実装
 */
function _drawCompetitorStrengthTable(slide, x, y, w, colDefs, rows) {
  var col1W = Math.round(w * 0.20);   // 観点列 134pt
  var col2W = w - col1W;              // 強み列 537pt
  var hdrH = 28;
  var rowH = 56;  // 2-3行入る高さ
  var totalH = hdrH + rowH * rows.length;

  // ヘッダ行: 紺背景・白文字
  var tbl = slide.insertTable(rows.length + 1, 2, x, y, w, totalH);
  var c0 = tbl.getCell(0, 0);
  c0.getFill().setSolidFill(C_DARK);
  c0.getText().setText(colDefs[0].label);
  _styleCell(c0, 11, true, C_WHITE, SlidesApp.ParagraphAlignment.CENTER);

  var c1 = tbl.getCell(0, 1);
  c1.getFill().setSolidFill(C_ACCENT3);  // マスター #0F3460 紺
  c1.getText().setText(colDefs[1].label);
  _styleCell(c1, 11, true, C_WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // ボディ行: 観点列=白bold / 強み列=#FFF5F7 薄ピンク
  for (var r = 0; r < rows.length; r++) {
    var observePoint = tbl.getCell(r + 1, 0);
    observePoint.getFill().setSolidFill(C_WHITE);
    observePoint.getText().setText(rows[r][0]);
    _styleCell(observePoint, 12, true, C_DARK, SlidesApp.ParagraphAlignment.CENTER);

    var strength = tbl.getCell(r + 1, 1);
    strength.getFill().setSolidFill("#FFF5F7");
    strength.getText().setText(rows[r][1]);
    _styleCell(strength, 9.5, false, C_TEXT, SlidesApp.ParagraphAlignment.START);
  }

  // ※ 列幅をAPIで直接設定する手段は限定的だが、後でユーザーがUIで微調整可能
}

// ============================================================
// S11: メディアパワーの因数分解（3カード）
// ============================================================
function insertS11v2_factoring(pres) {
  var s = newSlide(pres, "メディアパワーの因数分解");

  // 結論メッセージ（中央）
  drawConclusionMessage(s, 88,
    "メディアパワー = 集める力 × 引き留める力 × 再訪する力",
    "再訪する力");

  // サブ説明
  txt(s, "（＝視聴者数 × エンゲージメント深度［視聴分数・CTR・コメント率］ × リピーター比率）",
    25, 118, 671, 18, {
      size: 10, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 3カード: 集める力 / 引き留める力 / 再訪する力
  var cardY = 155;
  var cardH = 165;
  var cardW = 215;
  var cards = [
    { x: 30,  label: "① 集める力",         status: "業界上位・頭打ち",        desc: "視聴者数 平均5,000人超を安定維持。\n競合と比較しても高水準で、\nここからの倍増は困難。",                                            color: C_ACCENT3 },  // 紺
    { x: 252, label: "② 引き留める力",     status: "競合と大差・伸び代大",     desc: "視聴分数・CTR・コメント率すべてで\n競合トップの約2〜11%水準。\nここを埋めればメディア価値を\n数倍化できる。",                color: C_ACCENT  },  // #FA006D
    { x: 474, label: "③ 再訪する力",       status: "②を底上げするレバー",     desc: "リピーターはCTR・コメント率・\n視聴分数が非リピーターの数倍。\n再訪する視聴者を増やすことが、\n②の平均値を引き上げる最短経路。", color: C_ACCENT2 }   // 緑
  ];

  cards.forEach(function(c) {
    rect(s, c.x, cardY, cardW, cardH, C_WHITE, C_BORDER, 1);
    rect(s, c.x, cardY, 4, cardH, c.color);
    txt(s, c.label, c.x + 14, cardY + 12, cardW - 24, 16, {
      size: 11, color: c.color, bold: true,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
    txt(s, c.status, c.x + 14, cardY + 36, cardW - 24, 22, {
      size: 14, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
    txt(s, c.desc, c.x + 14, cardY + 66, cardW - 24, 92, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.START
    });
  });

  // 因数分解の結論（下部・黄色強調BOX）
  rect(s, 50, 335, 620, 48, "#FFF8E1", "#F0C040", 1);
  txt(s, "因数分解の結論：リピーターを増やすことが、メディアパワー全体を底上げする最短経路",
    60, 348, 600, 24, {
      size: 13, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S12: マイルストーン②：リピーター数値（追加提案・3指標）
// ============================================================
function insertS12v2_milestone2(pres) {
  var s = newSlide(pres, "マイルストーン②：リピーター数値（追加提案）");

  // 結論メッセージ
  drawConclusionMessage(s, 88,
    "マイルストーン①を底上げするレバーとして、リピーター3指標を追加提案",
    "リピーター3指標を追加提案");

  // 3カード: A=リピーター比率 / B=リピーターCTR / C=リピーターコメント率
  var cardY = 130;
  var cardH = 195;
  var cardW = 215;
  var cards = [
    { x: 30,  no: "マイルストーン② A", name: "リピーター比率",       now: "4.88%", target: "8.0%",  multiple: "—",                                  note: "自然成長(+約2pt/年)に\n出演者SNS×推し再訪施策で\n達成見込み", color: C_ACCENT  },
    { x: 252, no: "マイルストーン② B", name: "リピーターCTR",        now: "3.87%", target: "8.0%",  multiple: "全視聴者CTRの約3.7倍",              note: "限定特典＋ファン育成で\n「この人がすすめるなら買いたい」\n層を増やす",  color: "#E07A5F" },
    { x: 474, no: "マイルストーン② C", name: "リピーターコメント率", now: "3.78%", target: "7.5%",  multiple: "全視聴者コメント率の約5.3倍",      note: "出演者がInstagram等で\n話しかける配信外コミュニティ\n形成を新施策追加",  color: C_ACCENT2 }
  ];

  cards.forEach(function(c) {
    rect(s, c.x, cardY, cardW, cardH, C_WHITE, C_BORDER, 1);
    rect(s, c.x, cardY, cardW, 3, c.color);  // 上ボーダー

    txt(s, c.no, c.x + 12, cardY + 8, cardW - 24, 14, {
      size: 9, color: C_GRAY, bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
    txt(s, c.name, c.x + 12, cardY + 26, cardW - 24, 18, {
      size: 12, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
    txt(s, "現状", c.x + 12, cardY + 50, cardW - 24, 14, {
      size: 9, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
    txt(s, c.now, c.x + 12, cardY + 64, cardW - 24, 26, {
      size: 18, color: "#555555", bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
    txt(s, "2026年度目標", c.x + 12, cardY + 92, cardW - 24, 14, {
      size: 9, color: C_ACCENT, bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
    txt(s, c.target, c.x + 12, cardY + 106, cardW - 24, 32, {
      size: 24, color: C_ACCENT, bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
    if (c.multiple !== "—") {
      txt(s, c.multiple, c.x + 12, cardY + 138, cardW - 24, 18, {
        size: 8, color: C_GRAY,
        align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
      });
    }
    txt(s, c.note, c.x + 12, cardY + 158, cardW - 24, 32, {
      size: 8, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
  });

  // 下部：論拠注釈（緑bg）
  rect(s, 30, 340, 660, 42, "#F0FAF8", C_ACCENT2, 1);
  txt(s, "マイルストーン②達成 → マイルストーン①底上げ → 最終的にKPI（協賛配信受注）達成 に至る最短経路",
    40, 350, 640, 22, {
      size: 12, color: C_ACCENT2, bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
}
