/**
 * MCCM提案スライド化 - §振り返り: 2025年度の振り返り（1スライド）
 *
 * 最終更新: 2026-04-29 20:03
 *
 * マスター情報源: https://riderkarubo.github.io/pages/mccm-proposal/
 *  （これ以外の場所からは情報を拾わない）
 *
 * Issy手直し解析パターン準拠:
 *  - タイトル黒帯: x=9, y=20, w=671, h=45（newSlide()自動配置）
 *  - 強調色:
 *    Highlight: C_ACCENT2 (#1B998B 緑)
 *    Lowlight : C_ACCENT  (#FA006D マゼンタ)
 *    学び     : C_ACCENT3 (#0F3460 紺)
 *
 * 構成（1スライド完結・上段2カラム＋下段1カラム）:
 *  Highlight 横並び ＋ Lowlight 横並び（上段）
 *  学び 幅広（下段）
 *
 * 80%カット原則:
 *  - HTMLのテーブル省略・注釈削除
 *  - 各ブロック2要素以内
 *  - 学びは「結論メッセージ」だけ残す
 *
 * 実行手順:
 *  1. 02_helpers.gs を最新版で貼り付け済み
 *  2. このファイルを Apps Script に新規追加
 *  3. insertSectionReview() を実行 → 末尾に1枚追加
 */

function insertSectionReview() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  Logger.log("§振り返り 開始 - 現在: " + pres.getSlides().length + "枚");

  insertSReview_yearReview(pres);
  Logger.log("S振り返り 完了");

  Logger.log("§振り返り 完了 - 現在: " + pres.getSlides().length + "枚");
}

function runReviewOnly() { insertSReview_yearReview(SlidesApp.openById(PRESENTATION_ID)); }

// ============================================================
// S振り返り: 2025年度の振り返り（1スライド）
// ============================================================
function insertSReview_yearReview(pres) {
  var s = newSlide(pres, "2025年度の振り返り");

  // ============================================================
  // サブタイトル（小見出し領域・y=76, h=34）
  // ============================================================
  txt(s, 'ハイライト・ローライト・学び',
      SUBTITLE_X, SUBTITLE_Y, SUBTITLE_W, SUBTITLE_H, {
    size: 12,
    color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // ============================================================
  // レイアウト計算（BODY領域: y=110, h=264）
  // ============================================================
  // 上段: Highlight + Lowlight 横並び
  var topY = 115;
  var topH = 130;
  var sideMargin = 25;
  var gap = 12;
  var topCardW = (BODY_W - gap) / 2;  // (671 - 12) / 2 = 329.5
  var hlX = sideMargin;
  var llX = sideMargin + topCardW + gap;

  // 下段: 学び 幅広
  var bottomY = topY + topH + 12;  // 257
  var bottomH = 110;
  var bottomX = sideMargin;
  var bottomW = BODY_W;

  // ============================================================
  // Highlight ブロック（緑 #1B998B）
  // ============================================================
  drawReviewBlock(s, hlX, topY, topCardW, topH, {
    label: "🎯 Highlight",
    accent: C_ACCENT2,
    bg: "#F0FAF8",
    items: [
      { tag: "効率性", tagBg: "#E0F5F0", tagColor: C_ACCENT2,
        main: "3名体制で毎週配信を実施",
        sub: "他社は5〜7名体制が標準" },
      { tag: "実施頻度", tagBg: "#FFF3CD", tagColor: "#856404",
        main: "協賛配信を毎週受注",
        sub: "リテール業界でもトップクラスの実施頻度" }
    ]
  });

  // ============================================================
  // Lowlight ブロック（マゼンタ #FA006D）
  // ============================================================
  drawReviewBlock(s, llX, topY, topCardW, topH, {
    label: "⚠️ Lowlight",
    accent: C_ACCENT,
    bg: "#FFF5F7",
    items: [
      { main: "視聴者数・コメント率は順調に成長",
        sub: null, positive: true },
      { main: "平均視聴分数・商品CTR は伸び悩み",
        sub: "リピーター率 5.0% と低水準" }
    ]
  });

  // ============================================================
  // 学び ブロック（紺 #0F3460・幅広）
  // ============================================================
  drawLearnBlock(s, bottomX, bottomY, bottomW, bottomH);

  // ============================================================
  // 出典（右下小さく）
  // ============================================================
  txt(s, '出典: 配信数値CSV（2025年度・中央値ベース）',
      sideMargin, 380, BODY_W, 12, {
    size: 8,
    color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.END,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
}

// ============================================================
// ブロック描画ヘルパー（Highlight / Lowlight 共通）
// ============================================================
function drawReviewBlock(s, x, y, w, h, cfg) {
  // 背景矩形（左に4pxアクセントバー風の左ボーダー）
  rect(s, x, y, w, h, cfg.bg);
  rect(s, x, y, 4, h, cfg.accent);  // 左ボーダー

  // ラベル（見出し）
  txt(s, cfg.label, x + 14, y + 8, w - 18, 22, {
    size: 13,
    color: C_DARK,
    bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // アイテム（最大2件）
  var itemY = y + 36;
  var itemH = (h - 40) / cfg.items.length - 2;

  cfg.items.forEach(function(item, i) {
    var iy = itemY + i * (itemH + 2);

    // タグ（あれば）
    var mainX = x + 14;
    if (item.tag) {
      var tagW = item.tag.length * 9 + 12;  // 推定幅
      rectTxt(s, item.tag, mainX, iy + 2, tagW, 14, item.tagBg, {
        size: 8,
        color: item.tagColor,
        bold: true,
        align: SlidesApp.ParagraphAlignment.CENTER,
        va: SlidesApp.ContentAlignment.MIDDLE
      });
      mainX += tagW + 4;
    }

    // メインテキスト
    var mainColor = item.positive ? C_ACCENT2 : C_DARK;
    txt(s, item.main, mainX, iy, w - (mainX - x) - 10, 16, {
      size: 11,
      color: mainColor,
      bold: true,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

    // サブテキスト（あれば）
    if (item.sub) {
      txt(s, item.sub, x + 14, iy + 18, w - 24, 14, {
        size: 8,
        color: C_GRAY,
        align: SlidesApp.ParagraphAlignment.START,
        va: SlidesApp.ContentAlignment.MIDDLE
      });
    }
  });
}

// ============================================================
// 学びブロック描画（幅広・1メッセージ）
// ============================================================
function drawLearnBlock(s, x, y, w, h) {
  // 背景矩形（紺左ボーダー）
  rect(s, x, y, w, h, "#F0F4FF");
  rect(s, x, y, 4, h, C_ACCENT3);  // 左ボーダー紺

  // ラベル
  txt(s, "💡 学び", x + 14, y + 8, w - 18, 22, {
    size: 13,
    color: C_DARK,
    bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // メッセージ本体（プレフィックス + 強調 + サフィックス を分けて描画）
  // インライン色変更が複雑なので、メッセージは1要素で書きつつ後でハイライト範囲だけ色変更

  var msgY = y + 38;
  var msgX = x + 18;
  var msgW = w - 36;
  var msgH = h - 46;
  var fullMsg = "平均視聴分数・商品CTR が伸び悩み。これまで企画・施策で改善を試みてきたが、ここから抜本的な底上げを期待するのは難しい。\n→ 熱量の高いリピーター・ファンを増やすことが、全体のエンゲージメントを底上げする ＝ 次年度の最重要テーマ";

  var box = s.insertTextBox(fullMsg, msgX, msgY, msgW, msgH);
  box.getFill().setTransparent();
  box.getBorder().setTransparent();
  box.setContentAlignment(SlidesApp.ContentAlignment.TOP);

  var tr = box.getText();
  tr.getParagraphs().forEach(function(p) {
    var st = p.getRange().getTextStyle();
    st.setFontFamily("Noto Sans JP").setFontSize(11).setForegroundColor(C_DARK);
    p.getRange().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.START)
      .setLineSpacing(140);
  });

  // 強調部分1: 「→」と「熱量の高い〜底上げする」を C_ACCENT 太字
  var emphasizeStart = fullMsg.indexOf("熱量の高いリピーター・ファンを増やすことが、全体のエンゲージメントを底上げする");
  var emphasizeText = "熱量の高いリピーター・ファンを増やすことが、全体のエンゲージメントを底上げする";
  if (emphasizeStart >= 0) {
    try {
      var range = tr.getRange(emphasizeStart, emphasizeStart + emphasizeText.length);
      range.getTextStyle().setForegroundColor(C_ACCENT).setBold(true);
    } catch(e) { Logger.log("学びブロック強調1失敗: " + e); }
  }

  // 強調部分2: 「→」も色変更
  var arrowIdx = fullMsg.indexOf("→");
  if (arrowIdx >= 0) {
    try {
      var arrowRange = tr.getRange(arrowIdx, arrowIdx + 1);
      arrowRange.getTextStyle().setForegroundColor(C_ACCENT).setBold(true);
    } catch(e) {}
  }
}
