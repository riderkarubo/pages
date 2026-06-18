/**
 * Firework標準スライドテンプレート用ヘルパー
 *
 * 最終更新: 2026年5月1日 21:02
 *
 * 用途:
 *  最終資料 (株式会社MCCマネジメント御中_年次ビジネスレビュー_20260430)
 *  P2-15, P41-50 のデザインを完全再現するためのヘルパー関数群。
 *  接頭辞 "tpl_" で既存の 02_helpers.gs と衝突しないようにしてある。
 *
 * 使い方:
 *  1. 02_helpers.gs を先に Apps Script に貼り付ける（基盤ヘルパー: txt, rect, ellipse, makeTable など）
 *  2. このファイルを追加で貼り付ける
 *  3. 各テンプレ関数を呼び出してスライドを生成する
 *
 * 設計原則:
 *  - すべての座標は実測値の絶対座標（pt 単位、16:9 = 720x405）
 *  - 色は中央集約された定数を介して呼び出す（ブランド差し替えが容易）
 *  - 1関数 = 1スライド（または再利用可能な1パターン）
 *
 * 仕様参照:
 *  learned_layouts/design_patterns.md 参照
 */

// ============================================================
// 追加カラー定数（design_patterns.md 6章で抽出）
// ============================================================
var C_TIMELINE_BG       = "#FFFCF3";  // P4 タイムライン全体ベース
var C_BG_PINK           = "#FFF5F7";  // マゼンタ系カード背景
var C_BG_PINK_ROW       = "#FCE7F3";  // テーブル強調行ピンク
var C_BG_PINK_BADGE     = "#FFE0EC";  // カテゴリバッジピンク
var C_BG_AMBER_BADGE    = "#FFF3CD";  // カテゴリバッジ薄黄
var C_BG_GREEN          = "#F0FAF8";  // 緑カード背景（学びBOX）
var C_BG_NAVY           = "#F0F4FF";  // 紺Lowlight背景
var C_BG_LIGHT          = "#F8F9FA";  // 薄背景（課題BOX）
var C_BG_TABLE_ALT      = "#F8F8F8";  // テーブル偶数行ゼブラ
var C_TEXT_DARK         = "#1E293B";  // 強調本文
var C_TEXT_GRAY2        = "#475569";  // カード内サブ
var C_LABEL_GRAY        = "#64748B";  // 小ラベル「資産01」
var C_ORANGE            = "#E07A5F";  // 課題❸・リピーターCTRバッジ
var C_INDIGO            = "#6366F1";  // 課題❹
var C_SKY               = "#06B6D4";  // 資産07 LIPSコラボ
var C_AMBER             = "#F59E0B";  // 資産02 91回
var C_AMBER2            = "#EAB308";  // 資産06 計10名
var C_PURPLE            = "#8B5CF6";  // 資産05 ノウハウ
var C_BROWN             = "#C07000";  // 「おトク」戦略
var C_BLUE_LINK         = "#3C78D8";  // 課題BOXタイトル帯

// 系統A（ベタ黒帯）用タイトル座標
var TITLE_X_SOLID = 24, TITLE_Y_SOLID = 19, TITLE_W_SOLID = 671, TITLE_H_SOLID = 45;

// ============================================================
// タイトル系（系統A：ベタ黒帯）
// ============================================================
/**
 * 系統A: ベタ黒帯タイトル
 *  TITLE_AND_BODY_2_1 レイアウトで使用
 *  P2,3,5,6,7,8,10,11,12,13,14,15,42,45,48,49 と同じ装飾
 */
function tpl_titleSolid(slide, titleText, opts) {
  opts = opts || {};
  rect(slide, TITLE_X_SOLID, TITLE_Y_SOLID, TITLE_W_SOLID, TITLE_H_SOLID, "#000000");
  txt(slide, titleText, TITLE_X_SOLID + 14, TITLE_Y_SOLID, TITLE_W_SOLID - 28, TITLE_H_SOLID, {
    size: opts.size || 18,
    bold: true,
    color: C_WHITE,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
}

/**
 * 系統A タイトル + 部分強調インライン（カッコ内をグレーで）
 *  例: tpl_titleSolidInline(slide, "年度別 協賛実績の推移", "（MCCM様よりご提供）")
 */
function tpl_titleSolidInline(slide, mainText, subText) {
  rect(slide, TITLE_X_SOLID, TITLE_Y_SOLID, TITLE_W_SOLID, TITLE_H_SOLID, "#000000");
  var box = slide.insertTextBox(mainText + (subText || ""), TITLE_X_SOLID + 14, TITLE_Y_SOLID, TITLE_W_SOLID - 28, TITLE_H_SOLID);
  box.getFill().setTransparent();
  box.getBorder().setTransparent();
  var tr = box.getText();
  // メイン部
  tr.getRange(0, mainText.length).getTextStyle()
    .setFontFamily("Noto Sans JP").setFontSize(18).setBold(true).setForegroundColor(C_WHITE);
  if (subText) {
    tr.getRange(mainText.length, mainText.length + subText.length).getTextStyle()
      .setFontFamily("Noto Sans JP").setFontSize(12).setBold(false).setForegroundColor(C_WHITE);
  }
  tr.getParagraphs()[0].getRange().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
}

// ============================================================
// アジェンダリスト（Pattern A）
// ============================================================
/**
 * アジェンダ箇条書きスライドを構築
 * @param {string[]} items 項目配列
 * @param {number} highlightIdx -1=なし, それ以外は当該indexをBoldに
 */
function tpl_agendaList(slide, items, highlightIdx) {
  if (highlightIdx === undefined) highlightIdx = -1;
  // 黒帯タイトル
  tpl_titleSolid(slide, "アジェンダ", { size: 20 });
  // 本文ボックス
  var box = slide.insertTextBox(items.join("\n"), 34, 122, 651, 211);
  box.getFill().setTransparent();
  box.getBorder().setTransparent();
  var paragraphs = box.getText().getParagraphs();
  paragraphs.forEach(function(p, i) {
    var st = p.getRange().getTextStyle();
    st.setFontFamily("Noto Sans JP")
      .setFontSize(16)
      .setForegroundColor(C_TEXT_GRAY)
      .setBold(i === highlightIdx);
    p.getRange().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.START)
      .setLineSpacing(150);
  });
}

// ============================================================
// タイムライン（Pattern B）
// ============================================================
/**
 * P4タイムラインを完全再現
 * @param {Array<{date:string, desc:string, isMilestone?:boolean}>} entries 9項目想定
 */
function tpl_timeline(slide, entries) {
  // ベース薄ベージュ
  rect(slide, 40, 83, 644, 291, C_TIMELINE_BG);
  // 縦線
  rect(slide, 156, 98, 2, 256, "#DDDDDD");
  // 各項目
  var startY = 90;
  var pitch = 32;
  entries.forEach(function(e, i) {
    var y = startY + i * pitch;
    // 日付
    txt(slide, e.date, 30, y, 110, 20, {
      size: 11, color: C_TEXT_GRAY, align: SlidesApp.ParagraphAlignment.START
    });
    // ドット（マイルストーン=マゼンタ系THEME, 過去=グレー）
    var dotColor = e.isMilestone === false ? "#666666" : C_ACCENT;
    ellipse(slide, 152, y + 4, 10, 10, dotColor);
    // 内容
    txt(slide, e.desc, 175, y, 520, 22, {
      size: 11, color: C_TEXT_DARK, align: SlidesApp.ParagraphAlignment.START
    });
  });
}

// ============================================================
// KPIカード4分割（Pattern C）
// ============================================================
/**
 * P5の「数字でみる実績」4カード配置
 * @param {Array<{label, value, unit, note, accent?: color}>} cards 4枚
 *   accent: 縦バー&数値色（省略時は黒 #1A1A2E）
 */
function tpl_kpiQuad(slide, cards) {
  var positions = [
    { x: 19,  y: 112 }, { x: 360, y: 112 },
    { x: 19,  y: 224 }, { x: 360, y: 224 }
  ];
  cards.forEach(function(c, i) {
    var p = positions[i];
    var accent = c.accent || C_DARK;
    // 背景白カード
    rect(slide, p.x, p.y, 341, 113, C_WHITE);
    // 縦バー
    rect(slide, p.x + 6, p.y + 4, 4, 101, accent);
    // ラベル（上）
    txt(slide, c.label, p.x + 19, p.y + 11, 306, 18, {
      size: 10, bold: true, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START
    });
    // KPI数値（中）— 主数値+単位の2要素を1ボックスでインライン化
    var valueBox = slide.insertTextBox(c.value + (c.unit || ""), p.x + 19, p.y + 30, 188, 45);
    valueBox.getFill().setTransparent();
    valueBox.getBorder().setTransparent();
    var tr = valueBox.getText();
    tr.getRange(0, c.value.length).getTextStyle()
      .setFontFamily("Noto Sans JP").setFontSize(36).setBold(true).setForegroundColor(accent);
    if (c.unit) {
      tr.getRange(c.value.length, c.value.length + c.unit.length).getTextStyle()
        .setFontFamily("Noto Sans JP").setFontSize(14).setBold(true).setForegroundColor(accent);
    }
    tr.getParagraphs()[0].getRange().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    // 注釈（下）
    txt(slide, c.note, p.x + 19, p.y + 82, 306, 15, {
      size: 9, bold: true, color: C_TEXT_DARK,
      align: SlidesApp.ParagraphAlignment.START
    });
  });
}

// ============================================================
// 資産カードグリッド（Pattern D）
// ============================================================
/**
 * P7の7枚カードレイアウトを再現
 * @param {Array<{no, category, kpi, kpiSize, sub, accent}>} assets 7枚
 *   accent: 縦バー色 + KPIテキスト色
 *   kpiSize: 18 (大KPI) or 14 (中KPI) or 16 (固有名詞)
 */
function tpl_assetGrid7(slide, assets) {
  var positions = [
    { x: 24,  y: 126 }, { x: 193, y: 126 }, { x: 362, y: 126 }, { x: 530, y: 126 },
    { x: 109, y: 255 }, { x: 278, y: 255 }, { x: 446, y: 255 }
  ];
  assets.forEach(function(a, i) {
    var p = positions[i];
    // ROUND_RECTANGLE 白カード
    var card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, p.x, p.y, 161, 120);
    card.getFill().setSolidFill(C_WHITE);
    card.getBorder().setTransparent();
    // 縦バー
    rect(slide, p.x, p.y + 12, 3, 30, a.accent);
    // 4レイヤーテキスト（1ボックスにまとめてインライン色）
    var tx = "資産 " + a.no + "\n" + a.category + "\n" + a.kpi + "\n" + (a.sub || "");
    var tb = slide.insertTextBox(tx, p.x + 10, p.y, 150, 120);
    tb.getFill().setTransparent();
    tb.getBorder().setTransparent();
    var paras = tb.getText().getParagraphs();
    if (paras[0]) paras[0].getRange().getTextStyle()
      .setFontFamily("Noto Sans JP").setFontSize(8).setBold(true).setForegroundColor(C_LABEL_GRAY);
    if (paras[1]) paras[1].getRange().getTextStyle()
      .setFontFamily("Noto Sans JP").setFontSize(13).setBold(true).setForegroundColor(C_TEXT_DARK);
    if (paras[2]) paras[2].getRange().getTextStyle()
      .setFontFamily("Noto Sans JP").setFontSize(a.kpiSize || 18).setBold(true).setForegroundColor(a.accent);
    if (paras[3]) paras[3].getRange().getTextStyle()
      .setFontFamily("Noto Sans JP").setFontSize(9).setBold(false).setForegroundColor(C_TEXT_GRAY2);
  });
}

// ============================================================
// Highlight/Lowlight/学び（Pattern E）
// ============================================================
/**
 * P9の振り返りBOXを完全再現
 * @param {{
 *   highlight: { title, badges: Array<{label, bg, color}>, items: string[], note },
 *   lowlight: { title, body },
 *   learning: { title, body, emphasis }
 * }} content
 */
function tpl_reviewTriBox(slide, content) {
  var H = content.highlight, L = content.lowlight, R = content.learning;
  // Highlight 左
  rect(slide, 25, 102, 330, 130, C_BG_PINK);
  rect(slide, 25, 102, 4, 130, C_ACCENT);
  txt(slide, "🎯 Highlight", 39, 110, 312, 22, {
    size: 13, bold: true, color: C_DARK
  });
  // Highlight badges
  var by = 138;
  (H.badges || []).forEach(function(badge, i) {
    rect(slide, 39, by, 46, 14, badge.bg || C_BG_PINK_BADGE);
    txt(slide, badge.label, 39, by, 46, 14, {
      size: 9, bold: true, color: badge.color || C_ACCENT,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
    txt(slide, (H.items || [])[i] || "", 86, by - 2, 263, 16, {
      size: 11, bold: true, color: C_DARK
    });
    by += 45;
  });
  if (H.note) {
    txt(slide, H.note, 39, 158, 306, 18, {
      size: 9, color: C_GRAY
    });
  }
  // Lowlight 右
  rect(slide, 366, 102, 330, 130, C_BG_NAVY);
  rect(slide, 366, 102, 4, 130, C_ACCENT3);
  txt(slide, "⚠️ Lowlight", 380, 110, 312, 22, {
    size: 13, bold: true, color: C_ACCENT3
  });
  txt(slide, L.body, 380, 138, 306, 75, {
    size: 11, color: C_DARK
  });
  // 学び 下
  rect(slide, 25, 244, 671, 110, C_BG_GREEN);
  rect(slide, 25, 244, 4, 110, C_ACCENT2);
  txt(slide, "💡 学び", 39, 252, 653, 22, {
    size: 14, bold: true, color: C_ACCENT2
  });
  // 学び本文（emphasisがあれば矢印強調）
  if (R.emphasis) {
    // 2段組: 上に通常本文, 下に強調
    txt(slide, R.body, 46, 270, 653, 30, {
      size: 12, color: C_DARK
    });
    txt(slide, "→ " + R.emphasis, 46, 305, 653, 22, {
      size: 14, bold: true, color: C_ACCENT
    });
    txt(slide, R.tail || "", 46, 330, 653, 18, {
      size: 12, color: C_DARK
    });
  } else {
    txt(slide, R.body, 46, 281, 653, 75, {
      size: 12, color: C_DARK
    });
  }
}

// ============================================================
// 4課題カードグリッド + 最重点バッジ（Pattern F）
// ============================================================
/**
 * P14の4カード配置を再現
 * @param {Array<{title, source, cause, accent, isPriority?:boolean}>} cards 4枚
 */
function tpl_issueQuad(slide, cards) {
  var positions = [
    { x: 16,  y: 142 }, { x: 362, y: 142 },
    { x: 16,  y: 249 }, { x: 362, y: 249 }
  ];
  cards.forEach(function(c, i) {
    var p = positions[i];
    // ROUND_RECTANGLE 背景
    var card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, p.x, p.y, 332, 77);
    card.getFill().setSolidFill(c.isPriority ? C_BG_PINK_ROW : C_WHITE);
    card.getBorder().setTransparent();
    // 底部1pt横線
    rect(slide, p.x, p.y + 75, 332, 1, "#E5E7EB");
    // 縦バー
    rect(slide, p.x, p.y, 3, 77, c.accent);
    // タイトル行
    txt(slide, c.title, p.x + 13, p.y + 9, 318, 20, {
      size: 11, bold: true, color: c.accent
    });
    // 本文行（数値: 〇〇 原因: 〇〇 をインライン）
    var bodyText = "数値：" + c.source + " 原因：" + c.cause;
    var tb = slide.insertTextBox(bodyText, p.x + 13, p.y + 28, 318, 44);
    tb.getFill().setTransparent();
    tb.getBorder().setTransparent();
    var paras = tb.getText().getParagraphs();
    paras.forEach(function(pp) {
      pp.getRange().getTextStyle()
        .setFontFamily("Noto Sans JP").setFontSize(9).setForegroundColor(C_TEXT);
    });
    // 「数値：」「原因：」だけBold #666666 にする
    var labels = ["数値：", "原因："];
    labels.forEach(function(lab) {
      var idx = bodyText.indexOf(lab);
      if (idx >= 0) {
        tb.getText().getRange(idx, idx + lab.length).getTextStyle()
          .setBold(true).setForegroundColor(C_GRAY);
      }
    });
  });
  // 最重点バッジ（指定があれば最重点カードの右上に配置）
  var priIdx = cards.findIndex(function(c) { return c.isPriority; });
  if (priIdx >= 0) {
    var p = positions[priIdx];
    rect(slide, p.x, p.y - 28, 118, 28, C_ACCENT);
    txt(slide, "最重点ポイント", p.x, p.y - 28, 118, 28, {
      size: 11, bold: true, color: C_WHITE,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
  }
}

// ============================================================
// 大KPIメッセージ + 縦バーカード2分割（Pattern H）
// ============================================================
/**
 * P41の「年間48件・売上2,400万円」スライドを再現
 * @param {Object} content
 *   content.message  = タイトル下大文字メッセージ + 強調インライン
 *   content.cards    = [{header, body, accent}] 2枚
 *   content.footer   = フッター進捗メッセージ
 */
function tpl_targetMsgWith2Cards(slide, content) {
  var m = content.message;
  var msgBox = slide.insertTextBox(m.prefix + m.emphasis + (m.suffix || ""), 34, 98, 651, 30);
  msgBox.getFill().setTransparent();
  msgBox.getBorder().setTransparent();
  var tr = msgBox.getText();
  tr.getRange(0, m.prefix.length).getTextStyle()
    .setFontFamily("Noto Sans JP").setFontSize(20).setBold(true).setForegroundColor(C_TEXT_DARK);
  tr.getRange(m.prefix.length, m.prefix.length + m.emphasis.length).getTextStyle()
    .setFontFamily("Noto Sans JP").setFontSize(20).setBold(true).setForegroundColor(C_ACCENT);
  if (m.suffix) {
    var suffStart = m.prefix.length + m.emphasis.length;
    tr.getRange(suffStart, suffStart + m.suffix.length).getTextStyle()
      .setFontFamily("Noto Sans JP").setFontSize(20).setBold(true).setForegroundColor(C_TEXT_DARK);
  }
  tr.getParagraphs()[0].getRange().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // 2分割カード
  var positions = [{ x: 30, y: 158 }, { x: 367, y: 158 }];
  content.cards.forEach(function(c, i) {
    var p = positions[i];
    var card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, p.x, p.y, 322, 180);
    card.getFill().setSolidFill(C_WHITE);
    card.getBorder().setTransparent();
    // 上端帯
    var top = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, p.x, p.y, 322, 6);
    top.getFill().setSolidFill(c.accent);
    top.getBorder().setTransparent();
    // 内テキスト（multi-line）
    txt(slide, c.header + "\n" + c.body, p.x, p.y + 6, 322, 174, {
      size: 14, bold: true, color: C_TEXT_DARK,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
  });

  // フッターメッセージ
  if (content.footer) {
    txt(slide, content.footer, 30, 349, 651, 30, {
      size: 20, bold: true, color: C_TEXT_DARK,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
  }
}

// ============================================================
// 3カラム戦略カード（Pattern I）
// ============================================================
/**
 * P44のリピーター戦略3カードを再現
 * @param {string} subtitle 副タイトル「視聴者が"また観たくなる理由"をつくる」等
 * @param {Array<{title, body, accent, iconUrl?}>} strategies 3枚
 * @param {string} footer ※注釈
 */
function tpl_strategyTri(slide, subtitle, strategies, footer) {
  // 副タイトル
  txt(slide, subtitle, 25, 79, 671, 34, {
    size: 16, bold: true, color: C_DARK,
    align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
  });
  // 3カラム
  var positions = [{ x: 25, y: 128 }, { x: 252, y: 128 }, { x: 479, y: 128 }];
  strategies.forEach(function(s, i) {
    var p = positions[i];
    // 全体白カード
    var card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, p.x, p.y, 215, 210);
    card.getFill().setSolidFill(C_WHITE);
    card.getBorder().setTransparent();
    // 上ヘッダ帯（ラウンド）
    var hdr = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, p.x, p.y, 215, 45);
    hdr.getFill().setSolidFill(s.accent);
    hdr.getBorder().setTransparent();
    // ヘッダ下細帯（直角）
    rect(slide, p.x, p.y + 30, 215, 15, s.accent);
    // ヘッダテキスト
    txt(slide, s.title, p.x + 7, p.y, 200, 45, {
      size: 14, bold: true, color: C_WHITE,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
    // 本文
    txt(slide, s.body, p.x + 11, p.y + 60, 193, 68, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START
    });
  });
  if (footer) {
    txt(slide, footer, 24, 345, 671, 34, {
      size: 12, color: C_DARK
    });
  }
}

// ============================================================
// 2カラム課題BOX（Pattern J）
// ============================================================
/**
 * P46/P47の現状課題2BOXを再現
 * @param {Array<{header, body}>} boxes 2枚
 * @param {string} accent ヘッダ帯色（既定 C_BLUE_LINK）
 */
function tpl_issue2Boxes(slide, boxes, accent) {
  accent = accent || C_BLUE_LINK;
  var positions = [{ x: 30, y: 124 }, { x: 367, y: 124 }];
  boxes.forEach(function(b, i) {
    var p = positions[i];
    // 全体BOX
    var card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, p.x, p.y, 323, 195);
    card.getFill().setSolidFill(C_BG_LIGHT);
    card.getBorder().setTransparent();
    // ヘッダ帯（ラウンド）
    var hdr = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, p.x, p.y, 323, 27);
    hdr.getFill().setSolidFill(accent);
    hdr.getBorder().setTransparent();
    // ヘッダ下細帯
    rect(slide, p.x, p.y + 16, 323, 10, accent);
    // ヘッダテキスト
    txt(slide, b.header, p.x + 15, p.y, 293, 27, {
      size: 14, bold: true, color: C_DARK,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
    // 本文
    txt(slide, b.body, p.x + 15, p.y + 40, 293, 150, {
      size: 11, color: C_TEXT_GRAY,
      align: SlidesApp.ParagraphAlignment.START
    });
  });
}

// ============================================================
// 結論メッセージ（タイトル下リード文）
// ============================================================
/**
 * 「→ 結論メッセージ」スタイル（系統A・系統B共通）
 * @param {string} text  テキスト
 * @param {string} emphasis 強調インライン部分（任意）
 * @param {number} y 縦位置（系統B=98, 系統A=86 が多い）
 */
function tpl_leadMessage(slide, text, emphasis, y) {
  y = y || 95;
  var box = slide.insertTextBox(text, 34, y, 651, 30);
  box.getFill().setTransparent();
  box.getBorder().setTransparent();
  var tr = box.getText();
  tr.getParagraphs().forEach(function(p) {
    p.getRange().getTextStyle()
      .setFontFamily("Noto Sans JP").setFontSize(15).setBold(true).setForegroundColor(C_TEXT_GRAY);
    p.getRange().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  });
  if (emphasis) {
    var idx = text.indexOf(emphasis);
    if (idx >= 0) {
      tr.getRange(idx, idx + emphasis.length).getTextStyle()
        .setForegroundColor(C_ACCENT);
    }
  }
}

/**
 * 注釈※印（標準8pt #666666）
 */
function tpl_note(slide, text, x, y, w) {
  txt(slide, text, x, y, w || 671, 16, {
    size: 8, color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.START
  });
}

// ============================================================
// テーブル系（Pattern G の各バリアント）
// ============================================================
/**
 * KPI比較テーブル（P42/P45相当）
 * @param {Array<{label, current, target, gap}>} rows
 * @param {string[]} headers ["指標","現状","2026年度目標","ギャップ"] 等
 * @param {{x,y,w}} pos
 */
function tpl_kpiCompareTable(slide, rows, headers, pos) {
  pos = pos || { x: 57, y: 110, w: 605 };
  var nCols = headers.length;
  var colWidths = [];
  if (nCols === 4) colWidths = [186, 148, 167, 101];
  else if (nCols === 3) colWidths = [203, 203, 213];
  var rowH = 35;
  var startX = pos.x, startY = pos.y;
  // ヘッダ
  var x = startX;
  headers.forEach(function(h, ci) {
    var w = colWidths[ci];
    var fillColor = (ci === 2) ? C_ACCENT3 : C_DARK;  // 目標列だけアクセント色
    rect(slide, x, startY, w, rowH, fillColor);
    txt(slide, h, x, startY, w, rowH, {
      size: 10, bold: true, color: C_WHITE,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
    x += w;
  });
  // データ行
  rows.forEach(function(r, ri) {
    var y = startY + rowH * (ri + 1);
    var bg = (ri % 2 === 1) ? C_BG_TABLE_ALT : C_WHITE;
    var x = startX;
    var values = [r.label, r.current, r.target, r.gap];
    values.slice(0, nCols).forEach(function(v, ci) {
      var w = colWidths[ci];
      // 目標列は黄ハイライト
      var cellBg = (ci === 2) ? C_HIGHLIGHT : bg;
      rect(slide, x, y, w, rowH, cellBg);
      var size = (ci === 0) ? 11 : (ci === 2 ? 13 : 11);
      var bold = (ci === 0 || ci === 2 || ci === 3);
      var color = (ci === 3) ? C_ACCENT2 : C_TEXT;
      txt(slide, v, x, y, w, rowH, {
        size: size, bold: bold, color: color,
        align: ci === 0 ? SlidesApp.ParagraphAlignment.START : SlidesApp.ParagraphAlignment.CENTER,
        va: SlidesApp.ContentAlignment.MIDDLE
      });
      x += w;
    });
  });
}

// ============================================================
// テンプレート全パターン デモ生成
// ============================================================
/**
 * 全パターンを連続生成して動作確認するデモ関数
 * 既存プレゼンの末尾に追加生成する
 */
function tpl_demoAll() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  var base = findBaseSlide(pres);
  if (!base) {
    Logger.log("base slide not found");
    return;
  }

  // Pattern A: アジェンダ
  var s = dup(pres, base);
  tpl_agendaList(s, [
    "2025年度施策振り返り",
    "国内外における他リテール企業のビジネストレンド",
    "ライブ配信は大きなデジタル接客の傘の一つに",
    "様々な収益化チャンス",
    "2026年度方針"
  ], 0);

  // Pattern C: KPIカード4分割
  s = dup(pres, base);
  tpl_titleSolid(s, "数字でみるマツココライブ実績", { size: 20 });
  tpl_kpiQuad(s, [
    { label: "累計視聴（ライブ＋アーカイブ）", value: "200", unit: "万人超", note: "2024年度 578,567人 ＋ 2025年度 1,440,000人", accent: C_ACCENT3 },
    { label: "累計 協賛売上", value: "2,935", unit: "万円", note: "2024年度 1,432万円 ＋ 2025年度 1,487万円", accent: C_ACCENT2 },
    { label: "配信回数(2024〜2025年度 累計)", value: "91", unit: "回", note: "2024年度開始〜2025年度までの累計配信数", accent: C_DARK },
    { label: "平均視聴者数", value: "4,650", unit: "人/回", note: "2024年度 3,043人→2025年度 5,854人", accent: C_DARK }
  ]);

  // Pattern E: Highlight/Lowlight/学び
  s = dup(pres, base);
  tpl_titleSolid(s, "2025年度の振り返り");
  tpl_reviewTriBox(s, {
    highlight: {
      badges: [
        { label: "効率性", bg: C_BG_PINK_BADGE, color: C_ACCENT },
        { label: "実施頻度", bg: C_BG_AMBER_BADGE, color: "#856404" }
      ],
      items: ["3名体制で毎週配信を実施", "協賛配信を毎週受注"],
      note: "他社は5〜7名体制が多い中、高効率でまわる体制を構築できている"
    },
    lowlight: {
      body: "視聴者数・コメント率は順調に成長する一方、平均視聴分数・商品CTR は伸び悩み"
    },
    learning: {
      body: "平均視聴分数・商品CTR が伸び悩み。これまで企画・施策で改善を試みてきたが、ここから抜本的な底上げは難しい。",
      emphasis: "熱量の高いリピーター・ファンを増やすことが、全体のエンゲージメントを底上げする",
      tail: "＝ 次年度の最重要テーマ"
    }
  });

  // Pattern F: 4課題カードグリッド
  s = dup(pres, base);
  tpl_titleSolid(s, "競合調査からみえたマツココライブ4つの課題");
  tpl_issueQuad(s, [
    { title: "❶ 商品カテゴリがバラバラで再訪動機が弱い", source: "平均視聴3.1分（60分の約5%）", cause: "配信ごとに商品ジャンルが変わり、視聴者が「次に何が見られるか」を予想できない。", accent: C_ACCENT2 },
    { title: "❷ リピーター不在", source: "リピーター比率5.0%（業界最下位・首位の1/10）", cause: "「戻り客を作る」思想が配信設計に組み込まれていない。", accent: C_ACCENT, isPriority: true },
    { title: "❸ 熱量のある視聴者の不在", source: "コメント率0.8%（A社33.4%の約1/40）", cause: "個人インフルエンサー型・固定BA型のいずれに該当せず、\"推し\"を持てない。", accent: C_ORANGE },
    { title: "❹ 購買動機の浅さ（自分ごと化できていない）", source: "商品CTR 1.0%（A社29.4%・C社11.7%の1/12〜1/30）", cause: "商品紹介止まりで「なぜ自分に必要か」が伝わっていない。", accent: C_INDIGO }
  ]);

  // Pattern I: 3カラム戦略
  s = dup(pres, base);
  setTitle(s, "リピーターを増やす3つの戦略");
  tpl_strategyTri(s,
    "視聴者が \"また観たくなる理由\"をつくる",
    [
      { title: "①「推し」でまた観たくなる", body: "出演者をきっかけにライブを観に来てもらう。", accent: C_ACCENT },
      { title: "②「おトク」でまた観たくなる", body: "顧客情報と連携できている強みを活かし、視聴をポイント・限定特典がもらえる場にする。", accent: C_BROWN },
      { title: "③「いつもの」でまた観たくなる", body: "コスメに特化することで、視聴者が「毎週コスメの最新情報・新商品が知れる」と分かる構造にする。", accent: C_ACCENT2 }
    ],
    "※具体的なアクション・施策は月例分科会などでディスカッションさせていただけますと幸いです。"
  );

  Logger.log("テンプレデモ生成完了");
}
