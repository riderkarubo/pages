/**
 * MCCM提案スライド化 - 共通ヘルパー
 *
 * 最終更新: 2026-04-28 01:32
 * 用途:
 *  全章で使う共通ヘルパー関数を定義する。
 *  ロレアル案件の replaceSlides_v2.js を MCCMテンプレ実測値に合わせて調整したもの。
 *
 * 必ず最初に Apps Script に貼り付け、各章のGASより前に定義しておくこと。
 *
 * テンプレ: TITLE_AND_BODY_2_1（全スライド共通）
 *  - TITLE (idx=0): x=25, y=21, w=671, h=45
 *  - BODY  (idx=1): x=25, y=76,  w=671, h=34（小見出し）
 *  - BODY  (idx=0): x=25, y=110, w=671, h=264（メイン本文）
 *  - BODY  (idx=2): x=0,  y=0,   w=205, h=17（dup時に削除）
 *  - SLIDE_NUMBER:  x=339, y=391, w=43, h=8
 */

// ============================================================
// 定数
// ============================================================
var PRESENTATION_ID = "14fq9YY5dAOksIIT-H_d1qsiCwyJEnhAY3dGp-KZSO6U";

// === Fireworkテーマカラー（ロレアル実測 + MCCM追加） ===
var C_DARK    = "#1A1A2E";  // タイトルバー・ダークカード
var C_WHITE   = "#FFFFFF";
var C_TEXT    = "#333333";
var C_ACCENT  = "#FA006D";  // マゼンタ・MCCM強調色（テキスト・KPI・タイムライン強調すべてここ・絶対変更禁止）
var C_ACCENT2 = "#1B998B";  // 緑・柱②・成長
var C_ACCENT3 = "#0F3460";  // 紺・KPI・タイトル
var C_GRAY    = "#666666";
var C_LIGHT   = "#F5F5F5";
var C_BORDER  = "#DDDDDD";
var C_SOFT    = "#F8F8F8";
var C_HIGHLIGHT = "#FFF2CC";  // 「2025年度=今年度」強調背景（実測値・P4-P5両方で使用）
var C_TIPS_BG   = "#F0F7FF";  // ポイントボックス背景（薄水色・実測値）
var C_TIPS_HEAD = "#1E3A8A";  // ポイント見出し色（紺青・実測値）
var C_NOTE      = "#434343";  // 標準テキスト（タイムライン・メッセージ系）

// === スライド寸法（16:9 = 720×405pt） ===
var SW = 720;
var SH = 405;

// === プレースホルダー座標（実測値） ===
// タイトル黒帯: P4実測値で確定 — 全スライドで絶対不変
var TITLE_X = 9,     TITLE_Y = 20,    TITLE_W = 671, TITLE_H = 45;
var SUBTITLE_X = 25, SUBTITLE_Y = 76, SUBTITLE_W = 671, SUBTITLE_H = 34;
var BODY_X = 25,     BODY_Y = 110,    BODY_W = 671, BODY_H = 264;

// === セーフゾーン（メイン本文エリア） ===
var SZ_X = BODY_X, SZ_Y = BODY_Y, SZ_W = BODY_W, SZ_H = BODY_H;
var SZ_CENTER_X = SZ_X + SZ_W / 2;
var SZ_CENTER_Y = SZ_Y + SZ_H / 2;

/**
 * セーフエリア中央に幅w・高さhのコンテンツブロックを配置するための左上座標を返す
 * @return {x, y, w, h}
 */
function centerInSafeZone(w, h) {
  return {
    x: SZ_X + (SZ_W - w) / 2,
    y: SZ_Y + (SZ_H - h) / 2,
    w: w,
    h: h
  };
}

// ============================================================
// スライド複製
// ============================================================
/**
 * 既存スライド（base）を複製→末尾に移動→非プレースホルダー要素を削除
 * 注: appendSlide(layout) はマスター不整合エラーが出るため使用禁止
 *     プレースホルダー idx=2 の小さい要素（205×17・セクションラベル相当）も削除する
 */
function dup(pres, base) {
  var d = base.duplicate();
  d.move(pres.getSlides().length);
  d.getPageElements().forEach(function(el) {
    try {
      if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        var shape = el.asShape();
        var phType = shape.getPlaceholderType();
        // 非プレースホルダーは削除
        if (phType === SlidesApp.PlaceholderType.NONE) {
          el.remove();
          return;
        }
        // idx=2 の小さいラベル相当も削除（205×17）
        if (phType === SlidesApp.PlaceholderType.BODY &&
            shape.getPlaceholderIndex() === 2 &&
            shape.getWidth() < 250) {
          el.remove();
          return;
        }
      } else if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
        el.remove();
      }
    } catch(e) {}
  });
  return d;
}

// ============================================================
// プレースホルダー操作
// ============================================================
/**
 * タイトル（黒帯の上に白文字テキストを重ねる方式）
 *
 *  座標は P4 実測値で確定: x=9, y=20, w=671, h=45
 *  この座標は全スライドで絶対不変（ずらさない）
 *
 *  注: スライドマスター/レイアウト由来の黒帯は「装飾」として描画されているだけで、
 *      スライド本体にはプレースホルダーとして存在しない（getPlaceholder で取得不可）。
 *      そのため、レイアウト座標の上に白文字テキストボックスを新規挿入する方式を採用する。
 */
function setTitle(slide, title) {
  txt(slide, title, TITLE_X, TITLE_Y, TITLE_W, TITLE_H, {
    size: 18,
    color: C_WHITE,
    bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
}

/**
 * デバッグ用: スライド上の全要素をログ出力
 */
function _debugLogElements(slide) {
  var els = slide.getPageElements();
  Logger.log("--- DEBUG: スライド要素一覧 (" + els.length + "個) ---");
  els.forEach(function(el, i) {
    var type = el.getPageElementType();
    var info = "[" + i + "] type=" + type +
      ", x=" + Math.round(el.getLeft()) +
      ", y=" + Math.round(el.getTop()) +
      ", w=" + Math.round(el.getWidth()) +
      ", h=" + Math.round(el.getHeight());
    if (type === SlidesApp.PageElementType.SHAPE) {
      var sh = el.asShape();
      info += ", phType=" + sh.getPlaceholderType() + ", phIdx=" + sh.getPlaceholderIndex();
    }
    Logger.log(info);
  });
}

// ============================================================
// フォントサイズ自動計算
// ============================================================
function calcFontSize(text, boxW, boxH, max, min) {
  max = max || 18; min = min || 6;
  var lines = text.split("\n");
  var maxLen = 0;
  lines.forEach(function(l) { if (l.length > maxLen) maxLen = l.length; });
  if (maxLen === 0) return max;
  var sizeByW = Math.floor(boxW / (maxLen * 0.95));
  var sizeByH = Math.floor(boxH / (lines.length * 1.35));
  return Math.max(min, Math.min(max, sizeByW, sizeByH));
}

// ============================================================
// 図形ヘルパー
// ============================================================
/** テキストボックス（背景・枠なし） */
function txt(slide, text, x, y, w, h, opts) {
  opts = opts || {};
  var size = opts.size || calcFontSize(text, w, h, 18, 6);
  var b = slide.insertTextBox(text, x, y, w, h);
  b.getFill().setTransparent();
  b.getBorder().setTransparent();
  b.getText().getParagraphs().forEach(function(p) {
    var st = p.getRange().getTextStyle();
    st.setFontFamily("Noto Sans JP");
    st.setFontSize(size);
    st.setForegroundColor(opts.color || C_TEXT);
    if (opts.bold) st.setBold(true);
    p.getRange().getParagraphStyle()
      .setParagraphAlignment(opts.align || SlidesApp.ParagraphAlignment.START);
  });
  b.setContentAlignment(opts.va || SlidesApp.ContentAlignment.MIDDLE);
  return b;
}

/** 矩形（テキストなし） */
function rect(slide, x, y, w, h, fill, stroke, strokeW) {
  var r = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  r.getFill().setSolidFill(fill);
  if (stroke) {
    r.getBorder().getLineFill().setSolidFill(stroke);
    r.getBorder().setWeight(strokeW || 1);
  } else {
    r.getBorder().setTransparent();
  }
  return r;
}

/** 矩形＋テキスト（直接埋め込み、1要素扱い） */
function rectTxt(slide, text, x, y, w, h, fill, opts, stroke, strokeW) {
  opts = opts || {};
  var r = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  r.getFill().setSolidFill(fill);
  if (stroke) {
    r.getBorder().getLineFill().setSolidFill(stroke);
    r.getBorder().setWeight(strokeW || 1);
  } else {
    r.getBorder().setTransparent();
  }
  if (text !== "") {
    var size = opts.size || calcFontSize(text, w, h, 18, 6);
    r.getText().setText(text);
    r.getText().getParagraphs().forEach(function(p) {
      var st = p.getRange().getTextStyle();
      st.setFontFamily("Noto Sans JP");
      st.setFontSize(size);
      st.setForegroundColor(opts.color || C_TEXT);
      if (opts.bold) st.setBold(true);
      p.getRange().getParagraphStyle()
        .setParagraphAlignment(opts.align || SlidesApp.ParagraphAlignment.CENTER);
    });
    r.setContentAlignment(opts.va || SlidesApp.ContentAlignment.MIDDLE);
  }
  return r;
}

/**
 * 文字種を考慮したテキスト描画幅推定（Noto Sans JP, Bold基準）
 *  Playwright実測値ベース（2026-04-27 計測）:
 *    半角数字 0-9        : fontSize × 0.71
 *    カンマ ,            : fontSize × 0.27
 *    ピリオド・スペース等 : fontSize × 0.27
 *    スラッシュ /        : fontSize × 0.52
 *    パーセント %        : fontSize × 0.85
 *    プラス・マイナス    : fontSize × 0.55
 *    アルファベット      : fontSize × 0.55（推定）
 *    全角(漢字・カナ)    : fontSize × 1.0
 *
 * Regular（非bold）の場合はやや狭くなるが、安全側で同一係数を使う。
 */
function estimateTextWidth(text, fontSize) {
  var w = 0;
  for (var i = 0; i < text.length; i++) {
    var c = text.charAt(i);
    var ratio;
    if (/[0-9]/.test(c))                ratio = 0.71;
    else if (/[,.\s]/.test(c))          ratio = 0.27;
    else if (c === '/')                 ratio = 0.52;
    else if (c === '%')                 ratio = 0.85;
    else if (/[+\-]/.test(c))           ratio = 0.55;
    else if (/[a-zA-Z]/.test(c))        ratio = 0.55;
    else                                ratio = 1.0;  // 全角
    w += fontSize * ratio;
  }
  return w;
}

/**
 * KPIカード描画ヘルパー（P3手動調整値に完全準拠）
 *
 * P3レイアウト実測値:
 *  カード: x, y, w=330, h=92, fill=#FFFFFF
 *  左ボーダー: x, y, w=4, h=92
 *  ラベル: x+12, y+6, w=306, h=14, size=9, bold, #666666, START
 *  数字: x+(w-(valueW+gap+unitW))/2 程度の調整 / size=28, bold, accent
 *  単位: 数字の右隣 / size=12, bold, accent
 *  注釈: x+12, y+74, w=306, h=14, size=8, #666666, START
 *
 *  数字+単位は「END+BOTTOM」「START+BOTTOM」ペアで右下基準にベースライン揃え。
 *
 * @param slide
 * @param x, y, w, h  カード位置とサイズ（標準: 330×92）
 * @param label       上部ラベル（例: "累計 協賛売上（2024〜2025年度）"）
 * @param value       数字文字列（例: "2,935"）
 * @param unit        単位文字列（例: "万円"）
 * @param note        下部注釈
 * @param accentColor 左ボーダー＋数字＋単位の色
 */
function drawKpiCard(slide, x, y, w, h, label, value, unit, note, accentColor) {
  // カード本体（白）
  rect(slide, x, y, w, h, C_WHITE);
  // 左ボーダー（4pt幅）
  rect(slide, x, y, 4, h, accentColor);

  var pad = 12;
  var labelW = w - pad * 2;

  // ラベル（上部・グレー）
  txt(slide, label, x + pad, y + 6, labelW, 14, {
    size: 9,
    color: C_GRAY,
    bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 数字 + 単位（カード中央領域・右寄せベースラインペア）
  var valueSize = 28;
  var unitSize = 12;
  // ボックス幅は実測係数 + 余裕20pt（漢字混じり「200万」等で実描画>推定のため）
  // Noto Sans JP の漢字は推定1.0倍だが実測でやや広く出るケースあり
  var valueW = estimateTextWidth(value, valueSize) + 20;
  var unitW = estimateTextWidth(unit, unitSize) + 10;
  var gap = 3;

  // ブロック全体をカード中央に配置
  var totalW = valueW + gap + unitW;
  var blockX = x + (w - totalW) / 2;

  // 数字ブロック（P3実測: 数字 y+26, 単位 y+38）
  var valueY = y + 26;
  var unitY = y + 38;

  // 数字（右寄せ・下揃え）
  txt(slide, value, blockX, valueY, valueW, 36, {
    size: valueSize,
    color: accentColor,
    bold: true,
    align: SlidesApp.ParagraphAlignment.END,
    va: SlidesApp.ContentAlignment.BOTTOM
  });
  // 単位（左寄せ・下揃え）
  txt(slide, unit, blockX + valueW + gap, unitY, unitW, 26, {
    size: unitSize,
    color: accentColor,
    bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.BOTTOM
  });

  // 注釈（下部・グレー）
  txt(slide, note, x + pad, y + 74, labelW, 14, {
    size: 8,
    color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
}

/**
 * メインメッセージ部（P3手動調整値に完全準拠）
 *  メイン: y=88, h=24, size=14, bold, #434343, CENTER
 *  サブ:   y=121, h=18, size=10, #434343, CENTER
 *  両方とも x=41, w=639
 */
function drawExecMessage(slide, mainText, subText) {
  var msgX = 41;
  var msgW = 639;
  txt(slide, mainText, msgX, 88, msgW, 24, {
    size: 14,
    color: "#434343",
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(slide, subText, msgX, 121, msgW, 18, {
    size: 10,
    color: "#434343",
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
}

/** 楕円 */
function ellipse(slide, x, y, w, h, fill) {
  var e = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, y, w, h);
  e.getFill().setSolidFill(fill);
  e.getBorder().setTransparent();
  return e;
}

/** 矢印 */
function arrow(slide, x1, y1, x2, y2, color, w) {
  var line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x1, y1, x2, y2);
  line.getLineFill().setSolidFill(color || C_DARK);
  line.setWeight(w || 1.5);
  line.setEndArrow(SlidesApp.ArrowStyle.FILL_ARROW);
  return line;
}

// ============================================================
// テーブル
// ============================================================
/**
 * テーブル生成
 * @param colDefs [{label, align, nameCol}, ...]
 * @param rows    [[r1c1, r1c2, ...], [r2c1, ...], ...]
 * @param opts    {h, headerBg, headerSize, headerAlign, bodySize, highlightRows: [行番号(0始まり)]}
 */
function makeTable(slide, x, y, w, colDefs, rows, opts) {
  opts = opts || {};
  var rowH = opts.rowH || 22;
  var hdrH = opts.hdrH || 26;
  var totalH = opts.h || (hdrH + rowH * rows.length);
  var tbl = slide.insertTable(rows.length + 1, colDefs.length, x, y, w, totalH);

  // ヘッダー（先にテキストを入れてからスタイル適用）
  for (var c = 0; c < colDefs.length; c++) {
    var cell = tbl.getCell(0, c);
    cell.getFill().setSolidFill(opts.headerBg || C_DARK);
    cell.getText().setText(colDefs[c].label || "");
    _styleCell(cell,
      opts.headerSize || 10,
      true,
      C_WHITE,
      opts.headerAlign || SlidesApp.ParagraphAlignment.CENTER);
  }

  // ボディ（先にテキストを入れてからスタイル適用）
  var highlightRows = opts.highlightRows || [];
  for (var r = 0; r < rows.length; r++) {
    var isHighlight = highlightRows.indexOf(r) >= 0;
    for (var c2 = 0; c2 < colDefs.length; c2++) {
      var cell2 = tbl.getCell(r + 1, c2);
      var bg = isHighlight ? C_HIGHLIGHT : (r % 2 === 0 ? C_WHITE : C_SOFT);
      cell2.getFill().setSolidFill(bg);
      cell2.getText().setText(rows[r][c2] || "");
      _styleCell(cell2,
        opts.bodySize || 9,
        colDefs[c2].nameCol || isHighlight,
        C_TEXT,
        colDefs[c2].align || SlidesApp.ParagraphAlignment.START);
    }
  }
  return tbl;
}

function _styleCell(cell, size, bold, color, align) {
  var tr = cell.getText();
  // セルが空だとgetParagraphs()が空 → スペース1文字を入れて段落を作る
  if (!tr.asString() || tr.asString().length === 0) {
    tr.setText(" ");
  }
  tr.getParagraphs().forEach(function(p) {
    var range = p.getRange();
    if (!range) return;
    var st = range.getTextStyle();
    if (st) {
      st.setFontFamily("Noto Sans JP").setFontSize(size).setForegroundColor(color);
      if (bold) st.setBold(true);
    }
    var ps = range.getParagraphStyle();
    if (ps) ps.setParagraphAlignment(align);
  });
  cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
}

// ============================================================
// 共通スライド作成
// ============================================================
/**
 * 複製ベースとなる「TITLE_AND_BODY_2_1」レイアウトのスライドを自動検索
 *  優先順位:
 *    1. レイアウト名が "TITLE_AND_BODY_2_1" のスライドを最初に見つけたもの
 *    2. それが無ければ末尾から探して最初に見つかった同レイアウトのスライド
 *    3. それも無ければ最後のスライド（フォールバック）
 */
function findBaseSlide(pres) {
  var slides = pres.getSlides();
  if (slides.length === 0) {
    throw new Error("スライドが1枚もありません。複製元となるスライドを少なくとも1枚配置してください。");
  }
  // 前方から TITLE_AND_BODY_2_1 を探す
  for (var i = 0; i < slides.length; i++) {
    if (slides[i].getLayout().getLayoutName() === "TITLE_AND_BODY_2_1") {
      return slides[i];
    }
  }
  // 見つからなければ最後のスライドをフォールバックに
  Logger.log("警告: TITLE_AND_BODY_2_1 のスライドが見つからないため、最終スライドを複製ベースにします");
  return slides[slides.length - 1];
}

/**
 * 新規スライド作成（TITLE_AND_BODY_2_1 を自動検出して複製）
 *  title引数を渡せば自動でタイトル設定。省略時は呼び出し側でtxt()でタイトル配置
 * 戻り値: 空のスライド
 */
function newSlide(pres, title) {
  var base = findBaseSlide(pres);
  var s = dup(pres, base);
  if (title) setTitle(s, title);
  return s;
}

// ============================================================
// Issy手直しテンプレ（P4-P7 解析実測値ベース）
// ============================================================

/**
 * 結論メッセージ（テーブル上に置く中央太字16pt）
 * 強調キーワードだけ赤(C_ACCENT)、それ以外は #434343
 *
 * @param slide
 * @param y    Y座標（推奨: 90〜100、テーブル上に置く）
 * @param text 全文
 * @param highlight 強調する部分文字列（その部分だけ赤太字）
 */
function drawConclusionMessage(slide, y, text, highlight) {
  var sh = slide.insertTextBox(text || "", 35, y, 651, 24);
  sh.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  var tr = sh.getText();
  tr.getParagraphs().forEach(function(p) {
    var ps = p.getRange().getParagraphStyle();
    ps.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    var st = p.getRange().getTextStyle();
    st.setFontFamily("Noto Sans JP").setFontSize(16).setBold(true).setForegroundColor(C_NOTE);
  });
  // インライン強調（部分文字列だけ赤）
  if (highlight && text.indexOf(highlight) >= 0) {
    var start = text.indexOf(highlight);
    var end = start + highlight.length;
    try {
      var range = tr.getRange(start, end);
      range.getTextStyle().setForegroundColor(C_ACCENT).setBold(true);
    } catch(e) {}
  }
  return sh;
}

/**
 * 列ハイライト型テーブル（P4=視聴者数推移パターン）
 * 「2025年度=今年度」を視覚強調するため、特定の列だけクリーム背景にする
 *
 * 偶数行=#F8F8F8 / 奇数行=#FFFFFF / 指定列=#FFF2CC で上書き
 *
 * @param slide
 * @param x, y, w
 * @param colDefs       [{label, align, nameCol}, ...]
 * @param rows          [[r1c1, ...], ...]
 * @param opts          {highlightCol: 列番号(0始まり), highlightRow: 行番号(0始まり), bodySize}
 */
function makeColHighlightTable(slide, x, y, w, colDefs, rows, opts) {
  opts = opts || {};
  var rowH = opts.rowH || 32;
  var hdrH = opts.hdrH || 30;
  var totalH = hdrH + rowH * rows.length;
  var tbl = slide.insertTable(rows.length + 1, colDefs.length, x, y, w, totalH);

  var headerSize = opts.headerSize || 10;
  var bodySize = opts.bodySize || 10;
  var highlightCol = (typeof opts.highlightCol === "number") ? opts.highlightCol : -1;
  var highlightRow = (typeof opts.highlightRow === "number") ? opts.highlightRow : -1;

  // ヘッダー
  for (var c = 0; c < colDefs.length; c++) {
    var cell = tbl.getCell(0, c);
    cell.getFill().setSolidFill(C_DARK);
    cell.getText().setText(colDefs[c].label || "");
    _styleCell(cell, headerSize, true, C_WHITE,
      colDefs[c].headerAlign || SlidesApp.ParagraphAlignment.CENTER);
  }

  // ボディ
  for (var r = 0; r < rows.length; r++) {
    var isRowHighlight = (r === highlightRow);
    for (var c2 = 0; c2 < colDefs.length; c2++) {
      var cell2 = tbl.getCell(r + 1, c2);
      var bg;
      if (c2 === highlightCol || isRowHighlight) {
        bg = C_HIGHLIGHT;  // クリーム
      } else {
        bg = (r % 2 === 0) ? C_WHITE : C_SOFT;
      }
      cell2.getFill().setSolidFill(bg);
      cell2.getText().setText(rows[r][c2] || "");
      var bold = colDefs[c2].nameCol || isRowHighlight || (rows[r].__bold === true);
      var color = colDefs[c2].forceColor || C_TEXT;
      _styleCell(cell2, bodySize, bold, color,
        colDefs[c2].align || SlidesApp.ParagraphAlignment.CENTER);
    }
  }

  // 特定行を bold にする(highlightRow指定時)
  if (highlightRow >= 0) {
    for (var c3 = 0; c3 < colDefs.length; c3++) {
      var cb = tbl.getCell(highlightRow + 1, c3);
      try {
        cb.getText().getParagraphs().forEach(function(p) {
          p.getRange().getTextStyle().setBold(true);
        });
      } catch(e) {}
    }
  }

  return tbl;
}

/**
 * ポイントボックス（P5=協賛実績パターン）
 * 薄水色背景 + 「💡 ポイント」見出し + 本文複数行 + 結論行(矢印つき bold)
 *
 * @param slide
 * @param x, y, w, h
 * @param lines [{ text, conclusion: true|false }, ...]
 */
function drawTipsBox(slide, x, y, w, h, lines) {
  // 背景
  rect(slide, x, y, w, h, C_TIPS_BG);

  // 見出し
  var head = slide.insertTextBox("💡 ポイント", x + 10, y + 10, w - 20, 18);
  head.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  head.getText().getParagraphs().forEach(function(p) {
    p.getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    p.getRange().getTextStyle().setFontFamily("Noto Sans JP").setFontSize(11).setBold(true).setForegroundColor(C_TIPS_HEAD);
  });

  // 本文
  var lineY = y + 32;
  lines.forEach(function(line) {
    var sh = slide.insertTextBox(line.text || "", x + 7, lineY, w - 15, 22);
    sh.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    sh.getText().getParagraphs().forEach(function(p) {
      p.getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
      var st = p.getRange().getTextStyle().setFontFamily("Noto Sans JP").setFontSize(11);
      if (line.conclusion) {
        st.setBold(true).setForegroundColor(C_ACCENT);
      } else {
        st.setForegroundColor(C_TEXT);
      }
    });
    lineY += 23;
  });
}

/**
 * タイムライン1行を描画（P6=マツココライブの変遷パターン）
 *
 * @param slide
 * @param y           行のY座標(マークの中心)
 * @param date        左の日付ラベル
 * @param desc        右の説明テキスト
 * @param highlight   先頭の強調キーワード（その部分だけ赤、無ければ全文 #434343）
 * @param isMilestone true=赤丸 / false=グレー丸
 */
function drawTimelineRow(slide, y, date, desc, highlight, isMilestone) {
  // 日付ラベル(左, x=30, w=110)
  var dateBox = slide.insertTextBox(date || "", 30, y, 110, 20);
  dateBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  dateBox.getText().getParagraphs().forEach(function(p) {
    p.getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
    p.getRange().getTextStyle().setFontFamily("Noto Sans JP").setFontSize(9).setBold(true).setForegroundColor(C_NOTE);
  });

  // ●マーク(中央、x=152, y=y+4 → 中心揃え)
  var dotColor = isMilestone ? C_ACCENT : C_GRAY;
  var dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 152, y + 4, 10, 10);
  dot.getFill().setSolidFill(dotColor);
  dot.getBorder().setTransparent();

  // 説明テキスト(右, x=175, w=520)
  var descBox = slide.insertTextBox(desc || "", 175, y, 520, 22);
  descBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  var tr = descBox.getText();
  tr.getParagraphs().forEach(function(p) {
    p.getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    p.getRange().getTextStyle().setFontFamily("Noto Sans JP").setFontSize(9.5).setBold(true).setForegroundColor(C_NOTE);
  });

  // インライン強調（先頭のキーワードだけ赤）
  if (highlight && desc.indexOf(highlight) >= 0) {
    var start = desc.indexOf(highlight);
    var end = start + highlight.length;
    try {
      tr.getRange(start, end).getTextStyle().setForegroundColor(C_ACCENT).setBold(true);
    } catch(e) {}
  }
}

/**
 * タイムラインの縦線（P6パターン）
 * @param slide
 * @param topY, bottomY    縦線の上下端
 */
function drawTimelineAxis(slide, topY, bottomY) {
  rect(slide, 156, topY, 2, bottomY - topY, C_BORDER);
}

/**
 * 経営資産カード（P7パターン）
 * 左ボーダー w=3 / 名前11pt bold / 値12pt bold アクセント色 / 注釈8pt #666
 *
 * @param slide
 * @param x, y, w=160, h=110
 * @param p {no, name, value, note, accent}
 */
function drawAssetCardP7(slide, x, y, w, h, p) {
  var accent = p.accent || C_ACCENT;
  // カード本体（白＋グレー枠線）
  rect(slide, x, y, w, h, C_WHITE, C_BORDER, 1);
  // 左ボーダー（資産ごとのアクセント色）
  rect(slide, x, y, 3, h, accent);

  // 番号(資産 NN) - y+6, h=12, 8pt bold #666
  var noBox = slide.insertTextBox(p.no || "", x + 10, y + 6, w - 20, 12);
  noBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  noBox.getText().getParagraphs().forEach(function(pp) {
    pp.getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    pp.getRange().getTextStyle().setFontFamily("Noto Sans JP").setFontSize(8).setBold(true).setForegroundColor(C_GRAY);
  });

  // 名前 - y+22, h=16, 11pt bold #1A1A2E
  var nameBox = slide.insertTextBox(p.name || "", x + 10, y + 22, w - 20, 16);
  nameBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  nameBox.getText().getParagraphs().forEach(function(pp) {
    pp.getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    pp.getRange().getTextStyle().setFontFamily("Noto Sans JP").setFontSize(11).setBold(true).setForegroundColor(C_DARK);
  });

  // 値 - y+44, h=36, 12pt bold アクセント色
  var valBox = slide.insertTextBox(p.value || "", x + 10, y + 44, w - 20, 36);
  valBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  valBox.getText().getParagraphs().forEach(function(pp) {
    pp.getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    pp.getRange().getTextStyle().setFontFamily("Noto Sans JP").setFontSize(12).setBold(true).setForegroundColor(accent);
  });

  // 注釈 - y+84, h=22, 8pt 通常 #666
  var noteBox = slide.insertTextBox(p.note || "", x + 10, y + 84, w - 20, 22);
  noteBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  noteBox.getText().getParagraphs().forEach(function(pp) {
    pp.getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    pp.getRange().getTextStyle().setFontFamily("Noto Sans JP").setFontSize(8).setBold(false).setForegroundColor(C_GRAY);
  });
}
