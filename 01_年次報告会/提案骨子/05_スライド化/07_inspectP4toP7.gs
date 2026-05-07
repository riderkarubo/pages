/**
 * MCCM提案スライド化 - Phase 0+++: P4-P7（ユーザー手直し版）の完全構造取得
 *
 * 最終更新: 2026-04-28 01:32
 * 用途:
 *  ユーザーが手動で整えたP4〜P7を完全再現するため、以下を取得:
 *   - 全要素の座標 (x, y, w, h)
 *   - 図形の塗り色（HEX + alpha）/ 枠線
 *   - テキスト全文・段落ごとのフォント・サイズ・bold/italic・色・背景色・配置
 *   - **テーブルの場合は全セルの色・背景・テキスト・スタイル**
 *   - 縦配置・回転・行間
 *
 * PDFスライドとの対応:
 *   PDF Page 4 = Slides P4 (index=3): 章扉？空白？
 *   PDF Page 5 = Slides P5 (index=4): 年度別 視聴者数推移（テーブル）
 *   PDF Page 6 = Slides P6 (index=5): 協賛実績の推移（テーブル）
 *   PDF Page 7 = Slides P7 (index=6): マツココライブの変遷（タイムライン）
 *
 * 実行手順:
 *  1. このコードを Apps Script に貼り付け
 *  2. inspectP4toP7() を実行（4枚一気に解析）
 *     または個別: inspectSlideAt(3) / inspectSlideAt(4) / inspectSlideAt(5) / inspectSlideAt(6)
 *  3. 実行ログ全文を Issy に共有
 */

var INSPECT_PRESENTATION_ID = "14fq9YY5dAOksIIT-H_d1qsiCwyJEnhAY3dGp-KZSO6U";

// ============================================================
// メインエントリ
// ============================================================
function inspectP4toP7() {
  var pres = SlidesApp.openById(INSPECT_PRESENTATION_ID);
  var total = pres.getSlides().length;
  Logger.log("=== Slides全枚数: " + total + " ===");
  Logger.log("");

  // P4 = index 3, P5 = index 4, P6 = index 5, P7 = index 6
  [3, 4, 5, 6].forEach(function(idx) {
    if (idx < total) {
      _inspectSlide(pres, idx);
      Logger.log("");
      Logger.log("================================================");
      Logger.log("");
    } else {
      Logger.log("⚠ P" + (idx + 1) + " はスライド数を超えているためスキップ");
    }
  });
}

// 個別実行用
function inspectSlideAt(idx) {
  var pres = SlidesApp.openById(INSPECT_PRESENTATION_ID);
  _inspectSlide(pres, idx);
}

function inspectP4() { inspectSlideAt(3); }
function inspectP5() { inspectSlideAt(4); }
function inspectP6() { inspectSlideAt(5); }
function inspectP7() { inspectSlideAt(6); }

// ============================================================
// スライド1枚の解析
// ============================================================
function _inspectSlide(pres, idx) {
  var slide = pres.getSlides()[idx];
  Logger.log("=== P" + (idx + 1) + " (index=" + idx + ") ===");

  try {
    Logger.log("レイアウト: " + slide.getLayout().getLayoutName());
  } catch(e) {
    Logger.log("レイアウト取得エラー: " + e.message);
  }

  var elements = slide.getPageElements();
  Logger.log("要素数: " + elements.length + "個");
  Logger.log("");

  elements.forEach(function(el, i) {
    var type = el.getPageElementType();
    Logger.log("--- [" + i + "] type=" + type + " ---");
    Logger.log("  座標: x=" + Math.round(el.getLeft()) +
      ", y=" + Math.round(el.getTop()) +
      ", w=" + Math.round(el.getWidth()) +
      ", h=" + Math.round(el.getHeight()));

    try {
      var rotation = el.getRotation();
      if (rotation !== 0) Logger.log("  回転: " + rotation + "度");
    } catch(e) {}

    if (type === SlidesApp.PageElementType.SHAPE) {
      _inspectShape(el.asShape());
    } else if (type === SlidesApp.PageElementType.LINE) {
      _inspectLine(el.asLine());
    } else if (type === SlidesApp.PageElementType.IMAGE) {
      Logger.log("  IMAGE");
    } else if (type === SlidesApp.PageElementType.TABLE) {
      _inspectTable(el.asTable());
    } else if (type === SlidesApp.PageElementType.GROUP) {
      var group = el.asGroup();
      Logger.log("  GROUP children=" + group.getChildren().length + "個");
      group.getChildren().forEach(function(child, ci) {
        Logger.log("    -- group child[" + ci + "] type=" + child.getPageElementType());
        try {
          Logger.log("       座標: x=" + Math.round(child.getLeft()) +
            ", y=" + Math.round(child.getTop()) +
            ", w=" + Math.round(child.getWidth()) +
            ", h=" + Math.round(child.getHeight()));
        } catch(e) {}
        if (child.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          _inspectShape(child.asShape(), "       ");
        }
      });
    }
    Logger.log("");
  });

  Logger.log("=== P" + (idx + 1) + " 完了 ===");
}

// ============================================================
// SHAPE 解析
// ============================================================
function _inspectShape(shape, indent) {
  indent = indent || "  ";
  try {
    Logger.log(indent + "PlaceholderType=" + shape.getPlaceholderType() +
      ", PlaceholderIndex=" + shape.getPlaceholderIndex());
  } catch(e) {}

  // 塗り色
  try {
    var fill = shape.getFill();
    if (fill.getType() === SlidesApp.FillType.SOLID) {
      var sf = fill.getSolidFill();
      var color = sf.getColor();
      var alpha = sf.getAlpha();
      if (color.getColorType() === SlidesApp.ColorType.RGB) {
        Logger.log(indent + "fill: #" + color.asRgbColor().asHexString().replace("#","") +
          ", alpha=" + alpha);
      } else {
        Logger.log(indent + "fill: THEME(" + color.asThemeColor().getThemeColorType() +
          "), alpha=" + alpha);
      }
    } else {
      Logger.log(indent + "fill: " + fill.getType());
    }
  } catch(e) {}

  // 枠線
  try {
    var border = shape.getBorder();
    if (border.isVisible()) {
      var lf = border.getLineFill();
      var w = border.getWeight();
      if (lf.getFillType() === SlidesApp.LineFillType.SOLID) {
        var bc = lf.getSolidFill().getColor();
        if (bc.getColorType() === SlidesApp.ColorType.RGB) {
          Logger.log(indent + "border: #" + bc.asRgbColor().asHexString().replace("#","") +
            ", weight=" + w);
        }
      }
    }
  } catch(e) {}

  // 中身配置
  try { Logger.log(indent + "contentAlign=" + shape.getContentAlignment()); } catch(e) {}

  // テキスト
  try {
    var tr = shape.getText();
    var fullText = tr.asString();
    if (fullText.length > 0 && fullText !== "\n") {
      Logger.log(indent + "テキスト全文: \"" + fullText.replace(/\n/g, "\\n") + "\"");
      _logParagraphs(tr, indent + "  ");
    } else {
      Logger.log(indent + "テキスト: (空)");
    }
  } catch(e) {}
}

// ============================================================
// LINE 解析
// ============================================================
function _inspectLine(line) {
  try {
    var lf = line.getLineFill();
    if (lf.getFillType() === SlidesApp.LineFillType.SOLID) {
      var lc = lf.getSolidFill().getColor();
      if (lc.getColorType() === SlidesApp.ColorType.RGB) {
        Logger.log("  line color: #" + lc.asRgbColor().asHexString().replace("#",""));
      }
    }
    Logger.log("  weight=" + line.getWeight());
  } catch(e) {}
  try {
    var startStyle = line.getStartArrow();
    var endStyle = line.getEndArrow();
    if (startStyle && startStyle !== SlidesApp.ArrowStyle.NONE) {
      Logger.log("  startArrow=" + startStyle);
    }
    if (endStyle && endStyle !== SlidesApp.ArrowStyle.NONE) {
      Logger.log("  endArrow=" + endStyle);
    }
  } catch(e) {}
  try {
    Logger.log("  dashStyle=" + line.getDashStyle());
  } catch(e) {}
}

// ============================================================
// TABLE 解析（全セル詳細取得）
// ============================================================
function _inspectTable(table) {
  var rows = table.getNumRows();
  var cols = table.getNumColumns();
  Logger.log("  TABLE rows=" + rows + ", cols=" + cols);

  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < cols; c++) {
      var cell = table.getCell(r, c);
      var prefix = "    [" + r + "][" + c + "]";

      // セル背景
      try {
        var fill = cell.getFill();
        if (fill.getType() === SlidesApp.FillType.SOLID) {
          var sf = fill.getSolidFill();
          var col = sf.getColor();
          if (col.getColorType() === SlidesApp.ColorType.RGB) {
            Logger.log(prefix + " bg=#" + col.asRgbColor().asHexString().replace("#","") +
              ", alpha=" + sf.getAlpha());
          }
        }
      } catch(e) {}

      // 縦配置
      try { Logger.log(prefix + " contentAlign=" + cell.getContentAlignment()); } catch(e) {}

      // テキスト
      try {
        var tr = cell.getText();
        var t = tr.asString();
        if (t.length > 0 && t !== "\n") {
          Logger.log(prefix + " text=\"" + t.replace(/\n/g, "\\n") + "\"");
          _logParagraphs(tr, prefix + "   ");
        }
      } catch(e) {}
    }
  }
}

// ============================================================
// 段落ごとの詳細出力（共通）
// ============================================================
function _logParagraphs(textRange, indent) {
  try {
    var paragraphs = textRange.getParagraphs();
    paragraphs.forEach(function(p, pi) {
      var pText = p.getRange().asString();
      if (pText.length === 0 || pText === "\n") return;

      var st = p.getRange().getTextStyle();
      var ps = p.getRange().getParagraphStyle();
      var info = indent + "[p" + pi + "] \"" + pText.substring(0, 40).replace(/\n/g, "\\n") + "\"";
      try { info += " | font=" + st.getFontFamily(); } catch(e) {}
      try { info += ", size=" + st.getFontSize(); } catch(e) {}
      try { info += ", bold=" + st.isBold(); } catch(e) {}
      try { info += ", italic=" + st.isItalic(); } catch(e) {}
      try {
        var fc = st.getForegroundColor();
        if (fc && fc.getColorType() === SlidesApp.ColorType.RGB) {
          info += ", color=#" + fc.asRgbColor().asHexString().replace("#","");
        }
      } catch(e) {}
      try {
        var bc = st.getBackgroundColor();
        if (bc && bc.getColorType() === SlidesApp.ColorType.RGB) {
          info += ", bgColor=#" + bc.asRgbColor().asHexString().replace("#","");
        }
      } catch(e) {}
      try { info += ", align=" + ps.getParagraphAlignment(); } catch(e) {}
      try {
        var ls = ps.getLineSpacing();
        if (ls !== 100) info += ", lineSpacing=" + ls;
      } catch(e) {}
      Logger.log(info);

      // 段落内のリッチテキスト範囲（部分的な色変化を取得）
      _logRichRanges(p.getRange(), indent + "  ");
    });
  } catch(e) {
    Logger.log(indent + "段落解析エラー: " + e.message);
  }
}

// ============================================================
// 段落内の部分テキスト色（インラインカラー検知）
// ============================================================
function _logRichRanges(textRange, indent) {
  try {
    var ranges = textRange.getRuns();
    if (!ranges || ranges.length <= 1) return;
    Logger.log(indent + "── インライン書式 (" + ranges.length + " run) ──");
    ranges.forEach(function(run, ri) {
      var t = run.asString();
      if (!t || t === "\n") return;
      var st = run.getTextStyle();
      var info = indent + "run[" + ri + "] \"" + t.substring(0, 30).replace(/\n/g, "\\n") + "\"";
      try { info += " size=" + st.getFontSize(); } catch(e) {}
      try { info += " bold=" + st.isBold(); } catch(e) {}
      try {
        var fc = st.getForegroundColor();
        if (fc && fc.getColorType() === SlidesApp.ColorType.RGB) {
          info += " color=#" + fc.asRgbColor().asHexString().replace("#","");
        }
      } catch(e) {}
      try {
        var bc = st.getBackgroundColor();
        if (bc && bc.getColorType() === SlidesApp.ColorType.RGB) {
          info += " bg=#" + bc.asRgbColor().asHexString().replace("#","");
        }
      } catch(e) {}
      Logger.log(info);
    });
  } catch(e) {
    // getRuns() が無いバージョンもあるので無視
  }
}
