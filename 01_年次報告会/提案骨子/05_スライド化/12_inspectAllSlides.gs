/**
 * MCCM提案スライド化 - 完成版全スライド構造解析
 *
 * 最終更新: 2026-04-28 02:29
 * 用途:
 *  ユーザー手直し含む完成版全スライドを完全解析。
 *  Fireworkテンプレ準拠スライド作成のテンプレライブラリ化のため、
 *  全要素の座標・色・テキスト・スタイル・テーブルセル詳細を全枚数で取得。
 *
 * 実行手順:
 *  1. このコードを Apps Script に貼り付け
 *  2. inspectAllSlides() を実行（全枚数解析）
 *  3. 実行ログ全文を Issy に共有
 *
 * 単独実行: inspectSlideAtIdx(N) で N 枚目だけ解析（0始まり）
 */

var ALL_INSPECT_PRESENTATION_ID = "14fq9YY5dAOksIIT-H_d1qsiCwyJEnhAY3dGp-KZSO6U";

function inspectAllSlides() {
  var pres = SlidesApp.openById(ALL_INSPECT_PRESENTATION_ID);
  var slides = pres.getSlides();
  var total = slides.length;
  Logger.log("=== Slides全枚数: " + total + " ===");
  Logger.log("");

  for (var i = 0; i < total; i++) {
    _inspectSlideFull(pres, i);
    Logger.log("");
    Logger.log("================================================");
    Logger.log("");
  }
}

function inspectSlideAtIdx(idx) {
  var pres = SlidesApp.openById(ALL_INSPECT_PRESENTATION_ID);
  _inspectSlideFull(pres, idx);
}

function _inspectSlideFull(pres, idx) {
  var slide = pres.getSlides()[idx];
  Logger.log("=== P" + (idx + 1) + " (index=" + idx + ") ===");
  try {
    Logger.log("レイアウト: " + slide.getLayout().getLayoutName());
  } catch(e) {}

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
      _inspShape(el.asShape(), "  ");
    } else if (type === SlidesApp.PageElementType.LINE) {
      _inspLine(el.asLine());
    } else if (type === SlidesApp.PageElementType.IMAGE) {
      Logger.log("  IMAGE");
    } else if (type === SlidesApp.PageElementType.TABLE) {
      _inspTable(el.asTable());
    } else if (type === SlidesApp.PageElementType.GROUP) {
      var group = el.asGroup();
      Logger.log("  GROUP children=" + group.getChildren().length);
      group.getChildren().forEach(function(child, ci) {
        Logger.log("    -- child[" + ci + "] type=" + child.getPageElementType());
        if (child.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          _inspShape(child.asShape(), "    ");
        }
      });
    }
    Logger.log("");
  });

  Logger.log("=== P" + (idx + 1) + " 完了 ===");
}

function _inspShape(shape, indent) {
  indent = indent || "  ";
  try {
    Logger.log(indent + "PlaceholderType=" + shape.getPlaceholderType() + ", PlaceholderIndex=" + shape.getPlaceholderIndex());
  } catch(e) {}
  try {
    var fill = shape.getFill();
    if (fill.getType() === SlidesApp.FillType.SOLID) {
      var sf = fill.getSolidFill();
      var col = sf.getColor();
      if (col.getColorType() === SlidesApp.ColorType.RGB) {
        Logger.log(indent + "fill: #" + col.asRgbColor().asHexString().replace("#","") + ", alpha=" + sf.getAlpha());
      } else {
        Logger.log(indent + "fill: THEME(" + col.asThemeColor().getThemeColorType() + ")");
      }
    } else {
      Logger.log(indent + "fill: " + fill.getType());
    }
  } catch(e) {}
  try {
    var border = shape.getBorder();
    if (border.isVisible()) {
      var lf = border.getLineFill();
      var w = border.getWeight();
      if (lf.getFillType() === SlidesApp.LineFillType.SOLID) {
        var bc = lf.getSolidFill().getColor();
        if (bc.getColorType() === SlidesApp.ColorType.RGB) {
          Logger.log(indent + "border: #" + bc.asRgbColor().asHexString().replace("#","") + ", weight=" + w);
        }
      }
    }
  } catch(e) {}
  try { Logger.log(indent + "contentAlign=" + shape.getContentAlignment()); } catch(e) {}
  try {
    var tr = shape.getText();
    var t = tr.asString();
    if (t.length > 0 && t !== "\n") {
      Logger.log(indent + "text: \"" + t.replace(/\n/g, "\\n") + "\"");
      _logParas(tr, indent + "  ");
    }
  } catch(e) {}
}

function _inspLine(line) {
  try {
    var lf = line.getLineFill();
    if (lf.getFillType() === SlidesApp.LineFillType.SOLID) {
      var lc = lf.getSolidFill().getColor();
      if (lc.getColorType() === SlidesApp.ColorType.RGB) {
        Logger.log("  line: #" + lc.asRgbColor().asHexString().replace("#",""));
      }
    }
    Logger.log("  weight=" + line.getWeight());
  } catch(e) {}
}

function _inspTable(table) {
  var rows = table.getNumRows();
  var cols = table.getNumColumns();
  Logger.log("  TABLE rows=" + rows + ", cols=" + cols);
  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < cols; c++) {
      var cell = table.getCell(r, c);
      var prefix = "    [" + r + "][" + c + "]";
      try {
        var fill = cell.getFill();
        if (fill.getType() === SlidesApp.FillType.SOLID) {
          var sf = fill.getSolidFill();
          var col = sf.getColor();
          if (col.getColorType() === SlidesApp.ColorType.RGB) {
            Logger.log(prefix + " bg=#" + col.asRgbColor().asHexString().replace("#",""));
          }
        }
      } catch(e) {}
      try {
        var t = cell.getText().asString();
        if (t.length > 0 && t !== "\n") {
          Logger.log(prefix + " text=\"" + t.replace(/\n/g, "\\n") + "\"");
          _logParas(cell.getText(), prefix + "  ");
        }
      } catch(e) {}
    }
  }
}

function _logParas(textRange, indent) {
  try {
    textRange.getParagraphs().forEach(function(p, pi) {
      var pText = p.getRange().asString();
      if (pText.length === 0 || pText === "\n") return;
      var st = p.getRange().getTextStyle();
      var info = indent + "[p" + pi + "] \"" + pText.substring(0, 30).replace(/\n/g, "\\n") + "\"";
      try { info += " font=" + st.getFontFamily(); } catch(e) {}
      try { info += " size=" + st.getFontSize(); } catch(e) {}
      try { info += " bold=" + st.isBold(); } catch(e) {}
      try {
        var fc = st.getForegroundColor();
        if (fc && fc.getColorType() === SlidesApp.ColorType.RGB) {
          info += " color=#" + fc.asRgbColor().asHexString().replace("#","");
        }
      } catch(e) {}
      Logger.log(info);
    });
  } catch(e) {}
}
