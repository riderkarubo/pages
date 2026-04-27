/**
 * MCCM提案スライド化 - Phase 0+++: P4（ユーザー再修正版）の完全構造取得
 *
 * 用途:
 *  ユーザーが手動で整えたP4（エグゼクティブサマリ）を完全再現するため、
 *  以下のすべてを取得する:
 *   - 全要素の座標 (x, y, w, h)
 *   - 図形の塗り色（HEX + 透明度）と枠線
 *   - テキストの内容
 *   - フォント・サイズ・太字・斜体
 *   - 文字色（HEX + 透明度）
 *   - 段落配置（START/CENTER/END/JUSTIFY）
 *   - 段落間隔（spaceAbove / spaceBelow）
 *   - 行間隔（lineSpacing）
 *   - インデント
 *   - shape contentAlignment（縦方向）
 *   - 回転角度
 *
 * 実行手順:
 *  1. このコードをApps Scriptに貼り付け
 *  2. inspectSlide4() を実行
 *  3. 実行ログ全文をIssyに共有
 */

function inspectSlide4() {
  var presentationId = "14fq9YY5dAOksIIT-H_d1qsiCwyJEnhAY3dGp-KZSO6U";
  var pres = SlidesApp.openById(presentationId);
  var slide = pres.getSlides()[3]; // index=3（=スライド4 = ユーザー再修正版）

  Logger.log("=== P4（再修正済みエグゼクティブサマリ）完全構造 ===");
  Logger.log("レイアウト: " + slide.getLayout().getLayoutName());
  Logger.log("");

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
      var shape = el.asShape();
      var phType = shape.getPlaceholderType();
      var phIdx  = shape.getPlaceholderIndex();
      Logger.log("  PlaceholderType=" + phType + ", PlaceholderIndex=" + phIdx);

      // 塗り色
      try {
        var fill = shape.getFill();
        if (fill.getType() === SlidesApp.FillType.SOLID) {
          var solidFill = fill.getSolidFill();
          var color = solidFill.getColor();
          var alpha = solidFill.getAlpha();  // 0.0〜1.0
          if (color.getColorType() === SlidesApp.ColorType.RGB) {
            var rgb = color.asRgbColor();
            Logger.log("  fill: #" + rgb.asHexString().replace("#","") +
              ", alpha=" + alpha);
          } else {
            Logger.log("  fill: THEME(" + color.asThemeColor().getThemeColorType() +
              "), alpha=" + alpha);
          }
        } else {
          Logger.log("  fill: " + fill.getType());
        }
      } catch(e) { Logger.log("  fill取得エラー: " + e.message); }

      // 枠線
      try {
        var border = shape.getBorder();
        if (border.isVisible()) {
          var lineFill = border.getLineFill();
          var weight = border.getWeight();
          if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
            var bcolor = lineFill.getSolidFill().getColor();
            if (bcolor.getColorType() === SlidesApp.ColorType.RGB) {
              Logger.log("  border: #" + bcolor.asRgbColor().asHexString().replace("#","") +
                ", weight=" + weight);
            }
          }
        }
      } catch(e) {}

      // 中身配置（縦方向）
      try {
        Logger.log("  contentAlign=" + shape.getContentAlignment());
      } catch(e) {}

      // テキスト
      var textRange = shape.getText();
      var fullText = textRange.asString();
      if (fullText.length > 0 && fullText !== "\n") {
        Logger.log("  テキスト全文: \"" + fullText.replace(/\n/g, "\\n") + "\"");

        // 段落ごとに詳細
        var paragraphs = textRange.getParagraphs();
        paragraphs.forEach(function(p, pi) {
          var pText = p.getRange().asString();
          if (pText.length === 0 || pText === "\n") return;

          var st = p.getRange().getTextStyle();
          var ps = p.getRange().getParagraphStyle();
          var info = "    [p" + pi + "] \"" + pText.substring(0, 40).replace(/\n/g, "\\n") + "\"";

          // フォント情報
          try { info += "\n      font=" + st.getFontFamily(); } catch(e) {}
          try { info += ", size=" + st.getFontSize(); } catch(e) {}
          try { info += ", bold=" + st.isBold(); } catch(e) {}
          try { info += ", italic=" + st.isItalic(); } catch(e) {}
          try { info += ", underline=" + st.isUnderline(); } catch(e) {}

          // 文字色
          try {
            var fc = st.getForegroundColor();
            if (fc) {
              if (fc.getColorType() === SlidesApp.ColorType.RGB) {
                info += "\n      color=#" + fc.asRgbColor().asHexString().replace("#","");
              } else {
                info += "\n      color=THEME(" + fc.asThemeColor().getThemeColorType() + ")";
              }
            }
          } catch(e) {}

          // 文字背景色
          try {
            var bc = st.getBackgroundColor();
            if (bc) {
              if (bc.getColorType() === SlidesApp.ColorType.RGB) {
                info += ", bgColor=#" + bc.asRgbColor().asHexString().replace("#","");
              }
            }
          } catch(e) {}

          // 段落配置（横）
          try {
            info += "\n      align=" + ps.getParagraphAlignment();
          } catch(e) {}

          // インデント
          try {
            var indentStart = ps.getIndentStart();
            var indentEnd = ps.getIndentEnd();
            var indentFirstLine = ps.getIndentFirstLine();
            if (indentStart || indentEnd || indentFirstLine) {
              info += "\n      indent: start=" + indentStart +
                ", end=" + indentEnd +
                ", firstLine=" + indentFirstLine;
            }
          } catch(e) {}

          // 行間隔（line spacing）
          try {
            var lineSpacing = ps.getLineSpacing();
            info += "\n      lineSpacing=" + lineSpacing;  // % (例: 100=1.0倍, 150=1.5倍)
          } catch(e) {}

          // 段落の前後スペース
          try {
            var spaceAbove = ps.getSpaceAbove();
            var spaceBelow = ps.getSpaceBelow();
            info += ", spaceAbove=" + spaceAbove + ", spaceBelow=" + spaceBelow;
          } catch(e) {}

          // 方向
          try {
            var direction = ps.getTextDirection();
            if (direction !== SlidesApp.TextDirection.LEFT_TO_RIGHT) {
              info += "\n      direction=" + direction;
            }
          } catch(e) {}

          Logger.log(info);
        });
      } else {
        Logger.log("  テキスト: (空)");
      }
    } else if (type === SlidesApp.PageElementType.LINE) {
      var line = el.asLine();
      try {
        var lineFill = line.getLineFill();
        if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
          var lc = lineFill.getSolidFill().getColor();
          if (lc.getColorType() === SlidesApp.ColorType.RGB) {
            Logger.log("  line color: #" + lc.asRgbColor().asHexString().replace("#",""));
          }
        }
        Logger.log("  weight=" + line.getWeight());
      } catch(e) {}
    } else if (type === SlidesApp.PageElementType.IMAGE) {
      Logger.log("  IMAGE");
    } else if (type === SlidesApp.PageElementType.TABLE) {
      var table = el.asTable();
      Logger.log("  TABLE rows=" + table.getNumRows() + ", cols=" + table.getNumColumns());
    }
    Logger.log("");  // 区切り
  });

  Logger.log("=== 完了 ===");
}
