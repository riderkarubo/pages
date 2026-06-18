/**
 * MCCM提案スライド化 - Phase 0++: 修正後のスライド3（P3）の詳細構造を取得
 *
 * 用途:
 *  ユーザーが手動で整えた P3（エグゼクティブサマリ）のレイアウトを取得し、
 *  そのレイアウトを他のKPIブロックに踏襲するための情報を得る。
 *
 * 実行手順:
 *  1. このコードを Apps Script に貼り付け
 *  2. inspectSlide3() を実行
 *  3. 実行ログ全文をIssyに共有
 */

function inspectSlide3() {
  var presentationId = "14fq9YY5dAOksIIT-H_d1qsiCwyJEnhAY3dGp-KZSO6U";
  var pres = SlidesApp.openById(presentationId);
  var slide = pres.getSlides()[2]; // index=2（=スライド3 = エグゼクティブサマリ・修正済み）

  Logger.log("=== スライド3（P3 修正済みエグゼクティブサマリ）の全要素 ===");
  Logger.log("レイアウト: " + slide.getLayout().getLayoutName());
  Logger.log("");

  var elements = slide.getPageElements();
  Logger.log("要素数: " + elements.length + "個");
  Logger.log("");

  elements.forEach(function(el, i) {
    var type = el.getPageElementType();
    Logger.log("[" + i + "] type=" + type +
      ", x=" + Math.round(el.getLeft()) +
      ", y=" + Math.round(el.getTop()) +
      ", w=" + Math.round(el.getWidth()) +
      ", h=" + Math.round(el.getHeight()));

    if (type === SlidesApp.PageElementType.SHAPE) {
      var shape = el.asShape();
      var phType = shape.getPlaceholderType();
      var phIdx  = shape.getPlaceholderIndex();
      Logger.log("  PlaceholderType=" + phType + ", PlaceholderIndex=" + phIdx);

      // 図形の塗り色
      try {
        var fill = shape.getFill();
        if (fill.getType() === SlidesApp.FillType.SOLID) {
          var color = fill.getSolidFill().getColor();
          if (color.getColorType() === SlidesApp.ColorType.RGB) {
            var rgb = color.asRgbColor();
            Logger.log("  fill=#" + rgb.asHexString());
          } else {
            Logger.log("  fill=THEME(" + color.asThemeColor().getThemeColorType() + ")");
          }
        } else {
          Logger.log("  fill=" + fill.getType());
        }
      } catch(e) { Logger.log("  fill取得エラー: " + e.message); }

      // テキスト内容とスタイル
      var text = shape.getText().asString();
      if (text.length > 0) {
        var preview = text.substring(0, 80).replace(/\n/g, "\\n");
        Logger.log("  text=\"" + preview + (text.length > 80 ? "..." : "") + "\"");

        shape.getText().getParagraphs().forEach(function(p, pi) {
          if (pi > 3) return;  // 最初の4段落まで
          var pText = p.getRange().asString().replace(/\n/g, "\\n");
          if (pText.length === 0) return;
          var st = p.getRange().getTextStyle();
          var ps = p.getRange().getParagraphStyle();
          var info = "    p" + pi + ": \"" + pText.substring(0, 30) + "\"";
          info += " | size=" + st.getFontSize();
          info += ", bold=" + st.isBold();
          info += ", font=" + st.getFontFamily();
          try {
            var fc = st.getForegroundColor();
            if (fc) {
              if (fc.getColorType() === SlidesApp.ColorType.RGB) {
                info += ", color=#" + fc.asRgbColor().asHexString();
              } else {
                info += ", color=THEME";
              }
            }
          } catch(e) {}
          try {
            info += ", align=" + ps.getParagraphAlignment();
          } catch(e) {}
          Logger.log(info);
        });
      } else {
        Logger.log("  text=(空)");
      }

      // 中身配置
      try {
        Logger.log("  contentAlign=" + shape.getContentAlignment());
      } catch(e) {}
    }
  });

  Logger.log("");
  Logger.log("=== 完了 ===");
}
