/**
 * MCCM提案スライド化 - Phase 0+: 既存スライド2の詳細構造を取得
 *
 * 用途:
 *  既存スライド2（複製ベース・TITLE_AND_BODY_2_1）の全要素の座標・サイズ・テキスト・スタイルを取得する。
 *  この結果から、複製後の操作対象を正しく特定できるようにする。
 *
 * 実行手順:
 *  1. 既に Apps Script に貼ってあるコードに「追加」する形でこの関数を貼り付け
 *  2. inspectSlide2() を選んで実行
 *  3. 実行ログ全文をIssyに共有
 */

function inspectSlide2() {
  var presentationId = "14fq9YY5dAOksIIT-H_d1qsiCwyJEnhAY3dGp-KZSO6U";
  var pres = SlidesApp.openById(presentationId);
  var slide = pres.getSlides()[1]; // index=1（=スライド2）

  Logger.log("=== スライド2の全要素 ===");
  Logger.log("レイアウト: " + slide.getLayout().getLayoutName());

  var elements = slide.getPageElements();
  elements.forEach(function(el, i) {
    var type = el.getPageElementType();
    Logger.log("\n[" + i + "] type=" + type +
      ", x=" + Math.round(el.getLeft()) +
      ", y=" + Math.round(el.getTop()) +
      ", w=" + Math.round(el.getWidth()) +
      ", h=" + Math.round(el.getHeight()));

    if (type === SlidesApp.PageElementType.SHAPE) {
      var shape = el.asShape();
      var phType = shape.getPlaceholderType();
      var phIdx  = shape.getPlaceholderIndex();
      Logger.log("  PlaceholderType=" + phType + ", PlaceholderIndex=" + phIdx);

      var text = shape.getText().asString();
      var preview = text.substring(0, 60).replace(/\n/g, "\\n");
      Logger.log("  text=\"" + preview + (text.length > 60 ? "..." : "") + "\"");

      if (text.length > 0) {
        shape.getText().getParagraphs().forEach(function(p, pi) {
          if (pi > 1) return;
          var st = p.getRange().getTextStyle();
          Logger.log("  p" + pi + ": fontFamily=" + st.getFontFamily() +
                     ", size=" + st.getFontSize() +
                     ", bold=" + st.isBold() +
                     ", color=" + st.getForegroundColor());
        });
      }
    } else if (type === SlidesApp.PageElementType.TABLE) {
      var table = el.asTable();
      Logger.log("  TABLE rows=" + table.getNumRows() + ", cols=" + table.getNumColumns());
    } else if (type === SlidesApp.PageElementType.IMAGE) {
      Logger.log("  IMAGE");
    } else if (type === SlidesApp.PageElementType.LINE) {
      Logger.log("  LINE");
    }
  });

  Logger.log("\n=== レイアウト由来のプレースホルダー要素 ===");
  // レイアウトが提供している要素を確認（複製時にこれらが継承される）
  var layout = slide.getLayout();
  var layoutShapes = layout.getShapes();
  layoutShapes.forEach(function(shape, i) {
    var phType = shape.getPlaceholderType();
    if (phType !== SlidesApp.PlaceholderType.NONE) {
      Logger.log("[layout-" + i + "] PH type=" + phType +
        ", index=" + shape.getPlaceholderIndex() +
        ", x=" + Math.round(shape.getLeft()) +
        ", y=" + Math.round(shape.getTop()) +
        ", w=" + Math.round(shape.getWidth()) +
        ", h=" + Math.round(shape.getHeight()));
    }
  });
}
