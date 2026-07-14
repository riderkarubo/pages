// 最終更新: 2026年7月14日 11:17
// 2026年4-6月期 四半期報告会デッキ トンマナ実測スクリプト
// 対象: https://docs.google.com/presentation/d/1LI0mtiSuAXpexHHxNTqDAbHndN7f5GdWY6gMRKH_ZCc/edit
// 用途: 全スライドの座標・fill・フォント・テキストを実測し、Fireworkトンマナ（C_ACCENT=#FA006D等）と照合する
// 実行方法: このファイルをApps Scriptエディタに貼り付け → inspectQuarterlyDeck を実行 → 実行ログ（表示→ログ）を共有

var Q4_INSPECT_PRES_ID = '1LI0mtiSuAXpexHHxNTqDAbHndN7f5GdWY6gMRKH_ZCc';

function inspectQuarterlyDeck() {
  var pres = SlidesApp.openById(Q4_INSPECT_PRES_ID);
  var slides = pres.getSlides();
  Logger.log('=== 総スライド数: ' + slides.length + ' ===');

  slides.forEach(function (slide, idx) {
    Logger.log('\n########## Slide ' + (idx + 1) + ' (objectId=' + slide.getObjectId() + ') ##########');
    var elements = slide.getPageElements();
    elements.forEach(function (el, i) {
      try {
        var type = el.getPageElementType();
        var x = Math.round(el.getLeft());
        var y = Math.round(el.getTop());
        var w = Math.round(el.getWidth());
        var h = Math.round(el.getHeight());
        Logger.log('  [' + i + '] type=' + type + ' x=' + x + ' y=' + y + ' w=' + w + ' h=' + h);

        if (type === SlidesApp.PageElementType.SHAPE) {
          var shape = el.asShape();
          var fill = shape.getFill();
          var fillType = fill ? fill.getType() : null;
          var fillColor = '';
          try {
            if (fillType === SlidesApp.FillType.SOLID) {
              fillColor = fill.getSolidFill().getColor().asRgbColor().asHexString();
            }
          } catch (e) {}
          Logger.log('      fillType=' + fillType + ' fillColor=' + fillColor + ' placeholderType=' + shape.getPlaceholderType());

          var textRange = shape.getText();
          var textStr = textRange.asString();
          if (textStr && textStr.trim().length > 0) {
            Logger.log('      TEXT: "' + textStr.replace(/\n/g, ' / ').substring(0, 120) + '"');
            var paragraphs = textRange.getParagraphs();
            paragraphs.forEach(function (p, pi) {
              var runs = p.getRange().getRuns();
              runs.forEach(function (run) {
                var style = run.getTextStyle();
                var runText = run.asString();
                if (!runText || runText.trim().length === 0) return;
                var fontFamily = '';
                var fontSize = '';
                var bold = '';
                var color = '';
                try { fontFamily = style.getFontFamily(); } catch (e) {}
                try { fontSize = style.getFontSize(); } catch (e) {}
                try { bold = style.isBold(); } catch (e) {}
                try { color = style.getForegroundColor().getColor().asRgbColor().asHexString(); } catch (e) {}
                Logger.log('        run p' + pi + ': font=' + fontFamily + ' size=' + fontSize + ' bold=' + bold + ' color=' + color + ' text="' + runText.substring(0, 40) + '"');
              });
            });
          }
        } else if (type === SlidesApp.PageElementType.TABLE) {
          var table = el.asTable();
          Logger.log('      TABLE rows=' + table.getNumRows() + ' cols=' + table.getNumColumns());
        }
      } catch (e) {
        Logger.log('  [' + i + '] ERROR: ' + e);
      }
    });
  });

  Logger.log('\n=== 実測完了。上記ログを全てコピーして共有してください ===');
}
