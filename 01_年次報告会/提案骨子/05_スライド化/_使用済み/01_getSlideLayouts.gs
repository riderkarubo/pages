/**
 * MCCM提案スライド化 - Phase 0: テンプレ調査
 *
 * 用途:
 *  出力先Slidesのマスター・レイアウト・既存スライドを調査し、
 *  どのスライドを「複製ベース」にするかを決めるための情報を取得する。
 *
 * 実行手順:
 *  1. 対象のGoogleスライドを開く
 *  2. 「拡張機能」→「Apps Script」をクリック
 *  3. このコードをまるごと貼り付け
 *  4. 関数 getSlideLayouts() を実行（▶ 実行ボタン）
 *  5. 初回は権限承認のポップアップが出るので許可
 *  6. 「実行ログ」（時計アイコン）に結果が出るので、その全文をIssyに共有
 *
 * Issyに共有してほしい情報:
 *  - 「=== スライドマスター一覧 ===」以降のすべてのログ
 *  - 「=== 既存スライド一覧 ===」以降のすべてのログ
 *
 * 共有方法: ログの「コピー」ボタン → チャットに貼り付け
 */

function getSlideLayouts() {
  var presentationId = "14fq9YY5dAOksIIT-H_d1qsiCwyJEnhAY3dGp-KZSO6U";
  var presentation = SlidesApp.openById(presentationId);
  var masters = presentation.getMasters();

  Logger.log("=== スライドマスター一覧 ===");
  masters.forEach(function(master, mIdx) {
    Logger.log("マスター " + mIdx + ": " + master.getObjectId());
    var layouts = master.getLayouts();
    layouts.forEach(function(layout, lIdx) {
      Logger.log("  レイアウト " + lIdx + ": " + layout.getLayoutName() + " (ID: " + layout.getObjectId() + ")");
      var shapes = layout.getShapes();
      shapes.forEach(function(shape, sIdx) {
        var phType = shape.getPlaceholderType();
        if (phType !== SlidesApp.PlaceholderType.NONE) {
          Logger.log("    PH " + sIdx + ": type=" + phType + ", index=" + shape.getPlaceholderIndex() + ", ID=" + shape.getObjectId());
        }
      });
    });
  });

  Logger.log("\n=== 既存スライド一覧 ===");
  var slides = presentation.getSlides();
  slides.forEach(function(slide, sIdx) {
    var layout = slide.getLayout();
    Logger.log("スライド " + (sIdx + 1) + ": レイアウト=" + layout.getLayoutName() + ", ID=" + slide.getObjectId());
  });
}
