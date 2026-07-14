// 最終更新: 2026年7月14日 11:30
//
// MCCM 2026年4-6月期 四半期報告会デッキ
// ① P4「協賛配信 受注目標」スライドを元の状態にロールバック
// ② P4を複製してP5「協賛配信 受注進捗と今後の見通し」を新規デザインで作成
//
// 対象プレゼン: https://docs.google.com/presentation/d/1LI0mtiSuAXpexHHxNTqDAbHndN7f5GdWY6gMRKH_ZCc/edit
// 注意: 同一Apps Scriptプロジェクト内の 02_helpers.js に PRESENTATION_ID / C_ACCENT 等の
//       グローバル変数・ヘルパー（txt/rect/rectTxt/drawKpiCard/drawConclusionMessage/makeTable等）が
//       既に定義されているため、本ファイルでは再宣言せず、対象プレゼンIDのみローカル変数として持つ。
//
// 実行方法: Apps Scriptエディタで rollbackP4AndCreateP5() を実行 → ログで結果確認 → Slidesを確認

var Q4_TARGET_PRES_ID = '1LI0mtiSuAXpexHHxNTqDAbHndN7f5GdWY6gMRKH_ZCc';

function rollbackP4AndCreateP5() {
  rollbackP4();
  createP5KyosanForecast();
}

// ============================================================
// ① P4ロールバック（前回追加した2BOXを削除・進捗テキストを復元）
// ============================================================
function rollbackP4() {
  var pres = SlidesApp.openById(Q4_TARGET_PRES_ID);
  var targetSlide = findP4Slide(pres);
  if (!targetSlide) {
    Logger.log('エラー: P4（協賛配信 受注目標）が見つかりませんでした');
    return;
  }

  var shapes = targetSlide.getShapes();
  var removed = 0;
  shapes.forEach(function (shape) {
    var text = '';
    try { text = shape.getText().asString(); } catch (e) {}

    if (text.indexOf('残り枠') !== -1 && text.indexOf('26枠') !== -1) {
      shape.remove();
      removed++;
      return;
    }
    if (text.indexOf('横断型配信') !== -1 && text.indexOf('射程圏内') !== -1) {
      shape.remove();
      removed++;
      return;
    }
    if (text.indexOf('進捗') !== -1 && text.indexOf('件') !== -1) {
      shape.getText().setText('進捗：14件 655万円（4/30時点）');
      shape.getText().getParagraphs().forEach(function (p) {
        var st = p.getRange().getTextStyle();
        st.setFontFamily('Noto Sans JP');
        st.setFontSize(11);
        st.setForegroundColor(C_TEXT);
      });
    }
  });
  Logger.log('P4ロールバック完了: 削除' + removed + '件・進捗テキストを4/30時点表記に復元');
}

// ============================================================
// ② P5新規作成（協賛配信 受注進捗と今後の見通し）
// ============================================================
function createP5KyosanForecast() {
  var pres = SlidesApp.openById(Q4_TARGET_PRES_ID);
  var slides = pres.getSlides();
  var p4Slide = findP4Slide(pres);
  if (!p4Slide) {
    Logger.log('エラー: P4（複製元）が見つかりませんでした');
    return;
  }
  var p4Idx = slides.indexOf(p4Slide);

  // 既に同名P5が存在するか確認（再実行時の重複防止）
  for (var i = 0; i < slides.length; i++) {
    var shapes = slides[i].getShapes();
    for (var j = 0; j < shapes.length; j++) {
      var t = '';
      try { t = shapes[j].getText().asString(); } catch (e) {}
      if (t.indexOf('協賛配信 受注進捗と今後の見通し') !== -1) {
        Logger.log('既にP5相当のスライドが存在します（index=' + i + '）。重複作成を避けるため中断します。手動で削除してから再実行してください。');
        return;
      }
    }
  }

  // P4を複製 → 直後に配置
  var p5 = p4Slide.duplicate();
  p5.move(p4Idx + 1);

  // 複製後の要素を一旦全部削除（一から作り直す）
  var els = p5.getPageElements();
  for (var k = els.length - 1; k >= 0; k--) {
    try { els[k].remove(); } catch (e) {}
  }

  // --- タイトル黒帯 ---
  setTitle(p5, '協賛配信 受注進捗と今後の見通し');

  // --- 結論メッセージ（強調語を C_ACCENT で部分強調） ---
  var conclusionText = '残り26枠のうち、横断型配信あと5枠獲得で売上目標2,400万円は達成射程圏内';
  drawConclusionMessage(p5, 90, conclusionText, '横断型配信あと5枠獲得');

  // --- KPIカード×2（横並び） ---
  var cardY = 125, cardH = 90;
  var cardGap = 16;
  var cardW = (BODY_W - cardGap) / 2;
  var card1X = BODY_X;
  var card2X = BODY_X + cardW + cardGap;

  drawKpiCard(p5, card1X, cardY, cardW, cardH,
    '受注件数（2026年度累計）', '18', '/ 48件',
    '進捗 37.5%（7/14時点）', C_ACCENT3);

  drawKpiCard(p5, card2X, cardY, cardW, cardH,
    '売上換算（2026年度累計）', '920', '/ 2,400万円',
    '進捗 38.3%（7/14時点）', C_ACCENT);

  // --- 残り枠インフォバー ---
  var barY = cardY + cardH + 10; // 225
  rectTxt(
    p5,
    '残り枠：26枠（年度内・火曜配信ベース、祝日5件除く） ／ 件数上限は44件（枠制約により48件到達は困難）',
    BODY_X, barY, BODY_W, 24,
    C_SOFT,
    { size: 9, color: C_GRAY, align: SlidesApp.ParagraphAlignment.START, bold: false },
    C_BORDER, 0.5
  );

  // --- シナリオテーブル（横断型比率別の売上着地見込み・達成ライン行をハイライト） ---
  var tableY = barY + 34; // 259
  var colDefs = [
    { label: '横断型比率', align: SlidesApp.ParagraphAlignment.CENTER },
    { label: '内訳（横断型／1社協賛等）', align: SlidesApp.ParagraphAlignment.CENTER },
    { label: '着地見込み売上', align: SlidesApp.ParagraphAlignment.CENTER },
    { label: '対目標', align: SlidesApp.ParagraphAlignment.CENTER }
  ];
  var rows = [
    ['0%', '0枠／26枠', '2,220万円', '92.5%'],
    ['約19%（5枠）', '5枠／21枠', '約2,400万円', '達成ライン'],
    ['100%', '26枠／0枠', '3,260万円', '135.8%']
  ];
  makeTable(p5, BODY_X, tableY, BODY_W, colDefs, rows, {
    rowH: 22, hdrH: 26,
    headerBg: C_ACCENT3, headerSize: 9, bodySize: 9,
    highlightRows: [1]
  });

  Logger.log('P5作成完了: index=' + (p4Idx + 1) + ' (P' + (p4Idx + 2) + ')');
}

// ============================================================
// 共通: 「協賛配信」+「受注目標」を含むスライド（P4）を検索
// ============================================================
function findP4Slide(pres) {
  var slides = pres.getSlides();
  for (var i = 0; i < slides.length; i++) {
    var shapes = slides[i].getShapes();
    for (var j = 0; j < shapes.length; j++) {
      var text = '';
      try { text = shapes[j].getText().asString(); } catch (e) {}
      if (text.indexOf('協賛配信') !== -1 && text.indexOf('受注目標') !== -1) {
        return slides[i];
      }
    }
  }
  return null;
}
