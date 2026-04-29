/**
 * MCCM提案スライド化 - §補足: リピーターを増やす3つの戦略（1スライド）
 *
 * マスター情報源: https://riderkarubo.github.io/pages/mccm-proposal/
 *  （これ以外の場所からは情報を拾わない）
 *
 * Issy手直し解析パターン準拠:
 *  - タイトル黒帯: x=9, y=20, w=671, h=45（newSlide()自動配置）
 *  - 結論メッセージ相当: drawConclusionMessage() で y≈90、16pt bold #434343 + 強調語だけアクセント色
 *  - 強調色: ① C_ACCENT (#FA006D マゼンタ) / ② #C07000 オレンジ / ③ C_ACCENT2 (#1B998B 緑)
 *  - rectTxt() で1要素化を基本（複数色・複数サイズが必要なところだけレイヤー）
 *
 * 構成（1スライド完結・3カラム横並び）:
 *  S20: リピーターを増やす3つの戦略
 *    - ①「推し」できっかけを作る ― 出演者でまた観たくなる
 *    - ②「おトク」できっかけを作る ― おトクでまた観たくなる
 *    - ③「いつもの」できっかけを作る ― いつもの番組としてまた観たくなる
 *
 * レイアウト座標:
 *  カード x: ① 25 / ② 252 / ③ 479（gap 12pt・w=215）
 *  カード y: 110、h: 250
 *
 * 実行手順:
 *  1. 02_helpers.gs を最新版で貼り付け済み
 *  2. このファイルを Apps Script に新規追加
 *  3. insertSection5Repeater() を実行 → 末尾に1枚追加
 */

function insertSection5Repeater() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  Logger.log("§補足 リピーター戦略 開始 - 現在: " + pres.getSlides().length + "枚");

  insertS20_repeaterStrategy(pres);
  Logger.log("S20 完了");

  Logger.log("§補足 完了 - 現在: " + pres.getSlides().length + "枚");
}

function runS20Only() { insertS20_repeaterStrategy(SlidesApp.openById(PRESENTATION_ID)); }

// ============================================================
// S20: リピーターを増やす3つの戦略（1スライド）
// ============================================================
function insertS20_repeaterStrategy(pres) {
  var s = newSlide(pres, "リピーターを増やす3つの戦略");

  // ============================================================
  // サブタイトル（小見出し領域・y=76, h=34）
  // ============================================================
  txt(s, '視聴者が "また観たくなる理由" を3つ用意する',
      SUBTITLE_X, SUBTITLE_Y, SUBTITLE_W, SUBTITLE_H, {
    size: 12,
    color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // ============================================================
  // 3カード（横並び）
  // ============================================================
  var cardY = 115;
  var cardH = 245;
  var cardW = 215;
  var gap = 12;
  var cardX1 = 25;
  var cardX2 = cardX1 + cardW + gap;  // 252
  var cardX3 = cardX2 + cardW + gap;  // 479

  // 戦略①「推し」（マゼンタ #FA006D）
  drawRepeaterCard(s, cardX1, cardY, cardW, cardH, {
    no: "①",
    code: "STRATEGY 01",
    name: "「推し」",
    sub: "出演者でまた観たくなる",
    core: "「あの人が出ているから観る」を作る",
    aim: "商品を見に来る視聴者から、人を見に来る視聴者へ転換。リピート理由を商品ラインナップではなく出演者のキャラクターに固定する。",
    tags: [
      "曜日担当制でレギュラー化",
      "出演者個人SNSの強化",
      "出演者プロフィールページ"
    ],
    accent: C_ACCENT,
    bgTint: "#FFF5F7"
  });

  // 戦略②「おトク」（オレンジ #C07000）
  drawRepeaterCard(s, cardX2, cardY, cardW, cardH, {
    no: "②",
    code: "STRATEGY 02",
    name: "「おトク」",
    sub: "おトクでまた観たくなる",
    core: "「観ないと損」を仕組みで作る",
    aim: "「ふらっと来た人」から、会員ログインして観る人へ転換。CDP×マツキヨアプリ×全国会員基盤を活用し、視聴をポイント・限定特典の場に位置付ける(マツキヨ独自)。",
    tags: [
      "会員ログイン視聴特典の常設化",
      "ライブ限定SKU・バンドル",
      "CDP起点の1to1通知"
    ],
    accent: "#C07000",
    bgTint: "#FFFAF0"
  });

  // 戦略③「いつもの」（グリーン #1B998B）
  drawRepeaterCard(s, cardX3, cardY, cardW, cardH, {
    no: "③",
    code: "STRATEGY 03",
    name: "「いつもの」",
    sub: "いつもの番組としてまた観たくなる",
    core: "「火曜と言えば、いつものあれ」を作る",
    aim: "毎回ジャンルが変わる現状から、コスメに特化した定番番組へ転換。視聴者が「来週も同じ世界観のものが見られる」と分かって戻ってくる構造を作る。",
    tags: [
      "コスメ特化フォーマットの確立",
      "「悩み診断」コーナー導入",
      "「火曜ナイトコスメ」番組ブランド化"
    ],
    accent: C_ACCENT2,
    bgTint: "#F0FAF8"
  });

  // ============================================================
  // フッター：3戦略を重ねがけし、リピーター比率 5.0% → 8.0% を達成する
  //   ※ page-num領域（y=391付近）の手前に配置
  // ============================================================
  var footerY = 370;
  var footerH = 18;
  var footerText = "▼ 3戦略を重ねがけし、リピーター比率 5.0% → 8.0% を達成する";
  var footerBox = txt(s, footerText, 25, footerY, 671, footerH, {
    size: 11,
    color: C_NOTE,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  // 「5.0% → 8.0%」をマゼンタ強調
  try {
    var tr = footerBox.getText();
    var hl = "5.0% → 8.0%";
    var idx = footerText.indexOf(hl);
    if (idx >= 0) {
      tr.getRange(idx, idx + hl.length).getTextStyle()
        .setForegroundColor(C_ACCENT).setBold(true);
    }
  } catch(e) {}
}

/**
 * リピーター戦略カード描画（縦長カード・3カラム並べる前提）
 *
 * 構成（縦方向）:
 *  - カード本体（白＋枠線）
 *  - 左ボーダー（4pt幅・アクセント色）
 *  - 上部バッジ帯（薄ティント背景に番号＋STRATEGY 0X）
 *  - 戦略名（「推し」など・大きく・bold・アクセント色）
 *  - サブヘッダー（出演者でまた観たくなる など・グレー）
 *  - コアコピー（「あの人が...」 など・bold・濃色）
 *  - 区切り線（薄グレー）
 *  - 狙い文（小さめ・行間広め）
 *  - タグ3つ（●先頭）
 *
 * @param slide
 * @param x, y, w, h
 * @param p {no, code, name, sub, core, aim, tags[], accent, bgTint}
 */
function drawRepeaterCard(slide, x, y, w, h, p) {
  // カード本体（白＋グレー枠線）
  rect(slide, x, y, w, h, C_WHITE, C_BORDER, 1);
  // 左ボーダー（4pt幅・アクセント色）
  rect(slide, x, y, 4, h, p.accent);

  var pad = 12;
  var contentX = x + pad;
  var contentW = w - pad * 2;

  // 上部バッジ帯（薄ティント背景） y+8〜y+30
  rect(slide, x + 4, y, w - 4, 38, p.bgTint);

  // 番号 ① ② ③（左寄せ・大きめ・アクセント色）
  txt(slide, p.no, contentX, y + 8, 20, 22, {
    size: 18,
    color: p.accent,
    bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // STRATEGY 0X（番号の右隣・グレー小文字）
  txt(slide, p.code, contentX + 22, y + 12, contentW - 22, 14, {
    size: 8,
    color: p.accent,
    bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // できっかけを作る（番号の右隣・小・グレー）
  txt(slide, "できっかけを作る", contentX + 22, y + 24, contentW - 22, 12, {
    size: 8,
    color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 戦略名（大・bold・アクセント色） y+44〜y+72
  txt(slide, p.name, contentX, y + 44, contentW, 28, {
    size: 22,
    color: p.accent,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // サブヘッダー（小・グレー） y+76〜y+90
  txt(slide, p.sub, contentX, y + 76, contentW, 14, {
    size: 9,
    color: C_GRAY,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // コアコピー（中・bold・濃色） y+94〜y+114
  txt(slide, p.core, contentX, y + 94, contentW, 20, {
    size: 11,
    color: C_DARK,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 区切り線 y+120
  rect(slide, contentX + 30, y + 120, contentW - 60, 1, C_BORDER);

  // 狙い文（小・行間広め） y+128〜y+200
  txt(slide, p.aim, contentX, y + 128, contentW, 70, {
    size: 9,
    color: C_TEXT,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.TOP
  });

  // タグ3つ（●先頭） y+200〜y+240
  var tagY = y + 200;
  var tagH = 13;
  p.tags.forEach(function(tag, i) {
    txt(slide, "● " + tag, contentX, tagY + i * tagH, contentW, tagH, {
      size: 8.5,
      color: p.accent,
      bold: true,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  });
}
