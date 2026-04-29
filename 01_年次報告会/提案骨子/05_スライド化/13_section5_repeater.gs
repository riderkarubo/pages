/**
 * MCCM提案スライド化 - §補足: リピーターを増やす3つの戦略（1スライド）
 *
 * 最終更新: 2026-04-29 18:16
 *
 * マスター情報源: https://riderkarubo.github.io/pages/mccm-proposal/
 *  （これ以外の場所からは情報を拾わない）
 *
 * Issy手直し解析パターン準拠:
 *  - タイトル黒帯: x=9, y=20, w=671, h=45（newSlide()自動配置）
 *  - 強調色: ① C_ACCENT (#FA006D マゼンタ) / ② #C07000 オレンジ / ③ C_ACCENT2 (#1B998B 緑)
 *  - rectTxt() で1要素化を基本（複数色・複数サイズが必要なところだけレイヤー）
 *
 * 構成（1スライド完結・3カラム横並び）:
 *  S20: リピーターを増やす3つの戦略
 *    - ①「推し」でまた観たくなる
 *    - ②「おトク」でまた観たくなる
 *    - ③「いつもの」でまた観たくなる
 *
 * レイアウト座標:
 *  カード x: ① 25 / ② 252 / ③ 479（gap 12pt・w=215）
 *  カード y: 110、h: 215
 *
 * 修正履歴:
 *  - 2026-04-29 18:16: フィードバック反映
 *    - サブヘッダー削除（「出演者でまた観たくなる」など）
 *    - タグ3つ削除（●曜日担当制でレギュラー化など）
 *    - 戦略名を「①「推し」でまた観たくなる」の統合形式に変更
 *    - 狙い文を短縮版に書き換え
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
  var cardY = 125;
  var cardH = 215;
  var cardW = 215;
  var gap = 12;
  var cardX1 = 25;
  var cardX2 = cardX1 + cardW + gap;  // 252
  var cardX3 = cardX2 + cardW + gap;  // 479

  // 戦略①「推し」（マゼンタ #FA006D）
  drawRepeaterCard(s, cardX1, cardY, cardW, cardH, {
    no: "①",
    code: "STRATEGY 01",
    name: '①「推し」でまた観たくなる',
    core: "「あの人が出ているから観る」を作る",
    aim: "出演者をきっかけにライブを観に来てもらう。",
    accent: C_ACCENT,
    bgTint: "#FFF5F7"
  });

  // 戦略②「おトク」（オレンジ #C07000）
  drawRepeaterCard(s, cardX2, cardY, cardW, cardH, {
    no: "②",
    code: "STRATEGY 02",
    name: '②「おトク」でまた観たくなる',
    core: "「観ないと損」を仕組みで作る",
    aim: "顧客情報と連携できている強みを活かし、視聴をポイント・限定特典がもらえる場にする。",
    accent: "#C07000",
    bgTint: "#FFFAF0"
  });

  // 戦略③「いつもの」（グリーン #1B998B）
  drawRepeaterCard(s, cardX3, cardY, cardW, cardH, {
    no: "③",
    code: "STRATEGY 03",
    name: '③「いつもの」でまた観たくなる',
    core: "「火曜と言えば、いつものあれ」を作る",
    aim: "コスメに特化することで、視聴者が「毎週コスメの最新情報・新商品が知れる」と分かる構造にする。",
    accent: C_ACCENT2,
    bgTint: "#F0FAF8"
  });

  // ============================================================
  // フッター：3戦略を重ねがけし、リピーター比率 5.0% → 8.0% を達成する
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
 * 構成（縦方向・サブヘッダーとタグを廃止した簡潔版）:
 *  - カード本体（白＋枠線）
 *  - 左ボーダー（4pt幅・アクセント色）
 *  - 上部バッジ帯（薄ティント背景に番号＋STRATEGY 0X＋できっかけを作る）
 *  - 戦略名（「①「推し」でまた観たくなる」・bold・アクセント色）
 *  - コアコピー（「あの人が...」 など・bold・濃色）
 *  - 区切り線
 *  - 狙い文（短縮版・行間広め）
 *
 * @param slide
 * @param x, y, w, h
 * @param p {no, code, name, core, aim, accent, bgTint}
 */
function drawRepeaterCard(slide, x, y, w, h, p) {
  // カード本体（白＋グレー枠線）
  rect(slide, x, y, w, h, C_WHITE, C_BORDER, 1);
  // 左ボーダー（4pt幅・アクセント色）
  rect(slide, x, y, 4, h, p.accent);

  var pad = 12;
  var contentX = x + pad;
  var contentW = w - pad * 2;

  // 上部バッジ帯（薄ティント背景） y〜y+38
  rect(slide, x + 4, y, w - 4, 38, p.bgTint);

  // 番号 ① ② ③（左寄せ・大きめ・アクセント色）
  txt(slide, p.no, contentX, y + 8, 20, 22, {
    size: 18,
    color: p.accent,
    bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // STRATEGY 0X（番号の右隣・アクセント色）
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

  // 戦略名（中・bold・アクセント色） y+50〜y+76
  // 「①「推し」でまた観たくなる」を1行に収めるため14ptに調整
  txt(slide, p.name, contentX, y + 50, contentW, 26, {
    size: 14,
    color: p.accent,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // コアコピー（中・bold・濃色） y+82〜y+106
  txt(slide, p.core, contentX, y + 82, contentW, 24, {
    size: 12,
    color: C_DARK,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 区切り線 y+114
  rect(slide, contentX + 30, y + 114, contentW - 60, 1, C_BORDER);

  // 狙い文（短縮版・行間広め） y+124〜y+205
  txt(slide, p.aim, contentX, y + 124, contentW, 80, {
    size: 11,
    color: C_TEXT,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.TOP
  });
}
