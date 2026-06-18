/**
 * MCCM提案スライド化 - §4: 2026年度の2本柱（7スライド・v2 実測テンプレ準拠版）
 *
 * 最終更新: 2026-04-28 01:52
 * マスター情報源: https://riderkarubo.github.io/pages/mccm-proposal/
 *  （これ以外の場所からは情報を拾わない）
 *
 * Issy手直し解析パターン準拠:
 *  - タイトル黒帯: x=9, y=20, w=671, h=45（newSlide()自動配置）
 *  - 結論メッセージ: drawConclusionMessage() で y≈90, 16pt bold #434343 + 強調語だけ #FA006D
 *  - 強調色: 柱① = C_ACCENT (#FA006D マゼンタ) / 柱② = C_ACCENT2 (#1B998B 緑)
 *  - rectTxt() で1要素化を基本（複数色・複数サイズ・インライン強調が必要なところだけレイヤー）
 *  - ※注釈は8pt統一
 *
 * 構成（マスター完全踏襲・グレーアウト部除外）:
 *  S13: 2026年度の2本柱（章扉・サマリーカード2枚）
 *  S14: 柱① 現状課題（①価格 ②粗利率 の2課題）
 *  S15: 柱① 提案①：単価是正（A社比較・粗利率インパクト/広告換算は除外）
 *  S16: 柱① 体制改善：現状課題（工数膨張・薬機法CK詰まり）
 *  S17: 柱① 体制改善 上半期①：レギュレーション設定（プレースホルダー含む）
 *  S18: 柱① 体制改善：薬機法CK + AI活用（上半期② + 下半期③）
 *  S19: 柱② 社員インフルエンサー強化（4アクション）
 *
 * 除外項目（グレーアウト・打ち消し線=社内検討中）:
 *  - 提案②（セールス訴求強化・50万円内訳）
 *  - 提案③（メニューラインナップ3層構造）
 *  - 65万円・年間プラン
 *  - 粗利率インパクト・広告換算（CPM/CPV比較）
 *
 * 実行手順:
 *  1. 02_helpers.gs を最新版で貼り付け済み
 *  2. このファイルを Apps Script に新規追加
 *  3. insertSection4V2() を実行 → 末尾に7枚追加
 *
 * 単独実行用ラッパー:
 *  runS13Only() / runS14Only() / runS15Only() / runS16Only() / runS17Only() / runS18Only() / runS19Only()
 */

function insertSection4V2() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  Logger.log("§4 v2 開始 - 現在: " + pres.getSlides().length + "枚");

  insertS13v2_pillarsCover(pres);     Logger.log("S13 完了");
  insertS14v2_pillar1Issues(pres);    Logger.log("S14 完了");
  insertS15v2_priceRevision(pres);    Logger.log("S15 完了");
  insertS16v2_opsIssues(pres);        Logger.log("S16 完了");
  insertS17v2_regulation(pres);       Logger.log("S17 完了");
  insertS18v2_yakkihouAI(pres);       Logger.log("S18 完了");
  insertS19v2_employeeInfluencer(pres); Logger.log("S19 完了");

  Logger.log("§4 v2 完了 - 現在: " + pres.getSlides().length + "枚");
}

function runS13Only() { insertS13v2_pillarsCover(SlidesApp.openById(PRESENTATION_ID)); }
function runS14Only() { insertS14v2_pillar1Issues(SlidesApp.openById(PRESENTATION_ID)); }
function runS15Only() { insertS15v2_priceRevision(SlidesApp.openById(PRESENTATION_ID)); }
function runS16Only() { insertS16v2_opsIssues(SlidesApp.openById(PRESENTATION_ID)); }
function runS17Only() { insertS17v2_regulation(SlidesApp.openById(PRESENTATION_ID)); }
function runS18Only() { insertS18v2_yakkihouAI(SlidesApp.openById(PRESENTATION_ID)); }
function runS19Only() { insertS19v2_employeeInfluencer(SlidesApp.openById(PRESENTATION_ID)); }

// ============================================================
// S13: 2026年度の2本柱（章扉）
// ============================================================
function insertS13v2_pillarsCover(pres) {
  var s = newSlide(pres, "2026年度の2本柱");

  // 結論メッセージ
  drawConclusionMessage(s, 88,
    "KPI（協賛配信受注 年間48件・2,400万円）達成に向けた2つの重点施策",
    "2つの重点施策");

  // 2カード（柱①=#FA006D / 柱②=#1B998B）
  var cardY = 145;
  var cardH = 200;
  var cardW = 320;
  var gap = 20;
  var startX = (SW - 2 * cardW - gap) / 2;  // 中央配置

  // 柱①
  var c1x = startX;
  rect(s, c1x, cardY, cardW, cardH, "#FFF5F7", C_ACCENT, 2);
  rect(s, c1x, cardY, cardW, 6, C_ACCENT);  // 上ライン強調
  txt(s, "柱①", c1x + 20, cardY + 18, cardW - 40, 18, {
    size: 11, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "協賛配信のメニュー・体制改善", c1x + 20, cardY + 50, cardW - 40, 30, {
    size: 18, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "受注後の運営体制を標準化し、\n月4回以上の安定受注を実現する。",
    c1x + 20, cardY + 95, cardW - 40, 80, {
      size: 12, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.START
    });

  // 柱②
  var c2x = startX + cardW + gap;
  rect(s, c2x, cardY, cardW, cardH, "#F0FAF8", C_ACCENT2, 2);
  rect(s, c2x, cardY, cardW, 6, C_ACCENT2);
  txt(s, "柱②", c2x + 20, cardY + 18, cardW - 40, 18, {
    size: 11, color: C_ACCENT2, bold: true,
    align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "社員インフルエンサー強化", c2x + 20, cardY + 50, cardW - 40, 30, {
    size: 18, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "出演者を育成・サポートし、\nリピーター比率の底上げを実現する。",
    c2x + 20, cardY + 95, cardW - 40, 80, {
      size: 12, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.START
    });
}

// ============================================================
// S14: 柱① 現状課題（価格・粗利率）
// ============================================================
function insertS14v2_pillar1Issues(pres) {
  var s = newSlide(pres, "柱① 協賛配信のメニュー・体制改善 ― 現状課題");

  // 結論メッセージ
  drawConclusionMessage(s, 88,
    "現状の協賛配信50万円は、価格と粗利率の両面で見直し余地がある",
    "見直し余地がある");

  // 2カード横並び
  var cardY = 130;
  var cardH = 230;
  var cardW = 325;
  var gap = 18;
  var startX = (SW - 2 * cardW - gap) / 2;

  // 課題①
  var c1x = startX;
  rect(s, c1x, cardY, cardW, cardH, C_WHITE, C_ACCENT, 2);
  rect(s, c1x, cardY, 5, cardH, C_ACCENT);
  txt(s, "現状課題 ①", c1x + 20, cardY + 14, cardW - 40, 14, {
    size: 10, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "競合他社と比べて\n提供価値に対する価格が低い",
    c1x + 20, cardY + 36, cardW - 40, 50, {
      size: 14, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s,
    "フルサポート50万円のみ。\n" +
    "A社の15分プラン（45万円）と同水準の\n" +
    "価格で60分配信＋データ連動を提供\n" +
    "しており、市場価格との乖離がある。",
    c1x + 20, cardY + 100, cardW - 40, 120, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.START
    });

  // 課題②
  var c2x = startX + cardW + gap;
  rect(s, c2x, cardY, cardW, cardH, C_WHITE, C_ACCENT, 2);
  rect(s, c2x, cardY, 5, cardH, C_ACCENT);
  txt(s, "現状課題 ②", c2x + 20, cardY + 14, cardW - 40, 14, {
    size: 10, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "貴社の他メニューと比較して\n粗利率が低い（69.3%）",
    c2x + 20, cardY + 36, cardW - 40, 50, {
      size: 14, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s,
    "ショート動画83.3%・Xフォロリポ81.6%と\n" +
    "比べて低い。制作外注費の原価\n" +
    "（台本作成、現場ディレクション、\n" +
    "クリエイティブ制作）が50万円に\n" +
    "内包されている構造による。",
    c2x + 20, cardY + 100, cardW - 40, 120, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.START
    });
}

// ============================================================
// S15: 柱① 提案①：単価是正（A社比較）
// ============================================================
function insertS15v2_priceRevision(pres) {
  var s = newSlide(pres, "柱① 提案①：フルサポートプランの単価是正（A社比較）");

  // 結論メッセージ
  drawConclusionMessage(s, 88,
    "A社の15分プラン（45万円）と同水準の価格で60分・データ連動を提供している",
    "同水準の価格で60分・データ連動");

  // 3列テーブル: プラン / 価格 / 内容
  var colDefs = [
    { label: "プラン",   align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "価格",     align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "内容",     align: SlidesApp.ParagraphAlignment.START }
  ];
  var rows = [
    ["A社 スタンダード",                    "45万円", "15分・最大3商品"],
    ["A社 ポップアップ中継",                "70万円", "生中継"],
    ["マツキヨココカラライブ\n（現行）",   "50万円", "60分・データ連動・5,000人"]
  ];

  // マツキヨ行（index=2）をクリーム背景でハイライト
  makeColHighlightTable(s, 60, 145, 600, colDefs, rows, {
    rowH: 44,
    hdrH: 32,
    bodySize: 12,
    headerSize: 11,
    highlightRow: 2
  });

  // 注釈
  txt(s, "※ A社価格はA社公式サイト掲載のメニュー価格を参照（2026年4月時点）。",
    25, 320, 671, 16, {
      size: 8, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 結論注釈
  txt(s, "→ 60分配信・データ連動・平均5,000人視聴という提供価値に対して、現行50万円は競合比較で割安",
    25, 348, 671, 22, {
      size: 11, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S16: 柱① 体制改善：現状課題（工数膨張・薬機法CK詰まり）
// ============================================================
function insertS16v2_opsIssues(pres) {
  var s = newSlide(pres, "柱① 協賛配信の運営体制改善 ― 現状課題");

  // 結論メッセージ
  drawConclusionMessage(s, 88,
    "月5回以上の安定受注は実現。次の課題は受注後の運営フローを最適化し作業工数を下げること",
    "受注後の運営フローを最適化");

  // 2カード
  var cardY = 130;
  var cardH = 230;
  var cardW = 325;
  var gap = 18;
  var startX = (SW - 2 * cardW - gap) / 2;

  // 課題①
  var c1x = startX;
  rect(s, c1x, cardY, cardW, cardH, C_WHITE, C_ACCENT, 2);
  rect(s, c1x, cardY, 5, cardH, C_ACCENT);
  txt(s, "現状課題 ①", c1x + 20, cardY + 14, cardW - 40, 14, {
    size: 10, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "受け入れ無制限による\n工数膨張",
    c1x + 20, cardY + 36, cardW - 40, 50, {
      size: 14, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s,
    "台本初稿後の商品変更・追加に上限が\n" +
    "なく、台本量・薬機法チェック工数が\n" +
    "読めない。競合A社が約2週間で準備\n" +
    "完了するのに対し、現状は1.5〜2ヶ月。\n" +
    "実例：初稿後の商品変更で台本が\n" +
    "6稿に達したケースも。",
    c1x + 20, cardY + 100, cardW - 40, 120, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.START
    });

  // 課題②
  var c2x = startX + cardW + gap;
  rect(s, c2x, cardY, cardW, cardH, C_WHITE, C_ACCENT, 2);
  rect(s, c2x, cardY, 5, cardH, C_ACCENT);
  txt(s, "現状課題 ②", c2x + 20, cardY + 14, cardW - 40, 14, {
    size: 10, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "薬機法チェックが\n詰まる構造的問題",
    c2x + 20, cardY + 36, cardW - 40, 50, {
      size: 14, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s,
    "修正稿ベースの後ろ倒し運用に加え、\n" +
    "貴社法務からのNGの代替表現が提示\n" +
    "されず現場が毎回詰まる。\n" +
    "台本作成者・貴社担当者ともに\n" +
    "「どこまでOKか」の共通基準がなく、\n" +
    "都度確認が発生している。",
    c2x + 20, cardY + 100, cardW - 40, 120, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.START
    });
}

// ============================================================
// S17: 柱① 体制改善 上半期①：レギュレーション設定（プレースホルダー）
// ============================================================
function insertS17v2_regulation(pres) {
  var s = newSlide(pres, "柱① 体制改善 上半期①：レギュレーション設定");

  // 結論メッセージ
  drawConclusionMessage(s, 88,
    "配信ごとの商品変更・追加・修正の無制限受け入れに、レギュレーションを設ける",
    "レギュレーションを設ける");

  // 上半期ラベル
  txt(s, "上半期（〜2026年9月）", 60, 125, 600, 18, {
    size: 11, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 2列テーブル: 項目 / 内容
  var colDefs = [
    { label: "項目",  align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "内容",  align: SlidesApp.ParagraphAlignment.START }
  ];
  var rows = [
    ["商品確定締め切り",       "配信X週間前(要協議)"],
    ["主要紹介アイテム数",     "◯アイテムまで(カラバリ・香りなどのバリエーションは除く、要協議)"],
    ["台本修正回数",           "◯回まで(◯回目以降は別途修正費用発生。誤植除く)"]
  ];

  makeColHighlightTable(s, 50, 155, 620, colDefs, rows, {
    rowH: 42,
    hdrH: 32,
    bodySize: 12,
    headerSize: 11
  });

  // 注釈
  txt(s, "※ ◯部分は貴社・Firework間で具体値を協議のうえ確定。",
    25, 320, 671, 16, {
      size: 8, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 結論注釈
  txt(s, "→ 台本量・薬機法チェック工数が読めるようになり、月5回以上の並行受注でも管理が安定する",
    25, 348, 671, 22, {
      size: 11, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S18: 柱① 体制改善：薬機法CK + AI活用（上半期② + 下半期③）
// ============================================================
function insertS18v2_yakkihouAI(pres) {
  var s = newSlide(pres, "柱① 体制改善：薬機法CK + AI活用");

  // 結論メッセージ
  drawConclusionMessage(s, 88,
    "薬事ルールの制定とAI活用で、薬機法チェック詰まりを構造的に解消する",
    "構造的に解消する");

  // 上半期② 薬機法CKフロー改善（2項目）
  txt(s, "上半期②（〜2026年9月）薬機法CKフロー改善",
    30, 125, 660, 18, {
      size: 11, color: C_ACCENT, bold: true,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 上半期② カード1: ライブ配信向け薬事ルールの制定
  var card1Y = 150;
  rect(s, 30, card1Y, 660, 60, "#FFF5F7", C_ACCENT, 1);
  rect(s, 30, card1Y, 4, 60, C_ACCENT);
  txt(s, "ライブ配信向け薬事ルールの制定",
    44, card1Y + 8, 640, 18, {
      size: 12, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s,
    "ライブ配信に特化した判断基準・ルールを設ける。" +
    "参考：A社は薬機法を熟知した美容部員教育担当が出演者を兼任しており、法務の薬事CKを実施していない。\n" +
    "今後、社員インフルエンサーを育成していく際には、薬事知識のインプットも併せて進められるとよい。",
    44, card1Y + 28, 640, 30, {
      size: 9.5, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.START
    });

  // 上半期② カード2: GoogleAIを活用した薬事CKフロー
  var card2Y = card1Y + 70;
  rect(s, 30, card2Y, 660, 50, "#FFF5F7", C_ACCENT, 1);
  rect(s, 30, card2Y, 4, 50, C_ACCENT);
  txt(s, "GoogleAIを活用した薬事CKのフローを構築",
    44, card2Y + 8, 640, 18, {
      size: 12, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "別途Fireworkから具体的なフローをご提案。",
    44, card2Y + 28, 640, 18, {
      size: 9.5, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 下半期③ AI活用による業務効率化
  var s3Y = card2Y + 65;
  txt(s, "下半期（2026年10月〜）③ AI活用による業務効率化",
    30, s3Y, 660, 18, {
      size: 11, color: C_ACCENT2, bold: true,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });

  var card3Y = s3Y + 25;
  rect(s, 30, card3Y, 660, 50, "#F0FAF8", C_ACCENT2, 1);
  rect(s, 30, card3Y, 4, 50, C_ACCENT2);
  txt(s,
    "貴社が活用中のGoogle AI(Gemini・NotebookLM・AI Studio)を、配信後の実績レポート作成などの定型業務に活用。\n" +
    "貴社セキュリティポリシーに則ったかたちでFireworkが具体的な活用法をご提案。",
    44, card3Y + 8, 640, 35, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S19: 柱② 社員インフルエンサー強化
// ============================================================
function insertS19v2_employeeInfluencer(pres) {
  var s = newSlide(pres, "柱② 社員インフルエンサー強化");

  // 結論メッセージ
  drawConclusionMessage(s, 88,
    "「推しの出演者が出るから観る」という再訪動機を生み出し、リピーター比率を底上げする",
    "再訪動機を生み出し、リピーター比率を底上げ");

  // サブ説明
  txt(s,
    "出演者(美容部員・店舗スタッフ)のライブ配信スキルを育成・サポート。" +
    "SNS発信は個人の自由な活動を基本とし、適性・志望がある人材を社員・パートを問わず発掘・育成する。",
    25, 118, 671, 32, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 4アクション（2x2 グリッド）
  var actY = 158;
  var actH = 95;
  var actW = 325;
  var actGap = 18;
  var startX = (SW - 2 * actW - actGap) / 2;
  var actions = [
    {
      x: startX,
      y: actY,
      no: "①",
      title: "SNSフォロワー上位の発掘・選定",
      desc: "SNSフォロワー上位のスタッフを優先的に発掘・選定し、ライブ配信出演者として教育・サポートする。"
    },
    {
      x: startX + actW + actGap,
      y: actY,
      no: "②",
      title: "既存出演者のSNS強化",
      desc: "既存出演者のSNS開設・強化(STさま施策との連動)。"
    },
    {
      x: startX,
      y: actY + actH + 12,
      no: "③",
      title: "自社メディア活用",
      desc: "社員インフルエンサーコンテンツの自社メディア活用(ECサイト、アプリなど)。"
    },
    {
      x: startX + actW + actGap,
      y: actY + actH + 12,
      no: "④",
      title: "薬事知識研修",
      desc: "出演者への薬事知識研修(A社はライブ配信台本において法務CK不要な体制を実現)。"
    }
  ];

  actions.forEach(function(a) {
    rect(s, a.x, a.y, actW, actH, C_WHITE, C_ACCENT2, 1);
    rect(s, a.x, a.y, 4, actH, C_ACCENT2);
    txt(s, a.no, a.x + 14, a.y + 8, 24, 18, {
      size: 14, color: C_ACCENT2, bold: true,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
    txt(s, a.title, a.x + 40, a.y + 8, actW - 50, 18, {
      size: 12, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.MIDDLE
    });
    txt(s, a.desc, a.x + 14, a.y + 32, actW - 28, 60, {
      size: 9.5, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START, va: SlidesApp.ContentAlignment.START
    });
  });

  // 根拠注釈（下部・緑bg）
  rect(s, 25, 365, 671, 24, "#F0FAF8", C_ACCENT2, 1);
  txt(s,
    "根拠：競合B社は個人SNSフォロワー数万〜16万超のBAが複数在籍しファン流入→高エンゲージを実現。" +
    "マツキヨは現状リピーター比率5.0%(業界最下位)で出演者ファンベース構築が最重点課題。",
    35, 367, 651, 20, {
      size: 8, color: C_ACCENT2, bold: false,
      align: SlidesApp.ParagraphAlignment.CENTER, va: SlidesApp.ContentAlignment.MIDDLE
    });
}
