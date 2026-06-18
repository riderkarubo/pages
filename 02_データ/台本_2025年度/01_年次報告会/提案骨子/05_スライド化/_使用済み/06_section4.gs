/**
 * MCCM提案スライド化 - §4: 2026年度の2本柱（7スライド）
 *
 * マスター情報源: https://riderkarubo.github.io/pages/mccm-proposal/
 * （これ以外の場所からは情報を拾わない）
 *
 * 構成:
 *  S13: 2026年度の2本柱（章扉）
 *  S14: 柱① 現状課題（価格・粗利率）
 *  S15: 柱① 提案①：単価是正（A社比較のみ・グレーアウト部分非掲載）
 *  S16: 柱① 体制改善：現状課題（工数膨張・薬機法CK詰まり）
 *  S17: 柱① 体制改善 上半期①：レギュレーション設定
 *  S18: 柱① 体制改善：薬機法CK + AI活用（統合）
 *  S19: 柱② 社員インフルエンサー強化
 *
 * 実行手順:
 *  1. 02_helpers.gs / 04_section2.gs / 05_section3.gs を貼り付け済みであること
 *  2. このコードを Apps Script に貼り付け
 *  3. insertSection4() を実行 → 末尾に7枚追加される
 */

function insertSection4() {
  var pres = SlidesApp.openById(PRESENTATION_ID);
  Logger.log("§4 開始 - 現在: " + pres.getSlides().length + "枚");

  insertS13_pillarsCover(pres);   Logger.log("S13 完了");
  insertS14_pillar1Issues(pres);  Logger.log("S14 完了");
  insertS15_priceRevision(pres);  Logger.log("S15 完了");
  insertS16_opsIssues(pres);      Logger.log("S16 完了");
  insertS17_regulation(pres);     Logger.log("S17 完了");
  insertS18_yakkihoAI(pres);      Logger.log("S18 完了");
  insertS19_pillar2(pres);        Logger.log("S19 完了");

  Logger.log("§4 完了 - 現在: " + pres.getSlides().length + "枚");
}

// ============================================================
// S13: 2026年度の2本柱（章扉）
// ============================================================
function insertS13_pillarsCover(pres) {
  var s = newSlide(pres, "2026年度の2本柱");

  // サブタイトル
  txt(s, "KPI（協賛配信受注 年間48件・2,400万円）達成に向けた2つの重点施策",
    25, 76, 671, 24, {
      size: 11, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 2本柱カード（横並び大型）
  var cardW = 320, cardH = 200, gap = 24;
  var startX = (SW - 2 * cardW - gap) / 2;
  var startY = 130;

  // 柱①
  rect(s, startX, startY, cardW, cardH, "#FFF5F7", C_ACCENT, 2);
  txt(s, "柱 ①", startX + 14, startY + 14, cardW - 28, 18, {
    size: 11, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "協賛配信のメニュー・体制改善", startX + 14, startY + 38, cardW - 28, 28, {
    size: 17, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "受注後の運営体制を標準化し、月4回以上の安定受注を実現する。",
    startX + 14, startY + 76, cardW - 28, 36, {
      size: 11, color: "#555",
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
  txt(s, "・提案①：フルサポートプラン単価是正（A社比較）\n・体制改善：レギュレーション設定\n・体制改善：薬機法CKフロー改善 + AI活用",
    startX + 14, startY + 116, cardW - 28, 70, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });

  // 柱②
  rect(s, startX + cardW + gap, startY, cardW, cardH, "#F0FAF8", C_ACCENT2, 2);
  txt(s, "柱 ②", startX + cardW + gap + 14, startY + 14, cardW - 28, 18, {
    size: 11, color: C_ACCENT2, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "社員インフルエンサー強化", startX + cardW + gap + 14, startY + 38, cardW - 28, 28, {
    size: 17, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "出演者を育成・サポートし、リピーター比率の底上げを実現する。",
    startX + cardW + gap + 14, startY + 76, cardW - 28, 36, {
      size: 11, color: "#555",
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
  txt(s, "・SNSフォロワー上位スタッフを優先発掘・教育\n・既存出演者のSNS開設・強化（&ST連動）\n・社員インフルエンサーコンテンツの自社メディア活用\n・出演者への薬事知識研修",
    startX + cardW + gap + 14, startY + 116, cardW - 28, 80, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
}

// ============================================================
// S14: 柱① 現状課題
// ============================================================
function insertS14_pillar1Issues(pres) {
  var s = newSlide(pres, "柱① 協賛配信のメニュー：現状課題");

  // 2カード（縦並び・厚め）
  var cardW = 671, cardH = 130, gap = 16;
  var startX = 25;
  var startY = 96;

  // 課題①
  rect(s, startX, startY, cardW, cardH, "#FFF5F7", C_ACCENT, 1);
  rect(s, startX, startY, 4, cardH, C_ACCENT);
  txt(s, "現状課題 ①", startX + 16, startY + 12, cardW - 32, 16, {
    size: 10, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "競合他社と比べて提供価値に対する価格が低い", startX + 16, startY + 32, cardW - 32, 24, {
    size: 14, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "フルサポート50万円のみ。A社の15分プラン（45万円）と同水準の価格で60分配信＋データ連動を提供しており、市場価格との乖離がある。",
    startX + 16, startY + 60, cardW - 32, 60, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });

  // 課題②
  var y2 = startY + cardH + gap;
  rect(s, startX, y2, cardW, cardH, "#FFF5F7", C_ACCENT, 1);
  rect(s, startX, y2, 4, cardH, C_ACCENT);
  txt(s, "現状課題 ②", startX + 16, y2 + 12, cardW - 32, 16, {
    size: 10, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "貴社の他メニューと比較して粗利率が低い（69.3%）", startX + 16, y2 + 32, cardW - 32, 24, {
    size: 14, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "ショート動画83.3%・Xフォロリポ81.6%と比べて低い。制作外注費の原価（台本作成、現場ディレクション、クリエイティブ制作）が50万円に内包されている構造による。",
    startX + 16, y2 + 60, cardW - 32, 60, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
}

// ============================================================
// S15: 柱① 提案①：単価是正（A社比較）
// ============================================================
function insertS15_priceRevision(pres) {
  var s = newSlide(pres, "提案①：フルサポートプランの単価是正（A社比較）");

  // 説明文
  txt(s, "A社の競合プランと比較し、マツキヨココカラライブの提供価値（60分・データ連動・5,000人）と価格の乖離を整理。",
    25, 96, 671, 22, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // テーブル
  var colDefs = [
    { label: "プラン", align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "価格", align: SlidesApp.ParagraphAlignment.CENTER },
    { label: "内容", align: SlidesApp.ParagraphAlignment.START }
  ];
  var rows = [
    ["A社 スタンダード",                "45万円", "15分・最大3商品"],
    ["A社 ポップアップ中継",            "70万円", "生中継"],
    ["マツキヨココカラライブ（現行）",   "50万円", "60分・データ連動・5,000人"]
  ];

  makeTable(s, 25, 130, 671, colDefs, rows, {
    rowH: 36,
    hdrH: 30,
    headerSize: 11,
    bodySize: 11,
    highlightRows: [2]
  });

  // 結論ボックス
  rect(s, 25, 280, 671, 90, "#FFF5F5", C_ACCENT, 1);
  txt(s, "結論", 35, 286, 651, 18, {
    size: 10, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "同水準の価格で60分配信＋データ連動を提供しており、市場価格との乖離がある。",
    35, 308, 651, 22, {
      size: 13, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "次期分科会にて、本提案に対する議論・粗利率インパクトを別途ご提示します。",
    35, 334, 651, 22, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S16: 柱① 体制改善 現状課題
// ============================================================
function insertS16_opsIssues(pres) {
  var s = newSlide(pres, "協賛配信の運営体制改善 ― 現状課題");

  // リード文
  txt(s, "月5回以上の安定受注はできている。次の課題は、受注後の運営フローを2社間で最適化し、貴社の作業工数負荷を下げること。",
    25, 96, 671, 30, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // 2カード（縦並び）
  var cardW = 671, cardH = 110, gap = 14;
  var startX = 25;
  var startY = 132;

  rect(s, startX, startY, cardW, cardH, "#FFF5F7", C_ACCENT, 1);
  rect(s, startX, startY, 4, cardH, C_ACCENT);
  txt(s, "現状課題 ①  受け入れ無制限による工数膨張", startX + 16, startY + 10, cardW - 32, 22, {
    size: 13, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "台本初稿後の商品変更・追加に上限がなく、台本量・薬機法チェック工数が読めない。競合（A社）が約2週間で準備を完了するのに対し、現状の準備期間は1.5〜2ヶ月。実例として初稿後の商品変更により台本が6稿に達したケースも。",
    startX + 16, startY + 36, cardW - 32, 70, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });

  var y2 = startY + cardH + gap;
  rect(s, startX, y2, cardW, cardH, "#FFF5F7", C_ACCENT, 1);
  rect(s, startX, y2, 4, cardH, C_ACCENT);
  txt(s, "現状課題 ②  薬機法チェックが詰まる構造的問題", startX + 16, y2 + 10, cardW - 32, 22, {
    size: 13, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "修正稿ベースの後ろ倒し運用に加え、貴社法務からのNGの代替表現が提示されず現場が毎回詰まる。台本作成者・貴社担当者ともに「どこまでOKか」の共通基準がなく、都度確認が発生している。",
    startX + 16, y2 + 36, cardW - 32, 70, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
}

// ============================================================
// S17: 柱① 体制改善 上半期①：レギュレーション設定
// ============================================================
function insertS17_regulation(pres) {
  var s = newSlide(pres, "上半期 ① レギュレーション設定");

  // バッジ＋リード文
  rect(s, 25, 88, 100, 22, C_ACCENT);
  txt(s, "上半期（〜2026年9月）", 25, 88, 100, 22, {
    size: 9, color: C_WHITE, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  txt(s, "配信ごとの商品変更・追加・修正の無制限受け入れが、台本・チェック工数を押し上げている。\n受注前にメーカー側と共通ルールを定義することで、工数負荷を構造的に下げる。",
    25, 118, 671, 40, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });

  // テーブル
  var colDefs = [
    { label: "項目", align: SlidesApp.ParagraphAlignment.START, nameCol: true },
    { label: "内容", align: SlidesApp.ParagraphAlignment.START }
  ];
  var rows = [
    ["商品確定締め切り", "配信X週間前（要協議）"],
    ["主要紹介アイテム数", "◯アイテムまで（カラバリ・香りなどのバリエーションは除く、要協議）"],
    ["台本修正回数", "◯回まで（◯回目以降は別途修正費用発生。誤植除く）"]
  ];

  makeTable(s, 25, 170, 671, colDefs, rows, {
    rowH: 50,
    hdrH: 30,
    headerSize: 11,
    bodySize: 11
  });

  // フッター
  txt(s, "※ 具体的な数値（締切週数・アイテム数・修正回数）は別途協議",
    25, 360, 671, 18, {
      size: 9, color: C_GRAY,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
}

// ============================================================
// S18: 柱① 体制改善：薬機法CK + AI活用（統合）
// ============================================================
function insertS18_yakkihoAI(pres) {
  var s = newSlide(pres, "体制改善 ② 薬機法CKフロー改善 ＋ ③ AI活用");

  // 上半期バッジ
  rect(s, 25, 88, 100, 20, C_ACCENT);
  txt(s, "上半期 ②", 25, 88, 100, 20, {
    size: 9, color: C_WHITE, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "薬機法CKフロー改善", 130, 88, 565, 20, {
    size: 12, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 上段2カード（横並び）
  var cardW = 330, cardH = 110, gap = 11;
  var startX = 25;
  var topY = 116;

  rect(s, startX, topY, cardW, cardH, "#FFF5F7");
  rect(s, startX, topY, 3, cardH, C_ACCENT);
  txt(s, "ライブ配信向け薬事ルールの制定", startX + 14, topY + 10, cardW - 28, 20, {
    size: 12, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "ライブ配信に特化した判断基準・ルールを設ける。参考：A社は薬機法を熟知した美容部員教育担当が出演者を兼任しており、法務の薬事CKを実施していない。今後、社員インフルエンサーを育成していく際には、薬事知識のインプットも併せて進められるとよい。",
    startX + 14, topY + 32, cardW - 28, 74, {
      size: 9, color: "#555",
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });

  rect(s, startX + cardW + gap, topY, cardW, cardH, "#FFF5F7");
  rect(s, startX + cardW + gap, topY, 3, cardH, C_ACCENT);
  txt(s, "GoogleAI を活用した薬事CKフロー構築", startX + cardW + gap + 14, topY + 10, cardW - 28, 20, {
    size: 12, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "別途、Fireworkから具体的なフローをご提案。法務確認の前段でAIによる一次チェックを導入し、現場が詰まる構造を解消する。",
    startX + cardW + gap + 14, topY + 32, cardW - 28, 74, {
      size: 9, color: "#555",
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });

  // 下半期バッジ
  var bottomY = 246;
  rect(s, 25, bottomY, 100, 20, C_ACCENT2);
  txt(s, "下半期 ③", 25, bottomY, 100, 20, {
    size: 9, color: C_WHITE, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "AI活用による業務効率化", 130, bottomY, 565, 20, {
    size: 12, color: C_ACCENT2, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 下段ブロック
  rect(s, 25, bottomY + 28, 671, 110, "#F0FAF8");
  rect(s, 25, bottomY + 28, 3, 110, C_ACCENT2);
  txt(s, "Google AI（Gemini・NotebookLM・AI Studio）を実績レポートに活用",
    39, bottomY + 36, 657, 22, {
      size: 12, color: C_DARK, bold: true,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "貴社が活用中のGoogle AI（Gemini・NotebookLM・AI Studio）を、配信後の実績レポート作成などの定型業務に活用。\n貴社セキュリティポリシーに則ったかたちでFireworkが具体的な活用法をご提案。",
    39, bottomY + 62, 657, 64, {
      size: 11, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
}

// ============================================================
// S19: 柱② 社員インフルエンサー強化
// ============================================================
function insertS19_pillar2(pres) {
  var s = newSlide(pres, "柱② 社員インフルエンサー強化");

  // メイン訴求ボックス
  rect(s, 25, 88, 671, 64, "#F0FAF8", C_ACCENT2, 2);
  txt(s, "「推しの出演者が出るから観る」という再訪動機を作り、リピーター比率を底上げ",
    35, 96, 651, 24, {
      size: 13, color: C_ACCENT2, bold: true,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.MIDDLE
    });
  txt(s, "出演者（美容部員・店舗スタッフ）のライブ配信スキルを育成・サポートする。SNSでの発信は個人の自由な活動を基本とし、ライブ配信の適性・志望がある人材を社員・パートを問わず発掘・育成するスタンスで推進。",
    35, 122, 651, 28, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });

  // 具体的アクション（4枚カード・横2×縦2）
  txt(s, "具体的アクション", 25, 162, 671, 18, {
    size: 11, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  var cardW = 330, cardH = 60, gap = 11, vgap = 10;
  var startX = 25, startY = 184;

  drawActionCard(s, startX, startY, cardW, cardH, {
    no: "①",
    title: "SNSフォロワー上位スタッフを優先発掘・教育",
    detail: "ライブ配信出演者として教育・サポートする"
  });
  drawActionCard(s, startX + cardW + gap, startY, cardW, cardH, {
    no: "②",
    title: "既存出演者のSNS開設・強化（&ST連動）",
    detail: "STさま施策との連動"
  });
  drawActionCard(s, startX, startY + cardH + vgap, cardW, cardH, {
    no: "③",
    title: "社員インフルエンサーコンテンツの自社メディア活用",
    detail: "ECサイト、アプリなどでの再活用"
  });
  drawActionCard(s, startX + cardW + gap, startY + cardH + vgap, cardW, cardH, {
    no: "④",
    title: "出演者への薬事知識研修（A社事例）",
    detail: "A社はライブ配信台本において法務CK不要な体制を実現"
  });

  // 根拠フッター
  rect(s, 25, 318, 671, 56, "#FFF5F5", C_ACCENT, 1);
  txt(s, "根拠", 35, 322, 651, 16, {
    size: 9, color: C_ACCENT, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });
  txt(s, "競合B社は個人SNSフォロワー数万〜16万超のBAが複数在籍しファン流入→高エンゲージを実現。\nマツキヨココカラライブはリピーター比率5.0%（業界最下位）であり、出演者ファンベース構築が最重点課題。",
    35, 340, 651, 32, {
      size: 10, color: C_TEXT,
      align: SlidesApp.ParagraphAlignment.START,
      va: SlidesApp.ContentAlignment.START
    });
}

/** アクションカード（番号＋タイトル＋詳細） */
function drawActionCard(slide, x, y, w, h, p) {
  rect(slide, x, y, w, h, C_WHITE, C_BORDER, 1);
  rect(slide, x, y, 4, h, C_ACCENT2);

  // 番号（左寄せ大きめ）
  txt(slide, p.no, x + 10, y + 6, 22, h - 12, {
    size: 18, color: C_ACCENT2, bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // タイトル
  txt(slide, p.title, x + 36, y + 8, w - 46, 20, {
    size: 10.5, color: C_DARK, bold: true,
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.MIDDLE
  });

  // 詳細
  txt(slide, p.detail, x + 36, y + 30, w - 46, h - 36, {
    size: 9, color: "#555",
    align: SlidesApp.ParagraphAlignment.START,
    va: SlidesApp.ContentAlignment.START
  });
}
