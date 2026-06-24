/**
 * PS_MCCM_新店舗OPEN_見積もりGAS_260624
 * 最終更新: 2026年6月24日 16:37（inspectSheet_v1 関数追加・手作業状態の恒久ダンプツール）
 * 制作パートナー: ディレクション=森田 / 技術=寺田さん / 回線=パンダスタジオ
 * マージン率: 基本45%（旅費・機材レンタルは実費＋15%）
 *
 * 【案件概要】
 *   - クライアント: MCCM（株式会社MCCマネジメント）
 *   - サービス: マツキヨココカラライブ 協賛配信
 *   - 案件: 新店舗OPEN記念 現地ライブ配信（メーカー協賛前提）
 *   - 配信回数: 1回のみ
 *   - 出張人数: ミニマム=2名 / リッチ=5名
 *   - ラフ提出期限: 2026年6月末
 *
 * 【機密情報の取り扱い】
 *   - 店舗の所在都市名・地名は記載しない
 *   - 交通費は「現地までの往復交通費」「現地宿泊費」と抽象化
 *   - 「フラッグシップ店OPEN」「POPUPスペース」「B1/2F」までは可
 *
 * 【プラン構成】
 *   プラン1: ミニマム（スマホ配信・2名体制）
 *   プラン2: リッチ（OBS配信・5名体制）
 *   共通: 現地遠征費（交通・宿泊・諸経費）
 *   オプション: 機材レンタル・追加カメラ・テロップ・切り抜き動画
 *
 * 【使い方】
 *   1. MCCM側で見積もり用スプレッドシートを開く
 *   2. 拡張機能 → Apps Script でこのコードを貼り付け
 *   3. addMccmFlagshipSheet() を実行 → 「🌟 MCCM新店舗OPEN_260624」シートが追加される
 *
 * 【列構成（A〜R 計18列・260624更新）】
 *   A=項目 / B=内容（旧A列付箋を昇格・常時表示） / C=Partner商流？ / D=Partner(代理店) / E=制作パートナー
 *   F=グロス=G*H / G=数量 / H=単価 / I=割引率 / J=割引ご価格=F*(1-I)
 *   K=Partnerマージン% / L=Partnerマージン=J*K / M=FWマージン=J-L-R / N=FWマージン%=M/J
 *   O=原価単価(IN・手入力) / P=原価単価(EX)=ROUND(O/1.1) / Q=原価合計(EX)=G*P / R=原価合計(IN)=G*O
 *
 * 【IN/EX 表記について（社内ルール準拠）】
 *   IN = 原価ブロック内側ベース（手入力する原価単価）
 *   EX = 原価ブロック外側ベース（IN/1.1で自動計算）
 *   ※ クライアント提示の売値（E列・G列）にはこの区分は適用しない
 */

// ─────────────────────────────────────────────
// 設定・定数
// ─────────────────────────────────────────────
var SPREADSHEET_ID = '1PsD_25_tEsoAewUTPDApQnnB0LNNstblQbuXBGvOfmg';  // MCCM見積もりスプシ（260624指定）

var CONFIG = {
  TAX_RATE:    1.1,
  FREEZE_ROWS: 5,
  NUM_COLS:    18,
};

var C = {
  DARK:       '#1A1A2E',
  NAVY:       '#0F3460',
  ACCENT:     '#FA006D',
  TEAL:       '#1B998B',
  WHITE:      '#FFFFFF',
  LIGHT_BG:   '#F4F5F7',
  HEADER_BG:  '#37474F',
  TOTAL_BG:   '#1A1A2E',
  FORMULA:    '#E8F0FE',
  INPUT:      '#FFFDE7',
  PINK_HL:    '#FCE4EC',
  TEAL_HL:    '#E0F2F1',
  MARGIN_HL:  '#F3E5F5',
  GRAY_TXT:   '#666666',
  STRIKEOUT:  '#AAAAAA',
  BORDER:     '#CCCCCC',
  PLAN1_BG:   '#BBDEFB',
  PLAN1_DATA: '#E3F2FD',
  PLAN2_BG:   '#C8E6C9',
  PLAN2_DATA: '#E8F5E9',
  TRAVEL_BG:  '#FFCDD2',
  TRAVEL_DATA:'#FFEBEE',
  OPT_BG:     '#FFE0B2',
  OPT_DATA:   '#FFF8E1',
};

var SHEET_NAME = '🌟 MCCM新店舗OPEN_260624';

// ─────────────────────────────────────────────
// メイン: MCCM新店舗OPEN シート追加
// ─────────────────────────────────────────────
function addMccmFlagshipSheet() {
  var ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  var existing = ss.getSheetByName(SHEET_NAME);
  var sh;
  if (existing) {
    sh = existing;
    var filter = sh.getFilter();
    if (filter) filter.remove();
    if (sh.getMaxRows() > 0 && sh.getMaxColumns() > 0) {
      sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
    }
    sh.setFrozenRows(0);
    sh.setFrozenColumns(0);
    sh.clear();
    sh.clearFormats();
    sh.clearNotes();
  } else {
    sh = ss.insertSheet(SHEET_NAME, ss.getNumSheets());
  }

  _setupColumns(sh);
  _buildTitle(sh,
    '🌟 MCCM 新店舗OPEN記念 現地ライブ配信 ご提案 見積もり',
    'メーカー協賛配信 | フラッグシップ店OPEN 1配信 | Partner=SoftBank経由(マージン10%) | 配信管理は貴社側実施 | ディレクション=森田/技術=寺田さん/回線=パンダスタジオ | FW粗利45%キープ(旅費は15%)');
  _buildLegendAt(sh, 3);
  _buildHeadersAt(sh, 5);

  var r = 6;

  // ════════════════════════════════════════════
  // プラン1: ミニマム（スマホ配信・2名体制）
  // ════════════════════════════════════════════
  r = _writeSectionHeader(sh, r, '▶ プラン1: ミニマムパターン（スマホ配信・2名体制）', C.PLAN1_BG, '#0D47A1');
  var plan1Start = r;

  r = _writeDataRow(sh, r,
    '制作進行（メインディレクター）',
    true, 'SoftBank', '森田', 1, 250000, 0.10, 112200,
    '※ 企画立案・進行管理・現地統括');
  r = _writeDataRow(sh, r,
    '台本作成',
    true, 'SoftBank', '森田', 1, 123000, 0.10, 55000,
    '※ お打ち合わせ1回含む / メーカー協賛社情報の組み込み');
  r = _writeDataRow(sh, r,
    '現地ディレクション（配信オペレーター兼務）',
    true, 'SoftBank', '寺田さん', 1, 98000, 0.10, 44000,
    '※ 2名目。スマホ撮影・配信操作を兼務 / 配信管理は貴社側実施前提');
  r = _writeDataRow(sh, r,
    '機材コーディネート・持ち込み費',
    true, 'SoftBank', '寺田さん', 1, 123000, 0.10, 55000,
    '※ リングライト・NEEWER照明2灯・ワイヤレスマイク4波・配信用iPhone1台');
  r = _writeDataRow(sh, r,
    'モバイル回線（ポケットWi-Fi・3日レンタル）',
    true, 'SoftBank', 'パンダスタジオ', 1, 156000, 0.10, 70000,
    '※ 現地ネット環境不可前提 / ポケットWi-Fi・TVUルーター');
  r = _writeDataRow(sh, r,
    '配信後レポート',
    true, 'SoftBank', '森田', 1, 98000, 0.10, 44000,
    '※ 視聴データ・コメント分析・改善提案 / 配信管理は貴社側実施前提');

  var plan1End = r - 1;
  r = _writeTotalRow(sh, r, '■ プラン1 合計（遠征費除く）', plan1Start, plan1End, C.PLAN1_DATA, '#0D47A1');
  var plan1TotalRow = r - 1;
  r = _writeMarginRow(sh, r, plan1TotalRow);
  r++;

  // ════════════════════════════════════════════
  // プラン2: リッチ（OBS配信・5名体制）
  // ════════════════════════════════════════════
  r = _writeSectionHeader(sh, r, '▶ プラン2: リッチパターン（OBS配信・1カメ・5名体制）', C.PLAN2_BG, '#1B5E20');
  var plan2Start = r;

  r = _writeDataRow(sh, r,
    'ディレクター',
    true, 'SoftBank', '森田', 1, 250000, 0.10, 112200,
    '※ 企画統括・現地全体進行・MCCM/出演者/協賛社との調整');
  r = _writeDataRow(sh, r,
    'アシスタントディレクター',
    true, 'SoftBank', '森田', 1, 98000, 0.10, 44000,
    '※ Dの補佐・フロアディレクション・出演者対応');
  r = _writeDataRow(sh, r,
    '台本作成',
    true, 'SoftBank', '森田', 1, 123000, 0.10, 55000,
    '※ お打ち合わせ1回含む / メーカー協賛社情報の組み込み');
  r = _writeDataRow(sh, r,
    'カメラマン（1カメ体制）',
    true, 'SoftBank', '寺田さん', 1, 98000, 0.10, 44000,
    '※ 1カメ体制 / 撮影オペレーション');
  r = _writeDataRow(sh, r,
    'OBSオペレーター',
    true, 'SoftBank', '寺田さん', 1, 196000, 0.10, 88000,
    '※ OBS操作・テロップ出し・配信品質管理');
  r = _writeDataRow(sh, r,
    '技術アシスタント',
    true, 'SoftBank', '寺田さん', 1, 98000, 0.10, 44000,
    '※ 機材セッティング補助・回線監視・現場テクニカルサポート');
  r = _writeDataRow(sh, r,
    '機材コーディネート・持ち込み費',
    true, 'SoftBank', '寺田さん', 1, 250000, 0.10, 112200,
    '※ 一眼カメラ1台・NEEWER照明・ワイヤレスマイク・OBS機材一式・スイッチャー');
  r = _writeDataRow(sh, r,
    'TVUルーター（3日レンタル）',
    true, 'SoftBank', 'パンダスタジオ', 1, '', 0.10, '',
    '※ 現地有線LAN確保不可前提 / パンダスタジオ3日レンタル / 単価Issy確認待ち');
  r = _writeDataRow(sh, r,
    '機材運搬費（現地搬送・前日入り）',
    true, 'SoftBank', '寺田さん', 1, 98000, 0.10, 44000,
    '※ 機材一式の現地搬送 / 配信前日搬入・終了後撤収');
  r = _writeDataRow(sh, r,
    'テロップ制作（協賛社・商品情報）',
    true, 'SoftBank', '寺田さん', 1, 123000, 0.10, 55000,
    '※ メーカー協賛社のロゴ・商品スペック・限定情報をテロップ化');
  r = _writeDataRow(sh, r,
    '配信後レポート',
    true, 'SoftBank', '森田', 1, 98000, 0.10, 44000,
    '※ 視聴データ・コメント分析・協賛社向け実績レポート / 配信管理は貴社側実施前提');

  var plan2End = r - 1;
  r = _writeTotalRow(sh, r, '■ プラン2 合計（遠征費除く）', plan2Start, plan2End, C.PLAN2_DATA, '#1B5E20');
  var plan2TotalRow = r - 1;
  r = _writeMarginRow(sh, r, plan2TotalRow);
  r++;

  // ════════════════════════════════════════════
  // 共通: 現地遠征費（プランごとに人数が変動）
  // ════════════════════════════════════════════
  r = _writeSectionHeader(sh, r, '▶ 現地遠征費（プラン別／実費＋諸経費）', C.TRAVEL_BG, '#B71C1C');
  var travelStart = r;

  r = _writeDataRow(sh, r,
    '【ミニマム】現地までの往復交通費（2名分）',
    true, 'SoftBank', '', 2, 38000, 0.10, 28000,
    '※ 1名あたり実費¥28,000 IN想定 / マージン15%程度');
  r = _writeDataRow(sh, r,
    '【ミニマム】現地宿泊費（2名×1泊）',
    true, 'SoftBank', '', 2, 20000, 0.10, 15000,
    '※ 1名1泊¥15,000 IN想定 / 配信前日入り想定');
  r = _writeDataRow(sh, r,
    '【ミニマム】現地諸経費（タクシー・食事等）',
    true, 'SoftBank', '', 1, 20000, 0.10, 15000,
    '※ 概算実費 / 領収書ベース精算も可');
  r = _writeDataRow(sh, r,
    '【リッチ】現地までの往復交通費（5名分）',
    true, 'SoftBank', '', 5, 38000, 0.10, 28000,
    '※ 1名あたり実費¥28,000 IN想定 / マージン15%程度');
  r = _writeDataRow(sh, r,
    '【リッチ】現地宿泊費（5名×1泊）',
    true, 'SoftBank', '', 5, 20000, 0.10, 15000,
    '※ 1名1泊¥15,000 IN想定 / 配信前日入り想定');
  r = _writeDataRow(sh, r,
    '【リッチ】現地諸経費（タクシー・食事等）',
    true, 'SoftBank', '', 1, 40000, 0.10, 30000,
    '※ 概算実費 / 領収書ベース精算も可');

  var travelEnd = r - 1;
  r = _writeTotalRow(sh, r, '■ 遠征費 合計（ミニマム+リッチ 両方含む参考値）', travelStart, travelEnd, C.TRAVEL_DATA, '#B71C1C');
  r++;

  // ════════════════════════════════════════════
  // オプション
  // ════════════════════════════════════════════
  r = _writeSectionHeader(sh, r, '▶ オプション（追加要望に応じて）', C.OPT_BG, '#E65100');
  var optStart = r;

  r = _writeDataRow(sh, r,
    'リハーサル日（前日テスト配信）',
    true, 'SoftBank', '寺田さん', 1, 156000, 0.10, 70000,
    '※ 配信前日に現地でリハーサル / 出演者の通し練習・機材確認');
  r = _writeDataRow(sh, r,
    '出演者キャスティング（インフルエンサー）',
    true, 'SoftBank', '', 1, '', 0.10, '',
    '※ 別途お見積もり / キャスティング費はキャスト次第');
  r = _writeDataRow(sh, r,
    '切り抜き動画制作（最大3本）',
    true, 'SoftBank', '寺田さん', 3, 49000, 0.10, 22000,
    '※ 配信アーカイブから縦型ショート / SNS・Firework埋込活用');
  r = _writeDataRow(sh, r,
    '練習動画制作（出演者向けトーク練習用）',
    true, 'SoftBank', '森田', 1, 123000, 0.10, 55000,
    '※ 配信前の出演者ウォームアップ動画');
  r = _writeDataRow(sh, r,
    '音声機材増設（ピンマイク3名以上）',
    true, 'SoftBank', '寺田さん', 1, 98000, 0.10, 44000,
    '※ 出演者3名以上の場合 / 複数ゲスト・対談形式');
  r = _writeDataRow(sh, r,
    '3カメ目追加（俯瞰カメラ・固定）',
    true, 'SoftBank', '寺田さん', 1, 123000, 0.10, 55000,
    '※ 全体俯瞰映像の追加 / POPUPスペース全体感を出す');

  var optEnd = r - 1;
  r = _writeTotalRow(sh, r, '■ オプション 合計', optStart, optEnd, C.OPT_DATA, '#E65100');
  r++;

  var lastDataRow = r - 1;

  r = _writeMemoSection(sh, r);

  _applyBorders(sh, 5, lastDataRow);
  _applyFreezeAndFilters(sh, 6, lastDataRow);

  SpreadsheetApp.flush();
  Logger.log('完了: ' + SHEET_NAME);
}

// ─────────────────────────────────────────────
// 列幅設定
// ─────────────────────────────────────────────
function _setupColumns(sh) {
  // A=項目(240) / B=内容(260) / C=商流? / D=Partner / E=制作パートナー /
  // F=グロス / G=数量 / H=単価 / I=割引率 / J=割引ご価格 /
  // K=Partnerマージン% / L=Partnerマージン / M=FWマージン / N=FWマージン% /
  // O=原価単価IN / P=原価単価EX / Q=原価合計EX / R=原価合計IN
  var widths = [240, 260, 75, 75, 130, 110, 50, 110, 60, 110, 94, 110, 110, 70, 110, 100, 110, 110];
  widths.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
}

// ─────────────────────────────────────────────
// タイトル行
// ─────────────────────────────────────────────
function _buildTitle(sh, title, subtitle) {
  var nc = CONFIG.NUM_COLS;
  sh.getRange(1, 1).setBackground(C.DARK);
  sh.getRange(1, 2, 1, nc - 1).merge()
    .setValue(title)
    .setBackground(C.DARK).setFontColor(C.ACCENT)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);

  sh.getRange(2, 1).setBackground(C.NAVY);
  sh.getRange(2, 2, 1, nc - 5).merge()
    .setValue(subtitle)
    .setBackground(C.NAVY).setFontColor(C.WHITE)
    .setFontSize(9).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(2, nc - 3, 1, 4).merge()
    .setValue('作成日: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月d日'))
    .setBackground(C.NAVY).setFontColor(C.WHITE)
    .setFontSize(9).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sh.setRowHeight(2, 26);
}

// ─────────────────────────────────────────────
// 凡例行
// ─────────────────────────────────────────────
function _buildLegendAt(sh, row) {
  var nc = CONFIG.NUM_COLS;
  sh.getRange(row, 1).setValue('凡例:')
    .setFontSize(9).setFontColor(C.GRAY_TXT).setFontWeight('bold');
  sh.getRange(row, 2).setValue('（内容列）')
    .setFontSize(9).setFontColor(C.GRAY_TXT).setHorizontalAlignment('center');
  sh.getRange(row, 3).setBackground(C.INPUT)
    .setValue('黄=手入力').setFontSize(9).setFontColor(C.GRAY_TXT).setHorizontalAlignment('center');
  sh.getRange(row, 4).setBackground(C.FORMULA)
    .setValue('青=自動計算').setFontSize(9).setFontColor(C.GRAY_TXT).setHorizontalAlignment('center');
  sh.getRange(row, 5, 1, 2).merge()
    .setValue('F列:グロス(ピンク)').setBackground(C.PINK_HL)
    .setFontSize(9).setFontColor(C.GRAY_TXT).setHorizontalAlignment('center');
  sh.getRange(row, 7, 1, 2).merge()
    .setValue('R列:原価IN(TEAL)').setBackground(C.TEAL_HL)
    .setFontSize(9).setFontColor(C.GRAY_TXT).setHorizontalAlignment('center');
  sh.getRange(row, 9, 1, nc - 8).merge()
    .setValue('※ 原価単価はINで手入力 → EX・合計は自動計算。売値は表記分離なし。基本マージン45% / 旅費・機材レンタル15%')
    .setFontSize(9).setFontColor(C.GRAY_TXT);
  sh.setRowHeight(row, 22);
  sh.setRowHeight(row + 1, 8);
}

// ─────────────────────────────────────────────
// ヘッダー行
// ─────────────────────────────────────────────
function _buildHeadersAt(sh, row) {
  var headers = [
    '項目', '内容', 'Partner\n商流？', 'Partner\n(代理店)', '制作\nパートナー',
    'グロス\n(提案出し値)', '数量', '単価', '割引率', '割引ご価格',
    'Partner\nマージン%', 'Partner\nマージン', 'FW\nマージン', 'FW\nマージン%',
    '原価単価\n(IN)', '原価単価\n(EX)', '原価合計\n(EX)', '原価合計\n(IN)',
  ];
  sh.getRange(row, 1, 1, CONFIG.NUM_COLS)
    .setValues([headers])
    .setBackground(C.HEADER_BG).setFontColor(C.WHITE)
    .setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(row, 44);
}

// ─────────────────────────────────────────────
// データ行
// ─────────────────────────────────────────────
function _writeDataRow(sh, row, item, partnerFlg, partner, subcontractor, qty, unitPrice, partnerMarginPct, costUnitIn, note) {
  var r = row;
  var FMT_YEN = '¥#,##0';

  // A=項目（手入力色）
  sh.getRange(r, 1).setValue(item).setBackground(C.INPUT).setFontSize(10).setFontWeight('bold');
  // B=内容（旧A列付箋を昇格・常時表示・薄背景）
  sh.getRange(r, 2).setValue(note || '').setBackground(C.LIGHT_BG)
    .setFontSize(9).setFontColor(C.GRAY_TXT).setWrap(true).setVerticalAlignment('top');
  // C=Partner商流？
  sh.getRange(r, 3).insertCheckboxes().setValue(partnerFlg === true).setBackground(C.INPUT);
  // D=Partner(代理店)
  sh.getRange(r, 4).setValue(partner || '').setBackground(C.INPUT).setFontSize(9).setHorizontalAlignment('center');
  // E=制作パートナー
  sh.getRange(r, 5).setValue(subcontractor || '').setBackground(C.INPUT).setFontSize(9);
  // F=グロス = G*H
  sh.getRange(r, 6)
    .setFormula('=IF(G' + r + '*H' + r + '=0,"",G' + r + '*H' + r + ')')
    .setBackground(C.PINK_HL).setNumberFormat(FMT_YEN);
  // G=数量（整数のみ・通貨形式不可）
  sh.getRange(r, 7).setValue(qty || '').setBackground(C.INPUT).setNumberFormat('0').setHorizontalAlignment('center');
  // H=単価
  sh.getRange(r, 8).setValue(unitPrice || '').setBackground(C.INPUT).setNumberFormat(FMT_YEN);
  // I=割引率
  sh.getRange(r, 9).setValue('').setBackground(C.INPUT).setNumberFormat('0%').setHorizontalAlignment('center');
  // J=割引ご価格 = F * (1 - I)
  sh.getRange(r, 10)
    .setFormula('=IF(F' + r + '="","",F' + r + '*(1-IF(I' + r + '="",0,I' + r + ')))')
    .setBackground(C.FORMULA).setNumberFormat(FMT_YEN);
  // K=Partnerマージン%
  sh.getRange(r, 11).setValue(partnerMarginPct || '').setBackground(C.INPUT)
    .setNumberFormat('0%').setHorizontalAlignment('center');
  // L=Partnerマージン = J * K
  sh.getRange(r, 12)
    .setFormula('=IF(J' + r + '="","",IF(K' + r + '="",0,J' + r + '*K' + r + '))')
    .setBackground(C.FORMULA).setNumberFormat(FMT_YEN);
  // M=FWマージン = J - L - R
  sh.getRange(r, 13)
    .setFormula('=IF(J' + r + '="","",J' + r + '-IF(L' + r + '="",0,L' + r + ')-IF(R' + r + '="",0,R' + r + '))')
    .setBackground(C.FORMULA).setNumberFormat(FMT_YEN);
  // N=FWマージン% = M / J
  sh.getRange(r, 14)
    .setFormula('=IF(OR(J' + r + '="",J' + r + '=0),"",M' + r + '/J' + r + ')')
    .setBackground(C.FORMULA).setNumberFormat('0.0%').setHorizontalAlignment('center');
  // O=原価単価IN（手入力）
  sh.getRange(r, 15).setValue(costUnitIn || '').setBackground(C.INPUT).setNumberFormat(FMT_YEN);
  // P=原価単価EX = ROUND(O/1.1)
  sh.getRange(r, 16)
    .setFormula('=IF(O' + r + '="","",ROUND(O' + r + '/1.1))')
    .setBackground(C.FORMULA).setNumberFormat(FMT_YEN);
  // Q=原価合計EX = G * P
  sh.getRange(r, 17)
    .setFormula('=IF(G' + r + '*P' + r + '=0,"",G' + r + '*P' + r + ')')
    .setBackground(C.FORMULA).setNumberFormat(FMT_YEN);
  // R=原価合計IN = G * O
  sh.getRange(r, 18)
    .setFormula('=IF(G' + r + '*O' + r + '=0,"",G' + r + '*O' + r + ')')
    .setBackground(C.TEAL_HL).setNumberFormat(FMT_YEN);

  sh.setRowHeight(r, 28);
  return r + 1;
}

// ─────────────────────────────────────────────
// セクションヘッダー行
// ─────────────────────────────────────────────
function _writeSectionHeader(sh, row, label, bgColor, textColor) {
  var bg  = bgColor  || C.NAVY;
  var txt = textColor || C.WHITE;
  sh.getRange(row, 1).setBackground(bg);
  sh.getRange(row, 2, 1, CONFIG.NUM_COLS - 1).merge()
    .setValue(label)
    .setBackground(bg).setFontColor(txt)
    .setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(row, 26);
  return row + 1;
}

// ─────────────────────────────────────────────
// 合計行
// ─────────────────────────────────────────────
function _writeTotalRow(sh, row, label, dataStart, dataEnd, bgColor, textColor) {
  var FMT_YEN = '¥#,##0';
  var bg  = bgColor  || C.TOTAL_BG;
  var txt = textColor || '#1A1A2E';

  sh.getRange(row, 1).setValue(label)
    .setBackground(bg).setFontColor(txt)
    .setFontSize(10).setFontWeight('bold');
  // ラベルの右隣（B=内容 / C=商流? / D=Partner / E=制作パートナー）を merge して同色塗装
  sh.getRange(row, 2, 1, 4).merge().setBackground(bg);

  // 合計対象列: F=グロス / J=割引ご価格 / L=Partnerマージン / M=FWマージン / Q=原価合計EX / R=原価合計IN
  var sumCols = { 6:'F', 10:'J', 12:'L', 13:'M', 17:'Q', 18:'R' };
  Object.keys(sumCols).forEach(function(col) {
    var letter = sumCols[col];
    var colNum = parseInt(col);
    var cellBg  = (colNum === 6)  ? C.PINK_HL
                : (colNum === 18) ? C.TEAL_HL
                : bg;
    var cellTxt = (colNum === 6 || colNum === 18) ? C.DARK : txt;
    sh.getRange(row, colNum)
      .setFormula('=IFERROR(SUMIF(A' + dataStart + ':A' + dataEnd + ',"<>"&"",' +
                  letter + dataStart + ':' + letter + dataEnd + '),"—")')
      .setBackground(cellBg).setFontColor(cellTxt)
      .setFontSize(10).setFontWeight('bold').setNumberFormat(FMT_YEN);
  });

  // N=FWマージン% = M / J（合計行の小計に対する％）
  sh.getRange(row, 14)
    .setFormula('=IF(OR(J' + row + '="",J' + row + '=0),"",M' + row + '/J' + row + ')')
    .setBackground(bg).setFontColor(txt)
    .setFontSize(10).setFontWeight('bold')
    .setNumberFormat('0.0%').setHorizontalAlignment('center');

  // 残り（sumCols外 かつ ラベルmerge外）の列を同色塗装: G H I K O P
  [7, 8, 9, 11, 15, 16].forEach(function(col) {
    sh.getRange(row, col).setBackground(bg).setFontColor(txt);
  });

  sh.setRowHeight(row, 28);
  return row + 1;
}

// ─────────────────────────────────────────────
// 粗利率行
// ─────────────────────────────────────────────
function _writeMarginRow(sh, row, totalRow) {
  var nc = CONFIG.NUM_COLS;
  var bg = C.MARGIN_HL;

  sh.getRange(row, 1).setBackground(bg);
  // 粗利率ラベルは B-E の4列にmerge
  sh.getRange(row, 2, 1, 4).merge()
    .setValue('粗利率（FWマージン / 割引後価格）')
    .setBackground(bg).setFontColor('#6A1B9A')
    .setFontSize(9).setFontWeight('bold').setVerticalAlignment('middle');
  // F列（新グロス位置）に率を表示。J=割引ご価格・M=FWマージン
  sh.getRange(row, 6)
    .setFormula('=IF(OR(J' + totalRow + '="",J' + totalRow + '=0),"",M' + totalRow + '/J' + totalRow + ')')
    .setBackground(bg).setFontColor('#6A1B9A')
    .setFontSize(10).setFontWeight('bold')
    .setNumberFormat('0.0%').setHorizontalAlignment('center');
  for (var col = 7; col <= nc; col++) {
    sh.getRange(row, col).setBackground(bg);
  }
  sh.setRowHeight(row, 20);
  return row + 1;
}

// ─────────────────────────────────────────────
// 案件メモ
// ─────────────────────────────────────────────
function _writeMemoSection(sh, row) {
  var nc = CONFIG.NUM_COLS;

  sh.getRange(row, 1).setBackground(C.LIGHT_BG);
  sh.getRange(row, 2, 1, nc - 1).merge()
    .setValue('💡 案件メモ・前提条件')
    .setBackground(C.LIGHT_BG).setFontSize(10).setFontWeight('bold').setFontColor(C.DARK);
  sh.setRowHeight(row, 24);
  row++;

  var memoNotes = [
    '※ 商流: SoftBank経由（Partnerマージン10%）。Partnerマージン控除後にFW粗利45%（旅費は15%）をキープする売値設計に再計算済み。',
    '※ 配信管理（コメント返信・回線監視）はミニマム・リッチともに貴社（MCCM）側で実施前提。Firework側は制作・配信オペレーションのみ担当。',
    '※ 売値計算式: 配信制作費 売値=原価IN÷0.45（K=10%・L=45%・原価=45%）／ 旅費 売値=原価IN÷0.75（K=10%・L=15%・原価=75%）。',
    '※ 旧マージン40%設計→今回SoftBank10%対応のためFW45%目標へ単価を引上げ（プラン1: ¥522k→¥692k、プラン2: ¥1,104k→¥1,432k）。',
    '※ 基本マージン率45%。旅費（交通費・宿泊費・諸経費）と機材レンタルはマージン15%程度の実費＋諸経費型。',
    '※ 制作パートナー振り分け: ディレクション（制作進行・台本・現場D・配信後レポート・練習動画）=森田 / 技術（カメラ・OBS・機材・運搬・テロップ・切り抜き・リハ・3カメ追加）=寺田さん / 回線（ポケットWi-Fi・TVUルーター）=パンダスタジオ。',
    '※ プラン1（ミニマム・スマホ配信）: 2名体制 / 機材はスマホ＋簡易照明＋ワイヤレスマイク。',
    '※ プラン2（リッチ・OBS配信）: 5名体制 / 一眼2カメ＋スイッチング＋テロップ＋有線LAN想定（不可ならTVUルーター）。',
    '※ 現地遠征費はミニマム・リッチ別行で計上。実際に採択されたプランの行のみ採用してください。',
    '※ ポケットWi-Fi・TVUルーター（パンダスタジオ3日レンタル）の単価はIssyから確認待ち。',
    '※ メーカー協賛社のエントリー獲得はMCCM側で対応。本見積もりには協賛収益は含まない（MCCM側で別途精算）。',
    '※ 配信日は12月下旬（OPEN週末）想定。前日入り（現地宿泊1泊）想定で旅費を計上。',
    '※ 売値（E列グロス・G列単価）はクライアント提示用 / N列・O列・P列・Q列は原価ブロック（社内計算用）。',
    '※ IN=原価ブロック内側ベース手入力 / EX=IN/1.1で自動計算（社内見積もり区分）。',
    '※ 出演者キャスティングは別途見積もり。キャスティング費はインフルエンサー本人費＋エージェント手数料で変動。',
    '※ 切り抜き動画3本はFirework埋込ECサイト・オウンドメディア活用が主用途（SNS転用は付帯）。',
  ];

  memoNotes.forEach(function(note) {
    sh.getRange(row, 1).setBackground(C.LIGHT_BG);
    sh.getRange(row, 2, 1, nc - 1).merge()
      .setValue(note)
      .setFontSize(8).setFontColor(C.GRAY_TXT).setBackground(C.LIGHT_BG);
    sh.setRowHeight(row, 18);
    row++;
  });

  return row;
}

// ─────────────────────────────────────────────
// ボーダー
// ─────────────────────────────────────────────
function _applyBorders(sh, startRow, endRow) {
  if (endRow < startRow) return;
  sh.getRange(startRow, 1, endRow - startRow + 1, CONFIG.NUM_COLS)
    .setBorder(true, true, true, true, true, true, C.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

// ─────────────────────────────────────────────
// フリーズ・フィルター
// ─────────────────────────────────────────────
function _applyFreezeAndFilters(sh, dataStartRow, lastRow) {
  sh.setFrozenRows(CONFIG.FREEZE_ROWS);
  sh.setFrozenColumns(1);
  if (lastRow > dataStartRow) {
    sh.getRange(5, 1, lastRow - 4, CONFIG.NUM_COLS).createFilter();
  }
}

// ═════════════════════════════════════════════════════════
// 部分パッチ関数 — 手作業修正済みのシートに対しピンポイント更新
// ═════════════════════════════════════════════════════════
/**
 * Issyさんが手作業でタイトル変更・金額変更・行削除等を入れたシートに対し、
 * 全消去せず指定セルだけを差分更新する関数。
 *
 * 使い方:
 *   Apps Script エディタで関数 patchSheet_260624_v1 を選択 → ▶ 実行
 *
 * 更新内容（2026-06-24 16:xx Issy指示反映）:
 *   - B10: 機材コーディネートの内容欄 — マイク「2波→4波」
 *   - B11: モバイル回線の内容欄 — 「ポケットWi-Fi・TVUルーター」へ書換
 *   - H11: 売値単価 ¥156,000（FW45%キープ・Partner10%適用：70000/0.45=155,556→切上）
 *   - O11: 原価単価IN ¥70,000（パンダスタジオ確定）
 *
 * 注意: シート上の他の手作業修正（タイトル・金額・削除行）は触らない。
 *       上記4セルのみ書き換え、他は全てそのまま保持。
 */
function patchSheet_260624_v1() {
  var ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    throw new Error('シートが見つかりません: ' + SHEET_NAME);
  }

  var FMT_YEN = '¥#,##0';
  // [セル, 値, フォーマット(null可), ラベル]
  var patches = [
    ['B10', '※ リングライト・NEEWER照明2灯・ワイヤレスマイク4波・配信用iPhone1台', null,
     'B10: 機材コーディネート 内容（マイク2波→4波）'],
    ['B11', '※ 現地ネット環境不可前提 / ポケットWi-Fi・TVUルーター', null,
     'B11: モバイル回線 内容（パンダスタジオ表記→ポケットWi-Fi・TVUルーター）'],
    ['O11', 70000, FMT_YEN,
     'O11: 原価単価IN ¥70,000（パンダスタジオ確定）'],
    ['H11', 156000, FMT_YEN,
     'H11: 売値単価 ¥156,000（FW45%キープ・Partner10%適用）'],
  ];

  var log = ['=== patchSheet_260624_v1 ==='];
  patches.forEach(function(p) {
    var cellA1 = p[0], val = p[1], fmt = p[2], label = p[3];
    var range = sh.getRange(cellA1);
    var before = range.getValue();
    range.setValue(val);
    if (fmt) range.setNumberFormat(fmt);
    log.push('  ' + label);
    log.push('    旧: ' + (before === '' ? '(空)' : before));
    log.push('    新: ' + val);
  });

  SpreadsheetApp.flush();
  Logger.log(log.join('\n'));
}

// ═════════════════════════════════════════════════════════
// Inspect関数 — シートの現状をログにダンプして Claude に渡せる形で出力
// ═════════════════════════════════════════════════════════
/**
 * シートの現在の手作業状態を全行ダンプする恒久ツール。
 * Issyさんが手作業（タイトル変更・金額変更・行削除）を入れた後、
 * これを実行して実行ログを Claude に共有すれば、Claude側で解析しマスターGAS／提案HTMLへ反映可能。
 *
 * 使い方:
 *   Apps Script エディタで関数 inspectSheet_v1 を選択 → ▶ 実行
 *   表示 → 実行ログ で出力をコピー → Claude に貼り付け
 *
 * 出力フォーマット:
 *   - シート名・最終行・最終列のヘッダ情報
 *   - 各行を「[行N] A列 | B列 | C列 | ... | R列」のパイプ区切りでダンプ
 *   - 通貨フォーマットの数値は「¥123,456」表記、％は「12.0%」、boolはチェックボックス記号
 *   - 空セルは「∅」（差分が一目でわかる）
 */
function inspectSheet_v1() {
  var ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    throw new Error('シートが見つかりません: ' + SHEET_NAME);
  }

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var formats = sh.getRange(1, 1, lastRow, lastCol).getNumberFormats();

  var log = [];
  var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  log.push('=== inspectSheet_v1 @ ' + stamp + ' ===');
  log.push('シート名: ' + sh.getName());
  log.push('行数: ' + lastRow + ' / 列数: ' + lastCol);
  log.push('列構成: A=項目 / B=内容 / C=商流? / D=Partner / E=制作パートナー / F=グロス / G=数量 / H=単価 / I=割引率 / J=割引ご価格 / K=Partnerマ% / L=Partnerマ / M=FWマ / N=FWマ% / O=原価IN / P=原価EX / Q=原価合計EX / R=原価合計IN');
  log.push('---');

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var fmts = formats[i];
    var rowDump = [];
    for (var j = 0; j < row.length; j++) {
      var v = row[j];
      var f = fmts[j] || '';
      var cell;
      if (v === '' || v === null || v === undefined) {
        cell = '∅';
      } else if (typeof v === 'number' && (f.indexOf('¥') >= 0 || f.indexOf('#,##0') >= 0) && f.indexOf('%') < 0) {
        cell = '¥' + Math.round(v).toLocaleString('en-US');
      } else if (typeof v === 'number' && f.indexOf('%') >= 0) {
        cell = (v * 100).toFixed(1) + '%';
      } else if (typeof v === 'boolean') {
        cell = v ? '☑' : '☐';
      } else if (typeof v === 'number') {
        cell = String(v);
      } else {
        cell = String(v).replace(/\n/g, ' ').substring(0, 100);
      }
      rowDump.push(cell);
    }
    log.push('[行' + (i + 1) + '] ' + rowDump.join(' | '));
  }

  log.push('---');
  log.push('END (' + lastRow + ' rows)');

  Logger.log(log.join('\n'));
}
