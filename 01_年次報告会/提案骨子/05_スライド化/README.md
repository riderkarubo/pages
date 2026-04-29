# MCCM提案 スライド化作業

提案HTML（§1〜§4 = アジェンダ1〜4）を、Fireworkテーマを踏襲したGoogle Slidesに変換する作業フォルダ。

## 出力先

- Google Slides URL: https://docs.google.com/presentation/d/14fq9YY5dAOksIIT-H_d1qsiCwyJEnhAY3dGp-KZSO6U/edit
- PRESENTATION_ID: `14fq9YY5dAOksIIT-H_d1qsiCwyJEnhAY3dGp-KZSO6U`

## マスター情報源

**`https://riderkarubo.github.io/pages/mccm-proposal/`**

これ以外の場所からは情報を拾わない（CLAUDE.md・ローカル他HTMLも参照禁止）。
ローカル正本ファイル: `mccm-proposal/index.html`（公開HTMLと同期されている）

## 進め方（フェーズ）

| フェーズ | 内容 | 状態 |
|---|---|---|
| Phase 0 | テンプレ調査GAS実行・ログ取得 | **完了** |
| Phase 1 | スライド構成案作成（1スライド1メッセージで再構成） | **完了** |
| Phase 2 | 章ごとにGAS実装・実行・スクショ確認・微調整 | **コード作成済・実行待ち** |
| Phase 3 | 全体通し確認・仕上げ | 未着手 |

## ファイル一覧

| ファイル | 役割 | 状態 |
|---|---|---|
| `00_スライド構成案.md` | Phase 1: 全20枚のスライド構成案 | ✅ |
| `02_helpers.gs` | 共通ヘルパー（dup/setTitle/txt/rect/makeTable/drawKpiCard等） | ✅ |
| `03_sections.gs` | §1〜§4 全20枚のGAS実装（`insertSection1〜4` の4関数を統合） | コード作成済 |
| `07_inspectP4toP7.gs` | Phase 0+++: 手直し済P4〜P7の完全構造取得（テーブルセル含む詳細解析） | ✅ 解析済 |
| `08_section3_v2.gs` | §3 (S7-S12 + S10b) v2版・実測テンプレ準拠 | コード作成済 |
| `10_s12b_self_data_trend.gs` | S12b 自社データ(コメント参加率・リピーター比率) 単独追加 | コード作成済 |
| `11_section4_v2.gs` | §4 (S13-S19) v2版・実測テンプレ準拠（2本柱・グレーアウト除外） | コード作成済 |
| `12_inspectAllSlides.gs` | 完成版全スライドの構造解析（テンプレライブラリ吸収用） | 実行待ち |
| `13_section5_repeater.gs` | S20（リピーターを増やす3つの戦略・1枚補足） | コード作成済（2026-04-29追加） |
| `KPI_tree_S8.pptx` | S8用 KPIツリー画像のPPTXソース（python-pptxで生成） | 確認待ち |
| `build_s8_kpi_tree.py` | KPI_tree_S8.pptx を再生成するPythonスクリプト | ✅ |
| `_使用済み/` | Phase 0 テンプレ調査GAS（`01_*.gs`）と統合前の章別GAS（`03_section1.gs` 〜 `06_section4.gs`）を保管 | アーカイブ |

## 関数構成（`03_sections.gs`）

| 関数 | 対象スライド | 状態 |
|---|---|---|
| `insertSection1()` | S1（エグゼクティブサマリ）| ✅ 実行済み |
| `insertSection2()` | S2〜S6（2025年度総括・5枚） | 実行待ち |
| `insertSection3()` | S7〜S12（2026年度KPI・6枚） | 実行待ち |
| `insertSection4()` | S13〜S19（2本柱・7枚） | 実行待ち |
| `insertSection5Repeater()` | S20（リピーター3戦略・1枚補足） | コード作成済 |

## スライド構成（全20枚）

| # | アジェンダ | タイトル |
|---|---|---|
| S1 | §1 | エグゼクティブサマリ |
| S2 | §2 | 2025年度総括（章扉） |
| S3 | §2 | 視聴者数 +149%（年度比） |
| S4 | §2 | 協賛実績の推移 |
| S5 | §2 | マイルストーン |
| S6 | §2 | 7つの経営資産 |
| S7 | §3 | 2026年度 主要KPI（協賛配信受注） |
| S8 | §3 | 売上構造の分解（KPIツリー） |
| S9 | §3 | マイルストーン①:メディアパワー(4指標) |
| S10 | §3 | 競合ベンチマーク（コスメカテゴリ） |
| S11 | §3 | メディアパワーの因数分解 |
| S12 | §3 | マイルストーン②:リピーター数値 |
| S13 | §4 | 2026年度の2本柱（章扉） |
| S14 | §4 | 柱① 現状課題 |
| S15 | §4 | 柱① 提案①:単価是正(A社比較) |
| S16 | §4 | 柱① 体制改善:現状課題 |
| S17 | §4 | 柱① 体制改善 上半期①:レギュレーション設定 |
| S18 | §4 | 柱① 体制改善:薬機法CK + AI活用 |
| S19 | §4 | 柱② 社員インフルエンサー強化 |
| S20 | 補足 | リピーターを増やす3つの戦略（1枚版・2026-04-29追加） |

## 使用スキル・参考資料

- `/fw-slide` コマンド（gasモード）
- ロレアル案件のGASヘルパー本体: `02_仕事/03_クライアント/ロレアル/04_ツール/replaceSlides_v2.js`
- 過去のノウハウ: `02_仕事/03_クライアント/ロレアル/01_提案・見積もり/引き継ぎ_GASスライド流し込み_260403.md`

## 安全ルール

- 既存Slidesの既存スライドは触らない（追加のみ）
- 失敗した追加スライドはユーザーが手動削除
- GASファイルは v2, v3 と増やさず同一ファイルを上書き更新

## 注意点

- Fireworkテーマは出力先Slidesに既にインポート済み
- 複製方式（`duplicate()` + `move()`）でレイアウト・装飾を踏襲
- `appendSlide(layout)` はマスター不整合エラーが出るため使用禁止
- ドロップシャドウは GAS API 制約で取得・設定不可 → 必要なら Slides UI から「書式設定 → 影」を一括適用
- タイトル黒帯座標は絶対不変: `TITLE_X=9, Y=20, W=671, H=45`（`02_helpers.gs` で定数化）
- S2〜S19はP4のような既存リファレンスがないため、テンプレ実測座標＋helpers関数で新規設計

## 実行手順（再開時）

1. Apps Scriptを開く（出力先Slidesの拡張機能 → Apps Script）
2. `02_helpers.gs` と `03_sections.gs` を新規ファイルとして追加
3. `insertSection2()` を実行 → ブラウザで確認 → 修正指示（5枚追加）
4. `insertSection3()` を実行 → ブラウザで確認 → 修正指示（6枚追加）
5. `insertSection4()` を実行 → ブラウザで確認 → 修正指示（7枚追加）
6. `insertSection5Repeater()` を実行 → ブラウザで確認 → 修正指示（1枚追加・S20）

### S20 単独実行
- `runS20Only()` で S20 のみ末尾に追加可能
- `13_section5_repeater.gs` を Apps Script に追加し、`insertSection5Repeater()` または `runS20Only()` を実行
