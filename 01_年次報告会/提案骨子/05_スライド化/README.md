# MCCM提案 スライド化作業

提案HTML（§1〜§4 = アジェンダ1〜4）を、Fireworkテーマを踏襲したGoogle Slidesに変換する作業フォルダ。

## 出力先

- Google Slides URL: https://docs.google.com/presentation/d/14fq9YY5dAOksIIT-H_d1qsiCwyJEnhAY3dGp-KZSO6U/edit
- PRESENTATION_ID: `14fq9YY5dAOksIIT-H_d1qsiCwyJEnhAY3dGp-KZSO6U`

## 元コンテンツ

- 提案HTML（正本）: `../01_提案書/提案に向けた情報まとめ_統合版_20260414.html`
- 対象範囲: §1〜§4（`#sec-pillar3` まで）

## 進め方（フェーズ）

| フェーズ | 内容 | 状態 |
|---|---|---|
| Phase 0 | テンプレ調査GAS実行・ログ取得 | **進行中** |
| Phase 1 | スライド構成案作成（1スライド1メッセージで再構成） | 未着手 |
| Phase 2 | 章ごとにGAS実装・実行・スクショ確認・微調整 | 未着手 |
| Phase 3 | 全体通し確認・仕上げ | 未着手 |

## ファイル一覧

| ファイル | 役割 |
|---|---|
| `01_getSlideLayouts.gs` | Phase 0: テンプレ調査用GAS |
| `00_スライド構成案.md` | Phase 1: スライド構成案（後で作成） |
| `02_helpers.gs` | Phase 2: 共通ヘルパー（dup/setTitle/txt/rect等）（後で作成） |
| `03_section1.gs` 〜 `06_section4.gs` | Phase 2: 章ごとの実装（後で作成） |

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
