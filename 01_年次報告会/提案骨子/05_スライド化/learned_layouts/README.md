# MCCM最終資料 デザイン吸い出し作業

最終資料 (`https://docs.google.com/presentation/d/1eKkxLPQLumtMZOQqX7AwJPGWZcJ_wUXAbcfTfmVfAcQ/edit`) の **P2-15 + P41-50 (計24枚)** を構造化JSONに吸い出し、テンプレート化するための作業フォルダ。

最終ゴール: 他案件にも転用できる「Firework標準スライドテンプレート」として再生成ヘルパー関数 + デザインスペックを整備する。

---

## 作業ステップ

### Step 1: 吸い出し（Issyさん実行）

1. **新規GASプロジェクトを開く**
   - 対象資料を Google Slides で開いた状態で、`拡張機能 > Apps Script` を選択
   - 既存のスクリプトがあればそのまま追記でOK（関数名衝突なし）

2. **`15_inspectFinalDeck.gs` の中身を全部コピーして貼り付け**
   - このフォルダの `15_inspectFinalDeck.gs` を開いてコピー

3. **`inspectFinalDeck` を実行**
   - 関数選択ドロップダウンから `inspectFinalDeck` を選んで「実行」
   - 初回は権限承認ダイアログが出るので承認（SlidesApp + DriveApp）
   - 実行時間: 約30秒〜1分

4. **出力JSONを保存（2つの方法）**
   - **方法A（推奨・自動）**: スクリプトが自動で Drive の `MCCM_inspect_output/` フォルダに `p02-15_p41-50.json` を作成する。実行ログの末尾に出る Drive URL から内容を取得できる
   - **方法B（手動）**: 実行ログ（`表示 > ログ`）の `===JSON_START===` 〜 `===JSON_END===` の間をコピーして `learned_layouts/p02-15_p41-50.json` に保存

5. **Issyさんから完了報告**
   - 「吸い出し完了」と教えてもらえれば、こちら（Claude）でJSONを取得して Step 2 に進む

#### ログ上限が出たら（範囲分割実行）

GASのLoggerは最大文字数があるため、24枚分のJSONが乗り切らない可能性あり。その場合：

```javascript
// Apps Script エディタの実行ボックスで個別実行
inspectRange(2, 15);   // → MCCM_inspect_output/p2-15_inspect.json
inspectRange(41, 50);  // → MCCM_inspect_output/p41-50_inspect.json
```

---

### Step 2: パターン抽出（Claude側）

1. Issyさんから吸い出したJSON（または Drive ファイルID）をもらう
2. JSON から以下の共通パターンを抽出
   - タイトル黒帯（座標・色・フォント）
   - カードレイアウト（KPIカード・資産カード・課題カード等）
   - テーブル装飾（ヘッダー色・行間・ボーダー）
   - 章扉（アジェンダ・セクション区切り）
   - メッセージBOX・注釈BOX
   - 色パレット（実測値）・フォントサイズ階層
3. 抽出結果を `design_patterns.md` にまとめる

---

### Step 3: 再生成ヘルパー関数化 + スキル化（Claude側）

1. `02_helpers.gs` に新規関数を追加（既存関数と命名衝突しないよう接頭辞 `tpl_` で）
   - 例: `tplKpiCard()`, `tplReviewCard()`, `tplAgendaList()`, `tplComparisonTable()`
2. `~/.claude/skills/learned/firework-standard-deck-template.md` を作成
   - 他案件（アシックス・ロレアル・サムスン等）でも使えるテンプレートとして整備
   - 色変数化（クライアントごとに `C_ACCENT` だけ変えれば適用可能に）
3. テスト用GAS `16_template_demo.gs` を追加（テンプレ関数で全パターンを再現するデモ）

---

## 対象ページの内容（参考）

### P2-15（前半）
- P2-3: アジェンダ
- P4: マツココライブの変遷（タイムライン）
- P5: 数字でみる実績（KPIカード4つ）
- P6: 年度別協賛実績（テーブル）
- P7: 7つの資産（カードグリッド）
- P8-9: 2025年度の振り返り（Highlight/Lowlight/学び）
- P10-12: 競合調査（数値比較・リピーター比率・運用開始時期）
- P13: 競合比較（テーブル）
- P14: 4つの課題（カードグリッド + 最重点バッジ）
- P15: リピーター比率テーブル

### P41-50（後半）
- 内容未確認（Drive MCPの content snippet で取得できなかった範囲）
- 吸い出し後に確認

---

## 注意事項

- **画像のバイナリは取得しない**（contentUrl のみ）
- **GASマスター由来の装飾**は API で取得不可なので、目視確認が必要なケースあり
- **ドロップシャドウ**は GAS API では取得・設定不可（既知の制約）
- 出力JSONは **約 200KB-1MB** になる見込み（要素数による）
