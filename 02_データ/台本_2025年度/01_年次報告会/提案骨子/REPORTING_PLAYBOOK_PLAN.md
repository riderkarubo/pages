# 年次／四半期／月次報告会 横断プレイブック化プラン

最終更新: 2026年5月1日 21:15
作成者: Issy + Claude
スコープ: MCCM年次報告会 (2026.4-5実施分) で得た知見を、**四半期報告会・他クライアント (アシックス・ロレアル・サムスン・コーセー等)・月次分科会** に再利用できる仕組みに昇格させる。

---

## 0. ゴール定義

### 何を実現するか

1. **次回の四半期報告会で、立ち上げから初回ドラフト提示まで「半分以下の時間」で到達する**
2. **他クライアントの年次/四半期報告会でも、ブランド色とコンテンツ差し替えだけで同水準のアウトプットを再現できる**
3. **プロセス品質（エグゼクティブセルフレビュー・データ前提条件チェック・薬機法CK等）を毎回再発見せずに、最初から織り込める**

### 何を再利用するか（3レイヤー）

| レイヤー | 内容 | 既存資産 |
|---|---|---|
| **L1: コンテンツ骨格** | アジェンダ構成、章立て、論理展開（Highlight/Lowlight/学び→競合→課題→次年度方針→施策） | mccm-proposal/index.html、提案骨子配下 |
| **L2: デザインシステム** | 黒帯タイトル系統A/B、KPIカード、3カラム戦略、テーブル装飾、色階層、フォント階層 | learned_layouts/ (作成済み) |
| **L3: プロセス品質** | エグゼクティブセルフレビュー、データ前提条件チェック、表記ルール、HTML編集ルーティーン | ~/.claude/rules/、project memory |

---

## 1. 現状の資産マップ（棚卸し結果）

### A. すでに整備済み

| カテゴリ | 場所 | 役割 |
|---|---|---|
| グローバルルール (rules/) | `~/.claude/rules/executive-self-review.md` | エグゼ向け資料の3ステップセルフレビュー |
| グローバルルール | `~/.claude/rules/file-sync.md` | MD↔HTML同期 |
| グローバルルール | `~/.claude/rules/html-ogp.md` | HTML作成時のOGP・Last Updated・印刷対応 |
| グローバルルール | `~/.claude/rules/sso-requirement.md` | GitHub Pages SSO保護 |
| 学習スキル | `~/.claude/skills/learned/firework-final-deck-template-v2.md` | 最新デザインテンプレ仕様 |
| 学習スキル | `~/.claude/skills/learned/firework-template-design-patterns.md` | v1デザインパターン |
| 学習スキル | `~/.claude/skills/learned/gas-slides-noto-sans-jp-text-width.md` | 文字幅実測係数 |
| 学習スキル | `~/.claude/skills/learned/data-comparison-prerequisite-check.md` | データ比較前のチェック |
| 学習スキル | `~/.claude/skills/learned/single-source-of-truth-lock.md` | マスター情報源固定 |
| 学習スキル | `~/.claude/skills/learned/csv-excel-ratio-scale-detection.md` | 比率スケール混在検知 |
| 学習スキル | `~/.claude/skills/learned/html-section-restructure-pattern.md` | HTML一括並び替え |
| 学習スキル | `~/.claude/skills/learned/html-to-slide-80percent-cut-rule.md` | HTMLをスライド化する圧縮率 |
| プロジェクトメモリ | `~/.claude/projects/-Users-issy-...-MCCM/memory/MEMORY.md` | MCCM固有知見 |
| 既存スキル(slash) | `/fw-slide` | スライド生成 |
| 既存スキル(slash) | `/proposal-builder` | HTML提案書生成 |
| 既存スキル(slash) | `/delivery-analysis` | 配信データ分析 |
| 既存スキル(slash) | `/strategy-plan` | 戦略プラン |
| 既存スキル(slash) | `/pitch-deck` | ピッチデック構成 |
| 既存スキル(slash) | `/asics-proposal` | アシックス提案書 |
| GASヘルパー | MCCM/01_年次報告会/提案骨子/05_スライド化/02_helpers.gs | 基盤ヘルパー (txt/rect/makeTable等) |
| GASテンプレ | MCCM/.../learned_layouts/16_template_helpers.gs | 11パターン再生成関数 |
| 仕様書 | MCCM/.../learned_layouts/design_patterns.md | 11パターン完全仕様 |
| 吸い出しJSON | MCCM/.../learned_layouts/p02-15_p41-50.json | 全要素絶対座標+色 |

### B. 不足している/再利用しづらいもの

1. **「報告会を回すワークフロー」全体を貫く統合スキルがない**
   - 現状: `/proposal-builder`(HTML)・`/fw-slide`(スライド)・`/delivery-analysis`(分析)・`/strategy-plan`(戦略) が**個別に存在**するが、報告会1サイクルを通じてどう繋ぐかのプレイブックが未整備
2. **コンテンツ骨格テンプレート (md) がない**
   - 「Highlight/Lowlight/学び」「7つの資産」「4つの課題」「3つの戦略」など章立てのひな型がない
   - 各クライアントごとに毎回考え直している
3. **プロジェクトCLAUDE.mdのテンプレートがない**
   - MCCM/CLAUDE.md は400行超えで成熟しているが、新規クライアントのCLAUDE.mdに「最低限これは入れる」という骨子テンプレが未定義
4. **クライアント横断の表記ルールがバラバラ**
   - 「貴社」呼称・「Firework」表記・「マツキヨココカラライブ」のようなブランド名統一は MCCM限定で運用中
5. **HTMLマスター・スライドマスターの2系統運用パターンが暗黙知**
   - 提案書HTML(GitHub Pages公開) → スライド(Google Slides)の橋渡しがMCCMだけの実績
6. **エグゼクティブセルフレビューが「ルール」として存在するが、ワークフロー内のチェックポイント化されていない**
   - 提出前に必ず通る関門になっていない（Issyさんが意識して呼ぶ必要あり）
7. **吸い出しGAS (15_inspectFinalDeck.gs) を他案件で使いまわすには、Presentation IDだけ変えれば良い設計になっていない**
   - `SOURCE_PRES_ID` が定数なので、他案件では複製してID差し替えが必要
8. **過去スライドJSONアーカイブの位置が決まっていない**
   - p02-15_p41-50.json は MCCMリポにあるが、**横断参照しやすい場所** (例: `_shared/learned_decks/`) に置くべき

---

## 2. プラン: 4つのアウトプット

### Output A: 統合スキル `/reporting-cycle` の新規作成 (最重要)

**目的**: 報告会1サイクルの全工程を1つのスキルから呼べるようにする。`/proposal-builder` `/fw-slide` `/delivery-analysis` 等の既存スキルを**オーケストレーション**する上位スキル。

#### スキル定義（プレイブック骨格）

```
/reporting-cycle <client> <type> [phase]
  client: mccm / asics / loreal / samsung / kose ...
  type:   annual (年次) / quarterly (四半期) / monthly (月次)
  phase:  setup / data-prep / draft / review / finalize / handoff
```

#### 6フェーズ・プレイブック

| Phase | 内容 | 呼び出すスキル/ツール | 成果物 |
|---|---|---|---|
| **1. setup** | クライアント情報・前回資料・データソース・ステークホルダー・締切を整理 | (新規 reporting-cycle 内) | `00_setup.md` |
| **2. data-prep** | 配信数値・KPI・競合データ・前期比較を集約 | `/delivery-analysis` | `01_data_summary.md` |
| **3. draft** | 章立て決定 + HTML提案書生成（マスター） | `/proposal-builder` + コンテンツ骨格テンプレ | `02_proposal_draft.html` |
| **4. review** | エグゼクティブセルフレビュー + データ前提条件チェック | `executive-self-review.md` + `data-comparison-prerequisite-check.md` | `03_review_log.md` |
| **5. finalize** | スライド化 (HTMLからの80%カット) + デザイン適用 | `/fw-slide` + `firework-final-deck-template-v2.md` | `04_final_deck.gs` (Slides化済み) |
| **6. handoff** | 次回への引き継ぎメモ + JSONアーカイブ + 学習スキル更新 | `/session-handoff` + (新規 archive-deck) | `99_handoff.md` |

#### 既存スキルとの差別化

- `/proposal-builder` = 単発HTML作成 → このスキルは Phase 3 のみ担当
- `/fw-slide` = 単発スライド作成 → Phase 5 のみ担当
- **`/reporting-cycle` は 6フェーズ全体の地図とチェックポイント** を提供する

#### 配置先

- スキルmd実体: `~/.claude/skills/reporting-cycle/SKILL.md`
- (もしくは既存配布構造に合わせて) `~/.claude/skills/learned/reporting-cycle.md` + slash command 登録

---

### Output B: コンテンツ骨格テンプレ集 `_shared/reporting-templates/`

**目的**: クライアント名・数字・固有名詞だけ差し替えれば「報告会の章立て」が出来る、 markdownテンプレート集。

#### 配置

```
03_クライアント/
└── _shared/
    ├── slides/                      ← 既存
    └── reporting-templates/          ← 新設
        ├── README.md                 ← 使い方
        ├── annual-report-skeleton.md ← 年次報告会の章立て骨子
        ├── quarterly-report-skeleton.md
        ├── monthly-report-skeleton.md
        ├── chapters/
        │   ├── 01_変遷タイムライン.md
        │   ├── 02_数字でみる実績_KPIカード.md
        │   ├── 03_資産の棚卸し_カードグリッド.md
        │   ├── 04_Highlight_Lowlight_学び.md
        │   ├── 05_競合調査_4観点.md
        │   ├── 06_4つの課題_最重点バッジ.md
        │   ├── 07_次年度KPI目標_テーブル.md
        │   ├── 08_3つの戦略_3カラム.md
        │   ├── 09_体制改善案_2BOX.md
        │   └── 10_エグゼクティブサマリ.md
        └── examples/
            ├── mccm-2026-annual/  ← 今回のMCCM事例（参照用）
            └── (将来の事例)
```

#### 各章テンプレの中身（例: 04_Highlight_Lowlight_学び.md）

```markdown
# 章: Highlight / Lowlight / 学び

## いつ使う
報告会で「振り返り」セクションを作るとき。年次・四半期・月次すべてで使える普遍パターン。

## デザインパターン
firework-final-deck-template-v2.md の Pattern E (`tpl_reviewTriBox`)

## コンテンツ穴埋めテンプレ

### Highlight (左上 ピンク)
- 効率性: {N}名体制で{頻度}実施
- {成果指標}: {数値}達成

### Lowlight (右上 紺)
- {良かった指標}は{推移}する一方、{悪かった指標}は伸び悩み
  ※必ず「〜する一方、〜」の1メッセージ統合（独立2項目並列はHighlight誤読リスク）

### 学び (下 緑)
- {悪かった指標}が伸び悩み。これまで{打ち手}を試みてきたが、{構造的限界の説明}。
- → {次年度の最重要テーマ} = {ポジティブな転換メッセージ}

## アンチパターン
- 自社の落ち度を不必要に強調する（コンサル立場は維持）
- ネガティブだけで終わる（学びは必ず未来への転換で締める）

## 参照
- ~/.claude/rules/executive-self-review.md
- 過去事例: examples/mccm-2026-annual/
```

→ 各章テンプレに **デザインパターン関数名・コンテンツ穴埋め枠・アンチパターン・参照** を記載することで、初見でも完成形に到達できる。

---

### Output C: 学習スキル拡充

#### 新規作成

| ファイル | 内容 |
|---|---|
| `~/.claude/skills/learned/reporting-cycle-playbook.md` | 6フェーズの全体ガイド |
| `~/.claude/skills/learned/firework-content-skeleton.md` | 章立て骨格集 (Output Bへの誘導) |
| `~/.claude/skills/learned/multi-pc-git-recovery-pattern.md` | 今回のrebase衝突解消の事例 |
| `~/.claude/skills/learned/slide-deck-inspect-and-template.md` | 「完成版から吸い出してテンプレ化」の汎用パターン |

#### 更新

| ファイル | 更新内容 |
|---|---|
| `~/.claude/skills/learned/firework-final-deck-template-v2.md` | 「他案件転用時の手順」セクション追加 (Issy指示で確認) |
| `~/.claude/skills/learned/firework-slide-helper-architecture.md` | tpl_接頭辞のテンプレ関数群を追記 |

---

### Output D: グローバルルール拡充

#### 新規

| ファイル | 内容 |
|---|---|
| `~/.claude/rules/reporting-quality-gates.md` | 報告会資料の品質ゲート（提出前チェックリスト） |
| `~/.claude/rules/client-naming-convention.md` | 「貴社」呼称・正式社名・サービス名の統一ルール（MCCM以外にも展開） |

#### 更新（ルール統合）

- `editing-conventions.md` に「マスター情報源の単一固定ルール」を統合
- `executive-self-review.md` に「報告会フェーズ4の必須関門」と明記

---

## 3. 実行順序（推奨）

優先度・依存関係を考慮した順番。**A→B→C→D の順** が論理的だが、**Issyさん実用優先なら B (コンテンツ骨格) を先に**するのが効率的。

### 推奨順 (実用優先・全部で5-7時間想定)

| # | 工程 | 所要 | アウトプット | 依存 |
|---|---|---|---|---|
| 1 | **B: コンテンツ骨格テンプレ集 作成** | 2h | `_shared/reporting-templates/` 全体 | なし |
| 2 | **D: グローバルルール `reporting-quality-gates.md` `client-naming-convention.md`** | 30min | rules/ に2ファイル追加 | なし |
| 3 | **C: 学習スキル4本** | 1.5h | learned/ に4ファイル追加 | B完了 |
| 4 | **A: `/reporting-cycle` 統合スキル新規作成** | 2h | reporting-cycle/SKILL.md + slash登録 | B,C,D 完了 |
| 5 | **検証: アシックス案件 or 次の四半期報告会で実走** | 実案件で検証 | プレイブック改善点フィードバック | A完了 |
| 6 | **更新: フィードバック反映 v2** | 30min | 各ファイル更新 | 5完了後 |

### 段階的アプローチ案

**Phase 1 (今週)**: B + D だけ整備 → MCCM/CLAUDE.md と他クライアントから即座に参照可能に
**Phase 2 (次の四半期報告会の前)**: C + A 整備 → スキル化して slash command で呼べるように
**Phase 3 (実走後)**: 検証→改善

---

## 4. 詳細仕様 (Output B 最優先のため先出し)

### `_shared/reporting-templates/annual-report-skeleton.md` ドラフト

```markdown
# 年次報告会 章立て骨子テンプレート

## 想定アウトプット
- 提案HTML (mccm-proposal/index.html 相当): GitHub Pages公開・SSO保護
- 提案スライド (Google Slides): プレゼン用
- ディスカッション用md: 内部議論用

## 章構成 (10-12章 / 50-60ページ想定)

### 第1部: 振り返り
1. {サービス名}の変遷 (Pattern B: タイムライン9項目)
2. 数字でみる{サービス名}実績 (Pattern C: KPIカード4分割)
3. 年度別 {主要指標} の推移 (Pattern G: 標準テーブル)
4. {サービス名}における「{N}つの資産」 (Pattern D: グリッドカード)
5. {年度}の振り返り (Pattern E: Highlight/Lowlight/学び)

### 第2部: 競合・課題
6. 競合調査: {観点1} (Pattern G + Pattern J)
7. 競合調査からみえた{N}つの課題 (Pattern F: 4課題カード+最重点バッジ)
8. {ポテンシャル領域}について (Pattern G: 標準テーブル)

### 第3部: 次年度方針
9. {次年度} {サービス名} 目標 (Pattern G: KPI比較テーブル)
10. {施策軸1}：方向性と考え方 (Pattern H: 大KPIメッセージ+2カード)
11. {重要施策2}を増やす{N}つの戦略 (Pattern I: 3カラム戦略)
12. {現状課題1/2/3}：プラン/体制/AI活用 (Pattern J: 2カラム課題BOX)

### 第4部: 締め
13. ディスカッション (Pattern A: アジェンダ風)
14. エグゼクティブサマリ (Pattern K + 大文字メッセージ)

## 各章で使うパターン → 関数 対応表
firework-final-deck-template-v2.md 参照

## 章ごとに必要な情報

### 1. {サービス名}の変遷
- 9-10件の重要マイルストーン
- 日付 + 出来事 + 重要度フラグ
- 起点〜現在〜未来の3区分

### 2. 数字でみる実績
- 4つの主要KPI
- 各KPIに「累計値・期間内訳・比較」
- 例: 累計視聴 / 累計売上 / 配信回数 / 平均視聴者数

(以下章ごとに必要情報を列挙)
```

---

## 5. リスクと対策

| リスク | 内容 | 対策 |
|---|---|---|
| 過剰汎化 | 「すべてのクライアントで使える」を狙うと薄くなる | クライアント別の例外を `examples/` に蓄積。共通骨格は最小限 |
| ドキュメント乖離 | 実案件と骨格テンプレが乖離していく | プロジェクト完了時に **Phase 6 handoff で必ず差分を骨格に反映** をワークフロー化 |
| スキル名衝突 | `/reporting-cycle` が既存スキルと重複 | 接頭辞 `report-` で `/report-cycle` `/report-skeleton` `/report-archive` 統一 |
| Issy以外のメンバーが使えない | コンテキスト依存（マゼンタ #FA006D 等） | テンプレ変数でブランド色を抽出し、初回setup時に質問する設計 |
| 過去資産が散在 | MCCM配下にだけある資産は他案件から参照しづらい | `_shared/` に **シンボリックリンク or コピー配置** で横断参照可能に |

---

## 6. 確認したいこと（Issyさんへ）

1. **このプランで進めますか？** 段階的 (Phase 1 から) でいきますか？
2. **`/reporting-cycle` のスキル名は良いですか？** 別案: `/report-flow` `/quarterly-prep` 等
3. **コンテンツ骨格テンプレの配置先**: `_shared/reporting-templates/` でOK？
4. **吸い出しJSON のアーカイブ場所**: `_shared/learned_decks/{client}-{date}.json` のような横断参照可能な場所に置きたいが、どうしますか？
5. **既存 `/proposal-builder` `/fw-slide` の再ラベル**: 「これは Phase 3/5 で呼ぶスキル」と位置付けを明示するために、既存スキルmdに誘導文を追加したい（簡単）。OK？
6. **次の四半期報告会の予定はありますか？** あれば、それを「最初の検証案件」として Phase 1 完了後に実走したい

---

## 7. 即着手できる最小セット (もし「とりあえず始めて」と言われたら)

**最小コスト・最大効果の3点だけ即着手**:

1. **`_shared/reporting-templates/annual-report-skeleton.md` を作成** (30min)
   - 今回のMCCM章立てを「クライアント名差し替え可能形式」で書き起こすだけ
2. **`~/.claude/rules/reporting-quality-gates.md` を作成** (15min)
   - 「報告会資料を提出する前に必ず通すチェックリスト」を1ファイル化
3. **`~/.claude/skills/learned/reporting-cycle-playbook.md` を作成** (30min)
   - 6フェーズの早見表だけ作る（詳細は後で）

これだけで、次回の四半期報告会で「あれ、何から始めるんだっけ？」が消える。

---

## 結論

**推奨**: まず Output B の `_shared/reporting-templates/` を整備し、ルール2本を追加するだけで Phase 1 完了。次の四半期報告会で実走 → フィードバック → A の統合スキル化（Phase 2）に進む。

「ulw」指定なので、 GO の合図があれば Phase 1 から順次着手します。
