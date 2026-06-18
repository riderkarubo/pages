# 02_分科会

マツキヨココカラライブの**月次分科会**プロジェクト。年次報告会とは独立したプロジェクトとして運用。

## フォルダ構成

```
02_分科会/
├── README.md                                       # 本ファイル
├── 01_提案書/
│   └── 分科会用資料_20260414.html                  # 骨子版（メインの分科会資料）
├── 02_完了タスクアーカイブ/
│   └── 完了タスク_A5_20260501.html                 # 資料反映済みタスク格納
└── 03_過去pptx/
    ├── 2024年度/                                   # 11個（2024-04 〜 2025-03の月次分科会pptx）
    └── 2025年度/                                   # 13個（2025-04 〜 2026-03の月次分科会pptx）
```

## 公開URL

- **分科会本体**: https://riderkarubo.github.io/pages/mccm-proposal/bunkakai/
- **完了タスクアーカイブ**: https://riderkarubo.github.io/pages/mccm-proposal/bunkakai/completed/

## 公開版との同期ルール

公開版HTMLは `MCCM/mccm-proposal/bunkakai/` 以下にある。骨子版を編集したら必ず公開版にも反映する：

| 骨子版 | 公開版 |
|---|---|
| `02_分科会/01_提案書/分科会用資料_20260414.html` | `mccm-proposal/bunkakai/index.html` |
| `02_分科会/02_完了タスクアーカイブ/完了タスク_A5_20260501.html` | `mccm-proposal/bunkakai/completed/index.html` |

公開版のみに存在する2行（SSO保護）：
```html
<script src="../auth.js"></script>
<script>document.addEventListener('DOMContentLoaded',function(){FireworkSSO.protect();});</script>
```
（completed/ 配下は `../../auth.js` パス）

## 運用ガイド

### 分科会本体の構造（2階層）

- **Layer 1**: 全体可視化フレーム（GA × 会員ID × POS）— 2026/4/30 中尾FB反映の最上位レイヤー
- **Layer 2**: 配信単体指標分析（A3・A6・A7・A11・A14・A9・A10）

### 完了タスクの扱い

タスクが「分析完了 + 資料反映済み」になったら、以下の2ステップで切り出す：

1. **本体から削除**：分析セクション・追加検証タスクをコメントアウトまたは削除
2. **アーカイブに格納**：完了タスクアーカイブHTMLに移植 + 反映先・完了日を明記

これにより、分科会本体は「進行中タスク」だけに集中できる。

### 過去pptxの命名規則

- `【MCCM様】月次コンテンツ分科会_YYYYMMDD.pptx`
- `YYYYMMDD_コンテンツ分科会.pptx`（旧形式）

## 関連プロジェクト（年次報告会）

- 統合版HTML: `01_年次報告会/提案骨子/01_提案書/提案に向けた情報まとめ_統合版_20260414.html`
- 公開URL: https://riderkarubo.github.io/pages/mccm-proposal/

## 履歴

- **2026-05-01**: 「02_分科会」フォルダ新設・年次報告会から独立。A5（Xフォロリポキャンペーン分析）を完了タスクアーカイブへ移設。過去pptx 24個を本フォルダに集約
