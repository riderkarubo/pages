// Creative Brain Inc. 経営データ（正本）
// 更新は分析官(growth-analyst)または参謀本部のみ。updatedAt は date コマンド実測値。
// followers は一次データのみ。未確認は null（ダッシュボードでは「未報告」表示）。
// 目標の正典: ../creative-brain/road-to-1000.html（年内1,000人 作戦命令書）
window.COMPANY_DATA = {
  company: {
    name: "Creative Brain Inc.",
    ceo: "Issy",
    chiefOfStaff: "Claude（参謀本部）",
    founded: "2026-07-15"
  },
  goal: {
    label: "年内 総フォロワー 1,000人",
    target: 1000,
    deadline: "2026-12-31",
    milestones: [100, 300, 600]
  },
  updatedAt: "2026年7月16日 17:27",
  platforms: [
    {
      id: "youtube",
      name: "YouTube",
      handle: "@ishijima_issy（【底辺クリエイター】）",
      url: "https://www.youtube.com/@ishijima_issy",
      followers: 91,
      history: [
        { date: "2026-07-15", count: 91 }
      ]
    },
    {
      id: "tiktok",
      name: "TikTok",
      handle: "@ishijima_issy",
      url: "https://www.tiktok.com/@ishijima_issy",
      followers: 53,
      history: [
        { date: "2026-07-15", count: 53 }
      ]
    }
  ],
  approvals: [
    {
      id: "AP-001",
      title: "アカウント情報とフォロワー数の初回報告",
      detail: "完了: YouTube 91人（実測）・TikTok 53人（社長報告）。総勢144人で開戦（2026-07-15）。",
      dept: "分析官",
      status: "approved"
    },
    {
      id: "AP-002",
      title: "初回戦略会議（朝会）の開催",
      detail: "Claude Code で /ceo を実行すると朝会が開き、各部門への初回指令を発令できます。",
      dept: "参謀本部",
      status: "pending"
    },
    {
      id: "AP-003",
      title: "第3のチャンネル「へんじん図鑑」（登録者116人）を目標に算入するか",
      detail: "Studio切替時に発見。算入すれば総勢260人（行程26%）に跳ね上がりますが、目標の定義（正典road-to-1000は創作チャンネル前提）に関わるため社長判断を仰ぎます。",
      dept: "参謀本部",
      status: "pending"
    }
  ],
  directives: [],
  tasks: [
    { dept: "分析官", title: "現状フォロワー数の棚卸しと初回戦況レポート", status: "完了", due: "2026-07-15" },
    { dept: "トレンド偵察", title: "直近30日の海外バズ形式・AI動画ハック偵察報告", status: "着手可", due: "初回朝会後" },
    { dept: "企画部", title: "8月コンテンツ計画（MV×欠片の二層カレンダー）3案", status: "着手可", due: "初回朝会後" },
    { dept: "運用部", title: "週次投稿カレンダー叩き台（プレイブック手②準拠）", status: "着手可", due: "初回朝会後" },
    { dept: "素材司書", title: "既存素材・過去作の棚卸し台帳 初版", status: "着手可", due: "随時" }
  ],
  log: [
    {
      date: "2026-07-16",
      dept: "参謀本部",
      text: "社長指示により基地俯瞰図（アイソメトリックマップ）を第1図として増設。全部門を建物として配置し、識別色ヘルメットの隊員アバターを配備。作戦基地をpages（creative-brain-inc/）へ公開デプロイ。"
    },
    {
      date: "2026-07-16",
      dept: "参謀本部",
      text: "社長指示により基地を山岳要塞型に改装。地上部（回転レーダー・格子アンテナ・射出レール・EV上屋）と地層断面を増設し、全乗員をコンソール卓に配置。"
    },
    {
      date: "2026-07-16",
      dept: "参謀本部",
      text: "バーチャルオフィス「作戦基地 BRAIN BASE」を開設（office.html）。基地内部透視図に9部門＋司令部を配置し、状態ランプ・目標接近レーダー・呼集ボタンを実装。"
    },
    {
      date: "2026-07-15",
      dept: "参謀本部",
      text: "維持率の取得経路「道B」を開通（貴殿のChromeのYouTube Studioを【底辺クリエイター】に切替済み）。初計測: 最新ショート平均視聴率53.9%・過去28日視聴1,338回。第3チャンネル「へんじん図鑑」116人を発見（AP-003）。"
    },
    {
      date: "2026-07-15",
      dept: "参謀本部",
      text: "社長指示によりダッシュボードを全面改装（ワークフローアプリ準拠のダークUI・明朝廃止・ステップ円トラッカー）。"
    },
    {
      date: "2026-07-15",
      dept: "参謀本部",
      text: "社長訂正により目標を「年内 総フォロワー1,000人」に改定（正典: road-to-1000.html）。改定後の必要ペースは週約36人。第一関門100人は突破済み。"
    },
    {
      date: "2026-07-15",
      dept: "分析官",
      text: "TikTok 53人を社長報告により記録。総フォロワー144人が開戦時の正式な現在地となる。AP-001決裁完了。"
    },
    {
      date: "2026-07-15",
      dept: "分析官",
      text: "初回戦況レポート納品（成果物/260715_分析官_初回戦況レポート.md）。発見面は突破済み・登録転換0.29%がボトルネックと診断。"
    },
    {
      date: "2026-07-15",
      dept: "分析官",
      text: "YouTube登録者91人・動画15本を実測（チャンネル名【底辺クリエイター】）。"
    },
    {
      date: "2026-07-15",
      dept: "参謀本部",
      text: "Creative Brain Inc. 創業。9部門を組閣し、経営ダッシュボードを開設。目標「年内 総フォロワー1,000人」を制定。"
    }
  ]
};
