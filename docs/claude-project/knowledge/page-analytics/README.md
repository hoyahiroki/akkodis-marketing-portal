# SharePoint ページ分析（Page Analytics）

Microsoft 365 / SharePoint の「ページ分析」ダイアログから取得した閲覧指標の記録です。

- **取得日（スナップショット）:** 2026-07-15
- **集計期間:** 過去 30 日（グラフ横軸のおおよその範囲: 2026-06-14 〜 2026-07-14）
- **対象:** 5 ページ（`top` / `partner` / `brand` / `client` / `candidate`）
- **指標の出典:** SharePoint 標準の「ページ分析」（Outlook と SharePoint での閲覧を含む）

各ダイアログには次の 4 指標が表示されます。

| 指標 | 説明 |
|------|------|
| 重複しない閲覧者数 | このページを表示した一意のユーザー数 |
| 合計閲覧回数 | ユーザーがこのページを表示した回数 |
| ユーザーあたりの平均時間 | Outlook と SharePoint でページ読み取りに費やした平均時間 |
| 時間別のページ トラフィック | 一意の閲覧者の 1 時間あたり平均数（曜日 × 時間帯ヒートマップ） |

数値の一覧は [`summary.md`](./summary.md)、機械可読データは
[`page-analytics.csv`](./page-analytics.csv) を参照してください。

## ページ一覧

各ページに 3 種類のスクリーンショットがあります。

- `*_ana.png` … ページ分析ダイアログ（数値の出典）
- `*_pc.png` … PC 表示のページ
- `*_sp.png` … スマートフォン表示のページ

| ページ | 分析 | PC | スマホ |
|--------|------|----|--------|
| top       | [assets/top_ana.png](./assets/top_ana.png)             | [assets/top_pc.png](./assets/top_pc.png)             | [assets/top_sp.png](./assets/top_sp.png)             |
| partner   | [assets/partner_ana.png](./assets/partner_ana.png)     | [assets/partner_pc.png](./assets/partner_pc.png)     | [assets/partner_sp.png](./assets/partner_sp.png)     |
| brand     | [assets/brand_ana.png](./assets/brand_ana.png)         | [assets/brand_pc.png](./assets/brand_pc.png)         | [assets/brand_sp.png](./assets/brand_sp.png)         |
| client    | [assets/client_ana.png](./assets/client_ana.png)       | [assets/client_pc.png](./assets/client_pc.png)       | [assets/client_sp.png](./assets/client_sp.png)       |
| candidate | [assets/candidate_ana.png](./assets/candidate_ana.png) | [assets/candidate_pc.png](./assets/candidate_pc.png) | [assets/candidate_sp.png](./assets/candidate_sp.png) |

## 補足 / TODO

- **ページ名の由来:** `assets/` にアップロードされた分析スクリーンショットのファイル名
  （`*_ana.png`）から確定。各画像の数値と転記データを突き合わせて対応付け済み。
- **URL:** 各ページの実 URL が分かれば下表に追記してください。

### ページ URL 対応表（判明したら記入）

| ページ | URL |
|--------|-----|
| top |  |
| partner |  |
| brand |  |
| client |  |
| candidate |  |
