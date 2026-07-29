# SharePoint Online 一次情報調査（V14〜V17）

調査日: 2026-07-29 / 出典は learn.microsoft.com・support.microsoft.com のみ（一次情報）。

---

## V14. カスタムリストは検索に引っかかるか

| # | 論点 | 結論 | 根拠URL | 確度 | 実務上の注意 |
|---|---|---|---|---|---|
| V14-1 | リストアイテムはMicrosoft Search（サイト内検索）にインデックス化され結果に表示されるか | **される（既定でYes）**。「By default every list and library is set to include all items in search results.」 | [Enable content on a site to be searchable](https://learn.microsoft.com/en-us/sharepoint/make-site-content-searchable) / [同内容 support版](https://support.microsoft.com/en-us/office/enable-content-to-be-searchable-d7ba92db-8618-43fe-87ee-adf03d973062) | 高 | 検索結果は常にセキュリティトリミングされる（閲覧権限が無いアイテムは表示されない）。site owner/site collection administratorが上位（サイト単位）で「検索結果への表示」をNoにしていると、リスト単位の設定は無効化される点に注意。 |
| V14-2 | 検索対象化の設定（詳細設定「このリストのアイテムを検索結果に表示する」）の有無 | **あり**。List設定 → 全般設定 → 「詳細設定」→ 検索セクション「Allow items from this list（document library）to appear in search results」= Yes/No を切替可能。変更には**リストの管理（Manage Lists）権限**が必要（Designer/Owner グループが保持）。 | [同上（Enable content on a site to be searchable）](https://learn.microsoft.com/en-us/sharepoint/make-site-content-searchable) | 高 | メンバー権限のみのユーザーにはこのメニュー自体が表示されない（Manage Listsが無いと操作不可）。 |
| V14-3 | ハイパーリンク列に入れた外部URL・表示テキストは検索対象になるか | **なる（列の値はクロールされ管理プロパティ化される）**。ハイパーリンク型のサイト列は自動的にクロールプロパティ `ows_q_URLH_<列名>` → 管理プロパティ `<列名>OWSURLH` にマッピングされ、データ形式は「URL, description」（例: `https://www.contoso.com, Welcome to the home page of Contoso.`）。**URLと表示テキストの両方**が値としてインデックスに取り込まれる。 | [Automatically created managed properties in SharePoint](https://learn.microsoft.com/en-us/sharepoint/technical-reference/automatically-created-managed-properties-in-sharepoint)（SPOの[Overview of crawled and managed properties](https://learn.microsoft.com/en-us/sharepoint/crawled-and-managed-properties-overview)ページ自身がこの記事を「SharePoint Serverから継承される既定プロパティの一覧」として参照しており、SPOにも同じ命名規則が適用される） | 中〜高 | **要確認点**: 自動生成された管理プロパティが既定で「Searchable（フリーワード検索でヒット）」なのか、「Queryable（`列名:値` のプロパティ指定クエリでのみヒット）」止まりなのかは、公式ドキュメント上「auto-generated managed propertyの設定値は検索スキーマ画面上では非表示（hidden）」と明記されており断定できない（[Manage the search schema](https://learn.microsoft.com/en-us/sharepoint/manage-search-schema)）。ただし「Enable content on a site to be searchable」ページの「By default, most content contained in a site, list, library, Web Part page, **or column** will be crawled and added to the search index.」という記述から、通常運用（詳細設定がYesのまま）ではハイパーリンク列の値も含めキーワード検索でヒットするのが一般的挙動と考えられる。厳密に確定させたい場合は実機で `列名:値` と素のキーワードの両方で検索して挙動を確認することを推奨。 |

---

## V15. AI-2（SharePoint agent）はリストの中身を回答根拠にできるか

| # | 論点 | 結論 | 根拠URL | 確度 | 実務上の注意 |
|---|---|---|---|---|---|
| V15-1 | SharePoint agent（サイトスコープの生成AIエージェント／"Agents in SharePoint"）のナレッジソースにリストを含められるか | **含められない**。現行（2026-07-29時点で取得）の公式サポートページに明記: **「Agents currently don't use data from Lists. Also, you can't add pages from the Site Pages library as source for an agent.」** サポートされるソースはサイト・ドキュメントライブラリ・フォルダ・ファイル（最大20項目）。 | [Frequently asked questions about Copilot in SharePoint](https://support.microsoft.com/en-us/sharepoint/copilot-in-sharepoint/frequently-asked-questions-about-copilot-in-sharepoint)（同一内容が別URLの [support.microsoft.com/.../eb1b7668-...](https://support.microsoft.com/en-us/office/frequently-asked-questions-about-copilot-in-sharepoint-eb1b7668-3d98-4a93-98ef-f0c6dfc694f0) でも確認） | 高（ただし要確認: 下記注意参照） | **要確認**: 一部の非公式まとめサイト（Message Center記事の転載ブログ等、learn/support外）は「custom agents in SharePoint and OneDriveでリストが2026年3月〜5月にロールアウト完了」と主張しているが、learn/support一次情報では今回の取得時点でも上記の非対応の記載のまま。ロールアウト中で地域/テナントにより表示にラグがある可能性があるため、**自社テナントの実機で最新のSources選択画面を必ず確認**すること。 |
| V15-2 | 含まれない場合、リストに入れたリンクはAI-2の回答に出てこないか（テキストパーツの方が拾われやすいか） | **その通り**。V15-1の通りリストはナレッジソース対象外のため、リストのアイテム（資料リンク等）はSharePoint agentの回答根拠には使われない。ページ本文（テキストパーツ等）に書いた内容は、SharePoint agentのソースである「ページ」に含まれるため拾われる。 | 同上 | 高 | AI-2に確実に拾わせたい情報（資料リンク等）は、リストではなく**ページ本文（テキストweb part等）に明示的に記載**するのが確実。 |
| V15-3 | Copilot Studio・Agent Builder との違い（両者を区別） | 「SharePointをナレッジソースにできるか」は**製品/機能によって異なる**。①**SharePoint agent（サイトスコープ、V15-1で扱ったもの）**＝リスト非対応。②**Microsoft 365 CopilotのAgent Builder（宣言型エージェント、全社Copilotチャットで作る汎用エージェント。SharePoint agentとは別機能）**＝「SharePoint lists \| Select up to **1** SharePoint list for each agent.」と明記されており**リスト1つまで対応**（添付ファイル列は非対応、20,000行/50MBの上限あり、サイトURLを指定してもそのサイト配下のリストは自動的には含まれず個別にリストのURLを指定する必要がある）。③**Copilot Studio**＝サイトURL・リスト個別URLの両方をナレッジソースにできる（ユーザー提示の理解と一致）。 | [Add knowledge sources to your declarative agent in Microsoft 365 Copilot](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/agent-builder-add-knowledge)（更新日2026-06-18） / [Add SharePoint lists (preview) - Copilot Studio](https://learn.microsoft.com/en-us/microsoft-copilot-studio/agents-experience/knowledge-sharepoint-lists) / [Add SharePoint as a knowledge source - Copilot Studio](https://learn.microsoft.com/en-us/microsoft-copilot-studio/knowledge-add-sharepoint) | 高 | 「SharePointがリストに対応している」という情報を見た場合、それがSharePoint agent（非対応）の話なのか、Agent Builder/Copilot Studio（対応）の話なのかを必ず区別すること。混同した記事が社内外に多い。 |

---

## V16. 利用状況レポートの閲覧権限

| # | 論点 | 結論 | 根拠URL | 確度 | 実務上の注意 |
|---|---|---|---|---|---|
| V16-1 | Site usage（サイト利用状況）の閲覧に必要な権限 | **サイト管理者・所有者・メンバー・訪問者の全階層が閲覧可能**。「Site admins, owners, members, and visitors can view site usage data.」ただし訪問者（Visitor）は「外部ユーザーレポートの実行」「90日間利用状況レポートのダウンロード」はできない。またゲストユーザーはSite Owner権限を持っていてもサイト分析・利用状況データにアクセスできない。 | [View usage data for your SharePoint site](https://support.microsoft.com/en-us/office/view-usage-data-for-your-sharepoint-site-2fa8ddc2-c4b3-4268-8d26-a772dc55779e) | 高 | 「Site usage=管理者専用」という思い込みは誤り。ページ単位の閲覧数・一意閲覧者データが上記の権限階層のどこまで細分化されて見えるか（例: 訪問者が個別ページの閲覧数まで見えるか）はページ内に明示的な記載がなく**要確認**（同じSite usageページの一部として提供されると推測されるが未確定）。 |
| V16-2 | Search usage reports（検索利用状況・ゼロ件検索等）の閲覧に必要な権限 | **サイトコレクション管理者権限が必須**。「Only site collection admins can view search reports data.」と明記。所有者（Owner）権限だけでは閲覧不可（メニュー自体が Site Settings > **Site Collection Administration** 配下にあり、Owner権限のみでは通常このセクション自体が表示されない）。 | [Classic site collection search usage reports](https://learn.microsoft.com/en-us/troubleshoot/sharepoint/search/classic-site-collection-search-usage-reports) / [View search usage reports in modern sites](https://learn.microsoft.com/en-us/sharepoint/view-search-usage-reports-modern-sites)（アクセス経路: Site settings > Site collection administration > Microsoft Search > Configure search settings > Insights） | 高 | サイト所有者だけの体制だと検索利用状況（ゼロ件検索など）は原則見られない。必要なら管理者にサイトコレクション管理者権限の付与を依頼するか、下記の代替手段を使う。 |
| V16-3 | 管理者権限が無いユーザーの代替手段 | 公式ドキュメントに代替手段の明示的な一覧は無いため、以下は**実務上の提案（要検証）**：① Site usage（V16-1）はowner/memberでも見られるため、ページ閲覧数・訪問者数ベースのKPIはそちらで代替。② 検索クエリ・ゼロ件検索等はサイトコレクション管理者に定期エクスポート（Insightsのダウンロード機能）を依頼し、Excelで共有してもらう運用にする。③ Lists/Formsの回答数・アイテム数はアプリ自体のビュー集計やエクスポートで取得可能（ただし検索利用状況とは別軸のKPIであり代替にはならない点に注意）。 | 上記V16-1, V16-2と同じ | 中（提案部分は非一次情報のためCEOの実務判断） | KPI設計時は「検索系KPI（ゼロ件検索率等）は管理者依存」という制約を前提に、担当者の権限レベルに応じて代替指標（ページ閲覧数等）も併用する設計にするのが安全。 |

---

## V17. 強調表示されたコンテンツ web partの適用範囲

| # | 論点 | 結論 | 根拠URL | 確度 | 実務上の注意 |
|---|---|---|---|---|---|
| V17-1 | 「サイト内のファイル/ページ/ニュース等を動的クエリで一覧表示」するものであり、任意の外部URLへのリンク集には使えないか | **その理解で正しい**。ソースとして選択できるのは「This site」「A document library on this site」「This site collection」「The page library on this site」「Select sites（最大30サイト）」「All sites」「All sites in the hub」（ハブサイト接続時）のみで、いずれもSharePoint環境内のサイト/ライブラリを動的クエリの対象にする仕組み。個別の外部URLを1件ずつ指定してリンク集を作る選択肢は一覧に存在しない。 | [Use the Highlighted content web part](https://support.microsoft.com/en-us/sharepoint/pages-in-sharepoint/use-the-highlighted-content-web-part) | 高（ただし「外部URL不可」の直接的な禁止文言は無く、ソース一覧に選択肢が無いことからの論理的推論） | SharePoint Server 2019では選択肢がさらに少なく「This site」「A document library on this site」「This site collection」「All sites」のみ（Select sitesは不可）。GCC High/DoD/21Vianetでも「Select sites」不可。 |
| V17-2 | 「リンク先が1個や6個のリンク集」を見せる用途に適した標準パーツ | **Quick Links web part**が該当。ページ内・外部リソースへのリンクを個別に「ピン留め」でき、リンクごとにタイトル・画像/アイコン・説明文をカスタマイズ可能。Microsoft 365グループ/セキュリティグループ単位で表示対象を絞ることもできる。テキストweb part内にハイパーリンクを埋め込む方法も同様に使える（現行のSharePoint標準web partライブラリに独立した「Link web part」という名称のパーツは無く、この用途はQuick Linksが担う）。 | [Use the Quick Links web part](https://support.microsoft.com/en-us/office/use-the-quick-links-web-part-e1df7561-209d-4362-96d4-469f85ab2a82) | 高 | 6個程度の固定リンク集であれば、動的クエリ型のHighlighted contentより、手動で並び順・見た目を制御できるQuick Linksの方が保守しやすい。 |

---

## 推奨: クライアントページで「リスト」を使うべきか「テキストパーツ」にすべきか

V15-1/V15-2の通り、**SharePoint agent（AI-2）は現行仕様でリストのデータを回答根拠に一切使わない**（サイト・ドキュメントライブラリ・フォルダ・ファイルのみが対象）ため、AI-2に資料リンクを拾わせたいなら**ページ本文（テキストパーツ/Quick Links等）に書く一択**である。一方、V14の通りリストアイテムは既定でMicrosoft Search（サイト内検索）にはインデックスされ、人間がサイト内検索窓で探す用途には有効に機能する。したがって判断軸は「AI-2に拾わせたいか／人間の検索窓だけで足りるか」であり、**AI-2の回答根拠にする必要がある資料リンクは必ずページ本文（テキストパーツ）側に記載し、リストは補助的な一覧管理・サイト内検索用途に留める**のが現時点（2026-07-29）での推奨。将来リストがSharePoint agentのナレッジソースとして正式対応された場合（非公式情報では対応の兆しがあるが本調査時点の一次情報では未確認）は運用を見直す価値がある。
